"""
fs_menu.py — Un solo icono que pregunta qué quiere hacer.

Existe porque los dos caminos se confunden con facilidad: ambos aceptan que
se les arrastre el Excel, pero uno crea un documento nuevo y el otro
actualiza el que ya hay. Aquí se elige explícitamente.

NO duplica lógica: delega en los mismos módulos que usan los ejecutables
individuales, que siguen funcionando por su cuenta.

    generador_fs.ejecutar()   ->  GeneradorFS.exe  /  generar.bat
    refrescar_fs.ejecutar()   ->  RefrescarFS.exe  /  refrescar.bat
    fs_documento.estado()     ->  la radiografía del proyecto

Uso
---
    Arrastre su Excel sobre EstadosFinancieros.exe  (o doble clic)

    python fs_menu.py [libro.xlsx]              menú interactivo
    python fs_menu.py [libro.xlsx] --refrescar   sin menú
    python fs_menu.py [libro.xlsx] --generar     sin menú
    python fs_menu.py --estado                   sin menú
    python fs_menu.py --desbloquear              permite teclear las cifras
    python fs_menu.py --bloquear                 vuelve a protegerlas
    python fs_menu.py --consola                  menú de texto, sin ventana
"""
import shutil
import sys
import traceback
from pathlib import Path

_AQUI = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
if str(_AQUI) not in sys.path:
    sys.path.insert(0, str(_AQUI))

import generador_fs as G
import fs_documento as D
import refrescar_fs as R


ANCHO = 68


def _opcion(argv, nombre):
    """El valor que sigue a una bandera, o None si no lleva ninguno."""
    for i, a in enumerate(argv):
        if a.lower() == nombre and i + 1 < len(argv):
            siguiente = argv[i + 1]
            if not siguiente.startswith("--"):
                return siguiente
    return None


def _cabecera():
    print()
    print("=" * ANCHO)
    print(" ESTADOS FINANCIEROS — ¿qué quiere hacer?")
    print("=" * ANCHO)


def _describir_destino():
    """Una línea sobre el documento base, para que la opción 1 se entienda."""
    try:
        cfg = G.cargar_config()
        doc = D.resolver_documento(None, cfg)
        culpables = D.quien_bloquea(doc)
        if culpables:
            return f"{doc.name}  [ABIERTO: {culpables[0]}]"
        return doc.name
    except Exception:
        return "sin configurar (config.json -> documento_base)"


def _resolver_libro(args):
    """Qué Excel se va a leer, y de dónde ha salido.

    Devuelve (ruta o None, texto para la ventana, aviso o None).

    Hay que saberlo ANTES de dibujar el menú. Si no, el usuario arrastra su
    libro y la ventana no le dice si lo ha recogido; y cuando no arrastra
    nada, encontrar_excel_por_convencion() coge en silencio cualquier .xlsx
    de la carpeta cuyo nombre contenga la clave de convención —incluido el
    de ejemplos\— y las cifras de muestra acaban dentro del documento real
    sin que nadie lo haya pedido.
    """
    if args:
        ruta = Path(args[0]).resolve()
        if not ruta.exists():
            return (None, f"NO EXISTE: {ruta.name}",
                    f"El archivo indicado no existe:\n  {ruta}")
        return ruta, ruta.name, None
    try:
        cfg = G.cargar_config()
        ruta = G.encontrar_excel_por_convencion(cfg)
    except Exception:
        return None, "", None
    return ruta, f"{ruta.name}  (elegido solo, por convención de nombre)", (
        f"No se arrastró ningún Excel. Se ha elegido «{ruta.name}» porque su\n"
        f"nombre contiene «{cfg.get('buscar_por_convencion')}» y es el más reciente\n"
        f"de la carpeta. Si no es el libro que quería, cierre y arrástrelo encima."
    )


#: Ventana del menú. Se dibuja con WinForms desde PowerShell porque el
#: Python portable no trae tkinter y meterlo obligaría a arrastrar Tcl/Tk
#: dentro del .exe (unos 10 MB más y un empaquetado más frágil). PowerShell
#: y .NET están en cualquier Windows, así que no añade ninguna dependencia.
#: Si algo falla, se cae al menú de consola de siempre.
_PS_VENTANA = r"""
$ErrorActionPreference = 'Stop'

# --- Nitidez en pantallas de alta densidad -------------------------------
# Sin esto, Windows dibuja la ventana a 96 ppp y la escala como una imagen:
# el texto sale borroso ("pixelado"). Hay que declararlo ANTES de crear
# ninguna ventana.
$codigo = @"
using System;
using System.Runtime.InteropServices;
public static class Ppp {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern int SetProcessDpiAwarenessContext(IntPtr v);
}
"@
try { Add-Type -TypeDefinition $codigo -ErrorAction Stop } catch {}
try   { [void][Ppp]::SetProcessDpiAwarenessContext([IntPtr](-4)) }   # por monitor v2
catch { try { [void][Ppp]::SetProcessDPIAware() } catch {} }

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

$documento = if ($args.Count -ge 1) { $args[0] } else { '' }
$libro     = if ($args.Count -ge 2) { $args[1] } else { '' }
$captura   = if ($args.Count -ge 3) { $args[2] } else { '' }

# escala real de la pantalla
$g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
$k = $g.DpiX / 96.0
$g.Dispose()
function S($n) { [int][Math]::Round($n * $k) }

# --- Paleta --------------------------------------------------------------
$tinta   = [System.Drawing.Color]::FromArgb(26, 34, 38)
$suave   = [System.Drawing.Color]::FromArgb(102, 114, 120)
$fondo   = [System.Drawing.Color]::FromArgb(247, 246, 243)
$tarjeta = [System.Drawing.Color]::White
$borde   = [System.Drawing.Color]::FromArgb(224, 221, 214)
$acento  = [System.Drawing.Color]::FromArgb(17, 94, 89)
$ambar   = [System.Drawing.Color]::FromArgb(146, 94, 20)

# Segoe MDL2 Assets trae los iconos de Windows como tipografia: son
# vectoriales, se ven nitidos a cualquier escala y no hay que empaquetar
# ningun archivo. Si no estuviera, se cae a texto sin icono.
$hayIconos = $false
try {
  $fam = New-Object System.Drawing.FontFamily('Segoe MDL2 Assets')
  $hayIconos = $true
  $fam.Dispose()
} catch { $hayIconos = $false }

$fTitulo  = New-Object System.Drawing.Font('Segoe UI Semibold', 18, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$fSub     = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$fOpcion  = New-Object System.Drawing.Font('Segoe UI Semibold', 11.5, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$fDetalle = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$fIcono   = if ($hayIconos) { New-Object System.Drawing.Font('Segoe MDL2 Assets', 15, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point) } else { $fOpcion }
$fInfo    = if ($hayIconos) { New-Object System.Drawing.Font('Segoe MDL2 Assets', 11, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point) } else { $fDetalle }

$f = New-Object System.Windows.Forms.Form
$f.Text = 'Estados Financieros'
$f.ClientSize = New-Object System.Drawing.Size((S 660), (S 600))
$f.StartPosition = 'CenterScreen'
$f.BackColor = $fondo
$f.FormBorderStyle = 'FixedSingle'
$f.MaximizeBox = $false
$f.ShowIcon = $false
$f.Font = $fSub

$pistas = New-Object System.Windows.Forms.ToolTip
$pistas.InitialDelay = 200
$pistas.ReshowDelay = 100
$pistas.AutoPopDelay = 20000

$lTitulo = New-Object System.Windows.Forms.Label
$lTitulo.Text = 'Estados Financieros'
$lTitulo.Font = $fTitulo
$lTitulo.ForeColor = $tinta
$lTitulo.AutoSize = $true
$lTitulo.Location = New-Object System.Drawing.Point((S 30), (S 24))
$f.Controls.Add($lTitulo)

$lSub = New-Object System.Windows.Forms.Label
$lSub.Text = if ($documento) { "Documento: $documento" } else { 'Sin documento configurado — use «Cambiar el documento»' }
$lSub.Font = $fSub
$lSub.ForeColor = $suave
$lSub.AutoEllipsis = $true
$lSub.Location = New-Object System.Drawing.Point((S 32), (S 62))
$lSub.Size = New-Object System.Drawing.Size((S 600), (S 22))
$f.Controls.Add($lSub)

# De donde salen las cifras. Sin esta linea no habia forma de saber si el
# Excel que se arrastro se recogio o no, ni de distinguirlo del que la
# aplicacion elige sola por convencion de nombre.
$lLibro = New-Object System.Windows.Forms.Label
$lLibro.Text = if ($libro) { "Excel: $libro" } else { 'SIN EXCEL - arrastre el libro sobre el icono; las opciones 1 y 2 lo necesitan' }
$lLibro.Font = $fSub
$lLibro.ForeColor = if ($libro) { $suave } else { $ambar }
$lLibro.AutoEllipsis = $true
$lLibro.Location = New-Object System.Drawing.Point((S 32), (S 84))
$lLibro.Size = New-Object System.Drawing.Size((S 600), (S 22))
$f.Controls.Add($lLibro)

$script:eleccion = '0'

function Nueva-Opcion($icono, $texto, $detalle, $ayuda, $y, $valor, $color) {
  $alto = S 62

  $panel = New-Object System.Windows.Forms.Panel
  $panel.Location = New-Object System.Drawing.Point((S 28), (S $y))
  $panel.Size = New-Object System.Drawing.Size((S 604), $alto)
  $panel.BackColor = $tarjeta
  $panel.Cursor = [System.Windows.Forms.Cursors]::Hand
  $panel.BorderStyle = 'FixedSingle'

  $lIcono = New-Object System.Windows.Forms.Label
  $lIcono.Text = $icono
  $lIcono.Font = $fIcono
  $lIcono.ForeColor = $color
  $lIcono.TextAlign = 'MiddleCenter'
  $lIcono.Location = New-Object System.Drawing.Point((S 14), (S 14))
  $lIcono.Size = New-Object System.Drawing.Size((S 34), (S 34))
  $panel.Controls.Add($lIcono)

  $lTexto = New-Object System.Windows.Forms.Label
  $lTexto.Text = $texto
  $lTexto.Font = $fOpcion
  $lTexto.ForeColor = $color
  $lTexto.AutoSize = $true
  $lTexto.Location = New-Object System.Drawing.Point((S 56), (S 11))
  $panel.Controls.Add($lTexto)

  $lDetalle = New-Object System.Windows.Forms.Label
  $lDetalle.Text = $detalle
  $lDetalle.Font = $fDetalle
  $lDetalle.ForeColor = $suave
  $lDetalle.AutoSize = $true
  $lDetalle.Location = New-Object System.Drawing.Point((S 57), (S 34))
  $panel.Controls.Add($lDetalle)

  $lInfo = New-Object System.Windows.Forms.Label
  $lInfo.Text = if ($hayIconos) { [char]0xE946 } else { '?' }
  $lInfo.Font = $fInfo
  $lInfo.ForeColor = $suave
  $lInfo.TextAlign = 'MiddleCenter'
  $lInfo.Location = New-Object System.Drawing.Point((S 566), (S 20))
  $lInfo.Size = New-Object System.Drawing.Size((S 24), (S 24))
  $lInfo.Cursor = [System.Windows.Forms.Cursors]::Help
  $pistas.SetToolTip($lInfo, $ayuda)
  $panel.Controls.Add($lInfo)

  $alClic = { $script:eleccion = $valor; $f.Close() }.GetNewClosure()
  $entrar = { $panel.BackColor = [System.Drawing.Color]::FromArgb(240, 245, 244) }.GetNewClosure()
  $salir  = { $panel.BackColor = $tarjeta }.GetNewClosure()

  foreach ($c in @($panel, $lIcono, $lTexto, $lDetalle)) {
    $c.Add_Click($alClic)
    $c.Add_MouseEnter($entrar)
    $c.Add_MouseLeave($salir)
  }
  $lInfo.Add_MouseEnter($entrar)
  $lInfo.Add_MouseLeave($salir)

  $f.Controls.Add($panel)
}

$i = if ($hayIconos) { @{
  refrescar = [char]0xE72C; nuevo = [char]0xE8A5; carpeta = [char]0xE8E5
  abrir     = [char]0xE785; cerrar = [char]0xE72E
} } else { @{ refrescar=''; nuevo=''; carpeta=''; abrir=''; cerrar='' } }

Nueva-Opcion $i.refrescar 'Actualizar el documento de siempre' `
  'Conserva todo lo que haya escrito. Solo cambia las cifras.' `
  ("Lee el Excel y reescribe unicamente las regiones marcadas del documento:`n" +
   "la tabla del estado, los campos de encabezado y las cifras que haya`n" +
   "intercalado en la redaccion.`n`n" +
   "Su texto no se toca. Si borro un parrafo, sigue borrado.`n`n" +
   "Cierre el documento en Word antes de ejecutarlo.") 122 '1' $acento

Nueva-Opcion $i.nuevo 'Crear un documento nuevo' `
  'Sale de la plantilla, en la carpeta salidas\.' `
  ("Genera un Word desde cero con las cifras del Excel.`n`n" +
   "Es una foto desechable: sirve para una entrega puntual. Lo que`n" +
   "escriba en el NO pasa al siguiente que genere.`n`n" +
   "No toca el documento de siempre.") 200 '2' $tinta

Nueva-Opcion $i.carpeta 'Cambiar el documento que se actualiza' `
  'Abre el explorador para elegir otro documento de Word.' `
  ("Elige que archivo actualiza la opcion 1, y lo recuerda.`n`n" +
   "Comprueba que sea un .docx valido y le avisa si todavia le faltan`n" +
   "las regiones marcadas.`n`n" +
   "Queda guardado en config.json.") 278 '3' $tinta

Nueva-Opcion $i.abrir 'Permitir editar las cifras a mano en Word' `
  'Ojo: lo que teclee lo machaca el siguiente refresco.' `
  ("Quita el candado de las cifras y de la tabla para poder teclear`n" +
   "encima en Word.`n`n" +
   "AVISO: siguen vinculadas al Excel. Lo que escriba a mano`n" +
   "desaparece en el siguiente refresco.`n`n" +
   "Para que un valor escrito a mano sobreviva, hay que desvincularlo:`n" +
   "en Word, clic derecho sobre el recuadro -> Quitar control de contenido.") 356 '4' $ambar

Nueva-Opcion $i.cerrar 'Volver a proteger las cifras' `
  'Word deja de permitir teclear dentro de ellas.' `
  ("Vuelve a poner el candado a la tabla, a los campos de encabezado`n" +
   "y a las cifras intercaladas en la redaccion.`n`n" +
   "OJO: solo protege lo que esta dentro de una region marcada. Un`n" +
   "numero copiado y pegado del Excel como texto normal NO queda`n" +
   "protegido, porque el programa no puede saber que es una cifra.") 434 '5' $tinta

$salir = New-Object System.Windows.Forms.Button
$salir.Text = 'Salir'
$salir.Font = $fSub
$salir.ForeColor = $suave
$salir.BackColor = $fondo
$salir.FlatStyle = 'Flat'
$salir.FlatAppearance.BorderSize = 0
$salir.Location = New-Object System.Drawing.Point((S 540), (S 516))
$salir.Size = New-Object System.Drawing.Size((S 92), (S 34))
$salir.Cursor = [System.Windows.Forms.Cursors]::Hand
$salir.Add_Click({ $script:eleccion = '0'; $f.Close() })
$f.Controls.Add($salir)

if ($captura) {
  $f.Show(); [System.Windows.Forms.Application]::DoEvents()
  $bmp = New-Object System.Drawing.Bitmap($f.Width, $f.Height)
  $f.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0, 0, $f.Width, $f.Height)))
  $bmp.Save($captura, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose(); $f.Close()
  Write-Output "CAPTURA=$captura"
} else {
  [void]$f.ShowDialog()
}
Write-Output ("ELECCION=" + $script:eleccion)
"""


def menu_ventana(destino, libro=""):
    """Muestra el menú en una ventana. Devuelve la opción, o None si no
    se pudo dibujar (entonces el llamante usa el menú de consola)."""
    import subprocess
    import tempfile

    tmp = Path(tempfile.mkdtemp(prefix="fs_ventana_"))
    script = tmp / "ventana.ps1"
    try:
        script.write_text(_PS_VENTANA, encoding="utf-8")
        res = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-STA", "-File", str(script),
             str(destino or ""), str(libro or ""), ""],
            capture_output=True, text=True, timeout=1800,
        )
    except Exception:
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    for linea in (res.stdout or "").splitlines():
        linea = linea.strip()
        if linea.startswith("ELECCION="):
            return linea[len("ELECCION="):].strip()
    return None
def _menu(texto_libro=""):
    destino = _describir_destino()
    _cabecera()
    print()
    print("   1)  ACTUALIZAR el documento de siempre")
    print("       Conserva todo lo que haya escrito. Solo cambia las cifras.")
    print(f"       Documento: {destino}")
    print(f"       Excel:     {texto_libro or 'NINGUNO — arrástrelo sobre el icono'}")
    print()
    print("   2)  CREAR un documento nuevo")
    print("       Sale de la plantilla, en la carpeta salidas\\.")
    print("       Lo que escriba en él no pasa al siguiente.")
    print()
    print("   3)  CAMBIAR el documento que se actualiza")
    print("       Abre el explorador para elegir otro documento de Word.")
    print()
    print("   4)  PERMITIR editar las cifras a mano en Word")
    print("       Ojo: lo que teclee lo machaca el siguiente refresco.")
    print()
    print("   5)  VOLVER A PROTEGER las cifras")
    print("       Word deja de permitir teclear dentro de ellas.")
    print()
    print("   0)  Salir")
    print()
    print("=" * ANCHO)

    while True:
        try:
            eleccion = input(" Escriba una opción (0-5) y pulse Enter: ").strip()
        except (EOFError, KeyboardInterrupt):
            return "0"
        if eleccion in ("0", "1", "2", "3", "4", "5"):
            return eleccion
        print(" No entendí esa opción.")


def cambiar_documento():
    """Elige el documento que se actualizará y lo deja fijado en config.json.

    Abre el explorador de Windows filtrado a documentos de Word, comprueba
    que sirva, y ofrece prepararlo si todavía no tiene las regiones.
    """
    cfg = G.cargar_config()
    try:
        actual = D.resolver_documento(None, cfg)
        carpeta = actual.parent
        print()
        print(f" Documento actual: {actual.name}")
    except ValueError:
        carpeta = None
        print()
        print(" Ahora mismo no hay ningún documento configurado.")

    print(" Abriendo el explorador…")
    elegido = D.elegir_archivo_word(carpeta)
    if elegido is None:
        print()
        print(" No se eligió ninguno. Nada ha cambiado.")
        return 0

    ok, avisos, familias = D.revisar_candidato(elegido)
    print()
    print("=" * ANCHO)
    print(f" {elegido.name}")
    print(f" {elegido.parent}")
    print("=" * ANCHO)

    if not ok:
        print()
        print(" NO SIRVE COMO DOCUMENTO BASE:")
        for a in avisos:
            print(f"   {a}")
        print()
        print(" No se ha cambiado nada.")
        return 1

    if familias.get(D.C.FAM_TABLA):
        print(f"   Ya preparado: tablas={familias.get(D.C.FAM_TABLA, 0)}  "
              f"campos={familias.get(D.C.FAM_CAMPO, 0)}  "
              f"cifras={familias.get(D.C.FAM_DATO, 0)}")
    for a in avisos:
        print(f"   AVISO: {a}")

    D.fijar_documento_base(elegido)
    print()
    print(" Hecho: a partir de ahora la opción 1 actualiza este documento.")

    if not familias.get(D.C.FAM_TABLA):
        print()
        print(" Todavía le faltan las regiones. Para prepararlo, elija la")
        print(" opción 1 con el Excel y añada --preparar, o ejecute:")
        print("     EstadosFinancieros.exe MI_LIBRO.xlsx --refrescar --preparar")
    return 0
def ejecutar(argv):
    D.preparar_consola()
    args = [a for a in argv[1:] if not a.startswith("--")]
    flags = {a.lower() for a in argv[1:] if a.startswith("--")}

    if "--refrescar" in flags:
        eleccion = "1"
    elif "--generar" in flags:
        eleccion = "2"
    elif "--estado" in flags:
        eleccion = "estado"
    elif "--desbloquear" in flags:
        eleccion = "4"
    elif "--bloquear" in flags:
        eleccion = "5"
    elif "--elegir-documento" in flags or (
            "--documento" in flags and _opcion(argv, "--documento") is None):
        # Ojo: refrescar_fs.py usa «--documento OTRO.docx» para apuntar a otro
        # archivo. Antes el menú se quedaba con la bandera siempre, así que a
        # través de EstadosFinancieros.exe era imposible refrescar un
        # documento distinto del de config.json: se abría el explorador.
        # Ahora solo la intercepta cuando viene suelta, sin ruta detrás.
        eleccion = "3"
    elif "--consola" in flags:
        eleccion = _menu()
    else:
        eleccion = None

    libro, texto_libro, aviso_libro = _resolver_libro(args)

    if eleccion is None:
        # Primero la ventana; si el equipo no la puede dibujar, la consola.
        eleccion = menu_ventana(_describir_destino(), texto_libro)
        if eleccion is None:
            eleccion = _menu(texto_libro)

    if eleccion == "0":
        print(" Nada que hacer.")
        return 0

    # Cada modulo analiza sus propios argumentos: aqui solo se le quitan
    # las banderas del menu y se le pasa el resto tal cual.
    propias = {"--refrescar", "--generar", "--estado", "--elegir-documento",
               "--consola",
               "--desbloquear", "--bloquear"}
    if eleccion == "3":
        propias.add("--documento")
    resto = [argv[0]] + [a for a in argv[1:] if a.lower() not in propias]

    if eleccion in ("1", "2"):
        # Las dos opciones que leen cifras. Antes se entraba a ciegas: si no
        # había libro, el error saltaba en mitad del proceso; y si lo había
        # elegido la convención de nombre, no se decía.
        if libro is None:
            raise ValueError(
                (aviso_libro + "\n\n") if aviso_libro else
                "No hay ningún libro de Excel del que leer las cifras.\n\n"
                "Arrastre su .xlsx sobre el icono, o pase su ruta en la orden."
            )
        if aviso_libro:
            print()
            print(" AVISO — " + aviso_libro)
        print()
        print(f" Leyendo las cifras de: {libro.name}")
        print(f" Carpeta:               {libro.parent}")
        if eleccion == "1":
            R.ejecutar(resto)
        else:
            G.ejecutar(resto)
        return 0

    if eleccion == "3":
        return cambiar_documento()

    if eleccion in ("4", "5"):
        cfg = G.cargar_config()
        documento = D.resolver_documento(None, cfg)
        abrir = eleccion == "4"
        D._respaldar(documento)
        print()
        if abrir:
            D.desproteger(documento)
            D.cambiar_candado(documento, bloquear=False)
        else:
            # Candado de region + proteccion de documento. El candado solo no
            # basta: Buscar y reemplazar lo atraviesa y Word en el navegador
            # ni lo mira.
            D.cambiar_candado(documento, bloquear=True)
            D.proteger_salvo_datos(documento, str(cfg.get("clave_proteccion") or "fs"))
        print()
        if abrir:
            print(" Ya puede teclear encima de las cifras en Word.")
            print(" AVISO: lo que escriba a mano lo MACHACA el siguiente refresco.")
            print("        Para que un valor a mano sobreviva, hay que desvincularlo")
            print("        (clic derecho sobre el recuadro en Word -> Quitar control")
            print("        de contenido).")
        else:
            print(" Las cifras son intocables: ni tecleando, ni con Buscar y")
            print(" reemplazar, ni desde Word en el navegador.")
            print(" La redacción sigue libre en todo el documento.")
        return 0

    # Opción no listada en el menú: diagnóstico para quien da soporte.
    return D.estado(G.cargar_config(), args[0] if args else None)


def main():
    try:
        ejecutar(sys.argv)
    except ValueError as e:
        print()
        print("=" * ANCHO)
        print(" NO SE PUDO COMPLETAR")
        print("=" * ANCHO)
        print(str(e))
        print("=" * ANCHO)
        R._pausa()
        sys.exit(1)
    except Exception:
        print()
        print("=" * ANCHO)
        print(" OCURRIÓ UN ERROR INESPERADO — copie este texto para soporte")
        print("=" * ANCHO)
        traceback.print_exc()
        print("=" * ANCHO)
        R._pausa()
        sys.exit(1)
    else:
        R._pausa()


if __name__ == "__main__":
    main()
