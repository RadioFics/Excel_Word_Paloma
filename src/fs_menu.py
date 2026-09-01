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


#: Ventana del menú. Se dibuja con WinForms desde PowerShell porque el
#: Python portable no trae tkinter y meterlo obligaría a arrastrar Tcl/Tk
#: dentro del .exe (unos 10 MB más y un empaquetado más frágil). PowerShell
#: y .NET están en cualquier Windows, así que no añade ninguna dependencia.
#: Si algo falla, se cae al menú de consola de siempre.
_PS_VENTANA = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

$documento = if ($args.Count -ge 1) { $args[0] } else { '' }

$tinta   = [System.Drawing.Color]::FromArgb(26, 34, 38)
$suave   = [System.Drawing.Color]::FromArgb(88, 101, 107)
$fondo   = [System.Drawing.Color]::FromArgb(250, 249, 246)
$acento  = [System.Drawing.Color]::FromArgb(20, 92, 88)
$aviso   = [System.Drawing.Color]::FromArgb(138, 90, 18)

$f = New-Object System.Windows.Forms.Form
$f.Text = 'Estados Financieros'
$f.Size = New-Object System.Drawing.Size(620, 560)
$f.StartPosition = 'CenterScreen'
$f.BackColor = $fondo
$f.FormBorderStyle = 'FixedDialog'
$f.MaximizeBox = $false

$titulo = New-Object System.Windows.Forms.Label
$titulo.Text = 'Estados Financieros'
$titulo.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 17)
$titulo.ForeColor = $tinta
$titulo.Location = New-Object System.Drawing.Point(28, 22)
$titulo.Size = New-Object System.Drawing.Size(560, 34)
$f.Controls.Add($titulo)

$sub = New-Object System.Windows.Forms.Label
$sub.Text = if ($documento) { "Documento: $documento" } else { 'Sin documento configurado (use "Cambiar documento")' }
$sub.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$sub.ForeColor = $suave
$sub.Location = New-Object System.Drawing.Point(30, 58)
$sub.Size = New-Object System.Drawing.Size(560, 20)
$f.Controls.Add($sub)

$script:eleccion = '0'

function Nuevo-Boton($texto, $detalle, $y, $valor, $color) {
  $b = New-Object System.Windows.Forms.Button
  $b.Text = "  $texto"
  $b.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
  $b.ForeColor = $color
  $b.BackColor = [System.Drawing.Color]::White
  $b.FlatStyle = 'Flat'
  $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(222, 219, 211)
  $b.FlatAppearance.BorderSize = 1
  $b.TextAlign = 'MiddleLeft'
  $b.Location = New-Object System.Drawing.Point(28, $y)
  $b.Size = New-Object System.Drawing.Size(556, 40)
  $b.Cursor = [System.Windows.Forms.Cursors]::Hand
  $b.Add_Click({ $script:eleccion = $valor; $f.Close() }.GetNewClosure())
  $f.Controls.Add($b)

  $l = New-Object System.Windows.Forms.Label
  $l.Text = $detalle
  $l.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
  $l.ForeColor = $suave
  $l.Location = New-Object System.Drawing.Point(34, ($y + 42))
  $l.Size = New-Object System.Drawing.Size(556, 18)
  $f.Controls.Add($l)
}

Nuevo-Boton 'Actualizar el documento de siempre' `
            'Conserva todo lo que haya escrito. Solo cambia las cifras.' 96  '1' $acento
Nuevo-Boton 'Crear un documento nuevo' `
            'Sale de la plantilla, en la carpeta salidas\.' 162 '2' $tinta
Nuevo-Boton 'Cambiar el documento que se actualiza' `
            'Abre el explorador para elegir otro documento de Word.' 228 '3' $tinta
Nuevo-Boton 'Permitir editar las cifras a mano en Word' `
            'Ojo: lo que teclee lo machaca el siguiente refresco.' 294 '4' $aviso
Nuevo-Boton 'Volver a proteger las cifras' `
            'Word deja de permitir teclear dentro de ellas.' 360 '5' $tinta

$salir = New-Object System.Windows.Forms.Button
$salir.Text = 'Salir'
$salir.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$salir.ForeColor = $suave
$salir.BackColor = $fondo
$salir.FlatStyle = 'Flat'
$salir.FlatAppearance.BorderSize = 0
$salir.Location = New-Object System.Drawing.Point(492, 430)
$salir.Size = New-Object System.Drawing.Size(92, 32)
$salir.Add_Click({ $script:eleccion = '0'; $f.Close() })
$f.Controls.Add($salir)

[void]$f.ShowDialog()
Write-Output ("ELECCION=" + $script:eleccion)
"""


def menu_ventana(destino):
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
             "-STA", "-File", str(script), str(destino or "")],
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
def _menu():
    destino = _describir_destino()
    _cabecera()
    print()
    print("   1)  ACTUALIZAR el documento de siempre")
    print("       Conserva todo lo que haya escrito. Solo cambia las cifras.")
    print(f"       Documento: {destino}")
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
    elif "--documento" in flags:
        eleccion = "3"
    elif "--consola" in flags:
        eleccion = _menu()
    else:
        # Primero la ventana; si el equipo no la puede dibujar, la consola.
        eleccion = menu_ventana(_describir_destino())
        if eleccion is None:
            eleccion = _menu()

    if eleccion == "0":
        print(" Nada que hacer.")
        return 0

    # Cada modulo analiza sus propios argumentos: aqui solo se le quitan
    # las banderas del menu y se le pasa el resto tal cual.
    propias = {"--refrescar", "--generar", "--estado", "--documento",
               "--consola",
               "--desbloquear", "--bloquear"}
    resto = [argv[0]] + [a for a in argv[1:] if a.lower() not in propias]

    if eleccion == "1":
        R.ejecutar(resto)
        return 0

    if eleccion == "2":
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
        D.cambiar_candado(documento, bloquear=not abrir)
        print()
        if abrir:
            print(" Ya puede teclear encima de las cifras en Word.")
            print(" AVISO: lo que escriba a mano lo MACHACA el siguiente refresco.")
            print("        Para que un valor a mano sobreviva, hay que desvincularlo")
            print("        (clic derecho sobre el recuadro en Word -> Quitar control")
            print("        de contenido).")
        else:
            print(" Las cifras vuelven a ser intocables a mano en Word.")
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
