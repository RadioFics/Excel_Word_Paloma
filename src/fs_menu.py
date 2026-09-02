"""
fs_menu.py — Un solo icono que pregunta qué quiere hacer.

Existe porque los dos caminos se confunden con facilidad: ambos aceptan que
se les arrastre el Excel, pero uno crea un documento nuevo y el otro
actualiza el que ya hay. Aquí se elige explícitamente.

NO duplica lógica: delega en los mismos módulos que usan los ejecutables
individuales, que siguen funcionando por su cuenta.

    generador_fs.ejecutar()   ->  GeneradorFS.exe  /  generar.bat
    refrescar_fs.ejecutar()   ->  RefrescarFS.exe  /  refrescar.bat
    fs_documento.crear_base() ->  la plantilla viva, desde cero
    fs_documento.preparar()   ->  adaptar un Word cualquiera
    fs_documento.estado()     ->  la radiografía del proyecto

Los dos «crear» no son lo mismo y conviene no confundirlos:

    opción 6  crear_base()   -> un Word VIVO, con regiones dentro. Es la
                               base que se refresca cada cierre.
    opción 2  generador_fs   -> una FOTO en salidas\\, renderizada de una
                               plantilla de Word. No lleva regiones, así
                               que no se puede volver a actualizar.

Uso
---
    Arrastre su Excel sobre EstadosFinancieros.exe  (o doble clic)

    python fs_menu.py [libro.xlsx]              menú interactivo
    python fs_menu.py [libro.xlsx] --refrescar   sin menú
    python fs_menu.py [libro.xlsx] --plantilla   crear la base desde cero
    python fs_menu.py [libro.xlsx] --generar     copia desechable, sin menú
    python fs_menu.py --estado                   sin menú
    python fs_menu.py --desbloquear              permite teclear las cifras
    python fs_menu.py --bloquear                 vuelve a protegerlas
    python fs_menu.py --consola                  menú de texto, sin ventana
"""
import contextlib
import hashlib
import io
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


class _Eco(io.TextIOBase):
    """Escribe en la consola de siempre Y en un buffer, a la vez.

    La consola sigue siendo la fuente para quien ejecuta desde una terminal
    o un .bat; el buffer alimenta la ventana de resultado.
    """

    def __init__(self, real):
        self._real = real
        self._buffer = io.StringIO()

    def write(self, s):
        self._buffer.write(s)
        try:
            self._real.write(s)
        except Exception:
            pass          # una consola cerrada no debe tumbar la operacion
        return len(s)

    def flush(self):
        try:
            self._real.flush()
        except Exception:
            pass

    def getvalue(self):
        return self._buffer.getvalue()


@contextlib.contextmanager
def _consola_duplicada():
    eco = _Eco(sys.stdout)
    anterior = sys.stdout
    sys.stdout = eco
    try:
        yield eco
    finally:
        sys.stdout = anterior


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


def _hay_documento():
    """¿Se puede resolver el documento base ahora mismo?"""
    try:
        return D.resolver_documento(None, G.cargar_config()).exists()
    except Exception:
        return False


def _describir_cifras():
    """Estado del candado, en una línea, para la cabecera de la ventana.

    Vacío si no se puede saber: la ventana simplemente no lo muestra.
    """
    try:
        doc = D.resolver_documento(None, G.cargar_config())
        bloqueadas, total, proteccion = D.estado_candado(doc)
    except Exception:
        return ""
    if not total:
        return "el documento aún no tiene regiones de datos"
    if bloqueadas == total:
        extra = "  ·  modo estricto" if proteccion else ""
        return f"PROTEGIDAS ({total} regiones){extra}"
    if bloqueadas == 0:
        return f"EDITABLES a mano ({total} regiones sin candado)"
    return f"MIXTO ({bloqueadas} de {total} con candado)"


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
$cifras    = if ($args.Count -ge 3) { $args[2] } else { '' }
$captura   = if ($args.Count -ge 4) { $args[3] } else { '' }

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
$f.ClientSize = New-Object System.Drawing.Size((S 660), (S 644))
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

# Estado REAL del candado, leido del documento. Sin esta linea, las
# opciones de proteger/desproteger eran dos botones ciegos: no habia forma
# de saber en que estado estaba el documento, asi que pulsarlos no parecia
# tener ningun efecto.
$lCifras = New-Object System.Windows.Forms.Label
$lCifras.Text = if ($cifras) { "Cifras: $cifras" } else { '' }
$lCifras.Font = $fSub
$lCifras.ForeColor = if ($cifras -like 'EDITABLES*') { $ambar } else { $suave }
$lCifras.AutoEllipsis = $true
$lCifras.Location = New-Object System.Drawing.Point((S 32), (S 106))
$lCifras.Size = New-Object System.Drawing.Size((S 600), (S 22))
$f.Controls.Add($lCifras)

# La eleccion vive DENTRO de una tabla hash, no en una variable suelta.
#
# Motivo: los manejadores de clic se crean con .GetNewClosure(), que le da a
# cada bloque su propio ambito capturado. Una asignacion como
# «$script:eleccion = $valor» hecha ahi dentro escribe en ese ambito
# aislado, no en el del script: la ventana se cerraba, pero el valor que se
# leia al final seguia siendo el inicial ('0' = Salir). O sea, TODAS las
# tarjetas del menu acababan comportandose como «Salir», y por eso la
# aplicacion parecia un puñado de botones sin efecto.
#
# Mutar un objeto SI atraviesa el ambito: la closure y el script comparten
# la misma tabla hash. Es lo mismo que hace que el resaltado al pasar el
# raton si funcione ($panel.BackColor).
$estado = @{ valor = '0' }

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

  $alClic = { $estado.valor = $valor; $f.Close() }.GetNewClosure()
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
  abrir     = [char]0xE785; cerrar = [char]0xE72E; plantilla = [char]0xE8F4
} } else { @{ refrescar=''; nuevo=''; carpeta=''; abrir=''; cerrar=''; plantilla='' } }

Nueva-Opcion $i.refrescar 'Actualizar el documento de siempre' `
  'Conserva todo lo que haya escrito. Solo cambia las cifras.' `
  ("Lee el Excel y reescribe unicamente las regiones marcadas del documento:`n" +
   "la tabla del estado, los campos de encabezado y las cifras que haya`n" +
   "intercalado en la redaccion.`n`n" +
   "Si el documento todavia no tiene esas regiones, se las añade solo`n" +
   "antes de volcar las cifras.`n`n" +
   "Su texto no se toca. Si borro un parrafo, sigue borrado.`n`n" +
   "Cierre el documento en Word antes de ejecutarlo.") 140 '1' $acento

Nueva-Opcion $i.plantilla 'Crear la plantilla base desde el Excel' `
  'Un documento vivo, donde usted diga. Se puede actualizar siempre.' `
  ("Crea un Word NUEVO con las cifras del Excel y, dentro, todas las`n" +
   "regiones que hacen falta para poder actualizarlo cada cierre.`n`n" +
   "Le pregunta donde guardarlo: en el disco o en OneDrive, da igual.`n`n" +
   "Al terminar queda fijado como «el documento de siempre», asi que`n" +
   "puede escribir en el y volver a la opcion 1 en el proximo cierre.`n`n" +
   "Es lo que conviene la primera vez. La diferencia con «Crear un`n" +
   "documento nuevo» es que aquel es una foto y este es la base viva.") 218 '6' $acento

Nueva-Opcion $i.carpeta 'Cambiar el documento que se actualiza' `
  'Elija cualquier Word: vacio o ya escrito. Se adapta solo.' `
  ("Elige que archivo actualiza la opcion 1, y lo recuerda.`n`n" +
   "Vale cualquier .docx:`n" +
   "  - si ya esta integrado, se usa tal cual;`n" +
   "  - si esta en blanco, se usa de base;`n" +
   "  - si ya tiene redaccion, se le añade el estado como un apartado`n" +
   "    aparte y su texto no se toca.`n`n" +
   "Queda guardado en config.local.json (solo en este equipo).") 296 '3' $tinta

Nueva-Opcion $i.nuevo 'Crear un documento nuevo (copia desechable)' `
  'Una foto puntual, en salidas\. No se puede actualizar despues.' `
  ("Genera un Word desde cero con las cifras del Excel.`n`n" +
   "Es una foto desechable: sirve para una entrega puntual. NO lleva`n" +
   "regiones dentro, asi que no se puede volver a actualizar, y lo que`n" +
   "escriba en el no pasa al siguiente que genere.`n`n" +
   "Si lo que quiere es una base para trabajar, use «Crear la plantilla`n" +
   "base desde el Excel».") 374 '2' $tinta

# Un solo interruptor, no dos botones. La etiqueta dice que va a PASAR,
# y la linea «Cifras:» de arriba dice como estan AHORA.
if ($cifras -like 'EDITABLES*') {
  Nueva-Opcion $i.cerrar 'Volver a proteger las cifras' `
    'Ahora mismo se pueden teclear a mano. Esto lo impide.' `
    ("Vuelve a poner el candado a la tabla, a los campos de encabezado`n" +
     "y a las cifras intercaladas en la redaccion.`n`n" +
     "OJO: solo protege lo que esta dentro de una region marcada. Un`n" +
     "numero copiado y pegado del Excel como texto normal NO queda`n" +
     "protegido, porque el programa no puede saber que es una cifra.") 452 '5' $tinta
} else {
  Nueva-Opcion $i.abrir 'Permitir editar las cifras a mano en Word' `
    'Ojo: lo que teclee lo machaca el siguiente refresco.' `
    ("Quita el candado de las cifras y de la tabla para poder teclear`n" +
     "encima en Word.`n`n" +
     "AVISO: siguen vinculadas al Excel. Lo que escriba a mano`n" +
     "desaparece en el siguiente refresco.`n`n" +
     "Para que un valor escrito a mano sobreviva, hay que desvincularlo:`n" +
     "en Word, clic derecho sobre el recuadro -> Quitar control de contenido.") 452 '4' $ambar
}

$salir = New-Object System.Windows.Forms.Button
$salir.Text = 'Salir'
$salir.Font = $fSub
$salir.ForeColor = $suave
$salir.BackColor = $fondo
$salir.FlatStyle = 'Flat'
$salir.FlatAppearance.BorderSize = 0
$salir.Location = New-Object System.Drawing.Point((S 540), (S 560))
$salir.Size = New-Object System.Drawing.Size((S 92), (S 34))
$salir.Cursor = [System.Windows.Forms.Cursors]::Hand
$salir.Add_Click({ $estado.valor = '0'; $f.Close() })
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
Write-Output ("ELECCION=" + $estado.valor)
"""


#: Ventana de resultado. Sin esto, la aplicacion cerraba la ventana del
#: menu, hacia el trabajo y volcaba el informe en una consola que aparece
#: DETRAS: desde el punto de vista del usuario, el boton no hacia nada.
#: Aqui se le dice que paso, donde quedo, y se le ofrece abrirlo.
_PS_RESULTADO = r"""
$ErrorActionPreference = 'Stop'
$codigo = @"
using System;
using System.Runtime.InteropServices;
public static class PppRes {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern int SetProcessDpiAwarenessContext(IntPtr v);
}
"@
try { Add-Type -TypeDefinition $codigo -ErrorAction Stop } catch {}
try   { [void][PppRes]::SetProcessDpiAwarenessContext([IntPtr](-4)) }
catch { try { [void][PppRes]::SetProcessDPIAware() } catch {} }

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

$titulo  = if ($args.Count -ge 1) { $args[0] } else { 'Listo' }
# El informe llega por ARCHIVO, no como argumento: un texto de varias
# lineas pasado en la linea de ordenes se parte por los saltos y powershell
# lo recibe troceado en $args, con lo que la ventana no llegaba a abrirse.
$rutaCuerpo = if ($args.Count -ge 2) { $args[1] } else { '' }
$cuerpo = ''
if ($rutaCuerpo -and (Test-Path -LiteralPath $rutaCuerpo)) {
  $cuerpo = Get-Content -LiteralPath $rutaCuerpo -Raw -Encoding UTF8
}
$destino = if ($args.Count -ge 3) { $args[2] } else { '' }
$ok      = if ($args.Count -ge 4) { $args[3] -ne '0' } else { $true }

$g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
$k = $g.DpiX / 96.0
$g.Dispose()
function S($n) { [int][Math]::Round($n * $k) }

$tinta  = [System.Drawing.Color]::FromArgb(26, 34, 38)
$suave  = [System.Drawing.Color]::FromArgb(102, 114, 120)
$fondo  = [System.Drawing.Color]::FromArgb(247, 246, 243)
$acento = [System.Drawing.Color]::FromArgb(17, 94, 89)
$rojo   = [System.Drawing.Color]::FromArgb(155, 44, 44)

$f = New-Object System.Windows.Forms.Form
$f.Text = 'Estados Financieros'
$f.ClientSize = New-Object System.Drawing.Size((S 620), (S 420))
$f.StartPosition = 'CenterScreen'
$f.BackColor = $fondo
$f.FormBorderStyle = 'FixedSingle'
$f.MaximizeBox = $false
$f.ShowIcon = $false
$f.TopMost = $true

$lTitulo = New-Object System.Windows.Forms.Label
$lTitulo.Text = $titulo
$lTitulo.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$lTitulo.ForeColor = if ($ok) { $acento } else { $rojo }
$lTitulo.AutoSize = $true
$lTitulo.Location = New-Object System.Drawing.Point((S 26), (S 22))
$f.Controls.Add($lTitulo)

$caja = New-Object System.Windows.Forms.TextBox
$caja.Multiline = $true
$caja.ReadOnly = $true
$caja.ScrollBars = 'Vertical'
$caja.BorderStyle = 'FixedSingle'
$caja.BackColor = [System.Drawing.Color]::White
$caja.ForeColor = $tinta
$caja.Font = New-Object System.Drawing.Font('Consolas', 9.5, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$caja.Text = $cuerpo -replace "`n", "`r`n"
$caja.Location = New-Object System.Drawing.Point((S 26), (S 62))
$caja.Size = New-Object System.Drawing.Size((S 568), (S 272))
$f.Controls.Add($caja)

function Nuevo-Boton($texto, $x, $accion, $ancho) {
  $b = New-Object System.Windows.Forms.Button
  $b.Text = $texto
  $b.ForeColor = $tinta
  $b.BackColor = [System.Drawing.Color]::White
  $b.FlatStyle = 'Flat'
  $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(224, 221, 214)
  $b.Location = New-Object System.Drawing.Point((S $x), (S 350))
  $b.Size = New-Object System.Drawing.Size((S $ancho), (S 36))
  $b.Cursor = [System.Windows.Forms.Cursors]::Hand
  $b.Add_Click($accion)
  $f.Controls.Add($b)
  return $b
}

if ($destino -and (Test-Path -LiteralPath $destino)) {
  [void](Nuevo-Boton 'Abrir el documento' 26 { Start-Process $destino }.GetNewClosure() 150)
  [void](Nuevo-Boton 'Abrir la carpeta'  186 {
    Start-Process explorer.exe -ArgumentList ('/select,"' + $destino + '"')
  }.GetNewClosure() 150)
}

$cerrar = Nuevo-Boton 'Cerrar' 502 { $f.Close() } 92
$f.AcceptButton = $cerrar

# El foco va al boton, no a la caja: si se queda en la caja, Windows
# selecciona todo el informe y sale en azul, como si estuviera resaltado.
$f.Add_Shown({
  $cerrar.Focus()
  $caja.SelectionStart = 0
  $caja.SelectionLength = 0
})
[void]$f.ShowDialog()
"""


def _lanzar_ps(script_texto, *argumentos, timeout=1800):
    """Ejecuta un script de ventana. Devuelve (stdout, stderr) o None.

    Delega en D.ejecutar_ps, que es quien sabe escribir el .ps1 con BOM y
    leer la salida en UTF-8 en vez de con la página de códigos de la
    consola. No se duplica aquí para que el arreglo del juego de
    caracteres valga en un solo sitio.
    """
    try:
        stdout, stderr, _ = D.ejecutar_ps(script_texto, *argumentos,
                                          timeout=timeout)
        return stdout, stderr
    except Exception:
        return None


def menu_ventana(destino, libro="", cifras=""):
    """Muestra el menú en una ventana. Devuelve la opción, o None si no
    se pudo dibujar (entonces el llamante usa el menú de consola)."""
    salida = _lanzar_ps(_PS_VENTANA, destino, libro, cifras, "")
    if salida is None:
        return None
    for linea in salida[0].splitlines():
        linea = linea.strip()
        if linea.startswith("ELECCION="):
            return linea[len("ELECCION="):].strip()
    return None


def ventana_resultado(titulo, cuerpo, destino=None, ok=True):
    """Enseña el resultado en una ventana, con botones para abrirlo.

    Si no se puede dibujar, no pasa nada: el mismo informe ya se imprimió
    en la consola.
    """
    import tempfile

    tmp = Path(tempfile.mkdtemp(prefix="fs_res_"))
    try:
        # Por archivo: un informe de varias lineas no cabe en un argumento
        # de linea de ordenes sin partirse.
        cuerpo_txt = tmp / "informe.txt"
        cuerpo_txt.write_text(cuerpo or "", encoding="utf-8")
        _lanzar_ps(_PS_RESULTADO, titulo, cuerpo_txt, destino or "",
                   "1" if ok else "0", timeout=1800)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
def _menu(texto_libro=""):
    destino = _describir_destino()
    _cabecera()
    print()
    print("   1)  ACTUALIZAR el documento de siempre")
    print("       Conserva todo lo que haya escrito. Solo cambia las cifras.")
    print("       Si le faltan las regiones, se las añade antes de volcarlas.")
    print(f"       Documento: {destino}")
    print(f"       Excel:     {texto_libro or 'NINGUNO — arrástrelo sobre el icono'}")
    print()
    print("   6)  CREAR LA PLANTILLA BASE desde el Excel")
    print("       Un documento vivo, donde usted diga (disco u OneDrive).")
    print("       Queda fijado como el documento que se actualiza.")
    print()
    print("   3)  CAMBIAR el documento que se actualiza")
    print("       Vale cualquier Word: vacío o ya escrito. Se adapta solo.")
    print()
    print("   2)  CREAR un documento nuevo (copia desechable)")
    print("       Sale de la plantilla, en la carpeta salidas\\.")
    print("       No lleva regiones: no se puede actualizar después.")
    print()
    cifras = _describir_cifras()
    print("   4)  PERMITIR editar las cifras a mano en Word")
    print("       Ojo: lo que teclee lo machaca el siguiente refresco.")
    print()
    print("   5)  VOLVER A PROTEGER las cifras")
    print("       Word deja de permitir teclear dentro de ellas.")
    if cifras:
        print(f"       Ahora mismo: {cifras}")
    print()
    print("   0)  Salir")
    print()
    print("=" * ANCHO)

    while True:
        try:
            eleccion = input(" Escriba una opción (0-6) y pulse Enter: ").strip()
        except (EOFError, KeyboardInterrupt):
            return "0"
        if eleccion in ("0", "1", "2", "3", "4", "5", "6"):
            return eleccion
        print(" No entendí esa opción.")


def _contexto_del_libro(libro):
    """Las cifras del Excel, o None si no hay libro del que sacarlas.

    Se usa para poder montar el andamiaje YA relleno. Sin libro, el
    andamiaje se monta igual, solo que con los huecos a la vista.
    """
    if libro is None:
        return None, None
    cfg = G.cargar_config()
    ctx = G.leer_contexto(libro, cfg)
    ctx.pop("_meta", None)
    ctx.pop("_avisos", None)
    return ctx, cfg


def cambiar_documento(libro=None):
    """Elige el documento que se actualizará y lo deja listo para trabajar.

    Abre el explorador de Windows, mira cómo viene el documento y actúa en
    consecuencia. Las tres formas de venir valen:

      ya integrado  -> se fija y no se toca nada más
      en blanco     -> se usa de base y se le monta el estado encima
      con redacción -> se le añade el estado como apartado aparte

    Antes solo servía la primera, y las otras dos salían por pantalla como
    «NO SIRVE COMO DOCUMENTO BASE» o se quedaban a medias: el archivo se
    fijaba, pero el refresco siguiente no encontraba dónde escribir y
    terminaba con «no se actualizó NADA».
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

    ok, avisos, _familias, info = D.revisar_candidato(elegido)
    # revisar_candidato puede haber encontrado el archivo bajo un nombre
    # ligeramente distinto (espacios duros, tildes). Se sigue con ESE.
    elegido = info["ruta"]

    print()
    print("=" * ANCHO)
    print(f" {elegido.name}")
    print(f" {elegido.parent}")
    print("=" * ANCHO)

    if not ok:
        print()
        print(" NO SE PUEDE USAR ESTE ARCHIVO:")
        for a in avisos:
            print(f"   {a}")
        print()
        print(" No se ha cambiado nada.")
        return 1

    for a in avisos:
        print(f"   {a}")

    D.fijar_documento_base(elegido)

    if info["estado"] != D.LISTO:
        ctx, cfg_ctx = _contexto_del_libro(libro)
        print()
        print(" Preparándolo para que se le puedan volcar las cifras…")
        if ctx is None:
            print(" (Sin Excel a mano: se montan las regiones vacías. La")
            print("  primera actualización las rellena.)")
        D.preparar(elegido, ctx, cfg_ctx or cfg)

    print()
    print(" Hecho: a partir de ahora «Actualizar» trabaja sobre este documento.")
    return 0


def crear_plantilla(libro):
    """Crea desde cero el documento base, donde el usuario diga, y lo fija.

    Es lo que faltaba: «Crear un documento nuevo» produce una copia
    desechable en salidas\\ —renderizada de una plantilla de Word, sin
    regiones— que no se puede volver a actualizar nunca. Esto produce el
    documento VIVO, con todas las regiones dentro, y deja que el usuario
    elija dónde vive: en el disco o en OneDrive, da igual.
    """
    if libro is None:
        raise ValueError(
            "Para crear la plantilla hacen falta las cifras de un Excel.\n\n"
            "Arrastre su .xlsx sobre el icono y vuelva a elegir esta opción."
        )

    print()
    print(f" Leyendo las cifras de: {libro.name}")
    ctx, cfg = _contexto_del_libro(libro)

    sugerido = f"Estados financieros - {G.sanear(ctx.get('fecha_actual') or 'base')}.docx"
    print(" Elija dónde guardarla (local o en OneDrive)…")
    destino = D.elegir_destino_word(nombre_sugerido=sugerido)
    if destino is None:
        print()
        print(" No se eligió destino. No se ha creado nada.")
        return 0, None

    print()
    print(f" Creando: {destino.name}")
    print(f" En:      {destino.parent}")
    print()
    D.crear_base(destino, ctx, cfg)

    # Recién creada ya lleva las cifras dentro (construir las escribe al
    # montar las regiones), pero se refresca igual: así queda la foto de
    # metadatos y la bitácora, que es lo que hace comparables los refrescos
    # siguientes.
    sha = hashlib.sha256(libro.read_bytes()).hexdigest()[:12]
    D.refrescar(destino, ctx, origen=f"{libro.name} (sha {sha})", cfg=cfg)

    D.fijar_documento_base(destino)
    print()
    print("=" * ANCHO)
    print(" PLANTILLA CREADA")
    print("=" * ANCHO)
    print(f" Archivo:  {destino.name}")
    print(f" Carpeta:  {destino.parent}")
    print()
    print(" Ya está fijada como el documento que se actualiza: escriba en ella")
    print(" lo que quiera y use «Actualizar» cada cierre. Su redacción no se")
    print(" pierde; solo cambian las cifras.")
    print("=" * ANCHO)
    return 0, destino
def ejecutar(argv):
    D.preparar_consola()
    args = [a for a in argv[1:] if not a.startswith("--")]
    flags = {a.lower() for a in argv[1:] if a.startswith("--")}

    if "--refrescar" in flags:
        eleccion = "1"
    elif "--generar" in flags:
        eleccion = "2"
    elif "--plantilla" in flags or "--crear-base" in flags:
        eleccion = "6"
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

    # ¿Vino la orden de la ventana, o de una bandera? Solo la primera debe
    # terminar en una ventana de resultado: si «--bloquear» abriera una,
    # cualquier uso desde un .bat o una tarea programada se quedaria
    # colgado esperando a que alguien pulse «Cerrar».
    interactivo = eleccion is None

    if eleccion is None:
        # Arranque guiado: sin documento base, casi todas las opciones fallan
        # con el mismo error, asi que se resuelve eso primero. Ya no se
        # fuerza el explorador: con un Excel a mano, crear la plantilla
        # desde cero es lo que quiere quien empieza, y antes ni siquiera
        # existia como camino.
        if not _hay_documento():
            print()
            print(" No hay un documento de Word que actualizar todavia.")
            if libro is not None:
                print(" Puede crear la plantilla base desde el Excel (opción 6)")
                print(" o elegir un documento que ya tenga (opción 3).")
            else:
                print(" Eliga uno para empezar (queda guardado en config.local.json).")
                cambiar_documento(libro)

        # Primero la ventana; si el equipo no la puede dibujar, la consola.
        eleccion = menu_ventana(_describir_destino(), texto_libro,
                                _describir_cifras())
        if eleccion is None:
            eleccion = _menu(texto_libro)

    if eleccion == "0":
        print(" Nada que hacer.")
        return 0

    # Cada modulo analiza sus propios argumentos: aqui solo se le quitan
    # las banderas del menu y se le pasa el resto tal cual.
    propias = {"--refrescar", "--generar", "--estado", "--elegir-documento",
               "--consola", "--plantilla", "--crear-base",
               "--desbloquear", "--bloquear"}
    if eleccion == "3":
        propias.add("--documento")
    resto = [argv[0]] + [a for a in argv[1:] if a.lower() not in propias]

    if eleccion == "6":
        if libro is None:
            raise ValueError(
                (aviso_libro + "\n\n") if aviso_libro else
                "Para crear la plantilla hacen falta las cifras de un Excel.\n\n"
                "Arrastre su .xlsx sobre el icono, o pase su ruta en la orden."
            )
        if aviso_libro:
            print()
            print(" AVISO — " + aviso_libro)
        with _consola_duplicada() as eco:
            codigo, destino = crear_plantilla(libro)
        if interactivo:
            ventana_resultado(
                "Plantilla creada" if destino else "No se creó nada",
                eco.getvalue().strip(), destino, ok=True)
        return codigo

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

        # Se hace el trabajo con la consola duplicada a un buffer: lo mismo
        # que se imprime se enseña luego en una ventana. Sin esto el informe
        # acababa en una consola que aparece detras de todo, y la operacion
        # parecia no haber hecho nada.
        with _consola_duplicada() as eco:
            if eleccion == "1":
                destino = R.ejecutar(resto)
                titulo = "Documento actualizado"
            else:
                destino = G.ejecutar(resto)
                titulo = "Documento nuevo creado"
        if interactivo:
            ventana_resultado(titulo, eco.getvalue().strip(), destino, ok=True)
        return 0

    if eleccion == "3":
        # Con el libro delante: si el documento elegido no trae las regiones,
        # se le montan YA rellenas, en vez de dejarlas con los huecos a la
        # vista hasta la primera actualizacion.
        with _consola_duplicada() as eco:
            codigo = cambiar_documento(libro)
        if interactivo:
            try:
                destino = D.resolver_documento(None, G.cargar_config())
            except Exception:
                destino = None
            ventana_resultado("Documento cambiado" if codigo == 0 else "No se cambió",
                              eco.getvalue().strip(), destino, ok=(codigo == 0))
        return codigo

    if eleccion in ("4", "5"):
        cfg = G.cargar_config()
        # El documento puede venir en la orden: «--documento OTRO.docx».
        # Antes estaba fijado a None, asi que estas dos opciones solo sabian
        # operar sobre el de config.json.
        documento = D.resolver_documento(_opcion(argv, "--documento"), cfg)
        abrir = eleccion == "4"
        D._respaldar(documento)
        with _consola_duplicada() as eco:
            print(f" Documento: {documento.name}")
            print(f" Carpeta:   {documento.parent}")
            print()
            if abrir:
                D.desproteger(documento)
                D.cambiar_candado(documento, bloquear=False)
            else:
                # Candado de region + proteccion de documento. El candado solo
                # no basta: Buscar y reemplazar lo atraviesa y Word en el
                # navegador ni lo mira.
                D.cambiar_candado(documento, bloquear=True)
                D.proteger_salvo_datos(
                    documento, str(cfg.get("clave_proteccion") or "fs"))
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
        if interactivo:
            ventana_resultado(
                "Cifras editables a mano" if abrir else "Cifras protegidas",
                eco.getvalue().strip(), documento, ok=True)
        return 0

    # Opción no listada en el menú: diagnóstico para quien da soporte.
    return D.estado(G.cargar_config(), args[0] if args else None)


#: Banderas que piden una acción concreta sin pasar por el menú.
_BANDERAS_DIRECTAS = {
    "--refrescar", "--generar", "--plantilla", "--crear-base", "--estado",
    "--desbloquear", "--bloquear", "--elegir-documento", "--consola",
}


def _sin_menu(argv):
    """¿Se pidió una acción concreta por bandera, sin pasar por la ventana?

    Importa para los errores: si la orden viene de un .bat o de una tarea
    programada, abrir una ventana la dejaría colgada esperando a que
    alguien pulse «Cerrar».
    """
    return any(a.lower() in _BANDERAS_DIRECTAS for a in argv[1:])


def _fallo(titulo, cuerpo, argv):
    """Cuenta lo que ha fallado por los dos sitios: consola y ventana.

    La consola sola no basta: aparece DETRÁS de todo, así que quien pulsó un
    botón ve cerrarse la ventana y nada más. El error que de verdad importa
    —el libro abierto en Excel, el documento abierto en Word— acababa en un
    sitio donde nadie lo lee.
    """
    print()
    print("=" * ANCHO)
    print(f" {titulo}")
    print("=" * ANCHO)
    print(cuerpo)
    print("=" * ANCHO)
    if not _sin_menu(argv):
        ventana_resultado(titulo, cuerpo, None, ok=False)


def main():
    try:
        ejecutar(sys.argv)
    except ValueError as e:
        _fallo("NO SE PUDO COMPLETAR", str(e), sys.argv)
        R._pausa()
        sys.exit(1)
    except PermissionError as e:
        # Un archivo retenido por Word o Excel. Los caminos previstos ya lo
        # explican antes de llegar aquí; esto es la red por si alguno se
        # escapa, para que no salga como un volcado.
        _fallo("UN ARCHIVO ESTÁ ABIERTO EN OTRO PROGRAMA",
               f"{e}\n\nCierre el documento en Word y el libro en Excel, y\n"
               f"vuelva a intentarlo.", sys.argv)
        R._pausa()
        sys.exit(1)
    except Exception:
        _fallo("OCURRIÓ UN ERROR INESPERADO — copie este texto para soporte",
               traceback.format_exc(), sys.argv)
        R._pausa()
        sys.exit(1)
    else:
        R._pausa()


if __name__ == "__main__":
    main()
