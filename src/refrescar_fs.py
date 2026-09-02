"""
refrescar_fs.py — La vía fácil para actualizar el documento base.

Es la contraparte de generador_fs.py:

    generador_fs.py   crea un Word NUEVO en salidas\\      (no toca nada más)
    refrescar_fs.py   ACTUALIZA el documento que ya existe (conserva el texto)

Pensado para doble clic o para arrastrarle el Excel encima. Toma el
documento de config.json -> "documento_base" (el que vive en OneDrive) y le
refresca las cifras. Todo lo que alguien haya redactado se queda como está.

Uso
---
    Arrastre su Excel sobre RefrescarFS.exe  (o sobre refrescar.bat)
    Doble clic sin arrastrar nada            (busca el Excel por convención)

    python refrescar_fs.py [libro.xlsx] [--documento otro.docx] [--no-preparar]

Si al documento le faltan las regiones, se le añaden solas antes de
refrescar: en blanco se usa de base, y con redacción encima el estado entra
como un apartado aparte, sin tocar el texto que ya hubiera.

    --preparar      redundante (es lo que se hace ya). Se acepta por
                    compatibilidad con las órdenes que lo llevan.
    --no-preparar   no añadir nada: si faltan las regiones, no se escribe.
"""
import sys
import hashlib
import traceback
from pathlib import Path

_AQUI = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
if str(_AQUI) not in sys.path:
    sys.path.insert(0, str(_AQUI))

import generador_fs as G
import fs_documento as D


def _pausa():
    """Espera Enter solo si hay consola interactiva (doble clic / .bat)."""
    try:
        if sys.stdin and sys.stdin.isatty():
            input("\nPresione Enter para cerrar...")
    except (EOFError, OSError):
        pass


def ejecutar(argv):
    D.preparar_consola()
    args = [a for a in argv[1:] if not a.startswith("--")]
    flags = {a.lower() for a in argv[1:] if a.startswith("--")}

    def opcion(nombre):
        for i, a in enumerate(argv):
            if a.lower() == nombre and i + 1 < len(argv):
                return argv[i + 1]
        return None

    cfg = G.cargar_config()
    documento = D.resolver_documento(opcion("--documento"), cfg)

    # Lo primero, antes de leer el Excel siquiera: si Word tiene el
    # documento abierto, no se toca nada. Escribir encima lo destruiría.
    D.comprobar_escribible(documento)

    xlsx = Path(args[0]).resolve() if args else G.encontrar_excel_por_convencion(cfg)
    if not args:
        print(f"(Sin archivo indicado: usando '{xlsx.name}' por convención de nombre)")
    if not xlsx.exists():
        raise ValueError(f"No se encontró el libro de Excel:\n  {xlsx}")

    ctx = G.leer_contexto(xlsx, cfg)
    meta = ctx.pop("_meta", {})
    ctx.pop("_avisos", None)

    # La copia .bak va lo primero: es la de ANTES de tocar nada, que es la
    # única que sirve para deshacer. Estaba después de preparar el
    # documento, así que respaldaba el resultado en vez del original.
    D._respaldar(documento)

    # Un documento sin las regiones del contrato no tiene dónde recibir las
    # cifras: el refresco recorría sus anclas, no encontraba ninguna y
    # terminaba con «no se actualizó NADA», dejando al usuario con un
    # documento intacto y ninguna forma evidente de arreglarlo desde la
    # ventana (--preparar solo existía en la línea de órdenes).
    #
    # Ahora se prepara solo. El documento no pierde nada: lo que ya
    # estuviera escrito se respeta, y si traía redacción propia el estado
    # entra como un apartado aparte. Con --no-preparar se recupera el
    # comportamiento de antes.
    integrado = D.clasificar_documento(documento)[0] == D.LISTO
    if "--preparar" in flags or (not integrado and "--no-preparar" not in flags):
        print()
        print(D.imprimible("Preparando el documento (solo añade lo que falte)…"))
        D.preparar(documento, ctx, cfg, respaldar=False)

    sha = hashlib.sha256(xlsx.read_bytes()).hexdigest()[:12]
    inf = D.refrescar(documento, ctx, origen=f"{xlsx.name} (sha {sha})", cfg=cfg)

    total_filas = sum(n for _, n in inf["tablas"])
    print()
    print("=" * 68)
    print(" DOCUMENTO ACTUALIZADO")
    print("=" * 68)
    print(f" Documento:      {documento.name}")
    print(f" Carpeta:        {documento.parent}")
    print(f" Origen:         {xlsx.name}  sha256 {sha}")
    print(f" Hoja usada:     {meta.get('hoja', '?')}")
    print()
    print(f" Filas en tablas:      {total_filas}")
    print(f" Campos actualizados:  {inf['campos']}")
    print(f" Cifras en el texto:   {inf['datos']}")
    print(f" Zonas de redacción intactas: {inf['sin_ancla_prosa']}")

    if not inf["tablas"] and not inf["campos"] and not inf["datos"]:
        print()
        print(" AVISO — no se actualizó NADA, y eso no debería pasar: el")
        print("         documento se prepara solo antes de refrescarlo.")
        print("         Si lanzó la orden con --no-preparar, quítelo.")

    if inf["huerfanos"]:
        print()
        print(" AVISO — el documento pide cifras que el Excel ya no tiene:")
        for t in inf["huerfanos"]:
            print(f"   ? {t}")

    print()
    if inf.get("bitacora_archivo"):
        print(f" Bitácora:       {inf['bitacora_archivo']}")
    print()
    print(" Cambios respecto de la última actualización:")
    for c in inf["cambios"][:20]:
        print(D.imprimible(f"   • {c}"))
    if len(inf["cambios"]) > 20:
        print(D.imprimible(f"   • … y {len(inf['cambios']) - 20} más (están en la bitácora del documento)."))
    print("=" * 68)
    # Para que fs_menu pueda ofrecer «Abrir el documento» al terminar.
    return documento


def main():
    try:
        ejecutar(sys.argv)
    except ValueError as e:
        print()
        print("=" * 68)
        print(" NO SE PUDO ACTUALIZAR EL DOCUMENTO")
        print("=" * 68)
        print(str(e))
        print("=" * 68)
        _pausa()
        sys.exit(1)
    except PermissionError as e:
        # Puede ser el .docx retenido por Word o el .xlsx retenido por Excel:
        # decir «Word» a secas mandaba a cerrar el programa equivocado.
        print()
        print("=" * 68)
        print(" UN ARCHIVO ESTÁ ABIERTO EN OTRO PROGRAMA")
        print("=" * 68)
        print(f" {e}")
        print()
        print(" Ciérrelo (Word si es el documento, Excel si es el libro) y")
        print(" vuelva a intentarlo.")
        print("=" * 68)
        _pausa()
        sys.exit(1)
    except Exception:
        print()
        print("=" * 68)
        print(" OCURRIÓ UN ERROR INESPERADO — copie este texto para soporte")
        print("=" * 68)
        traceback.print_exc()
        print("=" * 68)
        _pausa()
        sys.exit(1)
    else:
        _pausa()


if __name__ == "__main__":
    main()
