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
"""
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
    print("   3)  VER el estado del proyecto")
    print("       Qué hay montado, qué falta y qué está bloqueando.")
    print()
    print("   0)  Salir")
    print()
    print("=" * ANCHO)

    while True:
        try:
            eleccion = input(" Escriba 1, 2, 3 o 0 y pulse Enter: ").strip()
        except (EOFError, KeyboardInterrupt):
            return "0"
        if eleccion in ("0", "1", "2", "3"):
            return eleccion
        print(" No entendí esa opción.")


def ejecutar(argv):
    D.preparar_consola()
    args = [a for a in argv[1:] if not a.startswith("--")]
    flags = {a.lower() for a in argv[1:] if a.startswith("--")}

    if "--refrescar" in flags:
        eleccion = "1"
    elif "--generar" in flags:
        eleccion = "2"
    elif "--estado" in flags:
        eleccion = "3"
    else:
        eleccion = _menu()

    if eleccion == "0":
        print(" Nada que hacer.")
        return 0

    # Cada modulo analiza sus propios argumentos: aqui solo se le quitan
    # las banderas del menu y se le pasa el resto tal cual.
    propias = {"--refrescar", "--generar", "--estado"}
    resto = [argv[0]] + [a for a in argv[1:] if a.lower() not in propias]

    if eleccion == "1":
        R.ejecutar(resto)
        return 0

    if eleccion == "2":
        G.ejecutar(resto)
        return 0

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
