"""
fs_contrato.py — Vocabulario de anclas entre el Excel y el documento de Word.

Este módulo NO toca archivos. Define el contrato: qué nombres de ancla
existen, cómo se deriva la clave de una etiqueta del Excel y cómo se
convierte un contexto (el que produce generador_fs.leer_contexto) en el
diccionario plano de valores que el motor de documento escribe región por
región.

Es la misma especificación que consume el add-in de Word. Ver CONTRATO.md.

Familias de ancla
-----------------
  fs-tabla-<nombre>        región de bloque: una tabla completa
  fs-campo-<nombre>        región en línea: un campo de encabezado
  fs-dato-<clave>-<campo>  región en línea: una cifra suelta dentro de prosa
  fs-prosa-<nombre>        región LIBRE (no se toca al refrescar)
  fs-registro              bloque de bitácora (se antepone, no se reemplaza)
  fs-meta                  foto oculta de la última actualización (JSON)

Regla de oro: el refresco SOLO escribe dentro de anclas de las familias
tabla / campo / dato / registro / meta. Todo lo demás del documento —
incluida la prosa que la persona redacte— no se visita siquiera.
"""
import re
import unicodedata

# --------------------------------------------------------------------------- #
#  Nombres de ancla
# --------------------------------------------------------------------------- #
PREFIJO = "fs"

FAM_TABLA = "tabla"
FAM_CAMPO = "campo"
FAM_DATO = "dato"
FAM_PROSA = "prosa"

TAG_REGISTRO = "fs-registro"
TAG_META = "fs-meta"

#: Campos de encabezado que el Excel alimenta. El nombre es también la clave
#: dentro del contexto que devuelve generador_fs.leer_contexto().
CAMPOS_ENCABEZADO = (
    "empresa",
    "titulo",
    "fecha_actual",
    "fecha_previa",
    "miles",
    "moneda",
    "estado_actual",
    "estado_previo",
)

#: Sufijos válidos de una cifra suelta.
CAMPOS_DATO = ("actual", "previo", "nota", "var_abs", "var_pct")

#: Tabla que siempre debe existir en una base conforme.
TABLA_PRINCIPAL = "principal"

#: Zonas de prosa que crea el andamiaje inicial. El usuario puede añadir,
#: renombrar o borrar las que quiera: no son obligatorias.
PROSA_SUGERIDA = ("introduccion", "analisis", "cierre")

#: Longitud máxima de un Tag de control de contenido en Word.
MAX_TAG = 64


def tag_tabla(nombre):
    return f"{PREFIJO}-{FAM_TABLA}-{nombre}"


def tag_campo(nombre):
    return f"{PREFIJO}-{FAM_CAMPO}-{nombre}"


def tag_dato(clave, campo):
    return f"{PREFIJO}-{FAM_DATO}-{clave}-{campo}"


def tag_prosa(nombre):
    return f"{PREFIJO}-{FAM_PROSA}-{nombre}"


def descomponer(tag):
    """'fs-dato-total_assets-actual' -> ('dato', 'total_assets', 'actual').

    Devuelve (familia, nombre, campo). `campo` es None salvo en la familia
    'dato'. Devuelve (None, None, None) si el tag no pertenece al contrato.
    """
    if not isinstance(tag, str):
        return (None, None, None)
    t = tag.strip()
    if t == TAG_REGISTRO:
        return ("registro", None, None)
    if t == TAG_META:
        return ("meta", None, None)
    partes = t.split("-")
    if len(partes) < 3 or partes[0] != PREFIJO:
        return (None, None, None)
    familia = partes[1]
    if familia == FAM_DATO:
        if len(partes) < 4:
            return (None, None, None)
        campo = partes[-1]
        clave = "-".join(partes[2:-1])
        if campo not in CAMPOS_DATO:
            return (None, None, None)
        return (FAM_DATO, clave, campo)
    if familia in (FAM_TABLA, FAM_CAMPO, FAM_PROSA):
        return (familia, "-".join(partes[2:]), None)
    return (None, None, None)


def es_region_de_datos(tag):
    """¿El refresco debe escribir dentro de esta ancla?"""
    fam, _, _ = descomponer(tag)
    return fam in (FAM_TABLA, FAM_CAMPO, FAM_DATO)


# --------------------------------------------------------------------------- #
#  Derivación de la clave de una etiqueta del Excel
# --------------------------------------------------------------------------- #
def clave(etiqueta):
    """Convierte la etiqueta de una fila del Excel en una clave estable.

    Debe dar el MISMO resultado en Python y en TypeScript:
      1. quitar tildes (NFKD, descartando los diacríticos)
      2. minúsculas
      3. todo lo que no sea [a-z0-9] pasa a '_'
      4. colapsar '_' repetidos y recortar los de los extremos
      5. cortar a 40 caracteres, recortando de nuevo los '_' del final

      'Cash and cash equivalents'   -> 'cash_and_cash_equivalents'
      'Total assets'               -> 'total_assets'
      'Provisión (neta)'           -> 'provision_neta'
    """
    if etiqueta is None:
        return ""
    s = unicodedata.normalize("NFKD", str(etiqueta))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s[:40].strip("_")


# --------------------------------------------------------------------------- #
#  Contexto -> valores por ancla
# --------------------------------------------------------------------------- #
def _num_contable(valor):
    """Mismo formato que generador_fs.num(), replicado para no importar."""
    if valor is None:
        return ""
    if isinstance(valor, str):
        return valor
    if not isinstance(valor, (int, float)):
        return str(valor)
    if valor < 0:
        return f"({abs(valor):,.0f})"
    return f"{valor:,.0f}"


def clave_de_linea(linea):
    """La clave con la que se identifica una línea en el documento.

    Prefiere el rango con nombre de Excel (identidad estable: sobrevive a
    que se renombre la etiqueta o se inserten filas). Si la línea no tiene
    rango, cae en la clave derivada del texto de la etiqueta.
    """
    return linea.get("clave") or clave(linea.get("etiqueta"))


def construir_valores(ctx):
    """Del contexto del Excel al mapa {tag: texto} de cifras sueltas y campos.

    No incluye las tablas: esas las arma el motor a partir de ctx['lineas'].
    Devuelve (valores, colisiones) donde `colisiones` lista las etiquetas
    que producen la misma clave (se conserva la primera aparición).
    """
    valores = {}
    colisiones = []
    vistas = {}

    for nombre in CAMPOS_ENCABEZADO:
        v = ctx.get(nombre, "")
        valores[tag_campo(nombre)] = "" if v is None else str(v)

    # Escalares de rangos con nombre que caen fuera de la tabla (una fecha
    # de corte, un tipo de cambio…). Se exponen solo con el campo 'actual'.
    # El escalar llega ya formateado por generador_fs.formatear_valor(),
    # que respeta el formato de la celda en Excel (decimales, %, fechas).
    for k, par in (ctx.get("escalares") or {}).items():
        texto_, crudo_ = par if isinstance(par, (tuple, list)) else (par, None)
        valores[tag_dato(k, "actual")] = "" if texto_ is None else str(texto_)
        for campo in ("previo", "nota", "var_abs", "var_pct"):
            valores.setdefault(tag_dato(k, campo), "")

    for linea in ctx.get("lineas", []):
        k = clave_de_linea(linea)
        if not k:
            continue
        if k in vistas:
            colisiones.append((linea.get("etiqueta"), vistas[k], k))
            continue
        vistas[k] = linea.get("etiqueta")

        valores[tag_dato(k, "actual")] = linea.get("actual", "")
        valores[tag_dato(k, "previo")] = linea.get("previo", "")
        valores[tag_dato(k, "nota")] = linea.get("nota", "")

        a = linea.get("actual_raw")
        p = linea.get("previo_raw")
        if isinstance(a, (int, float)) and isinstance(p, (int, float)):
            valores[tag_dato(k, "var_abs")] = _num_contable(a - p)
            valores[tag_dato(k, "var_pct")] = (
                f"{(a - p) / p * 100:,.1f}%" if p else ""
            )
        else:
            valores[tag_dato(k, "var_abs")] = ""
            valores[tag_dato(k, "var_pct")] = ""

    return valores, colisiones


def catalogo(ctx):
    """Lista legible de las cifras sueltas disponibles, para el panel/informe.

    Devuelve [(clave, origen, etiqueta, actual, previo), ...] en el orden del
    Excel. `origen` es 'rango' (identidad estable) o 'etiqueta' (derivada del
    texto, se rompe si alguien renombra la fila).
    """
    filas = []
    vistas = set()
    for linea in ctx.get("lineas", []):
        k = clave_de_linea(linea)
        if not k or k in vistas:
            continue
        vistas.add(k)
        filas.append((
            k,
            linea.get("clave_origen", "etiqueta"),
            linea.get("etiqueta", ""),
            linea.get("actual", ""),
            linea.get("previo", ""),
        ))
    for k, par in sorted((ctx.get("escalares") or {}).items()):
        if k in vistas:
            continue
        vistas.add(k)
        texto_ = par[0] if isinstance(par, (tuple, list)) else par
        filas.append((k, "rango", "(celda suelta del libro)",
                      "" if texto_ is None else str(texto_), ""))
    return filas
