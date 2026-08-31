# Generador de Estado de Situación Financiera (piloto)

Traslada el Estado de Situación Financiera de un libro de Excel a una
plantilla de Word. Cambia el Excel, corres el programa, obtienes un Word
actualizado. La plantilla nunca se toca; el Excel nunca se modifica.

Desde la versión de detección por contenido, el programa **ya no depende de
nombres ni posiciones fijas**: identifica la hoja y las columnas por lo que
contienen, y la columna `Tipo` pasó a ser **opcional** (si falta, infiere el
tipo de cada fila y deja un CSV de revisión).

## Estructura de la carpeta

```
generador_fs.py                              <- el programa
config.json                                  <- ajustes (todo opcional)
plantilla_estado_situacion_financiera.docx   <- la plantilla (no editar el formato)
generar.bat                                  <- lanzador de doble clic (Windows)
requirements.txt                             <- dependencias exactas
tools/bootstrap_python.ps1                   <- instala Python portable en .\python\
INSTALACION.md                               <- cómo configurar Python (sin admin)
DIRECCION.md                                 <- dirección del proyecto y caso ante TI
Copia_Editable_con_columna_Tipo.xlsx         <- ejemplo de Excel ya etiquetado
salidas/                                     <- documentos generados + revisar_tipos.csv
```

## Puesta en marcha (una vez por equipo)

No requiere permisos de administrador. Ver `INSTALACION.md`. En resumen:

```
powershell -ExecutionPolicy Bypass -File .\tools\bootstrap_python.ps1
```

Deja un Python portable en `.\python\`. `generar.bat` lo usa solo.

## Para el usuario que solo necesita generar el documento

1. Arrastra el archivo Excel sobre `generar.bat`, o haz doble clic en
   `generar.bat` sin arrastrar nada (buscará el Excel más reciente de esta
   carpeta cuyo nombre contenga "FS").
2. Se abre una ventana, muestra el resultado y espera a que presiones Enter.
3. El documento nuevo queda en `salidas`, con la fecha y hora en el nombre.
   Los anteriores no se borran ni se sobrescriben.
4. **Revisa `salidas\revisar_tipos.csv`**: lista fila por fila qué se
   trasladó, con qué tipo, y si ese tipo fue *declarado* (venía en la hoja)
   o *inferido* (lo dedujo el programa). Si algo no cuadra, corrígelo en el
   Excel y vuelve a generar.

Para ver solo esa revisión sin generar el Word:

```
generar.bat  y luego, en una terminal:  python\python.exe generador_fs.py "tu.xlsx" --revisar
```

## La columna "Tipo" (opcional)

Si la hoja tiene una columna cuyo encabezado es `Tipo` (o `Type`), el
programa usa esas letras. Si no la tiene, **infiere** el tipo de cada fila y
lo anota en `revisar_tipos.csv` con la señal que lo decidió.

| Letra | Significa | Cómo se infiere si falta |
|---|---|---|
| `H` | Encabezado de sección | etiqueta sin cifras, en mayúsculas / termina en ":" / en negrita |
| `I` | Línea de detalle | etiqueta + cifras (o + número de nota) |
| `S` | Subtotal (sin etiqueta) | cifras sin etiqueta |
| `T` | Total (con etiqueta) | la etiqueta empieza por "Total" |
| `N` | Nota de texto libre | texto largo sin cifras |
| `X` | Excluir a propósito | la etiqueta o la nota contienen "control check" / "cuadre" (configurable) |

La inferencia acierta en la práctica sobre este estado, pero **no es
infalible**: por eso está el CSV de revisión. Para dejar el mapeo fijo y sin
ambigüedad, añade la columna `Tipo` a la hoja.

## Ajustes (`config.json`)

Todo es opcional; borra una clave para usar su valor por defecto.

| Clave | Para qué |
|---|---|
| `empresa` | Nombre de la empresa (no viene de la hoja). |
| `hoja` | Nombre exacto de la hoja. Si no existe, se elige por contenido. |
| `hoja_marcadores` | Textos que delatan un estado de situación financiera. |
| `plantilla` | Nombre del `.docx` de plantilla. |
| `buscar_por_convencion` | Texto que debe contener el nombre del Excel al buscar sin argumentos. |
| `primera_fila` | Un número fija la fila de inicio; `"auto"` la detecta. |
| `columnas` | Letras (`A`, `C`, `E`, `F`, `G`) para forzar una columna; `null` = detectar. |
| `marcadores_excluir` | Etiquetas que marcan una fila de control/cuadre. |

## Advertencia operativa — edición del Excel

Abre y guarda el libro **siempre desde Excel o LibreOffice**. **Nunca**
reguardes este libro con un script de `openpyxl`: no recalcula fórmulas y,
al reguardar, descarta el valor cacheado de **todas** las fórmulas. El
síntoma es que el programa corre sin error pero el Word sale con las cifras
en blanco.

## Límites y alcance de esta versión

- Cubre un solo tipo de estado (Situación Financiera). Otro estado (Estado
  de Resultados, Flujo de Efectivo) necesita su propia plantilla y sus
  marcadores, siguiendo el mismo patrón.
- No traslada valores sueltos dentro de párrafos narrativos (solo tablas).
- La inferencia de tipos es heurística: revisa `revisar_tipos.csv`.
- Genera un documento **nuevo** cada vez. Para actualizar *el mismo*
  documento en su sitio, ver el frente C (add-in) en `DIRECCION.md`.
- No hay pruebas automatizadas todavía.
- Solo Windows x64 para el Python portable. Ver límites de gobernanza (TI)
  en `INSTALACION.md` y `DIRECCION.md`.

## Trazabilidad

Cada documento generado guarda, en sus propiedades de Word (Archivo →
Información → Propiedades → Comentarios), el nombre del Excel de origen, un
identificador único de su contenido (SHA-256), la hoja usada y la fecha de
generación.
