# Generador de Estado de Situación Financiera

Convierte la hoja del **Estado de Situación Financiera** de tu Excel en un
documento de **Word** con el formato de la plantilla ya aplicado.

Hay **dos caminos**, y hacen cosas distintas. Si no sabe cuál, use el
tercero: pregunta antes de hacer nada. Arrastrar el Excel funciona en
los dos; lo que cambia es el resultado.

| Arrastra tu Excel sobre… | Qué hace | Tu redacción |
|---|---|---|
| **`generar.bat`** · `GeneradorFS.exe` | Crea un Word **nuevo** en `salidas\` | Se pierde en cada corrida |
| **`refrescar.bat`** · `RefrescarFS.exe` | **Actualiza** el documento base de OneDrive | **Se conserva** |
| **`estados_financieros.bat`** · `EstadosFinancieros.exe` | **Pregunta cuál de los dos** antes de hacer nada | según elija |

Si lo que quieres es *«escribo mis párrafos y las cifras se actualizan
solas»*, el segundo. La guía de operación completa está en
**[`docs/GUIA.md`](docs/GUIA.md)**; el resumen, en
[Documento base vivo](#documento-base-vivo).

## Estructura del repositorio

```
generar.bat        crea un Word nuevo          config.json   ajustes
refrescar.bat      actualiza el documento base requirements.txt
estados_financieros.bat   pregunta cual de los dos

src/          codigo Python (motor y ordenes)
docs/         documentacion
plantillas/   documentos .docx modelo
ejemplos/     libro .xlsx de muestra
tools/        scripts de entorno, empaquetado y verificacion
addin/        complemento de Word (TypeScript)
salidas/      lo que genera generar.bat
```

Las carpetas se agrupan por **naturaleza**, no por extensión. Al añadir
archivos, respete esa división: un `.docx` de plantilla va a `plantillas/`,
uno de ejemplo a `ejemplos/`, y la documentación a `docs/`.

---

## Descargar y usar

**Un solo archivo. No hay que instalar nada.** (Windows 64 bits, ~13 MB)

### ⬇ [Descargar EstadosFinancieros.exe](https://github.com/RadioFics/Excel_Word_Paloma/releases/latest/download/EstadosFinancieros.exe)

**Arrastra tu Excel encima** (o haz doble clic) y te pregunta qué hacer:

```
   1)  ACTUALIZAR el documento de siempre
       Conserva todo lo que hayas escrito. Solo cambia las cifras.

   2)  CREAR un documento nuevo
       Sale de la plantilla, en la carpeta salidas\.

   3)  VER el estado del proyecto
```

Eso es todo. No se instala Python, no se descomprime nada, no hace falta
permiso de administrador.

### Configurarlo (solo para la opción 1)

Para que sepa **qué documento** actualizar, deja un `config.json` **junto al
`.exe`** con al menos esta clave:

```json
{
  "documento_base": "C:\\Users\\tu.usuario\\OneDrive - Collective Mining C-Suite\\MI_DOCUMENTO.docx"
}
```

Las barras van dobladas (`\\`), que es como JSON escribe las rutas de
Windows. Sin esta clave, las opciones 2 y 3 funcionan igual; la 1 te dirá que
falta configurarla.

La primera vez que uses un documento, prepáralo:

```
EstadosFinancieros.exe MI_LIBRO.xlsx --refrescar --preparar
```

Todas las claves de `config.json` están en la tabla de
[Ajustes opcionales](#ajustes-opcionales-configjson).

### Los tres ejecutables

`EstadosFinancieros.exe` reúne los dos caminos, pero **cada uno sigue
existiendo por separado** si prefieres un icono por tarea:

| Ejecutable | Qué hace | SHA-256 |
|---|---|---|
| **`EstadosFinancieros.exe`** | Pregunta cuál de los dos | `56835BD4883ADDBEF001F0156BBB1A69DEFC89770E867C9CD1351272AB57F2B7` |
| `GeneradorFS.exe` | Word **nuevo** en `salidas\` | `EFE41A8A5FFDFB4F5A262530EB8267E49CC49F2FF959940C2AD98EA099E438B1` |
| `RefrescarFS.exe` | **Actualiza** el documento base | `4AA1AFEB2049EDA6F67C2688A3D1CDBF1D54E7F21F860FAB0ABB27AA37D9D4DD` |

Los tres admiten que se les arrastre el Excel. Comprueba que descargaste el
archivo legítimo con:

```
powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '.\EstadosFinancieros.exe').Hash"
```

Las huellas cambian en cada recompilación: usa siempre las que acompañen a la
descarga.

### Sin menú, para automatizar

```
EstadosFinancieros.exe MI_LIBRO.xlsx --refrescar
EstadosFinancieros.exe MI_LIBRO.xlsx --generar
EstadosFinancieros.exe --estado
```

### La primera vez, Windows puede advertir

Si aparece la pantalla azul **«Windows protegió su PC»**:
**Más información → Ejecutar de todas formas.**

Es lo normal con cualquier programa recién descargado que todavía no tiene
firma digital. Abajo se explica qué contiene y cómo comprobar que es el
archivo legítimo.

## Qué obtienes

En la carpeta `salidas`, junto al `.exe`:

| Archivo | Qué es |
|---|---|
| `estado_situacion_financiera_<fecha>.docx` | El documento de Word con las cifras puestas. Los anteriores no se borran. |
| `revisar_tipos.csv` | Lista, fila por fila, qué se trasladó y cómo se clasificó cada renglón (encabezado, detalle, subtotal, total, nota). Ábrelo con Excel para revisar. |

Si tu hoja tiene una columna **`Tipo`** (con letras H/I/S/T/N/X), el programa
la respeta. Si no la tiene, deduce solo el papel de cada fila y lo apunta en
`revisar_tipos.csv` para que lo revises.

---

## Si algo sale mal

La ventana explica el problema en español y qué revisar. Los casos más
comunes:

| Mensaje | Qué hacer |
|---|---|
| «No encontré ningún Excel en esta carpeta» | Arrastra el Excel sobre el `.exe`, o cópialo a la misma carpeta. |
| «No pude identificar la hoja…» | Tu libro no tiene una hoja reconocible como Estado de Situación Financiera. |
| El Word sale con cifras en blanco | El Excel se guardó con una herramienta que no recalcula. Ábrelo y guárdalo desde Excel, y repite. |

---

## ¿Es seguro? ¿Es un virus?

No. Los avisos de Windows aparecen porque el programa **no está firmado
digitalmente** (la firma requiere un certificado de pago; es un paso
posterior). Windows advierte de **todo** archivo descargado que no reconoce,
sin mirar su contenido.

- **Qué hace:** lee un archivo `.xlsx` de tu equipo y escribe un `.docx` en
  tu equipo.
- **Qué NO hace:** no se instala, no toca la configuración de Windows, no se
  inicia solo, **no se conecta a internet**, no envía datos a ningún lado.
- **Cómo verificar que es el archivo correcto:** compara su huella SHA-256
  con la tabla de [Los tres ejecutables](#los-tres-ejecutables).
- **Origen:** repositorio `RadioFics/Excel_Word_Paloma`.

Si tu equipo lo **bloquea por completo** (política corporativa / antivirus
que lo pone en cuarentena sin opción de continuar): no lo fuerces.
Repórtalo — es uno de los datos que esta prueba busca confirmar.

---

## Documento base vivo

El modo en el que **el mismo documento** se mantiene al día sin perder lo
que hayas escrito.

### La idea

El documento **no se regenera**: se le **refrescan unas regiones concretas**
(la tabla, los campos de encabezado, las cifras que intercales en la
redacción). Todo lo demás ni se visita.

- Escribes un párrafo → refrescas → **el párrafo sigue ahí**, con las cifras
  al día.
- Borras ese párrafo → refrescas → **sigue borrado**. El motor nunca
  reinyecta prosa, porque nunca la guardó.
- Las tablas y las cifras van **bloqueadas**: nadie las pisa a mano.

La especificación completa está en **[`docs/CONTRATO.md`](docs/CONTRATO.md)**.

### Apunta al documento (una vez)

En `config.json`, la clave `documento_base` dice qué archivo se refresca:

```json
"documento_base": "C:\\Users\\...\\OneDrive\\MI_DOCUMENTO.docx"
```

### Preparar tu documento (una vez)

Parte de la plantilla lista, o de un documento tuyo:

```bash
refrescar.bat --preparar
```

Añade lo que le falte para cumplir el contrato **sin borrar nada**. Puedes
correrlo sobre un documento con meses de redacción encima: solo agrega las
anclas que no estén. También hay una plantilla ya armada:
[`plantillas/plantilla_base_EF.docx`](plantillas/plantilla_base_EF.docx).

### Cada cierre

Arrastra tu Excel sobre **`refrescar.bat`**. O, si prefieres escribirlo:

```bash
python src\fs_documento.py refrescar "MI_DOCUMENTO.docx" "MI_LIBRO.xlsx"
```

Reescribe la tabla y los campos, antepone a la bitácora del documento el
detalle de qué cifra cambió, y deja una copia `.bak` por si acaso.

> **Cierra el documento en Word antes.** Si Word lo tiene abierto, la orden
> se detiene sin tocar nada y te lo dice. No es una recomendación: escribir
> sobre un `.docx` que Word tiene abierto lo deja inservible.

### Intercalar una cifra viva en la redacción

Para escribir *«los activos totales ascendieron a **119,066,301**»* y que esa
cifra siga al Excel:

```bash
python src\fs_documento.py catalogo
```

te lista las claves disponibles. Luego, o bien la insertas desde Word
(control de contenido de texto con la Etiqueta `fs-dato-total_assets-actual`),
o bien:

```bash
python src\fs_documento.py insertar "MI_DOCUMENTO.docx" total_assets actual
```

Campos: `actual`, `previo`, `nota`, `var_abs`, `var_pct`.

### Los dos editores

Por defecto las regiones de datos están bloqueadas pero la prosa es libre en
todo el documento. Si quieres el modo estricto:

```bash
python src\fs_documento.py proteger "MI_DOCUMENTO.docx" --clave TU_CLAVE
```

| Rol | Puede | No puede |
|---|---|---|
| **Redactor** | escribir en las zonas `fs-prosa-*` | tocar tablas, campos ni cifras |
| **Editor de datos** | lo anterior + refrescar | — (tiene la clave) |

Para volver atrás: `python src\fs_documento.py desproteger "MI_DOCUMENTO.docx"`.

### Órdenes

| Orden | Para qué |
|---|---|
| `construir` / `reparar` | Añade las anclas que falten. No destruye nada. |
| `refrescar` | Actualiza las regiones de datos desde el Excel. |
| `insertar` | Coloca una cifra viva en una zona de prosa. |
| `catalogo` | Lista las claves disponibles en el Excel. |
| `nombrar` | Crea los rangos con nombre en el Excel (identidad estable). |
| `verificar` | Revisa el documento: anclas huérfanas, regiones vacías. |
| `plantilla` | Genera un documento base nuevo desde cero. |
| `proteger` / `desproteger` | Modo estricto de dos editores. |

### Fija las cifras con rangos con nombre

Por defecto, una cifra de la prosa se identifica por el **texto de la
etiqueta** de su fila. Si alguien renombra «Total assets» en el Excel, el
vínculo se rompe (se reporta como huérfano, no se rellena mal).

Para que aguante renombrados, reordenaciones e inserciones de filas:

```bash
python src\fs_documento.py nombrar "MI_LIBRO.xlsx" --aplicar
```

Crea un nombre de Excel `fs_<clave>` por cada fila. **Cierra el libro en
Excel antes.** Los nombres los escribe Excel, no openpyxl, así que las
fórmulas y sus valores cacheados quedan intactos.

### Antes de refrescar

**Cierra el documento en Word.** El motor escribe el archivo directamente; si
Word lo tiene abierto, se pelean. Y la co-autoría de Word en el navegador
maneja mal los controles de contenido: refresca desde el escritorio.

---

## Para desarrollo (no necesario para usar el `.exe`)

El `.exe` se genera desde `src\generador_fs.py` (Python). Para trabajar el código:

| Documento | Contenido |
|---|---|
| [`docs/GUIA.md`](docs/GUIA.md) | **Guía de operación paso a paso** del documento base vivo. Empieza por aquí. |
| [`docs/CONTRATO.md`](docs/CONTRATO.md) | Contrato de anclas Excel⇄Word. Lo comparten el motor de Python y el add-in. |
| [`docs/DATOS.md`](docs/DATOS.md) | **Dónde se editan las cifras** y por qué el Word está bloqueado. |
| [`docs/CAMBIAR_EXCEL.md`](docs/CAMBIAR_EXCEL.md) | Qué revisar **cuando llegue el Excel definitivo**. |
| [`docs/ESTRUCTURA.md`](docs/ESTRUCTURA.md) | Cómo están organizadas las carpetas y qué rutas dependen de ello. |
| [`docs/DESPLIEGUE_ADDIN.md`](docs/DESPLIEGUE_ADDIN.md) | Subir el complemento de Word y qué pedirle a TI. |
| [`docs/INSTALACION.md`](docs/INSTALACION.md) | Montar Python portable en `.\python\` sin permisos de administrador. |
| [`docs/PRUEBA_EXTERNA.md`](docs/PRUEBA_EXTERNA.md) | Reproducir la prueba en otro equipo (con el `.exe` o clonando). |
| [`docs/DIRECCION.md`](docs/DIRECCION.md) | Dirección del proyecto, alternativas y caso ante TI. |
| [`addin/`](addin/) | Add-in de Word (v0) que actualiza el mismo documento en vez de generar uno nuevo. |

Comandos (equipo con el entorno montado):

```
powershell -ExecutionPolicy Bypass -File .\tools\bootstrap_python.ps1   # monta .\python\
powershell -ExecutionPolicy Bypass -File .\tools\verificar.ps1          # prueba de humo
powershell -ExecutionPolicy Bypass -File .\tools\hacer_exe.ps1          # genera dist\GeneradorFS.exe
```

### Ajustes opcionales (`config.json`)

Va **embebido** en el `.exe` con valores por defecto. Para cambiarlos sin
recompilar, deja un `config.json` propio en la misma carpeta que el `.exe`:

| Clave | Para qué |
|---|---|
| `empresa` | Nombre de la empresa (no viene de la hoja). |
| `hoja` | Nombre exacto de la hoja. Si no existe, se elige por contenido. |
| `hoja_marcadores` | Textos que delatan un estado de situación financiera. |
| `plantilla` | Nombre de un `.docx` de plantilla alternativo (en la carpeta del `.exe`). |
| `primera_fila` | Un número fija la fila de inicio; `"auto"` la detecta. |
| `columnas` | Letras (`A`, `C`, `E`, `F`, `G`) para forzar una columna; `null` = detectar. |
| `marcadores_excluir` | Etiquetas que marcan una fila de control/cuadre. |
| `documento_base` | **El .docx que actualiza la opción 1.** Ruta absoluta, o relativa a la carpeta del `.exe`. |
| `prefijo_rangos` | Prefijo de los nombres de Excel que fijan la identidad de cada cifra (`fs_`). |
| `bitacora` | Dónde se anota el histórico: `"archivo"` (por defecto), `"documento"`, `"ambos"` o `"no"`. |
| `bitacora_archivo` | Ruta del `.log`. Vacío = `salidas\bitacora_<documento>.log`. |

### Cómo clasifica cada fila

| Letra | Significa | Cómo se deduce si no hay columna `Tipo` |
|---|---|---|
| `H` | Encabezado de sección | etiqueta sin cifras, en mayúsculas / termina en ":" / en negrita |
| `I` | Línea de detalle | etiqueta + cifras (o + número de nota) |
| `S` | Subtotal (sin etiqueta) | cifras sin etiqueta |
| `T` | Total (con etiqueta) | la etiqueta empieza por "Total" |
| `N` | Nota de texto libre | texto largo sin cifras |
| `X` | Excluir a propósito | la etiqueta o la nota contienen "control check" / "cuadre" |

### Límites

- Un solo tipo de estado (Situación Financiera). Otro estado = otra plantilla.
- La deducción de tipos es heurística: por eso está `revisar_tipos.csv`.
- `GeneradorFS.exe` genera un Word **nuevo** cada vez. Para actualizar *el
  mismo* documento conservando la redacción, use `fs_documento.py`
  ([arriba](#documento-base-vivo)); el add-in (`addin/`) hace lo mismo desde
  un panel dentro de Word.
- Las cifras sueltas dentro de párrafos requieren marcarlas una vez con un
  control de contenido (`fs-dato-…`); no se detectan solas.
- El vínculo cifra↔prosa se apoya en la etiqueta de la fila del Excel: si se
  renombra una fila, el ancla queda huérfana (se reporta, no falla).
- Solo Windows de 64 bits.

### Trazabilidad

Cada Word generado guarda en sus propiedades (Archivo → Información →
Propiedades → Comentarios) el nombre del Excel de origen, su huella SHA-256,
la hoja usada y la fecha.
