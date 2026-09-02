# 03 · La lectura del Excel y el contrato de anclas

> **Para quién.** Quien tenga que entender, ajustar o depurar de dónde salen
> las cifras: quien mantiene el código, y quien administra el libro de Excel y
> el `config.json`.
> **Qué encontrará.** El recorrido de `src/generador_fs.py` desde el `.xlsx`
> hasta el diccionario `ctx`, con los umbrales exactos de cada heurística; la
> especificación del contrato de anclas (`src/fs_contrato.py`); y el
> comportamiento del generador clásico.
> **Antes de leer.** Aquí no se escribe nada en Word. Esta capa solo lee Excel
> y construye el contexto que después consumen el generador clásico y el motor
> de refresco del documento vivo.

## Índice

- [1. Qué resuelve esta capa](#1-qué-resuelve-esta-capa)
- [2. El recorrido, de un vistazo](#2-el-recorrido-de-un-vistazo)
- [3. Configuración](#3-configuración)
- [4. Antes de abrir el libro](#4-antes-de-abrir-el-libro)
- [5. Elegir la hoja por su contenido](#5-elegir-la-hoja-por-su-contenido)
- [6. Materializar y perfilar las columnas](#6-materializar-y-perfilar-las-columnas)
- [7. La región de datos](#7-la-región-de-datos)
- [8. El encabezado por contenido](#8-el-encabezado-por-contenido)
- [9. Clasificar cada fila](#9-clasificar-cada-fila)
- [10. Armar las líneas](#10-armar-las-líneas)
- [11. El formato de los números](#11-el-formato-de-los-números)
- [12. La identidad estable](#12-la-identidad-estable)
- [13. El contexto (`ctx`)](#13-el-contexto-ctx)
- [14. El contrato de anclas](#14-el-contrato-de-anclas)
- [15. El generador clásico](#15-el-generador-clásico)
- [16. Puntos frágiles de esta capa](#16-puntos-frágiles-de-esta-capa)
- [Resumen del capítulo](#resumen-del-capítulo)

---

## 1. Qué resuelve esta capa

La tabla del estado de situación financiera no está siempre en el mismo sitio.
Puede empezar en `A1` o en `K14`; la hoja puede llamarse `FS`, `Balance` o
`Hoja3`; entre un cierre y otro alguien inserta filas, añade una columna de
notas o renombra la hoja. Un programa que lea coordenadas fijas funciona el
primer mes y falla el segundo, y falla **en silencio**: no da error, produce un
documento con las celdas equivocadas o con las cifras en blanco.

La tesis del autor, declarada en la cabecera del módulo
(`src/generador_fs.py:1`), es que **nada se localiza por coordenadas**. La
hoja, las columnas, el bloque de cabecera, la primera fila de datos y el tipo
de cada fila se identifican por lo que la celda *contiene*.

La consecuencia práctica es doble, y hay que aceptarla entera:

1. El libro puede moverse y renombrarse sin tocar el código ni la
   configuración.
2. A cambio, la lectura es **heurística**: acierta por reglas explícitas, no
   por certeza. Por eso cada decisión inferida deja rastro en
   `salidas/revisar_tipos.csv`, y por eso hay dos válvulas de escape: forzar
   las columnas por letra en `config.json`, y declarar el tipo de cada fila en
   una columna `Tipo` del propio Excel.

Del mismo principio nace el segundo mecanismo central: los **rangos con
nombre** `fs_*`, que dan a cada fila una identidad que sobrevive a que la
renombren (ver [§12](#12-la-identidad-estable)).

Hay además una restricción operativa que el autor repite en tres sitios del
archivo:

> *"Abra y guarde el libro SIEMPRE desde Excel o LibreOffice. NUNCA reguarde
> este libro con un script de openpyxl: openpyxl no recalcula fórmulas y, al
> reguardar, descarta el valor cacheado de TODAS las fórmulas. El síntoma es
> que el generador corre sin error pero el Word sale con las cifras en
> blanco."*

El libro se abre con `load_workbook(ruta, data_only=True, read_only=True)`
(`src/generador_fs.py:1010`): `data_only=True` devuelve el valor **cacheado**
de cada fórmula, que es justo lo que openpyxl destruye al reguardar.

---

## 2. El recorrido, de un vistazo

```
  <libro.xlsx>
      │
      ▼
  comprobar_legible ..... ¿existe? ¿lo suelta Excel? ¿está en la nube?
      │                   (aborta con prosa, no con traza)
      ▼
  load_workbook(data_only=True, read_only=True)
      │
      ▼
  elegir hoja ........... cfg["hoja"] como PISTA -> puntuar_hoja
      │                   si no llega a UMBRAL_HOJA (2):
      │                   _elegir_hoja_por_contenido sobre todo el libro
      ▼
  (reapertura del libro)  los iteradores read_only ya se consumieron
      │
      ▼
  _materializar ......... valores[][], negrita[][], n_filas  (1-indexado)
      │
      ▼
  detectar_columnas ..... etiqueta / nota / actual / previo / tipo
      │
      ▼
  detectar_region ....... (primera, ultima)
      │
      ▼
  detectar_encabezado ... fechas / escala / estado / moneda / fin
  leer_encabezado ....... nacen las 8 claves de cabecera del ctx
      │
      ▼
  inferir_tipo (fila a fila) ........ H I S T N X
  construir_lineas ...... ctx["lineas"] + revisar_tipos.csv
      │
      ▼
  leer_rangos_con_nombre  nombres fs_* -> identidad estable + escalares
      │
      ▼
  ctx = {cabecera..., lineas, escalares, _meta, _avisos}
      │
      ├──► generador clásico: docxtpl -> una FOTO en salidas\
      └──► motor de refresco: fs_contrato -> {tag: texto} por región
```

---

## 3. Configuración

### 3.1 `DEFAULTS`, clave por clave

`DEFAULTS` (`src/generador_fs.py:117`) es el esquema central del proyecto: no
todas sus claves se usan en este archivo, pero todas se cargan aquí.

| Clave | Valor por defecto | Para qué |
|---|---|---|
| `empresa` | `"Collective Mining Ltd."` | Se copia a `ctx["empresa"]`; no sale del Excel. |
| `hoja` | `"FS"` | Pista del nombre de hoja, que se verifica antes de aceptarse. |
| `hoja_marcadores` | 8 frases (`"situación financiera"`, `"total assets"`, `"assets"`…) | Señales con las que se puntúa cada hoja. |
| `plantilla` | `"plantilla_estado_situacion_financiera.docx"` | Plantilla del generador clásico. |
| `buscar_por_convencion` | `"FS"` | Bonus de nombre al puntuar la hoja **y** filtro del nombre del `.xlsx`. |
| `primera_fila` | `"auto"` | Un entero fija la fila de inicio; `"auto"` la detecta. |
| `columnas` | los cinco campos a `null` | Letras (`"C"`, `"E"`…) para forzar; `null` para detectar. |
| `marcadores_excluir` | `["control check", "check", "cuadre", "balance check"]` | Filas de cuadre que nunca llegan al documento. |
| `max_filas_scan` | `400` | **Mínimo** de filas a barrer, no tope. |
| `max_cols_scan` | `16` | **Mínimo** de columnas a barrer, no tope. |
| `prefijo_rangos` | `"fs_"` | Prefijo de los rangos con nombre. Vacío desactiva el mecanismo. |
| `documento_base` | `""` | El documento vivo. El generador clásico lo ignora. |
| `bitacora` | `"archivo"` | `"archivo"`, `"documento"`, `"ambos"` o `"no"`. |
| `bitacora_archivo` | `""` | Vacío equivale a `salidas\bitacora_<documento>.log`. |
| `apariencia_datos` | `"boundingBox"` | Aspecto de las regiones de datos; alternativa `"hidden"`. |
| `clave_proteccion` | `"fs"` | *"No es un secreto: solo evita ediciones accidentales."* |

Las cinco últimas no se leen en este archivo: las consume el módulo del
documento. La configuración es una sola para todo el proyecto.

### 3.2 La cascada

`cargar_config()` (`src/generador_fs.py:240`) parte de una copia profunda de
`DEFAULTS` —hecha con `json.loads(json.dumps(DEFAULTS))`, que además garantiza
que la configuración viva tenga los mismos tipos que un archivo JSON— y la
fusiona con hasta tres archivos, de menor a mayor prioridad:

| # | Origen | Ruta | Como script | Como `.exe` |
|---|---|---|---|---|
| 0 | `DEFAULTS` | en el código | siempre | siempre |
| 1 | `CONFIG_EMBEBIDA` | `RECURSOS/config.json` | coincide con la 2, se omite | la embebida el día de compilar |
| 2 | `CONFIG_PATH` | `BASE/config.json` | aplicada | aplicada (junto al `.exe`) |
| 3 | `CONFIG_LOCAL` | `BASE/config.local.json` | aplicada | aplicada |

El orden no es arbitrario:

> *"Los ajustes, de menos a más prioritario. El orden importa: lo de este
> equipo (config.local.json) tiene que poder ganarle a lo que venga por git, o
> cada «pull» reimpondria las rutas de la otra máquina."*

`_fusionar_config` (`src/generador_fs.py:224`) tiene cuatro comportamientos que
conviene conocer: un archivo inexistente no es error; un JSON mal formado se
convierte en `ValueError` con el nombre del archivo y el error de sintaxis, no
en una traza; **las claves que empiezan por `_` se ignoran** (es el mecanismo
de comentarios, que el `config.json` del repositorio usa); y `columnas` se
fusiona con `.update()`, mientras que **cualquier otra clave se reemplaza
entera** —quien redefina `hoja_marcadores` sustituye la lista completa.

### 3.3 Marcadores de ruta portables

`config.json` viaja por git; una ruta absoluta escrita ahí se le impone a la
otra máquina en cada `pull`. De ahí tres funciones.

`raiz_onedrive()` (`src/generador_fs.py:256`) busca la carpeta de OneDrive **de
empresa**, en este orden: `Path.home().glob("OneDrive - *")` ordenado
alfabéticamente (la primera gana, para que el resultado sea determinista con
varias cuentas); las variables `OneDriveCommercial`, `OneDrive`,
`OneDriveConsumer`; `~/OneDrive`; y si nada existe, `None`. No basta
`%OneDrive%` porque *"en un equipo con las dos cuentas apunta a la PERSONAL […]
no a la de la organizacion […] que es donde vive el documento"*.

`expandir_ruta(crudo)` (`src/generador_fs.py:276`) traduce una ruta de la
configuración a este equipo:

| Marcador | Se sustituye por |
|---|---|
| `${ONEDRIVE}` | `raiz_onedrive()`, y si no hay ninguna, `Path.home()` |
| `${USUARIO}` | `Path.home()` |
| `${USERPROFILE}` | `Path.home()` (admitido, pero ausente del docstring) |
| `${PROYECTO}` | `BASE`, la carpeta del proyecto o del `.exe` |
| `~` inicial | `Path.home()`, aplicado **después** de los marcadores |

La comparación es insensible a mayúsculas —*"nadie deberia fallar por escribir
${onedrive} en minuscula"*— pero solo se sustituye la **primera** aparición de
cada marcador.

`compactar_ruta(ruta)` (`src/generador_fs.py:310`) es la inversa: al guardar la
ruta del documento elegido intenta `relative_to` contra el OneDrive de empresa
y luego contra la carpeta del usuario, y devuelve `"${ONEDRIVE}\\…"` o
`"${USUARIO}\\…"`. Nunca emite `${PROYECTO}` ni `${USERPROFILE}`, y el
separador `"\\"` está fijado a Windows.

---

## 4. Antes de abrir el libro

`comprobar_legible(ruta)` (`src/generador_fs.py:933`) es el único portal por el
que pasan todos los caminos que leen cifras. Su razón de ser es de trato:

> *"Sin esta comprobación, ese caso salía como un volcado de PermissionError en
> mitad de openpyxl: veinte líneas de traza que no le dicen a nadie que lo
> único que hay que hacer es cerrar Excel."*

La secuencia: si el archivo no existe, `ValueError` con la ruta. Si
`open(ruta, "rb")` funciona, la función **retorna sin leer nada**. Un
`PermissionError` pasa al diagnóstico; cualquier otro `OSError` se convierte en
un `ValueError` explicativo.

El diagnóstico tiene dos ramas. `_esta_en_la_nube(ruta)`
(`src/generador_fs.py:926`) consulta los atributos de OneDrive «Archivos a
petición» —`_SOLO_EN_LA_NUBE = 0x00400000 | 0x00040000`, es decir
`FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS` y `FILE_ATTRIBUTE_RECALL_ON_OPEN`—: el
archivo figura en la carpeta, pero su contenido sigue en la nube y leerlo
dispara una descarga que puede fallar *"sin que nada indique que el archivo
está incompleto"*. El mensaje termina con la instrucción: *"Ábralo una vez en
Excel, o haga clic derecho sobre él -> «Conservar siempre en este
dispositivo»"*. `st_file_attributes` solo existe en Windows; fuera de Windows
la función degrada a `False`.

Si no está en la nube, se pregunta quién lo retiene, con un import diferido y
protegido de `fs_documento` (`quien_bloquea`, la misma maquinaria del Restart
Manager que usa Windows para el cartel «este archivo está siendo utilizado
por…»). Si ningún culpable contiene la frase `"en este equipo"`, el mensaje
concluye que probablemente lo tenga abierto otra persona a través de OneDrive.

Que Excel impida **incluso leer** sorprende a mucha gente: dentro de OneDrive y
con el autoguardado activo, Excel retiene el archivo en exclusiva mientras lo
tiene abierto. De ahí el cierre del mensaje, que es también la promesa de
diseño de toda la capa:

> *"El libro solo se LEE, nunca se escribe, así que no hay nada que perder: es
> Excel quien no deja ni leerlo mientras lo tiene abierto."*

---

## 5. Elegir la hoja por su contenido

`_texto_de_hoja(ws, max_filas=120, max_cols=60)` (`src/generador_fs.py:335`)
devuelve **una sola cadena** con todo el texto de una ventana de 120×60 celdas,
normalizado y unido por espacios. La ventana es ancha a propósito:

> *"antes se miraban 60 filas por 8 columnas, y una tabla movida a K14 caía
> entera fuera. La hoja se puntuaba con cero señales y se descartaba, aunque
> fuese la buena."*

`puntuar_hoja(ws, cfg, nombre=None)` (`src/generador_fs.py:354`) puntúa así:

| Señal | Puntos | Condición exacta |
|---|---|---|
| Cada marcador de `hoja_marcadores` | **+1,0** cada uno | el marcador normalizado es subcadena del texto de la hoja |
| Nombre de la hoja | **+0,5**, una vez | `buscar_por_convencion` normalizado es subcadena del nombre |
| Hoja visible | **+0,5**, una vez | `sheet_state == "visible"` |

Si el texto de la hoja sale vacío, devuelve `0.0` de inmediato: ni el nombre
suma. `UMBRAL_HOJA = 2` (`src/generador_fs.py:332`) es el corte: *"Cuántas
señales de contenido bastan para dar una hoja por buena."*

Dos consecuencias numéricas que no se ven a simple vista. Nombre más visible
suman 1,0: **nunca alcanzan el umbral por sí solos**, el contenido siempre es
necesario. Y los marcadores por defecto **se solapan**: una hoja con
`"Total assets"` dispara a la vez `"total assets"` y `"assets"`, y ya llega a 2.

`_elegir_hoja_por_contenido(wb, cfg)` (`src/generador_fs.py:379`) recorre las
hojas en el orden del libro y se queda con la de mayor puntuación
—estrictamente mayor, así que en empate gana la primera—. Si nadie llega al
umbral, `ValueError` sugiriendo indicar la hoja en `config.json -> "hoja"`.

La pieza que da sentido a todo está en `leer_contexto`
(`src/generador_fs.py:1013`): el nombre configurado **es una pista que se
verifica**.

> *"El nombre de config.json es una PISTA que se comprueba, no una orden. Se
> mira primero la hoja que se llama así, pero solo se acepta si además tiene
> forma de estado de situación financiera. […] Así el libro sigue funcionando
> aunque renombren la hoja, y una hoja que se llame 'FS' sin serlo no arrastra
> al resto."*

Si existe una hoja `FS` **y** puntúa 2 o más, se acepta y queda registrada la
justificación `"'FS' por nombre, confirmada por su contenido (3 señales)"`. Si
existe pero no puntúa, se descarta igual que si no existiera. La coincidencia
por nombre es exacta y sensible a mayúsculas.

Tras elegir, el libro **se cierra y se vuelve a abrir**
(`src/generador_fs.py:1027`): en modo `read_only` los iteradores de las hojas
ya se consumieron al puntuar y no se rebobinan de forma fiable.

---

## 6. Materializar y perfilar las columnas

### 6.1 `_materializar(ws, cfg)`

`_materializar` (`src/generador_fs.py:402`) convierte la hoja en tres
estructuras, en una sola pasada: `valores`, lista de listas **1-indexada en
ambos ejes** (`valores[0]` es una fila centinela), de modo que se indexa con
los mismos números que muestra Excel; `negrita`, con la misma forma, que
degrada a `False` si openpyxl no deja acceder al estilo; y `n_filas`, el número
de filas útiles tras recortar las filas finales vacías.

El detalle decisivo es el tamaño del barrido:

```python
max_filas = max(int(cfg["max_filas_scan"]), min(ws.max_row or 1, 20000))
max_cols  = max(int(cfg["max_cols_scan"]),  min(ws.max_column or 1, 256))
```

Los límites de configuración **son suelos, no techos**, por un fallo real:

> *"Antes eran un tope fijo (16 columnas). En cuanto alguien movia la tabla a
> la derecha —K14, por ejemplo— la columna 'Tipo' caia fuera del barrido y los
> tipos declarados desaparecian sin un solo aviso."*

Los techos duros —20 000 filas y 256 columnas— no son configurables y evitan
que una hoja con basura en `XFD1048576` obligue a materializar un millón de
filas.

### 6.2 `detectar_columnas(valores, cfg, n_filas, max_cols)`

`src/generador_fs.py:440`. Devuelve índices 1-based:
`{"etiqueta": int|None, "nota": int|None, "actual": int, "previo": int|None,
"tipo": int|None}`.

**Paso 0 — forzado.** Se leen las letras de `cfg["columnas"]` y se convierten a
índice. Una columna forzada **bloquea** la detección de esa clave: los pasos
siguientes están guardados con `col.get(k) is None`.

**Paso 1 — rótulos explícitos.** Se recorre **toda la altura barrida**, fila a
fila y de izquierda a derecha, comparando la celda normalizada contra una lista
cerrada: `tipo` si es exactamente `"tipo"` o `"type"`; `nota` si es `"note"`,
`"nota"`, `"notes"` o `"notas"`. Gana la coincidencia más alta de la hoja. Que
se busque en toda la altura responde al mismo fallo del barrido corto: *"El
resultado era mudo y peligroso: todos los tipos volvian a inferirse como si
nunca se hubieran declarado."*

**Paso 2 — perfilado.** Para cada columna se cuentan, **sobre las filas 2 a
`n_filas`** (la 1 se excluye por suponerse cabecera), `n_num` (celdas `int` o
`float`; `bool` no cuenta) y `n_txt` (texto no vacío). Un valor numérico nunca
cuenta como texto: la rama es `elif`.

**Paso 3 — las dos columnas de cifras.** Umbral: `n_num >= 2`.

```python
con_numeros = [c for c in perfil if perfil[c][1] >= 2]
con_numeros.sort(key=lambda c: (perfil[c][1], c), reverse=True)
elegidas = sorted(con_numeros[:2])
```

Se ordena por cantidad de números y, en empate, por índice descendente
(*"empate -> más a la derecha"*); se toman las dos primeras y se reordenan de
izquierda a derecha: `actual = elegidas[0]`, `previo = elegidas[1]`. Aquí vive
una **convención implícita** de la que depende todo el documento: se asume que
el periodo actual está **a la izquierda** de la comparativa. Si el libro las
invierte, el resultado se invierte sin aviso. Si solo hay una candidata,
`previo` queda en `None`.

**Paso 4 — etiqueta.** Umbral: `n_txt >= 3`. La **primera** columna
estrictamente a la izquierda de `actual` con al menos tres celdas de texto; si
no hay ninguna, la columna 1.

**Paso 5 — nota.** Solo si el rótulo no la fijó y hay `actual`. Se recorren las
columnas estrictamente **entre** etiqueta y actual y se acepta la primera que
cumpla las dos condiciones: al menos **un** entero en el intervalo `(0, 99]` y
**como mucho dos** celdas de texto en toda la columna.

**Validación final.** Sin `actual`, `ValueError`: *"No pude identificar las
columnas de cifras en la hoja. Defina las letras a mano en config.json ->
\"columnas\""*. Las demás pueden quedar en `None` sin abortar.

### 6.3 Un ejemplo concreto

Una hoja `Balance 2025` cuya tabla arranca en `C6`:

```
      A   B   C                             D      E            F              G
  6           Collective Mining Ltd.
  7           Statement of Financial Position
  8                                         Note   June 30,     December 31,
  9                                                1000         1000
 10                                                (Unaudited)  (Audited)
 11                                                US$          US$
 12           ASSETS
 13           Current assets
 14           Cash and cash equivalents     5      119,066,301   96,094,583
 15           Trade receivables             6        3,201,004    2,880,110
 16           Total current assets                 122,267,305   98,974,693
```

- **Paso 1.** `D8` contiene `Note`: fija `nota = D` (columna 4). No hay rótulo
  `Tipo`, así que esa columna queda en `None`.
- **Paso 2.** `C` acumula muchas celdas de texto y ninguna numérica; `D` tiene
  dos números (las notas 5 y 6) y un texto; `E` y `F` tienen las cifras del
  estado más el `1000` de la fila 9.
- **Paso 3.** Cumplen `n_num >= 2` las columnas `D`, `E` y `F`. `E` y `F` van
  muy por delante en cantidad; reordenadas queda **`actual = E`,
  `previo = F`**. `D`, con solo dos números, no entra.
- **Paso 4.** El límite es `E`; a su izquierda, la primera columna con tres o
  más textos es `C`: **`etiqueta = C`**.
- **Paso 5.** No se ejecuta, porque `nota` ya está fijada. Si la cabecera
  dijera `Ref.` en vez de `Note`, este paso miraría las columnas entre `C` y
  `E` —solo `D`—, hallaría dos enteros en `(0, 99]` y un solo texto, y elegiría
  `D` igual.

Resultado: `etiqueta=C, nota=D, actual=E, previo=F, tipo=—`. Obsérvese qué hace
fallar el paso 5: si los números de nota fueran `100` o mayores, o si la
columna llevara tres textos, la nota no se detectaría.

---

## 7. La región de datos

`detectar_region(valores, cols, cfg, n_filas)` (`src/generador_fs.py:522`)
devuelve `(primera, ultima)`.

Si `primera_fila` es un entero positivo, se usa tal cual. En `"auto"`: se llama
a `detectar_encabezado` y se toma `arranque = fin + 1` —los datos empiezan
justo detrás del bloque de cabecera—; se busca la primera fila desde ahí con
etiqueta de texto, cuya etiqueta normalizada no sea `""` ni `"$"` (descarta las
filas que solo llevan el signo de moneda) y que cumpla
`reconocido or r > 4 or _es_numero(va) or _es_numero(vp)`; si no encuentra
ninguna, `primera = arranque` si se reconoció cabecera, y `5` si no.

El `r > 4` y el `5` son **vestigios deliberados** de la heurística anterior
(«las filas 1 a 4 son cabecera»), conservados solo para cuando la detección por
contenido no reconozca nada: *"para no cambiarle el resultado a una hoja que
hoy funciona"*.

La última fila se calcula recorriendo **todas** las filas hasta `n_filas` y
guardando la última con etiqueta de texto o número en `actual` o `previo`. El
bucle no se detiene en el primer hueco: la región **tolera filas en blanco
intercaladas**, que en un estado financiero son la norma.

---

## 8. El encabezado por contenido

`detectar_encabezado(valores, cols, n_filas)` (`src/generador_fs.py:618`)
localiza en qué fila está cada dato de cabecera:

> *"Antes se leían las filas 1, 2, 3 y 4 de la hoja, en ese orden fijo. En
> cuanto alguien movía la tabla —a K14, por ejemplo— esas cuatro celdas
> quedaban vacías y el documento salía con «Al () — comparado con ()» y «Cifras
> expresadas en , en unidades de»."*

Los cuatro reconocedores, con sus patrones exactos:

| Función | Línea | Reconoce |
|---|---|---|
| `_es_fecha_cabecera(v)` | `577` | Un `datetime`/`date`; o una cadena con uno de los 24 nombres de mes de `_MESES` (12 en inglés y 12 en español, con `"setiembre"` además de `"septiembre"`); o que case `\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}`; o que contenga `\b(19|20)\d{2}\b`. |
| `_es_escala_cabecera(v)` | `593` | Un número que sea exactamente `100`, `1000`, `10000` o `1000000`; o una cadena con `thousand`, `million`, `miles`, `millones` o `millon`. |
| `_es_estado_cabecera(v)` | `603` | Cualquier cadena con la raíz `audit` (cubre `Audited`, `Unaudited`, `auditado`, `no auditado`). |
| `_es_moneda_cabecera(v)` | `608` | Una cadena de **6 caracteres o menos** tras quitar paréntesis y espacios exteriores, que case por completo con `[$€£¥]`, `[A-Za-z]{0,3}\$` (`$`, `US$`, `CDN$`) o uno de `USD COP CAD EUR MXN GBP`. |

El bucle va de la fila 1 hacia abajo y clasifica con una cadena `if/elif` de
**prioridad fija**: fechas, escala, estado, moneda. La escala exige además que
la fila **no tenga etiqueta** —*"si no, un importe de 1.000 en una linea de
detalle se confundiria con la fila de la unidad"*— y solo se busca en la
columna actual. De cada clase gana la **primera** fila que la dispare; `fin`
avanza hasta la última clasificada. El corte es explícito:

```python
if _texto_no_vacio(etq) and (_es_numero(va) or _es_numero(vp)):
    break
```

> *"Una fila con etiqueta Y cifras de verdad ya es un dato: el bloque de
> cabecera se acabó. Las filas de seccion (texto sin cifras) no cortan: pueden
> ir intercaladas antes del primer importe."*

`fin == 0` significa «no reconocí nada», y es la señal con la que
`detectar_region` vuelve a la heurística antigua.

`leer_encabezado(valores, cols, cfg, primera, hdr=None)`
(`src/generador_fs.py:665`) recoge las celdas y produce las ocho claves de
cabecera del `ctx`:

| Clave | De dónde sale | Forma |
|---|---|---|
| `empresa` | `cfg["empresa"]` | texto; no viene del Excel |
| `titulo` | primer texto de la columna de etiqueta **por encima** de la región | crudo |
| `fecha_actual` | fila `fechas`, columna actual | **crudo**: puede ser `datetime` |
| `fecha_previa` | fila `fechas`, columna previa | crudo |
| `miles` | fila `escala`, columna actual | crudo (p. ej. el entero `1000`) |
| `estado_actual` | fila `estado`, columna actual, por `texto()` | sin paréntesis: `Unaudited` |
| `estado_previo` | fila `estado`, columna previa, por `texto()` | ídem |
| `moneda` | fila `moneda`, columna actual | crudo, **conserva** los paréntesis |

La asimetría es deliberada: `(Unaudited)` debe llegar al Word sin paréntesis;
`(US$)` conserva su forma. En el ejemplo de [§6.3](#63-un-ejemplo-concreto),
`titulo` sería `Collective Mining Ltd.` —el primer texto de la columna `C` por
encima de la fila 12—, no `Statement of Financial Position`.

---

## 9. Clasificar cada fila

`inferir_tipo(vals, bold, cols, cfg)` (`src/generador_fs.py:690`) recibe una
fila materializada y devuelve `(tipo, señal)`, donde la señal es la frase que
explica la decisión y acaba en el CSV de revisión.

| Tipo | Significa | Cómo se reconoce |
|---|---|---|
| `H` | encabezado de sección | etiqueta sin cifras que además esté en mayúsculas, termine en `:`, vaya en negrita **o** mida 40 caracteres o menos |
| `I` | línea de detalle | etiqueta con cifras o con nota |
| `S` | subtotal (sin etiqueta) | cifras sin etiqueta |
| `T` | total | la etiqueta normalizada empieza por `total` |
| `N` | nota de texto libre | etiqueta sin cifras, de **más de 40 caracteres**, sin mayúsculas, sin `:` final y sin negrita |
| `X` | excluir a propósito | un marcador de `marcadores_excluir` aparece en la etiqueta o en la nota |

Las reglas se evalúan **en orden estricto** y decide la primera que dispara:
(1) exclusión, que gana a todo, incluso a `Total`; (2) fila totalmente vacía,
que devuelve `None`; (3) `total` al principio de la etiqueta, tenga o no
cifras; (4) cifras sin etiqueta; (5) etiqueta sin cifras, que da `H` con alguna
de las cuatro señales y `N` sin ninguna; (6) etiqueta con cifras o nota.

Tres precisiones que cambian el resultado: `hay_valor` es cierto también con un
**marcador de cero** —`_marcador_cero` (`src/generador_fs.py:205`) acepta `-`,
`–`, `—`, `−`, `ꟷ`, `‑`, `·` y la cadena `"0"`—, de modo que una línea de
detalle en cero no se degrada a encabezado; `en_negrita` mira las **tres**
columnas (actual, previo y etiqueta); y `TIPOS_VALIDOS`
(`src/generador_fs.py:114`) es `{"H","I","S","T","N"}`: **`X` no está**, porque
no produce línea sino descarte con registro.

Por qué es heurística y qué la hace fallar. La regla 5 convierte `H` en el caso
por defecto de todo texto sin cifras, y el corte de 40 caracteres es
arbitrario: una nota corta se clasifica como encabezado y un encabezado largo
como nota. La regla 3 ata el tipo `T` a la palabra literal `total`, así que un
`Suma del activo` no se reconoce como total. Y la regla 1 compara por
subcadena: un marcador corto como `check` excluye cualquier fila que lo
contenga.

La cura es determinista y está en manos de quien mantiene el Excel: **añadir
una columna `Tipo`** con las letras `H/I/S/T/N/X`. Si la fila trae un tipo
declarado y válido, se usa tal cual y el CSV lo registra con
`origen = "declarado"`. La columna se detecta por su rótulo, esté donde esté, y
es opcional: *"Ya NO se aborta solo porque falte la columna 'Tipo'"*. Cuando la
hoja no la tiene, la consola avisa de que **todos** los tipos se infirieron.

---

## 10. Armar las líneas

`construir_lineas(valores, negrita, cols, cfg, primera, ultima, hay_col_tipo)`
(`src/generador_fs.py:734`) recorre la región y devuelve
`(lineas, revision, sin_tipo, tipo_invalido, n_declarados, n_inferidos)`.

| Caso de la fila | Qué hace |
|---|---|
| `Tipo` declarado `X` | registra `"excluido a propósito"` en el CSV; no genera línea |
| `Tipo` declarado válido | lo usa; `origen = "declarado"` |
| `Tipo` declarado no reconocido | lo anota en `tipo_invalido` e infiere; si el inferido es `None` o `X`, descarta con registro |
| `Tipo` en blanco | infiere; si sale `None`, **descarta en silencio**; si sale `X`, descarta con registro; si no, `origen = "inferido"` y, **solo si la hoja tiene columna `Tipo`**, anota la fila en `sin_tipo` |

Esa última asimetría es deliberada: una fila dejada en blanco en una hoja que sí
tiene la columna es un olvido reportable; si la hoja no la tiene, el aviso
general de la consola ya cubre el caso.

La línea lleva ocho claves —`tipo`, `etiqueta`, `nota`, `actual`, `previo`,
`fila`, `actual_raw`, `previo_raw`—, y las dos últimas existen por una razón
concreta:

> *"*_raw: el valor sin formatear. Lo usa fs_contrato para calcular variaciones
> (var_abs / var_pct). docxtpl ignora las claves de más."*

`escribir_revision(revision)` (`src/generador_fs.py:793`) escribe
`salidas/revisar_tipos.csv` con la cabecera
`fila;etiqueta;nota;actual;previo;tipo;origen;señal`. Las tres decisiones de
formato apuntan a Excel en español: delimitador `";"`, codificación
`utf-8-sig` (con BOM, para que no se rompan las tildes) y `newline=""` para no
duplicar el salto de línea en Windows. El archivo se **sobrescribe** en cada
ejecución, y se escribe siempre, también al generar el documento.

---

## 11. El formato de los números

Hay dos formateadores, conscientemente distintos.

`num(valor)` (`src/generador_fs.py:165`) es el de la tabla del estado:
separador de miles `,`, **cero decimales**, negativos entre paréntesis
—`(142,204,537)`— y una regla que importa: si el valor es una cadena se
devuelve **tal cual**, de modo que los marcadores de cero llegan intactos.

`texto(valor)` (`src/generador_fs.py:179`) hace `str(valor).strip("() ")`.
`sanear(texto_)` (`src/generador_fs.py:185`) aplica una lista blanca estricta
—`re.sub(r"[^A-Za-z0-9_\-]", "_", …)` truncado a 40 caracteres— y se usa al
componer el nombre del `.docx` a partir de `ctx["fecha_actual"]`: neutraliza
`/`, `\`, `..` y `:`, porque el contenido del Excel es entrada no confiable y no
puede convertirse en una ruta.

`_decimales_de(formato)` (`src/generador_fs.py:808`) cuenta los `0` o `#` que
siguen al punto decimal de un formato de Excel (`'0.00'` da 2) y devuelve
`None` si no hay parte decimal.

`formatear_valor(valor, formato=None)` (`src/generador_fs.py:814`) es el
segundo formateador:

> *"Es distinto de num(): num() es para las cifras de la tabla del estado,
> siempre en miles y sin decimales. Aquí entran ratios, tipos de cambio,
> porcentajes y fechas, donde los decimales importan. Si el usuario puso '0.00'
> en Excel, aquí salen dos decimales; si puso '0.0%', sale con el signo de
> porcentaje."*

Su cascada: `None` da `""`; un `bool` da `"Sí"`/`"No"` —comprobado antes que
`int`, porque en Python `bool` es subclase de `int`—; una cadena se devuelve
recortada; cualquier objeto con `strftime` se formatea como `%Y-%m-%d`; y un
número se formatea con los decimales que pida el formato de la celda,
multiplicando por 100 y añadiendo `%` si el formato lleva `%`, y usando como
respaldo cero decimales para los enteros y dos para el resto.

Respetar el formato de la celda no es un capricho: quien administra el libro ya
decidió allí cuántos decimales tienen sentido para un tipo de cambio o un
porcentaje, y esa decisión debe llegar al documento sin volver a tomarse en el
código. `formatear_valor` se usa **solo** para los escalares de los rangos con
nombre; la tabla del estado nunca pasa por él.

---

## 12. La identidad estable

Sin rangos con nombre, una línea se identifica por el texto de su etiqueta.
Basta cambiar `Cash and cash equivalents` por `Cash & equivalents` para que la
clave pase de `cash_and_cash_equivalents` a `cash_equivalents` y la cifra que
ese texto alimentaba quede **huérfana**: un ancla presente en el documento para
la que el Excel ya no da valor.

El remedio es un nombre de Excel apuntando a la celda de etiqueta de la fila:

> *"un nombre de Excel \"fs_total_assets\" apunta a la celda de etiqueta de esa
> fila y le da una identidad ESTABLE. Si alguien renombra la fila o inserta
> filas encima, el nombre sigue apuntando a la misma línea y el vínculo con el
> Word no se rompe."*

`leer_rangos_con_nombre(wb, hoja_titulo, cfg)` (`src/generador_fs.py:862`)
recorre los nombres definidos del **libro entero** y devuelve `por_fila`
—`{fila: clave}` para los nombres que caen en la hoja del estado— y
`escalares` —`{clave: (texto_formateado, valor_crudo)}` para todo lo demás: una
celda de otra hoja, un ratio, un tipo de cambio, una fecha de corte—.

Detalles que hay que conocer: el filtro del prefijo `fs_` es insensible a
mayúsculas y la clave se pasa a minúsculas; un `prefijo_rangos` vacío desactiva
el mecanismo entero; los nombres rotos (`#REF!`, lo que queda cuando alguien
borra la fila) se saltan sin ruido; de un nombre que apunte a un rango se usa
**solo la esquina superior izquierda**, y de un nombre con varios destinos solo
el primero; y si dos nombres caen en la misma fila, gana el primero en el orden
de iteración.

`_leer_celda(wb, hoja, fila, columna)` (`src/generador_fs.py:845`) resuelve un
detalle de openpyxl: *"En modo read_only no hay acceso aleatorio con
ws.cell(), así que se pide esa fila con iter_rows."* Devuelve el valor y el
`number_format`, que es lo que consume `formatear_valor`.

En `leer_contexto` (`src/generador_fs.py:1052`) se hace el enlace: la línea
cuya fila aparezca en `por_fila` recibe `clave` y `clave_origen = "rango"`; el
resto solo recibe `clave_origen = "etiqueta"` —**sin** clave— y el consumidor
la deriva del texto. Después, las claves enlazadas a una línea se excluyen de
los escalares; un `fs_*` que caiga en la hoja del estado pero **fuera** de la
región de datos sobrevive como escalar.

Ese reparto es exactamente lo que formaliza `clave_de_linea` del contrato
(`src/fs_contrato.py:163`): `linea.get("clave") or clave(linea.get("etiqueta"))`.
El rango manda; la etiqueta es el respaldo. Los nombres se crean con la orden
`nombrar` del módulo del documento, que pilota Excel; nunca con openpyxl, por
la razón de siempre.

---

## 13. El contexto (`ctx`)

`leer_contexto(ruta_xlsx, cfg)` (`src/generador_fs.py:1007`) devuelve doce
claves: ocho nacen en `leer_encabezado` y cuatro se añaden después.

| Clave | Tipo | Contenido |
|---|---|---|
| `empresa` | `str` | de `cfg["empresa"]` |
| `titulo` | `str` | primer texto de la columna de etiqueta por encima de la región |
| `fecha_actual` | crudo | `datetime`, `str` o `""` — **sin formatear** |
| `fecha_previa` | crudo | ídem |
| `miles` | crudo | `int`, `float`, `str` o `""` |
| `estado_actual` | `str` | sin paréntesis ni espacios exteriores |
| `estado_previo` | `str` | ídem |
| `moneda` | crudo | `str` o `""`, **con** paréntesis si los tenía |
| `lineas` | `list[dict]` | las líneas del estado (forma abajo) |
| `escalares` | `dict[str, tuple[str, Any]]` | `{clave: (texto_formateado, valor_crudo)}` |
| `_meta` | `dict` | 9 claves de diagnóstico |
| `_avisos` | `dict` | 3 claves de revisión |

Forma exacta de **un elemento** de `ctx["lineas"]`:

| Clave | Tipo | Contenido |
|---|---|---|
| `tipo` | `str` | `"H"`, `"I"`, `"S"`, `"T"` o `"N"`. Nunca `"X"`: esas filas no llegan. |
| `etiqueta` | crudo | el valor de la celda, sin pasar por `texto()` |
| `nota` | `str` | `texto(…)`, o `""` si no hay columna de nota |
| `actual` | `str` | `num(…)`: miles con `,`, sin decimales, negativos entre paréntesis |
| `previo` | `str` | `num(…)`, o `""` si no hay columna comparativa |
| `fila` | `int` | número de fila 1-based en la hoja |
| `actual_raw` | `int\|float\|str\|None` | el valor sin formatear |
| `previo_raw` | `int\|float\|str\|None` | ídem |
| `clave_origen` | `str` | `"rango"` o `"etiqueta"` |
| `clave` | `str` | **solo si `clave_origen == "rango"`**: el nombre `fs_*` sin prefijo, en minúsculas |

`_meta` lleva `hoja`, `como_hoja` (la frase que justifica la elección),
`columnas` (las **letras**, con `"—"` para las no detectadas), `region`
(la tupla `(primera, ultima)`), `hay_col_tipo`, `n_declarados`, `n_inferidos`,
`n_con_rango` y `n_escalares`. `_avisos` lleva `sin_tipo` (lista de filas),
`tipo_invalido` (pares `(fila, valor)`) y `revision` (las 8-tuplas del CSV).
Ambas llevan guion bajo porque `ejecutar` las extrae con `pop` antes de
renderizar: no deben llegar a la plantilla.

El mismo contexto, como JSON comentado:

```jsonc
{
  "empresa":       "Collective Mining Ltd.",   // de config.json
  "titulo":        "Collective Mining Ltd.",   // primer texto sobre la región
  "fecha_actual":  "June 30,",                 // crudo: puede ser un datetime
  "fecha_previa":  "December 31,",
  "miles":         1000,                       // crudo: aquí un entero
  "estado_actual": "Unaudited",                // ya sin paréntesis
  "estado_previo": "Audited",
  "moneda":        "US$",                      // crudo: conserva paréntesis

  "lineas": [
    { "tipo": "H", "etiqueta": "ASSETS", "nota": "",
      "actual": "", "previo": "", "fila": 12,
      "actual_raw": null, "previo_raw": null,
      "clave_origen": "etiqueta" },            // sin "clave": se deriva del texto

    { "tipo": "I", "etiqueta": "Cash and cash equivalents", "nota": "5",
      "actual": "119,066,301", "previo": "96,094,583", "fila": 14,
      "actual_raw": 119066301, "previo_raw": 96094583,
      "clave_origen": "rango",
      "clave": "cash_and_cash_equivalents" }   // viene del nombre fs_...
  ],

  "escalares": {
    // clave -> [texto ya formateado por formatear_valor, valor crudo]
    "tipo_cambio": ["4,102.35", 4102.35],
    "fecha_corte": ["2025-06-30", "2025-06-30T00:00:00"]
  },

  "_meta": {
    "hoja": "Balance 2025",
    "como_hoja": "'Balance 2025' elegida por contenido (3.5 señales)",
    "columnas": { "etiqueta": "C", "nota": "D", "actual": "E",
                  "previo": "F", "tipo": "—" },
    "region": [12, 16],
    "hay_col_tipo": false,
    "n_declarados": 0, "n_inferidos": 5,
    "n_con_rango": 1, "n_escalares": 2
  },

  "_avisos": {
    "sin_tipo": [],                            // solo se puebla si hay columna Tipo
    "tipo_invalido": [],
    "revision": [ [12, "ASSETS", "", "", "", "H", "inferido",
                   "etiqueta sin cifras (encabezado)"] ]
  }
}
```

---

## 14. El contrato de anclas

`src/fs_contrato.py` no toca ningún archivo. Define el vocabulario que
comparten los dos lados del sistema:

> *"Es la misma especificación que consume el add-in de Word. Ver
> CONTRATO.md."*

Eso es lo esencial del módulo: es la **frontera compartida**. Cada constante y
cada función tiene su gemela en TypeScript (`addin/src/core/contrato.ts`), y
una divergencia entre las dos implementaciones no produce un error, produce
anclas huérfanas.

### 14.1 Las seis familias

Un ancla es la cadena `fs-…` que va en la etiqueta (`w:tag`) de un control de
contenido: la caja etiquetada que Word sabe delimitar dentro de un documento.

| Ancla | Familia | ¿La escribe el refresco? |
|---|---|---|
| `fs-tabla-<nombre>` | bloque: una tabla completa | Sí, reescribe la tabla entera |
| `fs-campo-<nombre>` | en línea: un campo de encabezado | Sí |
| `fs-dato-<clave>-<campo>` | en línea: una cifra suelta dentro de prosa | Sí |
| `fs-prosa-<nombre>` | bloque de redacción libre | **No**, ni se visita |
| `fs-registro` | bloque de bitácora | Antepone, no reemplaza |
| `fs-meta` | párrafo oculto con la foto JSON de la última corrida | Sobrescribe |

La regla de oro, textual (`src/fs_contrato.py:21`):

> *"el refresco SOLO escribe dentro de anclas de las familias tabla / campo /
> dato / registro / meta. Todo lo demás del documento —incluida la prosa que la
> persona redacte— no se visita siquiera."*

Los nombres se componen con `tag_tabla`, `tag_campo`, `tag_dato(clave, campo)`
y `tag_prosa` (`src/fs_contrato.py:68`–`81`). Los ocho campos de encabezado
válidos están en `CAMPOS_ENCABEZADO` (`src/fs_contrato.py:43`) y coinciden con
las ocho claves de cabecera del `ctx`; los cinco sufijos de una cifra suelta,
en `CAMPOS_DATO` (`src/fs_contrato.py:55`): `actual`, `previo`, `nota`,
`var_abs`, `var_pct`.

`descomponer(tag)` (`src/fs_contrato.py:84`) es la inversa:
`'fs-dato-total_assets-actual'` da `('dato', 'total_assets', 'actual')`. Un tag
ajeno al contrato da `(None, None, None)`, y ese es el mecanismo por el que un
control de contenido de otro uso queda intocado. En la familia `dato`, el
**último** segmento es el campo y todo lo de en medio es la clave; si el campo
no está en `CAMPOS_DATO`, el tag se rechaza entero. `es_region_de_datos(tag)`
(`src/fs_contrato.py:114`) responde a «¿el refresco debe escribir aquí?» y solo
es cierto para `tabla`, `campo` y `dato`: `registro` y `meta` se escriben por
caminos propios.

### 14.2 `clave()`: los cinco pasos

`clave(etiqueta)` (`src/fs_contrato.py:123`) convierte el texto de una fila en
un identificador estable, y su docstring empieza con la exigencia que define
todo el contrato: **«Debe dar el MISMO resultado en Python y en TypeScript»**.

1. Quitar tildes: normalizar a NFKD y descartar los diacríticos.
2. Pasar a minúsculas.
3. Todo lo que no sea `[a-z0-9]` se convierte en `_`.
4. Colapsar los `_` repetidos y recortar los de los extremos.
5. Cortar a 40 caracteres y volver a recortar los `_` finales.

Los ejemplos del propio código:

```
'Cash and cash equivalents'  ->  cash_and_cash_equivalents
'Total assets'               ->  total_assets
'Provisión (neta)'           ->  provision_neta
```

(Inferencia: el corte a 40 caracteres del paso 5 es lo que garantiza el límite
de 64 que Word impone a una etiqueta y que el módulo declara en `MAX_TAG` —
`fs-dato-` más 40 más `-var_abs` suman 56—. `MAX_TAG` no se valida en ninguna
parte del código.)

`clave_de_linea(linea)` (`src/fs_contrato.py:163`) es la que se usa en la
práctica: prefiere el rango con nombre y solo cae en la etiqueta si no lo hay.

### 14.3 `construir_valores(ctx)`

`construir_valores(ctx)` (`src/fs_contrato.py:173`) traduce el contexto al mapa
plano `{tag: texto}` que el motor escribe región por región; no incluye las
tablas, que el motor arma a partir de `ctx["lineas"]`. Devuelve
`(valores, colisiones)` y hace tres pasadas:

1. **Campos de encabezado.** Un `fs-campo-<nombre>` por cada uno de los ocho
   nombres de `CAMPOS_ENCABEZADO`, con `str(v)`.
2. **Escalares.** Cada escalar se expone como `fs-dato-<clave>-actual` con el
   texto que ya formateó `formatear_valor`, y los otros cuatro campos se
   rellenan con `""` mediante `setdefault`, para que un ancla
   `fs-dato-<clave>-previo` sobre un escalar no quede sin entrada.
3. **Líneas.** Por cada línea se escriben `actual`, `previo` y `nota`, y se
   **calculan** las variaciones: `var_abs` es `actual_raw − previo_raw` con el
   mismo formato contable que `num()`, y `var_pct` es
   `(actual − previo) / previo × 100` con un decimal y el signo `%`. Ambas
   exigen que los dos valores crudos sean numéricos; si alguno no lo es —una
   cadena, un marcador de cero, `None`— salen vacías, y si el previo es cero,
   `var_pct` sale vacía en lugar de dividir por cero.

**Las colisiones.** Dos etiquetas distintas pueden producir la misma clave
(`Otros activos` y `Otros Activos:` dan las dos `otros_activos`). El contrato
conserva **la primera aparición** y devuelve la lista de colisiones con la
etiqueta nueva, la primera y la clave compartida. No lo resuelve solo: lo
reporta, y así consta como deuda conocida en [`../CONTRATO.md`](../CONTRATO.md).

`catalogo(ctx)` (`src/fs_contrato.py:225`) produce la lista legible para el
panel y los informes —`[(clave, origen, etiqueta, actual, previo), …]`, en el
orden del Excel, más los escalares que no correspondan a una línea, ordenados
alfabéticamente y con la etiqueta `"(celda suelta del libro)"`—. El campo
`origen` es el que le dice a una persona qué claves son frágiles: `rango` es
identidad estable; `etiqueta` se rompe si alguien renombra la fila.

---

## 15. El generador clásico

`ejecutar(argv)` (`src/generador_fs.py:1136`) es el camino original: Excel más
plantilla, con `docxtpl`, hacia un `.docx` nuevo. El parseo de la línea de
órdenes es posicional puro, sin `argparse`: se separan los argumentos que
empiezan por `--` del resto, de modo que las banderas pueden ir en cualquier
posición.

| Argumento | Efecto |
|---|---|
| primer posicional | ruta del `.xlsx` |
| *(ninguno)* | `encontrar_excel_por_convencion`: el `.xlsx` más reciente de la raíz o de `ejemplos\` cuyo nombre contenga `FS`, excluidos los archivos de bloqueo `~$` |
| segundo posicional | ruta de la plantilla `.docx` |
| *(ninguno)* | `buscar_recurso`: junto al `.exe` o en la raíz primero, luego `plantillas\`, y por último la embebida |
| `--revisar` | **no genera el Word**: solo escribe `revisar_tipos.csv` y devuelve `None`; además omite la validación de que exista la plantilla |
| cualquier otra `--xxx` | se ignora en silencio |

El cuerpo lee el contexto, extrae `_meta` y `_avisos` con `pop`, escribe siempre
el CSV de revisión y —si no es `--revisar`— renderiza la plantilla con un
`SandboxedEnvironment(autoescape=True)` de Jinja. Ese entorno restringido no es
decorativo: el contenido del Excel es entrada no confiable y acaba dentro de una
plantilla. Después estampa en las propiedades del documento una línea de
trazabilidad con el nombre del origen, los primeros 12 dígitos del SHA-256 del
`.xlsx`, la hoja usada, la plantilla y el sello de tiempo, y guarda en
`salidas\estado_situacion_financiera_<fecha saneada>_<AAAAMMDD-hhmm>.docx`. Por
último imprime un informe con la hoja elegida y su justificación, las columnas
en letras, la región de datos, el recuento de líneas declaradas e inferidas, la
ruta del CSV y los avisos de tipos que faltan o no se reconocen.

**Por qué es una foto y no un documento vivo.** El resultado es un `.docx`
generado entero desde una plantilla, sin regiones dentro. No conserva nada de lo
que una persona hubiera escrito, porque no hay corrida anterior con la que
compararse: cada ejecución produce un archivo distinto, con su propio sello de
tiempo, en `salidas\`. Sirve para mirar las cifras y verificar que la lectura
del Excel es correcta. Todo lo que tenga que ver con conservar la redacción
humana pertenece al documento vivo y al refresco, que se apoyan en el mismo
`ctx` pero por el camino del contrato de anclas.

---

## 16. Puntos frágiles de esta capa

- **El orden de las columnas de cifras es una convención implícita.** Se asume
  que el periodo actual está a la izquierda de la comparativa. Un libro que las
  invierta produce un documento con los periodos cambiados y sin ningún aviso;
  solo se corrige forzando letras en `config.json -> "columnas"`.
- **Los marcadores de hoja por defecto se solapan.** `"assets"` es subcadena de
  `"total assets"`: una sola frase contable alcanza el `UMBRAL_HOJA` de 2. El
  umbral es más laxo de lo que el número sugiere.
- **`_texto_de_hoja` une las celdas con espacios**, de modo que un marcador de
  varias palabras puede formarse por concatenación de dos celdas contiguas.
  Tolerante, y también fuente de falsos positivos.
- **`_es_fecha_cabecera` reconoce `"may"` como subcadena**, así que `"mayor"` o
  `"maybe"` marcan una fila como fila de fechas.
- **`titulo` es el primer texto que haya encima de la región**, sea cual sea:
  con frecuencia acaba siendo el nombre de la empresa y no el del estado.
- **La detección de la columna de nota es estrecha**: exige enteros en `(0, 99]`
  y como mucho dos textos en toda la columna. Una nota `100` la vuelve
  invisible.
- **El corte de 40 caracteres entre `H` y `N` es arbitrario**, igual que la
  dependencia de `T` respecto de la palabra literal `total`. Sin columna `Tipo`,
  ambas fallan de forma predecible en libros con otra redacción.
- **Las banderas desconocidas se ignoran en silencio**: un `--revizar` mal
  escrito genera el documento sin avisar.
- **El sello de tiempo tiene resolución de minuto.** Dos ejecuciones dentro del
  mismo minuto, con la misma `fecha_actual`, escriben sobre el mismo archivo.
- **La configuración se reemplaza clave por clave, sin mezcla de listas.** Quien
  redefina `hoja_marcadores` para añadir uno los sustituye todos.

---

## Resumen del capítulo

- Nada se localiza por coordenadas: hoja, columnas, cabecera, región y tipo de
  fila se deciden por contenido, y cada decisión inferida queda registrada en
  `salidas/revisar_tipos.csv`.
- El nombre de hoja de `config.json` es una pista que se verifica: se acepta
  solo si la hoja puntúa `UMBRAL_HOJA = 2` o más.
- `max_filas_scan` y `max_cols_scan` son **suelos**, no topes; los techos duros
  de 20 000 filas y 256 columnas no son configurables.
- Solo la columna `actual` es obligatoria; forzar letras en `config.json` es la
  válvula de escape cuando la detección no acierta.
- La columna `Tipo` de la hoja es opcional, pero convierte la clasificación de
  heurística en determinista.
- Hay dos formateadores a propósito: `num()` para la tabla del estado y
  `formatear_valor()` para los escalares, que respeta el formato de la celda.
- Los rangos con nombre `fs_*` dan identidad estable a una fila; sin ellos la
  clave se deriva de la etiqueta y renombrar la fila rompe el vínculo.
- `fs_contrato.py` es la frontera compartida con el complemento de Office:
  `clave()` debe dar el mismo resultado en Python y en TypeScript, o aparecen
  anclas huérfanas sin que nada falle.
