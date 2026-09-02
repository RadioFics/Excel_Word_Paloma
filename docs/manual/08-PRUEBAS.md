# 08 · El banco de pruebas

> **Para quién.** Quien vaya a modificar el motor y necesite saber qué red
> tiene debajo, y quien tenga que justificar ante alguien que el programa no
> destruye documentos.
> **Qué encontrará.** El criterio con el que se prueba, la maquinaria del
> banco, el catálogo completo de las 27 comprobaciones con el invariante que
> defiende cada una, la prueba de humo del entorno, y una lista honesta de lo
> que no está cubierto.
> **Antes de leer.** Conviene tener presente qué es una región y qué hace un
> refresco: está en [El motor del documento](04-MOTOR-DEL-DOCUMENTO.md).

## Índice del capítulo

1. [Qué se prueba y por qué así](#1-qué-se-prueba-y-por-qué-así)
2. [Cómo se ejecuta](#2-cómo-se-ejecuta)
3. [La maquinaria](#3-la-maquinaria)
4. [El catálogo de pruebas](#4-el-catálogo-de-pruebas)
5. [La pasada con el libro real](#5-la-pasada-con-el-libro-real)
6. [La prueba de humo del entorno](#6-la-prueba-de-humo-del-entorno)
7. [Qué no está cubierto](#7-qué-no-está-cubierto)
8. [Cómo añadir una prueba](#8-cómo-añadir-una-prueba)

---

## 1. Qué se prueba y por qué así

Este programa escribe sobre documentos reales que viven en OneDrive y que
contienen trabajo de una persona. El riesgo que hay que descartar no es que una
cifra salga mal —eso se ve—, sino que el documento se corrompa despacio, o que
la redacción se pierda sin que nadie lo note hasta semanas después.

El docstring del banco plantea la pregunta en esos términos:

> «Responde a una pregunta concreta y repetible: si alguien añade una línea al
> Excel, si la quita, si la renombra o si mete filas por encima, ¿el documento
> de Word sale bien, y sigue siendo un .docx sano después de decenas de
> actualizaciones seguidas?»

De ahí salen las tres decisiones de método:

1. **Nada real se toca.** Cada prueba trabaja sobre copias en una carpeta
   temporal que se borra al terminar. El libro de Excel y el documento de Word
   de producción no intervienen; el banco se fabrica los suyos.
2. **El libro de prueba es sintético pero realista.** `FILAS_BASE`
   (`tools/probar_refresco.py:58`) es un estado de situación financiera en
   miniatura con la misma forma que el real: encabezados de sección, líneas de
   detalle, subtotales sin etiqueta y un total. Diez filas bastan para ejercitar
   los seis tipos de línea.
3. **Se verifica después de CADA escritura**, no solo al final. Es la única
   forma de localizar el refresco concreto que estropeó el archivo.

El encabezado del libro sintético se define como **desplazamientos respecto de
la esquina de la tabla**, no como celdas fijas, precisamente porque varias
pruebas mueven la tabla por la hoja a propósito.

---

## 2. Cómo se ejecuta

```bash
python tools\probar_refresco.py                   todas las pruebas
python tools\probar_refresco.py --libro X.xlsx    además, una pasada con un libro real
python tools\probar_refresco.py --verboso         enseña la traza completa de cada fallo
```

| Bandera | Efecto |
|---|---|
| `--libro <ruta>` | Añade al final una pasada sobre un `.xlsx` de verdad. Solo lo lee. |
| `--verboso` | Ante un error inesperado, imprime la traza completa además del mensaje. |

Desde el explorador, `probar.bat` hace lo mismo: localiza un intérprete de
Python con `tools\buscar_python.bat` y lanza el banco. Si se le arrastra un
`.xlsx` encima, lo pasa como `--libro`. Deja la ventana abierta al terminar
para que se pueda leer el resultado.

El programa devuelve `0` si todo pasó y `1` si algo falló, de modo que sirve
tal cual en una comprobación automatizada.

---

## 3. La maquinaria

### 3.1 Fabricar el material

| Función | Línea | Qué hace |
|---|---|---|
| `escribir_libro(ruta, filas, origen=(1,1))` | `:98` | Crea un `.xlsx` con esas filas y la tabla anclada en la celda `origen`. Es lo que permite mover la tabla por la hoja. |
| `poner_rangos(ruta, nombres)` | `:128` | Añade rangos con nombre al libro, para probar la identidad estable. |
| `contexto_de(ruta, cfg)` | `:148` | Lee el libro y devuelve el contexto, descartando los avisos. |

### 3.2 La clase `Banco`

`tools/probar_refresco.py:233`. Su docstring la describe con exactitud: *«Un
documento y un libro recién hechos, listos para maltratarlos.»*

Al construirse: carga la configuración y **desactiva la bitácora**
(`cfg["bitacora"] = "no"`, con el comentario *«sin .log: ensucia y no se mide
aquí»*), escribe el libro sintético, crea un `.docx` en blanco, le monta el
andamiaje con `construir()` y —salvo que se pida lo contrario— le siembra
redacción humana.

Esa siembra es el núcleo del método. `_sembrar_prosa` añade dos párrafos:

```
Texto de Pamela que no se debe tocar jamás.
Segundo párrafo, con acentuación: gestión, análisis.
```

Y su docstring dice por qué:

> «Es el corazón de la prueba: esta redacción NO puede cambiar nunca, por
> muchas veces que se refresque ni por mucho que cambie el Excel.»

El segundo párrafo lleva acentos a propósito: es también una prueba de
codificación.

El método `Banco.refrescar(filas=None, origen="prueba")` reescribe el libro con
otras filas si se le pasan, relee el contexto y refresca el documento. Toda la
batería se apoya en esa única operación.

### 3.3 La radiografía

`radiografia(docx)` (`:176`) extrae de un documento todo lo que hace falta para
detectar que algo se estropeó: la lista de etiquetas de región, cuántas hay,
las filas de tabla como listas de textos de celda, el contenido crudo de
`fs-meta`, los párrafos de redacción, el texto completo y el tamaño en bytes.

Un detalle de implementación que merece la pena conocer, porque es una trampa
clásica: `_texto(el)` (`:157`) baja al XML en vez de usar los accesores
cómodos de la biblioteca. El comentario lo explica:

> «Hace falta bajar al XML: doc.paragraphs y doc.tables de python-docx solo
> miran los hijos DIRECTOS del cuerpo, y aquí casi todo vive dentro de un
> control de contenido (w:sdt). Buscar con ellos da listas vacías y las
> comprobaciones pasan sin comprobar nada.»

Es decir: una prueba escrita de la manera obvia habría pasado siempre, sin
comprobar nada.

### 3.4 `comprobar_sano(docx, etiqueta="")`

`:207`. Son *«las invariantes que deben cumplirse SIEMPRE, pase lo que pase»*.
Se ejecuta después de cada escritura y comprueba cuatro cosas:

| # | Invariante | Qué fallo detecta |
|---|---|---|
| 1 | El archivo sigue siendo un ZIP válido y se puede abrir | Escritura interrumpida, archivo a medias, corrupción. |
| 2 | No hay etiquetas de región duplicadas | Un refresco que **inserta** en lugar de **reemplazar**. Dos regiones con el mismo nombre se pisarían: solo se refrescaría la primera. |
| 3 | `fs-meta` sigue siendo JSON legible | La memoria del documento se rompió; el siguiente diff sería inútil. |
| 4 | El documento no se quedó sin ninguna región | Un borrado accidental del andamiaje. |

A eso las pruebas individuales añaden la comprobación que las define: que la
redacción siga palabra por palabra como estaba, y que el número de regiones no
crezca solo.

### 3.5 El registro de pruebas

`@prueba("título")` (`:276`) es un decorador que anexa la función a una lista
global. El ejecutor recorre esa lista en orden, da a cada prueba su propia
carpeta temporal numerada, y clasifica el resultado en tres categorías:

- `BIEN` con el detalle que devuelva la función;
- `FALLO` si se incumplió una afirmación —un fallo previsto y explicado—;
- `ERROR` si saltó cualquier otra excepción, que es un defecto del banco o del
  motor.

Es la misma convención de errores que usa el motor: lo previsto se explica, lo
imprevisto se muestra entero.

---

## 4. El catálogo de pruebas

Son 27, agrupadas por lo que defienden.

### 4.1 La tabla puede estar en cualquier sitio

| # | Título de la prueba | Invariante que defiende |
|---|---|---|
| 1 | El encabezado (fechas, moneda, escala) se lee con la tabla en A1 | El caso base: los cuatro reconocedores del encabezado funcionan. |
| 2 | La MISMA tabla movida a K14 se lee exactamente igual | Nada se localiza por coordenadas. Es la promesa central de la capa de lectura. |
| 3 | La columna Tipo se detecta aunque quede lejos a la derecha | El rótulo se busca en toda la altura y anchura, no en una posición fija. |
| 4 | La hoja se encuentra por su ESTRUCTURA aunque la renombren | El nombre de hoja de la configuración es una pista que se verifica, no una orden. |

### 4.2 El refresco no desgasta el documento

| # | Título de la prueba | Invariante que defiende |
|---|---|---|
| 5 | Refrescar dos veces seguidas no cambia nada la segunda vez | Idempotencia. Si fallara, cada refresco iría acumulando basura. |
| 6 | Veinte refrescos seguidos no hinchan ni corrompen el documento | Que el desgaste no sea acumulativo. Es la prueba que descarta el riesgo real: la corrupción lenta. |
| 15 | Un documento con la redacción borrada sigue refrescándose | El motor no depende de que la prosa exista; si se borra, se queda borrada y no pasa nada. |

### 4.3 Los cambios en el Excel llegan bien

| # | Título de la prueba | Invariante que defiende |
|---|---|---|
| 7 | Cambiar una cifra en el Excel cambia esa cifra y solo esa | Que el refresco sea quirúrgico. |
| 8 | Añadir una línea al Excel la añade a la tabla del Word | La tabla sigue al libro hacia arriba. |
| 9 | Quitar una línea del Excel la quita de la tabla del Word | La tabla sigue al libro hacia abajo. |
| 14 | Cambiar el orden de las filas no descoloca las cifras | El emparejamiento del diff es por nombre, no por posición. |

### 4.4 La identidad de una fila

| # | Título de la prueba | Invariante que defiende |
|---|---|---|
| 10 | Una cifra intercalada en la redacción se actualiza sola | Las regiones `fs-dato-*` dentro de un párrafo funcionan igual que las de la tabla. |
| 11 | Si el Excel pierde una fila que la redacción cita, se avisa | Las anclas huérfanas se detectan y se reportan en vez de fallar en silencio. |
| 12 | Renombrar una etiqueta rompe el vínculo si no hay rango con nombre | Documenta la limitación como comportamiento esperado, no como sorpresa. |
| 13 | Con rango con nombre, renombrar la etiqueta NO rompe el vínculo | La razón de ser de los rangos `fs_*`. Es la pareja de la anterior. |

### 4.5 Escribir sin destruir

| # | Título de la prueba | Invariante que defiende |
|---|---|---|
| 16 | El respaldo .bak se crea y sirve para volver atrás | La red de seguridad existe y de verdad restaura. |
| 17 | Un .docx corrupto se rechaza en vez de destruir la copia buena | La comprobación de ZIP en `_respaldar` impide sobrescribir el único respaldo sano con uno malo. |

### 4.6 Cualquier documento vale

| # | Título de la prueba | Invariante que defiende |
|---|---|---|
| 18 | Un documento ya integrado se acepta, no se rechaza | La clasificación reconoce el estado «listo». |
| 19 | Un nombre con tildes y espacios duros se encuentra igual | Las tres funciones de reparación de nombres hacen su trabajo. |
| 20 | Un documento en blanco se usa de base | El estado «en blanco» monta el andamiaje encima. |
| 21 | Un documento CON redacción recibe un apartado y conserva su texto | El caso más delicado: el salto de página y la redacción intacta. |
| 22 | Preparar dos veces no duplica nada | Idempotencia de `preparar`, no solo de `refrescar`. |
| 23 | Crear la plantilla desde cero deja un documento refrescable | La opción 6 produce algo que la opción 1 sabe mantener. |
| 24 | Un .doc antiguo se rechaza con una explicación útil | Que el rechazo venga con instrucciones y no con una traza. |

### 4.7 El libro no se puede leer

| # | Título de la prueba | Invariante que defiende |
|---|---|---|
| 25 | Un libro retenido por Excel se explica, no se vuelca la traza | La causa más frecuente de que una ejecución no arranque tiene mensaje propio. La prueba retiene el archivo de verdad, en exclusiva, para provocar el caso. |
| 26 | Un libro que no existe se explica antes de abrirlo | El fallo se detecta antes, no dentro de la biblioteca de lectura. |
| 27 | Un libro normal pasa la comprobación sin estorbar | Que la comprobación anterior no produzca falsos positivos. |

Las tres últimas merecen una nota: `_retener_en_exclusiva(ruta)` (`:802`) abre
el archivo en modo exclusivo desde el propio banco para reproducir exactamente
lo que hace Excel. Es una prueba de un caso de error, ejecutada de verdad y no
simulada.

---

## 5. La pasada con el libro real

`prueba_libro_real(ruta_libro, carpeta)` (`:876`) solo se ejecuta si se pasa
`--libro`. Monta un documento nuevo, le siembra un párrafo, y **refresca dos
veces** con el libro de verdad. Después comprueba tres cosas:

1. La tabla tiene el mismo número de filas en los dos refrescos.
2. La redacción es idéntica.
3. El segundo refresco reporta explícitamente *«Sin cambios»*.

Devuelve un resumen legible con la hoja detectada, el número de líneas leídas,
las filas escritas y las regiones del documento. **El libro solo se lee, nunca
se escribe.**

Es la comprobación que hay que correr cuando llega un Excel nuevo o cuando
cambia la estructura del modelo: contesta a la vez si la detección acierta y si
el refresco es estable.

---

## 6. La prueba de humo del entorno

`tools/verificar.ps1` responde a una pregunta distinta: *¿este equipo puede
ejecutar el programa?* No prueba la lógica, prueba la instalación. Se ejecuta
con clic derecho → *Ejecutar con PowerShell*, y pinta cada comprobación en
verde o en rojo.

Comprueba, en este orden y deteniéndose en cuanto algo falla:

| # | Comprobación |
|---|---|
| 1 | Existe `.\python\python.exe`; si no, remite a `tools\bootstrap_python.ps1`. |
| 2 | La versión del intérprete es una 3.1x. |
| 3 | Importan las cuatro dependencias: `openpyxl`, `docxtpl`, `jinja2`, `python-docx`. |
| 4 | Están el Excel de ejemplo y el código en `src\`. |
| 5 | El generador procesa el ejemplo y produce **33 líneas**. |
| 6 | Escribió `salidas\revisar_tipos.csv`. |
| 7 | `fs_documento.py plantilla` construye un documento base con sus regiones. |
| 8 | `fs_documento.py refrescar` escribe **33 filas** en la tabla principal. |

Las dos cifras de 33 son un valor esperado fijo contra el libro de ejemplo: si
alguien cambia la heurística de detección y el conteo se mueve, la prueba de
humo lo delata inmediatamente. El documento temporal y su `.bak` se borran al
terminar.

| | Banco de pruebas | Prueba de humo |
|---|---|---|
| Pregunta | ¿El motor se comporta bien? | ¿Este equipo puede ejecutarlo? |
| Material | Libros y documentos sintéticos | El Excel de ejemplo del repositorio |
| Alcance | 27 casos de comportamiento | 8 comprobaciones de entorno y extremo a extremo |
| Cuándo | Al tocar el código | Al montar un equipo nuevo |

---

## 7. Qué no está cubierto

La lista siguiente sale de contrastar el catálogo con el código del motor. Es
importante tenerla presente antes de confiar en un cambio solo porque las
pruebas pasen.

- **La protección con contraseña.** Ni `proteger`, ni `proteger_salvo_datos`,
  ni `desproteger`, ni el cálculo del hash tienen prueba alguna. Es código que
  escribe en la configuración del documento y que, si se equivoca de sitio,
  hace que Word se queje al abrir.
- **El candado por región.** `cambiar_candado` y `estado_candado` no se
  ejercitan.
- **La interoperación con Excel por COM.** `nombrar_rangos` y `fijar_tipos` no
  se prueban, y no se pueden probar sin Excel instalado. Son, además, las dos
  únicas operaciones que **escriben en el libro**.
- **La interfaz de ventana.** Nada de `fs_menu.py` se prueba: ni el menú, ni
  los diálogos de archivo, ni la ventana de resultado. Los diálogos dependen
  de PowerShell y de WinForms.
- **El complemento de Office.** No hay ninguna prueba, coherente con que nunca
  se ha compilado.
- **La apariencia y la simplificación.** `normalizar_apariencia`,
  `_ocultar_parrafo` y `simplificar_documento` no se comprueban.
- **La bitácora.** El banco la desactiva expresamente, así que ni el formato
  del `.log` ni la inserción cronológica inversa dentro del documento están
  cubiertos.
- **La cascada de resolución de rutas** más allá del caso del nombre con
  tildes: `_reubicar_perfil` y el barrido de carpetas OneDrive no se prueban.
- **El empaquetado.** Que los ejecutables se construyan y arranquen no se
  verifica de forma automatizada.
- **Las regiones ya duplicadas.** El banco comprueba que no se dupliquen, pero
  no hay prueba del comportamiento del motor ante un documento que llegue ya
  con duplicados.

Hay además una discrepancia documental que conviene corregir: el `README.md`
del repositorio habla de *«Catorce comprobaciones»*, y el banco tiene 27. El
texto se quedó en la versión 0.11 y la batería creció después.

---

## 8. Cómo añadir una prueba

La receta es corta:

1. Escriba una función que reciba una carpeta temporal y devuelva una cadena
   con el detalle que quiere ver en el informe.
2. Decórela con `@prueba("Frase que describe el invariante")`. El título debe
   afirmar lo que se garantiza, no lo que se ejecuta: *«Cambiar el orden de las
   filas no descoloca las cifras»*, no *«prueba de reordenación»*.
3. Monte el material con `Banco(carpeta)`, o con `escribir_libro` si necesita
   un libro fuera de lo común.
4. Provoque el cambio y llame a `banco.refrescar(...)`.
5. Cierre con `comprobar_sano(banco.docx, "etiqueta")` y con las afirmaciones
   propias del caso. Compare siempre la redacción contra la radiografía previa.
6. Si el caso es un error esperado, compruebe que el mensaje es explicativo, no
   solo que se lanza la excepción.

La función se registra sola: el decorador la anexa a la lista y el ejecutor la
recoge sin más cambios.

---

## Resumen del capítulo

- El banco trabaja siempre sobre copias temporales; el documento y el libro
  reales no se tocan nunca.
- Verifica después de **cada** escritura, no al final, porque el riesgo real es
  la corrupción lenta y hay que poder señalar el refresco culpable.
- Las cuatro invariantes de `comprobar_sano` son: ZIP válido, sin etiquetas
  duplicadas, `fs-meta` legible y el documento con regiones.
- La redacción sembrada en el documento es el patrón de control: no puede
  cambiar nunca, por muchos refrescos que se hagan.
- Son 27 pruebas agrupadas en siete familias; la del desgaste —veinte refrescos
  seguidos— es la que descarta el riesgo de fondo.
- `--libro` añade una pasada con el Excel de verdad, de solo lectura, y es lo
  que hay que correr cuando cambia el modelo.
- `tools/verificar.ps1` responde a otra pregunta: si el equipo puede ejecutar
  el programa. Fija dos valores esperados de 33 que delatan cualquier deriva de
  la heurística.
- Lo que no está cubierto es sustancial: protección, candados, Excel por COM,
  interfaz, complemento y bitácora.
