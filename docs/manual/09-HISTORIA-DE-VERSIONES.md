# 09 · Historia de las versiones

> **Para quién.** Quien quiera entender por qué el código es como es, y quien
> tenga que explicar la evolución del proyecto a alguien que no lo siguió.
> **Qué encontrará.** Las quince entregas del repositorio, en orden, con lo que
> cambió en cada una y qué significa ese cambio en términos de producto. Al
> final, las líneas de fuerza que atraviesan toda la historia.
> **Antes de leer.** Nada imprescindible. Se entiende mejor con los conceptos
> de [Panorama](01-PANORAMA.md) a mano.

## Cómo se reconstruyó esto

El repositorio no lleva notas de versión: los mensajes de commit son, casi
siempre, la etiqueta `V0.x` a secas. Lo que sigue se reconstruyó leyendo los
objetos de Git —los commits, sus árboles y los archivos de cada versión— y
comparando cada árbol con el de su padre para saber exactamente qué archivos se
añadieron, se modificaron, se eliminaron o se movieron. La lectura de qué
significa cada salto sale del contenido de los archivos que aparecen y
desaparecen, y del `README.md` y los documentos de cada momento.

Lo que sea deducción se marca como tal.

- **Repositorio:** `github.com/RadioFics/Excel_Word_Paloma`
- **Rama:** `main`, historia estrictamente lineal, sin fusiones.
- **Etiquetas:** `0.1`, `0.7` y `0.12`.
- **Periodo:** del 31 de agosto al 2 de septiembre de 2026. **Tres días.**

---

## La historia de un vistazo

| # | Versión | Fecha y hora | Añade | Modif. | Elim. | Archivos |
|---|---|---|---|---|---|---|
| 1 | Commit inicial | 31 ago 17:00 | 1 | 0 | 0 | 1 |
| 2 | V0.1 | 31 ago 17:04 | 28 | 1 | 0 | 29 |
| 3 | README de descarga | 31 ago 17:30 | 0 | 1 | 0 | 29 |
| 4 | Distribución `.exe` · etiqueta `0.1` | 31 ago 17:39 | 1 | 3 | 0 | 30 |
| 5 | V0.2 | 1 sep 11:09 | 5 | 10 | 0 | 35 |
| 6 | V0.3 | 1 sep 11:38 | 1 | 10 | 0 | 36 |
| 7 | V0.4 | 1 sep 13:32 | 9 | 8 | 5 | 40 |
| 8 | V0.5 | 1 sep 15:09 | 4 | 9 | 0 | 44 |
| 9 | V0.6 | 1 sep 16:04 | 0 | 5 | 0 | 44 |
| 10 | V0.7 · etiqueta `0.7` | 1 sep 16:28 | 1 | 3 | 0 | 45 |
| 11 | V0.8 | 1 sep 16:44 | 0 | 4 | 0 | 45 |
| 12 | V0.9 | 1 sep 17:39 | 0 | 2 | 0 | 45 |
| 13 | V0.10 | 1 sep 17:56 | 0 | 1 | 0 | 45 |
| 14 | V0.11 | 2 sep 00:00 | 3 | 11 | 0 | 48 |
| 15 | V0.12 · etiqueta `0.12` | 2 sep 09:07 | 0 | 8 | 0 | 48 |

Crecimiento de los tres módulos principales, en líneas:

```
  fs_documento.py    —      1.143  1.316  1.467  1.990  2.449 … 2.918  3.309
  generador_fs.py   676  704  709   795    823    888    895 …  1.167  1.258
  fs_menu.py         —     —     —     —     149    374    597 …  883   1.061
                    V0.1  ·0.1  V0.2  V0.3  V0.4   V0.5   V0.6 … V0.11 V0.12
```

---

## Día 1 · El generador (31 de agosto)

### Commit inicial — 17:00

Un solo archivo, el `.gitattributes` que genera GitHub. El esqueleto vacío al
que se empujó todo tres minutos después.

### V0.1 — 17:04 · el proyecto entra de golpe

Veintiocho archivos a la vez. No es un comienzo: es la publicación de algo que
ya estaba hecho. Un `generador_fs.py` de 676 líneas que traslada el estado de
Excel a una plantilla de Word, y cuya tesis arquitectónica ya está en el
docstring: **la hoja y las columnas se identifican por su contenido, no por
posiciones fijas**, y la columna `Tipo` es opcional.

Ya conviven los **tres frentes** que `DIRECCION.md` documenta: vínculos nativos
de Word, Python portable, y un complemento de Office en TypeScript que está
presente desde el primer minuto, con su manifiesto y su panel.

Y ya está identificado el obstáculo real, que no es técnico sino de
gobernanza: la organización bloquea instaladores, así que `bootstrap_python.ps1`
monta un intérprete embebido sin tocar el registro ni pedir administrador.

El `.gitattributes` pasa de la plantilla genérica a una política explícita:
finales de línea normalizados, **CRLF forzado para los `.bat` y los `.ps1`** —que
en Windows lo necesitan— y tratamiento binario para los documentos de Office.

### Commit 3 — 17:30 · cambia el público, no el código

No se toca una línea de código. El `README.md` pasa de 116 a 240 líneas y se
reorganiza en dos mitades: primero «Descargar y probar», con cinco pasos para
alguien sin conocimientos técnicos, y al final el material de desarrollo.

La sección nueva se titula *«¿Por qué Windows me advierte? ¿Es un virus?»*, y
es el reconocimiento explícito de que **el obstáculo del producto ya no es la
lógica contable sino la confianza y el antivirus**.

### Commit 4 — 17:39 · un solo archivo · etiqueta `0.1`

Salto de empaquetado. Se abandona el paquete comprimido de 21 MB con Python
dentro y se pasa a **un solo ejecutable de unos 13 MB**. La plantilla y la
configuración viajan embebidas, lo que obliga a introducir la fusión de
configuración: la embebida más la que el usuario deje junto al ejecutable.

El README se acorta a la mitad, porque ya no hay que explicar cómo verificar un
archivo comprimido: se reduce a un enlace de descarga. Es el primer entregable
real, y por eso lleva etiqueta.

---

## Día 2 · El documento vivo (1 de septiembre)

Entre las 11:09 y las 17:56 el proyecto pasa de V0.2 a V0.10. Nueve entregas en
menos de siete horas.

### V0.2 — 11:09 · el cambio conceptual

**Es el salto más grande de toda la historia.** Aparece la idea de *documento
vivo*: el Word ya no se regenera desde cero, sino que se le refrescan regiones
marcadas dejando intacta la redacción de alrededor.

Eso exige tres piezas nuevas y coordinadas, que nacen en el mismo commit:

- `CONTRATO.md`, la especificación de anclas;
- `fs_contrato.py`, el vocabulario —*«este módulo NO toca archivos»*—;
- `fs_documento.py`, 1.143 líneas, el motor que *«ACTUALIZA EN EL SITIO»*.

Y una cuarta que revela la intención de fondo: `addin/src/core/contrato.ts`. El
contrato se declara desde el principio **fuente única para los dos mundos**,
Python y complemento. Nace en el mismo commit que su gemelo en Python, no
después.

El `.gitignore` empieza a ignorar los `.docx.bak`: ya hay escritura destructiva
sobre documentos reales.

### V0.3 — 11:38 · la identidad de una fila

Se resuelve el punto frágil del contrato. Con la lectura de rangos con nombre,
el vínculo entre Excel y Word deja de depender del texto de la etiqueta y pasa
a apoyarse en rangos `fs_*`, de modo que renombrar una línea ya no rompe la
cifra que la cita en la redacción.

Nace `GUIA.md`, un manual de operación en lenguaje de usuario cuya frase-tesis
es que el motor solo escribe dentro de las regiones marcadas —y por eso, si se
borra un párrafo, se queda borrado. El `.gitignore` añade los `.xlsx.bak`: ahora
también se escribe sobre el libro.

### V0.4 — 13:32 · el orden

Commit de reorganización estructural. La raíz deja de ser un vertedero y se
reparte en `src/`, `docs/`, `plantillas/`, `ejemplos/`, `tools/` y `salidas/`.
Seis archivos se mueven de sitio y cinco desaparecen de la raíz para reaparecer
bajo `docs/`.

El añadido funcional es `src/refrescar_fs.py`, que formaliza el reparto en dos
órdenes gemelas: *`generador_fs.py` crea un Word nuevo; `refrescar_fs.py`
actualiza el que ya existe*, con `refrescar.bat` como lanzador de doble clic.

La configuración gana dos claves: la ruta del documento base y el prefijo de
los rangos. El script de empaquetado pasa a compilar **dos** ejecutables. Y
`docs/ESTRUCTURA.md` congela la disciplina de carpetas con una instrucción
explícita: *«Mantener esta división en los próximos commits.»*

*(Inferencia: los tres documentos que «desaparecen y reaparecen» bajo `docs/`
no son movimientos puros —cambiaron de contenido a la vez, y se les coló una
marca de orden de bytes al reeditarlos, que persiste hasta hoy.)*

### V0.5 — 15:09 · un solo icono

Cambio de cara del producto. Los dos caminos se confundían —a ambos se les
arrastra el Excel encima, pero uno crea y el otro actualiza—, así que nace
`fs_menu.py`: *«un solo icono que pregunta qué quiere hacer»*. Con él llega un
tercer ejecutable, `EstadosFinancieros.exe`, que pasa a ser **el** enlace de
descarga del README.

El motor gana `--estado`, una inspección en seco que dice qué hoja y qué
columnas leyó sin escribir nada. Y aparecen los dos documentos que atacan las
dudas operativas más grandes: `docs/DATOS.md` («las cifras se editan en Excel,
nunca en Word») y `docs/CAMBIAR_EXCEL.md` (qué revisar cuando llegue el libro
definitivo).

### V0.6 — 16:04 · la ventana, y el README se parte en dos

Dos movimientos simultáneos.

El README se reescribe: un bloque corto para el usuario y un «Detalle técnico»
claramente separado. Todo lo largo se empuja a `docs/`, que casi duplica su
tamaño.

Y `fs_menu.py` se convierte de menú de consola en **interfaz de ventana**. El
usuario ya puede elegir el documento base desde la aplicación en vez de editar
un archivo de configuración a mano. `DATOS.md` documenta lo nuevo y difícil:
cómo intercalar una cifra viva dentro de un párrafo, y los tres estados del
candado.

### V0.7 — 16:28 · honestidad técnica · etiqueta `0.7`

`docs/VINCULOS_NATIVOS.md` documenta el frente A del plan original: cómo
vincular Excel y Word con Pegado especial, **sin programar nada**, e incluye
explícitamente *«lo que se rompe, que lo he comprobado ejecutándolo»*.

Es decir: se escribe la alternativa que haría innecesario el propio proyecto,
con sus límites medidos. Es un commit que dice mucho del criterio con que está
hecho el repositorio.

### V0.8 — 16:44 · por qué hacían falta dos candados

Cierre del asunto de la protección. `DATOS.md` gana la sección *«Por qué hacían
falta DOS candados»*: el bloqueo deja de ser un solo interruptor y pasa a ser
dos mecanismos distintos de Word —protección de documento y bloqueo de
contenido de las regiones—, con el motor ajustado en consecuencia.

Pequeño en volumen, pero es el commit que hace que «las cifras no se puedan
teclear encima» sea cierto de verdad.

### V0.9 — 17:39 · afinado

Cuarenta y tres líneas repartidas entre el lector de Excel y el escritor de
Word, sin cambios en documentación ni interfaz. *(Inferencia: por su tamaño y
por no arrastrar cambios de README, la lectura razonable es corrección de
detalle sobre lo introducido en V0.8. El commit no lleva mensaje que lo
explique.)*

### V0.10 — 17:56 · el commit más pequeño

Un solo archivo tocado. Refactor de la interfaz: se extrae el manejo de
opciones y, sobre todo, la resolución del libro de Excel a funciones propias.
El menú deja de repetir la lógica de «¿qué Excel uso?» en cada rama.

Cierra una jornada en la que se pasó de V0.2 a V0.10.

---

## Día 3 · Robustez (2 de septiembre)

### V0.11 — 00:00 · el commit que vino de otro equipo

El más denso desde V0.4, y el único hecho **en otra máquina** —el autor figura
con otro nombre y llegó a este equipo por sincronización ocho horas después—.
Ese detalle explica exactamente su contenido. Tres frentes:

**1 · Multimáquina.** La configuración compartida deja de llevar la ruta
absoluta al documento, que hasta entonces estaba escrita literalmente en el
repositorio, y pasa a `config.local.json`, no versionado, más un sistema de
marcadores portables `${ONEDRIVE}`, `${USUARIO}` y `${PROYECTO}`. El comentario
del propio archivo lo justifica: *«este archivo viaja por git y una ruta
absoluta se le impondría a la otra maquina en cada pull»*. Es el problema que
se sufrió al trabajar desde dos equipos, arreglado precisamente en el commit
que viene del segundo.

**2 · Pruebas.** Nace `tools/probar_refresco.py` con su lanzador: comprobaciones
sobre copias temporales que, tras **cada** escritura, verifican que el archivo
siga siendo un ZIP válido, que no haya regiones duplicadas y que la redacción
no se haya movido. El README de V0.1 decía «no hay pruebas automatizadas
todavía»; aquí deja de ser cierto.

**3 · Detección robusta.** El lector de Excel crece 265 líneas con la detección
de encabezado, la puntuación de hojas y los reconocedores de fecha y moneda: la
tabla puede estar en cualquier esquina, y el nombre de hoja de la configuración
pasa a ser **una pista que se verifica**, no una orden.

*(Nota técnica: entre V0.5 y V0.10 el archivo del motor estaba guardado con
finales de línea corruptos. V0.11 los normaliza. Un conteo ingenuo haría
parecer que el archivo se redujo a la mitad; en realidad creció.)*

### V0.12 — 09:07 · «actualiza el documento que tengas» · etiqueta `0.12`

Ningún archivo nuevo: es un commit de acabado, escrito una hora después de
recibir V0.11 del otro equipo.

Lo central está en la sección nueva del README, *«Qué documentos valen»*: el
programa deja de exigir un `.docx` preparado y **clasifica el documento que le
den** en tres casos —ya tiene regiones, está en blanco, o tiene redacción
propia—. En el código eso son la clasificación de documentos, el conteo de
contenido visible y el diálogo de guardado; en el lector de Excel, la
comprobación de legibilidad y la detección de archivos de OneDrive no
descargados, otro problema típico de trabajar entre dos máquinas.

El banco de pruebas crece para cubrirlo. Con esto el producto pasa de
«actualiza el documento que preparamos» a **«actualiza el documento que
tengas»**.

---

## Las líneas de fuerza

Cinco cosas atraviesan la historia entera y explican el estado actual.

**1 · El producto se define hacia atrás, desde el usuario.**
Tres de las quince entregas no tocan lógica: reescriben el README para otro
público. El commit 3 no cambia una línea de código y es uno de los más
importantes, porque es donde el proyecto reconoce que su obstáculo es la
confianza y no la contabilidad.

**2 · El giro conceptual está en V0.2 y todo lo demás lo desarrolla.**
De «generar un Word» a «refrescar regiones de un Word que ya existe». Las diez
versiones siguientes son consecuencias de esa decisión: la identidad estable
(V0.3), el menú que distingue los dos caminos (V0.5), los dos candados (V0.8),
la clasificación de documentos (V0.12).

**3 · La fricción real se convierte en código.**
Casi todos los mecanismos raros del sistema tienen un incidente detrás: el
nombre con espacio duro que se corrompía en la consola, la ruta del perfil ajeno
que quedó embebida en un ejecutable, OneDrive revirtiendo los cambios, un
documento lleno de texto que se contaba como vacío. Trabajar desde dos equipos
produjo por sí solo el sistema de configuración local y buena parte de V0.11.

**4 · El complemento lleva congelado desde V0.4.**
Nació en V0.1 con manifiesto y panel, recibió su contrato en V0.2 y su última
corrección de código en V0.3. Desde entonces, once entregas sin tocarlo,
mientras el frente B absorbía todo el desarrollo. No es abandono: es que el
frente B empezó a cumplir el requisito antes de lo previsto. Pero conviene no
confundir «escrito» con «disponible».

**5 · Tres días, quince entregas.**
El ritmo explica tanto la calidad de las decisiones —están tomadas con el
problema delante— como la deuda: mensajes de commit sin cuerpo, un README que
todavía menciona catorce pruebas cuando hay veintisiete, y un archivo de
especificación de empaquetado versionado que ya no es el que se usa.

---

## Recomendaciones para las próximas versiones

Salen de la propia historia, no de una preferencia:

- **Escribir el cuerpo de los mensajes de commit.** Los dos únicos commits con
  mensaje largo (los que reescriben el README) son los únicos que se explican
  solos meses después. `V0.13` no dice nada.
- **Mantener un archivo de novedades.** Este capítulo tuvo que reconstruirse
  leyendo objetos de Git; un `CHANGELOG.md` de cinco líneas por versión habría
  bastado.
- **Etiquetar todas las entregas publicadas, no tres de quince.** El enlace de
  descarga del README apunta a la última publicación, así que una entrega sin
  etiqueta no llega al usuario.
- **Revisar el README cuando cambie lo que describe.** El desajuste del número
  de pruebas es inofensivo; el mismo mecanismo aplicado a una instrucción de
  uso no lo sería.

---

## Resumen del capítulo

- Quince entregas en tres días, en una historia lineal sin fusiones, con
  etiquetas en `0.1`, `0.7` y `0.12`.
- V0.1 publica un proyecto ya maduro, con los tres frentes planteados desde el
  primer minuto.
- **V0.2 es el giro**: de generar un documento a refrescar regiones de uno que
  ya existe. Todo lo posterior desarrolla esa decisión.
- V0.4 impone la estructura de carpetas; V0.5 y V0.6 construyen la interfaz;
  V0.8 resuelve los dos candados.
- V0.11 viene de otra máquina y por eso trae el sistema de configuración local,
  el banco de pruebas y la detección robusta.
- V0.12 completa el producto: ya no hace falta un documento preparado, vale el
  que se tenga.
- El complemento de Office lleva congelado desde V0.4, mientras el núcleo de
  Python absorbía todo el desarrollo.
