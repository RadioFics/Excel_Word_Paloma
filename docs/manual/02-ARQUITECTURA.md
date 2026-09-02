# 02 · Arquitectura y funcionamiento interno

> **Para quién.** Quien vaya a mantener el código, revisarlo o decidir sobre él.
> Supone que ya se conocen los conceptos del capítulo anterior.
> **Qué encontrará.** Cómo están repartidas las responsabilidades, por dónde
> circula el dato, qué garantías ofrece el sistema y a costa de qué, y los
> invariantes de diseño que hay que respetar al tocar cualquier pieza.
> **Antes de leer.** El vocabulario de [Panorama](01-PANORAMA.md): región,
> ancla, contrato, contexto, documento vivo, foto.

## Índice del capítulo

1. [Las cuatro capas](#1-las-cuatro-capas)
2. [El contrato como costura](#2-el-contrato-como-costura)
3. [Anatomía de una región](#3-anatomía-de-una-región)
4. [El ciclo de vida de un documento](#4-el-ciclo-de-vida-de-un-documento)
5. [Cómo circula el dato en un refresco](#5-cómo-circula-el-dato-en-un-refresco)
6. [Los invariantes de diseño](#6-los-invariantes-de-diseño)
7. [Las tres restricciones del entorno](#7-las-tres-restricciones-del-entorno)
8. [Qué garantiza el sistema y qué no](#8-qué-garantiza-el-sistema-y-qué-no)

---

## 1. Las cuatro capas

El sistema se reparte en cuatro capas con responsabilidades que no se solapan.
Cada una conoce a la de abajo y no sabe nada de la de arriba.

```
  ┌─ CAPA 4 · ENTRADA ────────────────────────────────────────────────┐
  │  fs_menu.py          la ventana, el menú, las banderas            │
  │  refrescar_fs.py     la vía corta: preparar + refrescar           │
  │  .bat / .exe         los lanzadores                               │
  │  ─ no contiene lógica de negocio: despacha ─                      │
  └───────────────────────────────┬───────────────────────────────────┘
                                  │
  ┌─ CAPA 3 · ESCRITURA ──────────▼───────────────────────────────────┐
  │  fs_documento.py     abre el .docx, localiza anclas, reescribe,   │
  │                      diffea, respalda, guarda, bloquea, protege   │
  │  ─ es la única que escribe en Word ─                              │
  └───────────────────────────────┬───────────────────────────────────┘
                                  │
  ┌─ CAPA 2 · CONTRATO ───────────▼───────────────────────────────────┐
  │  fs_contrato.py      nombres de ancla, derivación de claves,      │
  │                      traducción contexto → {ancla: texto}         │
  │  ─ no toca ningún archivo. Es vocabulario puro ─                  │
  └───────────────────────────────┬───────────────────────────────────┘
                                  │
  ┌─ CAPA 1 · LECTURA ────────────▼───────────────────────────────────┐
  │  generador_fs.py     abre el .xlsx, detecta, clasifica, formatea  │
  │  ─ es la única que lee Excel. También renderiza la «foto» ─       │
  └───────────────────────────────────────────────────────────────────┘
```

Dos observaciones sobre este reparto.

La primera es que **la capa 2 no toca archivos**. Su docstring lo dice: *«Este
módulo NO toca archivos. Define el contrato.»* Esa pureza es deliberada: es lo
que permite que exista una segunda implementación del contrato en TypeScript,
dentro del complemento de Office, sin arrastrar nada del resto.

La segunda es que **`generador_fs.py` tiene dos papeles**. Es la capa de
lectura, sí, pero también contiene el generador clásico: el que renderiza una
plantilla de Word y produce una foto. Son dos responsabilidades en un mismo
archivo, y es la única mezcla de capas del sistema. Conviene tenerlo presente
al leerlo.

### Dependencias reales entre módulos

```
  fs_menu.py ──────┬──> generador_fs.py ──> (openpyxl)
                   ├──> fs_documento.py ──> (python-docx, lxml)
                   └──> fs_contrato.py

  refrescar_fs.py ─┬──> generador_fs.py
                   └──> fs_documento.py

  fs_documento.py ─┬──> fs_contrato.py
                   └──> generador_fs.py   (import diferido: config y rutas)

  fs_contrato.py ───> (nada del proyecto)
```

El único punto sutil es la dependencia de `fs_documento.py` hacia
`generador_fs.py`: es un **import diferido**, dentro de las funciones que lo
necesitan, y solo para reutilizar la carga de configuración y la resolución de
rutas. No hay ciclo en tiempo de importación.

---

## 2. El contrato como costura

El contrato es la pieza más pequeña del sistema y la más importante. Es la
costura por la que se separan los dos mundos: **de dónde salen las cifras** y
**dónde se escriben**.

Un ancla tiene esta forma:

```
   fs  -  dato  -  total_assets  -  actual
   │      │        │               │
   │      │        │               └─ campo: actual | previo | nota
   │      │        │                          | var_abs | var_pct
   │      │        └─ clave: el identificador estable de la línea
   │      └─ familia: tabla | campo | dato | prosa
   └─ prefijo fijo
```

`descomponer(tag)` hace el camino inverso y devuelve `(familia, nombre,
campo)`. Si el tag no pertenece al contrato, devuelve una terna vacía, y esa es
la razón por la que un control de contenido ajeno —uno que alguien insertara a
mano en Word— es sencillamente **invisible** para el refresco.

### La derivación de la clave

`clave(etiqueta)` convierte el texto de una fila en un identificador estable,
en cinco pasos: quitar tildes, minúsculas, todo lo que no sea alfanumérico a
guion bajo, colapsar los repetidos, cortar a 40 caracteres.

```
   'Cash and cash equivalents'  →  cash_and_cash_equivalents
   'Total assets'               →  total_assets
   'Provisión (neta)'           →  provision_neta
```

El docstring lleva una exigencia que no es un detalle: *«Debe dar el MISMO
resultado en Python y en TypeScript.»* Si las dos implementaciones divergieran,
el complemento de Office escribiría en anclas distintas de las que escribe el
núcleo, y el fallo no sería ruidoso: el documento se quedaría con las cifras
viejas y reportaría anclas huérfanas, sin que nada se rompiera.

### Los dos orígenes de la identidad

Una línea puede identificarse de dos maneras, y no son equivalentes:

| Origen | Cómo se obtiene | Qué la rompe |
|---|---|---|
| `etiqueta` | Derivada del texto de la fila | Renombrar la fila en el Excel |
| `rango` | Un rango con nombre de Excel con prefijo `fs_` | Nada razonable: sobrevive a renombrados e inserciones |

`clave_de_linea(linea)` prefiere siempre el rango y cae en la etiqueta si no
hay. La orden `nombrar` crea esos rangos apuntando a la **celda de la
etiqueta**, no a la de la cifra, para que Excel reajuste la referencia solo
cuando se inserten o borren filas encima.

Es la diferencia entre un sistema que se rompe cuando alguien corrige una falta
de ortografía en el Excel y uno que no.

---

## 3. Anatomía de una región

Un `.docx` es un archivo ZIP con documentos XML dentro. El texto vive en
`word/document.xml`, y la norma que lo describe es ECMA-376. El sistema
manipula ese XML directamente.

Una región es un elemento `w:sdt` —*structured document tag*, control de
contenido— con esta estructura:

```xml
<w:sdt>
  <w:sdtPr>                                   propiedades
    <w:alias w:val="Tabla — estado principal"/>   nombre que ve la persona
    <w:tag   w:val="fs-tabla-principal"/>         LA IDENTIDAD
    <w:id    w:val="1834729103"/>                 derivado del tag
    <w:lock  w:val="sdtContentLocked"/>           el candado
    <w:richText/>                                 tipo: bloque o en línea
  </w:sdtPr>
  <w:sdtContent>                              lo que el motor reescribe
    …
  </w:sdtContent>
</w:sdt>
```

Cuatro decisiones concentradas ahí:

- **`w:tag` es la clave funcional; `w:alias` es puramente humano.** El motor
  indexa y compara por el tag. El alias es lo que Word muestra en la pestaña
  del control, para que la persona sepa qué está mirando.
- **El identificador se deriva del tag**, no es aleatorio. Es lo que hace que
  dos construcciones seguidas produzcan el mismo XML, condición necesaria para
  que la operación sea idempotente.
- **El tipo distingue bloque de línea.** Una región de bloque puede contener
  párrafos y tablas; una en línea vive dentro de un párrafo y solo lleva texto
  plano. La tabla y las zonas de redacción son de bloque; los campos y las
  cifras sueltas, en línea.
- **El candado no estorba al motor.** El motor escribe el XML directamente, así
  que el bloqueo no le afecta; sí obliga al complemento de Office a desbloquear
  antes de escribir, porque este pasa por la API de Word.

¿Por qué controles de contenido y no marcadores o campos? Porque un control de
contenido **delimita** una zona (tiene principio y fin explícitos), sobrevive a
la edición del texto de alrededor, puede llevar una etiqueta arbitraria, y Word
lo respeta al copiar y pegar. Un marcador es un punto, no una zona; un campo
tiene semántica propia de Word que interfiere.

---

## 4. El ciclo de vida de un documento

```
                     ┌──────────────┐
                     │  .docx que   │
                     │  llega       │
                     └──────┬───────┘
                            │  clasificar_documento()
             ┌──────────────┼──────────────┐
             ▼              ▼              ▼
        ┌─────────┐   ┌───────────┐   ┌──────────┐
        │ LISTO   │   │ EN BLANCO │   │CON TEXTO │
        │ ya tiene│   │ ≤3 párr.  │   │ redacción│
        │ regiones│   │ 0 tablas  │   │  propia  │
        └────┬────┘   └─────┬─────┘   └────┬─────┘
             │              │              │
             │              │ construir()  │ construir(apartado=True)
             │              │ encima       │ tras salto de página
             └──────────────┴──────┬───────┘
                                   ▼
                          ┌─────────────────┐
                          │ DOCUMENTO VIVO  │◄────────┐
                          └────────┬────────┘         │
                                   │                  │
                       refrescar() │                  │ cada cierre
                                   └──────────────────┘
```

Las tres entradas convergen en el mismo estado, y ese es el punto: **ninguna
forma de llegar es un rechazo**. El comentario del código lo dice: *«Las tres
formas en que puede llegar un documento que el usuario elige. Ninguna es un
rechazo: las tres se saben trabajar.»*

Solo se rechaza lo que de verdad no se puede abrir: un `.doc` del Word antiguo,
un archivo que no es un ZIP válido, o uno que la biblioteca no consigue leer.
Que le falten las regiones no es motivo de rechazo: se le añaden.

El estado `CON TEXTO` merece atención porque es el más delicado. El andamiaje
entra **detrás de un salto de página**, como un apartado aparte, y el salto se
inserta *retroactivamente*: primero se añade todo, después se identifica cuál
fue el primer elemento nuevo y se le pone el salto delante. Poner el separador
al final del cuerpo no habría servido, porque ahí ya está la tabla.

---

## 5. Cómo circula el dato en un refresco

```
  ┌─ 1 ─────────────────────────────────────────────────────────────┐
  │  comprobar_escribible(documento)                                │
  │  ¿lo tiene Word abierto? ¿hay un ~$ al lado? ¿quién lo retiene? │
  │  → si está en uso, ABORTA sin haber tocado nada                 │
  └────────────────────────────┬────────────────────────────────────┘
                               ▼
  ┌─ 2 ─────────────────────────────────────────────────────────────┐
  │  _respaldar(documento)   →  documento.docx.bak                  │
  │  antes comprueba que el original es un ZIP válido: nunca se     │
  │  machaca una copia buena con una mala                           │
  └────────────────────────────┬────────────────────────────────────┘
                               ▼
  ┌─ 3 ── LECTURA ──────────────────────────────────────────────────┐
  │  leer_contexto(libro.xlsx)  →  ctx                              │
  │  hoja por contenido · columnas por perfil · encabezado por      │
  │  patrones · tipo de fila por heurística · números formateados   │
  └────────────────────────────┬────────────────────────────────────┘
                               ▼
  ┌─ 4 ── CONTRATO ─────────────────────────────────────────────────┐
  │  construir_valores(ctx)  →  {ancla: texto}                      │
  │  campos del encabezado, cifras sueltas y sus variaciones        │
  │  (las tablas NO van aquí: las arma el motor)                    │
  └────────────────────────────┬────────────────────────────────────┘
                               ▼
  ┌─ 5 ── DIFF ─────────────────────────────────────────────────────┐
  │  _leer_meta()  →  la foto del refresco anterior                 │
  │  _calcular_cambios()  →  lista de frases legibles               │
  │  SE CALCULA ANTES DE ESCRIBIR NADA                              │
  └────────────────────────────┬────────────────────────────────────┘
                               ▼
  ┌─ 6 ── ESCRITURA ────────────────────────────────────────────────┐
  │  un solo recorrido del índice de anclas, despachando por        │
  │  familia. Las de prosa se cuentan y NO SE ABREN.                │
  └────────────────────────────┬────────────────────────────────────┘
                               ▼
  ┌─ 7 ── RASTRO ───────────────────────────────────────────────────┐
  │  bitácora al .log y/o dentro del documento                      │
  │  _guardar_meta()  →  la foto nueva, al final de todo            │
  └────────────────────────────┬────────────────────────────────────┘
                               ▼
  ┌─ 8 ── GUARDADO ─────────────────────────────────────────────────┐
  │  guardar_seguro(): serializar a un temporal FUERA de OneDrive,  │
  │  comprobar que es un ZIP legible, y solo entonces volcar sobre  │
  │  el original conservando su identidad de archivo                │
  └─────────────────────────────────────────────────────────────────┘
```

Tres cosas de este flujo no son obvias y conviene subrayarlas.

**El diff se calcula antes de escribir.** Si se calculara después no habría con
qué comparar: la foto anterior ya estaría sobrescrita.

**El emparejamiento del diff es por nombre, no por posición.** Reordenar filas
en el Excel no genera ruido en la bitácora. Y las filas de subtotal, que no
tienen etiqueta, se nombran por la sección en la que caen —«Subtotal de Current
assets»— porque si no todas se llamarían igual y el diff las daría por nuevas
en cada corrida.

**Las zonas de redacción se cuentan pero no se abren.** El contador
`sin_ancla_prosa` es lo que la consola imprime como *«Zonas de prosa
intactas: N»*. Es la garantía central del sistema, hecha visible.

---

## 6. Los invariantes de diseño

Estas doce reglas explican el porqué de casi todo el código. Romper cualquiera
de ellas rompe una garantía del sistema.

**1 · Idempotencia por construcción, no por comprobación.**
El identificador de cada región se deriva de su etiqueta; el andamiaje congela
el índice al empezar y solo añade lo que faltaba antes; la normalización de
apariencia salta lo que ya coincide; insertar una cifra que ya existe no la
duplica. Correr dos veces cualquier operación produce el mismo resultado.

**2 · El guardado no es atómico, y es a propósito.**
Lo habitual sería escribir a un temporal y reemplazar. Aquí se hace al revés,
porque reemplazar el archivo cambia su identificador en el sistema de archivos
y OneDrive lo interpreta como «el documento desapareció y hay uno desconocido
en su sitio». Su forma de resolver ese conflicto es **reponer la versión del
servidor**, y los cambios se pierden minutos después sin un solo mensaje de
error. El detalle completo está en [el capítulo 04](04-MOTOR-DEL-DOCUMENTO.md).

**3 · Todo nombre de archivo que pasa por PowerShell viaja por archivo.**
Nunca por la salida estándar. PowerShell 5.1 escribe en la página de códigos de
la consola, y los caracteres que esa página no tiene se sustituyen **en
silencio**. Eso convertía una ruta válida en una inexistente y hacía que un
documento con el que se llevaba trabajando semanas se rechazara con «No
existe».

**4 · En lxml, identidad por referencia viva, nunca por dirección de memoria.**
Los envoltorios de Python se crean al vuelo y se reciclan; dos nodos distintos
pueden acabar en la misma dirección. El código lo documenta dos veces porque
provocó un fallo real: *«un documento lleno de redaccion se contaba como
vacio»*.

**5 · Dos niveles de bloqueo, jerarquizados explícitamente.**
El candado de región es *«una comodidad de la interfaz»*: Buscar y reemplazar
lo atraviesa y Word en el navegador ni lo mira. La protección de documento es
*«la protección que Word IMPONE de verdad»*. Por eso existen las dos, y por eso
hay dos modos de protección con huecos editables en sitios opuestos.

**6 · Para leer el Excel, `openpyxl`; para escribirlo, Excel.** Nunca al revés.
`openpyxl` no recalcula fórmulas y al reguardar descartaría los valores que
Excel dejó cacheados, dejando el Word en blanco. Es un fallo catastrófico y
silencioso, y por eso el código lo advierte dos veces.

**7 · PowerShell en lugar de dependencias binarias.**
Los diálogos de archivo y el control de Excel se hacen desde PowerShell y no
con bibliotecas nativas de Python, porque el intérprete portable no trae
interfaz gráfica y añadir dependencias binarias complicaría el empaquetado en
un ejecutable de un solo archivo.

**8 · La bitácora vive fuera del documento por defecto.**
Una bitácora creciendo dentro estorba la lectura y provoca conflictos de
sincronización cuando dos personas abren el documento.

**9 · Nada se rechaza si se puede arreglar.**
La falta de regiones se resuelve añadiéndolas; un nombre con tildes comidas se
busca; una ruta de otro perfil se reubica.

**10 · Ante ambigüedad, fallar antes que adivinar.**
Las búsquedas por aproximación exigen **exactamente un** candidato: *«ante la
duda, mejor decir que no existe que refrescar el documento equivocado»*.

**11 · La configuración de máquina se separa de la de proyecto.**
`config.json` viaja por Git y se embebe en el ejecutable; `config.local.json`
no se versiona y manda. Los marcadores `${ONEDRIVE}`, `${USUARIO}` y
`${PROYECTO}` permiten que una ruta compartida siga siendo válida en cualquier
equipo.

**12 · Los errores previstos se explican; los imprevistos muestran la traza.**
Todas las validaciones lanzan un tipo de excepción reservado para lo previsto,
con mensajes largos y con instrucciones de recuperación. Cualquier otra
excepción se trata como defecto del programa y se muestra entera, pidiendo que
se copie para soporte.

---

## 7. Las tres restricciones del entorno

Buena parte del diseño no responde a preferencias técnicas sino a tres hechos
del sitio donde el programa tiene que funcionar.

**OneDrive.** El documento vive en una carpeta sincronizada. Eso condiciona
cómo se guarda (invariante 2), dónde vive la bitácora (invariante 8), y obliga
a detectar los archivos que están en la nube pero no descargados localmente,
porque abrirlos dispara una descarga que puede tardar o fallar.

**No se puede instalar nada.** De ahí el intérprete de Python embebido, que se
descomprime en una carpeta del proyecto sin tocar el registro ni pedir permisos
de administrador; el ejecutable de un solo archivo, que no se instala; y la
apuesta a largo plazo por un complemento de Office, que se despliega de forma
centralizada desde el centro de administración y no requiere tocar los equipos.

**Windows en español.** La consola viene en una página de códigos que no sabe
escribir muchos caracteres. Eso genera el invariante 3 y las tres funciones de
reparación de nombres de archivo, que existen exclusivamente para deshacer un
daño ya hecho.

---

## 8. Qué garantiza el sistema y qué no

Conviene ser preciso, porque de esto depende la confianza que se le puede
poner.

### Garantiza

| Garantía | Cómo |
|---|---|
| La redacción sobrevive a cada refresco | El motor no abre las zonas de prosa. Verificado por el banco de pruebas con veinte refrescos seguidos. |
| Si se borra un párrafo, se queda borrado | El motor nunca reinyecta contenido de prosa. |
| El documento no se corrompe al escribir | Validación de ZIP antes de volcar, comprobación de archivo en uso, y una copia `.bak` previa. |
| Se puede saber qué cambió y de qué Excel salió | La foto en `fs-meta` más la huella SHA-256 del libro en el origen. |
| La tabla se puede mover por la hoja | Detección por contenido, no por coordenadas. |
| El ancho de columna ajustado a mano sobrevive | El refresco conserva el grid de la tabla anterior. |
| Nadie teclea encima de una cifra por accidente | El candado de región, y la protección de documento si se activa. |

### No garantiza

| Límite | Consecuencia práctica |
|---|---|
| La clasificación de filas es heurística | Sin columna `Tipo`, un cambio de formato en Excel puede cambiar cómo se clasifica una fila. El CSV de revisión existe por eso. |
| Renombrar una fila rompe el vínculo | Salvo que esa fila tenga un rango con nombre. |
| Las cifras dentro de párrafos no se detectan solas | Hay que marcarlas una vez, con la orden `insertar`. |
| Una región duplicada pasa desapercibida | El motor solo refresca la primera aparición y no avisa. |
| El candado de región no es una protección real | Buscar y reemplazar lo atraviesa; Word en el navegador lo ignora. |
| Solo hay una generación de respaldo | El `.bak` se sobrescribe en cada operación. |
| Un solo tipo de estado financiero | Otro estado requiere otra plantilla. |
| Solo Windows de 64 bits | El intérprete embebido y los diálogos son específicos de Windows. |

El capítulo [Límites y riesgos](10-LIMITES-Y-RIESGOS.md) desarrolla cada uno de
estos puntos y añade los que salen del código.

---

## Resumen del capítulo

- Cuatro capas con responsabilidades separadas: lectura, contrato, escritura y
  entrada. Solo la de lectura toca Excel, solo la de escritura toca Word, y la
  del contrato no toca ningún archivo.
- El contrato es la costura del sistema y la frontera compartida con el
  complemento de Office; la derivación de claves debe ser idéntica en los dos
  lenguajes.
- Una región es un control de contenido cuya etiqueta es la identidad y cuyo
  identificador se deriva de esa etiqueta, lo que hace idempotente la
  construcción.
- Las tres formas en que puede llegar un documento convergen en el mismo estado:
  ninguna es un rechazo.
- El diff se calcula antes de escribir y empareja por nombre, no por posición.
- Los doce invariantes de diseño explican casi todo el código; el más
  contraintuitivo es que el guardado **no** sea atómico, y la razón es OneDrive.
- Tres restricciones del entorno —OneDrive, la imposibilidad de instalar y la
  página de códigos de Windows en español— explican lo que de otro modo
  parecerían rarezas.
