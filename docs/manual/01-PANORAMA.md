# 01 · Panorama y conceptos básicos

> **Para quién.** Cualquiera que se acerque al proyecto por primera vez: quien
> lo va a usar, quien lo va a mantener y quien tiene que decidir sobre él. No
> hace falta saber programar para leer este capítulo.
> **Qué encontrará.** Qué problema resuelve el sistema, la idea que lo sostiene,
> el vocabulario que se usa en todo el manual, las piezas que lo componen y las
> tres alternativas que se evaluaron antes de construirlo.
> **Antes de leer.** Nada. Este es el punto de entrada.

## Índice del capítulo

1. [El problema](#1-el-problema)
2. [La idea](#2-la-idea)
3. [El vocabulario](#3-el-vocabulario)
4. [Las piezas](#4-las-piezas)
5. [Los dos caminos que no hay que confundir](#5-los-dos-caminos-que-no-hay-que-confundir)
6. [Un cierre contado de principio a fin](#6-un-cierre-contado-de-principio-a-fin)
7. [Los tres frentes que se evaluaron](#7-los-tres-frentes-que-se-evaluaron)
8. [Qué hay que saber antes de tocar nada](#8-qué-hay-que-saber-antes-de-tocar-nada)

---

## 1. El problema

Cada cierre contable hay que pasar las cifras de un modelo de Excel a un
documento de Word: el Estado de Situación Financiera, con su tabla de partidas,
sus dos columnas comparativas y sus notas. Alrededor de esa tabla hay
redacción: párrafos de análisis, comentarios, explicaciones que escribe una
persona y que son la parte que aporta valor.

Hacerlo a mano es lento y se cuelan erratas. Y la solución obvia —regenerar el
documento entero desde una plantilla— tiene un defecto que la descarta: **borra
la redacción**. Cada cierre habría que volver a escribirla.

El requisito del cliente, en sus palabras, es explícito: *el mismo documento
actualiza sus datos*, no se genera uno nuevo.

De ahí sale la única especificación que de verdad importa:

```
   Escribes un párrafo  →  refrescas  →  el párrafo sigue ahí, cifras al día
   Borras ese párrafo   →  refrescas  →  sigue borrado
```

La segunda línea es tan importante como la primera. Un sistema que reinyectara
el párrafo borrado sería igual de inservible que uno que lo borra: la persona
perdería el control de su propio texto.

Hay un segundo requisito, menos visible pero igual de duro: **la organización
bloquea instaladores** y revisa los ejecutables. Cualquier solución que exija
instalar algo en el equipo de la usuaria muere antes de nacer. Eso condiciona
el diseño entero, y explica decisiones que de otro modo parecerían raras.

---

## 2. La idea

El sistema no regenera el documento. **Solo escribe dentro de unas regiones
marcadas, y no visita nada más.**

Word tiene un mecanismo pensado exactamente para esto: los *controles de
contenido*, unas cajas etiquetadas que delimitan un trozo del documento y le
ponen un nombre. El motor busca esas etiquetas, reescribe su interior y se va.
Un párrafo que esté fuera de una región no es que se respete: es que el motor
ni siquiera pasa por él.

```
   ┌──────────────────────────────────────────────────┐
   │  Collective Mining Ltd.          ← región campo  │
   │  Estado de Situación Financiera  ← región campo  │
   │                                                  │
   │  Este trimestre los activos crecieron por la     │  ← redacción libre:
   │  entrada de la concesión minera, que suma        │     el motor NO entra
   │  2.267.625 y explica la mayor parte de la        │     ↑ salvo esta cifra,
   │  variación.                                      │       que sí es región
   │                                                  │
   │  ┌────────────────────────────────────────────┐  │
   │  │  ASSETS                    2026      2025  │  │  ← región tabla:
   │  │  Cash and cash equiv.  72.957.812   81.370 │  │     se regenera entera
   │  │  Total assets          77.315.954   88.811 │  │
   │  └────────────────────────────────────────────┘  │
   └──────────────────────────────────────────────────┘
```

Sobre esa idea se levantan tres decisiones que definen el carácter del sistema:

- **Nada se localiza por coordenadas.** Ni en el Excel ni en el Word. La tabla
  del estado se puede mover a otra esquina de la hoja, se le pueden insertar
  filas encima, se puede renombrar la hoja: se sigue leyendo igual, porque
  todo se identifica por su contenido y no por su posición.
- **El documento recuerda cómo estaba.** Dentro del propio `.docx`, oculta,
  hay una foto del último refresco. Gracias a ella el sistema puede decir qué
  cambió sin depender de ningún archivo externo.
- **Ante la duda, no se adivina.** Si hay dos documentos que podrían ser el
  buscado, no se elige ninguno. Es preferible un error claro a refrescar en
  silencio el documento equivocado.

---

## 3. El vocabulario

Estos términos se usan con el mismo significado en todo el manual. Conviene
fijarlos ahora.

| Término | Qué es |
|---|---|
| **Región** | Un control de contenido de Word: la caja etiquetada que delimita un trozo del documento. Es lo que el motor abre y reescribe. |
| **Ancla** | El nombre de una región. Siempre empieza por `fs-`. Es la identidad: `fs-tabla-principal`, `fs-dato-total_assets-actual`. |
| **Contrato** | La especificación de qué anclas existen y qué significa cada una. Vive en `src/fs_contrato.py` y en `docs/CONTRATO.md`, y es la frontera compartida entre el núcleo de Python y el complemento de Office. |
| **Contexto** (`ctx`) | Lo que se extrae del Excel: encabezado, líneas, escalares. Es un diccionario, y es lo único que viaja entre la capa que lee y la que escribe. |
| **Línea** | Una fila del estado ya interpretada: su etiqueta, su tipo, su cifra actual, su cifra previa, su nota. |
| **Clave** | El identificador estable de una línea: `total_assets`, `cash_and_cash_equivalents`. Se deriva de la etiqueta o, mejor, de un rango con nombre del Excel. |
| **Documento vivo** (o **documento base**) | El `.docx` que tiene regiones dentro. Es el que se refresca cada cierre. |
| **Foto** | Un `.docx` desechable, generado desde una plantilla, **sin** regiones. Sirve para mirarlo o imprimirlo; no se puede volver a actualizar. |
| **Andamiaje** | El conjunto de regiones que hay que añadir a un documento para que sea refrescable. |
| **Refresco** | La operación de reescribir solo las regiones de datos. Es lo que se hace cada cierre. |
| **Candado** | El bloqueo de una región: Word no deja teclear dentro. |
| **Protección** | El bloqueo de documento de Word, con contraseña. Es más fuerte que el candado. |
| **Bitácora** | El registro de qué cambió en cada refresco. Por defecto va a un archivo `.log` aparte. |
| **Ancla huérfana** | Una región que sigue en el documento pero para la que el Excel ya no da valor. Típicamente, una fila que se borró del libro. |

Las seis familias de anclas son estas:

| Familia | Ejemplo | Qué contiene | ¿La toca el refresco? |
|---|---|---|---|
| `fs-tabla-` | `fs-tabla-principal` | Una tabla completa | Sí, la regenera |
| `fs-campo-` | `fs-campo-empresa` | Un dato del encabezado | Sí |
| `fs-dato-` | `fs-dato-total_assets-actual` | Una cifra suelta dentro de un párrafo | Sí |
| `fs-prosa-` | `fs-prosa-analisis` | Una zona de redacción libre | **No. Nunca la abre.** |
| `fs-registro` | — | La bitácora dentro del documento | Solo antepone entradas |
| `fs-meta` | — | La foto oculta del último refresco | La reescribe al final |

---

## 4. Las piezas

El sistema tiene cinco módulos de Python, un complemento de Office sin
compilar, y una capa de scripts para montar el entorno y empaquetar.

```
    Excel (.xlsx)
        │
        │  openpyxl, solo lectura
        ▼
  ┌───────────────────┐
  │  generador_fs.py  │  Detecta hoja, columnas, región y encabezado.
  │  la LECTURA       │  Clasifica cada fila. Formatea los números.
  └────────┬──────────┘  Produce el CONTEXTO.
           │
           │  ctx = {empresa, titulo, fechas, lineas: [...], escalares: {...}}
           ▼
  ┌───────────────────┐
  │  fs_contrato.py   │  Traduce el contexto a {ancla: texto}.
  │  el CONTRATO      │  Define los nombres. No toca ningún archivo.
  └────────┬──────────┘
           │
           ▼
  ┌───────────────────┐
  │  fs_documento.py  │  Abre el .docx, localiza cada ancla, reescribe
  │  el MOTOR         │  su interior, calcula el diff, guarda con cuidado.
  └────────┬──────────┘
           │
           ▼
    Word (.docx) vivo
```

Por encima de esos tres hay dos puntos de entrada:

- **`fs_menu.py`** es la ventana que ve la usuaria. No duplica lógica: es un
  despachador que llama a los tres de abajo. Se empaqueta como
  `EstadosFinancieros.exe`.
- **`refrescar_fs.py`** es la vía corta: prepara el documento si hace falta y
  lo refresca. Se empaqueta como `RefrescarFS.exe`.

Y al margen, **`addin/`**, un complemento de Word escrito en TypeScript que
replica la misma operación dentro de Office. Está escrito y **nunca se ha
compilado**; el capítulo [06](06-COMPLEMENTO-DE-OFFICE.md) explica su estado
real.

| Carpeta | Contiene |
|---|---|
| `src/` | Los cinco módulos de Python. |
| `docs/` | La documentación, incluido este manual en `docs/manual/`. |
| `plantillas/` | Documentos `.docx` modelo. |
| `ejemplos/` | El libro de muestra. |
| `tools/` | Entorno, empaquetado, verificación y el banco de pruebas. |
| `addin/` | El complemento de Word. |
| `salidas/` | Lo que genera el programa y el `.log` de la bitácora. No se versiona. |
| `dist/` | Los ejecutables publicables. No se versiona. |

---

## 5. Los dos caminos que no hay que confundir

Es la confusión más frecuente del sistema, y la razón por la que existe el
menú. A los dos caminos se les arrastra el mismo Excel encima, pero hacen cosas
distintas y solo uno de ellos sirve para trabajar mes a mes.

| | **Documento vivo** | **Foto** |
|---|---|---|
| Opción del menú | 6 (crear) y 1 (refrescar) | 2 |
| Qué produce | Un `.docx` **con regiones dentro** | Un `.docx` sin regiones, en `salidas\` |
| Se puede volver a actualizar | Sí, cada cierre | **No** |
| Conserva la redacción | Sí | No aplica: nace de una plantilla |
| Para qué sirve | Es el documento de trabajo | Mirar, imprimir, enviar una versión puntual |
| Módulo que lo hace | `fs_documento.crear_base()` | `generador_fs.py` |

Dicho de otro modo: la opción 6 crea **la base que se refresca cada cierre**; la
opción 2 saca **una fotografía** de este momento que ya no se puede volver a
poner al día.

---

## 6. Un cierre contado de principio a fin

Así es como se usa el sistema en la práctica, la primera vez y las siguientes.

**La primera vez**, hay dos caminos según de dónde se parta:

- *No hay documento todavía.* Se elige la opción 6. El programa pregunta dónde
  guardarlo —disco u OneDrive, da igual—, lo crea ya con las cifras dentro y lo
  deja fijado como el documento que se actualizará.
- *Ya hay un Word.* Se elige la opción 3 y se selecciona el archivo. El
  programa lo examina y se adapta solo: si ya tiene regiones lo usa tal cual;
  si está en blanco monta el estado encima; y si trae redacción propia, añade
  el estado **detrás de un salto de página**, como un apartado aparte, sin
  tocar lo que hubiera escrito.

**Cada cierre**, el trabajo es este:

1. Cerrar el Excel y el Word. Si alguno está abierto, el programa se detiene y
   dice cuál es y quién lo retiene. El libro solo se lee, pero Excel lo retiene
   en exclusiva mientras lo tiene abierto y no deja ni leerlo: es la causa más
   frecuente de que una ejecución no arranque.
2. Arrastrar el Excel sobre el ejecutable y elegir la opción 1.
3. El programa deja una copia `.bak` del documento, lee el libro, reescribe las
   regiones, escribe la bitácora y guarda.
4. Al terminar, una ventana resume qué cambió y ofrece abrir el documento.

Y entre cierres, la persona escribe. Esa redacción no se toca.

---

## 7. Los tres frentes que se evaluaron

El proyecto no empezó decidiendo construir esto. `docs/DIRECCION.md` documenta
tres caminos que se estudiaron en paralelo, y conviene conocerlos porque
explican por qué el sistema es como es:

| Frente | Qué era | Aprobación de TI | Estado |
|---|---|---|---|
| **A. Tablas vinculadas** | Pegado especial → Pegar vínculo de rangos de Excel en Word, y «actualizar todo» | Ninguna: es nativo de Office | Documentado en `docs/VINCULOS_NATIVOS.md`, con sus límites medidos |
| **B. Python portable** | Un intérprete embebido y un ejecutable de un solo archivo | Consultar la política de ejecutables | **Es la solución operativa** |
| **C. Complemento de Office** | Un panel dentro de Word con un botón «Actualizar desde Excel» | Una revisión y despliegue centralizado | Escrito, sin compilar |

El frente A funciona sin programar nada, y por eso está documentado en el
repositorio con honestidad, incluyendo lo que se rompe. Su límite es que el
vínculo guarda la ruta del libro —si el archivo se mueve, hay que
revincular— y que el rango no se ajusta cuando se insertan o borran filas.
Sirve para tablas de estructura estable durante el año.

El frente B, que empezó como puente, acabó cumpliendo el requisito del cliente:
refresca en el sitio conservando la redacción. Por eso pasó a ser la solución
operativa mientras el frente C consigue su revisión.

El frente C ya no tiene que demostrar el concepto —el contrato de anclas está
especificado y probado—, solo trasladar la misma operación a un panel dentro de
Word.

---

## 8. Qué hay que saber antes de tocar nada

Cinco advertencias que evitan la mayoría de los problemas:

1. **Cierre el Excel y el Word antes de ejecutar.** Escribir sobre un documento
   que Word tiene abierto lo deja inservible. El programa lo comprueba y se
   detiene a propósito, pero es mejor no llegar ahí.
2. **Las cifras se editan en el Excel, nunca en el Word.** Una cifra escrita a
   mano dentro de una región la machaca el siguiente refresco, porque la región
   sigue vinculada al libro. Si de verdad hay que escribirla a mano, hay que
   *desvincular* esa región.
3. **`config.json` viaja por Git y no debe llevar rutas de ninguna máquina.**
   Lo que depende del equipo va en `config.local.json`, que no se versiona y
   manda sobre todo lo demás.
4. **Abra y guarde el libro siempre desde Excel.** El sistema lee los valores
   que Excel dejó calculados en cada celda; una herramienta que reguarde el
   libro sin recalcular los destruiría, y el Word saldría en blanco.
5. **Solo Windows de 64 bits, y un solo tipo de estado.** Otro estado
   financiero requiere otra plantilla. El capítulo
   [10](10-LIMITES-Y-RIESGOS.md) recoge los límites completos.

---

## Resumen del capítulo

- El problema no es trasladar cifras: es trasladarlas **sin borrar la redacción
  que alguien escribió alrededor**.
- La solución es escribir solo dentro de regiones marcadas y no visitar nada
  más; lo que está fuera no se toca porque el motor no pasa por ahí.
- Nada se localiza por coordenadas: la hoja, las columnas y el encabezado se
  identifican por su contenido.
- Hay dos caminos que se confunden: el **documento vivo**, que se refresca cada
  cierre, y la **foto**, que es una versión puntual imposible de actualizar.
- La restricción de fondo no es técnica sino organizativa: no se puede instalar
  nada, y eso explica el Python embebido, el ejecutable de un solo archivo y la
  apuesta a largo plazo por un complemento de Office.
- El contrato de anclas es la frontera compartida entre el núcleo de Python y
  el complemento: cambiarlo afecta a los dos mundos.
