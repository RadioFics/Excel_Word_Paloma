# 04 · El motor del documento

> **Para quién.** Quien vaya a mantener, auditar o ampliar
> `src/fs_documento.py`, y quien tenga que explicarle a un usuario por qué el
> refresco no le tocó un párrafo, o por qué se negó a escribir.
> **Qué encontrará.** El recorrido completo del módulo: qué es una región por
> dentro, cómo se escribe sobre un documento de OneDrive sin destruirlo, cómo
> se monta el andamiaje, qué hace exactamente un refresco, dónde guarda el
> documento su propia memoria, los dos niveles de bloqueo y la línea de
> órdenes entera.
> **Antes de leer.** Conviene tener presente el contrato de anclas
> (`src/fs_contrato.py`) y la separación entre documento vivo y foto descrita
> en [Arquitectura](02-ARQUITECTURA.md). Este capítulo no repite el contrato:
> lo usa.

## Índice del capítulo

1. [La tesis del módulo](#1-la-tesis-del-módulo)
2. [Las primitivas de OOXML](#2-las-primitivas-de-ooxml)
3. [Escribir sin destruir](#3-escribir-sin-destruir)
4. [PowerShell y el problema de la codificación](#4-powershell-y-el-problema-de-la-codificación)
5. [Construir el andamiaje](#5-construir-el-andamiaje)
6. [Clasificar el documento que llega](#6-clasificar-el-documento-que-llega)
7. [El refresco](#7-el-refresco)
8. [La memoria del documento](#8-la-memoria-del-documento)
9. [La bitácora](#9-la-bitácora)
10. [Cifras dentro de la redacción](#10-cifras-dentro-de-la-redacción)
11. [Apariencia](#11-apariencia)
12. [Los dos candados](#12-los-dos-candados)
13. [Interoperar con Excel por COM](#13-interoperar-con-excel-por-com)
14. [Diagnóstico](#14-diagnóstico)
15. [La línea de órdenes de `fs_documento.py`](#15-la-línea-de-órdenes-de-fs_documentopy)
16. [Puntos frágiles de este módulo](#16-puntos-frágiles-de-este-módulo)

---

## 1. La tesis del módulo

`src/fs_documento.py` son 3.309 líneas con una sola idea dentro. El docstring
la enuncia en el primer párrafo (`src/fs_documento.py:1`):

> "A diferencia de generador_fs.py (que RENDERIZA una plantilla y produce un
> .docx nuevo y desechable), este módulo ACTUALIZA EN EL SITIO un documento
> que ya existe: localiza las anclas del contrato (fs_contrato.py), reescribe
> únicamente su interior y no visita nada más. La prosa que una persona
> redacte alrededor sobrevive intacta a cada refresco; si la borra, se queda
> borrada, porque el motor nunca la vuelve a inyectar."

Las dos mitades de esa frase importan por igual. La primera es la promesa
habitual de cualquier sistema de plantillas: lo que el programa mantiene, lo
mantiene al día. La segunda es la promesa inversa, y es la difícil: **lo que
el programa no mantiene, no lo toca jamás, ni siquiera para reponerlo.** Si
alguien borra el párrafo de introducción, el refresco no lo devuelve. Si
alguien mueve la línea del nombre de la empresa al pie de página, ahí se
queda. El motor no tiene una idea previa de cómo debería verse el documento;
solo tiene una lista de anclas y la instrucción de escribir dentro de ellas.

La frontera con el otro motor del proyecto es explícita. `generador_fs.py`
produce una **foto**: renderiza una plantilla de Word y escupe un `.docx`
nuevo cada vez, sin regiones dentro. Esa foto es perfecta para imprimir y
archivar, y es imposible de refrescar: no hay nada en ella que diga dónde
acaba una cifra y empieza una frase. `fs_documento.py` produce y mantiene el
**documento vivo**: el que lleva las regiones dentro y por eso se puede
volver a actualizar la semana que viene sin perder la redacción de nadie. El
docstring de `crear_base` lo resume (`src/fs_documento.py:1001`):

> "Es la contraparte de generador_fs: aquel renderiza una plantilla de
> Word y produce un documento sin regiones —una foto, imposible de
> refrescar después—; este produce el documento vivo, el que la opción de
> actualizar sabe mantener al día."

De esa tesis se desprende la segunda propiedad del módulo, la idempotencia:

> "'construir' y 'reparar' son la misma operación: añaden lo que falte para
> que el documento cumpla el contrato, sin duplicar ni borrar lo que ya haya.
> Se pueden correr sobre un documento con meses de redacción encima. Si el
> documento ya traía redacción, lo que se añade entra detrás de un salto de
> página, como un apartado aparte."

Y una advertencia operativa que el propio docstring pone en mayúsculas:

> "ADVERTENCIA: el refresco escribe sobre el archivo indicado. Si vive en
> OneDrive, ciérrelo en Word antes de refrescar (o Word y el motor pelearán
> por el archivo). Siempre se deja una copia .bak junto al original."

Las secciones 3 y 5 de este capítulo son, en el fondo, el desarrollo técnico
de esas tres citas.

---

## 2. Las primitivas de OOXML

### 2.1 Qué es un `.docx` por dentro

Un archivo `.docx` no es un documento: es un **ZIP**. Si se le cambia la
extensión a `.zip` y se abre, dentro hay una carpeta `word/` con varios
archivos XML. El principal es `word/document.xml`, que contiene el cuerpo del
documento en un dialecto llamado **WordprocessingML**, parte de la norma
ECMA-376 (también conocida como OOXML). Junto a él viaja
`word/settings.xml`, donde se guardan los ajustes del documento, entre ellos
la protección.

El vocabulario de WordprocessingML que hace falta para leer este capítulo es
corto:

| Etiqueta | Qué es |
|---|---|
| `w:body` | El cuerpo del documento. Sus hijos directos son los bloques de nivel superior. |
| `w:p` | Un párrafo. |
| `w:r` | Un *run*: un tramo de texto con un formato homogéneo dentro de un párrafo. |
| `w:t` | El texto literal, siempre dentro de un `w:r`. |
| `w:rPr` / `w:pPr` | Las propiedades de un run / de un párrafo (negrita, sangría, alineación…). |
| `w:tbl` / `w:tr` / `w:tc` | Tabla, fila de tabla, celda de tabla. |
| `w:sectPr` | Las propiedades de sección (márgenes, tamaño de papel). Va al final del cuerpo. |
| `w:sdt` | Un *control de contenido*: una caja etiquetada que Word sabe delimitar. |

El módulo usa la librería `python-docx` casi exclusivamente como contenedor:
abre el ZIP, expone `doc.element.body` y `doc.settings.element`, y vuelve a
guardar. Toda la manipulación real se hace en lxml, elemento a elemento, con
`OxmlElement` y `qn()`. No se usa ni una vez la interfaz de alto nivel del
tipo `add_paragraph()`. La razón es directa: la interfaz de alto nivel sabe
crear documentos, pero no sabe operar quirúrgicamente dentro de un documento
ajeno sin alterar el resto.

### 2.2 El control de contenido como mecanismo de anclaje

Un **control de contenido** (`w:sdt`, de *structured document tag*) es una
caja con nombre que Word dibuja alrededor de un trozo de documento. Tiene dos
partes: `w:sdtPr`, sus propiedades, y `w:sdtContent`, lo que hay dentro.
Entre las propiedades está `w:tag`, una cadena arbitraria que Word almacena y
muestra, pero que no interpreta. Ese `w:tag` es el **ancla**: la cadena
`fs-…` que identifica la región.

Word ofrece al menos tres mecanismos que podrían servir para lo mismo, y los
tres se descartaron de hecho:

| Mecanismo | Por qué no |
|---|---|
| **Marcadores** (`w:bookmarkStart`) | Son puntos, no cajas: no delimitan contenido de forma fiable. Word los mueve, los duplica al copiar y pegar y los pierde con facilidad; no admiten bloqueo ni etiqueta visible. |
| **Campos** (`w:fldSimple`, DOCPROPERTY, vínculos DDE) | Los actualiza Word, no el motor. Dependen de que el usuario acepte actualizar campos al abrir, y arrastran el problema de los vínculos rotos entre archivos. |
| **Buscar y reemplazar sobre marcas de texto** | No sobrevive a que alguien edite alrededor, y no distingue la cifra de una mención de la cifra en la prosa. |

El control de contenido, en cambio, delimita, sobrevive a la edición
circundante, admite un candado propio (`w:lock`), se puede ocultar
visualmente y —lo decisivo— es exactamente el mismo objeto que manipula la
API JavaScript de Office, lo que permite que el complemento de Word sea un
consumidor de primera clase del mismo contrato. El comentario de `_sdt` lo
menciona de pasada (`src/fs_documento.py:228`):

> "bloqueado -> w:lock="sdtContentLocked": Word impide editar el interior a
> mano. No estorba a este motor (escribimos el XML directamente), pero sí
> obliga al add-in a desbloquear antes de escribir."

### 2.3 Anatomía de una región

```
 w:sdt                                   ← la REGIÓN
 |
 +-- w:sdtPr                             ← propiedades
 |   |
 |   +-- w:alias  w:val="Campo — empresa"      rótulo humano (Word lo enseña)
 |   +-- w:tag    w:val="fs-campo-empresa"     EL ANCLA: la identidad real
 |   +-- w:id     w:val="1244501893"           md5 del tag, determinista
 |   +-- w:lock   w:val="sdtContentLocked"     candado por región (opcional)
 |   +-- w:text  |  w:richText                 en línea  |  de bloque
 |   +-- w15:appearance w15:val="boundingBox"  recuadro gris | "hidden"
 |
 +-- w:sdtContent                        ← LO ÚNICO que el refresco reescribe
     +-- w:r
         +-- w:t xml:space="preserve"  ->  Collective Mining Ltd.
```

El orden del diagrama es el orden real en el archivo: `_sdt` emite alias,
tag, id, lock y tipo de control, en ese orden, y `_poner_apariencia` añade
después el `w15:appearance` al final del `w:sdtPr`
(`src/fs_documento.py:1425`). El valor del `w:id` del ejemplo es ilustrativo
(inferencia: depende del md5 del tag concreto).

### 2.4 El catálogo de primitivas

**`_sdt(tag, alias=None, bloqueado=True, en_linea=False)`**
(`src/fs_documento.py:228`) es el corazón del sistema. Genera exactamente
esto:

```xml
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="…"/>                  <!-- solo si se pasa alias -->
    <w:tag   w:val="fs-…"/>               <!-- SIEMPRE: es la identidad -->
    <w:id    w:val="<_id_estable(tag)>"/>
    <w:lock  w:val="sdtContentLocked"/>   <!-- solo si bloqueado=True -->
    <w:text/>   |   <w:richText/>         <!-- en_linea ? text : richText -->
  </w:sdtPr>
  <w:sdtContent/>                         <!-- vacío: lo llena el llamante -->
</w:sdt>
```

Dos reglas de diseño se leen en ese XML. La primera: `w:tag` es la clave
funcional y `w:alias` es puramente humano. El motor indexa y compara por
`w:tag`; el `w:alias` solo sirve para que Word muestre un rótulo legible en
la pestaña del control. Los alias que el módulo usa son `Redacción — {nombre}`,
`Campo — {nombre}`, `Tabla — estado principal`, `Bitácora de actualizaciones`,
`Metadatos (oculto)` y `{clave} ({campo})`. La segunda: `en_linea=True`
produce un `w:text` (control de texto plano, que vive dentro de un párrafo y
solo puede contener texto) y `en_linea=False` produce un `w:richText` (control
de bloque, que puede contener párrafos y tablas enteras). Los campos de
encabezado y las cifras sueltas son `w:text`; las tablas, las zonas de
redacción, la bitácora y los metadatos son `w:richText`.

**`_id_estable(tag)`** (`src/fs_documento.py:129`) es el md5 del tag,
truncado a 8 dígitos hexadecimales, convertido a entero, reducido módulo
2.000.000.000 y desplazado 1000. Su docstring dice por qué existe:

> "Un w:id determinista por tag: dos construcciones seguidas producen el
> mismo XML (condición necesaria para que el refresco sea idempotente)."

Word exige un `w:id` numérico en cada control y, si se dejara al azar, cada
`construir` produciría un `.docx` distinto byte a byte aunque no hubiera
cambiado nada semánticamente. En un archivo dentro de OneDrive eso significa
una versión nueva en el servidor y un conflicto potencial por cada
ejecución inútil. Con el md5, dos construcciones sobre el mismo documento
producen el mismo XML, y el sistema de sincronización no ve nada que
propagar.

**`_tag_de(sdt)`** (`src/fs_documento.py:261`) lee
`w:sdtPr/w:tag/@w:val` y devuelve `None` si falta cualquiera de los dos
elementos. Es la única puerta de entrada al ancla, y su tolerancia a la
ausencia es lo que permite que el módulo conviva con controles de contenido
ajenos —los que el usuario haya puesto por su cuenta, los que traiga una
plantilla corporativa— sin tropezar.

**`_contenido(sdt)`** (`src/fs_documento.py:257`) es `sdt.find(qn("w:sdtContent"))`.
Es la frontera de la escritura: **todo lo que el refresco modifica está
dentro de un `w:sdtContent`, y nada de lo que hay fuera se toca.**

**`_indexar(doc)`** (`src/fs_documento.py:269`) recorre
`doc.element.body.iter(qn("w:sdt"))` y devuelve `{tag: elemento}`. Al usar
`.iter()`, el recorrido es recursivo: encuentra también las regiones
anidadas, como un `fs-dato-*` colocado dentro de un `fs-prosa-*`, que es
justo el caso de una cifra intercalada en la redacción. Su docstring cabe en
cuatro palabras:

> "{tag: elemento sdt} de todo el documento. Ante duplicados, el primero."

La consecuencia real de esa frase merece decirse sin rodeos: **si un ancla
aparece dos veces en el documento, solo se refresca la primera aparición.**
La segunda queda congelada con el valor que tuviera, sigue pareciendo una
cifra viva —conserva su recuadro y su candado— y no aparece en ningún informe
ni aviso. Un duplicado se produce con facilidad: basta con que alguien copie
un párrafo que contenga una cifra y lo pegue más abajo. Word copia el control
con su `w:tag` intacto. Ni `refrescar` ni `verificar` detectan la situación,
porque ambos trabajan sobre el índice ya deduplicado.

**`_cuerpo_append(doc, el)`** (`src/fs_documento.py:279`) añade al final del
cuerpo *antes* del `w:sectPr` final. Es un detalle de formato de archivo con
consecuencia práctica: las propiedades de sección tienen que ser el último
hijo del `w:body`, y colgar cualquier cosa detrás corrompe la sección final
del documento.

**`_vaciar(el)`** (`src/fs_documento.py:289`) borra todos los hijos de un
elemento iterando sobre `list(el)`. La copia de la lista es obligatoria: en
lxml, mutar la colección mientras se recorre salta elementos.

**`_run`, `_parrafo`, `_parrafo_texto`, `_borde`, `_celda`** son los ladrillos
de construcción. Sus firmas:

| Primitiva | Firma | Notas |
|---|---|---|
| `_run` | `_run(texto_, negrita=None, cursiva=False, oculto=False)` | El `w:rPr` solo se emite si alguna propiedad se usa. `negrita=None` no emite `w:b`; `negrita=False` emite `<w:b w:val="0"/>` (apagado explícito). `oculto` emite `w:vanish`. |
| `_parrafo` | `_parrafo(estilo=None, sangria=None, alineacion=None, espacio_antes=None)` | Devuelve un `w:p` sin runs. Orden de emisión: `w:pStyle`, `w:spacing`, `w:ind`, `w:jc`. Sangría y espacio en *twips* (1/1440 de pulgada). |
| `_parrafo_texto` | `_parrafo_texto(texto_, **kw)` | Extrae `negrita`, `cursiva` y `oculto` y pasa el resto a `_parrafo`. |
| `_borde` | `_borde(nombre, tipo="single", sz="4")` | Un `w:top`, `w:bottom`… con `w:val`, `w:sz`, `w:space="0"`, `w:color="auto"`. |
| `_celda` | `_celda(ancho, parrafos, borde_sup=None, borde_inf=None)` | `w:tc` con ancho en twips y tipo `dxa`. |

Un detalle transversal: **todo `w:t` que crea el módulo lleva
`xml:space="preserve"`** (la constante `XMLSPACE`, `src/fs_documento.py:73`).
Sin ese atributo, Word colapsa los espacios de guarda, y una frase montada
por piezas —«Al », fecha, « (», estado, «)»— pierde los espacios y queda
pegada.

Los estilos de fila viven en `ESTILO_ETIQUETA` (`src/fs_documento.py:79`) y
traducen el tipo de línea a formato: `H` (encabezado de sección) en negrita
sin sangría, `I` (ítem) normal con sangría 220, `S` (subtotal) igual que el
ítem, `T` (total) en negrita sin sangría, `N` (nota) normal y en cursiva. Los
bordes los pone `_fila_de_linea` (`src/fs_documento.py:682`) según la
convención contable: una `S` lleva línea superior simple; una `T` lleva línea
superior simple e inferior doble. Y una regla que conviene conocer antes de
depurar un documento raro: **en las filas de tipo `H` y `N` se vacían nota,
actual y previo** aunque el Excel traiga números en ellas.

---

## 3. Escribir sin destruir

Esta es la sección que justifica la mitad del código del módulo. El escenario
de trabajo es el peor posible para escribir archivos: un `.docx` que vive en
una carpeta de OneDrive, que Word puede tener abierto en este mismo momento,
en el mismo equipo o en otro, y sobre el que un proceso externo va a volcar
bytes.

### 3.1 `comprobar_escribible(ruta)`

Firma: `comprobar_escribible(ruta)` (`src/fs_documento.py:512`). No devuelve
nada; o pasa en silencio o lanza `ValueError`. Su docstring:

> "Aborta ANTES de tocar nada si el documento está en uso.
> Word mantiene el .docx abierto mientras lo tiene en pantalla. Si se
> escribe encima en ese momento, el archivo queda inservible (bytes en
> cero). Antes esto solo se avisaba en la documentación; ahora se impide."

El orden de las comprobaciones es el siguiente:

1. Si el archivo **no existe**, retorna sin más: crear un documento nuevo es
   una operación legítima.
2. Calcula `abierto_en_office` como la existencia del archivo de bloqueo de
   Office, `~$nombre.docx`, junto al original.
3. Intenta `open(ruta, "r+b")`. **Si abre y además no hay `~$`, retorna: todo
   está bien.** Si abre pero hay `~$`, no retorna: sigue adelante y acabará
   fallando. Es decir, el archivo de bloqueo de Office es autoridad
   suficiente por sí solo, aunque el sistema operativo permita escribir.
4. Un `PermissionError` cae al bloque de error. Cualquier otro `OSError` se
   traduce a `ValueError("No puedo escribir en el documento: …")`.
5. Llama a `quien_bloquea(ruta)`. Si hay culpables y **ninguno** de ellos
   contiene la cadena `"en este equipo"`, añade al mensaje la explicación de
   que probablemente lo tenga abierto otra persona a través de OneDrive.
6. Lanza `ValueError` con el nombre del archivo, la carpeta, el detalle de
   quién lo retiene y el cierre: *"No se ha modificado nada. Escribir sobre
   un documento que Word tiene abierto lo deja inservible, así que la
   operación se detiene aquí a propósito."*

**Dónde se invoca.** Al principio de `construir`, `crear_base` (solo si el
destino existe), `quitar_registro_del_documento`, `insertar_dato`,
`refrescar`, `simplificar_documento`, `cambiar_candado`,
`desvincular_region`, `proteger`, `proteger_salvo_datos`, `desproteger` y —de
forma inline, dentro de `main`— la orden `apariencia`.

**Dónde no se invoca, y con razón.** En `verificar`, `estado_candado`,
`clasificar_documento`, `revisar_candidato` y `_estado_documento`. Las cinco
son de solo lectura, y bloquear un diagnóstico porque el usuario tiene el
documento abierto sería contraproducente: el diagnóstico es justo lo que
necesita para saber que lo tiene abierto.

### 3.2 Quién tiene el archivo

Tres funciones responden a esa pregunta desde ángulos distintos.

**`_duenos_office(ruta)`** (`src/fs_documento.py:359`) lee el archivo de
propietario que Office deja al lado del documento:

> "Al abrir un documento, Office crea al lado un '~$nombre.docx' cuyo primer
> byte es la longitud del nombre de usuario y el resto ese nombre. Es lo
> que Word lee para decir 'bloqueado por Fulano'. Si el documento lo tiene
> abierto OTRA PERSONA (por OneDrive/SharePoint), este suele ser el unico
> rastro visible desde aqui."

Prueba dos candidatos —`~$` más el nombre completo y, si el nombre pasa de
dos caracteres, `~$` más el nombre recortado por delante, porque *"Con
nombres largos Office recorta los dos primeros caracteres"*— y luego intenta
dos formatos: el clásico (primer byte = longitud, decodificación en `cp1252`
y después `latin-1`) y, como respaldo, la lectura desde el desplazamiento
`0x36` en `utf-16-le`. Valida siempre con `str.isprintable()`. Devuelve una
lista, posiblemente vacía.

**`_procesos_que_bloquean(ruta)`** (`src/fs_documento.py:406`) es la pieza más
exótica del módulo: una llamada al **Restart Manager** de Windows por
`ctypes`.

> "Es la misma API que usa Windows para el cartel «este archivo esta siendo
> utilizado por...». No requiere permisos de administrador. Si algo falla
> (no es Windows, la DLL no esta) devuelve lista vacia en vez de romper:
> saber quien bloquea es un extra, no la comprobacion en si."

El procedimiento es el canónico de esa API: `RmStartSession` para abrir la
sesión, `RmRegisterResources` para declarar el archivo que interesa, y **dos
llamadas a `RmGetList`** —la primera con el búfer a `None`, aceptando
`ERROR_MORE_DATA` (234), solo para averiguar cuántas entradas hacen falta; la
segunda ya con el búfer del tamaño correcto—. Todo dentro de un `try/finally`
que llama siempre a `RmEndSession`. Devuelve una lista de diccionarios
`{"pid", "app", "servicio"}`. Las estructuras `FILETIME`,
`RM_UNIQUE_PROCESS` y `RM_PROCESS_INFO` se declaran a mano como
`ctypes.Structure`. Fuera de Windows (`os.name != "nt"`) sale de inmediato
con lista vacía.

**`quien_bloquea(ruta)`** (`src/fs_documento.py:487`) envuelve las dos
anteriores en dos `try/except Exception` independientes —su docstring promete
*"Nunca lanza excepcion"*— y compone las líneas explicativas:

- Por proceso local: `"{app}   (PID {pid}[, servicio {svc}])   en este equipo"`.
  La coletilla literal **`"en este equipo"`** no es decorativa:
  `comprobar_escribible` la busca por subcadena para distinguir un bloqueo
  local de uno remoto.
- Por dueño de Office: `"Figura como abierto por: {d}"`.
- Si no hay nada de lo anterior pero el `~$` existe: `"Hay un archivo de
  bloqueo de Office (~$) junto al documento."`

### 3.3 `guardar_seguro(doc, ruta)` — la decisión de diseño del módulo

Firma: `guardar_seguro(doc, ruta)` (`src/fs_documento.py:560`). Todo el
módulo guarda por aquí; no hay un solo `doc.save()` sobre la ruta final.

Lo que hace esta función es **abandonar deliberadamente la escritura
atómica**, que es la buena práctica universal para escribir archivos. El
razonamiento del autor está entero en el docstring y merece leerse completo:

> "Guarda SIN cambiar la identidad del archivo en el disco.
>
> Antes esto era doc.save() a un temporal de la misma carpeta y luego
> os.replace(). Es atómico y en un disco normal está bien, pero en una
> carpeta de OneDrive rompe la sincronización, y de forma silenciosa:
>
>   - os.replace() borra el archivo original y pone otro en su sitio. El
>     archivo que queda tiene un File ID de NTFS NUEVO.
>   - OneDrive lleva su base de datos indexada por ese File ID, no por la
>     ruta. Al no reconocerlo, no lo lee como «el documento cambió», sino
>     como «el documento desapareció y hay uno desconocido en su sitio».
>   - Su forma de resolver ese conflicto es reponer la versión que tiene
>     en el servidor. Minutos después el archivo local vuelve a ser el de
>     antes y los cambios se han perdido sin un solo mensaje de error.
>
> Con Archivos a Petición es todavía más claro: el original es un punto
> de reanalisis (placeholder) y el temporal no, así que ni siquiera son
> el mismo tipo de archivo.
>
> Ahora se hace al revés, que es como escribe Word: se serializa entero a
> un temporal FUERA de la carpeta sincronizada (para que OneDrive no vea
> aparecer y desaparecer archivos sueltos), se comprueba que el resultado
> es un .docx legible, y solo entonces se vuelca sobre el archivo original
> abriéndolo en modo r+b. El archivo conserva su identidad y OneDrive lo
> ve como lo que es: una modificación normal."

El algoritmo final, paso a paso:

1. `tmp = Path(tempfile.gettempdir()) / f"fs_{os.getpid()}_{ruta.name}"`. El
   temporal va al directorio temporal del sistema, **fuera** de la carpeta
   sincronizada, y lleva el PID en el nombre para que dos procesos
   simultáneos no colisionen.
2. `doc.save(str(tmp))`: la serialización completa ocurre ahí.
3. **Validación**: `if not zipfile.is_zipfile(tmp)` lanza
   `ValueError("El documento generado no es un .docx legible; no se escribe
   nada sobre el original.")`. Es una comprobación barata —`zipfile` solo
   mira la firma y el directorio central— y es la que garantiza que nunca se
   vuelque basura sobre un original sano.
4. `datos = tmp.read_bytes()`: el documento entero pasa a memoria.
5. Un `finally` borra el temporal con `unlink()`, tragándose el `OSError`.
6. Volcado: si la ruta existe, se abre en `r+b` y se hace `write(datos)`,
   **`truncate()`**, `flush()` y **`os.fsync(f.fileno())`**. Si no existe, un
   simple `write_bytes`.

Los tres últimos detalles no son adorno. `truncate()` es imprescindible
porque `r+b` no vacía el archivo: si el documento nuevo es más corto que el
viejo —cosa habitual al retirar la bitácora, por ejemplo— sin truncar
quedaría una cola de bytes del archivo anterior pegada al final del ZIP, y el
`.docx` dejaría de abrirse. `flush()` y `fsync()` fuerzan que los bytes lleguen
al disco antes de que la función retorne, en vez de quedarse en la caché del
sistema operativo a merced de un corte de corriente o de que OneDrive lea el
archivo a medias.

**La ventana de riesgo que se asume.** Entre el `write` y el `truncate` el
archivo de destino no es válido: tiene los bytes nuevos y, si el anterior era
más largo, una cola sobrante. Si el proceso muere exactamente ahí —o si el
disco se llena— el documento queda corrupto. Esa es la contrapartida exacta
de no usar `os.replace()`: se cambia una garantía de atomicidad por una
garantía de identidad de archivo. El precio se cubre con dos redes: la
copia `.bak` que hace `_respaldar` antes de cada operación de escritura desde
la línea de órdenes, y la comprobación previa de `comprobar_escribible`, que
elimina la causa más frecuente de escritura fallida.

### 3.4 `_respaldar(ruta)` y el `.bak`

Firma: `_respaldar(ruta)` (`src/fs_documento.py:2308`). Devuelve la ruta de
la copia. Su docstring enuncia una regla que parece obvia y casi nunca se
implementa:

> "Copia previa, pero NUNCA sobre una copia buena con una mala.
>
> Un .docx es un ZIP. Si el archivo de partida no lo es (quedó a medio
> escribir, o Word lo tenía abierto), respaldarlo destruiría la única
> copia sana que queda. Mejor abortar y decirlo."

Por eso lo primero que hace es `zipfile.is_zipfile(ruta)`. Si el original ya
está dañado, no copia: lanza `ValueError` explicando que *"Suele significar
que se escribió sobre él mientras Word lo tenía abierto"* y, **si el `.bak`
que ya existe sí es un ZIP válido**, añade la orden de recuperación literal:

```
copy /Y "documento.docx.bak" "documento.docx"
```

Si el original está sano, hace `shutil.copy2` (que conserva la fecha de
modificación) sobre `ruta.with_suffix(ruta.suffix + ".bak")`. Nótese la doble
extensión: el respaldo de `informe.docx` es `informe.docx.bak`, no
`informe.bak`. Y nótese lo más importante: **solo hay una generación de
respaldo**. Cada operación machaca la anterior. Dos refrescos seguidos dejan
el `.bak` del estado inmediatamente anterior al segundo, no del inicial.

---

## 4. PowerShell y el problema de la codificación

El módulo lanza PowerShell para cuatro cosas: crear nombres de rango en
Excel, escribir la columna «Tipo», abrir el diálogo de «Abrir» de Windows y
abrir el de «Guardar como». Todo eso se apoya en una sola función, y esa
función existe tal como es por un fallo concreto.

### 4.1 `_PS_UTF8` y el nombre con espacio duro

`_PS_UTF8` (`src/fs_documento.py:315`) es una única línea que se antepone a
todos los scripts:

```powershell
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch {}
```

El comentario que la acompaña documenta el caso real que la motivó:

> "Windows PowerShell 5.1 no escribe su salida en Unicode: usa la pagina de
> codigos de la consola (cp437 / cp850 en un Windows en español). Cuando esa
> salida se redirige a una tuberia, los caracteres que la pagina NO tiene no
> dan error: se sustituyen en silencio por el parecido mas cercano.
>
> Eso convertia una ruta perfectamente valida en una que no existe:
>
>     ...EDICIÓN<espacio duro>- PAMELA.docx    (el nombre real)
>     ...EDICION<byte 0xFF>- PAMELA.docx       (lo que llegaba)
>              ^ perdio la tilde   ^ y el espacio duro se rompio
>
> y por eso un documento con el que se llevaba trabajando semanas se
> rechazaba con «No existe» en cuanto se elegia desde el explorador.
>
> Los argumentos que van HACIA PowerShell viajan por CreateProcessW, que ya
> es Unicode: ese sentido nunca ha tenido perdida."

El archivo `config.local.json` del repositorio contiene exactamente ese
nombre: `"${ONEDRIVE}\\DOCUMENTO DE PRUEBA DE EDICIÓN - PAMELA.docx"`. El
fallo no era teórico.

Merece la pena separar las tres piezas del problema, porque cada una tiene su
propia defensa en el código:

- **El BOM del script.** El `.ps1` temporal se escribe con
  `encoding="utf-8-sig"`, es decir, con marca de orden de bytes al principio.
  La razón, citada: *"El .ps1 se escribe con BOM porque PowerShell 5.1 lee un
  archivo sin BOM como ANSI, y entonces los acentos del propio script llegan
  como basura."* Sin BOM, un literal como `'Elija el documento de Word que se
  actualizara'` puede llegar mutilado al intérprete.
- **La lectura como bytes.** `subprocess.run` se llama **sin `text=True`**, y
  la salida se decodifica a mano con `.decode("utf-8", "replace")`. La cita:
  *"La salida se lee como BYTES y se descodifica en UTF-8: ver _PS_UTF8. Con
  `text=True` Python la descodificaba con la pagina de codigos local
  (cp1252), que no es la que PowerShell usa para escribir."* Es decir: había
  dos codificaciones equivocadas encadenadas, y `_PS_UTF8` arregla la de
  salida mientras que la lectura manual arregla la de entrada.
- **La devolución por archivo.** Para las rutas, ni siquiera eso se considera
  suficiente: los diálogos escriben la ruta elegida en un archivo temporal en
  UTF-8 sin BOM y Python lo lee de ahí. `_dialogo_ruta`
  (`src/fs_documento.py:2421`) lo dice: *"El rodeo por archivo es deliberado:
  es la unica via que conserva el nombre byte a byte (ver _PS_UTF8)."*

### 4.2 `ejecutar_ps`

Firma: `ejecutar_ps(script_texto, *argumentos, timeout=600, sta=True)`
(`src/fs_documento.py:321`). Devuelve la tupla `(stdout, stderr, codigo)`.

1. Crea un directorio temporal con `tempfile.mkdtemp(prefix="fs_ps_")` y
   escribe dentro `orden.ps1` con `_PS_UTF8 + script_texto`, en `utf-8-sig`.
2. Construye la orden:
   `["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass"]`, más
   `["-STA"]` si `sta`, más `["-File", str(script)]`, más los argumentos
   convertidos con `str(a if a is not None else "")`.
3. Ejecuta con `subprocess.run(orden, capture_output=True, timeout=timeout)`.
4. Un `finally` borra el directorio temporal.
5. Decodifica las dos salidas en UTF-8 con `errors="replace"`.

`-NoProfile` evita que el perfil de PowerShell del usuario contamine la
ejecución. `-ExecutionPolicy Bypass` evita que la directiva de ejecución
del equipo impida correr el `.ps1`. **`-STA`** pone el hilo en *Single
Threaded Apartment*, que es el modelo de hilos que WinForms exige: sin él, un
`OpenFileDialog` no se abre. Los scripts que pilotan Excel por COM pasan
`sta=False`, porque no crean ventanas y COM prefiere el modelo por defecto.

---

## 5. Construir el andamiaje

Llamamos **andamiaje** al conjunto de regiones que el documento necesita para
poder refrescarse: los campos del encabezado, la tabla principal, las zonas de
redacción, la región de metadatos y —si se pide— la de bitácora. Montarlo es
trabajo de `construir()`.

### 5.1 `construir(ruta, ctx=None, verbose=True, cfg_bitacora=None, apartado=None)`

`src/fs_documento.py:813`. Su docstring fija las dos propiedades que la hacen
segura:

> «Idempotente: correrla dos veces no duplica nada. No borra prosa ni reordena
> lo que ya exista; solo agrega las anclas ausentes al final. `apartado`: si el
> documento ya trae redacción propia, lo que se añade va detrás de un salto de
> página y un título […] None = decidirlo mirando el documento; True/False =
> forzarlo.»

El recorrido real, paso a paso:

1. `comprobar_escribible(ruta)`, abrir con `Document(str(ruta))`, `idx =
   _indexar(doc)`, lista vacía `añadidos`.
2. `parrafos_previos, tablas_previas = _contenido_visible(doc)` — cuánto texto
   ajeno al contrato hay ya en el documento (ver §6.1).
3. **Decisión de apartado.** Si el llamante no la forzó:
   `apartado = parrafos_previos > _UMBRAL_EN_BLANCO or tablas_previas > 0`.
4. `ya_estaban = list(doc.element.body)` — una foto de los hijos de primer
   nivel. El comentario explica por qué es una **lista** y no un conjunto de
   `id()`:

   > «La lista se GUARDA, no se convierte en un conjunto de id(): mientras haya
   > una referencia viva, lxml devuelve siempre el mismo objeto para el mismo
   > nodo y la comparacion por identidad vale. Sin la referencia, los
   > envoltorios se reciclan y las direcciones se repiten.»

5. Se define el predicado `falta(tag) = tag not in idx`. **`idx` no se
   recalcula durante la construcción**, así que «falta» significa siempre «no
   estaba antes de empezar». Es lo que hace que la operación sea idempotente
   sin necesidad de comprobaciones adicionales.
6. **El auxiliar `linea_campos(piezas, **kw)`** (`src/fs_documento.py:850`)
   monta una frase mezclando texto fijo con regiones de campo. Su guarda de
   idempotencia es una sola línea: si **todos** los campos de esa frase ya
   existen en el documento, la frase entera no se vuelve a escribir. Si falta
   alguno, se crea el párrafo y, pieza a pieza, se inserta una región nueva
   (para el campo que falta) o simplemente el valor en texto plano (para el que
   ya vive en otro sitio del documento). El comentario del autor:

   > «Cada campo va dentro de una frase con sentido, no en una línea suelta: el
   > documento tiene que poder imprimirse tal cual desde el primer día. Un
   > campo que ya exista en el documento no se vuelve a poner, así que la
   > persona puede moverlos donde quiera y "reparar" los respeta.»

   Las cuatro frases que monta son exactamente estas:

   | # | Piezas | Formato |
   |---|---|---|
   | 1 | `[("campo","empresa")]` | centrado, `espacio_antes=240` |
   | 2 | `[("campo","titulo")]` | centrado |
   | 3 | `["Al ", ("campo","fecha_actual"), " (", ("campo","estado_actual"), ") — comparado con ", ("campo","fecha_previa"), " (", ("campo","estado_previo"), ")"]` | centrado |
   | 4 | `["Cifras expresadas en ", ("campo","moneda"), ", en unidades de ", ("campo","miles")]` | centrado |

7. **Zona de redacción `introduccion`**, si falta, con el texto guía: *«Zona de
   redacción libre. Escriba aquí lo que quiera: este texto NO se toca al
   refrescar las cifras.»*
8. **Tabla principal**, si falta `fs-tabla-principal`: una región de bloque con
   alias `Tabla — estado principal` que contiene la tabla construida a partir
   de `_lineas_de_tabla("principal", ctx)`. Justo después se añade un párrafo
   vacío, porque Word exige un párrafo tras una tabla que cierre el cuerpo.
9. **Zona de redacción `analisis`**, si falta: *«Zona de análisis. Aquí puede
   intercalar cifras vivas: use "fs_documento.py catalogo" para ver las claves
   disponibles.»*
10. **Bitácora dentro del documento**, solo si la configuración lo pide
    (`bitacora` con valor `documento` o `ambos`). Se crea la región
    `fs-registro` con un párrafo de título en negrita, `Bitácora de
    actualizaciones`. El comentario es explícito:

    > «La bitácora solo se crea DENTRO del documento si se pide expresamente.
    > Por defecto vive en un .log aparte para no estorbar la redacción.»

11. **Metadatos**, siempre: región `fs-meta` con alias `Metadatos (oculto)`,
    con un único párrafo cuyo texto inicial es `{}` y va marcado como oculto.
12. **Inserción retroactiva del separador.** Si se decidió apartado y de hecho
    se añadió algo, se calculan los hijos del cuerpo que no estaban en
    `ya_estaban` —comparando con `is`, identidad de objeto, no con `id()`— y se
    inserta un salto de página **delante del primero de ellos**. Si el campo
    `titulo` no forma parte de lo añadido (es decir, ya existía en el
    documento, quizá movido por alguien), se añade además un rótulo. El
    razonamiento:

    > «La separación del apartado va DELANTE de lo primero que se haya añadido:
    > no al final del cuerpo, que ahí ya está la tabla. Solo el salto de
    > página: el propio andamiaje empieza por las líneas centradas de empresa y
    > título, que ya hacen de encabezado. Poner otro encima lo diría dos veces.
    > Si esas líneas ya existían en el documento (alguien las movió), entonces
    > sí hace falta un rótulo.»

13. `normalizar_apariencia(doc, datos=…)` con el valor de `apariencia_datos`
    (ver §11).
14. Se vuelve a indexar y se oculta de verdad cada párrafo de `fs-meta` con
    `_ocultar_parrafo`.
15. `guardar_seguro(doc, ruta)` — **incondicional**: el archivo se reescribe
    aunque no se haya añadido nada.
16. Informe por consola: `Anclas añadidas (N):` con una línea `+ <tag>` por
    cada una, o bien `Nada que añadir: el documento ya cumple el contrato.`

Devuelve la lista de tags añadidos.

### 5.2 `preparar(...)` — construir con red de seguridad

`preparar(ruta, ctx=None, cfg=None, verbose=True, respaldar=True)`
(`src/fs_documento.py:969`) es, en palabras del autor, *«`construir` con la
copia de seguridad delante y un informe legible detrás»*. Devuelve
`(estado_previo, añadidos)`.

1. Clasifica el documento (§6.2).
2. **Atajo:** si ya está integrado, imprime `Ya está integrado: no hay que
   añadirle nada.` y devuelve sin tocar el archivo ni crear `.bak`.
3. Si `respaldar`, hace `_respaldar(ruta)`. El parámetro existe por un motivo
   preciso: *«`respaldar=False` para cuando el llamante ya hizo la copia (el
   refresco la hace él): un segundo .bak machacaría el bueno con el ya
   modificado.»* Así lo invoca `src/refrescar_fs.py:97`.
4. Traduce la clasificación en la decisión de apartado de forma explícita —
   `apartado=(estado_ == CON_TEXTO)`— en vez de delegar en `apartado=None`.

### 5.3 `crear_base(...)` — el documento vivo desde cero

`crear_base(destino, ctx=None, cfg=None, verbose=True)`
(`src/fs_documento.py:1001`). Es la opción 6 del menú, y el docstring marca la
frontera con el generador clásico:

> «Es la contraparte de generador_fs: aquel renderiza una plantilla de Word y
> produce un documento sin regiones —una foto, imposible de refrescar
> después—; este produce el documento vivo, el que la opción de actualizar sabe
> mantener al día.»

Crea las carpetas que falten, guarda un `Document()` en blanco —*«Un Word en
blanco de verdad: […] trae los estilos normales de Office y ni un párrafo con
texto»*— y le aplica `construir(..., apartado=False)`, porque no hay nada
delante de lo que separarse.

---

## 6. Clasificar el documento que llega

El sistema acepta cualquier `.docx` que le den. Para eso tiene que responder
antes a dos preguntas: *¿este archivo es el que me piden?* y *¿en qué estado
llega?*

### 6.1 Cuánto texto ajeno hay: `_contenido_visible(doc)`

`src/fs_documento.py:2562`. Devuelve la pareja `(párrafos con texto, tablas)`
contando **solo lo que queda fuera de toda región del contrato**. Un párrafo
vacío no cuenta. La pertenencia a una región se comprueba subiendo por los
padres del nodo, y el comentario documenta el fallo real que obligó a hacerlo
así:

> «La pertenencia se mira subiendo por los padres, no comparando id(). En lxml
> el objeto de Python es un envoltorio que se crea al vuelo y se descarta en
> cuanto nadie lo referencia: dos nodos distintos pueden acabar en la misma
> direccion de memoria, y un conjunto de id() da falsos positivos. Con eso, un
> documento lleno de redaccion se contaba como vacio.»

Es el mismo problema que `construir` resuelve por la otra vía, guardando
`ya_estaban` como lista viva.

### 6.2 Los tres estados: `clasificar_documento(ruta)`

`src/fs_documento.py:2586`. Devuelve `(estado, familias, detalle)`. Las reglas,
en este orden:

| Orden | Condición | Estado | Qué implica |
|---|---|---|---|
| 1 | Existe al menos una región `fs-tabla-*` | `LISTO` | Se refresca tal cual. |
| 2 | 3 párrafos o menos con texto y ninguna tabla | `EN_BLANCO` | Se usa de base: se monta el estado encima. |
| 3 | Cualquier otro caso | `CON_TEXTO` | El estado entra detrás de un salto de página. |

El criterio de «ya integrado» es la sola existencia de una tabla del contrato,
nada más. El umbral de «prácticamente en blanco» está justificado en el
código: *«Un Word recien creado trae uno o dos parrafos vacios; uno que alguien
empezo suele traer un titulo y poco mas.»*

### 6.3 `revisar_candidato(ruta)` — qué se rechaza y qué no

`src/fs_documento.py:2608`. Devuelve `(ok, avisos, familias, info)`. Su
filosofía está escrita:

> «Solo se rechaza lo que de verdad no se puede abrir. Que le falten las
> regiones NO es motivo de rechazo: se le añaden. Antes esto era mas estricto
> de lo que hacia falta y dejaba fuera documentos utilizables.»

| Situación | Resultado |
|---|---|
| El archivo no existe | Se buscan variantes del nombre (§6.4). Si no aparece nada, rechazo. |
| Es un `.doc` antiguo | Rechazo, con la instrucción de reguardarlo como `.docx`. |
| Extensión inusual | Solo aviso, no rechazo. |
| No es un ZIP válido | Rechazo: *«No es un .docx válido […] puede que se quedara a medio escribir.»* |
| Word no lo puede abrir | Rechazo, con el tipo de excepción. |
| Está abierto ahora mismo | **Solo aviso**, no rechazo. |

### 6.4 Encontrar el documento aunque el nombre no coincida

Un nombre de archivo escrito en Office o pasado por una consola llega con
frecuencia alterado. Por eso `resolver_documento(argumento, cfg)`
(`src/fs_documento.py:2154`) aplica una cascada de cinco niveles:

```
  1. argumento explícito de la línea de órdenes
  2. cfg["documento_base"]  (si está vacío → error con instrucciones)
  3. expandir ${ONEDRIVE} / ${USUARIO} / ${PROYECTO} / ~   →  ¿existe?
  4. _buscar_parecido()  o  _buscar_sin_tildes()           →  ¿existe?
  5. _reubicar_perfil()  (la ruta viene de OTRO equipo)    →  ¿existe?
                          ↓ nada encaja
                    error explicado
```

- **`_normalizar_nombre`** (`:2250`) unifica la forma Unicode a NFC, colapsa
  cualquier tipo de espacio —incluido el espacio duro U+00A0— y pasa a
  minúsculas sin acentos de comparación.
- **`_buscar_parecido`** (`:2258`) lista los archivos de la carpeta con la
  misma extensión y se queda con el que coincide tras normalizar. Exige
  **exactamente un** candidato.
- **`_sin_tildes`** (`:2276`) va un paso más allá: primero sustituye por
  espacios los caracteres de `_COMODINES` —`ÿ`, `?`, el carácter de reemplazo
  de Unicode— y solo después descompone y quita diacríticos. El orden importa,
  y está justificado: *«Los comodines se cambian ANTES de normalizar: la "ÿ" se
  descompone en "y" + diéresis, así que después de un NFKD ya no habría forma
  de distinguirla de una "y" de verdad.»*
- **`_reubicar_perfil`** (`:2212`) ataca un problema distinto: la ruta apunta al
  perfil de otra máquina. Prueba a sustituir el nombre de usuario por el de
  este equipo, y si eso falla barre las carpetas `OneDrive*` del usuario
  actual, solo en su primer nivel. El motivo está escrito:

  > «config.json viaja entre maquinas y, sobre todo, se queda EMBEBIDO en el
  > .exe con la ruta de quien lo compilo: en cuanto el .exe cambia de manos,
  > ese "C:\\Users\\Fulano\\..." deja de existir y todas las opciones que
  > dependen del documento mueren a la vez.»

La regla común a las tres búsquedas es *«ante la duda, mejor decir que no
existe que refrescar el documento equivocado»*: si hay más de un candidato, no
se elige ninguno.

### 6.5 `fijar_documento_base(ruta_doc, verbose=True)`

`src/fs_documento.py:2679`. Escribe la elección del usuario en
`config.local.json`, nunca en `config.json`, y guarda la ruta **compactada**
con marcadores:

> «Va a config.local.json, no a config.json: el segundo viaja por git, y una
> ruta absoluta escrita ahí se le impone a la otra máquina en cada "pull".
> Además la ruta se guarda compactada (${ONEDRIVE}\\…) para que, si alguien la
> copia al config.json compartido, siga valiendo en ambas.»

Si el archivo existe pero tiene un error de sintaxis, **no lo sobrescribe a
ciegas**: lanza un error explicando que no se atreve a reescribirlo.

---

## 7. El refresco

Es la operación central del sistema y la que se ejecuta en cada cierre.

### 7.1 El ciclo completo

```
   refrescar(ruta, ctx, origen, con_registro, cfg)
        │
        ├─ 1. comprobar_escribible(ruta)      ¿lo tiene Word abierto? → abortar
        ├─ 2. Document(ruta) + _indexar()     {tag → elemento}
        ├─ 3. C.construir_valores(ctx)        {tag → texto} de campos y cifras
        ├─ 4. _leer_meta(idx)                 la foto del refresco anterior
        ├─ 5. _calcular_cambios(meta, ctx)    el diff, ANTES de escribir nada
        │
        ├─ 6. bucle único sobre idx.items(), despachando por familia:
        │        fs-tabla-*   → regenerar la tabla, conservando anchos y tblPr
        │        fs-campo-*   → escribir el valor, o anotar huérfana
        │        fs-dato-*    → escribir el valor, o anotar huérfana
        │        fs-prosa-*   → contar y NO ABRIR
        │        fs-registro  → ignorar aquí
        │        fs-meta      → ignorar aquí
        │        cualquier otro control → invisible para el refresco
        │
        ├─ 7. bitácora: al .log, dentro del documento, a los dos, o a ninguno
        ├─ 8. _guardar_meta(idx, ctx, origen)  la foto nueva, al final
        └─ 9. guardar_seguro(doc, ruta)
```

### 7.2 Qué se conserva de la tabla anterior

Al regenerar una tabla, el motor **rescata dos cosas del XML viejo antes de
vaciarlo**: el `w:tblGrid` (los anchos de columna) y el `w:tblPr` (las
propiedades de tabla). Todo lo demás se construye de nuevo. Esa decisión, que
cabe en dos líneas de código, es la que hace que el ajuste manual de columnas
del usuario sobreviva al refresco: `_anchos_de` (`src/fs_documento.py:622`) lo
dice sin rodeos —*«Si el usuario arrastra las columnas en Word, el tblGrid se
actualiza y el refresco respeta la nueva medida.»*

### 7.3 Anclas huérfanas

Una región **huérfana** es una región de campo o de dato presente en el
documento para la que el mapa de valores no trae entrada: el documento pide una
cifra que el Excel ya no produce. Ocurre cuando se borra una fila del libro, o
cuando se renombra una etiqueta sin tener un rango con nombre que sostenga la
identidad.

El motor las acumula en el informe **pero no las borra ni las vacía**:
conservan su último valor. La decisión es deliberada: borrar automáticamente
una cifra que alguien citó en un párrafo dejaría un hueco en la redacción sin
avisar. Se reportan por consola y en la bitácora, y la decisión queda en manos
de la persona.

### 7.4 El informe que devuelve

| Clave | Contenido |
|---|---|
| `tablas` | Lista de `(nombre, filas escritas)`. |
| `campos` | Cuántas regiones de campo se escribieron. |
| `datos` | Cuántas cifras sueltas se escribieron. |
| `huerfanos` | Lista de tags sin valor en el Excel. |
| `colisiones` | Etiquetas distintas que producen la misma clave. |
| `cambios` | Lista de frases legibles con el diff. |
| `sin_ancla_prosa` | Cuántas zonas de redacción se dejaron intactas. |
| `con_registro` | Si se escribió la bitácora dentro del documento. |
| `bitacora_archivo` | Ruta del `.log`, o `None`. |

### 7.5 Qué NO hace el refresco

Conviene tenerlo presente, porque delimita las responsabilidades del módulo:

- No crea anclas que falten. Eso es `construir` / `preparar`.
- No hace la copia de seguridad. La hace el llamante.
- No borra las anclas huérfanas.
- No vuelve a normalizar la apariencia.
- No toca la protección del documento.

Dos refrescos seguidos con el mismo Excel producen el mismo XML salvo por el
campo `fecha` de `fs-meta` y, si la bitácora está activa, una entrada nueva
que dirá *«Sin cambios en las cifras respecto de la última actualización.»*

---

## 8. La memoria del documento

Para poder decir qué cambió, el documento tiene que recordar cómo estaba la
vez anterior. Esa memoria vive dentro del propio `.docx`, en la región
`fs-meta`, y no en un archivo aparte. La razón práctica es que así viaja con el
documento: si alguien lo copia, lo renombra o lo mueve de carpeta, la memoria
va con él.

### 8.1 El formato de `fs-meta`

`_guardar_meta(idx, ctx, origen)` (`src/fs_documento.py:1066`) escribe un JSON
compacto, sin espacios y con los acentos reales:

```json
{"fecha":"2026-09-02T14:03:21",
 "origen":"libro.xlsx (sha a1b2c3d4e5f6)",
 "lineas":[{"e":"Cash and cash equivalents","t":"I","a":"12,345","p":"11,001"}]}
```

- `fecha`: hora local en formato ISO, al segundo, sin zona horaria.
- `origen`: lo que pase el llamante. La línea de órdenes usa el nombre del
  libro más los doce primeros caracteres de su huella SHA-256, de modo que
  siempre se puede saber **de qué versión exacta del Excel** salió un refresco.
- `lineas`: una entrada por línea, con **claves de una sola letra** para ahorrar
  espacio: `e` etiqueta, `t` tipo, `a` actual, `p` previo. No se guardan la
  nota, la fila, la clave ni los valores numéricos crudos.

`_leer_meta(idx)` (`:1055`) degrada en silencio: si la región no existe o el
JSON no es legible, devuelve un diccionario vacío, que el resto del sistema
interpreta como «primera actualización».

### 8.2 Cómo se compara: `_claves_diff` y `_calcular_cambios`

El problema de fondo es que dos listas de líneas no se pueden comparar por
posición: entre un cierre y otro se insertan filas, se borran y se reordenan.
La solución es dar a cada fila un **nombre estable y legible** y comparar por
nombre.

`_claves_diff(filas, etiqueta_de, tipo_de)` (`:1090`) hace eso. Está
parametrizado con dos funciones de acceso precisamente para poder aplicarse
tanto a la foto antigua (claves `e` y `t`) como al contexto nuevo (claves
`etiqueta` y `tipo`). Su docstring explica el caso difícil:

> «Las filas de subtotal (S) no traen etiqueta: todas se llamarían igual y el
> diff las daría por nuevas en cada corrida. Se las nombra por la sección en la
> que caen ("Subtotal de Current assets:"), y si aun así dos coinciden se
> numeran.»

`_calcular_cambios(meta_previa, ctx)` (`:1117`) construye entonces el diff:

1. Si no hay líneas previas: *«Primera actualización: no hay versión anterior
   con la que comparar.»*
2. Se indexan las filas antiguas por nombre. Al recorrer las nuevas se usa
   `pop`, no una simple consulta: **lo que quede sin consumir al terminar son
   exactamente las filas desaparecidas**.
3. Se emiten tres tipos de frase: `Nueva fila: …`, `<nombre>: <antes> →
   <ahora>  (comparativo …)`, y `Fila retirada: …`.
4. Si no cambió nada: *«Sin cambios en las cifras respecto de la última
   actualización.»*

Un detalle que conviene conocer: **la comparación es textual sobre los valores
ya formateados**, no numérica. Cambiar el formato de una celda en Excel —de
cero decimales a dos, por ejemplo— se reporta como un cambio de cifra aunque el
número sea el mismo.

---

## 9. La bitácora

Cada refresco deja constancia de lo que cambió. Por defecto, esa constancia va
a un archivo `.log` **fuera** del documento. `ruta_bitacora(cfg, documento)`
(`src/fs_documento.py:1143`) lo justifica:

> «Por defecto FUERA del documento: un .log junto al ejecutable, en salidas\\.
> El documento de OneDrive tiene que quedar limpio para redactar; una bitácora
> creciendo dentro estorba la lectura y provoca conflictos de sincronización
> cuando dos personas lo abren.»

Si `bitacora_archivo` no está configurado, la ruta se deriva del nombre del
documento: `salidas\bitacora_<clave del nombre>.log`.

### 9.1 El archivo `.log`

`escribir_bitacora_archivo(...)` (`:1161`) **anexa, nunca reescribe**. Una
entrada tiene esta forma exacta:

```
========================================================================
2026-09-02 14:03:21   MI_DOCUMENTO.docx
  origen: libro.xlsx (sha a1b2c3d4e5f6)
  escrito: 42 filas de tabla, 8 campos, 3 cifras en el texto
  anclas sin dato en el Excel: fs-dato-foo-actual, fs-campo-bar
  cambios:
    - Total assets: 1,000 → 1,200  (comparativo 900 → 1,000)

```

Las líneas `escrito:` y `anclas sin dato…` solo aparecen cuando hay
información que dar.

### 9.2 La bitácora dentro del documento

`_escribir_registro(idx, cambios, origen)` (`:1211`) escribe en la región
`fs-registro` cuando la configuración lo pide. Tres particularidades:

- El sello llega al minuto, no al segundo, a diferencia del `.log`.
- Se escriben **hasta 40 cambios**; si sobran, una línea final los resume.
- La inserción es **cronológica inversa**: el bloque nuevo se coloca justo
  después del párrafo de título, de modo que lo más reciente queda arriba.

`quitar_registro_del_documento(ruta)` (`:1191`) elimina la región entera con su
contenido. Se usa al pasar la bitácora de dentro a fuera. A diferencia del
índice —que ante duplicados se queda con el primero— este recorrido borra
**todas** las apariciones.

Los tres modos de `bitacora` son `archivo` (por defecto), `documento` y
`ambos`; el valor `no` la desactiva. Las dos ramas del código son
independientes, de modo que `ambos` escribe en los dos sitios.

---

## 10. Cifras dentro de la redacción

La tabla del estado no es el único sitio donde aparecen números. Un párrafo de
análisis puede decir *«los activos totales ascendieron a 119.066.301»*, y ese
número debería seguir al Excel igual que los de la tabla. Para eso existen las
regiones de la familia `fs-dato-`.

### 10.1 `insertar_dato(...)`

`insertar_dato(ruta, clave_, campo, zona="analisis", antes="", despues="",
valor_inicial="—", verbose=True)` (`src/fs_documento.py:1236`). El docstring la
sitúa:

> «Es el equivalente por línea de órdenes al botón "insertar dato" del add-in:
> deja el control de contenido en línea, bloqueado y con la etiqueta correcta,
> listo para que el siguiente refresco lo rellene.»

1. Valida que `campo` sea uno de los cinco válidos (`actual`, `previo`, `nota`,
   `var_abs`, `var_pct`); si no, el error enumera los válidos.
2. Busca la zona de redacción indicada. Si no existe, el error **enumera las
   zonas disponibles** en el documento.
3. Si el ancla ya existe, no la duplica: avisa y devuelve `False`.
4. Construye un párrafo con el texto de antes, la región en línea con el valor
   inicial —por defecto una raya, `—`— y el texto de después, y lo añade **al
   final de la zona de redacción**, no del cuerpo.

El valor real llega en el siguiente refresco.

### 10.2 `desvincular_region(...)`

`desvincular_region(ruta, seleccion, verbose=True)`
(`src/fs_documento.py:1630`) es la operación inversa, y el docstring advierte
de que no hay marcha atrás automática:

> «Es la salida definitiva para una cifra que hay que escribir a mano y que
> debe SOBREVIVIR a los refrescos. El texto que hubiera dentro se conserva tal
> cual, pero ya no hay ancla: el motor no volverá a tocarlo ni lo reportará
> como huérfano. No tiene vuelta atrás automática; para volver a vincularla se
> usa "insertar".»

Técnicamente, promueve los hijos de la región al nivel del padre y elimina el
control: el texto sobrevive, la etiqueta desaparece. Acepta el tag completo o
solo la clave; con una clave suelta como `total_assets` desvincula a la vez
todas sus variantes (`-actual`, `-previo`, `-nota`…), porque comparten nombre.

---

## 11. Apariencia

Word dibuja un recuadro gris alrededor de cada control de contenido. El
comentario del bloque explica por qué eso conviene en unos sitios y estorba en
otros:

> «Para la tabla y las cifras eso es útil: se ve de un vistazo qué lo mantiene
> el Excel. Para las zonas de redacción y para los metadatos es un estorbo — la
> persona quiere escribir sobre papel en blanco, no dentro de una caja.»

El mecanismo es el atributo `w15:appearance`, de Word 2013 en adelante. Las
versiones anteriores lo ignoran y siguen mostrando el recuadro, *«que es un
fallo inofensivo»*.

| Familia | Apariencia | ¿Configurable? |
|---|---|---|
| `prosa`, `registro`, `meta` | `hidden` siempre | No |
| `tabla`, `campo`, `dato` | El valor de `apariencia_datos` | Sí: `boundingBox` o `hidden` |
| Controles ajenos al contrato | No se tocan | — |

`apariencia_de(familia, datos="boundingBox")` (`:1417`) resuelve esa tabla.
`normalizar_apariencia(doc, ...)` (`:1437`) la aplica a todo el documento y
**salta las regiones que ya coinciden**, que es lo que la hace idempotente;
opera en memoria y no guarda.

`_ocultar_parrafo(p)` (`:1458`) resuelve un detalle de OOXML que se nota mucho
en pantalla:

> «Sin esto, un párrafo de metadatos con el texto oculto sigue ocupando una
> línea en blanco: lo que se ve en pantalla es un renglón vacío que nadie sabe
> de dónde sale.»

La marca de párrafo tiene su propio conjunto de propiedades de fuente, distinto
del de los fragmentos de texto, y hay que ocultarla ahí. La función además
respeta el orden que exige la norma ECMA-376: el estilo va primero.

`simplificar_documento(ruta, quitar_prosa=False, datos="boundingBox")`
(`:1482`) es la orden de usuario: quita los recuadros de redacción y
metadatos, oculta de verdad el párrafo de metadatos y, opcionalmente, disuelve
las zonas de redacción promoviendo su contenido. Esto último lleva una
advertencia literal, porque tiene consecuencias:

> «OJO: sin zonas de redacción no se puede usar el modo estricto de dos
> editores (proteger). Se pueden recrear con "reparar".»

Las cifras y la tabla no se tocan nunca: su recuadro es informativo y su
candado es lo que impide pisarlas.

---

## 12. Los dos candados

Aquí está una de las distinciones más importantes del sistema, y la que más
confusión genera si no se explica. **Hay dos mecanismos de bloqueo, no uno**, y
protegen cosas distintas con fuerza distinta.

```
   CANDADO POR REGIÓN                    PROTECCIÓN DE DOCUMENTO
   w:lock="sdtContentLocked"             w:documentProtection

   Vive en cada región                   Vive en la configuración del .docx
   Word no deja teclear dentro           Word impone el modo solo lectura
   Buscar y reemplazar LO ATRAVIESA      No se atraviesa
   Word en el navegador no lo mira       Lo respeta
   Se pone y se quita sin contraseña     Contraseña con hash SHA-1
   cambiar_candado()                     proteger() / proteger_salvo_datos()
```

El propio código lo dice sin ambigüedad: el candado por región es *«una
comodidad de la interfaz»*; `documentProtection` es *«la protección que Word
IMPONE de verdad»*.

### 12.1 El candado por región

`cambiar_candado(ruta, bloquear=True, solo=None, verbose=True)`
(`src/fs_documento.py:1541`) pone o quita el candado sobre las regiones de
datos. Acepta un filtro `solo` que admite tanto el tag completo como la clave.
Al poner el candado, el elemento se inserta en el sitio que exige la norma:
después del identificador y antes del tipo de control. Al quitarlo, se elimina
el elemento en lugar de ponerlo a un valor «desbloqueado».

La advertencia del docstring es el punto que hay que entender:

> «Sin candado, sí [se puede teclear] — pero OJO: lo que se escriba a mano lo
> machaca el siguiente refresco, porque la región sigue vinculada al Excel.
> Para conservar un valor escrito a mano hay que DESVINCULAR la región.»

`estado_candado(ruta)` (`:1597`) devuelve la terna `(bloqueadas, total,
protección)` y existe por una razón de interfaz muy concreta:

> «Existe porque las opciones "permitir editar" y "volver a proteger" eran dos
> botones ciegos: nadie decia en que estado estaba el documento, asi que
> pulsarlos no parecia tener efecto.»

`_contar_candados(doc)` (`:1614`) acepta los dos valores posibles de bloqueo,
aunque el motor solo escriba uno: el otro puede venir de Word o del
complemento.

### 12.2 La protección de documento

`_hash_proteccion(clave_, salt, vueltas=100000)` (`:1693`) implementa el
algoritmo de contraseña de ECMA-376 tal como lo espera Word:

- Algoritmo **SHA-1**, identificado en el XML por `cryptAlgorithmSid="4"`.
- Semilla inicial: `SHA1(salt ‖ contraseña)`, con el salt **primero** y la
  contraseña codificada en UTF-16 little-endian, sin marca de orden de bytes ni
  terminador.
- **100 000 vueltas.** En cada iteración se recalcula `SHA1(hash ‖ i)`, donde
  `i` es el número de iteración en cuatro bytes little-endian, anexado
  **después** del hash.
- El salt son 16 bytes aleatorios, regenerados en cada protección.

`_poner_documentProtection(doc, clave_)` (`:1701`) escribe el elemento con los
atributos exactos que Word espera (`w:edit="readOnly"`, `w:enforcement="1"`,
`w:cryptProviderType="rsaFull"`, `w:cryptAlgorithmClass="hash"`,
`w:cryptAlgorithmType="typeAny"`, `w:cryptAlgorithmSid="4"`,
`w:cryptSpinCount="100000"`, más el hash y el salt en base64) y —esto importa—
lo inserta **en la posición que exige la norma dentro de la configuración del
documento**. La lista `_ORDEN_SETTINGS` existe justo para eso; el comentario:
*«documentProtection debe ir en su sitio o Word se queja al abrir.»*

### 12.3 Los dos modos de protección

Sobre esa base hay dos funciones que abren huecos editables en sitios opuestos:

**`proteger(ruta, clave_)`** (`:1739`) — el modelo de dos roles:

> «Rol REDACTOR: abre y solo puede escribir dentro de las zonas de prosa.
> Rol EDITOR DE DATOS: conoce la clave (o usa el add-in), que desprotege,
> refresca y vuelve a proteger.»

Abre un rango editable por cada zona de redacción. Es idempotente: si una zona
ya tiene su rango, la salta.

**`proteger_salvo_datos(ruta, clave_)`** (`:1793`) — la vuelta del revés:

> «En vez de abrir huecos en las zonas de redacción, abre huecos en TODO menos
> en la tabla, los campos y las cifras intercaladas. Resultado: se escribe
> donde se quiera, y los números son intocables de verdad.»

Este modo agrupa los hijos del cuerpo en tramos consecutivos que no sean
regiones de datos y abre un rango editable por tramo. Antes limpia todos los
rangos anteriores para no acumularlos, y se asegura de que el documento termine
en un párrafo libre:

> «Un párrafo libre de cierre. Sin él, si el documento termina en la tabla o en
> los metadatos, el último punto del documento cae fuera de todo rango editable
> y el redactor no puede añadir un párrafo al final.»

**Límite de granularidad que conviene conocer.** El criterio de agrupación
opera sobre los hijos de primer nivel del cuerpo. Una cifra `fs-dato-*`
embebida dentro de un párrafo de redacción queda, por tanto, **dentro de un
tramo editable**: a esa solo la protege su candado de región. La protección
«de verdad» cubre las regiones de datos que cuelgan directamente del cuerpo,
que en la práctica son la tabla y los campos sueltos.

`desproteger(ruta)` (`:1861`) retira la protección, pero **no retira los rangos
editables**: quedan en el XML, inertes, hasta la siguiente protección.

---

## 13. Interoperar con Excel por COM

Dos órdenes del sistema escriben en el libro de Excel: `nombrar`, que crea los
rangos con nombre que dan identidad estable a cada fila, y `tipos`, que fija en
una columna la clasificación que hoy se infiere. Las dos pilotan **Excel por
COM**, nunca `openpyxl`, y la razón está escrita dos veces en el código porque
la consecuencia de equivocarse es catastrófica y silenciosa:

> «Se hace con Excel (COM), NO con openpyxl: openpyxl no recalcula fórmulas y
> al reguardar descartaría el valor cacheado de todas ellas, dejando el Word en
> blanco. Excel guarda el libro con sus propias reglas y no toca ningún valor.»

El razonamiento completo: la capa de lectura abre el libro con `openpyxl` en
modo «solo valores», es decir, leyendo los resultados que Excel dejó guardados
en cada celda con fórmula. Si `openpyxl` reescribiera el archivo, esos
resultados se perderían —`openpyxl` no tiene motor de cálculo— y la siguiente
lectura devolvería vacío para toda celda con fórmula. El Word saldría en
blanco. La regla, por tanto, es fija: **`openpyxl` solo lee; Excel solo
escribe.**

Y por qué PowerShell en lugar de una biblioteca de Python para COM:

> «Conduce Excel desde PowerShell en vez de con pywin32. Evita una dependencia
> binaria pesada (que además complica el empaquetado con PyInstaller) y
> funciona en cualquier Windows con Excel instalado.»

### 13.1 `nombrar_rangos(...)`

`nombrar_rangos(xlsx, ctx, cfg, solo_simular=True, verbose=True)`
(`src/fs_documento.py:1878`). Cada nombre apunta a la **celda de la etiqueta**
de su fila, no a la de la cifra:

> «Así Excel reajusta la referencia solo cuando se insertan o borran filas
> encima, y el vínculo con el documento de Word no depende del texto de la
> etiqueta.»

Construye un plan con un nombre por línea que tenga fila y etiqueta —los
subtotales sin etiqueta se omiten—, usando el prefijo configurado (`fs_` por
defecto) y una referencia absoluta a la hoja. **El modo por defecto es
simulación**: sin `--aplicar` no toca nada, solo informa de cuántos nombres se
crearían.

### 13.2 `fijar_tipos(...)`

`fijar_tipos(xlsx, ctx, cfg, solo_simular=True, verbose=True)` (`:2041`). La
motivación es eliminar heurística frágil:

> «Mientras no exista esa columna, el tipo de cada fila se deduce de señales
> fragiles —negrita en las cifras, que la etiqueta empiece por "Total", que la
> fila no traiga numeros—. Funciona, pero cualquier retoque de formato puede
> cambiar en silencio como se clasifica una fila, y con ello el aspecto del
> documento. Fijarla convierte esa adivinanza en un dato declarado. No cambia
> nada hoy: escribe exactamente lo que ya se estaba infiriendo.»

Si el libro ya tiene columna `Tipo`, no toca nada. La columna de destino es la
forzada en la configuración o, si no hay ninguna, la primera libre a la derecha
de las que ya se usan.

### 13.3 Cómo se ejecutan

Ambas siguen el mismo patrón: se escribe el plan como JSON en un directorio
temporal, se invoca un script de PowerShell que abre Excel invisible y sin
diálogos, aplica los cambios uno a uno —el fallo de uno no aborta el lote—,
guarda y cierra, y devuelve el recuento por una línea con formato conocido. El
bloque `finally` del script llama a `Quit()` y libera el objeto COM: sin eso
quedaría un proceso `EXCEL.EXE` huérfano.

Los errores se traducen a lenguaje de usuario. Si no hay PowerShell:
*«Cree los nombres a mano: en Excel, seleccione la celda de la etiqueta y
escriba el nombre en el Cuadro de nombres (arriba a la izquierda).»* Si Excel
tarda demasiado: *«¿Está el libro abierto o pidiendo algo en pantalla?
Ciérrelo y vuelva a intentarlo.»*

Con `--aplicar`, la línea de órdenes deja siempre una copia `.bak` del libro
antes de tocarlo.

---

## 14. Diagnóstico

### 14.1 `verificar(ruta, ctx=None)`

`src/fs_documento.py:1349`. Solo lectura. Recorre el documento y devuelve un
inventario de lo que encuentra: tablas (y si tienen tabla dentro), campos y su
valor actual, cifras sueltas, zonas de redacción, si hay bitácora y metadatos,
y tres listas de problemas:

| Lista | Significa |
|---|---|
| `desconocidos` | Controles de contenido con etiqueta ajena al contrato. La orden los rotula *«Controles ajenos al contrato (se ignoran al refrescar)»*. |
| `huerfanos` | Anclas del documento sin valor en el Excel. |
| `sin_usar` | Cifras que el Excel ofrece, no vacías, que ninguna ancla del documento consume. |

Las dos últimas son direcciones opuestas del mismo desajuste, y juntas dan una
imagen completa de la correspondencia entre el libro y el documento.

### 14.2 `estado(cfg, xlsx_arg=None)`

`src/fs_documento.py:2760`. Es la radiografía del proyecto entero. Imprime
cinco secciones: el documento vivo (integridad, bloqueo, regiones, último
refresco y origen, estado de las cifras y modo de protección), el camino del
Word desechable, el libro de Excel tal como se está leyendo, los ejecutables
presentes en `dist/` y el estado del complemento de Word.

`_estado_documento(cfg)` (`:2716`) es la parte que reúne todo lo que se puede
saber del documento base **sin modificarlo**. Está enteramente envuelta en
manejo de errores: cualquier problema se recoge en una clave `error` en vez de
interrumpir el diagnóstico. Un diagnóstico que falla no sirve de nada.

Un detalle amable: si no encuentra el libro por convención, la inspección cae
sobre el libro de muestra de `ejemplos\`, porque *«para una simple inspección
vale el libro de muestra»*.

---

## 15. La línea de órdenes de `fs_documento.py`

El módulo se puede invocar directamente. Su despachador (`main`,
`src/fs_documento.py:2932`) reparte en dos grupos.

### 15.1 Órdenes que no necesitan documento

| Orden | Argumentos y banderas | Qué hace |
|---|---|---|
| `estado` | `[libro.xlsx]` | La radiografía completa del proyecto. |
| `catalogo` | `[libro.xlsx]` | Lista las claves disponibles en el Excel, con su origen (rango o etiqueta) y sus valores. |
| `tipos` | `[libro.xlsx]`, `--aplicar` | Escribe la columna `Tipo` en el libro. Sin `--aplicar`, solo simula. |
| `nombrar` | `[libro.xlsx]`, `--aplicar` | Crea los rangos con nombre `fs_*`. Sin `--aplicar`, solo simula. |
| `plantilla` | `<destino.docx>`, `--excel <libro>` | Crea un documento vivo desde cero. El contexto es opcional. |

### 15.2 Órdenes sobre el documento

Antes de despachar, el documento se resuelve con la cascada de §6.4. Si no se
pasó ninguno, se toma el de la configuración y se avisa por consola.

| Orden | Argumentos y banderas | Llama a | Respalda |
|---|---|---|---|
| `construir` / `reparar` | `[libro.xlsx]` | `construir` | Sí |
| `refrescar` | `[libro.xlsx]`, `--sin-registro` | `refrescar` | Sí |
| `insertar` | `<clave> <campo>`, `--zona`, `--antes`, `--despues` | `insertar_dato` | Sí |
| `apariencia` | `<visible\|invisible>` | `normalizar_apariencia` | Sí |
| `simplificar` | `--quitar-zonas` | `simplificar_documento` | Sí |
| `bloquear` / `desbloquear` | `[clave o tag]` | `cambiar_candado` | Sí |
| `desvincular` | `<clave o tag>` | `desvincular_region` | Sí |
| `limpiar-bitacora` | — | `quitar_registro_del_documento` | Sí |
| `verificar` | `[libro.xlsx]` | `verificar` | **No** |
| `proteger` | `--clave <X>` (obligatoria), `--salvo-datos` | `proteger` o `proteger_salvo_datos` | Sí |
| `desproteger` | — | `desproteger` | Sí |

Dos detalles del comportamiento:

- **`refrescar`** calcula la huella SHA-256 del libro y mete sus doce primeros
  caracteres en el origen. Ese dato queda en `fs-meta` y en la bitácora, de
  modo que siempre se puede saber de qué versión exacta del Excel salió cada
  refresco.
- **`desbloquear`** imprime la advertencia que cierra el modelo de datos:
  *«AVISO — lo que escriba a mano lo MACHACA el siguiente refresco: la región
  sigue vinculada al Excel. Si quiere que un valor escrito a mano sobreviva,
  desvincúlelo.»*

### 15.3 La convención de errores

El punto de entrada distingue dos clases de fallo, y la distinción es
consistente en todo el módulo:

- **`ValueError` = fallo previsto.** Se imprime el mensaje limpio, sin traza,
  bajo un rótulo `NO SE PUDO COMPLETAR`. Por eso todas las validaciones del
  módulo lanzan `ValueError` con textos largos, multilínea y con instrucciones
  de recuperación.
- **Cualquier otra excepción = defecto del programa.** Se muestra la traza
  completa bajo el rótulo `ERROR INESPERADO — copie este texto para soporte`.

---

## 16. Puntos frágiles de este módulo

Se recogen aquí para que quien mantenga el código los tenga a la vista. El
capítulo [Límites y riesgos](10-LIMITES-Y-RIESGOS.md) los sitúa en el conjunto
del sistema.

1. **Regiones duplicadas.** `_indexar` se queda con la primera aparición de
   cada ancla. Si un documento acaba con la misma etiqueta dos veces —al
   copiar y pegar un bloque, por ejemplo— la segunda queda congelada y **no se
   reporta**. El banco de pruebas comprueba que no se dupliquen, pero nada
   avisa si ya están duplicadas.
2. **La ventana de escritura no atómica** de `guardar_seguro` es una decisión
   consciente, pero real: entre el volcado y el truncado, un corte de energía
   deja el archivo inconsistente. La red de seguridad es el `.bak`, del que
   **solo hay una generación**: se sobrescribe en cada operación.
3. **El diff es textual.** Cambiar el formato de una celda en Excel se reporta
   como cambio de cifra.
4. **La protección no cubre las cifras en línea.** Una `fs-dato-*` dentro de un
   párrafo de redacción queda en un tramo editable bajo
   `proteger_salvo_datos`; solo la protege su candado de región, que Buscar y
   reemplazar atraviesa.
5. **`desproteger` deja los rangos editables** en el XML. Son inertes, pero se
   acumulan si se alterna entre los dos modos de protección.
6. **`_guardar_meta` no oculta la marca de párrafo.** Tras un refresco, el
   párrafo de metadatos puede dejar un renglón vacío visible hasta que se
   ejecute `simplificar`.
7. **El barrido de `_reubicar_perfil` no es recursivo:** solo mira el primer
   nivel de cada carpeta `OneDrive*`. Un documento en una subcarpeta no se
   encuentra por esa vía.
8. **El parseo de la línea de órdenes es artesanal.** El valor de una opción no
   empieza por `--` y por tanto también entra en la lista de argumentos
   posicionales. Además, la bandera `--desde` figura en el docstring del módulo
   pero **no está implementada**.
9. **La interoperación con Excel por COM no tiene pruebas automatizadas** y
   depende de que Excel esté instalado y no muestre ningún diálogo.
10. **`normalizar_apariencia` declara un parámetro `verbose` que no usa.** Es
    inocuo, pero conviene saberlo antes de confiar en él.

---

## Resumen del capítulo

- El motor actualiza el documento **en el sitio**: solo abre las regiones del
  contrato y no visita nada más, por lo que la redacción sobrevive a cada
  refresco y, si se borra, se queda borrada.
- Una región es un control de contenido de Word (`w:sdt`) cuya etiqueta es la
  identidad; el identificador se deriva del propio tag para que dos
  construcciones seguidas produzcan el mismo XML.
- `guardar_seguro` **no es atómico a propósito**: `os.replace()` cambia el
  identificador de archivo de NTFS y OneDrive lo interpreta como una
  desaparición, revirtiendo los cambios en silencio.
- Todo nombre de archivo que pasa por PowerShell viaja por archivo en UTF-8, no
  por la salida estándar, porque la página de códigos de la consola corrompía
  los acentos y los espacios duros.
- `construir` es idempotente por diseño: congela el índice al empezar y solo
  añade lo que faltaba antes de comenzar.
- El refresco reconstruye las tablas pero conserva sus anchos de columna, de
  modo que el ajuste manual del usuario sobrevive.
- El documento guarda su propia memoria en `fs-meta`, y el diff se calcula
  antes de escribir nada, emparejando por nombre y no por posición.
- Hay **dos candados**: el de región es una comodidad de la interfaz; la
  protección de documento es la que Word impone de verdad.
- `openpyxl` solo lee el libro y Excel solo lo escribe: invertirlo destruiría
  los valores cacheados de las fórmulas y dejaría el Word en blanco.
- La convención de errores es firme: `ValueError` para lo previsto, con
  instrucciones; cualquier otra excepción se trata como defecto y muestra la
  traza.
