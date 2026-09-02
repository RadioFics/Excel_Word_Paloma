# Manual de ingeniería · Estados Financieros (Excel → Word)

Documentación completa del sistema que mantiene al día un documento de Word con
las cifras de un libro de Excel **sin borrar la redacción** que alguien
escribió alrededor.

Este manual describe el repositorio `Excel_Word_Paloma` en su versión **0.12**.
Es la documentación de fondo: el `README.md` de la raíz explica cómo *usar* el
programa; esto explica **cómo funciona y por qué es así**.

---

## Por dónde empezar

Depende de a qué haya venido.

| Si usted… | Lea, en este orden |
|---|---|
| Quiere entender qué es esto y para qué sirve | [01](01-PANORAMA.md) y nada más |
| Va a mantener o modificar el código | [01](01-PANORAMA.md) → [02](02-ARQUITECTURA.md) → [04](04-MOTOR-DEL-DOCUMENTO.md) → [03](03-LECTURA-DEL-EXCEL.md) |
| Tiene que dar soporte a una usuaria | [01](01-PANORAMA.md) → [05](05-INTERFAZ-Y-ORDENES.md) → [10](10-LIMITES-Y-RIESGOS.md) |
| Va a montar el proyecto en un equipo nuevo | [07](07-DISTRIBUCION-Y-ENTORNO.md) → [08](08-PRUEBAS.md) |
| Tiene que decidir si se retoma el complemento de Office | [01 §7](01-PANORAMA.md) → [06](06-COMPLEMENTO-DE-OFFICE.md) → [10 §6](10-LIMITES-Y-RIESGOS.md) |
| Tiene que explicar el proyecto a dirección o a TI | [01](01-PANORAMA.md) → [09](09-HISTORIA-DE-VERSIONES.md) → [10 §7](10-LIMITES-Y-RIESGOS.md) |
| Quiere saber hasta dónde confiar en el sistema | [10](10-LIMITES-Y-RIESGOS.md) directamente |

---

## Los capítulos

| # | Capítulo | De qué trata |
|---|---|---|
| 01 | [Panorama y conceptos básicos](01-PANORAMA.md) | El problema, la idea, el vocabulario, las piezas y los tres frentes que se evaluaron. Sin tecnicismos. |
| 02 | [Arquitectura y funcionamiento interno](02-ARQUITECTURA.md) | Las cuatro capas, el contrato de anclas como costura, el ciclo de vida de un documento y los doce invariantes de diseño. |
| 03 | [La lectura del Excel y el contrato de anclas](03-LECTURA-DEL-EXCEL.md) | Cómo se identifica la hoja, las columnas, el encabezado y el tipo de cada fila, con los umbrales exactos. La especificación del contrato. |
| 04 | [El motor del documento](04-MOTOR-DEL-DOCUMENTO.md) | El capítulo central. Cómo se escribe sobre un `.docx` sin destruirlo, cómo se monta el andamiaje, qué hace un refresco, la memoria del documento y los dos candados. |
| 05 | [La interfaz y las órdenes](05-INTERFAZ-Y-ORDENES.md) | El menú, la ventana (que no es tkinter), todas las banderas de la línea de órdenes y la configuración en uso. |
| 06 | [El complemento de Office](06-COMPLEMENTO-DE-OFFICE.md) | El complemento de Word en TypeScript: qué hay escrito, qué le falta y qué haría falta para publicarlo. |
| 07 | [Distribución, empaquetado y entorno](07-DISTRIBUCION-Y-ENTORNO.md) | El Python embebido, los lanzadores, los tres ejecutables, la publicación de una versión y la higiene del repositorio. |
| 08 | [El banco de pruebas](08-PRUEBAS.md) | Las 27 comprobaciones, qué invariante defiende cada una, la prueba de humo y lo que no está cubierto. |
| 09 | [Historia de las versiones](09-HISTORIA-DE-VERSIONES.md) | Las quince entregas, qué cambió en cada una y qué significa. Reconstruida desde los objetos de Git. |
| 10 | [Límites, riesgos y deuda técnica](10-LIMITES-Y-RIESGOS.md) | Todo lo que el sistema no garantiza, reunido y ordenado por gravedad, con una propuesta de orden de trabajo. |

---

## Las tres ideas que hay que llevarse

Si solo se leen tres frases de todo el manual, que sean estas.

**1 · El motor solo escribe dentro de regiones marcadas.** No es que respete la
redacción: es que no pasa por ella. Por eso el texto sobrevive a cada refresco
y, si se borra, se queda borrado.

**2 · Nada se localiza por coordenadas.** Ni la hoja, ni las columnas, ni el
encabezado, ni las regiones del documento. Todo se identifica por su contenido
o por su etiqueta, y por eso el libro y el documento pueden cambiar de forma
sin romper nada.

**3 · Casi cada mecanismo raro del sistema tiene un incidente detrás.** El
guardado que no es atómico, los nombres de archivo que viajan por archivo y no
por la consola, las tres funciones que reparan nombres con las tildes comidas:
ninguna es una preferencia estética. El capítulo 02 las reúne como invariantes
de diseño, y conviene leerlas antes de «simplificar» algo.

---

## Relación con la documentación existente

Este manual no sustituye a los documentos de `docs/`, los complementa. Aquellos
responden a «cómo hago X»; este responde a «cómo funciona X y por qué».

| Documento existente | Para qué |
|---|---|
| [`../GUIA.md`](../GUIA.md) | Guía de operación paso a paso. Es el punto de partida para usar el programa. |
| [`../DATOS.md`](../DATOS.md) | Dónde se editan las cifras y cómo llevarlas a la redacción. |
| [`../CONTRATO.md`](../CONTRATO.md) | La especificación normativa de las anclas. El capítulo 03 la explica; este documento la define. |
| [`../CAMBIAR_EXCEL.md`](../CAMBIAR_EXCEL.md) | Qué revisar cuando llegue un Excel nuevo. |
| [`../VINCULOS_NATIVOS.md`](../VINCULOS_NATIVOS.md) | Vincular Excel y Word sin programar, y qué se rompe. |
| [`../ESTRUCTURA.md`](../ESTRUCTURA.md) | Organización de las carpetas. |
| [`../DESPLIEGUE_ADDIN.md`](../DESPLIEGUE_ADDIN.md) | Cómo subir el complemento y qué pedirle a TI. |
| [`../INSTALACION.md`](../INSTALACION.md) | Montar el entorno de desarrollo. |
| [`../PRUEBA_EXTERNA.md`](../PRUEBA_EXTERNA.md) | Reproducir la prueba en otro equipo. |
| [`../DIRECCION.md`](../DIRECCION.md) | Dirección del proyecto y el caso ante TI. |

---

## Cómo se escribió este manual

Leyendo el código fuente completo —los cinco módulos de Python, los ocho
archivos del complemento, los scripts de PowerShell y de proceso por lotes— y
reconstruyendo la historia de versiones desde los objetos de Git, commit a
commit y archivo a archivo.

Las citas entre comillas angulares son textuales de los comentarios del autor.
Las referencias con la forma `src/archivo.py:123` apuntan a la línea exacta en
la versión 0.12. Lo que sea deducción y no afirmación del código está marcado
como inferencia.

Si el código cambia, lo primero que envejece son los números de línea; el
segundo, los umbrales de las heurísticas del capítulo 03.
