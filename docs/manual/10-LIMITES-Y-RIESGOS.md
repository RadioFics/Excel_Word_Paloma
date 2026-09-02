# 10 · Límites, riesgos y deuda técnica

> **Para quién.** Quien tenga que decidir hasta dónde confiar en el sistema,
> quién lo vaya a mantener, y quien deba priorizar el trabajo que queda.
> **Qué encontrará.** Todo lo que el sistema no hace o no garantiza, reunido en
> un solo sitio y ordenado por gravedad; los riesgos abiertos que no dependen
> del código; y una propuesta de orden de trabajo.
> **Antes de leer.** Cada punto remite al capítulo donde se explica en detalle.
> Este capítulo reúne, no repite.

## Cómo leer este capítulo

Las limitaciones se clasifican así:

| Marca | Significa |
|---|---|
| **Por diseño** | Es una decisión consciente con una razón detrás. No se «arregla»: se conoce. |
| **Deuda** | Habría que corregirlo; no es urgente. |
| **Riesgo** | Puede causar un daño silencioso. Merece atención. |
| **Abierto** | No depende del código; depende de una decisión de otro. |

---

## 1. Los límites declarados por el propio proyecto

El `README.md` los enumera, y son el punto de partida:

| Límite | Clase |
|---|---|
| Un solo tipo de estado financiero: Situación Financiera. Otro estado exige otra plantilla. | Por diseño |
| La deducción del tipo de cada fila es heurística; por eso existe `revisar_tipos.csv`. | Por diseño |
| Las cifras sueltas dentro de párrafos hay que marcarlas una vez: no se detectan solas. | Por diseño |
| El complemento de Word está escrito pero **sin compilar**. | Abierto |
| La coautoría de Word en el navegador maneja mal los controles de contenido: hay que refrescar desde Word de escritorio. | Por diseño |
| Solo Windows de 64 bits. | Por diseño |

---

## 2. Riesgos de pérdida o corrupción de datos

Son los que importan de verdad, porque el sistema escribe sobre un documento
que contiene trabajo de una persona.

### 2.1 La ventana de escritura no atómica · **Por diseño / Riesgo**

El guardado escribe sobre el archivo original en lugar de reemplazarlo, para
que OneDrive no lo interprete como una desaparición y reponga la versión del
servidor. El precio es que, entre el volcado y el truncado, un corte de energía
o un fallo del disco deja el archivo inconsistente.

La red de seguridad es la copia `.bak`, y **solo hay una generación**: se
sobrescribe en cada operación. Un fallo detectado dos refrescos después ya no
tiene copia a la que volver.

*Mitigación práctica:* el historial de versiones de OneDrive cubre ese hueco.
Conviene saber que existe y cómo se usa. → [04 §3](04-MOTOR-DEL-DOCUMENTO.md)

### 2.2 Regiones duplicadas · **Riesgo**

El índice de anclas se queda con la primera aparición de cada etiqueta. Si un
documento acaba con la misma región dos veces —al copiar y pegar un bloque en
Word, por ejemplo—, la segunda **queda congelada y nada avisa**. El banco de
pruebas comprueba que el motor no las duplique, pero no hay comprobación del
comportamiento ante un documento que ya llegue duplicado.

*Trabajo pendiente:* que `verificar` reporte duplicados, y que el refresco los
señale en el informe. → [04 §16](04-MOTOR-DEL-DOCUMENTO.md)

### 2.3 Una cifra escrita a mano se pierde en el siguiente refresco · **Por diseño**

Es coherente y está advertido en la propia orden de desbloqueo, pero es el
error de uso más probable. La salida correcta es *desvincular* la región, y esa
operación **no tiene vuelta atrás automática**.

### 2.4 El candado de región no es una protección · **Por diseño**

Buscar y reemplazar lo atraviesa, y Word en el navegador lo ignora. Solo la
protección de documento impone algo de verdad. Y esta última tiene un límite de
granularidad: una cifra embebida dentro de un párrafo de redacción queda en un
tramo editable, protegida únicamente por su candado.
→ [04 §12](04-MOTOR-DEL-DOCUMENTO.md)

---

## 3. Riesgos de lectura silenciosa

Estos son más insidiosos que los anteriores, porque no producen un error: dan
un resultado plausible pero equivocado.

| # | Qué pasa | Clase |
|---|---|---|
| 1 | **El orden de las columnas de cifras es una convención implícita.** Se asume que el periodo actual está a la izquierda del comparativo. Un libro que las invierta produce el documento con los periodos cambiados, sin ningún aviso. Solo se corrige forzando las letras en la configuración. | Riesgo |
| 2 | **La detección de la columna de nota es estrecha:** exige enteros pequeños y casi ningún texto. Una nota con valor `100` la vuelve invisible. | Deuda |
| 3 | **`titulo` es el primer texto que haya encima de la región**, sea cual sea. Con frecuencia acaba siendo el nombre de la empresa y no el del estado. | Deuda |
| 4 | **Los marcadores de hoja por defecto se solapan** (`assets` es subcadena de `total assets`), de modo que el umbral de puntuación es más laxo de lo que su valor sugiere. | Deuda |
| 5 | **El reconocedor de fechas acepta subcadenas:** `mayor` o `maybe` marcan una fila como fila de fechas. | Deuda |
| 6 | **La clasificación de filas depende de la palabra literal `total` y de un corte de longitud arbitrario.** Sin columna `Tipo`, falla de forma predecible en libros redactados de otra manera. | Por diseño |
| 7 | **La configuración se reemplaza clave por clave, sin mezclar listas.** Quien redefina los marcadores de hoja para añadir uno los sustituye todos. | Deuda |
| 8 | **Renombrar una fila rompe el vínculo** si esa fila no tiene rango con nombre. | Por diseño |
| 9 | **El diff es textual sobre valores ya formateados:** cambiar los decimales de una celda en Excel se reporta como cambio de cifra. | Por diseño |

Los puntos 1 a 5 tienen la misma forma: heurísticas afinadas contra un libro
concreto. Funcionan, y el sistema ofrece la válvula de escape de forzar
columnas y hoja en la configuración. Pero conviene correr la inspección en seco
(`--estado`) **cada vez que llegue un libro nuevo**, que es exactamente para lo
que existe. → [03 §16](03-LECTURA-DEL-EXCEL.md), [05](05-INTERFAZ-Y-ORDENES.md)

---

## 4. Deuda en el código

| # | Dónde | Qué | Clase |
|---|---|---|---|
| 1 | Línea de órdenes del motor | El parseo es artesanal: el valor de una opción también entra en la lista de argumentos posicionales. | Deuda |
| 2 | Línea de órdenes del motor | La bandera `--desde` figura en la documentación del módulo pero **no está implementada**. | Deuda |
| 3 | Lector de Excel | Las banderas desconocidas se ignoran en silencio: un `--revizar` mal escrito genera el documento sin avisar. | Deuda |
| 4 | Lector de Excel | El sello de tiempo de los archivos generados tiene resolución de minuto: dos ejecuciones en el mismo minuto escriben sobre el mismo archivo. | Deuda |
| 5 | Motor | Al guardar los metadatos no se oculta la marca de párrafo, así que tras un refresco puede quedar un renglón vacío visible hasta ejecutar `simplificar`. | Deuda |
| 6 | Motor | `desproteger` no retira los rangos editables: quedan inertes en el XML y se acumulan al alternar modos. | Deuda |
| 7 | Motor | El barrido que reubica un documento en otro perfil **no es recursivo**: solo mira el primer nivel de cada carpeta OneDrive. | Deuda |
| 8 | Motor | `normalizar_apariencia` declara un parámetro `verbose` que no usa. | Deuda |
| 9 | `generador_fs.py` | Mezcla dos responsabilidades: la capa de lectura y el generador clásico de la foto. Es la única mezcla de capas del sistema. | Deuda |
| 10 | Empaquetado | `GeneradorFS.spec` está versionado pero **no participa en ninguna compilación**: el flujo real llama a PyInstaller por línea de órdenes. Además cubre uno de los tres ejecutables y no vacía la ruta del documento base. | Riesgo |
| 11 | Documentación | El `README.md` habla de catorce pruebas; el banco tiene veintisiete. | Deuda |

El punto 10 es el más peligroso de la lista, porque un archivo versionado que
parece la especificación de empaquetado y no lo es invita a que alguien lo
edite creyendo que cambia algo. → [07 §4](07-DISTRIBUCION-Y-ENTORNO.md)

---

## 5. Lo que no está probado

El banco de pruebas cubre bien el refresco y la clasificación de documentos. Lo
que queda fuera es sustancial:

- La protección con contraseña, en sus dos modos, y el cálculo del hash.
- El candado por región.
- La interoperación con Excel por COM, que es la única operación que **escribe
  en el libro**.
- Toda la interfaz de ventana y los diálogos de archivo.
- El complemento de Office, entero.
- La apariencia y la simplificación de documentos.
- La bitácora, que el banco desactiva expresamente.
- La cascada de resolución de rutas, salvo el caso del nombre con tildes.
- Que los ejecutables se construyan y arranquen.

→ [08 §7](08-PRUEBAS.md)

---

## 6. El complemento de Office

Es el capítulo aparte, porque su estado no es «con defectos» sino «sin
verificar». **Nunca se ha compilado ni ejecutado**, así que todo lo que sigue
son defectos leídos en el código, no observados.

Los más relevantes:

| # | Qué | Consecuencia |
|---|---|---|
| A | El refresco desbloquea, escribe y vuelve a bloquear **sin `try/finally`**. | Un fallo intermedio deja las regiones de datos desbloqueadas. |
| B | El diff se indexa por el texto de la etiqueta, no por la clave estable. | Renombrar una fila genera falsos «retirada» más «nueva»; los subtotales sin etiqueta colapsan en una sola entrada. |
| C | **`fs-meta` nunca se escribe.** | Los dos motores no comparten estado: refrescar con Python y luego con el complemento produce bitácoras incoherentes. |
| D | Las colisiones de clave se descartan. | Dos filas homónimas se pisan en silencio. |
| E | Hay tres versiones declaradas y dos valores distintos. | Confusión al desplegar. |
| F | El manifiesto conserva un identificador de ejemplo y todas las URL apuntan a `localhost`. | No se puede publicar tal cual. |

Y a eso se suman los obstáculos previsibles del primer intento de compilación
—dependencias que Webpack 5 ya no inyecta solo, la hoja de estilos que nunca
llega al paquete, cargadores que faltan— que se detallan en el capítulo 6.

*(Inferencia: ninguno de esos obstáculos se ha podido verificar, precisamente
porque nunca se ha compilado.)* → [06 §11](06-COMPLEMENTO-DE-OFFICE.md)

---

## 7. Riesgos abiertos, que no dependen del código

Son los que pueden detener el proyecto y no se resuelven programando.

**Control de aplicaciones · Abierto.** Es el riesgo declarado desde V0.1, y
sigue vigente. El README y `docs/DIRECCION.md` coinciden:

> «Si el equipo aplica control de aplicaciones, puede bloquear la ejecución de
> `python.exe` desde una carpeta de usuario o una unidad de red aunque no haya
> "instalación". **Confirmar con TI** antes de depender de esto.»

Si se confirma que hay control de aplicaciones, el frente del Python portable
queda descartado y hay que acelerar el complemento de Office.

**Sin firma de código · Abierto.** Los ejecutables no están firmados. Eso
produce la advertencia de SmartScreen en cada descarga y hace que la integridad
descanse enteramente en publicar y comparar las huellas SHA-256. Es funcional,
pero exige disciplina en cada publicación. Y hay una frontera que conviene
tener clara: la advertencia de SmartScreen se salta; **un bloqueo por política
corporativa no se salta, se escala**.

**Dueño de mantenimiento · Abierto.** `docs/DIRECCION.md` lo plantea como
decisión pendiente: *«¿Quién es el dueño de mantenimiento de plantillas y del
add-in a un año?»* Sigue sin respuesta, y es la que decide si conviene o no
retomar el complemento.

**Concentración de conocimiento · Riesgo.** El proyecto entero se escribió en
tres días por una persona. La documentación es buena y los comentarios del
código explican el porqué de cada decisión rara —este manual se ha podido
escribir gracias a eso—, pero nadie más lo ha modificado nunca.

---

## 8. Propuesta de orden de trabajo

No es una hoja de ruta del producto: es el orden en que conviene pagar la
deuda, ordenado por relación entre riesgo evitado y esfuerzo.

**Primero, lo barato que evita daño silencioso:**

1. Que `verificar` reporte regiones duplicadas y que el refresco las señale
   (§2.2). Es el único punto donde el sistema puede fallar sin decirlo.
2. Borrar o rehacer `GeneradorFS.spec` para que el repositorio no contenga un
   archivo engañoso (§4.10).
3. Corregir el número de pruebas del README (§4.11).

**Después, lo que reduce la fragilidad de la lectura:**

4. Ejecutar `nombrar --aplicar` sobre el libro de producción. Convierte la
   identidad de cada fila en estable y elimina de un golpe el riesgo de que
   renombrar una etiqueta rompa un vínculo.
5. Ejecutar `tipos --aplicar`. Convierte la clasificación heurística en un dato
   declarado, y con ello desaparecen los puntos 5 y 6 de §3.
6. Documentar en la configuración el orden esperado de las columnas de cifras,
   o detectarlo por las fechas del encabezado en lugar de asumirlo (§3.1).

**Luego, ampliar la red de pruebas** hacia lo que hoy no cubre, empezando por lo
que escribe: la protección y los candados.

**Y en paralelo, la decisión abierta:** confirmar con TI el control de
aplicaciones. De esa respuesta depende si el complemento de Office es un
proyecto futuro o una urgencia.

---

## Resumen del capítulo

- Los límites declarados por el proyecto son honestos y están en el README; lo
  que este capítulo añade sale de leer el código.
- El riesgo de pérdida de datos está acotado por diseño, pero descansa en una
  sola generación de respaldo: conviene apoyarse también en el historial de
  versiones de OneDrive.
- El único fallo verdaderamente silencioso del motor son las **regiones
  duplicadas**: se refresca la primera y nada avisa.
- Las heurísticas de lectura están afinadas contra un libro concreto; correr la
  inspección en seco cuando llegue un libro nuevo es una obligación, no una
  recomendación.
- Aplicar `nombrar` y `tipos` sobre el libro de producción elimina de golpe
  buena parte de la fragilidad de la capa de lectura.
- El complemento de Office no está «con defectos»: está **sin verificar**, y
  eso incluye si compila.
- El riesgo que puede detener el proyecto no es técnico: es la política de
  control de aplicaciones, pendiente de confirmar con TI desde la primera
  versión.
