# Add-in de Word — actualizar el Estado Financiero in situ

Panel dentro de **Word** con un botón que trae las cifras del Excel modelo a
las tablas de **este mismo documento** — sin generar un archivo nuevo — y
deja una entrada de bitácora con los cambios.

> **Estado: v0.2 (andamiaje, sin compilar).** El código está escrito pero
> **no se ha compilado ni probado** (necesita un equipo con Node.js). Los
> criterios de mapeo son el mismo modelo que `generador_fs.py`, portado a
> TypeScript.
>
> Lo que sí está probado de punta a punta es **el motor de Python
> equivalente** (`fs_documento.py`, en la raíz): mismo contrato de anclas,
> misma semántica de refresco, verificado abriendo los documentos en Word.
> El add-in es la misma operación desde un panel; el motor de Python es la
> referencia de comportamiento.

## El contrato

Ambos lados —este add-in y `fs_documento.py`— hablan el mismo vocabulario de
anclas, especificado en **[`../CONTRATO.md`](../CONTRATO.md)** y portado en
`src/core/contrato.ts`. **Si cambia uno, cambia el otro.**

Regla central: el refresco solo escribe dentro de controles de las familias
`fs-tabla-*`, `fs-campo-*`, `fs-dato-*`, `fs-registro` y `fs-meta`. Las zonas
`fs-prosa-*` y la prosa suelta no se visitan, y por eso la redacción
sobrevive intacta.

## Qué hace

1. Eliges el `.xlsx` modelo (de tu OneDrive sincronizado). Se parsea en el
   navegador con `exceljs` — **no** sube a ningún servidor, no usa Graph.
2. Detecta hoja y columnas por contenido; infiere el `Tipo` de fila si la
   hoja no trae la columna. Muestra vista previa y qué se infirió.
3. Al pulsar **Actualizar**: reescribe las filas de la tabla marcada,
   reemplaza los campos de encabezado, calcula el diff contra la última
   versión aplicada (guardada dentro del propio `.docx`) y lo antepone a la
   sección de bitácora.

## Preparar el documento base (una vez)

Lo más rápido es dejar que lo arme el motor de Python:

```
python src\fs_documento.py construir "MI_DOCUMENTO.docx"
```

Deja todas las anclas puestas, bloqueadas y con el formato correcto. Es
idempotente: se puede correr sobre un documento con redacción encima.

A mano, desde Word: pestaña **Programador** (Archivo → Opciones →
Personalizar cinta) → **Control de contenido** → *Propiedades* → **Etiqueta**.
Ver [`../CONTRATO.md`](../CONTRATO.md) para el vocabulario completo.

## Desarrollo (equipo con Node.js 18+)

```
cd addin
npm install
npm start          # compila, levanta https://localhost:3000 y abre Word con el add-in cargado
```

`npm start` usa `office-addin-debugging`, que instala un certificado de
desarrollo local y hace *sideload* del `manifest.xml` en Word.

Para parar: `npm run stop`.

Faltan los iconos: crea `assets/icon-16.png`, `icon-32.png`, `icon-80.png`
(PNG cuadrados). Sin ellos el add-in carga igual pero sin icono en la cinta.

## Despliegue para las usuarias (sin instalar nada)

1. `npm run build` → genera `dist/` (archivos estáticos).
2. Sube `dist/` a una carpeta de **SharePoint** o de un sitio interno HTTPS.
3. En `manifest.xml`, cambia todas las URLs `https://localhost:3000` por la
   URL real de esa carpeta, y genera un `<Id>` propio (`[guid]::NewGuid()`).
4. En el **Centro de administración de Microsoft 365 → Aplicaciones
   integradas → Cargar aplicación personalizada**, sube el `manifest.xml` y
   acótalo a Pamela y Violeta. (Para pruebas, cada usuaria puede hacer
   *sideload* ella misma: Word → Inicio → Complementos → «Cargar mi
   complemento».)

Todo el tráfico queda en el tenant. La única intervención de TI es aprobar
ese despliegue centralizado (y, si más adelante se lee el Excel por Graph en
vez de elegir archivo, consentir un permiso de solo lectura).

## Archivos

```
manifest.xml            manifiesto del add-in (Word, ReadWriteDocument)
package.json             dependencias y scripts
webpack.config.js        build + dev-server HTTPS
src/taskpane/            panel: HTML, CSS, TS (orquesta el flujo)
src/core/contrato.ts      vocabulario de anclas (puerto de fs_contrato.py)
src/core/tipos.ts         modelo de tipos de fila + inferencia (== Python)
src/core/leer-excel.ts     parseo del .xlsx -> Contexto (detección por contenido)
src/core/escribir-word.ts  refresca las regiones de datos del documento actual
src/core/registro.ts       snapshot en el .docx + cálculo de cambios
```

## Limitaciones conocidas

- **Sin compilar ni probar** (no hay Node en el equipo de desarrollo). El
  comportamiento de referencia es `fs_documento.py`, que sí está verificado.
- `addRows` asume que la tabla tiene exactamente 4 columnas.
- El formato por tipo de fila que aplica el add-in (negrita en H/S/T) es más
  pobre que el del motor de Python (sangrías, filetes de subtotal y total).
  Un documento construido con Python y refrescado con el add-in **pierde
  esos filetes**. Igualarlo es el siguiente trabajo pendiente.
- Falta en el panel el catálogo de claves y el botón «insertar dato»
  (`insertarDato()` ya existe en `escribir-word.ts`, pero no está cableado
  al HTML).
- No aplica ni retira `documentProtection`: el modo estricto de dos editores
  hay que gestionarlo hoy desde Python o desde Word.
- Lee el Excel eligiéndolo a mano. La variante con Microsoft Graph (leer el
  libro directamente de OneDrive) es un paso posterior.
