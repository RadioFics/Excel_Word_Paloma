# Add-in de Word — actualizar el Estado Financiero in situ

Panel dentro de **Word** con un botón que trae las cifras del Excel modelo a
las tablas de **este mismo documento** — sin generar un archivo nuevo — y
deja una entrada de bitácora con los cambios.

> **Estado: v0 (andamiaje).** El código está escrito pero **no se ha
> compilado ni probado** todavía (necesita un equipo con Node.js). Sirve
> como base para iterar. Los criterios de mapeo (`src/core/tipos.ts`,
> `leer-excel.ts`) son el mismo modelo que `generador_fs.py`, portado a
> TypeScript.

## Qué hace

1. Eliges el `.xlsx` modelo (de tu OneDrive sincronizado). Se parsea en el
   navegador con `exceljs` — **no** sube a ningún servidor, no usa Graph.
2. Detecta hoja y columnas por contenido; infiere el `Tipo` de fila si la
   hoja no trae la columna. Muestra vista previa y qué se infirió.
3. Al pulsar **Actualizar**: reescribe las filas de la tabla marcada,
   reemplaza los campos de encabezado, calcula el diff contra la última
   versión aplicada (guardada dentro del propio `.docx`) y lo antepone a la
   sección de bitácora.

## Preparar la plantilla de Word (una vez)

Activa la pestaña **Programador** (Archivo → Opciones → Personalizar cinta).
Envuelve cada zona en un **Control de contenido de texto enriquecido** y
ponle una **Etiqueta** (Propiedades del control → Etiqueta):

| Etiqueta | Envuelve |
|---|---|
| `fs-tabla-principal` | la tabla del estado (4 columnas: Concepto · Nota · Periodo actual · Periodo comparativo, + fila de encabezado) |
| `fs-registro` | un párrafo/sección vacía para la bitácora (opcional) |
| `fs-titulo`, `fs-fecha-actual`, `fs-fecha-previa`, `fs-miles`, `fs-moneda` | los textos de encabezado (opcional) |

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
src/core/tipos.ts         modelo de tipos de fila + inferencia (== Python)
src/core/leer-excel.ts     parseo del .xlsx -> Contexto (detección por contenido)
src/core/escribir-word.ts  reescribe tablas y campos en el documento actual
src/core/registro.ts       snapshot en el .docx + cálculo de cambios
```

## Limitaciones conocidas de v0

- Sin probar/compilar (no había Node en el equipo de desarrollo inicial).
- Una sola tabla (`fs-tabla-principal`). Añadir notas = repetir el patrón
  con otras etiquetas.
- `addRows` asume que la tabla de la plantilla tiene exactamente 4 columnas.
- El formato fino (sangrías, líneas de subtotal) se hereda de la fila de
  encabezado de la plantilla; puede requerir ajuste por estilo de celda.
- Lee el Excel eligiéndolo a mano. La variante con Microsoft Graph
  (leer el libro directamente de OneDrive) es un paso posterior.
