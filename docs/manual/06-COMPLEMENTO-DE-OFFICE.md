# 06 · El complemento de Office

> **Para quién.** Quien tenga que decidir si se retoma el complemento de Word,
> quien deba compilarlo por primera vez, y quien tenga que sostener la
> conversación con el área de TI.
> **Qué encontrará.** El argumento organizativo que justifica el complemento,
> el inventario exacto de lo que hay escrito, la lectura archivo por archivo
> de los ocho fuentes, la comparación con el núcleo de Python, y la lista
> ordenada de lo que falta para que alguien pueda usarlo.
> **Antes de leer.** Conviene haber leído [Arquitectura](02-ARQUITECTURA.md)
> —de ahí sale la comparación de los tres frentes— y tener presente el
> contrato de anclas (`docs/CONTRATO.md`). Este capítulo describe código que
> **nunca se ha compilado ni ejecutado**.

## Índice del capítulo

1. [Por qué existe](#1-por-qué-existe)
2. [Estado real: escrito y sin compilar](#2-estado-real-escrito-y-sin-compilar)
3. [El manifiesto](#3-el-manifiesto)
4. [La compilación prevista](#4-la-compilación-prevista)
5. [El modelo de datos](#5-el-modelo-de-datos)
6. [El contrato compartido](#6-el-contrato-compartido)
7. [La lectura del libro](#7-la-lectura-del-libro)
8. [La escritura en Word](#8-la-escritura-en-word)
9. [La bitácora y la foto de la última actualización](#9-la-bitácora-y-la-foto-de-la-última-actualización)
10. [El panel](#10-el-panel)
11. [Qué falta para poder usarlo](#11-qué-falta-para-poder-usarlo)
12. [Si algún día se retoma](#12-si-algún-día-se-retoma)

---

## 1. Por qué existe

El diagnóstico está escrito en `docs/DIRECCION.md`, en la primera frase del
documento: el piloto en Python **demostró que la lógica de mapeo funciona**;
lo que no encaja en una organización que bloquea instaladores y revisa
ejecutables es el **vehículo de entrega**. Un intérprete de Python o un
`.exe` sin firma digital, ejecutado desde una carpeta de usuario, queda fuera
de los controles del inquilino corporativo (el espacio de Microsoft 365 de la
organización). Es material que AppLocker, WDAC o el antivirus de puesto
pueden frenar sin previo aviso.

De ahí el plan en tres frentes paralelos que fija `docs/DIRECCION.md` y que
[Arquitectura](02-ARQUITECTURA.md) desarrolla:

| Frente | Qué es | Horizonte | Aprobación de TI |
|---|---|---|---|
| **A. Tablas vinculadas** | Pegado con vínculo de rangos de Excel en Word y «actualizar todo» | Ya | Ninguna (es nativo) |
| **B. Python portable** | Paquete embebido en `.\python\`, herramienta de las dos usuarias | Ya | Consultar AppLocker / EDR |
| **C. Complemento de Office** | Panel dentro de Word con un botón que actualiza **el mismo documento** | Semanas | Una revisión y despliegue centralizado |

El frente C es el complemento que describe este capítulo. Su argumento es de
naturaleza administrativa, no técnica, y `docs/DESPLIEGUE_ADDIN.md` lo enuncia
con precisión: un complemento de Office **no es un programa que se instala**;
es una página web que Word carga dentro de un panel lateral. De ahí se sigue
todo lo demás. La usuaria final no descarga nada, no necesita permisos de
administrador y no hay ningún ejecutable que un antivirus pueda poner en
cuarentena. El código vive en una biblioteca de SharePoint del propio
inquilino; el manifiesto —el archivo XML que le dice a Word dónde está esa
página y qué permisos pide— se publica una sola vez desde el centro de
administración de Microsoft 365, en **Aplicaciones integradas**, y se acota a
las dos usuarias. La única intervención de TI es aprobar ese despliegue.

Hay un matiz que conviene no perder. `docs/DIRECCION.md` incorpora una
actualización posterior: el frente B ya cumple el requisito explícito del
cliente —que **el mismo documento actualice sus datos**, sin generar uno
nuevo— porque `fs_documento.py` refresca en el sitio conservando la
redacción, con el mismo contrato de anclas. La consecuencia está escrita
allí mismo:

> Con eso, **el frente B deja de ser un puente y pasa a ser la solución
> operativa** mientras el frente C consigue su revisión de TI. El frente C
> ya no tiene que demostrar el concepto: solo trasladar la misma operación
> —ya especificada y probada— a un panel dentro de Word.

Es decir: el complemento no es un experimento pendiente de validar. Es un
traslado de vehículo. La especificación que debe cumplir ya está probada en
Python.

---

## 2. Estado real: escrito y sin compilar

Conviene decirlo antes que nada, porque el volumen de código puede dar una
impresión equivocada. El complemento **está escrito y nunca se ha
compilado**. Su propio `addin/README.md` lo declara sin rodeos:

> **Estado: v0.2 (andamiaje, sin compilar).** El código está escrito pero
> **no se ha compilado ni probado** (necesita un equipo con Node.js). Los
> criterios de mapeo son el mismo modelo que `generador_fs.py`, portado a
> TypeScript.

El estado del directorio lo confirma. No existe `node_modules/` (las
dependencias descargadas), ni `dist/` (el resultado de la compilación), ni
`package-lock.json` (el archivo que fija las versiones exactas y que `npm`
crea en la primera instalación), ni pruebas, ni configuración de linter. La
carpeta `assets/`, que debería contener los tres iconos del botón, contiene
un único archivo de texto que avisa de que faltan.

```
addin/
├── README.md              5.362 B
├── manifest.xml           4.338 B
├── package.json           1.023 B
├── tsconfig.json            371 B
├── webpack.config.js      1.458 B
├── assets/
│   └── LEEME.txt            303 B   <- único archivo; NO hay iconos PNG
└── src/
    ├── core/
    │   ├── contrato.ts       209 líneas
    │   ├── escribir-word.ts  194 líneas
    │   ├── leer-excel.ts     408 líneas
    │   ├── registro.ts        79 líneas
    │   └── tipos.ts          136 líneas
    └── taskpane/
        ├── taskpane.css     2.137 B
        ├── taskpane.html       39 líneas
        └── taskpane.ts       132 líneas
```

Son 1.158 líneas de TypeScript que cubren el flujo completo y son
internamente coherentes, pero de las que no existe evidencia de una sola
compilación. Y están congeladas desde una versión temprana del proyecto: el
comentario de cabecera de `addin/src/taskpane/taskpane.ts:8` todavía enumera
un mapa de anclas que el contrato vigente ya no reconoce (`fs-fecha-actual`,
`fs-miles`, `fs-moneda`, `fs-titulo`), cuando el contrato actual exige la
familia `campo` y guiones bajos: `fs-campo-fecha_actual`. Es documentación
fósil dentro del código.

A esto se suma una incoherencia de numeración que hay que resolver antes de
publicar nada:

| Fuente | Versión declarada |
|---|---|
| `addin/manifest.xml:11` | `0.1.0` |
| `addin/package.json:3` | `0.1.0` |
| `addin/README.md` | «Estado: **v0.2**» |

---

## 3. El manifiesto

El manifiesto (`addin/manifest.xml`) es el único archivo que TI mira. Son
noventa y ocho líneas de XML que le dicen a Word tres cosas: dónde está la
página del panel, qué permisos pide y qué botón poner en la cinta.

| Elemento | Valor literal | Qué significa |
|---|---|---|
| Formato | `OfficeApp`, `xsi:type="TaskPaneApp"` | Manifiesto XML clásico (no el unificado en JSON). «TaskPane» = panel lateral. |
| `<Id>` | `d3a9b1c2-7e4f-4a1b-9c8d-1f2e3a4b5c6d` | Identificador único del complemento. **Es un marcador de posición** (ver abajo). |
| `<Version>` | `0.1.0` | Se debe subir cada vez que el manifiesto cambie. |
| `<ProviderName>` | `Collective Mining - Finanzas` | Quién publica. |
| `<DefaultLocale>` | `es-CO` | Idioma de los textos por defecto. |
| `<DisplayName>` | `Estados Financieros - Actualizar tablas` | Nombre visible. Sin tildes, presumiblemente por prudencia ante recodificaciones. |
| `<Description>` | `Actualiza las tablas del Estado de Situacion Financiera en este mismo documento a partir del Excel modelo.` | Texto del catálogo. |
| `<Hosts>` | `<Host Name="Document"/>` | `Document` significa Word. **No hay host de Excel.** |
| `<Requirements>` | `<Set Name="WordApi" MinVersion="1.3"/>`, `DefaultMinVersion="1.3"` | Conjunto de requisitos: la versión mínima de la API de Word que el equipo debe soportar. |
| `<Permissions>` | `ReadWriteDocument` | El nivel más alto: leer y escribir **el documento abierto**. Nada más. |
| `<SourceLocation>` | `https://localhost:3000/taskpane.html` | La página del panel. |
| `<IconUrl>` | `https://localhost:3000/assets/icon-32.png` | Icono del catálogo. |
| `<HighResolutionIconUrl>` | `https://localhost:3000/assets/icon-80.png` | Variante grande. |
| `<SupportUrl>` | `https://localhost:3000/` | Enlace de soporte. |

Dos observaciones de peso. La primera: `addin/manifest.xml:9` lleva un
comentario del autor —`Genere un GUID propio antes de desplegar: PowerShell
[guid]::NewGuid()`— que confirma que el identificador es de ejemplo. La
segunda: **todas las URL apuntan a `https://localhost:3000`**, la dirección
del servidor de desarrollo. Tal como está, el manifiesto solo funciona en el
equipo de quien lo compila.

El complemento es **mono-anfitrión**: solo Word. Pese al nombre del
repositorio, el Excel no se toca como aplicación, sino como archivo (§7).

El conjunto de requisitos `WordApi 1.3` es coherente con lo que el código
usa. `Table.addRows`, `TableRow.delete()`, `TableRow.font` y
`ContentControl.tables` son de la versión 1.3; el resto —los controles de
contenido, `getByTag`, `cannotEdit`, `insertContentControl`, `appearance`—
son de la 1.1, cubiertas por la 1.3.

### El botón de la cinta

El bloque `<VersionOverrides>` define un único botón, alojado en la pestaña
**Inicio** de Word (`OfficeTab id="TabHome"`), no en una pestaña propia:

| Nodo | Identificador |
|---|---|
| Grupo | `FS.Group`, etiqueta `FS.Group.Label` |
| Botón | `FS.OpenTaskpane`, etiqueta `FS.Button.Label` |
| Acción | `ShowTaskpane` → `<TaskpaneId>FS.Taskpane</TaskpaneId>` |
| Iconos | `Icon.16`, `Icon.32`, `Icon.80` |

Los textos visibles, literales:

| Recurso | Texto |
|---|---|
| `FS.Group.Label` | `Estados Financieros` |
| `FS.Button.Label` | `Actualizar tablas` |
| `FS.Button.Tooltip` | `Trae las cifras del Excel modelo a las tablas de este documento, sin generar un archivo nuevo.` |
| `GetStarted.Title` | `Actualizar tablas del Estado Financiero` |
| `GetStarted.Description` | `Abra el panel para traer las cifras del Excel modelo a las tablas de este documento.` |

No hay comandos sin interfaz (`ExecuteFunction`), ni `FunctionFile`, ni
`MobileFormFactor`, ni `<AppDomains>` —esto último es correcto mientras todo
cuelgue del mismo origen—. Un solo botón que abre un solo panel.

---

## 4. La compilación prevista

El código está en TypeScript, un dialecto de JavaScript con tipos que el
navegador no entiende: hay que traducirlo. Esa traducción es la única razón
por la que hace falta Node.js, y `docs/DESPLIEGUE_ADDIN.md` lo explica bien:
**Node.js solo sirve para compilar**, igual que un compilador de C. Convierte
el TypeScript de `addin/src/` en JavaScript estático y después no vuelve a
aparecer en ningún sitio. Las usuarias no instalan nada.

### `package.json` — órdenes y dependencias

| Orden npm | Comando exacto | Para qué |
|---|---|---|
| `build` | `webpack --mode production` | Genera `dist/` optimizado. |
| `build:dev` | `webpack --mode development` | Igual, sin optimizar y con mapas de origen. |
| `dev-server` | `webpack serve --mode development` | Servidor local con recarga. |
| `start` | `office-addin-debugging start manifest.xml` | Compila, levanta HTTPS y abre Word con el complemento cargado. |
| `stop` | `office-addin-debugging stop manifest.xml` | Lo retira. |
| `validate` | `office-addin-manifest validate manifest.xml` | Comprueba el XML contra el esquema. |
| `lint` | `tsc --noEmit` | **No es un linter**: es una comprobación de tipos. |

Dependencia de ejecución, una sola: `exceljs` `^4.4.0`. Office.js no se
empaqueta, se carga desde la red de distribución de Microsoft (§10).

Dependencias de desarrollo: `@types/office-js` `^1.0.377`,
`copy-webpack-plugin` `^12.0.2`, `html-webpack-plugin` `^5.6.0`,
`office-addin-debugging` `^6.0.0`, `office-addin-dev-certs` `^1.13.0`,
`office-addin-manifest` `^1.13.0`, `ts-loader` `^9.5.1`, `typescript`
`^5.5.4`, `webpack` `^5.94.0`, `webpack-cli` `^5.1.4` y `webpack-dev-server`
`^5.0.4`. No hay campo `engines` que fije la versión de Node, pese a que el
README exige la 18 o superior; no hay orden `test`.

### `webpack.config.js`

Exporta una función asíncrona porque debe esperar al certificado de
desarrollo (`devCerts.getHttpsServerOptions()`). Un solo punto de entrada,
`taskpane: "./src/taskpane/taskpane.ts"`, salida en `dist/[name].js` con
`clean: true`. Tres reglas de módulo: `ts-loader` para `.ts`, la pareja
`style-loader` + `css-loader` para `.css` (`addin/webpack.config.js:25`), y
`asset/resource` para imágenes. Dos complementos: `HtmlWebpackPlugin`, que
produce `dist/taskpane.html` a partir de la plantilla e inyecta la etiqueta
`<script>`, y `CopyWebpackPlugin`, que copia `assets/` con
`noErrorOnMissing: true`. El servidor de desarrollo sirve `dist/` por HTTPS
en el puerto 3000 con `Access-Control-Allow-Origin: *`.

### `tsconfig.json`

`target: "ES2020"`, `module: "ESNext"`, `moduleResolution: "bundler"`
—que exige TypeScript 5 o superior—, `lib: ["ES2020", "DOM"]`,
`strict: true` y `noImplicitAny: true`, `esModuleInterop: true`,
`skipLibCheck: true`, `forceConsistentCasingInFileNames: true`,
`types: ["office-js"]` y `outDir: "dist"`, con `include: ["src/**/*.ts"]`.

Dos consecuencias prácticas. `types: ["office-js"]` es lo que permite
escribir `Word.run` u `Office.onReady` sin importarlos: son globales
declarados. Y el modo estricto es lo que fuerza las aserciones de no nulidad
del panel (`document.getElementById("resumen")!`). El `outDir` es inerte: la
emisión la controla `ts-loader` dentro de webpack.

### El flujo

```
src/taskpane/taskpane.ts   --(ts-loader)---------->  dist/taskpane.js
src/taskpane/taskpane.html --(HtmlWebpackPlugin)-->  dist/taskpane.html
assets/*                   --(CopyWebpackPlugin)-->  dist/assets/*
office.js                  <-- se carga del CDN, no se empaqueta
```

`dist/` es todo el complemento: archivos estáticos que se suben a SharePoint.

---

## 5. El modelo de datos

`addin/src/core/tipos.ts` declara el mismo modelo que el núcleo de Python.
Su cabecera lo dice: *«Modelo de tipos de fila e inferencia. Es el mismo
criterio que usa generador_fs.py (Python), portado a TypeScript para el
add-in.»*

```ts
export type Tipo = "H" | "I" | "S" | "T" | "N" | "X";
```

`H` encabezado de sección, `I` línea de detalle, `S` subtotal, `T` total,
`N` nota de texto, `X` fila de control o cuadre (se descarta).

### `interface Linea` (`addin/src/core/tipos.ts:8`)

| Campo | Tipo | Notas |
|---|---|---|
| `fila` | `number` | Fila del Excel, base 1. |
| `tipo` | `Tipo` | |
| `etiqueta` | `string` | |
| `nota` | `string` | |
| `actual` | `string` | Ya formateado en estilo contable. |
| `previo` | `string` | |
| `actualRaw?` | `unknown` | Sin formatear, para calcular variaciones. |
| `previoRaw?` | `unknown` | |
| `clave?` | `string` | Clave venida de un rango con nombre `fs_<clave>`. |
| `claveOrigen?` | `"rango" \| "etiqueta"` | De dónde salió la clave. |
| `origen` | `"declarado" \| "inferido"` | Obligatorio. |
| `senal` | `string` | Obligatorio; texto humano de la señal que disparó la inferencia. |

### `interface Contexto` (`addin/src/core/tipos.ts:29`)

Campos: `empresa`, `titulo`, `fechaActual`, `fechaPrevia`, `miles`,
`estadoActual`, `estadoPrevio`, `moneda`, `lineas: Linea[]` y un objeto
`meta` anónimo con `hoja`, `comoHoja`, `columnas` (`Record<string,string>`),
`region` (tupla `[number, number]`), `hayColTipo`, `nDeclarados` y
`nInferidos`.

La correspondencia con el contexto de Python es directa, con dos
diferencias de forma y una de fondo:

| `Contexto` (TypeScript) | `ctx` (Python) | Observación |
|---|---|---|
| `fechaActual`, `fechaPrevia`, `estadoActual`, `estadoPrevio` | `fecha_actual`, `fecha_previa`, `estado_actual`, `estado_previo` | Camello frente a guion bajo. `contrato.ts` traduce entre ambos. |
| `Linea.actualRaw` / `previoRaw` | `linea["actual_raw"]` / `["previo_raw"]` | Mismo papel. |
| Interfaz tipada | Diccionario | Python accede con `.get()` y tolera claves ausentes. |
| **No existe** | `ctx["escalares"]` | **Divergencia de fondo.** Ver §6 y §7. |

### Funciones auxiliares

`MARC_CERO` reúne ocho marcadores de «cero visual»: `-`, `–`, `—`, `−`, `ꟷ`,
`‑`, `·` y `0`. Sobre esa base:

| Función | Firma | Qué hace |
|---|---|---|
| `norm` | `(s: unknown) => string` | Recorta, pasa a minúsculas y colapsa espacios. |
| `esNumero` | `(v: unknown) => v is number` | Número que no es `NaN`. |
| `marcadorCero` | `(v: unknown) => boolean` | Texto cuyo recorte está en `MARC_CERO`. |
| `num` | `(v: unknown) => string` | Formato contable: entero con separador de millares, negativos entre paréntesis. |
| `texto` | `(v: unknown) => string` | Convierte a texto quitando paréntesis y espacios de los extremos. |

`inferirTipo(c: Celdas, marcadoresExcluir: string[]): { tipo: Tipo | null;
senal: string }` (`addin/src/core/tipos.ts:86`) es el motor heurístico que
adivina el tipo de fila cuando la hoja no trae columna `Tipo`. Decide en
cascada, en este orden exacto:

| # | Condición | Tipo | Señal literal |
|---|---|---|---|
| 1 | Algún marcador de exclusión aparece en la etiqueta o la nota | `X` | `fila de control/cuadre` |
| 2 | Sin etiqueta, sin valor y sin nota | `null` | `fila vacía` |
| 3 | La etiqueta empieza por `total` | `T` | `empieza por 'Total'` |
| 4 | Sin etiqueta pero con valor | `S` | `cifras sin etiqueta (subtotal)` |
| 5 | Con etiqueta, sin valor, y en mayúsculas **o** terminada en `:` **o** en negrita **o** de 40 caracteres o menos | `H` | `etiqueta sin cifras (encabezado)` |
| 6 | Con etiqueta y sin valor (resto) | `N` | `texto largo sin cifras (nota)` |
| 7 | Con etiqueta y con valor o nota | `I` | `etiqueta con cifras o nota` |
| 8 | Cualquier otro caso | `null` | `sin señales suficientes` |

La regla 5 es muy permisiva: basta con que el texto tenga 40 caracteres o
menos. En la práctica, casi todo texto corto sin cifras se clasifica como
encabezado. (Inferencia: la rama 8 es inalcanzable, porque las reglas 2, 5,
6 y 7 agotan el espacio de casos.)

---

## 6. El contrato compartido

`addin/src/core/contrato.ts` es el corazón del acuerdo entre los dos
motores. Su cabecera no deja lugar a interpretación: *«Es el puerto exacto de
fs_contrato.py. Si cambia uno, cambia el otro.»*

Declara las mismas familias de ancla que `src/fs_contrato.py`:

```
fs-tabla-<nombre>        bloque: una tabla completa        -> se refresca
fs-campo-<nombre>        en línea: campo de encabezado     -> se refresca
fs-dato-<clave>-<campo>  en línea: cifra dentro de prosa   -> se refresca
fs-prosa-<nombre>        bloque: redacción libre           -> NO se toca
fs-registro              bloque: bitácora                  -> se antepone
fs-meta                  bloque oculto: foto JSON          -> se sobrescribe
```

### Paridad función por función

| Elemento | `addin/src/core/contrato.ts` | `src/fs_contrato.py` | Paridad |
|---|---|---|---|
| `PREFIJO`, `FAM_*`, `TAG_REGISTRO`, `TAG_META`, `TABLA_PRINCIPAL` | Idénticos | Idénticos | **Total** |
| `CAMPOS_ENCABEZADO` (8 nombres, mismo orden) | `as const` | Tupla | **Total** |
| `CAMPOS_DATO` (`actual`, `previo`, `nota`, `var_abs`, `var_pct`) | Idénticos | Idénticos | **Total** |
| `tagTabla` / `tagCampo` / `tagDato` / `tagProsa` | Idénticos | `tag_tabla`… | **Total.** TS restringe el campo al tipo `CampoDato`; Python acepta cualquier texto. |
| `descomponer` | Devuelve objeto `{familia, nombre, campo}` | Devuelve tupla `(familia, nombre, campo)` | Semántica idéntica, forma distinta |
| `esRegionDeDatos` | Idéntico, **nunca se llama** | `es_region_de_datos` | Total, pero código muerto en el complemento |
| `clave()` | Regex `[̀-ͯ]` | `unicodedata.combining()` | Equivalente para latín acentuado |
| `claveDeLinea` | Idéntico | `clave_de_linea` | **Total** |
| `construirValores` — campos de encabezado | Idéntico | Idéntico | **Total** |
| `construirValores` — `var_pct` | `.toFixed(1)`, sin separador de millares | `f"{…:,.1f}%"`, con separador | **Divergencia** |
| `construirValores` — redondeo | `Math.round`, medio hacia arriba | `:,.0f`, medio al par | **Divergencia menor** |
| `construirValores` — `escalares` | **Ausente** | Presente (`src/fs_contrato.py:192`) | **Divergencia** |
| `catalogo` | Cuádruple `[clave, etiqueta, actual, previo]`, **nunca se llama** | Quíntuple con `origen`, más los escalares | **Divergencia** |
| `lineasDeTabla` | Presente | Vive en `fs_documento.py` | Coherente con `CONTRATO.md` §5 |
| `PROSA_SUGERIDA` | **No portada** | `("introduccion","analisis","cierre")` | No aplica: el complemento no construye andamiaje |
| `MAX_TAG = 64` | **No portada** | `src/fs_contrato.py:65` | **Divergencia**: `insertarDato()` no valida la longitud |
| Persistencia de la foto | `settings["fs_snapshot"]` | Región `fs-meta` | **Divergencia de estado** (§9) |

Lo que **no** está portado, en resumen: el bloque de escalares de
`construir_valores`, la columna `origen` y los escalares del `catalogo`, la
constante `PROSA_SUGERIDA`, la constante `MAX_TAG = 64`, el separador de
millares del porcentaje de variación y la escritura de la región `fs-meta`.

### Por qué `clave()` debe coincidir en los dos lenguajes

`clave()` (`addin/src/core/contrato.ts:97`) convierte la etiqueta de una fila
del Excel —`Provisión (neta)`— en la clave estable que forma parte del
ancla: `provision_neta`. Los cinco pasos están documentados en el propio
código y deben ser los mismos que en `src/fs_contrato.py:123`: quitar tildes
mediante normalización NFKD descartando los diacríticos, pasar a minúsculas,
sustituir por `_` todo lo que no sea `[a-z0-9]`, colapsar y recortar los `_`,
y cortar a 40 caracteres recortando de nuevo.

La razón por la que esto es crítico es de circuito. El documento vivo lleva
las anclas escritas por el motor de Python (`python fs_documento.py construir
<doc.docx>`), que las nombró con **su** `clave()`. Al refrescar, el
complemento lee cada `w:tag` del documento y busca ese texto exacto en el
mapa de valores que construyó con **su** `clave()`. Si las dos derivaciones
no producen la misma cadena, la búsqueda falla.

Y falla en silencio. `refrescarDocumento` no lanza excepción ante un ancla
que no encuentra: la anota como huérfana (`addin/src/core/escribir-word.ts:78`)
y deja el texto que hubiera. El resultado sería un documento que parece
actualizado, que informa «Listo», y que conserva las cifras del trimestre
anterior en las regiones afectadas. Es el peor modo de fallo posible en un
estado financiero: silencioso y plausible.

La única divergencia conocida hoy es teórica. `unicodedata.combining()`
cubre todos los bloques de marcas combinantes de Unicode; el regex de
TypeScript solo el bloque básico `U+0300–U+036F`. Para etiquetas contables en
español o inglés son equivalentes.

Una nota sobre `lineasDeTabla` (`addin/src/core/contrato.ts:196`): resuelve
qué líneas alimentan cada tabla. Con el nombre `principal` devuelve todas;
con cualquier otro, recorre las líneas con una máquina de estados de un bit y
se queda con la sección que va desde el encabezado `H` cuya clave coincide
hasta el siguiente `H`, exclusive. Usa `clave(l.etiqueta)` y **no**
`claveDeLinea(l)`: un rango con nombre puesto sobre una fila `H` no sirve
para direccionar una tabla. Es consistente con Python, pero contradice la
regla general de que «el rango manda», y merece documentarse.

---

## 7. La lectura del libro

`addin/src/core/leer-excel.ts` es el módulo más largo del complemento (408
líneas) y el que más conviene entender antes de tocar nada, porque hay una
sorpresa de arquitectura: **no usa la API de Excel de Office.js en absoluto**.
No hay una sola llamada a `Excel.run`. El libro se lee con la biblioteca
`exceljs` sobre el `ArrayBuffer` (el contenido binario en memoria) que
devuelve el campo de selección de archivo del panel. El README lo confirma:
*«Se parsea en el navegador con `exceljs` — **no** sube a ningún servidor, no
usa Graph.»* Para el complemento, Excel no es una aplicación anfitriona sino
un archivo de entrada.

### Ajustes

`AJUSTES_DEFECTO` (`addin/src/core/leer-excel.ts:30`) fija todo el
comportamiento:

| Campo | Valor por defecto |
|---|---|
| `empresa` | `"Collective Mining Ltd."` |
| `hoja` | `"FS"` |
| `hojaMarcadores` | siete cadenas: `situación financiera`, `statement of financial position`, `financial position`, `total assets`, `total liabilities and equity`, `assets`, `liabilities and equity` |
| `primeraFila` | `"auto"` |
| `columnas` | `{}` (vacío: todo se detecta) |
| `marcadoresExcluir` | `["control check", "check", "cuadre", "balance check"]` |
| `maxFilasScan` | `400` |
| `maxColsScan` | `16` |
| `prefijoRangos` | `"fs_"` |

Aquí está la primera simplificación frente a Python: **estos ajustes están
cableados en el código**. El motor de Python los lee de `config.json`. El
complemento no lee configuración alguna y no ofrece interfaz para cambiarla,
pese a que sus mensajes de error remiten a «los ajustes».

### Detección de la hoja

`elegirHoja` (`addin/src/core/leer-excel.ts:112`) trabaja en dos niveles.
Primero busca una hoja cuyo nombre coincida exactamente con `FS`; si la
encuentra, informa `'FS' por nombre exacto`. Si no, puntúa cada hoja: para
cada una concatena el texto normalizado de las primeras 60 filas por 8
columnas y suma un punto por cada marcador de contenido que aparezca, más un
punto adicional si el nombre de la hoja contiene el nombre buscado. Gana la
de mayor puntuación **siempre que alcance 2**. Si ninguna llega, lanza:
`No pude identificar la hoja del Estado de Situación Financiera. Indique el
nombre exacto en los ajustes (hoja).`

### Detección de columnas

`detectarColumnas` (`addin/src/core/leer-excel.ts:145`) procede en fases.
Parte de lo que digan los ajustes (convertidos de letra a número con
`letraACol`). Busca encabezados literales en las ocho primeras filas: `tipo`
o `type` para la columna de tipo, y `note`/`nota`/`notes`/`notas` para la de
nota. Construye después un perfil por columna contando textos y números
**desde la fila 2** —la primera se excluye a propósito—. Las candidatas a
cifra son las columnas con dos números o más; se toman las dos con más
números y se reordenan de izquierda a derecha: la izquierda es `actual`, la
derecha `previo`. La columna de etiqueta, si no se dedujo antes, es la
primera a la izquierda de `actual` con tres textos o más, y si no hay
ninguna, la columna A. La de nota, si falta, se busca entre etiqueta y
actual: la primera columna con al menos un entero entre 1 y 99 y como mucho
dos celdas de texto. Sin columna `actual`, lanza:
`No pude identificar las columnas de cifras. Fíjelas en los ajustes
(columnas).`

### Región y encabezado

`detectarRegion` (`addin/src/core/leer-excel.ts:218`) devuelve la tupla
`[primera, ultima]`. En modo automático arranca con `5` como respaldo y busca
la primera fila con etiqueta textual no vacía, distinta de `$`, que además
cumpla `r > 4` o traiga cifras. Es decir: en las cuatro primeras filas exige
números, para no confundir el bloque de encabezado con la primera línea de
datos. La última es la última fila con etiqueta textual o con número en
alguna de las dos columnas de cifras.

El encabezado se extrae por posición, por encima de la región:

| Campo del contexto | Origen |
|---|---|
| `empresa` | De los ajustes, **no** del Excel |
| `titulo` | Primer texto de la columna de etiquetas por encima de la región |
| `fechaActual` / `fechaPrevia` | Fila 1 de la columna actual / previa |
| `miles` | Fila 2 de la columna actual |
| `estadoActual` / `estadoPrevio` | Fila 3 de la columna actual / previa |
| `moneda` | Fila 4 de la columna actual |

### Claves estables

`leerRangosConNombre` (`addin/src/core/leer-excel.ts:259`) recoge los nombres
definidos del libro que empiezan por `fs_` y los traduce a un mapa
`{fila → clave}`. Acepta los dos formatos habituales —`'Mi Hoja'!$A$16` y
`FS!$A$16:$G$16`—, exige coincidencia exacta con la hoja elegida y conserva
el primer nombre asignado a cada fila. Es lo que permite que renombrar una
fila del Excel no rompa la identidad de la línea.

Tiene una fragilidad conocida: obtiene los nombres de
`(wb as any).definedNames?.model`, que es **API interna de `exceljs`**, no
pública ni tipada. Cualquier cambio menor de la biblioteca puede romperla.

### Qué replica y qué deja fuera

Replica: la detección de hoja por nombre y por contenido, la detección de
columnas por perfil, la región automática, la inferencia de tipo de fila con
las mismas señales, la lectura de rangos `fs_*` y el formato contable con
paréntesis para negativos.

Simplifica u omite:

- **No lee `config.json`.** Los ajustes no son configurables.
- **No expone escalares.** Los rangos `fs_*` que apuntan fuera de la región
  se leen en el mapa y **nunca se consultan**: se pierden en silencio. El
  motor de Python los publica como `fs-dato-<clave>-actual`.
- **No respeta el formato de celda.** Python tiene
  `generador_fs.formatear_valor()`, que honra decimales, porcentajes y fechas
  del libro; aquí todo pasa por `num()`, que redondea a cero decimales.
- **Techo duro de 400 filas por 16 columnas.** Un estado más largo o más
  ancho queda truncado sin aviso.
- No maneja celdas combinadas, hojas ocultas ni varias hojas de estado.

---

## 8. La escritura en Word

`addin/src/core/escribir-word.ts` es el módulo que da sentido al proyecto. Su
cabecera enuncia la filosofía completa:

> *No genera un archivo nuevo y no renderiza el documento entero: recorre los
> controles de contenido del documento, y solo entra en los que pertenecen a
> las familias de datos del contrato (ver CONTRATO.md). La prosa que la
> persona haya escrito no se visita siquiera, así que sobrevive intacta a
> cada actualización — y si la borra, se queda borrada.*
>
> *Los controles de datos van con LockContents = true para que nadie los pise
> a mano. Este módulo los desbloquea justo antes de escribir y los vuelve a
> bloquear después: es la única forma de que Office.js pueda tocarlos.*

Un *control de contenido* es la caja etiquetada que Word sabe delimitar
dentro de un documento; en el vocabulario de este manual, una **región**. Su
etiqueta (`w:tag`) es el **ancla**.

### El candado

El segundo párrafo de esa cabecera describe una restricción de la API que
condiciona todo el diseño. Las regiones de datos llevan **candado**
(`cannotEdit = true`) para que nadie las edite a mano. Pero el candado no
distingue entre una persona y un programa: Office.js tampoco puede escribir
en una región bloqueada. De ahí la coreografía de cinco fases de
`refrescarDocumento` (`addin/src/core/escribir-word.ts:40`), un único
`Word.run`:

```
carga    ccs.load("items/tag,items/id")  +  sync
fase 1   clasificar por familia: tablas / en línea / prosa
fase 2   cc.cannotEdit = false  en todo lo que se va a escribir  +  sync
fase 3   campos de encabezado y cifras sueltas: insertText(valor,"Replace")
fase 4   tablas: lineasDeTabla() -> reescribirTabla()
fase 5   cc.cannotEdit = true  en todo  +  sync
```

Las regiones `fs-prosa-*` solo se cuentan (`informe.prosaIntacta += 1`);
nunca se abren. Todo lo demás —incluidas `fs-registro`, `fs-meta` y los
controles ajenos al contrato— se ignora en esta función.

En la fase 3, un ancla para la que el Excel no da valor se anota como
**huérfana** y se salta, sin tocar el texto existente.

> **Riesgo de robustez.** No hay `try/finally`. Si algo lanza en las fases 3
> o 4, la fase 5 no se ejecuta; y como el `context.sync()` de la fase 2 ya
> confirmó los `cannotEdit = false`, el documento puede quedar con todas sus
> regiones de datos **desbloqueadas**.

### Reescritura de una tabla

`reescribirTabla` (`addin/src/core/escribir-word.ts:109`) carga las tablas del
control; si no hay ninguna, lanza un error que además indica el remedio:
`El control "<tag>" no contiene ninguna tabla. Prepare el documento con:
python fs_documento.py construir <doc.docx>`. Toma la primera tabla, borra
las filas desde el final hacia el índice 1 —**conserva la fila 0**, el
encabezado, y con ella los anchos y el estilo de la plantilla—, y añade al
final las filas nuevas de golpe con `addRows`. Cada línea se convierte en
cuatro celdas según su tipo (`filaValores`, `addin/src/core/escribir-word.ts:21`):

| Tipo | Fila producida |
|---|---|
| `H` | `[etiqueta, "", "", ""]` |
| `N` | `[etiqueta (Nota n), "", "", ""]` |
| `S` | `["", "", actual, previo]` |
| `I`, `T` | `[etiqueta, nota, actual, previo]` |

Después aplica formato. Y aquí está la deuda más visible: **el único formato
que aplica es la negrita** (`row.font.bold = t === "S" || t === "T" || t ===
"H"`). El contrato especifica además sangría para `I`, filete superior para
`S`, negrita con filete superior y doble inferior para `T`, y cursiva para
`N`. El README lo reconoce: un documento construido con Python y refrescado
con el complemento **pierde esos filetes**.

Dos supuestos más: la tabla debe tener exactamente cuatro columnas, y el
patrón de sincronización es costoso —al menos cuatro viajes de ida y vuelta
por tabla, más tres globales—.

### Insertar una cifra suelta

`insertarDato(clave: string, campo: K.CampoDato): Promise<string>`
(`addin/src/core/escribir-word.ts:154`) crea una región en línea en la
posición del cursor: le pone el ancla del contrato, un título legible
`<clave> (<campo>)`, apariencia `BoundingBox`, el marcador de posición `—` y
el candado. Devuelve el ancla creada. Equivale a `fs_documento.py insertar`.
**No está cableada a ningún elemento del panel**: es código funcional pero
inalcanzable.

### `InformeEscritura`

```ts
export interface InformeEscritura {
  tablas: Array<{ nombre: string; filas: number }>;
  campos: number;
  datos: number;
  huerfanos: string[];
  prosaIntacta: number;
}
```

### Lo que este módulo no hace

- **Nunca escribe `fs-meta`.** La constante `TAG_META` se importa y no se usa
  en ninguna parte. La foto de la última actualización va a otro sitio (§9).
- **No aplica ni retira la protección de documento.** El modo estricto de dos
  editores hay que gestionarlo desde Python o desde Word.
- **Descarta las colisiones de clave.** La línea 41 desestructura solo
  `{ valores }` de `construirValores`; dos filas del Excel que produzcan la
  misma clave se pisan sin que nadie lo advierta.

---

## 9. La bitácora y la foto de la última actualización

`addin/src/core/registro.ts` (79 líneas) responde a una pregunta: ¿qué
cambió respecto de la vez anterior? Para contestarla necesita guardar una
foto del estado aplicado. (En el vocabulario de este manual, «foto» designa
también el `.docx` desechable de `generador_fs.py`; aquí se usa en el sentido
del contrato: la instantánea JSON de la última actualización.)

El contrato reserva para eso la región oculta `fs-meta`. **El complemento no
la usa.** Guarda la foto en `Office.context.document.settings`, bajo la clave
`fs_snapshot` (`addin/src/core/registro.ts:19`), que es un almacén de
propiedades que Word conserva dentro del propio archivo. Viaja con el
documento, sí, pero **el motor de Python no lo lee**.

La estructura guardada es deliberadamente pobre: por cada línea, `etiqueta`,
`actual`, `previo` y `tipo`, más la fecha en formato ISO. No guarda ni la
`clave` ni el número de `fila`.

`guardarSnapshot` envuelve la API de retrollamadas de Office en una promesa y
espera a `settings.saveAsync`. La persistencia solo cuaja de verdad cuando la
usuaria guarda el `.docx`: de ahí el `Guarda el documento (Ctrl+S).` con que
termina el mensaje de éxito del panel.

`calcularCambios(previo, ctx)` (`addin/src/core/registro.ts:50`) compara. Si
no hay foto previa, devuelve una sola frase: `Primera actualización: no hay
versión anterior con la que comparar.` Si la hay, construye un mapa de la
foto anterior y para cada línea nueva emite `Nueva fila: …` si no había
correspondencia, o `<clave>: <antes> → <ahora> (comparativo …)` si cambió
alguna cifra; lo que quede en el mapa sale como `Fila retirada: …`. Sin
diferencias, devuelve `Sin cambios en las cifras respecto de la última
actualización.`

> **Incoherencia arquitectónica.** El diff se indexa por **texto de
> etiqueta**, con una función local `(l) => l.etiqueta || "(sin etiqueta " +
> l.tipo + ")"`, no por `claveDeLinea()`. Todo el §4 del contrato existe
> precisamente para que renombrar una fila no rompa la identidad; aquí sí la
> rompe: cambiar «Cash and cash equivalents» por «Efectivo y equivalentes»
> produce un falso par «Fila retirada» más «Nueva fila». Y todos los
> subtotales sin etiqueta colapsan en la misma entrada `(sin etiqueta S)`, de
> modo que solo el último se compara.

La bitácora se escribe desde el otro módulo. `escribirRegistro(cambios,
origen)` (`addin/src/core/escribir-word.ts:170`) busca la región
`fs-registro`; si no existe, devuelve `false` sin lanzar. Si existe, la
desbloquea, **antepone** —no reemplaza— un bloque con la fecha localizada en
`es-CO`, el nombre del archivo de origen y hasta **cuarenta** cambios
(`slice(0, 40)`, sin indicar cuántos se omitieron), y la vuelve a bloquear.

---

## 10. El panel

### La interfaz

`addin/src/taskpane/taskpane.html` son 39 líneas: un título, un subtítulo y
una lista ordenada de tres pasos. Los textos, literales:

| Elemento | Texto exacto |
|---|---|
| `<h1>` | `Estado de Situación Financiera` |
| Subtítulo | `Trae las cifras del Excel modelo a las tablas de este documento. No genera un archivo nuevo.` |
| Paso 1 | `1 · Elige el Excel modelo (de tu OneDrive)` |
| Paso 2 | `2 · Revisa lo que se va a trasladar` |
| Estado inicial del resumen | `Aún no has elegido un Excel.` |
| Plegable | `Ver líneas y cambios` |
| Paso 3 | `3 · Aplícalo al documento` |
| Botón | `Actualizar tablas de este documento` |
| Pie | `Deja una entrada de bitácora en la sección «Registro de actualizaciones» del documento (control con etiqueta fs-registro, si existe).` |

Los identificadores que unen el HTML con el TypeScript son `archivo` (campo
de archivo, `accept=".xlsx"`), `resumen`, `detalle`, `cambios`,
`tabla-preview`, `aplicar` (deshabilitado de inicio) y `estado`. Las
cabeceras de la vista previa son `#`, `Tipo`, `Concepto`, `Nota`, `Actual`,
`Comp.` y `Origen`. Office.js se carga al final del cuerpo desde
`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`.

La hoja de estilos define una paleta sobria de ocho variables —verde `#2f6b4f`
para la acción, ámbar `#9a6512` para avisos, rojo `#9c382c` para errores— y
un detalle que vale por sí solo: `tr.inferido td` pinta de ámbar las filas
cuyo tipo se **infirió**, que son justo las que hay que revisar. El panel es
siempre claro: no contempla tema oscuro.

### El flujo

`Office.onReady` (`addin/src/taskpane/taskpane.ts:21`) comprueba primero que
el anfitrión sea Word y, si no lo es, sustituye la página por el mensaje
`Este complemento solo funciona en Word.` Después engancha dos oyentes.

`onArchivo` lee el archivo elegido a un `ArrayBuffer`, llama a
`leerContexto`, guarda el contexto y el nombre del origen, lee la foto
anterior, calcula los cambios y pinta el resumen: hoja y cómo se identificó,
letras de las cinco columnas, región de filas y recuento de líneas
declaradas frente a inferidas. Si la hoja no traía columna `Tipo`, añade en
ámbar: `La hoja no tiene columna 'Tipo': todos los tipos se infirieron.
Revisa la vista previa.` Solo entonces habilita el botón. Ante un error,
muestra `No se pudo leer: <mensaje>`.

`onAplicar` (`addin/src/taskpane/taskpane.ts:89`) ejecuta la secuencia
completa y compone el mensaje final: `Listo: N filas en M tabla(s), C campos
y D cifras en la redacción. P zona(s) de prosa sin tocar.`, más `Bitácora
añadida.` si la hubo, más `AVISO: n ancla(s) sin dato en el Excel (…)` si hay
huérfanas —lo que tiñe el estado de ámbar—, y siempre `Guarda el documento
(Ctrl+S).`

### Grafo de llamadas

```
Office.onReady
  ├─ guarda de anfitrión (Office.HostType.Word)
  ├─ #archivo .change  →  onArchivo
  │     ├─ File.arrayBuffer()
  │     ├─ leer-excel.leerContexto(buf, AJUSTES_DEFECTO)
  │     │     └─ elegirHoja → materializar → detectarColumnas
  │     │        → detectarRegion → leerRangosConNombre
  │     │        → [tipos.inferirTipo, tipos.num, tipos.texto]
  │     ├─ registro.leerSnapshot()
  │     ├─ registro.calcularCambios()
  │     └─ pintarDetalle() → escape()
  └─ #aplicar .click   →  onAplicar
        ├─ escribir-word.refrescarDocumento(ctx)
        │     ├─ contrato.construirValores(ctx)
        │     ├─ contrato.descomponer(cc.tag)
        │     ├─ contrato.lineasDeTabla(nombre, ctx)
        │     └─ reescribirTabla() → filaValores()
        ├─ registro.leerSnapshot() + calcularCambios()
        ├─ escribir-word.escribirRegistro(cambios, nombreOrigen)
        └─ registro.guardarSnapshot(ctx)

NO ALCANZABLES desde el panel:
  contrato.catalogo()        contrato.esRegionDeDatos()
  contrato.tagTabla()        contrato.tagProsa()
  escribir-word.insertarDato()
```

`calcularCambios` se ejecuta dos veces con la misma foto: una para la vista
previa y otra para la bitácora. Es correcto —`guardarSnapshot` va después—
pero duplica trabajo.

---

## 11. Qué falta para poder usarlo

En orden de ejecución.

1. **Compilar.** `npm install` y `npm run build` en un equipo con Node.js 18
   o superior. Es el paso que nunca se ha dado.
2. **Crear los tres iconos.** `assets/icon-16.png`, `icon-32.png` e
   `icon-80.png`, cuadrados. Sin ellos el complemento carga, pero el botón
   sale sin icono.
3. **Publicar los archivos en HTTPS** dentro del inquilino: una biblioteca de
   SharePoint, un IIS interno o Azure Static Web Apps. Son archivos
   estáticos; no hay puertos de entrada ni servidor de aplicación.
4. **Ajustar el manifiesto.** Generar un GUID propio para `<Id>`, sustituir
   **todas** las URL de `localhost:3000` —`SourceLocation`, `IconUrl`,
   `HighResolutionIconUrl`, `SupportUrl`, `<bt:Images>` y `<bt:Urls>`— y
   unificar el número de versión con `package.json` y el README.
5. **Publicar el manifiesto.** Centro de administración de Microsoft 365 →
   Configuración → Aplicaciones integradas → Cargar aplicación personalizada,
   asignado solo a las dos usuarias. Aparece en un plazo de entre minutos y
   24 horas. Para probar antes, cada usuaria puede hacer carga lateral desde
   Word → Inicio → Complementos → Mis complementos → Cargar mi complemento.
6. **Permisos de Graph, solo si se cambia la forma de leer el Excel.** Hoy
   **no hacen falta**: la usuaria elige el archivo a mano y se parsea en su
   equipo. Si más adelante se lee el libro directamente de SharePoint u
   OneDrive, hará falta consentimiento de administrador para `Files.Read`,
   acotado al sitio correspondiente.
7. **Revisión de TI y prueba real** contra un documento vivo construido con
   `python fs_documento.py construir`, en Word de escritorio y en Word web.

### Obstáculos de compilación previsibles

Son deducciones sobre el primer `npm install && npm run build`, no defectos
observados (inferencia: ninguno se ha podido verificar, porque nunca se ha
compilado).

1. **`exceljs` en Webpack 5 sin polirrellenos.** `webpack.config.js` no
   define `resolve.fallback`. `exceljs@4` arrastra dependencias que esperan
   `stream`, `buffer`, `crypto`, `zlib` y `fs`, y Webpack 5 dejó de
   inyectarlos automáticamente. Es el riesgo número uno.
2. **La hoja de estilos nunca llega a `dist/`.** `taskpane.ts` no importa
   `./taskpane.css`, `CopyWebpackPlugin` solo copia `assets/`, y
   `HtmlWebpackPlugin` no reescribe el `<link>` de la plantilla. Resultado
   esperado: el panel se renderiza sin estilos. Se corrige con una línea
   (`import "./taskpane.css";`).
3. **`style-loader` y `css-loader` no están en `devDependencies`** pese a
   figurar en `webpack.config.js:25`.
4. **`import ExcelJS from "exceljs"`** puede requerir la forma
   `import * as ExcelJS` según cómo resuelva `moduleResolution: "bundler"`.
5. **`(wb as any).definedNames.model`** es API interna de `exceljs`.

### Defectos lógicos e incoherencias detectados

| # | Dónde | Qué |
|---|---|---|
| A | `escribir-word.ts:70–99` | Sin `try/finally`: un fallo intermedio deja las regiones de datos **desbloqueadas**. |
| B | `registro.ts:53` | El diff se indexa por texto de etiqueta, no por `claveDeLinea()`. Renombrar una fila genera falsos «retirada» + «nueva»; los subtotales sin etiqueta colapsan en una sola entrada. |
| C | `escribir-word.ts` | `fs-meta` nunca se escribe. Los dos motores **no comparten estado**: refrescar con Python y luego con el complemento produce bitácoras incoherentes. |
| D | `escribir-word.ts:41` | Las colisiones de clave se descartan; dos filas homónimas se pisan en silencio. |
| E | `taskpane.ts:8–11` | Mapa de anclas obsoleto en el comentario de cabecera. |
| F | `leer-excel.ts` | Los rangos `fs_*` fuera de la región (los escalares del contrato) se leen y se descartan sin aviso. |
| G | `escribir-word.ts:185` | Tope silencioso de 40 cambios por entrada de bitácora. |
| H | `leer-excel.ts:45–46` | Techo duro de 400 filas por 16 columnas. |
| I | `contrato.ts` | `MAX_TAG = 64` no portado: `insertarDato()` no valida la longitud del ancla. |
| J | `manifest.xml`, `package.json`, `README.md` | Tres versiones declaradas, dos valores distintos (`0.1.0` frente a «v0.2»). |
| K | `leer-excel.ts:139`, `:213`, `:376` | Los mensajes de error piden cambiar «los ajustes», que la usuaria no puede tocar desde el panel. |
| L | `addin/README.md`, `contrato.ts:5` | Ambos apuntan a `CONTRATO.md` «en la raíz del repositorio»; el archivo real está en `docs/CONTRATO.md`. |

---

## 12. Si algún día se retoma

Las recomendaciones que siguen se apoyan en lo que ya dice
`docs/DIRECCION.md` y en lo que el propio código pide.

**El contrato como fuente única.** Hoy hay dos implementaciones de la misma
especificación y una regla escrita —*«Si cambia uno, cambia el otro»*— que
depende de que alguien se acuerde. Conviene reforzarla con algo mecánico: un
único archivo de datos (por ejemplo, un JSON con `CAMPOS_ENCABEZADO`,
`CAMPOS_DATO` y los prefijos) del que ambos lenguajes lean, o como mínimo una
prueba que compare las constantes de los dos lados y falle si divergen.

**Pruebas de «archivo dorado».** `docs/DIRECCION.md` ya las propone y las
califica de «afirmación medible»: un `.xlsx` de referencia y un `.docx`
esperado, y una comprobación automática que compara el texto resultante. Para
la paridad entre motores hay una prueba más barata y más urgente: un juego de
etiquetas problemáticas —con tildes, con paréntesis, con guiones, de más de
40 caracteres— sobre el que `clave()` de Python y `clave()` de TypeScript
deban producir exactamente la misma cadena. Es la única garantía real de que
las anclas del documento seguirán encontrándose (§6).

**Unificar el estado persistido.** Escribir `fs-meta` desde el complemento,
en lugar de —o además de— `settings["fs_snapshot"]`, para que las dos vías
compartan la misma foto y las bitácoras sean coherentes.

**Indexar el diff por `claveDeLinea()`.** Es un cambio de tres líneas en
`registro.ts` que alinea la bitácora con el resto del sistema.

**Igualar el formato de tabla.** Sangrías, filete superior de subtotal,
filete superior y doble inferior de total, cursiva de nota. El README lo
llama «el siguiente trabajo pendiente».

**Externalizar la configuración.** `docs/DIRECCION.md` lo recomienda para el
frente B —`config.json` con hoja, mapa de columnas y empresa— y vale igual
aquí: mientras los ajustes estén cableados, los mensajes de error del propio
complemento piden algo imposible.

**Cerrar el panel.** Cablear `catalogo()` e `insertarDato()`, que ya están
escritos, y mostrar las colisiones de clave.

**Un dueño de mantenimiento nombrado.** Es una de las cinco decisiones
abiertas de `docs/DIRECCION.md` («¿Quién es el dueño de mantenimiento de
plantillas y del add-in a un año?»). Un complemento publicado en el centro de
administración es infraestructura: necesita alguien que responda por él.

Y una última consideración de gobierno, que también está escrita en
`docs/DESPLIEGUE_ADDIN.md`: el camino de línea de órdenes no depende del
complemento y hace lo mismo. Si el frente C se atasca en revisión de TI, o se
cae, el trabajo no se detiene.

---

## Resumen del capítulo

- El complemento es el frente C del plan: existe porque la organización
  bloquea instaladores, no porque el motor de Python fallara.
- Se despliega desde el centro de administración de Microsoft 365 y no
  instala nada en el equipo de las usuarias; es una página web que Word carga
  en un panel.
- Está **escrito y sin compilar**: 1.158 líneas de TypeScript sin
  `node_modules/`, sin `dist/`, sin `package-lock.json` y sin iconos.
- El manifiesto es funcional pero apunta entero a `https://localhost:3000` y
  lleva un GUID de ejemplo; hay tres versiones declaradas y dos valores.
- `contrato.ts` es un puerto casi exacto de `fs_contrato.py`; lo que falta
  son los escalares, `MAX_TAG`, la columna `origen` del catálogo y la
  escritura de `fs-meta`.
- Si `clave()` divergiera entre los dos lenguajes, el refresco no fallaría:
  dejaría las cifras viejas y reportaría anclas huérfanas.
- El refresco desbloquea las regiones, escribe y vuelve a bloquear, pero sin
  `try/finally`: un fallo intermedio puede dejarlas abiertas.
- Para usarlo hacen falta, por este orden: compilar, iconos, alojamiento
  HTTPS, manifiesto ajustado, publicación centralizada y una prueba real
  contra un documento vivo. Graph solo si se cambia la forma de leer el
  libro.
