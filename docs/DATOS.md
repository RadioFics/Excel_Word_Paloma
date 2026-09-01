# Dónde se editan las cifras

La duda más razonable del proyecto: *si el Word está bloqueado, ¿cómo cambio
un número?*

---

## La regla, en una frase

> **Las cifras se editan en Excel. Nunca en Word.**

El Word no es el sitio donde viven los datos: es donde se *muestran*. Por eso
sus tablas y sus cifras están bloqueadas — no para fastidiar, sino para que
nadie escriba a mano un número que el siguiente refresco va a machacar.

```
   Excel  ──────────────►  Word
   (aquí se editan)        (aquí se leen)

   Un solo sentido. Nunca al revés.
```

## Por qué el bloqueo es lo que quieres

Sin él, esto pasa a la primera:

1. Alguien corrige «72.957.812» a mano en el Word porque vio un error.
2. El Excel sigue diciendo 72.957.812.
3. Al siguiente refresco, el Word vuelve a poner 72.957.812.
4. La corrección desaparece, y nadie sabe cuándo ni por qué.

Con el bloqueo, el paso 1 no ocurre: Word no deja escribir ahí. La persona va
al Excel, que es donde el cambio queda registrado y donde cuadran las
fórmulas.

---

## Cómo cambiar una cifra, en la práctica

1. **Abra el Excel** y cambie el valor donde corresponda.
   Si es una celda con fórmula, cambie lo que alimenta la fórmula, no el
   resultado.
2. **Guarde y cierre el Excel.**
3. **Cierre el Word**, si lo tiene abierto.
4. **Arrastre el Excel** sobre `refrescar.bat` (o elija la opción 1 en
   `estados_financieros.bat`).
5. El documento queda al día. Qué cambió se anota en un **.log aparte**
   (`salidas\bitacora_<documento>.log`), no dentro del Word:
   ```
   2026-09-01 14:13   MI_DOCUMENTO.docx
     origen: MI_LIBRO.xlsx (sha 95d5f53b20ac)
     cambios:
       - Cash and cash equivalents: 72,957,812 -> 99,999,999
   ```

### Comprobar antes de aplicar

Para ver qué se va a mover sin tocar el documento:

```bash
python src\fs_documento.py catalogo
```

Lista todas las cifras disponibles y su valor actual.

---

## Dos archivos, dos papeles

| | El Excel | El Word base (OneDrive) |
|---|---|---|
| **Qué contiene** | los números y sus fórmulas | la redacción, y las cifras *reflejadas* |
| **Quién lo edita** | quien lleva la contabilidad | quien redacta |
| **Qué se edita** | cualquier celda | solo el texto |
| **Si se toca lo otro** | — | se pierde en el siguiente refresco |

---

## ¿Y los documentos de `salidas\`?

Los que crea `generar.bat` son **fotos desechables**. Sirven para enviar algo
puntual.

**No los edite esperando conservar los cambios.** No están vinculados a nada:
la siguiente corrida crea otro archivo con otro nombre y el anterior se queda
donde está, congelado. Si quiere que su texto sobreviva, trabaje en el
documento base, no en una salida.

| | `salidas\...docx` | Documento base |
|---|---|---|
| Se crea | uno nuevo cada vez | una sola vez |
| Su redacción | se queda en ese archivo | pasa a la siguiente versión |
| Las cifras | congeladas | se refrescan |
| Para qué sirve | una entrega puntual | el documento de trabajo |

---

## Excepción: texto que sí se escribe en Word

Todo lo que **no** sea una cifra del Excel:

- párrafos de análisis y comentarios
- explicaciones de una variación
- notas al pie, encabezados propios, anexos

Eso vive **solo** en el Word y no está en ningún otro sitio. El refresco no lo
toca. Es exactamente el trabajo que el sistema existe para proteger.

La frontera es limpia: **si el número sale del Excel, se edita en Excel. Si es
prosa, se escribe en Word.**

---

## Intercalar una cifra en medio de un párrafo

Para escribir *«los activos totales ascendieron a **119.066.301**»* y que ese
número siga al Excel:

```bash
python src\fs_documento.py insertar "MI_DOCUMENTO.docx" total_assets actual
```

Queda un recuadro bloqueado dentro de la frase. Usted escribe alrededor; la
cifra la mantiene el Excel. Campos disponibles: `actual`, `previo`, `nota`,
`var_abs` (la diferencia) y `var_pct` (la variación porcentual).

---

---

## Paso a paso: una cifra de la tabla dentro de un párrafo

El caso concreto: quiere escribir

> «Al cierre del periodo los activos totales ascendieron a **169.266.900**,
> con un incremento notable frente al ejercicio anterior.»

…y que ese número lo mantenga el Excel, **se lea como texto corrido** (nada
de cuadros flotantes) y **nadie lo pueda cambiar a mano**.

### 1 · Averigüe la clave de esa cifra

```bash
EstadosFinancieros.exe MI_LIBRO.xlsx --estado
```

o, con más detalle:

```bash
python src\fs_documento.py catalogo "MI_LIBRO.xlsx"
```

Verá una lista así. La primera columna es la **clave**:

```
 clave                        origen      actual
 ---------------------------------------------------
 cash_and_cash_equivalents    rango      93,662,415
 total_assets                 rango     169,266,900
 total_liabilities            rango      40,169,722
```

### 2 · Cierre el documento en Word

Si está abierto, la orden se detiene y le dice quién lo retiene. No es
opcional: escribir sobre un `.docx` abierto lo deja inservible.

### 3 · Coloque la cifra

```bash
python src\fs_documento.py insertar "MI_DOCUMENTO.docx" total_assets actual --antes "Al cierre del periodo los activos totales ascendieron a " --despues ", con un incremento notable."
```

`--antes` y `--despues` son opcionales: si los omite, deja solo la cifra y
usted escribe alrededor en Word.

Campos que puede pedir: `actual`, `previo`, `nota`, `var_abs` (la
diferencia) y `var_pct` (la variación porcentual).

### 4 · Refresque para que tome el valor

```bash
EstadosFinancieros.exe MI_LIBRO.xlsx --refrescar
```

### 5 · Quite el recuadro para que se lea como texto normal

```bash
python src\fs_documento.py apariencia "MI_DOCUMENTO.docx" invisible
```

Y ya está. El párrafo se lee como cualquier otro:

> Al cierre del periodo los activos totales ascendieron a 169.266.900, con
> un incremento notable.

Sin recuadro, sin sombreado, sin cuadro flotante. Pero:

- **la cifra sigue bloqueada** — si alguien intenta teclear encima, Word no
  le deja
- **sigue vinculada** — el próximo refresco la pone al día sola

Para que valga por defecto en todos los documentos, ponga en `config.json`:

```json
"apariencia_datos": "hidden"
```

Para volver a ver los recuadros (útil mientras monta el documento, porque
enseña de un vistazo qué mantiene el Excel):

```bash
python src\fs_documento.py apariencia "MI_DOCUMENTO.docx" visible
```

### Hacerlo desde Word, sin consola

También se puede, si prefiere elegir el sitio exacto con el cursor:

1. Pestaña **Programador** (si no está: *Archivo → Opciones → Personalizar
   cinta → marcar Programador*).
2. Ponga el cursor donde va la cifra, dentro de su frase.
3. **Control de contenido de texto sin formato** (el icono `Aa`).
4. Con el control seleccionado, **Propiedades**:
   - **Etiqueta** → `fs-dato-total_assets-actual`
   - marque **«No se puede editar el contenido»**
   - en **Aparición** elija **Ninguno** (así no dibuja recuadro)
5. Guarde, cierre Word y refresque.

La *Etiqueta* es lo único que importa: es el nombre por el que el programa
la reconoce. El formato del texto (negrita, tamaño, color) es el que usted
le dé — el refresco cambia el número, no su aspecto.

### Por qué no son cuadros de texto

Un **cuadro de texto** de Word es un objeto flotante: se mueve, se
superpone, rompe el flujo del párrafo. Esto no es eso.

Es un **control de contenido en línea**: por dentro es una marca invisible
alrededor de unos caracteres. Para Word, esos caracteres son parte del
párrafo como cualquier otro — se justifican, se parten de línea, heredan el
estilo, salen en el índice y se imprimen igual. La marca solo sirve para que
el programa sepa dónde está esa cifra.

Por eso puede darle negrita, cambiarle la fuente o meterla en una viñeta: el
formato es suyo, el valor es del Excel.

## Llevar al Word cifras que NO son filas de la tabla

Ratios, tipos de cambio, fechas de corte, un importe de otra hoja. Todo lo
que quiera citar en la redacción sin que forme parte del estado.

**No hace falta tocar el programa.** Se nombra la celda en Excel y se coloca
el hueco en Word. Dos gestos, y ya queda conectado para siempre.

### Paso 1 — Nombre la celda en Excel

Seleccione la celda —**de cualquier hoja del libro**— y escriba un nombre
que empiece por `fs_` en el **Cuadro de nombres** (arriba a la izquierda,
donde pone `A1`):

```
   fs_ratio_corriente
   fs_margen_bruto
   fs_tipo_cambio_cop
   fs_capex_comprometido
```

Pulse Enter y guarde el libro. El nombre queda pegado a esa celda: aunque
inserte filas encima o mueva la tabla, Excel reajusta la referencia sola.

### Paso 2 — Compruebe que el programa la ve

```bash
python src\fs_documento.py catalogo
```

```
 clave                        origen      actual
 ---------------------------------------------------
 ratio_corriente              rango         1.23
 margen_bruto                 rango        23.9%
 tipo_cambio_cop              rango     4,123.46
 capex_comprometido           rango    1,500,000
```

### Paso 3 — Coloque el hueco en el Word

```bash
python src\fs_documento.py insertar "MI_DOCUMENTO.docx" ratio_corriente actual
```

O desde Word, si prefiere elegir el sitio exacto dentro de una frase:
**Programador → Control de contenido de texto**, y en *Propiedades* →
**Etiqueta** escriba `fs-dato-ratio_corriente-actual`.

Refresque, y la cifra aparece. Usted redacta alrededor:

> «El ratio corriente cerró en **1.23**, por encima del covenant, con un
> tipo de cambio de **4,123.46** COP/USD.»

Las dos cifras en negrita las mantiene el Excel. El resto de la frase es
suya y el refresco no la toca.

### El formato lo pone Excel, no el programa

Esto es lo que hace cómodo el mecanismo: **se respeta el formato de número
que usted puso en la celda**.

| En Excel | Formato de celda | En el Word |
|---|---|---|
| `1.2345` | `0.00` | `1.23` |
| `0.2387` | `0.0%` | `23.9%` |
| `4123.456` | `#,##0.00` | `4,123.46` |
| `1500000` | `#,##0` | `1,500,000` |
| `-8300` | `#,##0` | `(8,300)` |
| una fecha | cualquiera | `2026-06-30` |

Si quiere más decimales, cámbieselos a la celda en Excel. No hay nada que
configurar en el programa.

> Ojo: las cifras **de la tabla del estado** siguen usando el formato
> contable de siempre (miles, sin decimales, negativos entre paréntesis).
> Esta tabla aplica solo a las celdas sueltas que usted nombre.

### Por qué esto es una conexión, y no una exportación

El Excel sigue siendo el único sitio donde se editan los números — eso no
cambia. Lo que cambia es que **ya no hace falta pasar por la tabla**: usted
decide qué celda quiere citar, le pone nombre, y el Word la reclama por ese
nombre.

```
   Word                          Excel
   ────                          ─────
   «...cerró en [   ] »   ──►  ¿quién es fs_ratio_corriente?
     fs-dato-ratio_corriente-actual        │
                            ◄──────────────┘  Hoja "Ratios", celda H1
                                              formato 0.00  ->  "1.23"
```

El Word **pide por nombre**; el Excel **responde**. Añadir una cifra nueva
no toca el código ni la plantilla: es un nombre en Excel y una etiqueta en
Word.

### Si el nombre no existe

El refresco **no falla ni inventa un valor**. Reporta el hueco como
huérfano y deja lo que hubiera:

```
 AVISO — el documento pide cifras que el Excel ya no tiene:
   ? fs-dato-ratio_corriente-actual
```

Es el aviso correcto: alguien borró o renombró esa celda en Excel, y hay
que decidir qué hacer. Nunca verá una cifra equivocada por este camino.

---

## Editar los números **desde Word**

Sí se puede. Pero hay que entender qué pasa después, porque hay **tres
estados** distintos y solo uno conserva lo que usted teclee.

| Estado | ¿Se puede teclear en Word? | ¿Sobrevive al refresco? | Para qué |
|---|---|---|---|
| **Bloqueado** *(por defecto)* | No | — | El uso normal |
| **Desbloqueado** | Sí | **No** | Un retoque de un rato, antes de refrescar |
| **Desvinculado** | Sí | **Sí** | Una cifra que no está en el Excel |

### Estado 1 · Bloqueado — el de siempre

Word no deja escribir dentro de la tabla ni de las cifras. Es lo que protege
la referencia: nadie puede teclear un número que el Excel contradiga.

```bash
python src\fs_documento.py bloquear "MI_DOCUMENTO.docx"
```

### Estado 2 · Desbloqueado — editable, pero temporal

```bash
python src\fs_documento.py desbloquear "MI_DOCUMENTO.docx"
```

A partir de ahí Word deja teclear encima de cualquier cifra.

> **Lo que escriba a mano lo machaca el siguiente refresco.** La región
> sigue vinculada al Excel: al refrescar vuelve el valor del libro. No es un
> fallo — es lo que impide que el Word y el Excel se contradigan en silencio.

Sirve para un ajuste que vaya a durar poco: enseñar un supuesto en una
reunión, imprimir una versión con una cifra provisional. Para volver atrás:

```bash
python src\fs_documento.py bloquear "MI_DOCUMENTO.docx"
```

Se puede desbloquear **una sola cifra** en vez de todas:

```bash
python src\fs_documento.py desbloquear "MI_DOCUMENTO.docx" total_assets
```

### Estado 3 · Desvinculado — el valor a mano se queda

Cuando la cifra **no va a venir nunca del Excel** (un dato de un tercero, un
importe pendiente de incorporar), lo correcto no es desbloquear: es cortar el
vínculo.

```bash
python src\fs_documento.py desvincular "MI_DOCUMENTO.docx" total_assets
```

El recuadro desaparece y su contenido se queda como **texto normal**. A
partir de ahí:

- se edita como cualquier palabra del documento
- el refresco **no lo toca**
- **no** sale reportado como huérfano, porque ya no hay ancla

Para volver a vincularlo algún día: `insertar` otra vez.

### Desde Word, sin línea de órdenes

Lo mismo se hace a mano si prefiere no tocar la consola:

1. Pestaña **Programador** (si no está: *Archivo → Opciones → Personalizar
   cinta → Programador*).
2. Clic dentro de la cifra. Se marca el recuadro del control.
3. **Propiedades**.
4. Desmarque **«No se puede editar el contenido»** → equivale a
   *desbloquear*.
   Marque **«Quitar control de contenido al editarlo»**, edite, y el control
   desaparece → equivale a *desvincular*.

El menú **Quitar control de contenido** (clic derecho sobre el recuadro) es
exactamente `desvincular` para esa cifra.

---

## ¿Y desde OneDrive?

**OneDrive no tiene órdenes.** Es almacenamiento: guarda y sincroniza el
archivo, no sabe nada de las cifras. Todo lo de arriba se hace sobre el
documento, esté donde esté.

Dos avisos concretos:

- **Word en el navegador maneja mal los controles de contenido.** Puede que
  no vea el recuadro o que lo trate como texto suelto. Para todo lo de esta
  página, **use Word de escritorio**.
- **Cierre el documento antes de refrescar.** Si Word lo tiene abierto, la
  orden se detiene y le dice quién lo retiene. Escribir encima de un `.docx`
  abierto lo deja inservible.

---

## Los dos candados, que no son lo mismo

Es fácil confundirlos:

| | Candado de región | Protección del documento |
|---|---|---|
| **Orden** | `bloquear` / `desbloquear` | `proteger` / `desproteger` |
| **Alcance** | cada cifra y cada tabla, por separado | el documento entero |
| **Qué impide** | teclear dentro de una cifra | escribir en cualquier sitio salvo las zonas de redacción |
| **Contraseña** | no | sí |
| **Para qué** | que nadie pise una cifra del Excel | separar el rol Redactor del de Editor de datos |

Se combinan: lo normal es **candado de región puesto** y **sin protección de
documento** — las cifras intocables, la prosa libre. Es como está hoy.

## Si de verdad hace falta escribir un número a mano

A veces hay una cifra que no está en el modelo: un dato de un tercero, un
importe que aún no se ha incorporado.

**Escríbala como prosa normal**, fuera de toda región marcada. El motor no la
tocará, porque solo escribe dentro de las anclas. Pero tenga presente que
entonces **no está vinculada a nada**: si el modelo cambia, esa cifra se
queda como está y nadie avisará.

Si esa cifra va a **repetirse cada cierre**, llévela al Excel y déle un
rango con nombre; si es puntual, use `desvincular` (arriba).

```bash
python src\fs_documento.py nombrar "MI_LIBRO.xlsx" --aplicar
```

---

## Resumen

| Quiero… | Dónde |
|---|---|
| Cambiar un número | **Excel** → guardar → `refrescar.bat` |
| Teclear una cifra en Word, solo por un rato | `desbloquear` (se pierde al refrescar) |
| Teclear una cifra en Word **para siempre** | `desvincular` |
| Volver a proteger las cifras | `bloquear` |
| Escribir un párrafo | **Word**, en el documento base |
| Intercalar una cifra viva en un párrafo | `insertar`, y luego escribir alrededor en Word |
| Una entrega puntual congelada | `generar.bat` → `salidas\` |
| Saber qué cifras hay disponibles | `catalogo` |
| Saber qué cambió en el último refresco | la bitácora del propio documento |
