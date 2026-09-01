# Guía de operación — el documento que se actualiza solo

Cómo funciona hoy el documento base de Word: se le refrescan las cifras desde
el Excel sin perder una coma de lo que hayas escrito.

---

## Qué cambió

| Antes | Ahora |
|---|---|
| El programa creaba un Word **nuevo** cada vez, desde una plantilla fija. Lo que alguien hubiera redactado se quedaba en el documento anterior. | Hay un documento **base, vivo**, en OneDrive. Se le refrescan solo unas regiones marcadas. La redacción vive fuera de ellas, así que no se toca nunca. |

**La regla que lo explica todo:** el motor **solo escribe dentro de las
regiones marcadas**. Lo demás ni lo visita. Por eso tu párrafo sobrevive al
refresco — y por eso, si lo borras, se queda borrado: el motor nunca lo
guardó, así que no puede reinyectarlo.

---

## Sobre «tiempo real»

**No es tiempo real.** El documento no se actualiza solo cuando alguien toca
el Excel. La actualización ocurre cuando *alguien la lanza*: una orden, o un
botón en el panel de Word cuando el complemento esté listo.

Lo que sí es inmediato es el resultado: la orden tarda unos segundos y el
documento queda al día, con un registro de qué cifra cambió. Un refresco
automático de verdad (al guardar el Excel) es posible más adelante con Power
Automate, pero **hoy no está montado**.

---

## Cómo se opera, paso a paso

Todas las órdenes se ejecutan desde la carpeta del proyecto. Los pasos 1 y 2
se hacen una sola vez; el 5 es el del día a día.

### 1. Fijar las cifras del Excel · *una vez*

Cada fila del estado recibe un **nombre de Excel** (`fs_total_assets`,
`fs_cash_and_cash_equivalents`…) que le da identidad propia. Sin esto, el
vínculo se apoya en el texto de la etiqueta y se rompe si alguien renombra
una fila.

```bash
python fs_documento.py nombrar "MI_LIBRO.xlsx"
```

```bash
python fs_documento.py nombrar "MI_LIBRO.xlsx" --aplicar
```

*Ya aplicado sobre `Copia_Editable_con_columna_Tipo.xlsx`: 28 nombres.*

> **Cierra el Excel antes.** Los nombres los escribe **Excel**, no el script.
> Es deliberado: openpyxl no recalcula fórmulas y, al reguardar, descartaría
> el valor de todas ellas — el Word saldría en blanco. Si el libro está
> abierto, la orden falla y te lo dice.

### 2. Preparar el documento base · *una vez*

Le añade al documento las regiones que le falten: la tabla, los campos de
encabezado, las zonas de redacción, la bitácora. **No borra nada**, así que
puedes correrlo sobre un documento con meses de trabajo encima.

```bash
python fs_documento.py construir "MI_DOCUMENTO.docx"
```

¿No tienes documento todavía? Parte de la plantilla del repositorio,
[`plantilla_base_EF.docx`](plantilla_base_EF.docx), o genera una nueva:

```bash
python fs_documento.py plantilla "NUEVO.docx"
```

### 3. Redactar · *libre*

Se abre el documento en Word y se escribe con normalidad. Párrafos, listas,
comentarios, control de cambios — lo que sea. Hay dos zonas señaladas
(*introducción* y *análisis*) que existen para marcar dónde se espera texto,
pero en el modo actual **se puede escribir en cualquier parte**.

Lo que *no* se puede tocar a mano es la tabla ni las cifras: están
bloqueadas. Word simplemente no deja escribir dentro.

### 4. Intercalar cifras vivas en la redacción · *cuando haga falta*

Para escribir *«los activos totales ascendieron a **119.066.301**»* y que esa
cifra siga al Excel sin que nadie la vuelva a teclear. Primero se consulta qué
hay disponible:

```bash
python fs_documento.py catalogo
```

Y se coloca la cifra donde toque:

```bash
python fs_documento.py insertar "MI_DOCUMENTO.docx" total_assets actual
```

Campos posibles: `actual`, `previo`, `nota`, `var_abs` (la diferencia) y
`var_pct` (la variación porcentual).

También se puede hacer desde Word: **Programador → Control de contenido de
texto**, y ponerle la etiqueta `fs-dato-total_assets-actual`.

### 5. Refrescar · *cada cierre*

El paso del día a día. Un comando, unos segundos.

```bash
python fs_documento.py refrescar "MI_DOCUMENTO.docx" "MI_LIBRO.xlsx"
```

Qué hace, exactamente:

- Reescribe la tabla del estado con las cifras nuevas.
- Actualiza los campos de encabezado y las cifras de la redacción.
- Añade a la bitácora del documento qué cambió, cifra por cifra.
- Deja una copia `.bak` del documento anterior.
- Avisa si alguna cifra de tu texto ya no existe en el Excel.

> **Cierra el documento en Word antes.** El motor escribe el archivo
> directamente; si Word lo tiene abierto, se pelean por él. Y evita refrescar
> mientras alguien más lo edita en el navegador: la coautoría de Word Online
> maneja mal este tipo de documento.

### 6. Comprobar · *cuando dudes*

```bash
python fs_documento.py verificar "MI_DOCUMENTO.docx"
```

Lista las regiones que tiene el documento, qué cifras hay intercaladas y
cuáles se han quedado huérfanas.

---

## Cómo están los bloqueos

Hoy el documento está en **modo abierto**: las regiones de datos están
bloqueadas una por una, pero el documento en sí no lleva contraseña.

| Región | Etiqueta | ¿Editable a mano? | ¿La refresca el motor? |
|---|---|---|---|
| Tabla del estado | `fs-tabla-principal` | Bloqueada | Sí |
| Campos de encabezado | `fs-campo-*` | Bloqueados | Sí |
| Cifras dentro del texto | `fs-dato-*` | Bloqueadas | Sí |
| Zonas de redacción | `fs-prosa-*` | **Libres** | Nunca |
| Bitácora | `fs-registro` | Bloqueada | Añade encima |
| Texto suelto | *sin etiqueta* | **Libre** | Nunca |

### Los dos editores, cuando se quiera activar

Existe un **modo estricto** que pone el documento entero en solo lectura y
abre huecos únicamente en las zonas de redacción:

```bash
python fs_documento.py proteger "MI_DOCUMENTO.docx" --clave TU_CLAVE
```

| Rol | Puede | No puede |
|---|---|---|
| **Redactor** | Escribir dentro de las zonas de redacción | Tocar tablas, campos ni cifras |
| **Editor de datos** | Lo anterior, y además refrescar | — (tiene la clave) |

Está probado y funciona, pero **no está activado** en el documento actual:
conviene decidir antes quién custodia la clave. Se quita con `desproteger`.

---

## Qué falta del proceso anterior

### Lo que se conserva íntegro

- La detección de la hoja y de las columnas por contenido.
- La clasificación de filas (`H`/`I`/`S`/`T`/`N`/`X`) y la inferencia cuando
  no hay columna *Tipo*.
- El formato contable, con negativos entre paréntesis.
- `GeneradorFS.exe` sigue funcionando igual, para quien lo prefiera.

### Lo que todavía no está

- **El complemento de Word está sin compilar.** El código está escrito y
  alineado con el mismo contrato, pero no se ha probado. Hoy todo se opera
  por línea de órdenes.
- **El complemento formatea peor que el motor de Python.** Si se refrescara
  un documento con él, se perderían los filetes de subtotal y total. Hay que
  igualarlo antes de usarlo.
- **Los subtotales no tienen nombre.** No tienen etiqueta donde anclarlo, así
  que no se pueden intercalar en la redacción por clave propia. Se
  identifican por su sección.
- **Una sola tabla y un solo tipo de estado.** La estructura para partir el
  estado en varias tablas ya existe, pero no está en uso. Las notas al pie
  siguen fuera.
- **El Excel se lee desde el disco.** Leerlo directamente de OneDrive
  requiere el complemento y un permiso de solo lectura aprobado por TI.

---

## Por dónde seguir

| Fase | Qué es | Estado |
|---|---|---|
| 0 | Contrato de regiones y documento base | ✅ Hecho |
| 1 | Motor de refresco en el sitio | ✅ Hecho |
| 2 | Identidad estable por rangos con nombre | ✅ Hecho |
| 3 | Modo estricto de dos editores | ⏸ Listo, sin activar |
| 4 | Complemento de Word | ○ Pendiente |
| 5 | Pruebas de archivo dorado en CI | ○ Pendiente |

**Lo siguiente, en orden:**

1. **Usar el documento un cierre entero** por línea de órdenes. Es lo único
   que dirá si el modelo aguanta el trabajo real antes de invertir semanas en
   el complemento.
2. **Decidir si se activa el modo estricto**, y quién guarda la clave.
3. **Instalar Node.js y cerrar el complemento.** Recuerda: Node.js es solo
   para compilar. Pamela y Violeta no instalan nada — el panel les aparece en
   la cinta de Word y ya.

---

Especificación completa en [`CONTRATO.md`](CONTRATO.md) · Motor en
[`fs_documento.py`](fs_documento.py)
