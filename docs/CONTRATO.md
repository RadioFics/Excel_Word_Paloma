# Contrato de anclas — Excel ⇄ Word

Especificación única que consumen `fs_contrato.py` / `fs_documento.py`
(Python) y el add-in de Word (`addin/`). Si cambias algo aquí, cambia en los
dos lados.

---

## 1. La idea en una frase

El documento de Word **no se regenera**: se le **refrescan unas regiones
concretas**. Todo lo que no sea una región marcada ni se visita, y por eso
la redacción de una persona sobrevive intacta a cada actualización.

```
        ┌─────────────────── EF_base.docx (vive en OneDrive) ────────────────┐
        │                                                                    │
        │  Collective Mining Ltd.        ← fs-campo-empresa      REFRESCADO  │
        │  Estado de Situación Financiera  ← fs-campo-titulo     REFRESCADO  │
        │                                                                    │
        │  Durante el semestre la compañía...   ← fs-prosa-*     INTACTO     │
        │  ...escrito por Pamela, a mano.                                    │
        │                                                                    │
        │  ┌──────────────────────────────┐                                  │
        │  │ tabla del estado             │  ← fs-tabla-principal REFRESCADO │
        │  └──────────────────────────────┘                                  │
        │                                                                    │
        │  Los activos totales fueron [119,066,301], un [23.9%] más...       │
        │   └ prosa INTACTA ──┘        └ fs-dato-* REFRESCADO ┘              │
        └────────────────────────────────────────────────────────────────────┘
```

**Consecuencia buscada:** si borras ese párrafo y vuelves a refrescar, el
párrafo sigue borrado y las cifras salen actualizadas. El motor nunca
reinyecta prosa porque nunca la guardó.

---

## 2. Familias de ancla

Cada ancla es un **control de contenido** de Word (`w:sdt`) identificado por
su **Etiqueta** (Tag). En Word: pestaña *Programador* → *Control de
contenido* → *Propiedades* → *Etiqueta*.

| Etiqueta | Tipo de control | ¿La refresca el motor? | Para qué |
|---|---|---|---|
| `fs-tabla-<nombre>` | Texto enriquecido (bloque) | **Sí** — reescribe la tabla entera | Una tabla del estado |
| `fs-campo-<nombre>` | Texto sin formato (en línea) | **Sí** | Un campo de encabezado |
| `fs-dato-<clave>-<campo>` | Texto sin formato (en línea) | **Sí** | Una cifra suelta dentro de la prosa |
| `fs-prosa-<nombre>` | Texto enriquecido (bloque) | **No** | Zona de redacción libre |
| `fs-registro` | Texto enriquecido (bloque) | Antepone | Bitácora de actualizaciones |
| `fs-meta` | Texto enriquecido, oculto | Sobrescribe | Foto de la última corrida (JSON) |

Todo control cuya etiqueta no encaje en el contrato **se ignora**: puedes
usar controles de contenido para tus propios fines sin que el motor los toque.

### Reglas de nombre

- Los segmentos van separados por `-`.
- `<nombre>` y `<clave>` solo pueden llevar `[a-z0-9_]`.
- Máximo 64 caracteres (límite de Word para una Etiqueta).

---

## 3. Campos de encabezado — `fs-campo-<nombre>`

Los ocho nombres válidos, tal como salen del Excel:

| `<nombre>` | Contenido | De dónde sale |
|---|---|---|
| `empresa` | Collective Mining Ltd. | `config.json` → `empresa` |
| `titulo` | As at prueba 2 | primer texto de la columna etiqueta |
| `fecha_actual` | June 30, | encabezado de la columna actual |
| `fecha_previa` | December 31, | encabezado de la columna comparativa |
| `miles` | 1000 | fila 2 de la columna actual |
| `estado_actual` | Unaudited | fila 3 de la columna actual |
| `estado_previo` | Audited | fila 3 de la columna comparativa |
| `moneda` | $ | fila 4 de la columna actual |

---

## 4. Cifras sueltas en la prosa — `fs-dato-<clave>-<campo>`

Este es el mecanismo para escribir *"los activos totales ascendieron a
**119,066,301**"* y que esa cifra siga al Excel sin que nadie la retipee.

### La clave — dos orígenes

Una cifra de la prosa se identifica por una **clave**. Hay dos maneras de
obtenerla, y no son igual de firmes:

| Origen | Cómo | Aguanta |
|---|---|---|
| **`rango`** *(recomendado)* | un nombre de Excel `fs_<clave>` que apunta a la celda de etiqueta de esa fila | renombrar la fila, insertar/borrar filas encima, reordenar |
| `etiqueta` *(respaldo)* | se deriva del texto de la etiqueta | **se rompe si alguien renombra la fila** |

El motor prefiere siempre el rango; si la fila no tiene nombre, cae en la
etiqueta. `catalogo` muestra el origen de cada clave.

#### Crear los rangos

```bash
python src\fs_documento.py nombrar <libro.xlsx>            # simulación
python src\fs_documento.py nombrar <libro.xlsx> --aplicar  # los escribe
```

Crea un nombre `fs_<clave>` por cada línea con etiqueta, apuntando a su
celda de etiqueta. A partir de ahí, esa fila tiene identidad propia.

> **Por qué con Excel y no con openpyxl.** openpyxl no recalcula fórmulas y
> al reguardar descarta el valor cacheado de **todas** ellas: el Word saldría
> con las cifras en blanco. La orden pilota Excel (vía PowerShell) para que
> sea Excel quien guarde el libro. Cierre el libro antes de ejecutarla.

También puede crearlos a mano: en Excel, seleccione la celda de la etiqueta
y escriba `fs_lo_que_sea` en el **Cuadro de nombres** (arriba a la izquierda).

#### Escalares fuera de la tabla

Un nombre `fs_*` que apunte **fuera** de la región de datos (una fecha de
corte, un tipo de cambio, un dato de otra hoja) se expone igual, con el campo
`actual`. Sirve para intercalar en la redacción cifras que no son filas del
estado.

### La clave derivada de la etiqueta

Cuando no hay rango, se deriva del texto con un algoritmo que debe dar
idéntico resultado en Python y en TypeScript:

1. quitar tildes (NFKD, descartando diacríticos)
2. minúsculas
3. todo lo que no sea `[a-z0-9]` pasa a `_`
4. colapsar `_` repetidos, recortar los de los extremos
5. cortar a 40 caracteres y volver a recortar `_` del final

```
'Cash and cash equivalents'  ->  cash_and_cash_equivalents
'Total assets'               ->  total_assets
'Provisión (neta)'           ->  provision_neta
```

Para ver las claves reales de tu libro:

```bash
python src\fs_documento.py catalogo
```

### El campo

| `<campo>` | Qué escribe | Ejemplo |
|---|---|---|
| `actual` | cifra del periodo actual, formato contable | `119,066,301` |
| `previo` | cifra del periodo comparativo | `96,094,583` |
| `nota` | número de nota | `10` |
| `var_abs` | diferencia actual − previo | `22,971,718` |
| `var_pct` | variación porcentual | `23.9%` |

Los negativos van entre paréntesis: `(142,204,537)`.

### Cómo insertar una

1. En Word, pon el cursor donde va la cifra.
2. *Programador* → **Control de contenido de texto sin formato**.
3. *Propiedades* → **Etiqueta** = `fs-dato-total_assets-actual`.
4. Marca **«No se puede editar el contenido»** para que nadie la pise a mano.
5. Refresca. El motor pone la cifra.

Si escribes una clave que el Excel no tiene, el motor **no falla**: la
reporta como **huérfana** y deja el texto que hubiera.

---

## 5. Tablas — `fs-tabla-<nombre>`

- `fs-tabla-principal` recibe **todas** las líneas del estado.
- Cualquier otro nombre recibe **solo la sección** cuyo encabezado (fila de
  tipo `H`) produce esa misma clave. Para partir el estado en varias tablas
  no hay que tocar código: basta con nombrar el control
  `fs-tabla-current_assets` y esa tabla recibe solo esa sección.

### Qué se conserva al refrescar

El motor **rehace las filas** pero **respeta**:

- `w:tblPr` — estilo de tabla, bordes generales, sombreado
- `w:tblGrid` — los anchos de columna. Si arrastras una columna en Word, el
  refresco siguiente respeta la medida nueva.

Las filas se generan siempre desde el modelo, con formato por tipo:

| Tipo | Formato |
|---|---|
| `H` encabezado de sección | negrita, sin sangría, sin cifras |
| `I` línea de detalle | sangría, cifras a la derecha |
| `S` subtotal | filete superior, sin etiqueta |
| `T` total | negrita, filete superior + doble inferior |
| `N` nota de texto | cursiva, sin cifras |

---

## 6. Zonas de prosa — `fs-prosa-<nombre>`

El motor **nunca** las toca. Existen por dos razones:

1. Documentan dónde se espera redacción.
2. En **modo estricto** (§8) son los únicos rangos que el rol Redactor puede
   editar.

Puedes crear, renombrar o borrar las que quieras. No son obligatorias: en
modo abierto la persona puede escribir en cualquier parte del documento.

---

## 7. Bitácora y metadatos

- `fs-registro` — el motor **antepone** un bloque por cada refresco con la
  fecha, el Excel de origen y las diferencias contra la corrida anterior.
  No borra lo anterior.
- `fs-meta` — párrafo **oculto** con un JSON: la foto de las cifras de la
  última corrida. Es lo que permite calcular el «antes → después». Viaja
  dentro del `.docx`, así que funciona igual en OneDrive y entre equipos.

---

## 8. Los dos editores

### Modo abierto (por defecto)

Sin protección de documento. Las regiones de datos llevan
`w:lock="sdtContentLocked"`: Word impide editarlas a mano, pero la prosa es
libre en todo el documento. **Es el modo recomendado para empezar.**

### Modo estricto — `proteger`

```bash
python src\fs_documento.py proteger EF_base.docx --clave <clave>
```

Pone el documento en solo lectura (`w:documentProtection edit="readOnly"`) y
abre rangos editables (`w:permStart`/`w:permEnd`, grupo *everyone*) alrededor
de cada zona `fs-prosa-*`.

| Rol | Puede | No puede |
|---|---|---|
| **Redactor** | escribir dentro de las zonas `fs-prosa-*` | tocar tablas, campos ni cifras sueltas |
| **Editor de datos** | todo lo anterior + refrescar | — (conoce la clave, o usa el add-in) |

El add-in del Editor de datos hace: **desproteger → refrescar → reproteger**.

> **Detalle práctico:** el rango editable termina en la última marca de
> párrafo de la zona. Por eso cada zona se crea con un párrafo vacío de
> holgura: sin él, el Redactor puede corregir el texto que ya hay pero no
> arrancar un párrafo nuevo al final.

---

## 9. Garantías que el motor debe cumplir

1. **Idempotencia.** Refrescar dos veces con el mismo Excel deja las
   regiones de datos idénticas byte a byte. (Los `w:id` de los controles se
   derivan del hash de la etiqueta justamente para esto.)
2. **No visitar la prosa.** El refresco solo escribe dentro de anclas de las
   familias tabla / campo / dato / registro / meta.
3. **Reparación no destructiva.** `construir` sobre un documento con meses
   de redacción encima solo añade las anclas que falten.
4. **Fallo explícito.** Un ancla sin dato en el Excel se reporta como
   huérfana; nunca se rellena con una cifra equivocada.

---

## 10. Deuda conocida

- Dos etiquetas distintas del Excel que produzcan la misma clave colisionan.
  El motor conserva la primera y **lo reporta**; no lo resuelve solo.
- Las filas **sin rango con nombre** siguen dependiendo del texto de la
  etiqueta. Corra `nombrar --aplicar` para fijarlas; `catalogo` dice cuáles
  faltan.
- Las filas de **subtotal (S) no reciben nombre**: no tienen etiqueta donde
  anclarlo. Se identifican por su sección (*«Subtotal de Current assets»*).
- Si alguien **borra** una fila con nombre en Excel, el nombre queda en
  `#REF!` y el motor lo ignora: el ancla del Word sale reportada como
  huérfana, que es el aviso correcto.
- La co-autoría de Word en el navegador maneja mal los controles de
  contenido. **El refresco debe correrse desde Word de escritorio o desde la
  línea de órdenes, con el archivo cerrado por los demás.**
