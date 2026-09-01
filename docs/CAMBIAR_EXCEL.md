# Cuando llegue el Excel definitivo

El libro que hay hoy en `ejemplos\` es una **base desactualizada**. Cuando
llegue el bueno —o cuando el actual cambie de estructura— esto es lo que hay
que revisar, en orden. Nada de esto obliga a tocar código.

---

## Lo primero: mirar antes de refrescar

```bash
EstadosFinancieros.exe MI_LIBRO_NUEVO.xlsx --estado
```

Le dice cómo está leyendo el libro **sin escribir nada**: qué hoja eligió,
qué columnas detectó, cuántas líneas salieron, cuántas tienen rango con
nombre y si alguna cifra de su Word se quedó huérfana.

Si eso sale bien, el resto probablemente también.

---

## 1. La hoja

El programa busca una hoja llamada `FS`. Si no existe, elige la que más
señales tenga de ser un estado de situación financiera.

**Si se equivoca**, fíjela a mano en `config.json`:

```json
"hoja": "Balance Consolidado"
```

Su libro real tiene **44 hojas**: merece la pena fijar el nombre en vez de
confiar en la detección.

---

## 2. Las columnas

Se detectan por contenido: las dos con más números son las cifras, la de más
texto a la izquierda es la etiqueta, etc. Hoy salen `A / C / E / F / G`.

**Si el libro nuevo las mueve**, fíjelas:

```json
"columnas": {
  "etiqueta": "B", "nota": "D", "actual": "F", "previo": "G", "tipo": "H"
}
```

Deje en `null` las que quiera que siga detectando.

---

## 3. La columna `Tipo`

Es **opcional**. Si el libro nuevo no la trae, el programa deduce el papel de
cada fila y deja el detalle en `salidas\revisar_tipos.csv`. Revíselo la
primera vez.

Si prefiere fijarlo, añada una columna `Tipo` con `H`/`I`/`S`/`T`/`N`/`X`.

---

## 4. Los rangos con nombre — **esto es lo que hay que rehacer**

Los nombres `fs_*` viven **dentro del libro**. Un libro nuevo no los tiene:
hay que crearlos otra vez.

```bash
python src\fs_documento.py nombrar "MI_LIBRO_NUEVO.xlsx"
```

Eso **simula** y enseña el plan. Si convence:

```bash
python src\fs_documento.py nombrar "MI_LIBRO_NUEVO.xlsx" --aplicar
```

> **Cierre el libro en Excel antes.** Los nombres los escribe Excel, no el
> programa, para no destruir las fórmulas.

### Qué pasa si no lo hace

Nada se rompe. Las cifras vuelven a identificarse por el **texto de la
etiqueta**, que es como funcionaba antes. Solo pierde la protección frente a
renombrados. `catalogo` le dirá cuántas claves dependen de cada cosa:

```
 28 de 33 claves vienen de un rango con nombre (identidad estable).
 Las de origen 'etiqueta' se rompen si alguien renombra la fila.
```

### Si las etiquetas cambian de texto

Aquí es donde se nota. Con rango con nombre, el vínculo aguanta. Sin él, el
ancla del Word queda **huérfana** y se reporta:

```
 AVISO — el documento pide cifras que el Excel ya no tiene:
   ? fs-dato-total_assets-actual
```

No se rellena con un valor equivocado: se avisa y usted decide.

---

## 5. Las cifras sueltas de otras hojas

Si el libro nuevo trae ratios, tipos de cambio o cualquier celda que quiera
citar en la redacción, nómbrela a mano en Excel con el prefijo `fs_`. Ver
[DATOS.md](DATOS.md) → *Llevar al Word cifras que NO son filas de la tabla*.

---

## 6. El documento de Word no hay que rehacerlo

Esto es lo importante: **el documento sobrevive al cambio de Excel**. Sus
regiones se llaman por su nombre, no por posiciones. Si el libro nuevo tiene
las mismas claves, el refresco simplemente funciona.

Solo tendrá que tocar el Word si:

- **aparecen filas nuevas** que quiera citar en la prosa → `insertar`
- **desaparecen filas** que ya citaba → el refresco las reporta como
  huérfanas y usted borra esa mención

En ninguno de los dos casos hay que reconstruir nada.

---

## Lista de comprobación

```
[ ] EstadosFinancieros.exe MI_LIBRO.xlsx --estado     ¿lee bien el libro?
[ ] revisar salidas\revisar_tipos.csv                 ¿clasificó bien las filas?
[ ] fs_documento.py nombrar MI_LIBRO.xlsx             ¿qué nombres crearía?
[ ] fs_documento.py nombrar MI_LIBRO.xlsx --aplicar   crearlos
[ ] fs_documento.py catalogo MI_LIBRO.xlsx            ¿salen todas las claves?
[ ] refrescar sobre una COPIA del documento primero
[ ] revisar el .log de salidas\                       ¿los cambios cuadran?
[ ] refrescar el documento de verdad
```

El paso de la copia no es paranoia: la primera vez con un libro nuevo es
justo cuando conviene ver el resultado antes de tocar el documento bueno.
