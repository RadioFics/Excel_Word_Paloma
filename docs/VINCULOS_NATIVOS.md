# Vincular Excel y Word sin programar

Sí se puede, y es una función de Word de toda la vida: **Pegado especial →
Pegar vínculo**. No hace falta este programa ni ningún código.

Aquí está el paso a paso, y también **lo que se rompe**, que lo he
comprobado ejecutándolo.

---

## El paso a paso

### 1 · En Excel: copie la celda

Seleccione la celda con la cifra y **Ctrl+C**. Deje Excel abierto.

### 2 · En Word: pegue como vínculo

Ponga el cursor donde va la cifra, dentro de su frase, y:

**Inicio → Pegar (la flecha de abajo) → Pegado especial…**

En el cuadro que sale:

1. Marque **Pegar vínculo** (la opción de la izquierda, no *Pegar*).
2. En la lista, elija **Texto con formato (RTF)**.
3. Aceptar.

> **Elija bien el formato.** *Texto con formato (RTF)* mete la cifra en la
> línea, como una palabra más: se justifica, hereda el estilo y se imprime
> igual. *Objeto de hoja de cálculo* mete un objeto incrustado que estorba
> en medio de un párrafo. *Texto sin formato* funciona pero pierde el
> formato de número.

Ya está. El párrafo se lee así:

> Total de activos: 119.066.301

Y por dentro Word ha escrito un **campo**:

```
{ LINK Excel.Sheet.12 "C:\ruta\libro.xlsx" "FS!F16C5" \a \f 4 \r }
```

Para verlo: **Alt+F9** alterna entre el valor y el código.

### 3 · Actualizar

| Cómo | Qué actualiza |
|---|---|
| **F9** con el cursor en el campo | ese vínculo |
| **Ctrl+E** y luego **F9** | todos |
| Clic derecho → **Actualizar vínculo** | ese vínculo |
| Al abrir el documento | Word pregunta; conteste *Sí* |

Para que pregunte siempre: **Archivo → Opciones → Avanzadas → General →
«Actualizar vínculos automáticos al abrir»**.

### 4 · Que no se pueda editar a mano

Un campo ya no se edita como texto normal: si alguien escribe encima, la
siguiente actualización lo devuelve a su sitio.

Para **congelarlo** del todo (que ni siquiera se actualice):

- **Ctrl+F11** bloquea el campo
- **Ctrl+Mayús+F11** lo desbloquea

O desde **Archivo → Información → Editar vínculos a archivos**, marcando
**Bloqueado**.

### 5 · Romper el vínculo

**Ctrl+Mayús+F9** convierte el campo en texto normal, para siempre. Es lo
que se hace cuando esa cifra ya no debe seguir al Excel.

---

## Lo que se rompe

Esto lo he ejecutado y medido, no es teoría.

### El vínculo guarda **coordenadas**, no el concepto

El campo apunta a `"FS!F16C5"` — hoja FS, fila 16, columna 5. **No** al
nombre de la fila ni a un rango con nombre.

**Consecuencia: si inserta una fila arriba en Excel, el Word muestra un
número equivocado y no avisa.**

Medido:

```
   Antes:   E16 = 119.066.301  (Total assets)     Word: 119.066.301   correcto

   -- se inserta una fila arriba en Excel --

   Ahora:   E16 =  44.066.794  (otro subtotal)
            E17 = 119.066.301  (Total assets se movio aqui)

   Word tras actualizar:  44.066.794    <-- EQUIVOCADO, y sin ningun aviso
```

Es el fallo peligroso: no da error, da una cifra falsa.

### Los rangos con nombre **no sirven** aquí

Lo intenté: editar el campo a mano para que apunte a un nombre de Excel en
vez de a las coordenadas.

```
{ LINK Excel.Sheet.12 "C:\ruta\libro.xlsx" "mi_rango_con_nombre" \a \f 4 \r }
```

Resultado: **«¡Error! Vínculo no válido.»** El vínculo OLE de Word solo
resuelve referencias de celda, no nombres definidos. Probado con Excel
abierto y con Excel cerrado.

### La ruta es absoluta

El campo guarda `C:\Users\TuNombre\OneDrive - ...\libro.xlsx`. Si el libro
se mueve, se renombra, o **lo abre otra persona cuya carpeta de OneDrive
tiene otra ruta**, el vínculo deja de encontrarlo.

En un OneDrive compartido esto importa: la ruta de Pamela no es la de
Violeta. Un documento con vínculos hechos en un equipo llega roto al otro.

### Otras cosas a tener en cuenta

- **Rutas largas.** Si la ruta pasa de 260 caracteres, el vínculo falla con
  «vínculos a archivos que no se encuentran». Me pasó al probar.
- **Word en el navegador no actualiza vínculos.** Solo el de escritorio.
- **Aviso al abrir.** Cada vez que se abra el documento saldrá el cartel de
  «contiene vínculos a otros archivos».
- **Muchos vínculos = lento.** Un estado entero celda a celda hace que abrir
  el documento tarde.

---

## Entonces, ¿cuál uso?

Los dos, para cosas distintas.

| | Vínculo nativo de Word | Este programa |
|---|---|---|
| **Quién lo monta** | cualquiera, desde Word | una orden, o marcar el control |
| **Referencia** | coordenadas de celda | nombre estable (`fs_*`) |
| **Insertar filas en Excel** | **da una cifra falsa, sin avisar** | aguanta |
| **Renombrar una fila** | aguanta | aguanta (con rango con nombre) |
| **Mover el Excel** | se rompe | da igual: se elige en cada refresco |
| **Otro equipo / otro usuario** | se rompe (ruta absoluta) | funciona |
| **Si falta el dato** | «Error! Vínculo no válido» | lo reporta como huérfano |
| **La tabla del estado** | pega el rango tal cual | formato por tipo de fila (H/I/S/T/N) |
| **Actualizar** | F9, o al abrir | al refrescar |
| **Depende de** | nada | el `.exe` |

### Recomendación

**Use el vínculo nativo** para algo puntual y suyo: un dato que quiere citar
hoy en un documento que no va a salir de su equipo, sin pedirle nada a
nadie. Es rápido y no depende de este programa.

**Use este programa** para el documento de cierre: el que se comparte, el
que otra persona abre desde su OneDrive, y el que tiene que aguantar que el
modelo de Excel cambie de estructura de un trimestre a otro.

El riesgo del vínculo nativo no es que se rompa —eso se ve—, es que **siga
funcionando enseñando el número de al lado**.

### Si aun así quiere vínculos nativos, protéjase

1. **No inserte ni borre filas** en la zona vinculada del Excel. Si hay que
   hacerlo, revise después cada vínculo del Word.
2. Deje el Excel y el Word **en la misma carpeta**, y no los mueva.
3. Tras cada actualización, **compare un par de cifras** contra el Excel
   antes de dar el documento por bueno.
4. Cuando el documento esté cerrado, **rompa los vínculos**
   (Ctrl+Mayús+F9) en la copia que va a circular, para que nadie la
   actualice contra un Excel que ya cambió.
