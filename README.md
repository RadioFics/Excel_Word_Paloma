# Estados Financieros — Excel a Word

Mantiene al día un documento de Word con las cifras de tu Excel, **sin
borrar lo que hayas escrito**.

### ⬇ [Descargar EstadosFinancieros.exe](https://github.com/RadioFics/Excel_Word_Paloma/releases/latest/download/EstadosFinancieros.exe)

Windows 64 bits · ~13 MB · no se instala nada

---

## Para qué sirve

Cada cierre hay que pasar las cifras del modelo de Excel al documento de
Word. Hacerlo a mano es lento y se cuelan erratas; regenerar el Word entero
borra los comentarios y el análisis que alguien escribió.

Este programa **refresca solo las cifras**. La redacción no se toca.

```
   Escribes un párrafo  →  refrescas  →  el párrafo sigue ahí, cifras al día
   Borras ese párrafo   →  refrescas  →  sigue borrado
```

Funciona así porque solo escribe dentro de unas **regiones marcadas** del
documento. Todo lo demás ni lo visita.

---

## Qué hace

Arrastra tu Excel sobre el `.exe` (o haz doble clic) y elige:

| | Opción | Qué hace |
|---|---|---|
| **1** | Actualizar el documento de siempre | Refresca las cifras. Conserva toda la redacción. |
| **2** | Crear un documento nuevo | Un Word desde cero en `salidas\`. Para una entrega puntual. |
| **3** | Cambiar el documento que se actualiza | Elige otro documento de Word desde el explorador. |
| **4** | Permitir editar las cifras a mano | Desbloquea temporalmente. |
| **5** | Volver a proteger las cifras | Las vuelve intocables. |

Además:

- **Cifras dentro de la redacción.** Puedes escribir *«los activos totales
  ascendieron a 119.066.301»* y que ese número siga al Excel, leyéndose como
  texto normal. Ver [`docs/DATOS.md`](docs/DATOS.md).
- **Las cifras van bloqueadas.** Nadie puede teclear encima de un número que
  el Excel contradiga.
- **Registro de cambios** en un `.log` aparte, no dentro del documento.
- **Aguanta que renombres filas** en el Excel, si usas rangos con nombre.

---

## Empezar

**1 · Descarga el `.exe`** y déjalo en una carpeta.

**2 · Dile qué documento actualizar.** Ejecútalo y elige la opción **3**:
se abre el explorador, eliges tu `.docx` y queda guardado. (También puedes
escribirlo a mano en un `config.json` junto al `.exe`.)

**3 · Prepáralo la primera vez:**

```bash
EstadosFinancieros.exe MI_LIBRO.xlsx --refrescar --preparar
```

**4 · A partir de ahí**, cada cierre: arrastra el Excel sobre el `.exe` y
elige la opción **1**.

> **Cierra el documento en Word antes de refrescar.** Si está abierto, el
> programa se detiene y te dice quién lo retiene.

La guía completa está en **[`docs/GUIA.md`](docs/GUIA.md)**.

---

## La primera vez, Windows puede advertir

Si aparece **«Windows protegió su PC»**: *Más información → Ejecutar de
todas formas*. Es lo normal con cualquier programa recién descargado que
todavía no tiene firma digital.

**Qué hace:** lee un `.xlsx` de tu equipo y escribe un `.docx` de tu equipo.
**Qué no hace:** no se instala, no toca la configuración de Windows, **no se
conecta a internet**, no envía datos a ningún lado.

Para comprobar que descargaste el archivo legítimo, compara su huella con la
publicada junto a la descarga:

```
powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '.\EstadosFinancieros.exe').Hash"
```

---

## Documentación

| Documento | Contenido |
|---|---|
| [`docs/GUIA.md`](docs/GUIA.md) | **Guía de operación paso a paso.** Empieza por aquí. |
| [`docs/DATOS.md`](docs/DATOS.md) | Dónde se editan las cifras, y cómo llevarlas a la redacción. |
| [`docs/CAMBIAR_EXCEL.md`](docs/CAMBIAR_EXCEL.md) | Qué revisar cuando llegue un Excel nuevo. |
| [`docs/VINCULOS_NATIVOS.md`](docs/VINCULOS_NATIVOS.md) | Vincular Excel y Word **sin programar**, y qué se rompe. |
| [`docs/CONTRATO.md`](docs/CONTRATO.md) | Especificación de las regiones. Para quien toque el código. |
| [`docs/ESTRUCTURA.md`](docs/ESTRUCTURA.md) | Organización de las carpetas. |
| [`docs/DESPLIEGUE_ADDIN.md`](docs/DESPLIEGUE_ADDIN.md) | Subir el complemento de Word y qué pedirle a TI. |
| [`docs/INSTALACION.md`](docs/INSTALACION.md) | Montar el entorno de desarrollo. |
| [`docs/DIRECCION.md`](docs/DIRECCION.md) | Dirección del proyecto y caso ante TI. |

---
---

# Detalle técnico

Lo de abajo no hace falta para usar el programa.

## Los tres ejecutables

`EstadosFinancieros.exe` reúne los dos caminos, pero cada uno existe también
por separado si prefieres un icono por tarea:

| Ejecutable | Qué hace |
|---|---|
| **`EstadosFinancieros.exe`** | Pregunta cuál de los dos |
| `GeneradorFS.exe` | Word **nuevo** en `salidas\` |
| `RefrescarFS.exe` | **Actualiza** el documento base |

Los tres admiten que se les arrastre el Excel encima.

### Sin menú, para automatizar

```
EstadosFinancieros.exe MI_LIBRO.xlsx --refrescar
EstadosFinancieros.exe MI_LIBRO.xlsx --generar
EstadosFinancieros.exe --documento        elegir documento
EstadosFinancieros.exe --desbloquear
EstadosFinancieros.exe --bloquear
EstadosFinancieros.exe --estado           diagnostico
EstadosFinancieros.exe --consola          menu de texto, sin ventana
```

## Estructura del repositorio

```
EstadosFinancieros.exe  ->  src/fs_menu.py        (menu y ventana)
generar.bat             ->  src/generador_fs.py   (Word nuevo)
refrescar.bat           ->  src/refrescar_fs.py   (actualiza el base)
config.json                 ajustes

src/          codigo Python (motor y ordenes)
docs/         documentacion
plantillas/   documentos .docx modelo
ejemplos/     libro .xlsx de muestra
tools/        entorno, empaquetado y verificacion
addin/        complemento de Word (TypeScript)
salidas/      lo que genera generar.bat, y el .log de cambios
```

Las carpetas se agrupan por **naturaleza**, no por extensión. Ver
[`docs/ESTRUCTURA.md`](docs/ESTRUCTURA.md).

## Ajustes (`config.json`)

Va embebido en el `.exe` con valores por defecto. Para cambiarlos sin
recompilar, deja un `config.json` propio junto al `.exe`:

| Clave | Para qué |
|---|---|
| `documento_base` | **El .docx que actualiza la opción 1.** |
| `empresa` | Nombre de la empresa (no viene de la hoja). |
| `hoja` | Nombre exacto de la hoja. Si no existe, se elige por contenido. |
| `hoja_marcadores` | Textos que delatan un estado de situación financiera. |
| `plantilla` | Un `.docx` de plantilla alternativo. |
| `primera_fila` | Un número fija la fila de inicio; `"auto"` la detecta. |
| `columnas` | Letras (`A`, `C`, `E`…) para forzar una columna; `null` = detectar. |
| `marcadores_excluir` | Etiquetas que marcan una fila de control/cuadre. |
| `prefijo_rangos` | Prefijo de los nombres de Excel que fijan la identidad (`fs_`). |
| `bitacora` | `"archivo"` (por defecto), `"documento"`, `"ambos"` o `"no"`. |
| `bitacora_archivo` | Ruta del `.log`. Vacío = `salidas\bitacora_<documento>.log`. |
| `apariencia_datos` | `"boundingBox"` (recuadro) o `"hidden"` (se lee como texto normal). |

## Órdenes de `src/fs_documento.py`

| Orden | Para qué |
|---|---|
| `construir` / `reparar` | Añade las regiones que falten. No destruye nada. |
| `refrescar` | Actualiza las regiones de datos desde el Excel. |
| `insertar` | Coloca una cifra viva en la redacción. |
| `catalogo` | Lista las claves disponibles en el Excel. |
| `nombrar` | Crea los rangos con nombre en el Excel. |
| `verificar` | Revisa el documento: regiones huérfanas, vacías. |
| `estado` | Radiografía del proyecto. |
| `plantilla` | Genera un documento base nuevo desde cero. |
| `bloquear` / `desbloquear` | Impide (o permite) teclear las cifras a mano. |
| `desvincular` | Convierte una cifra en texto normal: deja de refrescarse. |
| `apariencia` | `visible` (recuadro) o `invisible` (texto corrido). |
| `simplificar` | Quita los recuadros que estorban al redactar. |
| `limpiar-bitacora` | Retira del documento la bitácora incrustada. |
| `proteger` / `desproteger` | Modo estricto de dos editores. |

## Cómo clasifica cada fila

| Letra | Significa | Cómo se deduce si no hay columna `Tipo` |
|---|---|---|
| `H` | Encabezado de sección | etiqueta sin cifras, en mayúsculas / termina en ":" / en negrita |
| `I` | Línea de detalle | etiqueta + cifras (o + número de nota) |
| `S` | Subtotal | cifras sin etiqueta |
| `T` | Total | la etiqueta empieza por "Total" |
| `N` | Nota de texto libre | texto largo sin cifras |
| `X` | Excluir a propósito | la etiqueta contiene "control check" / "cuadre" |

## Desarrollo

```
powershell -ExecutionPolicy Bypass -File .\tools\bootstrap_python.ps1   # monta .\python\
powershell -ExecutionPolicy Bypass -File .\tools\verificar.ps1          # prueba de humo
powershell -ExecutionPolicy Bypass -File .\tools\hacer_exe.ps1          # genera los .exe
```

### Publicar una versión para que se pueda descargar

El enlace de descarga de arriba apunta a `releases/latest/download/`, así que
**empieza a funcionar solo** en cuanto publiques la primera release. No hay
que tocar el README en cada versión.

1. `tools\hacer_exe.ps1` → deja los tres `.exe` en `dist\` y escribe sus
   huellas SHA-256 en pantalla.
2. En GitHub, en la página del repositorio: **Releases** (columna derecha) →
   **Draft a new release**.
3. **Choose a tag** → escribe una etiqueta nueva, p. ej. `v0.4`, y elige
   *Create new tag on publish*.
4. Pon un título y, en el cuerpo, pega las huellas SHA-256.
5. Arrastra los tres `.exe` de `dist\` a la caja *Attach binaries by
   dropping them here*.
6. **Publish release.**

El repositorio tiene que ser **público** para que el enlace funcione sin
iniciar sesión. Si es privado, quien descargue necesitará permisos.

## Límites

- Un solo tipo de estado (Situación Financiera). Otro estado = otra plantilla.
- La deducción de tipos es heurística: por eso está `revisar_tipos.csv`.
- Las cifras sueltas dentro de párrafos hay que marcarlas una vez; no se
  detectan solas.
- El complemento de Word (`addin/`) está escrito pero **sin compilar**.
- La coautoría de Word en el navegador maneja mal los controles de
  contenido: refresca desde Word de escritorio.
- Solo Windows de 64 bits.
