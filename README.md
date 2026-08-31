# Generador de Estado de Situación Financiera (piloto)

Traslada el Estado de Situación Financiera de un libro de Excel a una
plantilla de Word. Cambias el Excel, corres el programa, obtienes un Word
actualizado. La plantilla nunca se toca; el Excel nunca se modifica.

Identifica la hoja y las columnas por su **contenido**, no por nombres ni
posiciones fijas. La columna `Tipo` es **opcional**: si falta, infiere el
tipo de cada fila y deja un CSV de revisión.

---

# Descargar y probar (usuario final, equipo externo)

Para comprobar que funciona en otro computador **sin instalar nada** y sin
conocimientos técnicos. Solo Windows de 64 bits.

## 1. Descargar

Descarga **`GeneradorFS_portable_AAAAMMDD.zip`** (~21 MB) desde:

- la página **Releases** del repositorio
  (`https://github.com/RadioFics/Excel_Word_Paloma/releases`), **o**
- el enlace de **OneDrive** que te compartan.

Contiene todo lo necesario, incluida una copia de Python. No hay instalador.

Si el navegador dice *"…no se descarga habitualmente"*, elige
**Conservar / Mantener**.

## 2. Comprobar que es el archivo correcto (recomendado)

Abre PowerShell en la carpeta de descargas y ejecuta:

```
powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '.\GeneradorFS_portable_20260831.zip').Hash"
```

El resultado debe coincidir, carácter por carácter, con el hash que publique
quien te compartió el archivo. Hash de la compilación del 2026-08-31:

```
4E48FD857919292C619994812B806FE1CEE172B38F7B5EB6D9467B8DA7E66074
```

Si coincide, es exactamente ese archivo, sin modificar. (Cada vez que se
reconstruye el `.zip` el hash cambia; usa siempre el que acompañe a la
descarga.)

## 3. Desbloquear y extraer

1. Clic derecho en el `.zip` → **Propiedades**.
2. Si abajo aparece *"Este archivo proviene de otro equipo…"*, marca
   **Desbloquear** → **Aceptar**. (Esto evita los avisos en todos los
   archivos de dentro.)
3. Clic derecho en el `.zip` → **Extraer todo…** → elige una carpeta normal
   (Escritorio, Documentos). **No** dentro de `C:\Archivos de programa`.
4. **No ejecutes nada desde dentro del `.zip`**; primero extrae.

## 4. Ejecutar

- **Comprobación automática:** entra en la carpeta `tools`, clic derecho en
  `verificar.ps1` → **Ejecutar con PowerShell**. Debe terminar en verde:
  `TODO CORRECTO. La aplicacion funciona en este equipo.`
  - Si dice *"la ejecución de scripts está deshabilitada"*, ábrelo así (una
    sola vez, no cambia la configuración del equipo):
    ```
    powershell -ExecutionPolicy Bypass -File .\tools\verificar.ps1
    ```
- **Prueba real:** arrastra tu Excel sobre `generar.bat`, o haz doble clic
  en `generar.bat` sin arrastrar nada (usa el `Copia_Editable_con_columna_Tipo.xlsx`
  incluido).
- Si aparece la pantalla azul **"Windows protegió su PC"**:
  **Más información** → **Ejecutar de todas formas**.

## 5. Resultado esperado

En la carpeta `salidas` aparecen:

- `estado_situacion_financiera_*.docx` — el documento generado.
- `revisar_tipos.csv` — qué se trasladó, con qué tipo y por qué.

Con el Excel de ejemplo: **33 líneas**, hoja `FS`, columnas
`etiqueta=A nota=C actual=E previo=F tipo=G`, región filas 5–47.

---

## ¿Por qué Windows me advierte? ¿Es un virus?

No. Los avisos aparecen porque el paquete **no está firmado digitalmente**
(firmar código requiere un certificado de pago; es un paso posterior).
Hasta entonces, Windows advierte de **cualquier** archivo descargado que no
reconoce, sin mirar su contenido.

Qué hay dentro del `.zip`:

- Una copia de **Python** (el lenguaje, tal como se baja de python.org).
- La **plantilla de Word** y un **Excel de ejemplo**.
- `generador_fs.py`: **texto plano** que puedes abrir con el Bloc de notas y
  leer entero.

Lo que **no** hace: no hay instalador, no toca el registro de Windows, no se
arranca solo con el sistema, no se conecta a internet, no envía datos a
ningún lado. Lee un `.xlsx` de tu equipo y escribe un `.docx` en tu equipo.

Cómo asegurarte de que es el archivo legítimo: verifica el **SHA-256**
(paso 2). Origen: repositorio `RadioFics/Excel_Word_Paloma`.

**Si tu equipo lo bloquea por completo** (política corporativa, AppLocker,
antivirus que lo pone en cuarentena y no da opción de continuar): **no
insistas ni lo fuerces**. Anótalo y repórtalo — ese es precisamente uno de
los datos que esta prueba busca confirmar.

---

## Cómo usarlo en el día a día

1. Arrastra el Excel sobre `generar.bat`, o doble clic sin arrastrar nada
   (toma el Excel más reciente de la carpeta cuyo nombre contenga "FS").
2. Se abre una ventana, muestra el resultado y espera a que pulses Enter.
3. El documento nuevo queda en `salidas`, con fecha y hora en el nombre.
   Los anteriores no se borran ni se sobrescriben.
4. **Revisa `salidas\revisar_tipos.csv`**: fila por fila, qué se trasladó,
   con qué tipo, y si fue *declarado* (venía en la hoja) o *inferido* (lo
   dedujo el programa). Si algo no cuadra, corrígelo en el Excel y repite.

Para ver solo esa revisión sin generar el Word, en una terminal en la
carpeta:

```
python\python.exe generador_fs.py "tu.xlsx" --revisar
```

## La columna "Tipo" (opcional)

Si la hoja tiene una columna cuyo encabezado es `Tipo` (o `Type`), el
programa usa esas letras. Si no la tiene, **infiere** el tipo de cada fila y
lo anota en `revisar_tipos.csv` con la señal que lo decidió.

| Letra | Significa | Cómo se infiere si falta |
|---|---|---|
| `H` | Encabezado de sección | etiqueta sin cifras, en mayúsculas / termina en ":" / en negrita |
| `I` | Línea de detalle | etiqueta + cifras (o + número de nota) |
| `S` | Subtotal (sin etiqueta) | cifras sin etiqueta |
| `T` | Total (con etiqueta) | la etiqueta empieza por "Total" |
| `N` | Nota de texto libre | texto largo sin cifras |
| `X` | Excluir a propósito | la etiqueta o la nota contienen "control check" / "cuadre" (configurable) |

La inferencia acierta en la práctica sobre este estado, pero **no es
infalible**: por eso está el CSV de revisión. Para dejar el mapeo fijo y sin
ambigüedad, añade la columna `Tipo` a la hoja.

## Ajustes (`config.json`)

Todo es opcional; borra una clave para usar su valor por defecto.

| Clave | Para qué |
|---|---|
| `empresa` | Nombre de la empresa (no viene de la hoja). |
| `hoja` | Nombre exacto de la hoja. Si no existe, se elige por contenido. |
| `hoja_marcadores` | Textos que delatan un estado de situación financiera. |
| `plantilla` | Nombre del `.docx` de plantilla. |
| `buscar_por_convencion` | Texto que debe contener el nombre del Excel al buscar sin argumentos. |
| `primera_fila` | Un número fija la fila de inicio; `"auto"` la detecta. |
| `columnas` | Letras (`A`, `C`, `E`, `F`, `G`) para forzar una columna; `null` = detectar. |
| `marcadores_excluir` | Etiquetas que marcan una fila de control/cuadre. |

## Advertencia operativa — edición del Excel

Abre y guarda el libro **siempre desde Excel o LibreOffice**. **Nunca**
reguardes este libro con un script de `openpyxl`: no recalcula fórmulas y,
al reguardar, descarta el valor cacheado de **todas** las fórmulas. El
síntoma es que el programa corre sin error pero el Word sale con las cifras
en blanco.

---

# Para desarrollo o equipo propio (con permisos)

## Estructura de la carpeta

```
generador_fs.py                              <- el programa
config.json                                  <- ajustes (todo opcional)
plantilla_estado_situacion_financiera.docx   <- la plantilla (no editar el formato)
generar.bat                                  <- lanzador de doble clic (Windows)
requirements.txt                             <- dependencias exactas
tools/bootstrap_python.ps1                   <- monta Python portable en .\python\
tools/verificar.ps1                          <- prueba de humo del entorno
tools/hacer_paquete.ps1                      <- crea el .zip portable (usuario final)
INSTALACION.md                               <- Python sin permisos de admin
PRUEBA_EXTERNA.md                            <- guía de la prueba en otro equipo
DIRECCION.md                                 <- dirección del proyecto y caso ante TI
addin/                                       <- add-in de Word (v0, actualiza in situ)
Copia_Editable_con_columna_Tipo.xlsx         <- ejemplo de Excel ya etiquetado
salidas/                                     <- documentos generados + revisar_tipos.csv
```

## Montar el entorno (una vez, necesita internet)

Sin permisos de administrador. Ver `INSTALACION.md`. En resumen:

```
powershell -ExecutionPolicy Bypass -File .\tools\bootstrap_python.ps1
```

Deja un Python portable en `.\python\`. `generar.bat` lo usa solo.

## Producir el `.zip` para el usuario final

```
powershell -ExecutionPolicy Bypass -File .\tools\hacer_paquete.ps1
```

Crea `dist\GeneradorFS_portable_AAAAMMDD.zip` (incluye `.\python\`). Publícalo
como *release* del repositorio o en OneDrive, **junto con su hash SHA-256**:

```
powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '.\dist\GeneradorFS_portable_AAAAMMDD.zip').Hash"
```

## Límites y alcance de esta versión

- Cubre un solo tipo de estado (Situación Financiera). Otro estado (Estado
  de Resultados, Flujo de Efectivo) necesita su propia plantilla y sus
  marcadores, siguiendo el mismo patrón.
- No traslada valores sueltos dentro de párrafos narrativos (solo tablas).
- La inferencia de tipos es heurística: revisa `revisar_tipos.csv`.
- Genera un documento **nuevo** cada vez. Para actualizar *el mismo*
  documento en su sitio, ver el add-in (`addin/`) y el frente C de `DIRECCION.md`.
- No hay pruebas automatizadas todavía.
- Solo Windows x64 para el Python portable. Ver límites de gobernanza (TI)
  en `INSTALACION.md` y `DIRECCION.md`.

## Trazabilidad

Cada documento generado guarda, en sus propiedades de Word (Archivo →
Información → Propiedades → Comentarios), el nombre del Excel de origen, un
identificador único de su contenido (SHA-256), la hoja usada y la fecha de
generación.
