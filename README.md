# Generador de Estado de Situación Financiera

Convierte la hoja del **Estado de Situación Financiera** de tu Excel en un
documento de **Word** con el formato de la plantilla ya aplicado.

---

## Descargar y usar

**Un solo archivo. No hay que instalar nada.** (Windows, ~13 MB)

### ⬇ [Descargar GeneradorFS.exe](https://github.com/RadioFics/Excel_Word_Paloma/releases/latest/download/GeneradorFS.exe)

Luego:

1. **Arrastra tu archivo de Excel encima de `GeneradorFS.exe`.**
   (O deja el Excel en la misma carpeta que el `.exe` y haz doble clic en el `.exe`.)
2. Se abre una ventana, trabaja unos segundos y muestra
   **DOCUMENTO GENERADO CORRECTAMENTE**. Pulsa Enter para cerrarla.
3. El Word queda en una carpeta **`salidas`**, junto al `.exe`.

Eso es todo. No se instala Python, no se descomprime nada, no hace falta
permiso de administrador.

### La primera vez, Windows puede advertir

Si aparece la pantalla azul **«Windows protegió su PC»**:
**Más información → Ejecutar de todas formas.**

Es lo normal con cualquier programa recién descargado que todavía no tiene
firma digital. Abajo se explica qué contiene y cómo comprobar que es el
archivo legítimo.

---

## Qué obtienes

En la carpeta `salidas`, junto al `.exe`:

| Archivo | Qué es |
|---|---|
| `estado_situacion_financiera_<fecha>.docx` | El documento de Word con las cifras puestas. Los anteriores no se borran. |
| `revisar_tipos.csv` | Lista, fila por fila, qué se trasladó y cómo se clasificó cada renglón (encabezado, detalle, subtotal, total, nota). Ábrelo con Excel para revisar. |

Si tu hoja tiene una columna **`Tipo`** (con letras H/I/S/T/N/X), el programa
la respeta. Si no la tiene, deduce solo el papel de cada fila y lo apunta en
`revisar_tipos.csv` para que lo revises.

---

## Si algo sale mal

La ventana explica el problema en español y qué revisar. Los casos más
comunes:

| Mensaje | Qué hacer |
|---|---|
| «No encontré ningún Excel en esta carpeta» | Arrastra el Excel sobre el `.exe`, o cópialo a la misma carpeta. |
| «No pude identificar la hoja…» | Tu libro no tiene una hoja reconocible como Estado de Situación Financiera. |
| El Word sale con cifras en blanco | El Excel se guardó con una herramienta que no recalcula. Ábrelo y guárdalo desde Excel, y repite. |

---

## ¿Es seguro? ¿Es un virus?

No. Los avisos de Windows aparecen porque el programa **no está firmado
digitalmente** (la firma requiere un certificado de pago; es un paso
posterior). Windows advierte de **todo** archivo descargado que no reconoce,
sin mirar su contenido.

- **Qué hace:** lee un archivo `.xlsx` de tu equipo y escribe un `.docx` en
  tu equipo.
- **Qué NO hace:** no se instala, no toca la configuración de Windows, no se
  inicia solo, **no se conecta a internet**, no envía datos a ningún lado.
- **Cómo verificar que es el archivo correcto:** compara su huella SHA-256
  con la que se publica junto a la descarga:
  ```
  powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '.\GeneradorFS.exe').Hash"
  ```
  Huella de la versión actual:
  ```
  61C57FFE7AE31A6C4A864E26CD6F0B56E7F831224B5F2DD63D893AE7465145E6
  ```
  (Cambia cada vez que se recompila; usa siempre la que acompañe a la descarga.)
- **Origen:** repositorio `RadioFics/Excel_Word_Paloma`.

Si tu equipo lo **bloquea por completo** (política corporativa / antivirus
que lo pone en cuarentena sin opción de continuar): no lo fuerces.
Repórtalo — es uno de los datos que esta prueba busca confirmar.

---
---

## Para desarrollo (no necesario para usar el `.exe`)

El `.exe` se genera desde `generador_fs.py` (Python). Para trabajar el código:

| Documento | Contenido |
|---|---|
| [`INSTALACION.md`](INSTALACION.md) | Montar Python portable en `.\python\` sin permisos de administrador. |
| [`PRUEBA_EXTERNA.md`](PRUEBA_EXTERNA.md) | Reproducir la prueba en otro equipo (con el `.exe` o clonando). |
| [`DIRECCION.md`](DIRECCION.md) | Dirección del proyecto, alternativas y caso ante TI. |
| [`addin/`](addin/) | Add-in de Word (v0) que actualiza el mismo documento en vez de generar uno nuevo. |

Comandos (equipo con el entorno montado):

```
powershell -ExecutionPolicy Bypass -File .\tools\bootstrap_python.ps1   # monta .\python\
powershell -ExecutionPolicy Bypass -File .\tools\verificar.ps1          # prueba de humo
powershell -ExecutionPolicy Bypass -File .\tools\hacer_exe.ps1          # genera dist\GeneradorFS.exe
```

### Ajustes opcionales (`config.json`)

Va **embebido** en el `.exe` con valores por defecto. Para cambiarlos sin
recompilar, deja un `config.json` propio en la misma carpeta que el `.exe`:

| Clave | Para qué |
|---|---|
| `empresa` | Nombre de la empresa (no viene de la hoja). |
| `hoja` | Nombre exacto de la hoja. Si no existe, se elige por contenido. |
| `hoja_marcadores` | Textos que delatan un estado de situación financiera. |
| `plantilla` | Nombre de un `.docx` de plantilla alternativo (en la carpeta del `.exe`). |
| `primera_fila` | Un número fija la fila de inicio; `"auto"` la detecta. |
| `columnas` | Letras (`A`, `C`, `E`, `F`, `G`) para forzar una columna; `null` = detectar. |
| `marcadores_excluir` | Etiquetas que marcan una fila de control/cuadre. |

### Cómo clasifica cada fila

| Letra | Significa | Cómo se deduce si no hay columna `Tipo` |
|---|---|---|
| `H` | Encabezado de sección | etiqueta sin cifras, en mayúsculas / termina en ":" / en negrita |
| `I` | Línea de detalle | etiqueta + cifras (o + número de nota) |
| `S` | Subtotal (sin etiqueta) | cifras sin etiqueta |
| `T` | Total (con etiqueta) | la etiqueta empieza por "Total" |
| `N` | Nota de texto libre | texto largo sin cifras |
| `X` | Excluir a propósito | la etiqueta o la nota contienen "control check" / "cuadre" |

### Límites

- Un solo tipo de estado (Situación Financiera). Otro estado = otra plantilla.
- No traslada cifras sueltas dentro de párrafos; solo tablas.
- La deducción de tipos es heurística: por eso está `revisar_tipos.csv`.
- Genera un Word **nuevo** cada vez. Actualizar *el mismo* documento en su
  sitio es el add-in (`addin/`).
- Solo Windows de 64 bits.

### Trazabilidad

Cada Word generado guarda en sus propiedades (Archivo → Información →
Propiedades → Comentarios) el nombre del Excel de origen, su huella SHA-256,
la hoja usada y la fecha.
