# Prueba externa — replicar el proceso en otro equipo

Objetivo: que una persona en otro computador reproduzca la generación del
Word desde el Excel **sin instalar nada** (o, como mucho, ejecutando un
único script de descarga), y compruebe que funciona igual.

Hay dos formas. La **A** no necesita internet en el equipo destino.

---

## Forma A — con el paquete portable (recomendada para la prueba)

Nada que instalar. Requiere solo Windows de 64 bits.

**Quien prepara el paquete (una vez, en el equipo que ya funciona):**

```
powershell -ExecutionPolicy Bypass -File .\tools\hacer_paquete.ps1
```

Genera `dist\GeneradorFS_portable_AAAAMMDD.zip` (~50 MB, incluye Python).
Súbelo a OneDrive o como *release* del repositorio.

**Quien hace la prueba (en el equipo destino):**

1. Descarga el `.zip` y descomprímelo en una carpeta cualquiera
   (Escritorio, Documentos…). **No** dentro de `C:\Archivos de programa`.
2. Comprobación automática — clic derecho en `tools\verificar.ps1` →
   *Ejecutar con PowerShell*. Debe terminar en verde:
   `TODO CORRECTO. La aplicacion funciona en este equipo.`
3. Prueba real — arrastra un Excel sobre `generar.bat`, o haz doble clic
   sin arrastrar nada (toma el `Copia_Editable_con_columna_Tipo.xlsx`
   incluido).
4. Revisa el resultado en `salidas\`:
   - `estado_situacion_financiera_*.docx` — el documento generado.
   - `revisar_tipos.csv` — qué se trasladó y con qué tipo.

**Resultado esperado con el Excel de ejemplo:** 33 líneas, hoja `FS`,
columnas `etiqueta=A nota=C actual=E previo=F tipo=G`, región filas 5–47.

---

## Forma B — clonando el repositorio (necesita internet una vez)

Para quien va a iterar sobre el código.

1. `git clone <url-del-repo>` y entra en la carpeta.
2. Monta el Python portable (descarga ~30 MB de python.org y PyPI):
   ```
   powershell -ExecutionPolicy Bypass -File .\tools\bootstrap_python.ps1
   ```
3. `tools\verificar.ps1` → debe salir verde.
4. Uso: `generar.bat`, o `python\python.exe generador_fs.py "tu.xlsx"`.

El repositorio **no** incluye la carpeta `python\` (por tamaño); la crea el
paso 2. El resto es idéntico a la Forma A.

---

## Qué se está probando

| Punto | Cómo se comprueba |
|---|---|
| No hace falta instalar Python | Forma A: no se instala nada. Forma B: un solo script, sin permisos de admin. |
| Portabilidad entre equipos | El mismo `.zip`/repo corre en otra máquina Windows x64. |
| Funciona "tal cual" | `verificar.ps1` en verde + `salidas\*.docx` con 33 líneas sobre el ejemplo. |
| Detección por contenido | Cambia el nombre de la hoja en `config.json` a uno inexistente y vuelve a correr: debe encontrar `FS` "por contenido". |
| Tolera Excel sin columna Tipo | Pásale un Excel sin columna `Tipo`: genera igual e infiere los tipos (ver `revisar_tipos.csv`). |

## Límites conocidos de la prueba

- Solo Windows x64 (el Python portable no es multiplataforma).
- Si el equipo destino aplica **AppLocker / WDAC / control de aplicaciones**,
  puede bloquear `python.exe` desde una carpeta de usuario aunque no haya
  instalación. Es exactamente el dato que la prueba busca confirmar con TI.
- Genera un documento **nuevo** cada vez. La actualización *in situ* del
  mismo documento es el add-in (`addin\`, ver `DIRECCION.md`).
