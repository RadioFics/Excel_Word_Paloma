# Configuración de Python para este proyecto

Sigue las restricciones del equipo: **sin permisos de administrador, sin
instalador MSI, sin `winget`, sin Microsoft Store**. Usa el *paquete
embebido* de Python, que es solo archivos descomprimidos dentro de la
carpeta del proyecto — no toca el registro ni el sistema.

## Instalación (una vez por equipo)

1. Abre la carpeta del proyecto en el Explorador.
2. Clic derecho en `tools\bootstrap_python.ps1` → **Ejecutar con PowerShell**.
   - Si Windows lo bloquea, abre PowerShell en la carpeta y ejecuta:
     ```
     powershell -ExecutionPolicy Bypass -File .\tools\bootstrap_python.ps1
     ```
3. El script descarga Python 3.12.10 embebido + pip + las dependencias de
   `requirements.txt` y lo deja todo en `.\python\`. Necesita internet solo
   esta vez (python.org, bootstrap.pypa.io, pypi.org).

Al terminar debe verse `Python 3.12.10` y la lista de paquetes
(`docxtpl`, `openpyxl`, …).

## Uso

- Doble clic en `generar.bat`, o arrastra el Excel sobre él.
- `generar.bat` usa `.\python\python.exe` automáticamente si existe; si no,
  cae al `python` del PATH.

## Llevarlo a otro computador

**Con internet:** copia el proyecto **sin** la carpeta `python\` y ejecuta
otra vez `tools\bootstrap_python.ps1` en el equipo nuevo.

**Sin internet:** copia la carpeta del proyecto **completa, incluyendo
`python\`**. Es portable entre equipos **Windows de 64 bits**. No requiere
instalación ni permisos.

## Límites de esta vía (leer antes de comprometerla con TI)

- **Solo Windows x64.** El paquete embebido no sirve en macOS/Linux ni en
  Windows ARM. Para este piloto (2 usuarias, equipos Windows) es suficiente.
- **AppLocker / WDAC.** Si el equipo aplica control de aplicaciones, puede
  bloquear la ejecución de `python.exe` desde una carpeta de usuario o una
  unidad de red aunque no haya "instalación". **Confirmar con TI** antes de
  depender de esto.
- **Antivirus / EDR.** Un ejecutable nuevo en una ruta compartida puede
  generar alertas. Conviene avisar a seguridad y, si se puede, alojar el
  proyecto en una ruta de solo lectura para las usuarias.
- **Sin `tkinter`, sin `venv`.** El paquete embebido no incluye la interfaz
  gráfica ni entornos virtuales. Si más adelante se quiere un botón gráfico,
  habrá que empaquetar con PyInstaller (otro camino, mismo dilema de firma).
- **Actualizaciones de dependencias.** Se hacen volviendo a correr el
  bootstrap o `.\python\python.exe -m pip install -U ...`. Queda registro en
  `pip list`.

## Por qué esto no resuelve la portabilidad "para siempre"

Python portable resuelve la portabilidad **técnica** para las 2 usuarias en
sus equipos Windows. **No** resuelve la portabilidad de **gobernanza**: en
una flota gestionada, cualquier equipo nuevo puede tener AppLocker/EDR que
frene un intérprete no firmado. La respuesta permanente para "que funcione
en cualquier computador de la organización sin fricción" es el **add-in de
Office** (no instala nada, corre donde corra M365). Ver `DIRECCION.md`.
