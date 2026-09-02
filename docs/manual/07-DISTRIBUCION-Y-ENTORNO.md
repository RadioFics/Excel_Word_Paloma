# 07 · Distribución, empaquetado y entorno

> **Para quién.** Quien monte el proyecto en un equipo nuevo, reconstruya
> los ejecutables, publique una versión o tenga que justificar ante TI qué
> se instala exactamente.
> **Qué encontrará.** La cadena completa desde el repositorio hasta el
> archivo que descarga la usuaria: el Python embebido, los lanzadores
> `.bat`, los tres `.exe`, el paquete portable, la publicación de una
> versión y las reglas de higiene del repositorio.
> **Antes de leer.** Basta con saber que el programa lee un `.xlsx` y
> escribe un `.docx`. No hace falta conocer el contrato de anclas.

## Índice

- [1. El problema de fondo](#1-el-problema-de-fondo)
- [2. Montar el entorno de desarrollo](#2-montar-el-entorno-de-desarrollo)
- [3. Los lanzadores `.bat`](#3-los-lanzadores-bat)
- [4. Construir los ejecutables](#4-construir-los-ejecutables)
- [5. El paquete portable](#5-el-paquete-portable)
- [6. Los tres ejecutables](#6-los-tres-ejecutables)
- [7. Publicar una versión](#7-publicar-una-versión)
- [8. La advertencia de SmartScreen](#8-la-advertencia-de-smartscreen)
- [9. Higiene del repositorio](#9-higiene-del-repositorio)
- [Resumen del capítulo](#resumen-del-capítulo)

---

## 1. El problema de fondo

El proyecto no se distribuye como se distribuye normalmente un programa en
Python porque no puede. `docs/DIRECCION.md` lo dice sin rodeos:

> «Lo que no encaja en una organización que bloquea instaladores y revisa
> ejecutables es el **vehículo de entrega**: un intérprete/`.exe` no
> gestionado, sin firma, fuera de los controles del tenant.»

De ahí salen las dos vías de entrega, y ninguna instala nada:

```
   repositorio (código Python)
        │
        ├── tools\bootstrap_python.ps1 ──► .\python\  (intérprete embebido)
        │        │                              │
        │        │                              └── los cuatro .bat lo usan
        │        └── tools\hacer_paquete.ps1 ──► dist\GeneradorFS_portable_*.zip
        │                                            (código + Python, ~50 MB)
        └── tools\hacer_exe.ps1 ──► dist\EstadosFinancieros.exe
                 (PyInstaller)       dist\GeneradorFS.exe
                                     dist\RefrescarFS.exe
                                            └── release de GitHub ──► la usuaria
```

La vía del `.zip` lleva el intérprete dentro y necesita que el equipo
destino tolere ejecutar `python.exe` desde una carpeta de usuario. La vía
del `.exe` no necesita Python, pero es un binario sin firma digital: es la
que dispara la advertencia de la [sección 8](#8-la-advertencia-de-smartscreen).
Ninguna de las dos resuelve el fondo del asunto, y el repositorio no
pretende lo contrario: *«Python portable resuelve la portabilidad técnica
(…). No resuelve la portabilidad de gobernanza»* (`docs/INSTALACION.md`).

---

## 2. Montar el entorno de desarrollo

### 2.1 `tools/bootstrap_python.ps1`

Monta un intérprete dentro de la carpeta del proyecto *«sin permisos de
administrador y sin instalador MSI (no usa winget ni la Microsoft Store)»*
(`tools/bootstrap_python.ps1:3-5`). La versión está fijada en el código,
`$PyVersion = '3.12.10'` (`:22`), con el criterio al lado: *«El paquete
"embeddable" es solo archivos: se descomprime, no se instala, no toca el
registro ni el sistema.»* Ese es el argumento entero ante TI.

Es idempotente: si ya hay `.\python\python.exe`, imprime `Ya existe
.\python\python.exe -- nada que reinstalar.` y sale con código `0` (`:28-32`).

| Paso | Qué hace | Línea |
|---|---|---|
| 1 | Descarga `python-3.12.10-embed-amd64.zip` de python.org a `%TEMP%` | `:38-42` |
| 2 | `Expand-Archive -Force` sobre `.\python\`; borra el `.zip` | `:44-48` |
| 3 | Habilita `site-packages` en el archivo `python*._pth` | `:50-61` |
| 4 | Descarga `get-pip.py` de bootstrap.pypa.io, lo ejecuta y lo borra | `:63-68` |
| 5 | `pip install --no-cache-dir -r requirements.txt` | `:70-72` |

El paso 3 no se puede saltar:

> «El paquete embebido trae `import site` comentado y no busca en
> `Lib\site-packages`. Sin este ajuste, pip instala pero nada se importa.»

Sin ese arreglo el paso 5 terminaría en verde y el programa seguiría
fallando con `ModuleNotFoundError`. La descarga es **`amd64`
exclusivamente**: no hay variante para Windows ARM, macOS ni Linux.
Internet hace falta una sola vez y solo contra `python.org`,
`bootstrap.pypa.io` y `pypi.org`; en ejecución normal el programa no abre
ninguna conexión.

### 2.2 `requirements.txt`

```
docxtpl==0.20.2
openpyxl==3.1.5
```

Dos dependencias directas con versión exacta. `python-docx`, `jinja2`,
`lxml`, `et-xmlfile` y `MarkupSafe` entran como transitivas de `docxtpl`.
No incluye `pyinstaller`: `tools/hacer_exe.ps1:13` lo pide aparte.

### 2.3 `tools/buscar_python.bat`

Deja en la variable `PY` la ruta de un intérprete **usable**, o explica qué
falta y devuelve error. Existe por un motivo concreto:

> «Existe porque "si no esta el portable, usa python" no basta en Windows:
> el 'python' del PATH suele ser el ATAJO de la Microsoft Store, que no es
> Python. Imprime un aviso, devuelve 9009, y la ventana se cierra antes de
> que nadie lea nada.»

| # | Candidato | Cómo lo comprueba | Línea |
|---|---|---|---|
| 1 | `.\python\python.exe` del proyecto | `if exist` | `:18-21` |
| 2 | El lanzador `py -3` | `py -3 -c "import sys;print(sys.executable)"` | `:25-26` |
| 3 | El `python` del PATH | `python -c "import sys;print(sys.executable)"` | `:31-32` |

Los candidatos 2 y 3 no comprueban que el ejecutable responda: le preguntan
por `sys.executable`. *«El atajo de la Store falla aqui y deja PY vacia, en
vez de colarse»* (`:29-30`). Después, la etiqueta `:comprobar` valida las
dependencias con `"%PY%" -c "import openpyxl, docx, docxtpl"`.

| Código | Etiqueta | Mensaje |
|---|---|---|
| `0` | — | `PY` está listo |
| `1` | `:sin_python` | `NO HAY PYTHON EN ESTE EQUIPO`: usar el `.exe`, o correr `bootstrap_python.ps1` |
| `2` | `:sin_dependencias` | `FALTAN LAS DEPENDENCIAS`, con la ruta de `%PY%` y la orden de `pip` |

La comprobación mira `openpyxl`, `docx` y `docxtpl`, pero **no `jinja2`** —
que `tools/verificar.ps1:33` sí verifica. (Inferencia: `jinja2` llega
siempre como transitiva, así que la diferencia no se manifiesta; un entorno
instalado a mano y a medias pasaría aquí y fallaría después.)

---

## 3. Los lanzadores `.bat`

Los cuatro comparten esqueleto: `cd /d "%~dp0"`, llamada a
`buscar_python.bat` y, si este falla, `pause` para que el mensaje se pueda
leer antes de que la ventana se cierre.

| Archivo | Qué invoca | Argumentos |
|---|---|---|
| `estados_financieros.bat` | `src\fs_menu.py %*` | `%*` — todos, tal cual |
| `generar.bat` | `src\generador_fs.py` o `... "%~1"` | Solo el primero, entrecomillado |
| `refrescar.bat` | `src\refrescar_fs.py` o `... "%~1" %2 %3` | El primero entrecomillado; dos más en crudo |
| `probar.bat` | `tools\probar_refresco.py` o `... --libro "%~1"` | Convierte el arrastre en `--libro` |

Tres detalles que importan al mantenerlo:

- **`estados_financieros.bat` reenvía `%*` íntegro** (`:29`). Es lo que
  permite `--consola`, `--estado` o `--bloquear`. Los otros tres usan
  `if "%~1"==""` para distinguir el doble clic del arrastre.
- **`refrescar.bat` pasa `%2` y `%3` sin comillas** (`:35`): sirve para
  `--no-preparar` o `--documento X`, pero un argumento con espacios en esas
  posiciones se partiría. (Inferencia sobre el comportamiento de
  `cmd.exe`; no hay comentario del autor.)
- **`probar.bat` siempre pausa** (`:29-30`), mientras los otros solo pausan
  si hubo error: el banco de pruebas termina con un recuento que hay que
  leer. Véase [08 · El banco de pruebas](08-PRUEBAS.md).

La cabecera de `refrescar.bat:17` lleva el aviso operativo más importante:
**«IMPORTANTE: cierre el documento en Word antes de refrescar.»**

---

## 4. Construir los ejecutables

### 4.1 `tools/hacer_exe.ps1`

Produce tres `.exe` en `dist\`, declarados en un diccionario ordenado
(`:35-39`) cuyo orden se conserva para compilar y para el resumen final.

**Intérprete.** Prefiere `.\python\python.exe`; si no está, acepta `py` o
`python` del sistema y lo avisa en amarillo, porque *«antes se abortaba en
seco, y en un equipo sin `.\python\` no habia forma de reconstruir los .exe
(que es justo cuando hace falta)»*.

**Carpeta de trabajo fuera de OneDrive** (`:47`):

> «Si se deja dentro, la sincronizacion bloquea archivos a medio escribir y
> el empaquetado falla con "Acceso denegado". `--specpath` mueve la base de
> las rutas relativas, asi que los recursos se pasan en absoluto.»

**La configuración embebida se higieniza antes de empaquetar** (`:50-61`):
se lee `config.json`, se vacía `documento_base` y se escribe una copia en
UTF-8 sin BOM en la carpeta temporal. Esa copia, no el original, entra en
el binario:

> «'documento_base' es una ruta absoluta de la maquina que compila, y
> quedaba congelada dentro del binario. En cuanto el .exe cambiaba de manos
> (o de perfil de usuario) apuntaba a una carpeta inexistente, y TODAS las
> opciones que dependen del documento morian a la vez con el mismo error.»

**La orden de PyInstaller**, idéntica para los tres (`:72-85`):

```powershell
& $py -m PyInstaller `
    --onefile --name $nombre --console --noconfirm --clean `
    --workpath $trabajo --specpath $trabajo `
    --paths 'src' `
    --add-data "$plantilla;." --add-data "$configuracion;." `
    --collect-submodules docxtpl --collect-submodules docx `
    $entrada
```

`--paths 'src'` existe *«para que encuentre `fs_contrato` /
`fs_documento`, que se importan por nombre y no cuelgan de un paquete»*.
Los dos `--add-data` usan el separador `;` de Windows y dejan ambos
recursos en la **raíz** del directorio temporal que el `.exe` despliega al
arrancar; por eso `buscar_recurso()` los busca también en plano
(`src/generador_fs.py:98-112`). `--collect-submodules` compensa los imports
dinámicos de esas dos librerías. **Los tres son `--console`**: no hay
`--windowed`, porque la ventana gráfica la levanta PowerShell desde el
propio programa, no el binario.

**Las huellas SHA-256** se calculan al final sobre los archivos ya
producidos (`:96-103`): tamaño en MB y
`(Get-FileHash -Algorithm SHA256 $exe).Hash` por cada uno, con el cierre
`Publiquelos como release del repositorio junto con esos hashes.`

### 4.2 `GeneradorFS.spec`

```python
# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = []
hiddenimports += collect_submodules('docxtpl')
hiddenimports += collect_submodules('docx')


a = Analysis(
    ['src/generador_fs.py'],
    pathex=['src'],
    binaries=[],
    datas=[('plantillas/plantilla_estado_situacion_financiera.docx', '.'), ('config.json', '.')],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='GeneradorFS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
```

| Clave | Valor | Qué significa |
|---|---|---|
| entrada | `src/generador_fs.py` | Solo la **foto**: ni `refrescar_fs.py` ni `fs_menu.py` |
| `pathex` | `['src']` | Equivale a `--paths 'src'`: resuelve `fs_documento` y `fs_contrato` como módulos planos |
| `datas` | plantilla `.docx` y `config.json`, ambos a `.` | Aterrizan en la raíz del temporal, donde el código los busca |
| `hiddenimports` | `collect_submodules` de `docxtpl` y `docx` | Fuerza los submódulos que el análisis estático no ve |
| `binaries`, `hookspath`, `hooksconfig`, `runtime_hooks`, `excludes` | vacíos | Sin hooks ni exclusiones propias |
| `noarchive` | `False` | El bytecode viaja dentro del PYZ, no suelto |
| `optimize` | `0` | Sin `-O`: se conservan docstrings y `assert` |
| `EXE(pyz, a.scripts, a.binaries, a.datas, [], …)` | firma **onefile** | Todo en un binario. En *onedir* habría un `COLLECT` y `EXE` solo recibiría `a.scripts` |
| `console` | `True` | Aplicación de consola, coherente con el `--console` del `.ps1` |
| `upx` | `True` | Comprime con UPX solo si UPX está en el PATH; si no, se ignora sin avisar |
| `debug`, `strip`, `argv_emulation`, `disable_windowed_traceback` | `False` | Valores por defecto |
| `runtime_tmpdir` | `None` | El binario se despliega en el `%TEMP%` del sistema |
| `target_arch`, `codesign_identity`, `entitlements_file` | `None` | Campos de macOS, inertes aquí. **No hay firma de código** |

### 4.3 Discrepancias entre el `.spec` y `hacer_exe.ps1`

1. **El `.spec` no se usa.** El `.ps1` llama a PyInstaller por línea de
   órdenes con `--specpath $trabajo`, que genera un `.spec` nuevo en la
   carpeta temporal y lo borra al terminar (`:78-79` y `:91`).
   `GeneradorFS.spec` está versionado pero no participa en ninguna
   compilación del flujo actual.
2. **Cubre un ejecutable de tres.** Reconstruir con él dejaría sin
   `EstadosFinancieros.exe` ni `RefrescarFS.exe`, que son los que la
   usuaria descarga.
3. **Embebe `config.json` tal cual**, sin vaciar `documento_base`: es
   exactamente el fallo que el `.ps1` corrige. Hoy el `config.json` del
   repositorio trae `"documento_base": ""`, así que el daño no se
   materializa; pero quien rellene esa clave para probar y compile con el
   `.spec` congelará su ruta dentro del binario.
4. **Las rutas de `datas` son relativas**, así que el `.spec` solo funciona
   ejecutado desde la raíz del repositorio; el `.ps1` pasa rutas absolutas
   precisamente porque `--specpath` mueve la base.
5. `upx=True` **no es una diferencia real**: también es el comportamiento
   por defecto de la línea de órdenes, y ninguno de los dos pasa
   `--upx-dir`.
6. **`.gitignore` contiene `*.spec`** y el archivo está versionado de todos
   modos. (Inferencia: se añadió antes de esa regla, o con `git add -f`.)

---

## 5. El paquete portable

`tools/hacer_paquete.ps1` empaqueta *«el proyecto COMPLETO (incluido
`.\python\`) en un unico .zip que se ejecuta en cualquier Windows x64 sin
instalar nada y sin internet»* (`:3-5`). Exige que el Python portable
exista y, si no, aborta indicando el remedio (`:13-15`). El artefacto es
`dist\GeneradorFS_portable_<yyyyMMdd>.zip` y se rehace cada vez.

Entra esto (`:23-29`), más una carpeta `salidas\` vacía creada a propósito:

```
generar.bat  refrescar.bat  estados_financieros.bat
config.json  requirements.txt  README.md
src\  docs\  plantillas\  ejemplos\  python\  tools\
salidas\   (vacía)
```

No entra: `config.local.json`, `dist\`, `build\`, `addin\`, `.git\`,
`GeneradorFS.spec` ni los `.exe`. El `.zip` es la vía *código más
intérprete*, no la binaria. Tampoco entra `probar.bat`, aunque sí viaja
`tools\probar_refresco.py`: en el equipo destino el banco de pruebas hay
que lanzarlo a mano.

Antes de comprimir borra los `__pycache__` (`:36-37`). Comprime con
`tar.exe` y no con `Compress-Archive`, con razón escrita: *«`tar.exe`
genera un .zip con separador "/" que abre bien en cualquier extractor.
`Compress-Archive` de Windows PowerShell usa "\" y algunos extractores
fuera de Windows fallan.»* Si `tar.exe` no está, avisa y cae al otro.

**Cuándo tiene sentido frente al `.exe`.** El `.zip` sirve para iterar
sobre el código, para no pedirle a nadie que ejecute un binario opaco y
para enseñarle a TI exactamente qué archivos hay. El `.exe` sirve cuando la
usuaria solo quiere un icono: unos 13 MB frente a los ~50 MB del paquete, y
sin depender de que el equipo tolere `python.exe`. `docs/PRUEBA_EXTERNA.md`
recomienda el `.zip` para la prueba externa porque no necesita internet en
el destino.

---

## 6. Los tres ejecutables

| Ejecutable | Entrada | Qué hace |
|---|---|---|
| `EstadosFinancieros.exe` | `src/fs_menu.py` | Un solo icono que pregunta qué hacer. Es la descarga principal |
| `GeneradorFS.exe` | `src/generador_fs.py` | Crea una **foto**: un `.docx` desechable en `salidas\`, sin regiones, que ya no se puede refrescar |
| `RefrescarFS.exe` | `src/refrescar_fs.py` | Hace el **refresco** del documento vivo: reescribe solo las regiones de datos y conserva la redacción |

Los tres aceptan que se les arrastre el `.xlsx` encima y los tres llevan
embebidos los mismos dos recursos: la plantilla
`plantilla_estado_situacion_financiera.docx` y el `config.json`
higienizado.

### Sobrescribir la configuración sin recompilar

Los valores embebidos son solo el punto de partida.
`cargar_config()` (`src/generador_fs.py:240-253`) fusiona cuatro capas:

```
   DEFAULTS del código
        └─► config.json EMBEBIDA en el .exe   (documento_base vaciado)
                └─► config.json junto al .exe (viaja por git)
                        └─► config.local.json (este equipo; no se versiona)
```

> «El orden importa: lo de este equipo (`config.local.json`) tiene que
> poder ganarle a lo que venga por git, o cada «pull» reimpondria las
> rutas de la otra máquina.»

Para cambiar cualquier ajuste basta con dejar un `config.json` propio junto
al `.exe`. Hay además un rescate que evita un fallo silencioso: si el
`.exe` se queda dentro de `dist\` y ahí no hay `config.json` pero sí lo hay
un nivel por encima, el programa toma esa carpeta como raíz
(`src/generador_fs.py:74-82`). Sin eso leería la configuración del día en
que se compiló y escribiría `salidas\` y la bitácora dentro de `dist\`.

`RefrescarFS.exe` es el único que **necesita** un `config.json` a su lado:
sin `documento_base` no sabe qué documento actualizar (`tools/hacer_exe.ps1:106`).

---

## 7. Publicar una versión

El enlace del `README.md` apunta a
`https://github.com/RadioFics/Excel_Word_Paloma/releases/latest/download/EstadosFinancieros.exe`.
GitHub resuelve la forma `releases/latest/download/<nombre>` contra la
última publicación, así que **el README no se toca en cada versión**:
empieza a funcionar solo en cuanto exista la primera release, mientras el
nombre del archivo adjunto no cambie.

1. `tools\hacer_exe.ps1` → deja los tres `.exe` en `dist\` y escribe sus
   huellas SHA-256 en pantalla.
2. En GitHub: **Releases** → **Draft a new release**.
3. **Choose a tag** → una etiqueta nueva (`v0.4`, por ejemplo) con *Create
   new tag on publish*.
4. Título, y en el cuerpo **pegar las huellas SHA-256**.
5. Arrastrar los tres `.exe` a *Attach binaries by dropping them here*.
6. **Publish release.**

El repositorio tiene que ser **público** para que el enlace funcione sin
iniciar sesión. Del lado del usuario, la verificación es una orden:

```
powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '.\EstadosFinancieros.exe').Hash"
```

El resultado debe coincidir carácter por carácter con la huella publicada.
Es el único mecanismo de integridad que existe: no hay firma de código en
ningún punto de la cadena.

---

## 8. La advertencia de SmartScreen

**Por qué aparece.** Los `.exe` no llevan firma digital —`codesign_identity`
está en `None` y la orden de PyInstaller tampoco firma— y son binarios
recién publicados, sin reputación acumulada. Windows los recibe descargados
de internet y muestra **«Windows protegió su PC»**.

**Qué decirle al usuario.** *Más información → Ejecutar de todas formas*,
con el contexto que el `README.md` da al lado: *«lee un `.xlsx` de tu
equipo y escribe un `.docx` de tu equipo (…) no se instala, no toca la
configuración de Windows, no se conecta a internet, no envía datos a ningún
lado»*. Lo prudente es acompañar esa instrucción con la verificación de la
huella SHA-256: es lo que distingue «este archivo es el que publicamos» de
«este archivo me lo dio alguien».

**Cuándo NO hay que forzar la ejecución.** SmartScreen es un aviso de
reputación y ofrece continuar. Un **bloqueo por política corporativa** no
es lo mismo. `docs/INSTALACION.md`, `docs/PRUEBA_EXTERNA.md` y
`docs/DIRECCION.md` coinciden en señalar AppLocker, WDAC y el antivirus o
EDR como el riesgo abierto del proyecto:

> «Si el equipo aplica control de aplicaciones, puede bloquear la ejecución
> de `python.exe` desde una carpeta de usuario o una unidad de red aunque
> no haya "instalación". **Confirmar con TI** antes de depender de esto.»
> (`docs/INSTALACION.md:42-45`)

Si lo que aparece es un mensaje del administrador o de la herramienta de
seguridad del equipo —y no el diálogo de SmartScreen con su enlace *Más
información*—, la respuesta correcta es **detenerse y escalar a TI**, no
buscar la forma de saltárselo. `docs/DIRECCION.md:72-74` lo trata como una
decisión de proyecto: si el control de aplicaciones frena el binario, la
vía queda descartada y se acelera el complemento de Office. (Inferencia: la
diferencia visible es que el bloqueo por política no ofrece ninguna opción
de continuar; el repositorio no describe ese mensaje.)

---

## 9. Higiene del repositorio

### 9.1 `.gitignore`

```gitignore
# --- Python portable del proyecto ---------------------------------------
# NO se versiona (52 MB, especifico de Windows x64). En un equipo nuevo:
#   - con internet:  tools\bootstrap_python.ps1
#   - sin internet:  use el .zip de tools\hacer_paquete.ps1 (lo incluye)
/python/

# Entornos virtuales sueltos (de intentos previos; generar.bat NO los usa)
/.venv/
/venv/

# Paquete distribuible generado (el .exe / .zip se publican como release)
/dist/
/build/
*.spec

# Artefactos de Python
__pycache__/
*.pyc

# Documentos generados y fotos de estado (la carpeta se conserva vacia)
/salidas/*
!/salidas/.gitkeep

# Node / add-in
/addin/node_modules/
/addin/dist/
/addin/.env

# Temporales de Office y del sistema
~$*
Thumbs.db
.DS_Store

# Copias previas que deja fs_documento.py antes de escribir un archivo
*.docx.bak
*.xlsx.bak

# Temporal de la escritura atomica (se reemplaza de un golpe al terminar)
.*.tmp.docx

# Ajustes de ESTE equipo (rutas absolutas, documento base). NO se versiona:
# config.json si viaja por git, y una ruta de una maquina escrita ahi se le
# impone a la otra en cada pull.
/config.local.json
```

| Grupo | Regla | Motivo |
|---|---|---|
| Intérprete | `/python/` | 52 MB de binarios Windows x64, reconstruibles |
| Entornos sueltos | `/.venv/`, `/venv/` | Restos de intentos previos; el flujo actual no los usa |
| Artefactos de compilación | `/dist/`, `/build/`, `*.spec` | Los binarios se publican como release |
| Bytecode | `__pycache__/`, `*.pyc` | Regenerable |
| Documentos generados | `/salidas/*` con `!/salidas/.gitkeep` | Se ignora el contenido pero **se conserva la carpeta vacía**, porque el generador escribe ahí |
| Complemento | `/addin/node_modules/`, `/addin/dist/`, `/addin/.env` | Compilación de TypeScript y secretos |
| Basura del sistema | `~$*`, `Thumbs.db`, `.DS_Store` | `~$*` son los archivos de propietario que Word y Excel dejan al abrir |
| Copias y temporales | `*.docx.bak`, `*.xlsx.bak`, `.*.tmp.docx` | El respaldo previo a cada escritura y el temporal de la escritura atómica |
| Ajustes del equipo | `/config.local.json` | El contra-veneno al pisoteo de rutas |

Las reglas que empiezan por `/` están ancladas a la raíz; `__pycache__/`,
`~$*` o `*.docx.bak` aplican a cualquier profundidad.

Las tres decisiones de fondo: **`python/`** son 52 MB de una sola
plataforma, y versionarlos ataría el repositorio a Windows x64;
**`dist/` y `salidas/`** son salida y no fuente, la una reconstruible con
un script y la otra distinta en cada ejecución; y **`config.local.json`**
queda fuera por el motivo que el propio archivo declara —*«una ruta de una
maquina escrita ahi se le impone a la otra en cada pull»*—, que es lo que
permite que dos equipos compartan `config.json` sin pisarse el
`documento_base`.

### 9.2 `.gitattributes`

```gitattributes
# Normaliza saltos de linea. Los .bat/.ps1 se quedan en CRLF (Windows).
* text=auto eol=lf
*.bat text eol=crlf
*.ps1 text eol=crlf

# Binarios: sin conversion ni diff textual
*.xlsx binary
*.docx binary
*.dotx binary
*.png binary
*.zip binary
```

| Regla | Efecto |
|---|---|
| `* text=auto eol=lf` | Git decide si es texto; se almacena y **se extrae** con LF, porque el `eol=lf` explícito anula el `core.autocrlf` de cada usuario |
| `*.bat text eol=crlf` | CRLF obligatorio en el árbol de trabajo |
| `*.ps1 text eol=crlf` | Igual, por coherencia |
| `*.xlsx *.docx *.dotx *.png *.zip binary` | Macro equivalente a `-text -diff`: sin conversión de finales de línea y sin diff textual |

**Por qué los `.bat` y `.ps1` se quedan en CRLF.** Los interpreta Windows,
no Git. `cmd.exe` puede comportarse de forma errática con un `.bat` en LF,
sobre todo con etiquetas `goto` y bloques entre paréntesis: exactamente la
forma de `tools/buscar_python.bat`, lleno de `goto :comprobar` y de
`if exist ( … )`. Forzar CRLF elimina esa clase entera de fallo.

**Por qué los ofimáticos van como `binary`.** Un `.docx`, un `.xlsx` y un
`.dotx` son archivos ZIP: la conversión de finales de línea los
corrompería. Marcarlos `binary` la desactiva y evita, de paso, que Git
intente un diff textual de datos comprimidos.

**No hay Git LFS**: ni filtros `filter=lfs` ni `.lfsconfig`. Las plantillas
y los libros de ejemplo van en los objetos de Git, y ese es justo el motivo
de que `.gitignore` excluya con tanta insistencia lo pesado y regenerable.

---

## Resumen del capítulo

- La organización bloquea instaladores: de ahí el Python embebido en
  `.\python\` y los `.exe` de un archivo, que no tocan el registro ni piden
  administrador.
- `bootstrap_python.ps1` fija Python 3.12.10 `amd64`; su paso crítico es
  habilitar `Lib\site-packages` en el `._pth`, sin el cual pip instala pero
  nada se importa.
- `buscar_python.bat` prueba el portable, luego `py -3`, luego `python`, y
  pregunta siempre por `sys.executable` para que el atajo de la Microsoft
  Store no se cuele.
- `hacer_exe.ps1` es el flujo real de compilación; `GeneradorFS.spec` está
  versionado pero no se usa, cubre un ejecutable de tres y no vacía
  `documento_base`.
- Vaciar `documento_base` antes de empaquetar es lo que evita que un `.exe`
  salga con la ruta de la máquina que lo compiló congelada dentro.
- El `.zip` de `hacer_paquete.ps1` es la vía código más intérprete, sin
  internet en el destino; el `.exe` es la vía de un solo icono.
- La release usa `releases/latest/download/`, de modo que el enlace del
  README nunca se actualiza; la integridad se apoya solo en las huellas
  SHA-256, porque no hay firma de código.
- SmartScreen se salta con *Más información → Ejecutar de todas formas*; un
  bloqueo de AppLocker, WDAC o EDR no se salta: se escala a TI.
