# 05 · La interfaz y las órdenes

> **Para quién.** Quien tenga que dar soporte al sistema, automatizarlo desde
> un `.bat` o una tarea programada, o modificar la ventana del menú.
> **Qué encontrará.** Cómo `src/fs_menu.py` decide qué hacer, cómo dibuja su
> ventana sin usar tkinter, qué banderas acepta la línea de órdenes, cómo
> trabaja `src/refrescar_fs.py` y qué contiene la configuración en uso.
> **Antes de leer.** Conviene tener presente la distinción entre el
> **documento vivo** (el `.docx` con regiones dentro, el que se refresca) y la
> **foto** (el `.docx` desechable que se genera en `salidas\`), explicada en
> [Arquitectura](02-ARQUITECTURA.md).

## Índice del capítulo

- [1. El papel de esta capa](#1-el-papel-de-esta-capa)
- [2. Cómo se dibuja la ventana](#2-cómo-se-dibuja-la-ventana)
- [3. La consola duplicada](#3-la-consola-duplicada)
- [4. La cabecera de estado](#4-la-cabecera-de-estado)
- [5. Cómo se resuelve el libro de Excel](#5-cómo-se-resuelve-el-libro-de-excel)
- [6. El menú, opción por opción](#6-el-menú-opción-por-opción)
- [7. Los dos flujos largos](#7-los-dos-flujos-largos)
- [8. La línea de órdenes](#8-la-línea-de-órdenes)
- [9. `refrescar_fs.py` paso a paso](#9-refrescar_fspy-paso-a-paso)
- [10. La configuración en uso](#10-la-configuración-en-uso)
- [11. Qué pasa cuando algo falla](#11-qué-pasa-cuando-algo-falla)
- [Resumen del capítulo](#resumen-del-capítulo)

---

## 1. El papel de esta capa

`src/fs_menu.py` es el punto de entrada único del proyecto: es lo que hay
detrás de `EstadosFinancieros.exe` y de `estados_financieros.bat`. No hace
trabajo propio. Su docstring lo declara sin rodeos (`src/fs_menu.py:7-15`):

> «NO duplica lógica: delega en los mismos módulos que usan los ejecutables
> individuales, que siguen funcionando por su cuenta.
>
>     generador_fs.ejecutar()   ->  GeneradorFS.exe  /  generar.bat
>     refrescar_fs.ejecutar()   ->  RefrescarFS.exe  /  refrescar.bat
>     fs_documento.crear_base() ->  la plantilla viva, desde cero
>     fs_documento.preparar()   ->  adaptar un Word cualquiera
>     fs_documento.estado()     ->  la radiografía del proyecto»

Importa los otros tres con alias fijos —`G` para `generador_fs`, `D` para
`fs_documento`, `R` para `refrescar_fs` (`src/fs_menu.py:50-52`)— y en todo el
archivo no abre por su cuenta ni un `.docx` ni un `.xlsx`, salvo para calcular
la huella del libro en `crear_plantilla`
(`hashlib.sha256(libro.read_bytes())`, `src/fs_menu.py:806`).

### La confusión que lo hizo nacer

El motivo de existir del menú está en las primeras líneas del docstring
(`src/fs_menu.py:4-6`):

> «Existe porque los dos caminos se confunden con facilidad: ambos aceptan
> que se les arrastre el Excel, pero uno crea un documento nuevo y el otro
> actualiza el que ya hay. Aquí se elige explícitamente.»

Y a continuación fija la distinción conceptual central del sistema
(`src/fs_menu.py:17-23`):

> «Los dos «crear» no son lo mismo y conviene no confundirlos:
>
>     opción 6  crear_base()   -> un Word VIVO, con regiones dentro. Es la
>                                base que se refresca cada cierre.
>     opción 2  generador_fs   -> una FOTO en salidas\\, renderizada de una
>                                plantilla de Word. No lleva regiones, así
>                                que no se puede volver a actualizar.»

La *región* es el control de contenido de Word (el elemento `w:sdt`, la caja
etiquetada que Word sabe delimitar dentro del documento) que lleva un
**ancla**: la cadena `fs-…` del atributo `w:tag`. Un documento sin regiones no
tiene dónde recibir las cifras, y por eso una foto no sirve como base de
trabajo: quien redacte encima pierde su texto en la siguiente entrega.

Antes del menú, ambos caminos eran dos iconos indistinguibles a los que se
arrastraba el mismo Excel. La ventana no añade capacidades: añade una
pregunta explícita antes de actuar.

---

## 2. Cómo se dibuja la ventana

Este es el punto arquitectónico más llamativo del archivo: **la interfaz no
es tkinter**. Es una ventana de WinForms (la biblioteca clásica de ventanas de
.NET) dibujada por un script de PowerShell que Python escribe en un archivo
temporal y ejecuta como un proceso aparte.

### 2.1 Por qué

El comentario de cabecera de `_PS_VENTANA` (`src/fs_menu.py:186-190`) da la
razón completa:

> «Ventana del menú. Se dibuja con WinForms desde PowerShell porque el
> Python portable no trae tkinter y meterlo obligaría a arrastrar Tcl/Tk
> dentro del .exe (unos 10 MB más y un empaquetado más frágil). PowerShell y
> .NET están en cualquier Windows, así que no añade ninguna dependencia. Si
> algo falla, se cae al menú de consola de siempre.»

Hay dos scripts embebidos como literales `r"""…"""`:

| Constante | Líneas | Qué dibuja |
|---|---|---|
| `_PS_VENTANA` | `src/fs_menu.py:191-473` | El menú de tarjetas |
| `_PS_RESULTADO` | `src/fs_menu.py:480-585` | La ventana de informe posterior a cada operación |

Ambos hacen `Add-Type -AssemblyName System.Windows.Forms` y
`System.Drawing`, y llaman a
`[System.Windows.Forms.Application]::EnableVisualStyles()`.

### 2.2 El puente: `_lanzar_ps`

```python
def _lanzar_ps(script_texto, *argumentos, timeout=1800):
```

(`src/fs_menu.py:588`). No reimplementa nada; delega en
`D.ejecutar_ps(script_texto, *argumentos, timeout=600, sta=True)`
(`src/fs_documento.py:321`). El docstring explica por qué se delega:

> «Delega en D.ejecutar_ps, que es quien sabe escribir el .ps1 con BOM y
> leer la salida en UTF-8 en vez de con la página de códigos de la consola.
> No se duplica aquí para que el arreglo del juego de caracteres valga en un
> solo sitio.»

`D.ejecutar_ps` escribe `orden.ps1` en un directorio temporal con
`encoding="utf-8-sig"` —la marca de orden de bytes (BOM) es obligatoria,
porque PowerShell 5.1 lee un archivo sin BOM como ANSI y destroza los acentos
del propio script—, lanza
`powershell -NoProfile -ExecutionPolicy Bypass -STA -File …` (el modo `-STA`
lo exigen los diálogos de WinForms) y decodifica la salida en UTF-8 con
`errors="replace"`.

La pieza clave de `_lanzar_ps` es su `except Exception: return None`
(`src/fs_menu.py:600-601`): ese `None` es el **conmutador de degradación**. Si
no hay PowerShell, si expira el tiempo o si falla el `Add-Type`, la aplicación
no aborta; cae al menú de texto.

### 2.3 Nitidez en pantallas de alta densidad

El bloque de `src/fs_menu.py:194-208` se ocupa de los puntos por pulgada (la
densidad de píxeles de la pantalla, «ppp» o «DPI»). El comentario:

> «Sin esto, Windows dibuja la ventana a 96 ppp y la escala como una imagen:
> el texto sale borroso ("pixelado"). Hay que declararlo ANTES de crear
> ninguna ventana.»

Se compila C# en caliente con `Add-Type -TypeDefinition`, declarando una
clase estática `Ppp` con dos llamadas a `user32.dll`. La escalera de intentos:

1. `[Ppp]::SetProcessDpiAwarenessContext([IntPtr](-4))` — el modo «por
   monitor v2», el más moderno.
2. Si falla, `[Ppp]::SetProcessDPIAware()` — la interfaz antigua, por sistema.
3. Si falla todo, se sigue igualmente: borroso, pero funcional. Todo va
   dentro de `try{}catch{}`.

Después se mide la escala real de la pantalla y se define una función que
convierte todas las medidas (`src/fs_menu.py:279-282`):

```powershell
$g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
$k = $g.DpiX / 96.0
$g.Dispose()
function S($n) { [int][Math]::Round($n * $k) }
```

Toda coordenada y todo tamaño del script pasan por `S(…)`. `_PS_RESULTADO`
repite el bloque con la clase renombrada a `PppRes` (`src/fs_menu.py:482-492`):
han de ser tipos distintos porque `Add-Type` no permite redefinir un nombre
dentro del mismo proceso.

### 2.4 La paleta

Declarada literalmente en `src/fs_menu.py:225-232`:

| Variable | `FromArgb` | Hex | Uso |
|---|---|---|---|
| `$tinta` | `26, 34, 38` | `#1A2226` | Texto principal, opciones neutras |
| `$suave` | `102, 114, 120` | `#667278` | Subtítulos, detalles, botón `Salir` |
| `$fondo` | `247, 246, 243` | `#F7F6F3` | Fondo del formulario |
| `$tarjeta` | `White` | `#FFFFFF` | Fondo de cada tarjeta |
| `$borde` | `224, 221, 214` | `#E0DDD6` | Bordes |
| `$acento` | `17, 94, 89` | `#115E59` | Opciones 1 y 6; título cuando todo fue bien |
| `$ambar` | `146, 94, 20` | `#925E14` | Avisos: `SIN EXCEL`, cifras `EDITABLES`, opción 4 |

`_PS_RESULTADO` reutiliza `$tinta`, `$suave`, `$fondo` y `$acento`, y añade
`$rojo` = `155, 44, 44` (`#9B2C2C`) para el título de un resultado fallido. El
resalte al pasar el ratón, `FromArgb(240, 245, 244)` (`#F0F5F4`), va en línea
dentro de `Nueva-Opcion` y no forma parte de la paleta declarada.

Los iconos no son archivos: son glifos de la tipografía `Segoe MDL2 Assets`
(`src/fs_menu.py:234-236`):

> «Segoe MDL2 Assets trae los iconos de Windows como tipografia: son
> vectoriales, se ven nitidos a cualquier escala y no hay que empaquetar
> ningun archivo. Si no estuviera, se cae a texto sin icono.»

Si la fuente no existe, el diccionario de iconos se rellena con cadenas
vacías, no con sustitutos.

### 2.5 El flujo de datos entre Python y la ventana

La comunicación es deliberadamente pobre: cuatro argumentos de entrada y una
línea de texto de salida.

```
  Python (fs_menu.ejecutar)
      |
      |  _describir_destino()  ->  "EF.docx"
      |  _resolver_libro(args) ->  "cierre FS.xlsx"
      |  _describir_cifras()   ->  "PROTEGIDAS (37 regiones)"
      v
  menu_ventana(destino, libro, cifras)
      |
      |  _lanzar_ps(_PS_VENTANA, destino, libro, cifras, "")
      v
  D.ejecutar_ps  --escribe-->  %TEMP%\fs_ps_xxxx\orden.ps1   (UTF-8 con BOM)
      |
      |  powershell -NoProfile -ExecutionPolicy Bypass -STA -File orden.ps1 …
      v
  WinForms dibuja la ventana; el usuario pulsa una tarjeta
      |
      |  stdout:  "ELECCION=6"
      v
  menu_ventana lee las líneas, busca el prefijo "ELECCION=" y
  devuelve "6"  (o None si no encontró ninguna)
```

En la ventana de resultado hay un matiz añadido: el informe **viaja por
archivo, no por argumento** (`src/fs_menu.py:500-502`):

> «El informe llega por ARCHIVO, no como argumento: un texto de varias
> lineas pasado en la linea de ordenes se parte por los saltos y powershell
> lo recibe troceado en $args, con lo que la ventana no llegaba a abrirse.»

Por eso `ventana_resultado` crea un temporal `fs_res_*`, escribe `informe.txt`
en UTF-8, pasa la ruta al script y borra el directorio en un `finally`
(`src/fs_menu.py:617-634`).

### 2.6 El detalle que hacía inservible el menú

En `src/fs_menu.py:306-320` está documentado un fallo que dejaba la
aplicación aparentemente muerta:

> «La eleccion vive DENTRO de una tabla hash, no en una variable suelta.
> Motivo: los manejadores de clic se crean con .GetNewClosure(), que le da a
> cada bloque su propio ambito capturado. Una asignacion como
> «$script:eleccion = $valor» hecha ahi dentro escribe en ese ambito
> aislado, no en el del script: la ventana se cerraba, pero el valor que se
> leia al final seguia siendo el inicial ('0' = Salir). O sea, TODAS las
> tarjetas del menu acababan comportandose como «Salir», y por eso la
> aplicacion parecia un puñado de botones sin efecto.
> Mutar un objeto SI atraviesa el ambito: la closure y el script comparten
> la misma tabla hash.»

De ahí `$estado = @{ valor = '0' }` (`src/fs_menu.py:321`) y, en cada tarjeta,
`{ $estado.valor = $valor; $f.Close() }.GetNewClosure()`.

Hay además un **modo captura** (`src/fs_menu.py:462-471`): con un cuarto
argumento que sea una ruta `.png`, la ventana se muestra, se vuelca con
`DrawToBitmap` y se cierra sola. `menu_ventana` siempre pasa `""` en esa
posición (`src/fs_menu.py:607`), así que es un gancho para documentación y
pruebas, no una ruta de usuario.

---

## 3. La consola duplicada

`_Eco` (`src/fs_menu.py:68`) desdobla `sys.stdout`. Su docstring:

> «Escribe en la consola de siempre Y en un buffer, a la vez. La consola
> sigue siendo la fuente para quien ejecuta desde una terminal o un .bat; el
> buffer alimenta la ventana de resultado.»

Su `write` escribe primero en el buffer y luego en el flujo real, protegido:

```python
try:
    self._real.write(s)
except Exception:
    pass          # una consola cerrada no debe tumbar la operacion
```

`_consola_duplicada()` (`src/fs_menu.py:97-105`) es un gestor de contexto que
sustituye `sys.stdout` por un `_Eco`, cede el objeto y **restaura el flujo
anterior en el `finally`**. Se usa en las cinco ramas de trabajo de `ejecutar`
(opciones 1, 2, 3, 4/5 y 6). El problema que resuelve
(`src/fs_menu.py:925-928`):

> «Se hace el trabajo con la consola duplicada a un buffer: lo mismo que se
> imprime se enseña luego en una ventana. Sin esto el informe acababa en una
> consola que aparece detras de todo, y la operacion parecia no haber hecho
> nada.»

El mismo texto se emite una sola vez y sirve para los dos públicos —quien
ejecuta desde una terminal y quien pulsó un botón— sin duplicar ninguna línea
de código de informe.

---

## 4. La cabecera de estado

La ventana declara siempre tres cosas antes de ofrecer nada. No es adorno:
cada línea corrige un caso en que el usuario no podía saber sobre qué actuaba.

| Línea | La calcula | Texto si no se sabe |
|---|---|---|
| `Documento: …` | `_describir_destino()` (`src/fs_menu.py:115`) | `Sin documento configurado — use «Cambiar el documento»` |
| `Excel: …` | `_resolver_libro(args)` (`src/fs_menu.py:156`) | `SIN EXCEL - arrastre el libro sobre el icono; las opciones 1 y 2 lo necesitan` (en ámbar) |
| `Cifras: …` | `_describir_cifras()` (`src/fs_menu.py:136`) | la línea queda vacía |

`_describir_destino()` carga la configuración, resuelve el documento y
pregunta a `D.quien_bloquea(doc)` quién lo tiene abierto. Si hay alguien,
devuelve `f"{doc.name}  [ABIERTO: {culpables[0]}]"`; si algo falla,
`"sin configurar (config.json -> documento_base)"`. Nunca lanza excepción.
`_hay_documento()` (`src/fs_menu.py:128`) es su versión booleana, y alimenta
el arranque guiado.

`_describir_cifras()` interroga el **candado** (el bloqueo por región,
`w:lock`) y la **protección** (el bloqueo de documento de Word,
`w:documentProtection`) con `D.estado_candado(doc)`, que devuelve
`(bloqueadas, total, proteccion)`. Salidas literales:

| Condición | Texto devuelto |
|---|---|
| Cualquier excepción | `""` |
| `not total` | `el documento aún no tiene regiones de datos` |
| `bloqueadas == total` | `PROTEGIDAS (N regiones)`, con `  ·  modo estricto` si además hay protección |
| `bloqueadas == 0` | `EDITABLES a mano (N regiones sin candado)` |
| Resto | `MIXTO (B de N con candado)` |

El prefijo `EDITABLES` no es solo informativo: es la condición
(`$cifras -like 'EDITABLES*'`) que pinta la línea en ámbar y decide si la
quinta tarjeta ofrece proteger (opción 5) o desproteger (opción 4). El
comentario que la introduce (`src/fs_menu.py:294-297`):

> «Estado REAL del candado, leido del documento. Sin esta linea, las
> opciones de proteger/desproteger eran dos botones ciegos: no habia forma
> de saber en que estado estaba el documento, asi que pulsarlos no parecia
> tener ningun efecto.»

Y el de la línea del Excel (`src/fs_menu.py:283-285`):

> «De donde salen las cifras. Sin esta linea no habia forma de saber si el
> Excel que se arrastro se recogio o no, ni de distinguirlo del que la
> aplicacion elige sola por convencion de nombre.»

---

## 5. Cómo se resuelve el libro de Excel

```python
def _resolver_libro(args):
```

(`src/fs_menu.py:156`) devuelve una tupla de tres:
`(ruta o None, texto para la ventana, aviso o None)`. Su docstring explica el
riesgo concreto que evita:

> «Hay que saberlo ANTES de dibujar el menú. Si no, el usuario arrastra su
> libro y la ventana no le dice si lo ha recogido; y cuando no arrastra
> nada, encontrar_excel_por_convencion() coge en silencio cualquier .xlsx de
> la carpeta cuyo nombre contenga la clave de convención —incluido el de
> ejemplos\— y las cifras de muestra acaban dentro del documento real sin
> que nadie lo haya pedido.»

Orden de preferencia:

1. **Argumento arrastrado o escrito**: si hay posicionales, se toma
   `Path(args[0]).resolve()`.
   - Si no existe: `(None, "NO EXISTE: <nombre>", "El archivo indicado no existe:\n  <ruta>")`.
   - Si existe: `(ruta, ruta.name, None)` — sin aviso, porque la elección fue
     explícita.
2. **Convención de nombre**: sin argumentos, se llama a
   `G.encontrar_excel_por_convencion(cfg)`, que busca `*.xlsx` en la raíz del
   proyecto y luego en `ejemplos\`, filtra por que el nombre contenga la clave
   `buscar_por_convencion` (aquí, `FS`), descarta los temporales `~$*` y se
   queda con el más reciente por fecha de modificación.
   - Si falla: `(None, "", None)` — sin libro y sin aviso.
   - Si acierta: la ruta, el texto
     `"<nombre>  (elegido solo, por convención de nombre)"` y **siempre** un
     aviso que explica por qué se eligió ese y cómo cambiarlo.
3. **Diálogo de archivo**: no lo hay para el Excel. Los dos únicos diálogos
   del sistema son para el Word: `D.elegir_archivo_word` (opción 3) y
   `D.elegir_destino_word` (opción 6). Un libro solo entra arrastrado, escrito
   en la orden, o por convención.

Con libro ya resuelto, `_contexto_del_libro(libro)` (`src/fs_menu.py:679`)
obtiene el **contexto** (`ctx`, el diccionario que `leer_contexto()` extrae del
Excel) y le quita las dos claves de metadatos antes de entregarlo:

```python
cfg = G.cargar_config()
ctx = G.leer_contexto(libro, cfg)
ctx.pop("_meta", None)
ctx.pop("_avisos", None)
return ctx, cfg
```

Su docstring: *«Se usa para poder montar el andamiaje YA relleno. Sin libro,
el andamiaje se monta igual, solo que con los huecos a la vista.»* El
**andamiaje** es el conjunto de regiones que `construir()` añade al documento.

---

## 6. El menú, opción por opción

### 6.1 Las opciones

| Op. | Texto literal de la tarjeta | Invoca | Efecto sobre el disco |
|---|---|---|---|
| `1` | `Actualizar el documento de siempre` | `R.ejecutar(resto)` | Reescribe **solo las regiones** del documento vivo. Deja un `.bak` al lado. Prepara el documento antes si le faltan regiones. |
| `6` | `Crear la plantilla base desde el Excel` | `crear_plantilla(libro)` | Crea un `.docx` nuevo donde el usuario diga, con andamiaje completo, lo refresca y lo fija en `config.local.json`. |
| `3` | `Cambiar el documento que se actualiza` | `cambiar_documento(libro)` | Escribe `documento_base` en `config.local.json`. Si el elegido no está integrado, le monta el andamiaje (y deja `.bak`). |
| `2` | `Crear un documento nuevo (copia desechable)` | `G.ejecutar(resto)` | Crea una **foto** en `salidas\`. No toca el documento vivo. |
| `4` | `Permitir editar las cifras a mano en Word` | `D.desproteger` + `D.cambiar_candado(bloquear=False)` | Reescribe el documento vivo quitando candados y protección. `.bak` previo. |
| `5` | `Volver a proteger las cifras` | `D.cambiar_candado(bloquear=True)` + `D.proteger_salvo_datos` | Reescribe el documento vivo poniendo candados y protección. `.bak` previo. |
| `0` | `Salir` (botón, no tarjeta) | — | Nada. Imprime `Nada que hacer.` |

Detalles que conviene no perder de vista:

- **Las tarjetas 4 y 5 son un solo interruptor**, no dos botones: solo se
  dibuja una, según el prefijo `EDITABLES` (`src/fs_menu.py:428-429`): *«La
  etiqueta dice que va a PASAR, y la linea «Cifras:» de arriba dice como estan
  AHORA.»*
- **El orden visual no es el numérico**: las tarjetas van 1, 6, 3, 2 y luego
  el interruptor (`y` = 140, 218, 296, 374, 452), y el menú de texto lo repite.
- **La ventana no ofrece el diagnóstico.** `--estado` no tiene tarjeta
  (`src/fs_menu.py:996`): *«Opción no listada en el menú: diagnóstico para
  quien da soporte.»* Tampoco hay tarjeta para `--consola` ni para las
  banderas de compatibilidad (`--preparar`, `--no-preparar`, `--revisar`), que
  atraviesan el menú y las consume el módulo destino.

Cada tarjeta lleva un texto de detalle y una ayuda emergente larga (hasta 20
segundos en pantalla: `AutoPopDelay = 20000`). La de la opción 5 es la que más
importa para el soporte:

> «OJO: solo protege lo que esta dentro de una region marcada. Un numero
> copiado y pegado del Excel como texto normal NO queda protegido, porque el
> programa no puede saber que es una cifra.»

### 6.2 Las tres funciones de presentación

```python
def menu_ventana(destino, libro="", cifras="")
```
(`src/fs_menu.py:604`). Devuelve la cadena de la opción (`'0'`…`'6'`) o
`None`. **`None` significa «no se pudo dibujar»**, y es la señal para que el
llamante caiga al menú de texto.

```python
def ventana_resultado(titulo, cuerpo, destino=None, ok=True)
```
(`src/fs_menu.py:617`). Muestra el informe con tipografía monoespaciada
(`Consolas` 9.5 pt, porque los informes son tabulados) y, **solo si `destino`
existe en disco**, añade los botones `Abrir el documento` y `Abrir la
carpeta`. `Cerrar` se asigna a `AcceptButton`, de modo que Enter cierra. Si
`_lanzar_ps` devuelve `None`, no pasa nada: el informe ya está en la consola.
Su justificación (`src/fs_menu.py:476-479`):

> «Sin esto, la aplicacion cerraba la ventana del menu, hacia el trabajo y
> volcaba el informe en una consola que aparece DETRAS: desde el punto de
> vista del usuario, el boton no hacia nada. Aqui se le dice que paso, donde
> quedo, y se le ofrece abrirlo.»

```python
def _menu(texto_libro="")
```
(`src/fs_menu.py:635`). El menú de texto. Imprime la cabecera de `_cabecera()`
(`" ESTADOS FINANCIEROS — ¿qué quiere hacer?"` entre dos filas de 68 signos
`=`) y las siete entradas. Lee con
`input(" Escriba una opción (0-6) y pulse Enter: ")`; `EOFError` y
`KeyboardInterrupt` devuelven `"0"`, y cualquier otra cosa imprime
`" No entendí esa opción."` y repite.

Dos diferencias con la ventana: aquí **se ofrecen 4 y 5 a la vez** —no es un
interruptor— y el estado del candado aparece como una línea `Ahora mismo: …`
bajo la opción 5, solo si `_describir_cifras()` devolvió algo.

---

## 7. Los dos flujos largos

### 7.1 `cambiar_documento(libro=None)`

(`src/fs_menu.py:694`). Su docstring describe las tres formas en que puede
venir un documento y el fallo que corrige:

> «ya integrado  -> se fija y no se toca nada más
> en blanco     -> se usa de base y se le monta el estado encima
> con redacción -> se le añade el estado como apartado aparte
>
> Antes solo servía la primera, y las otras dos salían por pantalla como «NO
> SIRVE COMO DOCUMENTO BASE» o se quedaban a medias: el archivo se fijaba,
> pero el refresco siguiente no encontraba dónde escribir y terminaba con «no
> se actualizó NADA».»

Paso a paso:

1. `cfg = G.cargar_config()`.
2. Intenta `D.resolver_documento(None, cfg)` solo para saber la carpeta
   inicial del diálogo: si lo consigue imprime `" Documento actual: <nombre>"`;
   si lanza `ValueError`, `" Ahora mismo no hay ningún documento configurado."`
3. Imprime `" Abriendo el explorador…"` y llama a
   `D.elegir_archivo_word(carpeta)` (`src/fs_documento.py:2457`), un
   `OpenFileDialog` de WinForms lanzado por el mismo mecanismo de PowerShell.
4. Cancelado: `" No se eligió ninguno. Nada ha cambiado."` y **devuelve 0**.
5. `ok, avisos, _familias, info = D.revisar_candidato(elegido)` y acto seguido
   `elegido = info["ruta"]`, con este comentario: *«revisar_candidato puede
   haber encontrado el archivo bajo un nombre ligeramente distinto (espacios
   duros, tildes). Se sigue con ESE.»*
6. Imprime un bloque de 68 signos `=` con el nombre y la carpeta del elegido.
7. Si `not ok`: imprime `" NO SE PUEDE USAR ESTE ARCHIVO:"`, la lista de
   avisos, `" No se ha cambiado nada."` y **devuelve 1**.
8. Si `ok`: imprime los avisos y llama a `D.fijar_documento_base(elegido)`.
9. Si `info["estado"] != D.LISTO`, obtiene el contexto del libro (si lo hay),
   imprime `" Preparándolo para que se le puedan volcar las cifras…"` y llama a
   `D.preparar(elegido, ctx, cfg_ctx or cfg)`. Sin Excel a mano avisa de que se
   montan las regiones vacías y que la primera actualización las rellena.
10. Cierra con `" Hecho: a partir de ahora «Actualizar» trabaja sobre este
    documento."` y **devuelve 0**.

Lo que queda escrito en `config.local.json` lo decide
`D.fijar_documento_base` (`src/fs_documento.py:2679`): la clave
`documento_base` con la ruta **compactada** por `G.compactar_ruta` (convertida
a `${ONEDRIVE}\…` cuando se puede), más un `_comentario` puesto por
`setdefault` si el archivo aún no lo tenía. Su docstring:

> «Va a config.local.json, no a config.json: el segundo viaja por git, y una
> ruta absoluta escrita ahí se le impone a la otra máquina en cada «pull».
> Además la ruta se guarda compactada (${ONEDRIVE}\\…) para que, si alguien
> la copia al config.json compartido, siga valiendo en ambas.»

### 7.2 `crear_plantilla(libro)`

(`src/fs_menu.py:766`). Devuelve una tupla `(codigo, destino)`.

1. Si `libro is None`, lanza `ValueError` con el texto: *«Para crear la
   plantilla hacen falta las cifras de un Excel. / Arrastre su .xlsx sobre el
   icono y vuelva a elegir esta opción.»*
2. Imprime `" Leyendo las cifras de: <nombre>"` y obtiene el contexto.
3. Compone el nombre sugerido:
   `f"Estados financieros - {G.sanear(ctx.get('fecha_actual') or 'base')}.docx"`.
4. Imprime `" Elija dónde guardarla (local o en OneDrive)…"` y abre
   `D.elegir_destino_word(nombre_sugerido=sugerido)`
   (`src/fs_documento.py:2532`), un `SaveFileDialog` que fuerza la extensión
   `.docx`.
5. Cancelado: `" No se eligió destino. No se ha creado nada."` y devuelve
   `(0, None)`.
6. Imprime destino y carpeta, y llama a `D.crear_base(destino, ctx, cfg)`.
7. **Refresca inmediatamente lo que acaba de crear**, con su razón
   (`src/fs_menu.py:801-804`):

   > «Recién creada ya lleva las cifras dentro (construir las escribe al
   > montar las regiones), pero se refresca igual: así queda la foto de
   > metadatos y la bitácora, que es lo que hace comparables los refrescos
   > siguientes.»

   ```python
   sha = hashlib.sha256(libro.read_bytes()).hexdigest()[:12]
   D.refrescar(destino, ctx, origen=f"{libro.name} (sha {sha})", cfg=cfg)
   ```
8. `D.fijar_documento_base(destino)` — igual que en la opción 3, escribe la
   ruta compactada en `config.local.json`.
9. Imprime el bloque `PLANTILLA CREADA` con archivo, carpeta y la explicación
   de que ya queda fijada como el documento que se actualiza.
10. Devuelve `(0, destino)`.

---

## 8. La línea de órdenes

`ejecutar(argv)` (`src/fs_menu.py:819`) empieza llamando a
`D.preparar_consola()` —que reconfigura `stdout` y `stderr` con
`errors="replace"` para que un carácter fuera de la página de códigos no
tumbe el informe a mitad— y luego separa los argumentos:

```python
args  = [a for a in argv[1:] if not a.startswith("--")]
flags = {a.lower() for a in argv[1:] if a.startswith("--")}
```

### 8.1 Todas las banderas

La cadena de decisión es un `if/elif` **estrictamente ordenado**: si se pasan
varias banderas, gana la primera de esta lista.

| Bandera | Qué hace | ¿Necesita el libro? | Devuelve |
|---|---|---|---|
| `--refrescar` | Opción 1: actualiza el documento vivo vía `R.ejecutar(resto)` | Sí (o lo busca por convención) | `0` |
| `--generar` | Opción 2: crea la foto en `salidas\` vía `G.ejecutar(resto)` | Sí (o lo busca por convención) | `0` |
| `--plantilla` | Opción 6: crea el documento vivo donde el usuario diga | Sí, obligatorio | `0` |
| `--crear-base` | Sinónimo exacto de `--plantilla` | Sí, obligatorio | `0` |
| `--estado` | Radiografía del proyecto: `D.estado(cfg, args[0] if args else None)` | No | `0` |
| `--desbloquear` | Opción 4: quita protección y candados | No | `0` |
| `--bloquear` | Opción 5: pone candados y `proteger_salvo_datos` | No | `0` |
| `--elegir-documento` | Opción 3: abre el explorador y fija el documento base | No (si lo hay, rellena las regiones al preparar) | `0` o `1` |
| `--documento` **sin ruta detrás** | Equivale a `--elegir-documento` | No | `0` o `1` |
| `--documento RUTA.docx` | **No** captura el menú: se propaga al módulo destino (y las opciones 4 y 5 lo leen con `_opcion`) | Depende de la opción | — |
| `--consola` | Dibuja el menú de texto en vez de la ventana | Ver la nota de abajo | según la opción elegida |
| `--preparar` | Pasa de largo; la consume `refrescar_fs` | — | — |
| `--no-preparar` | Pasa de largo; la consume `refrescar_fs` | — | — |
| `--revisar` | Pasa de largo; la consume `generador_fs` | — | — |
| *(ninguna)* | Modo interactivo: ventana, y si no se puede dibujar, consola | Según la opción | según la opción |

El caso de `--documento` merece su propio comentario (`src/fs_menu.py:838-843`):

> «Ojo: refrescar_fs.py usa «--documento OTRO.docx» para apuntar a otro
> archivo. Antes el menú se quedaba con la bandera siempre, así que a través
> de EstadosFinancieros.exe era imposible refrescar un documento distinto del
> de config.json: se abría el explorador. Ahora solo la intercepta cuando
> viene suelta, sin ruta detrás.»

La distinción la hace `_opcion(argv, nombre)` (`src/fs_menu.py:58`), que
devuelve el valor que sigue a una bandera **solo si no empieza por `--`**,
para no tragarse la bandera siguiente.

### 8.2 La regla que impide colgar una automatización

```python
interactivo = eleccion is None
```

(`src/fs_menu.py:855`), con el comentario:

> «¿Vino la orden de la ventana, o de una bandera? Solo la primera debe
> terminar en una ventana de resultado: si «--bloquear» abriera una,
> cualquier uso desde un .bat o una tarea programada se quedaria colgado
> esperando a que alguien pulse «Cerrar».»

En consecuencia, **cada llamada a `ventana_resultado(...)` del archivo está
bajo `if interactivo:`**. Y como `--consola` asigna `eleccion` antes de ese
cálculo, el menú de texto tampoco abre ventana de resultado.

Antes de despachar, se reconstruye el `argv` que recibirán los módulos
delegados quitando solo las banderas propias del menú
(`src/fs_menu.py:884-890`). `--documento` se añade a esa lista **únicamente**
cuando la opción es la 3; en las opciones 1 y 2 se propaga íntegro, con su
valor.

### 8.3 `_BANDERAS_DIRECTAS`, `_sin_menu`, `_fallo` y `main`

```python
_BANDERAS_DIRECTAS = {
    "--refrescar", "--generar", "--plantilla", "--crear-base", "--estado",
    "--desbloquear", "--bloquear", "--elegir-documento", "--consola",
}
```

(`src/fs_menu.py:1001`). `_sin_menu(argv)` (`src/fs_menu.py:1007`) devuelve
`True` si alguno de los argumentos está en ese conjunto. Su docstring:

> «Importa para los errores: si la orden viene de un .bat o de una tarea
> programada, abrir una ventana la dejaría colgada esperando a que alguien
> pulse «Cerrar».»

Obsérvese que `--documento` **no** figura en `_BANDERAS_DIRECTAS`, a
diferencia del conjunto `propias`, que sí lo añade condicionalmente.

`_fallo(titulo, cuerpo, argv)` (`src/fs_menu.py:1017`) cuenta el error por los
dos canales: siempre por consola, y además en ventana **si no** se pidió una
acción por bandera. Su docstring:

> «La consola sola no basta: aparece DETRÁS de todo, así que quien pulsó un
> botón ve cerrarse la ventana y nada más. El error que de verdad importa —el
> libro abierto en Excel, el documento abierto en Word— acababa en un sitio
> donde nadie lo lee.»

`main()` (`src/fs_menu.py:1035`) envuelve `ejecutar(sys.argv)` en tres capas:

| Excepción | Título | Cuerpo |
|---|---|---|
| `ValueError` | `NO SE PUDO COMPLETAR` | `str(e)` |
| `PermissionError` | `UN ARCHIVO ESTÁ ABIERTO EN OTRO PROGRAMA` | el error, más `Cierre el documento en Word y el libro en Excel, y vuelva a intentarlo.` |
| `Exception` | `OCURRIÓ UN ERROR INESPERADO — copie este texto para soporte` | `traceback.format_exc()` |

Las tres terminan en `R._pausa()` y `sys.exit(1)`; el camino de éxito solo
llama a `R._pausa()`. Esa función (`src/refrescar_fs.py:41`) solo espera Enter
si `sys.stdin.isatty()`, con `EOFError` y `OSError` silenciados: es lo que
permite usar el ejecutable en una tarea programada sin que se quede colgado.

### 8.4 Ejemplos para automatizar

```bat
REM Cierre desatendido: refrescar el documento de config.local.json
"%PY%" src\fs_menu.py "C:\datos\cierre FS.xlsx" --refrescar

REM Refrescar OTRO documento sin tocar la configuración
"%PY%" src\fs_menu.py "C:\datos\cierre FS.xlsx" --refrescar --documento "C:\docs\EF trimestral.docx"

REM Foto desechable para una entrega puntual (queda en salidas\)
"%PY%" src\fs_menu.py "C:\datos\cierre FS.xlsx" --generar

REM Abrir el candado, editar a mano, y volver a cerrarlo
"%PY%" src\fs_menu.py --desbloquear
"%PY%" src\fs_menu.py --bloquear

REM Diagnóstico para soporte (no abre ninguna ventana)
"%PY%" src\fs_menu.py --estado

REM Menú, pero de texto: útil por escritorio remoto o sin .NET
"%PY%" src\fs_menu.py "C:\datos\cierre FS.xlsx" --consola
```

Con el ejecutable, sustituya `"%PY%" src\fs_menu.py` por
`EstadosFinancieros.exe`. Los `.bat` de la raíz hacen exactamente eso:
`estados_financieros.bat` reenvía `%*` a `src\fs_menu.py` tras localizar un
Python usable con `tools\buscar_python.bat`.

**Cuidado con el orden de los argumentos.** Como `args` recoge todo lo que no
empiece por `--`, una orden `fs_menu.py --documento OTRO.docx` sin libro deja
`args = ["OTRO.docx"]`, y `_resolver_libro` intentará tratar ese `.docx` como
si fuera el Excel. Ponga siempre el `.xlsx` primero.

---

## 9. `refrescar_fs.py` paso a paso

`src/refrescar_fs.py` (183 líneas) es la vía fácil de actualización: lo que
hay detrás de `RefrescarFS.exe` y `refrescar.bat`. Su docstring plantea el
contraste con el otro camino:

> «generador_fs.py   crea un Word NUEVO en salidas\\      (no toca nada más)
> refrescar_fs.py   ACTUALIZA el documento que ya existe (conserva el texto)»

Secuencia de `ejecutar(argv)` (`src/refrescar_fs.py:50-140`):

1. `D.preparar_consola()`.
2. Separa `args` y `flags` igual que el menú, y define una función local
   `opcion(nombre)` que —a diferencia de `_opcion` en `fs_menu`— **no**
   comprueba que el valor siguiente no empiece por `--`.
3. `cfg = G.cargar_config()`.
4. `documento = D.resolver_documento(opcion("--documento"), cfg)`.
5. `D.comprobar_escribible(documento)` — **lo primero, antes de leer siquiera
   el Excel**. El comentario: *«si Word tiene el documento abierto, no se
   toca nada. Escribir encima lo destruiría.»*
6. Resuelve el libro: `Path(args[0]).resolve()` si hay posicional, o
   `G.encontrar_excel_por_convencion(cfg)`, avisando en ese caso con
   `(Sin archivo indicado: usando '<nombre>' por convención de nombre)`. Si no
   existe, lanza `ValueError("No se encontró el libro de Excel:\n  <ruta>")`.
7. `ctx = G.leer_contexto(xlsx, cfg)`; extrae `meta = ctx.pop("_meta", {})` y
   descarta `_avisos`.
8. **`D._respaldar(documento)`**, con el comentario que documenta el fallo
   corregido (`src/refrescar_fs.py:78-80`):

   > «La copia .bak va lo primero: es la de ANTES de tocar nada, que es la
   > única que sirve para deshacer. Estaba después de preparar el documento,
   > así que respaldaba el resultado en vez del original.»

9. Preparación condicional:

   ```python
   integrado = D.clasificar_documento(documento)[0] == D.LISTO
   if "--preparar" in flags or (not integrado and "--no-preparar" not in flags):
       print(D.imprimible("Preparando el documento (solo añade lo que falte)…"))
       D.preparar(documento, ctx, cfg, respaldar=False)
   ```

   **Quién hace el `.bak` y por qué `respaldar=False`.** La firma completa es
   `D.preparar(ruta, ctx=None, cfg=None, verbose=True, respaldar=True)`
   (`src/fs_documento.py:969`): por omisión, `preparar` haría su propia copia
   de seguridad. Aquí no debe: `refrescar_fs` ya hizo el `.bak` en el paso 8,
   y ese es el único que sirve para deshacer, porque es el del documento
   **antes** de que se le montara el andamiaje. Un segundo respaldo
   machacaría el bueno con una copia del documento ya modificado. El respaldo
   es responsabilidad exclusiva del llamante.

10. Huella y refresco:

    ```python
    sha = hashlib.sha256(xlsx.read_bytes()).hexdigest()[:12]
    inf = D.refrescar(documento, ctx, origen=f"{xlsx.name} (sha {sha})", cfg=cfg)
    ```

11. Informe en consola, en un bloque de 68 caracteres titulado
    `DOCUMENTO ACTUALIZADO`, con documento, carpeta, origen (nombre más
    `sha256`), hoja usada, y cuatro contadores: filas en tablas, campos
    actualizados, cifras en el texto y zonas de redacción intactas.
12. Si no se tocó nada, un aviso explícito: *«no se actualizó NADA, y eso no
    debería pasar: el documento se prepara solo antes de refrescarlo. Si lanzó
    la orden con --no-preparar, quítelo.»*
13. **Huérfanas** (anclas presentes en el documento para las que el Excel ya
    no da valor): `" AVISO — el documento pide cifras que el Excel ya no
    tiene:"` seguido de una línea `   ? <ancla>` por cada una.
14. Ruta de la **bitácora**, si el informe la trae.
15. Hasta 20 líneas de cambios respecto del refresco anterior; si hay más,
    una línea final que remite a la bitácora del documento.
16. `return documento` — *«Para que fs_menu pueda ofrecer «Abrir el
    documento» al terminar.»*

Esa última línea es la razón de que la ventana de resultado pueda mostrar los
botones de abrir: tanto `R.ejecutar` como `G.ejecutar` devuelven la ruta del
documento producido, y `fs_menu` la pasa como `destino` a
`ventana_resultado`.

El `main()` de `refrescar_fs` (`src/refrescar_fs.py:143`) repite el esquema de
tres capas del menú, pero **imprimiendo solo por consola**: no abre ventanas,
porque el módulo se usa también como ejecutable independiente.

---

## 10. La configuración en uso

Hay dos archivos, y la diferencia entre ellos es una decisión de
funcionamiento, no de orden.

### 10.1 `config.json`, íntegro

```json
{
  "_comentario": "Todo es opcional. Borre una clave para usar su valor por defecto. 'primera_fila': un numero fija la fila de inicio; 'auto' la detecta. 'columnas': letras (A, C, E, F, G) para forzar una columna; null para detectarla por contenido.",
  "_documento_base": "Documento de Word que actualiza refrescar.bat / RefrescarFS.exe, conservando la redaccion. NO lo fije aqui: este archivo viaja por git y una ruta absoluta se le impondria a la otra maquina en cada pull. Uselo desde la opcion 3 («Cambiar el documento»), que lo guarda en config.local.json. Si aun asi quiere compartirlo, use marcadores independientes del equipo: ${ONEDRIVE}\\carpeta\\EF.docx, ${USUARIO}\\... o ${PROYECTO}\\...",
  "documento_base": "",
  "_prefijo_rangos": "Prefijo de los nombres de Excel que dan identidad estable a cada fila (fs_total_assets). Crearlos: fs_documento.py nombrar <libro> --aplicar",
  "prefijo_rangos": "fs_",
  "empresa": "Collective Mining Ltd.",
  "hoja": "FS",
  "hoja_marcadores": [
    "situación financiera",
    "statement of financial position",
    "financial position",
    "total assets",
    "total liabilities and equity",
    "assets",
    "liabilities and equity"
  ],
  "plantilla": "plantilla_estado_situacion_financiera.docx",
  "buscar_por_convencion": "FS",
  "primera_fila": "auto",
  "columnas": {
    "etiqueta": null,
    "nota": null,
    "actual": null,
    "previo": null,
    "tipo": null
  },
  "marcadores_excluir": [
    "control check",
    "check",
    "cuadre",
    "balance check"
  ]
}
```

### 10.2 `config.local.json`, íntegro

(La ruta lleva el nombre de la persona que usa ese equipo; aquí se sustituye
por `<usuario>`.)

```json
{
  "_comentario": "Ajustes de ESTE equipo. No se versiona: manda sobre config.json. Marcadores admitidos en las rutas: ${ONEDRIVE}, ${USUARIO}, ${PROYECTO}.",
  "documento_base": "${ONEDRIVE}\\DOCUMENTO DE PRUEBA DE EDICIÓN - <usuario>.docx"
}
```

### 10.3 Qué significa cada clave

| Clave | Valor aquí | Función |
|---|---|---|
| `_comentario`, `_documento_base`, `_prefijo_rangos` | texto | **Claves de documentación.** La fusión las descarta con `if k.startswith("_"): continue`. Nunca llegan al diccionario de configuración. |
| `documento_base` | `""` en el compartido; una ruta con marcador en el local | El `.docx` que actualiza la opción 1. Vacío produce un `ValueError` con instrucciones. |
| `prefijo_rangos` | `"fs_"` | Prefijo de los *nombres definidos* de Excel (`fs_total_assets`) que dan identidad estable a cada línea. Sin ellos, renombrar o insertar filas rompe el vínculo. |
| `empresa` | `"Collective Mining Ltd."` | Nombre que se inyecta en el contexto. |
| `hoja` | `"FS"` | Hoja preferida del libro. Si no existe, se elige por contenido. |
| `hoja_marcadores` | siete cadenas, en español e inglés | Señales de contenido para elegir la hoja cuando la preferida no está. |
| `plantilla` | `"plantilla_estado_situacion_financiera.docx"` | Plantilla del camino clásico (la foto). |
| `buscar_por_convencion` | `"FS"` | Subcadena que debe contener el nombre del `.xlsx` para que se le recoja en doble clic. |
| `primera_fila` | `"auto"` | `"auto"` detecta la fila de inicio; un entero la fija. |
| `columnas` | los cinco a `null` | `etiqueta`, `nota`, `actual`, `previo`, `tipo`. `null` = detectar por contenido; una letra la fuerza. Es la única clave que se **fusiona** en vez de reemplazarse, así que se puede forzar una sola columna sin declarar las otras cuatro. |
| `marcadores_excluir` | cuatro etiquetas | Filas de cuadre que no deben trasladarse al Word. |

Hay claves que no aparecen en ninguno de los dos archivos y actúan con su
valor por defecto: entre ellas `bitacora` (`"archivo"`), `apariencia_datos`
(`"boundingBox"`) y `clave_proteccion` (`"fs"`), esta última leída por el menú
como `str(cfg.get("clave_proteccion") or "fs")` en la opción 5.

### 10.4 La cascada de precedencia

`G.cargar_config()` compone la configuración en cuatro pasos, de menor a mayor
autoridad:

```
  DEFAULTS  (en generador_fs.py)
      |
      +--  config.json embebida en el .exe
      |
      +--  config.json del proyecto        (viaja por git)
      |
      +--  config.local.json               (este equipo; NO se versiona)
      |
      v
  cfg  ->  todo el sistema
```

El docstring lo justifica: *«El orden importa: lo de este equipo
(`config.local.json`) tiene que poder ganarle a lo que venga por git, o cada
«pull» reimpondria las rutas de la otra máquina.»* Un JSON con error de
sintaxis no se ignora en silencio: lanza
`ValueError(f"{ruta.name} tiene un error de sintaxis:\n  {e}")`.

### 10.5 Por qué `config.json` no lleva rutas de ninguna máquina

La respuesta está escrita dentro del propio archivo, en la clave
`_documento_base`:

> «NO lo fije aqui: este archivo viaja por git y una ruta absoluta se le
> impondria a la otra maquina en cada pull. Uselo desde la opcion 3
> («Cambiar el documento»), que lo guarda en config.local.json.»

Y hay una segunda razón, recogida en `D.resolver_documento`
(`src/fs_documento.py:2190-2195`): el `config.json` **se queda embebido en el
`.exe`** con la ruta de quien lo compiló. En cuanto el ejecutable cambia de
manos, ese `C:\Users\Fulano\…` deja de existir y todas las opciones que
dependen del documento mueren a la vez.

### 10.6 Los marcadores de ruta

Cuando aun así haya que compartir una ruta, se usan marcadores que cada
equipo resuelve a lo suyo. `G.expandir_ruta` los sustituye —y no distingue
mayúsculas de minúsculas:

| Marcador | Se sustituye por |
|---|---|
| `${ONEDRIVE}` | La raíz de OneDrive, o la carpeta personal si no hay |
| `${USUARIO}` | La carpeta personal del usuario |
| `${USERPROFILE}` | Igual que `${USUARIO}` |
| `${PROYECTO}` | La raíz del proyecto |
| `~` inicial | La carpeta personal |

`G.raiz_onedrive()` prefiere **siempre** la carpeta de empresa:

> «No sirve la variable de entorno `%OneDrive%` a secas: en un equipo con las
> dos cuentas apunta a la PERSONAL (`C:\Users\x\OneDrive`), no a la de la
> organizacion (`C:\Users\x\OneDrive - Empresa`), que es donde vive el
> documento.»

`G.compactar_ruta` es la operación inversa, y es la que aplica
`D.fijar_documento_base` al escribir en `config.local.json`. De ahí que el
`documento_base` local empiece por `${ONEDRIVE}\\`.

---

## 11. Qué pasa cuando algo falla

Los tres fallos que de verdad ocurren tienen un mensaje propio, escrito para
que el usuario sepa qué cerrar. Ninguno de los tres es un volcado de traza.

**El documento abierto en Word.** Lo detecta `D.comprobar_escribible`
(`src/fs_documento.py:512`), llamada por `refrescar_fs` antes de leer siquiera
el Excel. Comprueba si existe el archivo de bloqueo `~$<nombre>` junto al
documento e intenta abrirlo en modo `r+b`. Si no puede, pregunta a
`D.quien_bloquea` (que usa la misma maquinaria de Windows que produce el
cartel «este archivo está siendo utilizado por…») y compone:

```
El documento está abierto:
  archivo:  <nombre>
  carpeta:  <carpeta>

Quién lo tiene:
  - WINWORD.EXE   (PID 1234)   en este equipo

Ciérrelo y vuelva a ejecutar.

No se ha modificado nada. Escribir sobre un documento que Word
tiene abierto lo deja inservible, así que la operación se detiene
aquí a propósito.
```

Si ninguno de los procesos lleva la marca `en este equipo`, añade que es
probable que lo tenga abierto otra persona a través de OneDrive.

**El libro retenido por Excel.** Lo detecta `G.comprobar_legible`
(`src/generador_fs.py:933`). Su docstring dice por qué existe:

> «El libro NUNCA se escribe: solo se lee. Pero Excel lo retiene en exclusiva
> mientras lo tiene abierto —dentro de OneDrive y con el autoguardado puesto,
> casi siempre— y entonces ni siquiera se puede abrir en modo lectura. Sin
> esta comprobación, ese caso salía como un volcado de PermissionError en
> mitad de openpyxl: veinte líneas de traza que no le dicen a nadie que lo
> único que hay que hacer es cerrar Excel.»

El mensaje termina con `"El libro solo se LEE, nunca se escribe, así que no
hay nada que perder: es Excel quien no deja ni leerlo mientras lo tiene
abierto."`. Hay además un caso aparte: si el archivo figura en la carpeta
pero su contenido no se ha descargado, el aviso es que el libro *«todavía
está en la nube, no en este equipo»* y sugiere «Conservar siempre en este
dispositivo».

**Una ruta que no existe.** `D.resolver_documento` no se rinde al primer
intento. Prueba, en orden: expandir los marcadores; buscar un archivo de
nombre parecido (los nombres escritos en Office suelen colar espacios duros
U+00A0 y acentos descompuestos, invisibles al ojo pero distintos byte a
byte); buscar sin tildes; y reubicar la ruta bajo el perfil de este usuario.
Solo entonces lanza:

```
config.json apunta a un documento que no existe:
  <ruta>

Revise la clave 'documento_base'. Ojo con los espacios: un nombre
escrito en Office puede llevar espacios duros que no se ven.

Para elegirlo con el explorador:
    EstadosFinancieros.exe --elegir-documento
```

Si el libro indicado no existe, `_resolver_libro` no lanza nada: pone
`NO EXISTE: <nombre>` en la cabecera de la ventana y guarda el aviso, que se
convierte en el `ValueError` solo cuando se elige una opción que necesita
cifras (1, 2 o 6).

Por encima de todo esto, `_fallo` decide el canal: consola siempre, ventana
solo si la orden no vino de una bandera directa.

---

## Resumen del capítulo

- `fs_menu.py` es un despachador: fuera de una huella `sha256`, no abre por sí
  mismo ningún archivo de Office; toda la lógica vive en los otros tres módulos.
- La ventana es WinForms dibujada por PowerShell, no tkinter, para no arrastrar
  Tcl/Tk dentro del ejecutable; si falla, se degrada al menú de texto.
- El diálogo entre Python y la ventana es de una sola línea: `ELECCION=<n>` por
  la salida estándar, y el informe de resultado viaja por archivo.
- La cabecera declara siempre documento, Excel y estado del candado, porque sin
  esas tres líneas las opciones eran botones ciegos.
- El menú ofrece seis acciones; la 6 crea el documento vivo y la 2 una foto
  desechable, y esa es la confusión que el menú existe para evitar.
- Cada `ventana_resultado` está bajo `if interactivo:`, para que una bandera
  lanzada desde un `.bat` nunca deje un proceso esperando un clic.
- En `refrescar_fs`, el `.bak` lo hace el llamante antes de nada y por eso
  `preparar` se invoca con `respaldar=False`.
- `config.json` viaja por git y se embebe en el `.exe`: por eso las rutas de
  máquina van en `config.local.json`, que manda sobre él.
