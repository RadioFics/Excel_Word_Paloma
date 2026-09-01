# Cómo subir el complemento de Word

Qué hay que hacer, en orden, para que a Pamela y a Violeta les aparezca el
panel en su cinta de Word.

---

## Lo primero: quién instala qué

| Quién | Qué instala | Cuándo |
|---|---|---|
| Tú (desarrollo) | Node.js | Una vez, en tu equipo |
| Pamela y Violeta | **Nada** | — |
| TI | Aprueba el manifiesto | Una vez |

Un complemento de Office **no es un programa que se instala**: es una página
web que Word carga dentro de un panel. Por eso el usuario final no descarga
nada, no necesita permisos de administrador, y no hay `.exe` que un antivirus
pueda poner en cuarentena. Es el argumento fuerte frente a TI.

**Node.js solo sirve para compilar**, igual que un compilador de C. Convierte
el TypeScript de `addin/src/` en JavaScript estático. Después de eso, Node no
vuelve a aparecer en ningún sitio.

---

## Paso 1 — Compilar (en tu equipo)

Instala Node.js 18 o superior desde [nodejs.org](https://nodejs.org) (versión
LTS). Luego:

```bash
cd addin
npm install
npm run build
```

Queda una carpeta `addin/dist/` con archivos estáticos: HTML, CSS, JS. Eso es
todo el complemento.

Para probarlo en tu propio Word antes de subir nada:

```bash
npm start
```

Levanta un servidor local en `https://localhost:3000`, instala un certificado
de desarrollo y abre Word con el complemento ya cargado. Para pararlo:
`npm run stop`.

> Antes de dar por bueno el complemento, comprueba que el formato de la tabla
> coincide con el del motor de Python: hoy el add-in aplica menos formato
> (pierde los filetes de subtotal y total). Ver `addin/README.md`.

---

## Paso 2 — Publicar los archivos en HTTPS

El manifiesto apunta a una URL. Los archivos de `dist/` tienen que estar
accesibles por HTTPS **dentro del tenant**.

La opción con menos fricción: **una biblioteca de SharePoint**.

1. Crea una biblioteca de documentos, por ejemplo `Complementos`, en el sitio
   del área financiera.
2. Sube el contenido de `addin/dist/` a una carpeta, p. ej. `fs-addin/`.
3. Anota la URL pública de esa carpeta. Tendrá esta pinta:
   ```
   https://<tenant>.sharepoint.com/sites/<sitio>/Complementos/fs-addin/
   ```
4. Comprueba en el navegador que `.../fs-addin/taskpane.html` abre.

Alternativas si TI lo prefiere: un IIS interno, o Azure Static Web Apps. Lo
único que hace falta es HTTPS y que los equipos lo alcancen. **No hay puertos
de entrada ni servidor de aplicación**: son archivos estáticos.

---

## Paso 3 — Ajustar el manifiesto

En `addin/manifest.xml`:

1. **Genera un GUID propio.** El que viene es de ejemplo y no debe usarse en
   producción:
   ```bash
   powershell -Command "[guid]::NewGuid()"
   ```
   Ponlo en `<Id>`.

2. **Sustituye todas las URLs** `https://localhost:3000` por la URL real del
   paso 2. Hay que cambiarlas todas: `SourceLocation`, `IconUrl`,
   `HighResolutionIconUrl`, `SupportUrl`, y las de `<bt:Images>` y
   `<bt:Urls>`.

3. **Añade los iconos.** Faltan `assets/icon-16.png`, `icon-32.png` e
   `icon-80.png` (PNG cuadrados). Sin ellos el complemento carga, pero sale
   sin icono en la cinta.

Comprueba que no queda ninguna referencia a localhost:

```bash
grep -n "localhost" addin/manifest.xml
```

---

## Paso 4 — Desplegarlo para las usuarias

### Opción A · Despliegue centralizado (la buena)

Lo hace un administrador de Microsoft 365:

1. Centro de administración de Microsoft 365 → **Configuración** →
   **Aplicaciones integradas**.
2. **Cargar aplicación personalizada** → **Cargar manifiesto**.
3. Sube `manifest.xml`.
4. **Asignar usuarios:** solo Pamela y Violeta (o un grupo que las contenga).
   No hace falta desplegarlo a toda la organización.
5. Aceptar y terminar.

El panel les aparece en Word en un plazo de entre unos minutos y unas 24 h,
sin que ellas hagan nada.

### Opción B · Carga lateral, para probar

Cada usuaria puede cargarlo ella misma sin pasar por TI. Sirve para validar
antes de pedir el despliegue formal:

**Word → Inicio → Complementos → Más complementos → Mis complementos →
Cargar mi complemento →** elegir el `manifest.xml`.

Se pierde al cerrar Word, y hay que repetirlo. Está bien para una prueba, no
para el uso diario.

---

## Qué pedirle a TI, exactamente

Para que la conversación sea corta, ten preparado esto:

| Pregunta que harán | Respuesta |
|---|---|
| ¿Qué se instala en los equipos? | Nada. Es una página web que Word carga en un panel. |
| ¿Dónde vive el código? | En una biblioteca de SharePoint del propio tenant. |
| ¿Sale algún dato fuera? | No. El Excel se lee en el equipo; el Word se escribe en el equipo. Sin llamadas a internet. |
| ¿Qué permisos pide? | `ReadWriteDocument`: leer y escribir **el documento abierto**. Nada más. |
| ¿Se puede auditar? | Sí. El código es TypeScript en el repositorio; el manifiesto son 100 líneas de XML. |
| ¿A quién llega? | Solo a los usuarios que se asignen en Aplicaciones integradas. |
| ¿Cómo se retira? | Se quita desde el mismo panel de administración. Inmediato. |

Si más adelante se lee el Excel directamente de OneDrive (en vez de elegir el
archivo a mano), hará falta además un consentimiento de administrador para el
permiso **`Files.Read`**, acotado al sitio correspondiente. Hoy **no** hace
falta: el complemento lee el archivo que la usuaria elija.

---

## Actualizar el complemento después

- **Cambios en el código** (`dist/`): se suben a SharePoint y ya está. Word
  recoge la versión nueva sola. **No hay que volver a tocar el manifiesto.**
- **Cambios en el manifiesto** (permisos nuevos, botones nuevos): hay que
  subir el manifiesto otra vez en Aplicaciones integradas, subiendo el número
  de `<Version>`.

Por eso conviene que el manifiesto sea estable y que la iteración vaya por el
lado de los archivos estáticos.

---

## Si el complemento falla, queda la vía local

El camino de línea de órdenes **no depende del complemento** y hace lo mismo:

```bash
refrescar.bat
```

o `RefrescarFS.exe`, arrastrándole el Excel encima. Ambos refrescan el mismo
documento con el mismo contrato de anclas
([`CONTRATO.md`](CONTRATO.md)). Si el complemento se atasca en revisión de
TI, o se cae, el trabajo no se detiene.

Los ejecutables se generan con:

```bash
powershell -ExecutionPolicy Bypass -File .\tools\hacer_exe.ps1
```

que produce `dist\GeneradorFS.exe` (Word nuevo) y `dist\RefrescarFS.exe`
(actualiza el documento base). Publícalos como *release* del repositorio con
sus huellas SHA-256.
