# Dirección del proyecto y caso ante TI

## Resumen

El piloto en Python **demostró que la lógica de mapeo funciona** (columna
`Tipo`, plantilla con marcadores, modelo de tipos de fila). Lo que no encaja
en una organización que bloquea instaladores y revisa ejecutables es el
**vehículo de entrega**: un intérprete/`.exe` no gestionado, sin firma,
fuera de los controles del tenant.

Plan en tres frentes, en paralelo:

| Frente | Qué es | Horizonte | Aprobación TI |
|---|---|---|---|
| **A. Tablas vinculadas** | Pegar-vínculo de rangos de Excel en Word + "actualizar todo" | Ya | Ninguna (nativo) |
| **B. Python portable** | Paquete embebido en `.\python\`, herramienta de las 2 usuarias | Ya | Consultar AppLocker/EDR |
| **C. Add-in de Office** | Panel dentro de Word: botón "Actualizar desde Excel", **actualiza el mismo documento** | Semanas | 1 revisión + despliegue centralizado |

El requisito del cliente es explícito: **el mismo documento actualiza sus
datos**, no se genera uno nuevo.

> **Actualización.** El frente B ya cumple ese requisito. `fs_documento.py`
> refresca **en el sitio** un documento existente conservando la redacción,
> con el mismo contrato de anclas que usará el add-in
> ([`CONTRATO.md`](CONTRATO.md)). Verificado abriendo los documentos en Word:
> refresco idempotente, prosa intacta, tablas y cifras bloqueadas, y el modo
> de dos editores funcionando (`documentProtection` + rangos editables).
>
> Con eso, **el frente B deja de ser un puente y pasa a ser la solución
> operativa** mientras el frente C consigue su revisión de TI. El frente C
> ya no tiene que demostrar el concepto: solo trasladar la misma operación
> —ya especificada y probada— a un panel dentro de Word.

---

## Frente A — Tablas vinculadas (corto plazo, sin aprobaciones)

Para las 4 tablas de presentación fijas.

1. En Excel, selecciona el rango de una tabla → **Copiar**.
2. En Word, en la posición de esa tabla → **Pegar → Pegado especial →
   Pegar vínculo → Hoja de cálculo de Microsoft Excel (objeto)**.
   - Word inserta un campo `{ LINK Excel.Sheet.12 "ruta" "rango" }`.
3. Repite para las 4 tablas.
4. Actualización de todas a la vez:
   - **Archivo → Opciones → Avanzadas → General → "Actualizar vínculos
     automáticos al abrir"** = activado, **o**
   - con el documento abierto: `Ctrl+A` y luego `F9`.
5. Bloquea el resto del documento para edición
   (**Revisar → Restringir edición**) dejando libres solo las zonas de texto
   que Pamela modifica.

**Límites (por eso es puente, no solución):** el vínculo guarda la ruta del
libro; si el archivo se mueve o se renombra, hay que re-vincular. Si dentro
de la tabla se insertan/borran filas en Excel, el rango del vínculo no se
ajusta solo. Sirve para tablas de estructura estable durante el año, que es
justo el caso de las 4 de presentación.

---

## Frente B — Python portable (herramienta interina de las 2 usuarias)

Ver `INSTALACION.md`. Puntos para el caso ante TI:

- **No es una instalación:** archivos descomprimidos en la carpeta del
  proyecto. Sin admin, sin registro, sin servicio.
- **Sin red en ejecución:** el programa solo lee un `.xlsx` local y escribe
  un `.docx` local. La red se usa una única vez, al montar el entorno.
- **Trazabilidad:** cada `.docx` generado registra origen, huella SHA-256
  del Excel y fecha en sus propiedades.
- **Alcance:** 2 usuarias, 1 ruta compartida, equipos Windows x64.
- **Riesgo abierto:** AppLocker/WDAC/EDR podría frenar `python.exe` en un
  equipo nuevo. **Requiere confirmación de TI.** Si lo frena, el frente B
  queda descartado y se acelera el frente C.

Mejora recomendada antes de ampliar: externalizar la configuración
(`config.json`: nombre de hoja, mapa de columnas, empresa, lista de vínculos
hoja→plantilla) para que sumar tablas no toque el código.

---

## Frente C — Add-in de Office (solución permanente)

> Andamiaje v0 ya en el repositorio: carpeta `addin/` (manifiesto, panel y
> lógica de mapeo portada de `generador_fs.py`). Sin compilar todavía;
> necesita un equipo con Node.js. Ver `addin/README.md`.

### Experiencia de usuario

Pamela abre **su documento de Word** (el mismo de siempre) → panel lateral
"Estados financieros" → botón **Actualizar desde Excel** → el add-in
reescribe las tablas marcadas **en ese mismo documento** → `Ctrl+S`. Antes
de aplicar, el panel muestra un resumen de diferencias
("Efectivo 81.370.000 → 72.957.812; nueva fila 'IVA por cobrar'").

### Componentes

1. **Add-in de panel para Word** — `manifest.xml` + archivos estáticos
   (HTML/JS/TS) servidos desde una URL HTTPS interna (una biblioteca de
   SharePoint sirve como hosting estático, o IIS interno). Sin infraestructura
   de entrada.
2. **Acceso a las cifras del Excel.** Opción recomendada: el modelo grande
   vive en SharePoint/OneDrive y el add-in lee los **rangos con nombre** vía
   Microsoft Graph (`/workbook/names/{nombre}/range`), con permiso delegado
   **de solo lectura** (`Files.Read`) consentido una vez por un administrador
   y acotado al sitio correspondiente.
   Alternativa: un add-in complementario en Excel con un botón "Publicar
   cifras" que deja un JSON compacto de las filas `Tipo` en una lista de
   SharePoint o en el XML personalizado del documento.
3. **Modelo de mapeo:** el mismo de hoy. El add-in lee las filas etiquetadas
   con `Tipo` y arma la misma lista de líneas.
4. **Escritura en Word.** La plantilla lleva **controles de contenido**
   etiquetados (p. ej. `tag="fs-tabla-principal"`, `tag="nota-ar"`). El
   add-in localiza cada tabla por su etiqueta y reescribe sus filas con la
   API `Word.Table`. Documento intacto en todo lo demás; se guarda en su
   sitio.
5. **Informe de diferencias:** el add-in compara contra la última versión
   aplicada (guardada en el XML personalizado del propio documento) y la
   muestra antes de confirmar.

### Gobernanza (el argumento ante TI)

- **Despliegue centralizado:** el manifiesto se publica en *Aplicaciones
  integradas* del centro de administración de M365, acotado a Pamela y
  Violeta. Nada que instalar por equipo.
- **Datos en el tenant:** Graph de solo lectura, acotado por sitio,
  consentido por administrador. Sin egreso externo. Auditable en los
  registros de Office / inicios de sesión de Graph.
- **Código revisable:** TypeScript en un repositorio que TI puede auditar;
  una sola revisión de código + manifiesto.
- **Hosting:** archivos estáticos; escala trivialmente; sin puertos de
  entrada.

### Escalabilidad verificable

- Cada tipo de tabla = un control de contenido etiquetado + una entrada en
  el mapa `Tipo`→plantilla. Sumar tablas es lineal, sin infraestructura.
- Volumen real (un juego de estados financieros: decenas de tablas, cientos
  de filas) está muy por debajo de cualquier límite de Office.js o de Graph.
- Pruebas de "archivo dorado": un `.xlsx` de referencia y un `.docx`
  esperado; el CI compara el texto resultante. Es una afirmación medible.

### Coste

- v1 (1 tabla de presentación + 1 nota): unas semanas de desarrollo
  (TypeScript, Office.js, `yo office`, empaquetado con esbuild/webpack).
- Requiere un **dueño de mantenimiento nombrado** y un ciclo de revisión de
  TI.
- Extender al resto de tablas: incremental, sobre la misma base.

---

## Decisiones abiertas (confirmar con TI y con Pamela)

1. ¿Los documentos (Excel modelo y Word modelo) pueden vivir en
   SharePoint/OneDrive? — habilita Graph y el add-in.
2. ¿Hay AppLocker/WDAC/EDR que frene ejecutables no firmados desde carpetas
   de usuario o unidades de red? — decide si el frente B es viable.
3. ¿Power Automate está licenciado para las 2 usuarias? — alternativa de
   menor esfuerzo al add-in si la UX "panel en Word" no fuera imprescindible
   (aunque no cumple "mismo documento" tan limpio como el add-in).
4. ¿Quién es el dueño de mantenimiento de plantillas y del add-in a un año?
5. Dado que las 4 tablas + notas exceden los 3 días iniciales, ¿se amplía el
   encargo o se entrega el frente A + este documento de dirección?
