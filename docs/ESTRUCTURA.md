# Estructura de carpetas

Se agrupa por **naturaleza del archivo**, no por extensión. Un `.docx` puede
ser una plantilla o un ejemplo, y van a sitios distintos.

**Mantener esta división en los próximos commits.**

```
Excel_Word_Paloma/
│
├── README.md              punto de entrada
├── generar.bat            crea un Word NUEVO en salidas\
├── refrescar.bat          ACTUALIZA el documento base
├── config.json            ajustes del usuario
├── requirements.txt       dependencias de Python
│
├── src/                   codigo de la aplicacion
│   ├── generador_fs.py      lectura del Excel + generador clasico
│   ├── fs_contrato.py       vocabulario de anclas Excel<->Word
│   ├── fs_documento.py      motor de refresco en el sitio + ordenes
│   └── refrescar_fs.py      entrada amable (doble clic / arrastrar)
│
├── docs/                  documentacion
│   ├── GUIA.md              operacion paso a paso
│   ├── CONTRATO.md          especificacion de anclas
│   ├── ESTRUCTURA.md        este archivo
│   ├── DESPLIEGUE_ADDIN.md  como subir el complemento
│   ├── DIRECCION.md         direccion del proyecto y caso ante TI
│   ├── INSTALACION.md       montar el Python portable
│   └── PRUEBA_EXTERNA.md    reproducir la prueba en otro equipo
│
├── plantillas/            documentos .docx MODELO
│   ├── plantilla_base_EF.docx                       base viva, con anclas
│   └── plantilla_estado_situacion_financiera.docx   plantilla del generador clasico
│
├── ejemplos/              datos de muestra
│   └── Copia_Editable_con_columna_Tipo.xlsx
│
├── tools/                 scripts de entorno y empaquetado
│   ├── bootstrap_python.ps1   monta .\python\
│   ├── verificar.ps1          prueba de humo
│   ├── hacer_exe.ps1          genera los .exe
│   └── hacer_paquete.ps1      genera el .zip portable
│
├── addin/                 complemento de Word (TypeScript)
│
├── salidas/               generado por generar.bat  (ignorado por git)
├── python/                Python portable            (ignorado por git)
└── dist/                  .exe y .zip publicables    (ignorado por git)
```

## Dónde va cada cosa

| Si añades… | Va a |
|---|---|
| Un módulo de Python | `src/` |
| Documentación en Markdown | `docs/` |
| Una plantilla de Word | `plantillas/` |
| Un libro de ejemplo | `ejemplos/` |
| Un script de PowerShell de mantenimiento | `tools/` |
| Código del complemento | `addin/src/` |

**En la raíz solo va lo que el usuario toca directamente:** los dos `.bat`,
`config.json`, `requirements.txt` y el `README.md`.

## Rutas que dependen de esta estructura

Si se mueve algo, hay que revisar:

| Archivo | Qué asume |
|---|---|
| `src/generador_fs.py` | `BASE` es la carpeta **padre** de `src/`. Los recursos se buscan con `buscar_recurso()` en la raíz, `plantillas/` y `ejemplos/`. |
| `generar.bat`, `refrescar.bat` | llaman a `src\*.py` |
| `tools/verificar.ps1` | `src\`, `ejemplos\` |
| `tools/hacer_exe.ps1` | `--paths src`, y la plantilla en `plantillas\` |
| `tools/hacer_paquete.ps1` | la lista de carpetas a incluir |

Dentro del `.exe` todo queda plano (PyInstaller lo empaqueta en un temporal),
por eso `buscar_recurso()` prueba también la raíz.
