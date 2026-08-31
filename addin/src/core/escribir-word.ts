/**
 * Escribe el Contexto EN ESTE MISMO DOCUMENTO de Word.
 * No genera un archivo nuevo: localiza los controles de contenido
 * etiquetados en la plantilla y reemplaza su contenido.
 *
 * Preparación de la plantilla (una vez, en Word):
 *   - Pestaña "Programador" (si no aparece: Archivo > Opciones > Personalizar cinta > Programador).
 *   - Seleccione la tabla del estado -> "Control de contenido de texto enriquecido".
 *   - Propiedades del control -> Etiqueta (Tag) = "fs-tabla-principal".
 *   - Repita para otras tablas ("fs-tabla-nota-ar", etc.).
 *   - Para el registro de cambios: un control de contenido con Etiqueta = "fs-registro".
 *   - Para campos de encabezado: controles con Etiqueta = "fs-fecha-actual", "fs-fecha-previa", ...
 *
 * La tabla dentro de "fs-tabla-principal" debe tener 4 columnas:
 *   Concepto | Nota | Periodo actual | Periodo comparativo
 * y una fila de encabezado (que se conserva).
 */
import { Contexto, Linea } from "./tipos";

function filaValores(l: Linea): string[] {
  if (l.tipo === "H") return [l.etiqueta, "", "", ""];
  if (l.tipo === "N") return [l.nota ? `${l.etiqueta} (Nota ${l.nota})` : l.etiqueta, "", "", ""];
  if (l.tipo === "S") return ["", "", l.actual, l.previo];
  return [l.etiqueta, l.nota, l.actual, l.previo];
}

/** Reescribe la tabla marcada con `tag`. Devuelve el nº de filas escritas. */
export async function actualizarTabla(tag: string, ctx: Contexto): Promise<number> {
  return Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();

    if (ccs.items.length === 0) {
      throw new Error(
        `No encontré un control de contenido con la etiqueta "${tag}". ` +
          `Marque la tabla en la plantilla (ver escribir-word.ts).`
      );
    }

    const cc = ccs.items[0];
    cc.tables.load("items");
    await context.sync();
    if (cc.tables.items.length === 0) {
      throw new Error(`El control "${tag}" no contiene ninguna tabla.`);
    }

    const tabla = cc.tables.items[0];
    tabla.rows.load("items");
    await context.sync();

    // conservar la fila de encabezado; borrar el resto
    for (let i = tabla.rows.items.length - 1; i >= 1; i--) {
      tabla.rows.items[i].delete();
    }
    await context.sync();

    const filas = ctx.lineas.map(filaValores);
    if (filas.length > 0) {
      tabla.addRows("End", filas.length, filas);
      await context.sync();
    }

    // formato por tipo
    tabla.rows.load("items");
    await context.sync();
    for (let i = 0; i < ctx.lineas.length; i++) {
      const row = tabla.rows.items[i + 1];
      if (!row) continue;
      const t = ctx.lineas[i].tipo;
      row.font.bold = t === "S" || t === "T" || t === "H";
    }
    await context.sync();

    return ctx.lineas.length;
  });
}

/** Reemplaza el texto de controles de contenido de encabezado por su tag. */
export async function actualizarCampos(valores: Record<string, string>): Promise<void> {
  return Word.run(async (context) => {
    for (const [tag, val] of Object.entries(valores)) {
      const ccs = context.document.contentControls.getByTag(tag);
      ccs.load("items");
      await context.sync();
      for (const cc of ccs.items) cc.insertText(val ?? "", "Replace");
    }
    await context.sync();
  });
}

/** Antepone un bloque de bitácora al control "fs-registro" (si existe). */
export async function escribirRegistro(tag: string, cambios: string[]): Promise<boolean> {
  return Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) return false;

    const fecha = new Date().toLocaleString("es-CO");
    const bloque =
      `Actualización ${fecha}\n` + cambios.map((c) => `  • ${c}`).join("\n") + "\n\n";
    ccs.items[0].insertParagraph(bloque, "Start");
    await context.sync();
    return true;
  });
}
