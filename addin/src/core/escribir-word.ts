/**
 * Escribe el Contexto EN ESTE MISMO DOCUMENTO de Word.
 *
 * No genera un archivo nuevo y no renderiza el documento entero: recorre los
 * controles de contenido del documento, y solo entra en los que pertenecen a
 * las familias de datos del contrato (ver CONTRATO.md). La prosa que la
 * persona haya escrito no se visita siquiera, así que sobrevive intacta a
 * cada actualización — y si la borra, se queda borrada.
 *
 * Los controles de datos van con LockContents = true para que nadie los pise
 * a mano. Este módulo los desbloquea justo antes de escribir y los vuelve a
 * bloquear después: es la única forma de que Office.js pueda tocarlos.
 *
 * Para preparar un documento base desde cero, use el motor de Python:
 *     python fs_documento.py construir <doc.docx>
 */
import { Contexto, Linea } from "./tipos";
import * as K from "./contrato";

/** Cómo se ve cada tipo de fila en la tabla (mismo criterio que fs_documento.py). */
function filaValores(l: Linea): string[] {
  if (l.tipo === "H") return [l.etiqueta, "", "", ""];
  if (l.tipo === "N") return [l.nota ? `${l.etiqueta} (Nota ${l.nota})` : l.etiqueta, "", "", ""];
  if (l.tipo === "S") return ["", "", l.actual, l.previo];
  return [l.etiqueta, l.nota, l.actual, l.previo];
}

export interface InformeEscritura {
  tablas: Array<{ nombre: string; filas: number }>;
  campos: number;
  datos: number;
  huerfanos: string[];
  prosaIntacta: number;
}

/**
 * Recorre el documento una sola vez y refresca todas las regiones de datos.
 * Devuelve qué se tocó y qué anclas quedaron sin dato en el Excel.
 */
export async function refrescarDocumento(ctx: Contexto): Promise<InformeEscritura> {
  const { valores } = K.construirValores(ctx);
  const informe: InformeEscritura = {
    tablas: [],
    campos: 0,
    datos: 0,
    huerfanos: [],
    prosaIntacta: 0,
  };

  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag,items/id");
    await context.sync();

    // 1. clasificar por familia antes de tocar nada
    const tablas: Array<{ cc: Word.ContentControl; nombre: string }> = [];
    const enLinea: Array<{ cc: Word.ContentControl; tag: string }> = [];

    for (const cc of ccs.items) {
      const { familia, nombre } = K.descomponer(cc.tag);
      if (familia === K.FAM_TABLA) {
        tablas.push({ cc, nombre: nombre ?? K.TABLA_PRINCIPAL });
      } else if (familia === K.FAM_CAMPO || familia === K.FAM_DATO) {
        enLinea.push({ cc, tag: cc.tag });
      } else if (familia === K.FAM_PROSA) {
        informe.prosaIntacta += 1;
      }
    }

    // 2. desbloquear todo lo que vayamos a escribir
    for (const { cc } of tablas) cc.cannotEdit = false;
    for (const { cc } of enLinea) cc.cannotEdit = false;
    await context.sync();

    // 3. campos de encabezado y cifras sueltas
    for (const { cc, tag } of enLinea) {
      const valor = valores[tag];
      if (valor === undefined) {
        informe.huerfanos.push(tag);
        continue;
      }
      cc.insertText(valor, "Replace");
      const fam = K.descomponer(tag).familia;
      if (fam === K.FAM_CAMPO) informe.campos += 1;
      else informe.datos += 1;
    }
    await context.sync();

    // 4. tablas
    for (const { cc, nombre } of tablas) {
      const lineas = K.lineasDeTabla(nombre, ctx);
      await reescribirTabla(context, cc, lineas);
      informe.tablas.push({ nombre, filas: lineas.length });
    }

    // 5. volver a bloquear
    for (const { cc } of tablas) cc.cannotEdit = true;
    for (const { cc } of enLinea) cc.cannotEdit = true;
    await context.sync();
  });

  return informe;
}

/**
 * Reescribe las filas de la tabla que hay dentro de un control, conservando
 * la fila de encabezado (y con ella los anchos y el estilo de la plantilla).
 */
async function reescribirTabla(
  context: Word.RequestContext,
  cc: Word.ContentControl,
  lineas: Linea[]
): Promise<void> {
  cc.tables.load("items");
  await context.sync();
  if (cc.tables.items.length === 0) {
    throw new Error(
      `El control "${cc.tag}" no contiene ninguna tabla. ` +
        `Prepare el documento con: python fs_documento.py construir <doc.docx>`
    );
  }

  const tabla = cc.tables.items[0];
  tabla.rows.load("items");
  await context.sync();

  for (let i = tabla.rows.items.length - 1; i >= 1; i--) {
    tabla.rows.items[i].delete();
  }
  await context.sync();

  const filas = lineas.map(filaValores);
  if (filas.length > 0) {
    tabla.addRows("End", filas.length, filas);
    await context.sync();
  }

  tabla.rows.load("items");
  await context.sync();
  for (let i = 0; i < lineas.length; i++) {
    const row = tabla.rows.items[i + 1];
    if (!row) continue;
    const t = lineas[i].tipo;
    row.font.bold = t === "S" || t === "T" || t === "H";
  }
  await context.sync();
}

/**
 * Inserta una cifra viva en la posición del cursor: crea un control de
 * contenido en línea, bloqueado y con la etiqueta del contrato. Es el botón
 * «insertar dato» del panel (equivale a `fs_documento.py insertar`).
 */
export async function insertarDato(clave: string, campo: K.CampoDato): Promise<string> {
  const tag = K.tagDato(clave, campo);
  await Word.run(async (context) => {
    const sel = context.document.getSelection();
    const cc = sel.insertContentControl();
    cc.tag = tag;
    cc.title = `${clave} (${campo})`;
    cc.appearance = "BoundingBox";
    cc.insertText("—", "Replace");
    cc.cannotEdit = true;
    await context.sync();
  });
  return tag;
}

/** Antepone un bloque de bitácora al control fs-registro (si existe). */
export async function escribirRegistro(cambios: string[], origen: string): Promise<boolean> {
  return Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag(K.TAG_REGISTRO);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) return false;

    const cc = ccs.items[0];
    cc.cannotEdit = false;
    await context.sync();

    const fecha = new Date().toLocaleString("es-CO");
    const bloque =
      `Actualización ${fecha} — origen: ${origen}\n` +
      cambios
        .slice(0, 40)
        .map((c) => `  • ${c}`)
        .join("\n") +
      "\n";
    cc.insertParagraph(bloque, "Start");
    cc.cannotEdit = true;
    await context.sync();
    return true;
  });
}
