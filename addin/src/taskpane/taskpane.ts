/**
 * Panel del add-in. Flujo:
 *   1. El usuario elige el .xlsx modelo (de su OneDrive sincronizado).
 *   2. Se parsea en el navegador (exceljs) -> Contexto, con vista previa.
 *   3. "Aplicar": se reescriben las tablas de ESTE documento, se calcula
 *      la bitácora contra la última versión y se guarda la nueva foto.
 *
 * Mapa de etiquetas (Tag) de controles de contenido en la plantilla:
 *   fs-tabla-principal  -> la tabla del estado (4 columnas + encabezado)
 *   fs-registro         -> sección de bitácora (opcional)
 *   fs-fecha-actual / fs-fecha-previa / fs-miles / fs-moneda / fs-titulo
 */
import { Contexto } from "../core/tipos";
import { leerContexto, AJUSTES_DEFECTO } from "../core/leer-excel";
import { actualizarTabla, actualizarCampos, escribirRegistro } from "../core/escribir-word";
import { leerSnapshot, guardarSnapshot, calcularCambios } from "../core/registro";

const TAG_TABLA = "fs-tabla-principal";
const TAG_REGISTRO = "fs-registro";

let ctxActual: Contexto | null = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    document.body.innerHTML = "<p style='padding:16px'>Este complemento solo funciona en Word.</p>";
    return;
  }
  (document.getElementById("archivo") as HTMLInputElement).addEventListener("change", onArchivo);
  (document.getElementById("aplicar") as HTMLButtonElement).addEventListener("click", onAplicar);
});

async function onArchivo(ev: Event) {
  const input = ev.target as HTMLInputElement;
  const file = input.files?.[0];
  const resumen = document.getElementById("resumen")!;
  const btn = document.getElementById("aplicar") as HTMLButtonElement;
  if (!file) return;

  resumen.className = "resumen";
  resumen.textContent = "Leyendo el Excel…";
  btn.disabled = true;
  ctxActual = null;

  try {
    const buf = await file.arrayBuffer();
    const ctx = await leerContexto(buf, AJUSTES_DEFECTO);
    ctxActual = ctx;

    const previo = leerSnapshot();
    const cambios = calcularCambios(previo, ctx);

    const m = ctx.meta;
    resumen.className = "resumen ok";
    resumen.innerHTML =
      `Hoja <b>${m.hoja}</b> (${m.comoHoja}).<br>` +
      `Columnas: etiqueta ${m.columnas.etiqueta} · nota ${m.columnas.nota} · ` +
      `actual ${m.columnas.actual} · previo ${m.columnas.previo} · tipo ${m.columnas.tipo}.<br>` +
      `Región filas ${m.region[0]}–${m.region[1]}. ` +
      `<b>${ctx.lineas.length}</b> líneas (${m.nDeclarados} declaradas, ${m.nInferidos} inferidas).` +
      (m.hayColTipo ? "" : "<br><span style='color:var(--warn)'>La hoja no tiene columna 'Tipo': todos los tipos se infirieron. Revisa la vista previa.</span>");

    pintarDetalle(ctx, cambios);
    btn.disabled = false;
  } catch (e: any) {
    resumen.className = "resumen err";
    resumen.textContent = "No se pudo leer: " + (e?.message ?? String(e));
  }
}

function pintarDetalle(ctx: Contexto, cambios: string[]) {
  const det = document.getElementById("detalle") as HTMLDetailsElement;
  det.hidden = false;
  document.getElementById("cambios")!.textContent =
    "Cambios respecto de la última actualización:\n" + cambios.map((c) => "  • " + c).join("\n");

  const tb = document.querySelector("#tabla-preview tbody")!;
  tb.innerHTML = "";
  for (const l of ctx.lineas) {
    const tr = document.createElement("tr");
    if (l.origen === "inferido") tr.className = "inferido";
    tr.innerHTML =
      `<td>${l.fila}</td><td>${l.tipo}</td><td>${escape(l.etiqueta)}</td>` +
      `<td>${escape(l.nota)}</td><td class="num">${escape(l.actual)}</td>` +
      `<td class="num">${escape(l.previo)}</td>` +
      `<td>${l.origen === "inferido" ? "inferido — " + escape(l.senal) : "declarado"}</td>`;
    tb.appendChild(tr);
  }
}

async function onAplicar() {
  const ctx = ctxActual;
  const estado = document.getElementById("estado")!;
  const btn = document.getElementById("aplicar") as HTMLButtonElement;
  if (!ctx) return;

  btn.disabled = true;
  estado.className = "estado";
  estado.textContent = "Actualizando el documento…";

  try {
    await actualizarCampos({
      "fs-titulo": ctx.titulo,
      "fs-fecha-actual": ctx.fechaActual,
      "fs-fecha-previa": ctx.fechaPrevia,
      "fs-miles": ctx.miles,
      "fs-moneda": ctx.moneda,
    });

    const n = await actualizarTabla(TAG_TABLA, ctx);

    const previo = leerSnapshot();
    const cambios = calcularCambios(previo, ctx);
    const conRegistro = await escribirRegistro(TAG_REGISTRO, cambios);
    await guardarSnapshot(ctx);

    estado.className = "estado ok";
    estado.textContent =
      `Listo: ${n} filas actualizadas en la tabla. ` +
      (conRegistro ? "Bitácora añadida. " : "") +
      "Guarda el documento (Ctrl+S).";
  } catch (e: any) {
    estado.className = "estado err";
    estado.textContent = "Error: " + (e?.message ?? String(e));
  } finally {
    btn.disabled = false;
  }
}

function escape(s: string): string {
  return String(s ?? "").replace(/[&<>"]/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c] as string)
  );
}
