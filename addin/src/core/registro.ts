/**
 * Bitácora de cambios entre la última versión aplicada y la nueva.
 * La foto anterior se guarda dentro del propio documento
 * (Office.context.document.settings), así que viaja con el archivo.
 */
import { Contexto } from "./tipos";

interface FilaSnap {
  etiqueta: string;
  actual: string;
  previo: string;
  tipo: string;
}
interface Snapshot {
  fecha: string;
  lineas: FilaSnap[];
}

const CLAVE = "fs_snapshot";

export function leerSnapshot(): Snapshot | null {
  const raw = Office.context.document.settings.get(CLAVE) as string | null;
  if (!raw) return null;
  try {
    return JSON.parse(raw) as Snapshot;
  } catch {
    return null;
  }
}

export async function guardarSnapshot(ctx: Contexto): Promise<void> {
  const snap: Snapshot = {
    fecha: new Date().toISOString(),
    lineas: ctx.lineas.map((l) => ({
      etiqueta: l.etiqueta,
      actual: l.actual,
      previo: l.previo,
      tipo: l.tipo,
    })),
  };
  Office.context.document.settings.set(CLAVE, JSON.stringify(snap));
  await new Promise<void>((resolve, reject) => {
    Office.context.document.settings.saveAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(r.error);
    });
  });
}

export function calcularCambios(previo: Snapshot | null, ctx: Contexto): string[] {
  if (!previo) return ["Primera actualización: no hay versión anterior con la que comparar."];

  const clave = (l: { etiqueta: string; tipo: string }) =>
    l.etiqueta || `(sin etiqueta ${l.tipo})`;
  const antes = new Map(previo.lineas.map((l) => [clave(l), l]));
  const cambios: string[] = [];

  for (const l of ctx.lineas) {
    const k = clave(l);
    const a = antes.get(k);
    if (!a) {
      cambios.push(`Nueva fila: ${k}  ${l.actual || "—"}`);
      continue;
    }
    if (a.actual !== l.actual || a.previo !== l.previo) {
      cambios.push(
        `${k}: ${a.actual || "—"} → ${l.actual || "—"}` +
          `  (comparativo ${a.previo || "—"} → ${l.previo || "—"})`
      );
    }
    antes.delete(k);
  }
  for (const k of antes.keys()) cambios.push(`Fila retirada: ${k}`);

  if (cambios.length === 0) {
    cambios.push("Sin cambios en las cifras respecto de la última actualización.");
  }
  return cambios;
}
