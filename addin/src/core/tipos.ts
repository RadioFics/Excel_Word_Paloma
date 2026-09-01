/**
 * Modelo de tipos de fila e inferencia. Es el mismo criterio que usa
 * generador_fs.py (Python), portado a TypeScript para el add-in.
 */

export type Tipo = "H" | "I" | "S" | "T" | "N" | "X";

export interface Linea {
  fila: number;
  tipo: Tipo;
  etiqueta: string;
  nota: string;
  actual: string; // ya formateado en estilo contable
  previo: string;
  /** Valor sin formatear, para calcular variaciones. Ver contrato.ts. */
  actualRaw?: unknown;
  previoRaw?: unknown;
  origen: "declarado" | "inferido";
  senal: string;
}

export interface Contexto {
  empresa: string;
  titulo: string;
  fechaActual: string;
  fechaPrevia: string;
  miles: string;
  estadoActual: string;
  estadoPrevio: string;
  moneda: string;
  lineas: Linea[];
  meta: {
    hoja: string;
    comoHoja: string;
    columnas: Record<string, string>;
    region: [number, number];
    hayColTipo: boolean;
    nDeclarados: number;
    nInferidos: number;
  };
}

const MARC_CERO = new Set(["-", "–", "—", "−", "ꟷ", "‑", "·", "0"]);

export function norm(s: unknown): string {
  return String(s ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

export function esNumero(v: unknown): v is number {
  return typeof v === "number" && !Number.isNaN(v);
}

export function marcadorCero(v: unknown): boolean {
  return typeof v === "string" && MARC_CERO.has(v.trim());
}

/** Formato contable: negativos entre paréntesis, sin decimales. */
export function num(v: unknown): string {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "string") return v;
  if (typeof v !== "number" || Number.isNaN(v)) return String(v);
  const s = Math.abs(Math.round(v)).toLocaleString("en-US");
  return v < 0 ? `(${s})` : s;
}

export function texto(v: unknown): string {
  if (v === null || v === undefined) return "";
  return String(v).replace(/^[()\s]+|[()\s]+$/g, "");
}

export interface Celdas {
  etiqueta: unknown;
  nota: unknown;
  actual: unknown;
  previo: unknown;
  negrita: boolean;
}

export function inferirTipo(
  c: Celdas,
  marcadoresExcluir: string[]
): { tipo: Tipo | null; senal: string } {
  const etTxt =
    typeof c.etiqueta === "string"
      ? c.etiqueta.trim()
      : c.etiqueta == null
      ? ""
      : String(c.etiqueta);
  const etN = norm(etTxt);
  const notaN = norm(c.nota);
  const hayValor =
    esNumero(c.actual) ||
    esNumero(c.previo) ||
    marcadorCero(c.actual) ||
    marcadorCero(c.previo);
  const tieneNota =
    esNumero(c.nota) || (typeof c.nota === "string" && c.nota.trim() !== "");

  for (const m of marcadoresExcluir) {
    const mm = norm(m);
    if (mm && (etN.includes(mm) || notaN.includes(mm))) {
      return { tipo: "X", senal: "fila de control/cuadre" };
    }
  }

  if (!etTxt && !hayValor && !tieneNota) return { tipo: null, senal: "fila vacía" };

  if (etN.startsWith("total")) return { tipo: "T", senal: "empieza por 'Total'" };

  if (!etTxt && hayValor) return { tipo: "S", senal: "cifras sin etiqueta (subtotal)" };

  if (etTxt && !hayValor) {
    if (
      etTxt === etTxt.toUpperCase() ||
      etTxt.endsWith(":") ||
      c.negrita ||
      etTxt.length <= 40
    ) {
      return { tipo: "H", senal: "etiqueta sin cifras (encabezado)" };
    }
    return { tipo: "N", senal: "texto largo sin cifras (nota)" };
  }

  if (etTxt && (hayValor || tieneNota)) {
    return { tipo: "I", senal: "etiqueta con cifras o nota" };
  }

  return { tipo: null, senal: "sin señales suficientes" };
}
