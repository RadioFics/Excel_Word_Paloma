/**
 * Lee un .xlsx (ArrayBuffer) y devuelve el Contexto para la plantilla.
 * Detecta la hoja y las columnas por CONTENIDO, igual que generador_fs.py.
 */
import ExcelJS from "exceljs";
import {
  Contexto,
  Linea,
  Tipo,
  num,
  texto,
  norm,
  esNumero,
  inferirTipo,
} from "./tipos";

export interface Ajustes {
  empresa: string;
  hoja: string;
  hojaMarcadores: string[];
  primeraFila: number | "auto";
  columnas: Partial<Record<"etiqueta" | "nota" | "actual" | "previo" | "tipo", string>>;
  marcadoresExcluir: string[];
  maxFilasScan: number;
  maxColsScan: number;
}

export const AJUSTES_DEFECTO: Ajustes = {
  empresa: "Collective Mining Ltd.",
  hoja: "FS",
  hojaMarcadores: [
    "situación financiera",
    "statement of financial position",
    "financial position",
    "total assets",
    "total liabilities and equity",
    "assets",
    "liabilities and equity",
  ],
  primeraFila: "auto",
  columnas: {},
  marcadoresExcluir: ["control check", "check", "cuadre", "balance check"],
  maxFilasScan: 400,
  maxColsScan: 16,
};

const TIPOS_VALIDOS = new Set<Tipo>(["H", "I", "S", "T", "N"]);

function letraACol(x?: string): number | null {
  if (!x) return null;
  let n = 0;
  for (const ch of x.trim().toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n || null;
}

function colALetra(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function valorPlano(v: any): unknown {
  if (v == null) return null;
  if (typeof v === "object") {
    if ("result" in v) return valorPlano(v.result); // fórmula
    if ("richText" in v) return v.richText.map((t: any) => t.text).join("");
    if ("text" in v) return v.text; // hipervínculo
    if (v instanceof Date) return v;
  }
  return v;
}

function esTextoNoVacio(v: unknown): boolean {
  return typeof v === "string" && v.trim() !== "";
}

interface Grid {
  valores: unknown[][]; // [fila][col], base 1
  negrita: boolean[][];
  nFilas: number;
}

function materializar(ws: ExcelJS.Worksheet, aj: Ajustes): Grid {
  const maxF = Math.min(aj.maxFilasScan, ws.rowCount || aj.maxFilasScan);
  const maxC = aj.maxColsScan;
  const valores: unknown[][] = [Array(maxC + 1).fill(null)];
  const negrita: boolean[][] = [Array(maxC + 1).fill(false)];
  for (let r = 1; r <= maxF; r++) {
    const row = ws.getRow(r);
    const vrow: unknown[] = Array(maxC + 1).fill(null);
    const brow: boolean[] = Array(maxC + 1).fill(false);
    for (let c = 1; c <= maxC; c++) {
      const cell = row.getCell(c);
      vrow[c] = valorPlano(cell.value);
      brow[c] = Boolean(cell.font && cell.font.bold);
    }
    valores.push(vrow);
    negrita.push(brow);
  }
  let nFilas = valores.length - 1;
  while (nFilas > 1 && valores[nFilas].slice(1).every((v) => v === null || v === "")) nFilas--;
  return { valores, negrita, nFilas };
}

function elegirHoja(wb: ExcelJS.Workbook, aj: Ajustes): { ws: ExcelJS.Worksheet; como: string } {
  const exacta = wb.worksheets.find((w) => w.name === aj.hoja);
  if (exacta) return { ws: exacta, como: `'${aj.hoja}' por nombre exacto` };

  const marcadores = aj.hojaMarcadores.map(norm).filter(Boolean);
  const conv = norm(aj.hoja);
  let mejor: ExcelJS.Worksheet | null = null;
  let mejorScore = 0;
  for (const w of wb.worksheets) {
    let textoHoja = "";
    for (let r = 1; r <= Math.min(60, w.rowCount || 60); r++) {
      const row = w.getRow(r);
      for (let c = 1; c <= 8; c++) {
        const v = valorPlano(row.getCell(c).value);
        if (esTextoNoVacio(v)) textoHoja += " " + norm(v);
      }
    }
    let score = marcadores.reduce((acc, m) => acc + (textoHoja.includes(m) ? 1 : 0), 0);
    if (conv && norm(w.name).includes(conv)) score += 1;
    if (score > mejorScore) {
      mejor = w;
      mejorScore = score;
    }
  }
  if (mejor && mejorScore >= 2) {
    return { ws: mejor, como: `'${mejor.name}' por contenido (${mejorScore} señales)` };
  }
  throw new Error(
    `No pude identificar la hoja del Estado de Situación Financiera. ` +
      `Indique el nombre exacto en los ajustes (hoja).`
  );
}

function detectarColumnas(g: Grid, aj: Ajustes) {
  const maxC = aj.maxColsScan;
  const col: Record<string, number | null> = {
    etiqueta: letraACol(aj.columnas.etiqueta),
    nota: letraACol(aj.columnas.nota),
    actual: letraACol(aj.columnas.actual),
    previo: letraACol(aj.columnas.previo),
    tipo: letraACol(aj.columnas.tipo),
  };

  for (let r = 1; r <= Math.min(8, g.nFilas); r++) {
    for (let c = 1; c <= maxC; c++) {
      const t = norm(g.valores[r][c]);
      if (!t) continue;
      if (col.tipo == null && (t === "tipo" || t === "type")) col.tipo = c;
      if (col.nota == null && ["note", "nota", "notes", "notas"].includes(t)) col.nota = c;
    }
  }

  const perfil: Record<number, [number, number]> = {};
  for (let c = 1; c <= maxC; c++) {
    let txt = 0;
    let n = 0;
    for (let r = 2; r <= g.nFilas; r++) {
      const v = g.valores[r][c];
      if (esNumero(v)) n++;
      else if (esTextoNoVacio(v)) txt++;
    }
    perfil[c] = [txt, n];
  }

  if (col.actual == null || col.previo == null) {
    const conNums = Object.keys(perfil)
      .map(Number)
      .filter((c) => perfil[c][1] >= 2)
      .sort((a, b) => perfil[b][1] - perfil[a][1] || b - a);
    const elegidas = conNums.slice(0, 2).sort((a, b) => a - b);
    if (elegidas.length >= 2) {
      col.actual = col.actual ?? elegidas[0];
      col.previo = col.previo ?? elegidas[1];
    } else if (elegidas.length === 1) {
      col.actual = col.actual ?? elegidas[0];
    }
  }

  if (col.etiqueta == null) {
    const limite = col.actual ?? maxC + 1;
    const textuales: number[] = [];
    for (let c = 1; c < limite; c++) if (perfil[c][0] >= 3) textuales.push(c);
    col.etiqueta = textuales[0] ?? 1;
  }

  if (col.nota == null && col.actual) {
    for (let c = (col.etiqueta as number) + 1; c < col.actual; c++) {
      const [txt] = perfil[c];
      let enterosPeq = 0;
      for (let r = 2; r <= g.nFilas; r++) {
        const v = g.valores[r][c];
        if (esNumero(v) && Number.isInteger(v) && v > 0 && v <= 99) enterosPeq++;
      }
      if (enterosPeq >= 1 && txt <= 2) {
        col.nota = c;
        break;
      }
    }
  }

  if (!col.actual) {
    throw new Error("No pude identificar las columnas de cifras. Fíjelas en los ajustes (columnas).");
  }
  return col as { etiqueta: number; nota: number | null; actual: number; previo: number | null; tipo: number | null };
}

function detectarRegion(
  g: Grid,
  cols: { etiqueta: number; actual: number; previo: number | null },
  aj: Ajustes
): [number, number] {
  let primera: number;
  if (typeof aj.primeraFila === "number" && aj.primeraFila > 0) {
    primera = aj.primeraFila;
  } else {
    primera = 5;
    for (let r = 1; r <= g.nFilas; r++) {
      const et = g.valores[r][cols.etiqueta];
      const va = g.valores[r][cols.actual];
      const vp = cols.previo ? g.valores[r][cols.previo] : null;
      if (!esTextoNoVacio(et)) continue;
      if (["", "$"].includes(norm(et))) continue;
      if (r > 4 || esNumero(va) || esNumero(vp)) {
        primera = r;
        break;
      }
    }
  }
  let ultima = primera;
  for (let r = primera; r <= g.nFilas; r++) {
    if (
      esTextoNoVacio(g.valores[r][cols.etiqueta]) ||
      esNumero(g.valores[r][cols.actual]) ||
      (cols.previo && esNumero(g.valores[r][cols.previo]))
    ) {
      ultima = r;
    }
  }
  return [primera, ultima];
}

export async function leerContexto(buf: ArrayBuffer, aj: Ajustes = AJUSTES_DEFECTO): Promise<Contexto> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);

  const { ws, como } = elegirHoja(wb, aj);
  const g = materializar(ws, aj);
  const cols = detectarColumnas(g, aj);
  const [primera, ultima] = detectarRegion(g, cols, aj);

  const ce = cols.etiqueta;
  const ca = cols.actual;
  const cp = cols.previo;
  const cn = cols.nota;
  const ct = cols.tipo;

  const hdr = (c: number | null, off: number) => {
    const r = 1 + off;
    return c && r < primera && r < g.valores.length ? g.valores[r][c] : null;
  };
  const primerTexto = (c: number) => {
    for (let r = 1; r < Math.max(primera, 2); r++) if (esTextoNoVacio(g.valores[r][c])) return g.valores[r][c] as string;
    return "";
  };

  const lineas: Linea[] = [];
  let nDeclarados = 0;
  let nInferidos = 0;

  for (let r = primera; r <= ultima; r++) {
    const vals = g.valores[r];
    const crudo = ct ? vals[ct] : null;
    const declarado =
      crudo !== null && crudo !== undefined && String(crudo).trim() !== ""
        ? String(crudo).trim().toUpperCase()
        : "";

    if (declarado === "X") continue;

    let tipo: Tipo;
    let origen: "declarado" | "inferido";
    let senal = "";

    const celdas = {
      etiqueta: vals[ce],
      nota: cn ? vals[cn] : null,
      actual: ca ? vals[ca] : null,
      previo: cp ? vals[cp] : null,
      negrita: Boolean((ca && g.negrita[r][ca]) || (cp && g.negrita[r][cp]) || g.negrita[r][ce]),
    };

    if (TIPOS_VALIDOS.has(declarado as Tipo)) {
      tipo = declarado as Tipo;
      origen = "declarado";
    } else {
      const inf = inferirTipo(celdas, aj.marcadoresExcluir);
      if (inf.tipo == null || inf.tipo === "X") continue;
      tipo = inf.tipo;
      origen = "inferido";
      senal = inf.senal;
    }

    lineas.push({
      fila: r,
      tipo,
      etiqueta: String(vals[ce] ?? ""),
      nota: cn ? texto(vals[cn]) : "",
      actual: ca ? num(vals[ca]) : "",
      previo: cp ? num(vals[cp]) : "",
      // sin formatear: los necesita contrato.ts para var_abs / var_pct
      actualRaw: ca ? vals[ca] : null,
      previoRaw: cp ? vals[cp] : null,
      origen,
      senal,
    });
    origen === "declarado" ? nDeclarados++ : nInferidos++;
  }

  if (lineas.length === 0) {
    throw new Error(
      `Identifiqué la hoja (${ws.name}, ${como}) pero no obtuve líneas ` +
        `en la región filas ${primera}–${ultima}. Revise los ajustes (columnas / primeraFila).`
    );
  }

  return {
    empresa: aj.empresa,
    titulo: primerTexto(ce) || "",
    fechaActual: String(hdr(ca, 0) ?? ""),
    fechaPrevia: String(hdr(cp, 0) ?? ""),
    miles: String(hdr(ca, 1) ?? ""),
    estadoActual: texto(hdr(ca, 2)),
    estadoPrevio: texto(hdr(cp, 2)),
    moneda: String(hdr(ca, 3) ?? ""),
    lineas,
    meta: {
      hoja: ws.name,
      comoHoja: como,
      columnas: {
        etiqueta: colALetra(ce),
        nota: cn ? colALetra(cn) : "—",
        actual: colALetra(ca),
        previo: cp ? colALetra(cp) : "—",
        tipo: ct ? colALetra(ct) : "—",
      },
      region: [primera, ultima],
      hayColTipo: ct != null,
      nDeclarados,
      nInferidos,
    },
  };
}
