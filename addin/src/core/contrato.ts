/**
 * contrato.ts — Vocabulario de anclas entre el Excel y el documento de Word.
 *
 * Es el puerto exacto de fs_contrato.py. Si cambia uno, cambia el otro.
 * La especificación legible está en CONTRATO.md (raíz del repositorio).
 *
 * Familias:
 *   fs-tabla-<nombre>        bloque: una tabla completa        -> se refresca
 *   fs-campo-<nombre>        en línea: campo de encabezado     -> se refresca
 *   fs-dato-<clave>-<campo>  en línea: cifra dentro de prosa   -> se refresca
 *   fs-prosa-<nombre>        bloque: redacción libre           -> NO se toca
 *   fs-registro              bloque: bitácora                  -> se antepone
 *   fs-meta                  bloque oculto: foto JSON          -> se sobrescribe
 */
import { Contexto, Linea } from "./tipos";

export const PREFIJO = "fs";
export const FAM_TABLA = "tabla";
export const FAM_CAMPO = "campo";
export const FAM_DATO = "dato";
export const FAM_PROSA = "prosa";

export const TAG_REGISTRO = "fs-registro";
export const TAG_META = "fs-meta";

export const TABLA_PRINCIPAL = "principal";

export const CAMPOS_ENCABEZADO = [
  "empresa",
  "titulo",
  "fecha_actual",
  "fecha_previa",
  "miles",
  "moneda",
  "estado_actual",
  "estado_previo",
] as const;

export const CAMPOS_DATO = ["actual", "previo", "nota", "var_abs", "var_pct"] as const;

export type CampoDato = (typeof CAMPOS_DATO)[number];

export const tagTabla = (nombre: string) => `${PREFIJO}-${FAM_TABLA}-${nombre}`;
export const tagCampo = (nombre: string) => `${PREFIJO}-${FAM_CAMPO}-${nombre}`;
export const tagDato = (clave: string, campo: CampoDato) =>
  `${PREFIJO}-${FAM_DATO}-${clave}-${campo}`;
export const tagProsa = (nombre: string) => `${PREFIJO}-${FAM_PROSA}-${nombre}`;

export interface Ancla {
  familia: string | null;
  nombre: string | null;
  campo: string | null;
}

/** 'fs-dato-total_assets-actual' -> { familia:'dato', nombre:'total_assets', campo:'actual' } */
export function descomponer(tag: string | null | undefined): Ancla {
  const vacio: Ancla = { familia: null, nombre: null, campo: null };
  if (!tag) return vacio;
  const t = tag.trim();
  if (t === TAG_REGISTRO) return { familia: "registro", nombre: null, campo: null };
  if (t === TAG_META) return { familia: "meta", nombre: null, campo: null };

  const partes = t.split("-");
  if (partes.length < 3 || partes[0] !== PREFIJO) return vacio;
  const familia = partes[1];

  if (familia === FAM_DATO) {
    if (partes.length < 4) return vacio;
    const campo = partes[partes.length - 1];
    if (!(CAMPOS_DATO as readonly string[]).includes(campo)) return vacio;
    return { familia: FAM_DATO, nombre: partes.slice(2, -1).join("-"), campo };
  }
  if (familia === FAM_TABLA || familia === FAM_CAMPO || familia === FAM_PROSA) {
    return { familia, nombre: partes.slice(2).join("-"), campo: null };
  }
  return vacio;
}

/** ¿El refresco debe escribir dentro de esta ancla? */
export function esRegionDeDatos(tag: string | null | undefined): boolean {
  const f = descomponer(tag).familia;
  return f === FAM_TABLA || f === FAM_CAMPO || f === FAM_DATO;
}

/**
 * Deriva la clave estable de una etiqueta del Excel.
 * DEBE dar el mismo resultado que fs_contrato.clave() en Python:
 *   1. quitar tildes (NFKD sin diacríticos)  2. minúsculas
 *   3. [^a-z0-9] -> '_'                      4. colapsar y recortar '_'
 *   5. cortar a 40 y recortar de nuevo
 */
// Marcas diacríticas combinantes (U+0300–U+036F). Se construye con RegExp
// desde una cadena para que el rango no viaje como caracteres literales en
// el archivo fuente, que es frágil ante recodificaciones.
const DIACRITICOS = new RegExp("[\\u0300-\\u036f]", "g");

export function clave(etiqueta: unknown): string {
  if (etiqueta === null || etiqueta === undefined) return "";
  const s = String(etiqueta)
    .normalize("NFKD")
    .replace(DIACRITICOS, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
  return s.slice(0, 40).replace(/_+$/g, "");
}

/**
 * La clave con la que se identifica una línea en el documento.
 * Prefiere el rango con nombre de Excel (identidad estable ante renombrados);
 * si no lo hay, cae en la clave derivada del texto de la etiqueta.
 */
export function claveDeLinea(l: Linea): string {
  return l.clave || clave(l.etiqueta);
}

/** Formato contable, igual que tipos.num(): negativos entre paréntesis. */
function numContable(v: number): string {
  const s = Math.abs(Math.round(v)).toLocaleString("en-US");
  return v < 0 ? `(${s})` : s;
}

export interface ValoresConstruidos {
  valores: Record<string, string>;
  colisiones: Array<{ nueva: string; primera: string; clave: string }>;
}

/**
 * Del contexto del Excel al mapa {tag: texto} de campos y cifras sueltas.
 * No incluye las tablas: esas se arman desde ctx.lineas.
 */
export function construirValores(ctx: Contexto): ValoresConstruidos {
  const valores: Record<string, string> = {};
  const colisiones: ValoresConstruidos["colisiones"] = [];
  const vistas = new Map<string, string>();

  const porNombre: Record<string, unknown> = {
    empresa: ctx.empresa,
    titulo: ctx.titulo,
    fecha_actual: ctx.fechaActual,
    fecha_previa: ctx.fechaPrevia,
    miles: ctx.miles,
    moneda: ctx.moneda,
    estado_actual: ctx.estadoActual,
    estado_previo: ctx.estadoPrevio,
  };
  for (const nombre of CAMPOS_ENCABEZADO) {
    valores[tagCampo(nombre)] = String(porNombre[nombre] ?? "");
  }

  for (const l of ctx.lineas) {
    const k = claveDeLinea(l);
    if (!k) continue;
    if (vistas.has(k)) {
      colisiones.push({ nueva: l.etiqueta, primera: vistas.get(k)!, clave: k });
      continue;
    }
    vistas.set(k, l.etiqueta);

    valores[tagDato(k, "actual")] = l.actual ?? "";
    valores[tagDato(k, "previo")] = l.previo ?? "";
    valores[tagDato(k, "nota")] = l.nota ?? "";

    const a = l.actualRaw;
    const p = l.previoRaw;
    if (typeof a === "number" && typeof p === "number") {
      valores[tagDato(k, "var_abs")] = numContable(a - p);
      valores[tagDato(k, "var_pct")] = p ? `${(((a - p) / p) * 100).toFixed(1)}%` : "";
    } else {
      valores[tagDato(k, "var_abs")] = "";
      valores[tagDato(k, "var_pct")] = "";
    }
  }
  return { valores, colisiones };
}

/** Lista legible de cifras disponibles, para el panel: [clave, etiqueta, actual, previo]. */
export function catalogo(ctx: Contexto): Array<[string, string, string, string]> {
  const out: Array<[string, string, string, string]> = [];
  const vistas = new Set<string>();
  for (const l of ctx.lineas) {
    const k = claveDeLinea(l);
    if (!k || vistas.has(k)) continue;
    vistas.add(k);
    out.push([k, l.etiqueta, l.actual, l.previo]);
  }
  return out;
}

/**
 * Qué líneas alimentan la tabla `fs-tabla-<nombre>`.
 * 'principal' -> todas. Otro nombre -> solo la sección cuyo encabezado (H)
 * produce esa misma clave.
 */
export function lineasDeTabla(nombre: string, ctx: Contexto): Linea[] {
  if (nombre === TABLA_PRINCIPAL) return [...ctx.lineas];
  const sel: Linea[] = [];
  let dentro = false;
  for (const l of ctx.lineas) {
    if (l.tipo === "H") {
      dentro = clave(l.etiqueta) === nombre;
      if (dentro) sel.push(l);
      continue;
    }
    if (dentro) sel.push(l);
  }
  return sel;
}
