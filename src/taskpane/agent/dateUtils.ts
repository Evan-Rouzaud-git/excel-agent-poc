import { normalizeHeader } from "./normalizeHeader";

export type ParsedDate = { date: Date | null; ts: number | null; source: string | null; fallback?: boolean };

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 24 * 60 * 60 * 1000;

function excelSerialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH_UTC + serial * MS_PER_DAY);
}

function validateParts(year: number, month: number, day: number): boolean {
  if (month < 0 || month > 11 || day < 1 || year < 1) return false;
  const d = new Date(Date.UTC(year, month, day));
  return d.getUTCFullYear() === year && d.getUTCMonth() === month && d.getUTCDate() === day;
}

function parseDdMmYyyy(text: string, opts?: { onInvalid?: (msg: string) => void }): Date | null {
  const m = text.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
  if (!m) return null;
  const day = Number(m[1]);
  const month = Number(m[2]) - 1;
  const yearStr = m?.[3];
  if (!yearStr) return null;
  let year = Number(yearStr);
  if (yearStr.length === 2) {
    year = year >= 50 ? 1900 + year : 2000 + year;
  }
  if (!validateParts(year, month, day)) {
    opts?.onInvalid?.(`date_parse_invalid:${normalizeHeader(text)}`);
    return null;
  }
  const d = new Date(Date.UTC(year, month, day));
  return Number.isNaN(d.getTime()) ? null : d;
}

function parseIso(text: string): Date | null {
  const m = text.match(/^\s*(\d{4})-(\d{2})-(\d{2})\s*$/);
  if (!m) return null;
  const year = Number(m[1]);
  const month = Number(m[2]) - 1;
  const day = Number(m[3]);
  if (!validateParts(year, month, day)) return null;
  const d = new Date(Date.UTC(year, month, day));
  return Number.isNaN(d.getTime()) ? null : d;
}

type ParseOpts = { onFallback?: (msg: string) => void; onInvalid?: (msg: string) => void; allowFallback?: boolean };

export function parseDateCell(value: any, locale = "fr-FR", opts?: ParseOpts): ParsedDate {
  if (value === null || typeof value === "undefined") return { date: null, ts: null, source: null };
  if (value instanceof Date && !Number.isNaN(value.getTime())) return { date: value, ts: value.getTime(), source: "date" };

  // Excel serial
  if (typeof value === "number" && Number.isFinite(value)) {
    const d = excelSerialToDate(value);
    return { date: d, ts: d.getTime(), source: "serial" };
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return { date: null, ts: null, source: null };
    const fr = parseDdMmYyyy(trimmed, { onInvalid: opts?.onInvalid });
    if (fr) return { date: fr, ts: fr.getTime(), source: "fr" };
    const iso = parseIso(trimmed);
    if (iso) return { date: iso, ts: iso.getTime(), source: "iso" };
    if (opts?.allowFallback) {
      const fallback = new Date(trimmed);
      if (!Number.isNaN(fallback.getTime())) {
        opts?.onFallback?.(`date_parse_fallback:${normalizeHeader(trimmed)}`);
        return { date: fallback, ts: fallback.getTime(), source: "fallback", fallback: true };
      }
    }
  }
  return { date: null, ts: null, source: null };
}
