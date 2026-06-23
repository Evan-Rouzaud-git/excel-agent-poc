import { AgentPlan, PlanStep } from "./types";
import { normalizeHeader } from "./normalizeHeader";

export type CanonicalFilterOp =
  | "gt"
  | "gte"
  | "lt"
  | "lte"
  | "eq"
  | "neq"
  | "contains"
  | "not_contains"
  | "in"
  | "isEmpty"
  | "notEmpty"
  | "between";

function clone<T>(obj: T): T {
  return JSON.parse(JSON.stringify(obj));
}

function normalizeOpToken(raw: string): string {
  if (!raw) return "";
  const decomposed = raw.normalize("NFKD");
  let stripped = "";
  for (const ch of decomposed) {
    if (ch.charCodeAt(0) >= 0x0300 && ch.charCodeAt(0) <= 0x036f) continue; // drop diacritics
    stripped += ch;
  }
  stripped = stripped.replace(/[’']/g, "'").replace(/[–—]/g, "-");
  stripped = stripped.replace(/[^a-zA-Z0-9<>=\-_'\s]/g, " ");
  const lower = stripped.toLowerCase();
  return lower
    .split(/\s+/)
    .filter(Boolean)
    .join(" ")
    .trim();
}

export function normalizeFilterOp(opRaw: string, warnings?: string[]): CanonicalFilterOp | null {
  const raw = (opRaw || "").toString();
  const trimmed = raw.trim();
  // strict symbol handling (never fuzzy)
  const symbolMap: Record<string, CanonicalFilterOp> = {
    "=": "eq",
    "==": "eq",
    "!=": "neq",
    "<>": "neq",
    "<=": "lte",
    ">=": "gte",
    "<": "lt",
    ">": "gt",
  };
  if (symbolMap[trimmed] && trimmed.length <= 2) {
    const mapped = symbolMap[trimmed];
    warnings?.push?.(`filter_op_normalized raw="${raw}" norm="${trimmed}" -> ${mapped} (method=symbol)`);
    return mapped;
  }

  const token = normalizeOpToken(raw);
  if (!token) return null;

  const table: Record<CanonicalFilterOp, string[]> = {
    gt: [">", "gt", "greater", "greater than", "above", "over", "superieur", "sup a", "plus grand", "plus que", "superieur a"],
    gte: [">=", "gte", "ge", "at least", "minimum", "min", "superieur ou egal", "au moins", "superieur ou egal a"],
    lt: ["<", "lt", "less", "below", "under", "inferieur", "moins que", "inferieur a"],
    lte: ["<=", "lte", "le", "at most", "maximum", "max", "inferieur ou egal", "au plus", "inferieur ou egal a"],
    eq: ["=", "==", "eq", "egal", "equals", "is", "egal a"],
    neq: ["!=", "<>", "neq", "not equal", "different", "diff", "pas egal"],
    contains: ["contains", "contient", "contiens", "comprend"],
    not_contains: ["not contains", "does not contain", "doesn't contain", "not_contains", "notcontains", "contient pas", "ne contient pas", "n contient pas", "ne contient plus"],
    in: ["in", "dans", "parmi"],
    isEmpty: ["is empty", "isempty", "is_empty", "vide", "empty"],
    notEmpty: ["not empty", "notempty", "not_empty", "non vide", "pas vide"],
    between: ["between", "entre"],
  };

  const lookup: Record<string, CanonicalFilterOp> = {};
  Object.entries(table).forEach(([canon, aliases]) => {
    aliases.forEach((a) => {
      lookup[normalizeOpToken(a)] = canon as CanonicalFilterOp;
    });
  });

  const direct = lookup[token];
  if (direct) return direct;

  const tokens = token.split(" ").filter(Boolean);
  const synonymsSup = ["superieur", "supereur", "sup"];
  const synonymsInf = ["inferieur", "infereur"];
  const synonymsEgal = ["egal", "egale", "egaux", "egux", "egau"];
  const synonymsDiff = ["different", "differents", "differente"];
  const hasAny = (frags: string[]) => tokens.some((t) => frags.some((f) => t.includes(f)));
  const hasSup = hasAny(synonymsSup) || hasAny(["plus"]) || token.includes("au dessus") || token.includes("minimum") || token.includes("au moins");
  const hasInf = hasAny(synonymsInf) || hasAny(["moins"]) || token.includes("au dessous") || token.includes("maximum") || token.includes("au plus");
  const hasEgal = hasAny(synonymsEgal) || token.includes("egal");
  const hasDiff = hasAny(synonymsDiff) || token.includes("pas egal") || token.includes("pas egale");
  const hasNeg = tokens.some((t) => ["pas", "non", "sans"].includes(t));
  const hasContain = tokens.some((t) => t.includes("contient") || t.includes("contiens") || t.includes("contientp"));

  const chosenOp = (() => {
    if (hasContain && hasNeg) return "not_contains";
    if (hasContain) return "contains";
    if ((hasSup && hasInf) || (hasEgal && hasDiff)) return "ambiguous" as const;
    if (hasEgal && (hasNeg || hasDiff)) return "neq";
    if (hasEgal && hasSup) return "gte";
    if (hasEgal && hasInf) return "lte";
    if (hasSup) return "gt";
    if (hasInf) return "lt";
    if (hasEgal) return "eq";
    return null;
  })();
  if (chosenOp === "ambiguous") {
    warnings?.push?.(`ambiguous_filter_op:${raw}`);
    return null;
  }
  if (chosenOp) {
    warnings?.push?.(`filter_op_normalized raw="${raw}" norm="${token}" -> ${chosenOp} (method=token-fuzzy)`);
    return chosenOp as CanonicalFilterOp;
  }

  warnings?.push?.(`unsupported_filter_op:${raw}`);
  return null;
}

const isoFromYear = (year: number) => `${year.toString().padStart(4, "0")}-01-01`;
const excelSerialToISO = (serial: number) => {
  const base = new Date(Date.UTC(1899, 11, 30));
  base.setUTCDate(base.getUTCDate() + Math.floor(serial));
  return base.toISOString().slice(0, 10);
};

function normalizeDateValue(op: CanonicalFilterOp, value: any): { op: CanonicalFilterOp; value: any } {
  const asNum = typeof value === "number" ? value : Number((value || "").toString());
  if (Number.isFinite(asNum) && `${value}`.trim().length === 4) {
    const year = asNum;
    if (op === "gt") return { op: "gte", value: isoFromYear(year + 1) };
    if (op === "lt") return { op, value: isoFromYear(year) };
    if (op === "lte") return { op: "lt", value: isoFromYear(year + 1) };
    return { op, value: isoFromYear(year) };
  }
  if (Number.isFinite(asNum) && asNum >= 20000 && asNum <= 80000) {
    return { op, value: excelSerialToISO(asNum) };
  }
  if (typeof value === "string" && /^\d{4}$/.test(value.trim())) {
    const year = Number(value.trim());
    if (op === "gt") return { op: "gte", value: isoFromYear(year + 1) };
    if (op === "lt") return { op, value: isoFromYear(year) };
    if (op === "lte") return { op: "lt", value: isoFromYear(year + 1) };
    return { op, value: isoFromYear(year) };
  }
  return { op, value };
}
function normalizeRef(ref: any, keepExtras = false): any {
  if (!ref) return undefined;
  if (typeof ref === "string") return { blockRef: ref };
  if (typeof ref === "object") {
    const out: any = keepExtras ? { ...ref } : {};
    if (typeof ref.blockRef === "string") out.blockRef = ref.blockRef;
    if (typeof ref.artifactRef === "string") out.artifactRef = ref.artifactRef;
    if (typeof ref.tableName === "string") out.tableName = ref.tableName;
    if (typeof ref.sheetName === "string") out.sheetName = ref.sheetName;
    return out;
  }
  return undefined;
}

function canonicalizeSort(sort: any) {
  if (!sort) return undefined;
  if (typeof sort === "string") return { col: sort, dir: "asc" as const };
  if (Array.isArray(sort) && sort.length >= 1) {
    const col = sort[0];
    const dir = typeof sort[1] === "string" && sort[1].toLowerCase() === "desc" ? "desc" : "asc";
    return { col, dir };
  }
  if (typeof sort === "object" && sort.col) {
    const dir = typeof sort.dir === "string" && sort.dir.toLowerCase() === "desc" ? "desc" : "asc";
    return { col: sort.col, dir };
  }
  return undefined;
}

function canonicalizeFilter(filter: any) {
  if (!filter) return [];
  const arr = Array.isArray(filter) ? filter : [filter];
  const clean = arr
    .map((f) => {
      if (!f || typeof f !== "object") return null;
      const col = f.col || f.column || f.field;
      const op = (f.op || f.operator || "").toString().trim();
      if (!col || !op) return null;
      const entry: any = { col, op };
      const opLower = op.toLowerCase();
      const looksDateOp = opLower.includes("apres") || opLower.includes("avant") || opLower.includes("date");
      if (typeof f.value !== "undefined") entry.value = f.value;
      if (typeof f.type !== "undefined") entry.type = f.type;
      if (!entry.type && looksDateOp) entry.type = "date";
      return entry;
    })
    .filter(Boolean) as any[];
  return clean;
}

function canonicalizeStep(step: PlanStep, warnings?: string[]): PlanStep {
  const params = { ...(step.params || {}) };
  if (params.source) params.source = normalizeRef(params.source, true);
  if (params.target) params.target = normalizeRef(params.target, true);
  if (params.anchor) params.anchor = normalizeRef(params.anchor, true);
  if (params.dest && typeof params.dest === "object") {
    const dest: any = { ...params.dest };
    if (dest.anchor) dest.anchor = normalizeRef(dest.anchor, true);
    params.dest = dest;
  }
  if (params.sort) params.sort = canonicalizeSort(params.sort);
  if (params.filter) {
    const filt = canonicalizeFilter(params.filter).map((f: any) => {
      const opCanon = normalizeFilterOp(f.op, warnings);
      if (!opCanon) {
        warnings?.push?.(`unsupported_filter_op:${f.op}`);
        return null;
      }
      let value = f.value;
      let type = f.type;
      if (
        type === "date" ||
        (typeof value === "string" && /^\d{4}(-\d{2}(-\d{2})?)?$/.test(value.trim())) ||
        (typeof value === "number" && value >= 1000)
      ) {
        type = "date";
        const normalized = normalizeDateValue(opCanon, value);
        value = normalized.value;
        if (normalized.op !== opCanon) {
          warnings?.push?.(`filter_op_normalized_date:${opCanon}->${normalized.op}`);
        }
        return { ...f, op: normalized.op, value, type };
      }
      return { ...f, op: opCanon, type, value };
    }).filter(Boolean);
    params.filter = filt;
  }
  return { ...step, params };
}

export function canonicalizePlan(plan: AgentPlan, warnings?: string[]): AgentPlan {
  if (!plan || typeof plan !== "object") return plan;
  const copy: AgentPlan = clone(plan);
  // ensure required top-level defaults
  (copy as any).version = (copy as any).version || "1.0";
  (copy as any).goal = (copy as any).goal || "User request";
  copy.steps = (copy.steps || []).map((s) => canonicalizeStep(s, warnings));
  return copy;
}

export function headerMatch(a?: string | null, b?: string | null): boolean {
  return normalizeHeader(a) === normalizeHeader(b);
}
