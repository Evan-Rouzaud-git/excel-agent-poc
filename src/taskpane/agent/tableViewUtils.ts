import { WorkbookContextSnapshot } from "../context/types";
import { normalizeFilterOp } from "./canonicalizePlan";
import { findBlock, letterToCol } from "./utils";
import { normalizeHeader } from "./normalizeHeader";

type TableViewWarning =
  | "table_view_select_unresolved"
  | "table_view_select_partial"
  | "table_view_filter_dropped"
  | "table_view_filter_col_unresolved"
  | "table_view_sort_dropped"
  | "table_view_sort_col_unresolved";

export const normalizeHeaderText = normalizeHeader;

function buildHeaderAliasLookup(aliases?: Record<string, string[]>) {
  const map = new Map<string, string[]>();
  if (!aliases) return map;
  Object.entries(aliases).forEach(([key, values]) => {
    if (!key) return;
    const normKey = normalizeHeaderText(key);
    if (!normKey || !Array.isArray(values)) return;
    const entry = map.get(normKey) || [];
    values.forEach((raw) => {
      if (typeof raw !== "string") return;
      const trimmed = raw.trim();
      if (!trimmed) return;
      if (!entry.includes(trimmed)) entry.push(trimmed);
    });
    if (entry.length) map.set(normKey, entry);
  });
  return map;
}

function buildHeaderLookup(headers: string[]) {
  const map = new Map<string, string>();
  headers.forEach((header) => {
    if (!header) return;
    const norm = normalizeHeaderText(header);
    if (norm && !map.has(norm)) map.set(norm, header);
  });
  return map;
}

function resolveHeaderFromAlias(
  headerLookup: Map<string, string>,
  aliasLookup?: Map<string, string[]>,
  token?: string
): string | null {
  if (!token || !aliasLookup || aliasLookup.size === 0) return null;
  const normalized = normalizeHeaderText(token);
  if (!normalized) return null;
  const aliasValues = aliasLookup.get(normalized);
  if (!aliasValues?.length) return null;
  for (const alias of aliasValues) {
    const aliasNorm = normalizeHeaderText(alias);
    if (!aliasNorm) continue;
    const candidate = headerLookup.get(aliasNorm);
    if (candidate) return candidate;
  }
  return null;
}

function columnIndexFromToken(token: string): number | null {
  const trimmed = (token || "").trim();
  if (!trimmed) return null;
  const colLetter = trimmed.match(/^[A-Za-z]{1,3}$/);
  if (colLetter) return letterToCol(colLetter[0]);
  const colNum = trimmed.match(/^col(?:onne)?[-_\s]?([0-9]+)$/i);
  if (colNum) {
    const idx = Number(colNum[1]) - 1;
    return Number.isInteger(idx) && idx >= 0 ? idx : null;
  }
  const hashNum = trimmed.match(/^#?([0-9]+)$/);
  if (hashNum) {
    const idx = Number(hashNum[1]) - 1;
    return Number.isInteger(idx) && idx >= 0 ? idx : null;
  }
  return null;
}

export function matchHeaderToken(headers: string[], token?: string): { header: string; index: number; score: number; via: string } | null {
  if (!token || !headers || headers.length === 0) return null;
  const headerInfos = headers.map((h, i) => ({ header: h ?? "", norm: normalizeHeaderText(h ?? ""), index: i }));
  const normToken = normalizeHeaderText(token);
  const exact = headerInfos.find((h) => h.norm === normToken);
  if (exact) return { header: exact.header || "", index: exact.index, score: 1, via: "exact" };

  const colIdx = columnIndexFromToken(token);
  if (colIdx !== null && colIdx >= 0 && colIdx < headers.length) {
    return { header: headers[colIdx] || "", index: colIdx, score: 0.8, via: "index" };
  }
  return null;
}

function getSourceHeaders(context: WorkbookContextSnapshot | undefined, source: any): string[] {
  if (!context) return [];
  const blockRef = source?.blockRef;
  if (blockRef) {
    const found = findBlock(blockRef, context);
    if (found.block?.headers?.length) return found.block.headers;
  }
  // fallback: aggregate headers from all blocks (helps artifactRef or runtime tables)
  const anyHeaders =
    (context.sheets || [])
      .flatMap((s) => (s.blocks || []).flatMap((b) => b.headers || []))
      .filter((h) => typeof h === "string") || [];
  return anyHeaders;
}

function sanitizeSelect(
  selectRaw: any,
  headers: string[],
  warnings: string[],
  headerAliasLookup?: Map<string, string[]>,
  context?: WorkbookContextSnapshot,
  allowTokens: string[] = []
): string[] {
  const tokens = Array.isArray(selectRaw) ? selectRaw : typeof selectRaw === "string" ? [selectRaw] : [];
  const cleanedTokens = tokens.filter((t) => typeof t === "string" && (t as string).trim());
  if (cleanedTokens.length === 0) return [];
  const resolved: string[] = [];
  const seen = new Set<string>();
  const headerLookup = buildHeaderLookup(headers);

  cleanedTokens.forEach((tok) => {
    if (allowTokens.includes(tok)) {
      const norm = normalizeHeaderText(tok);
      if (!seen.has(norm)) {
        resolved.push(tok);
        seen.add(norm);
      }
      return;
    }
    let candidate: string | null = null;
    if (headers.length) {
      const match = matchHeaderToken(headers, tok);
      if (match) candidate = match.header;
    }
    if (!candidate) {
      candidate = resolveHeaderFromAlias(headerLookup, headerAliasLookup, tok);
    }
    if (!candidate) {
      warnings.push(`unknown_header:${tok}|headers=${headers.join(",")}`);
      return;
    }
    const norm = normalizeHeaderText(candidate);
    if (!seen.has(norm)) {
      resolved.push(candidate);
      seen.add(norm);
    }
  });
  if (resolved.length === 0) warnings.push("table_view_select_unresolved");
  return resolved;
}

function sanitizeRename(renameRaw: any, select: string[], warnings?: string[]): { map: Record<string, string>; provided: boolean } {
  const out: Record<string, string> = {};
  const provided = !!renameRaw && typeof renameRaw === "object" && !Array.isArray(renameRaw);
  if (!provided) return { map: out, provided: false };
  Object.entries(renameRaw as Record<string, any>).forEach(([k, v]) => {
    if (!k || typeof v !== "string") return;
    out[k] = v;
  });
  return { map: out, provided };
}

function buildRenameAliasLookup(rename?: Record<string, string>): Map<string, string> {
  const map = new Map<string, string>();
  if (!rename) return map;
  Object.entries(rename).forEach(([actual, alias]) => {
    if (!alias) return;
    const norm = normalizeHeaderText(alias);
    if (norm && !map.has(norm)) map.set(norm, actual);
  });
  return map;
}

function resolveToOutputHeader(
  col: string,
  select: string[],
  rename: Record<string, string>,
  outputHeaders: string[],
  renameAliasLookup?: Map<string, string>,
  headerAliasLookup?: Map<string, string[]>
): string | null {
  if (!col) return null;
  const matchOutput = matchHeaderToken(outputHeaders, col);
  if (matchOutput) {
    const val = outputHeaders[matchOutput.index];
    if (typeof val === "string") {
      const actual = renameAliasLookup?.get(normalizeHeaderText(val) || "") || null;
      return actual || val;
    }
  }
  const headerLookup = buildHeaderLookup(outputHeaders);
  const aliasResolved = resolveHeaderFromAlias(headerLookup, headerAliasLookup, col);
  if (aliasResolved) return aliasResolved;

  const selectMatch = matchHeaderToken(select, col);
  if (selectMatch) {
    const key = selectMatch.header ?? "";
    const candidate = (key && rename[key]) || key;
    if (!candidate) return null;
    const aliasActual = renameAliasLookup?.get(normalizeHeaderText(candidate) || "") || key;
    return aliasActual || null;
  }

  return null;
}

function resolveHeaderFromList(
  headers: string[],
  token: string,
  headerAliasLookup?: Map<string, string[]>
): string | null {
  if (!token || !headers.length) return null;
  const match = matchHeaderToken(headers, token);
  if (match) return match.header;
  const lookup = buildHeaderLookup(headers);
  return resolveHeaderFromAlias(lookup, headerAliasLookup, token);
}

function normalizeOpSymbol(opRaw: string) {
  const op = (opRaw || "").trim();
  const lower = op.toLowerCase();
  if (["=", "==", "eq"].includes(lower)) return "eq";
  if (["!=", "<>", "neq"].includes(lower)) return "neq";
  if (["notempty", "not empty"].includes(lower)) return "notEmpty";
  if (["isempty", "is empty"].includes(lower)) return "isEmpty";
  if (lower === ">" || lower === "gt") return "gt";
  if (lower === ">=" || lower === "gte") return "gte";
  if (lower === "<" || lower === "lt") return "lt";
  if (lower === "<=" || lower === "lte") return "lte";
  if (lower === "contains") return "contains";
  if (lower === "notcontains" || lower === "not_contains" || lower === "not contain") return "not_contains";
  if (lower === "between") return "between";
  if (lower === "in") return "in";
  return op || lower;
}

const allowedFilterTypes = new Set(["number", "text", "date"]);
function mapFilterType(type?: any): string | undefined {
  if (typeof type !== "string") return undefined;
  const normalized = type.trim().toLowerCase();
  if (normalized === "string") return "text";
  return allowedFilterTypes.has(normalized) ? normalized : undefined;
}

function sanitizeFilters(
  filterRaw: any,
  select: string[],
  rename: Record<string, string>,
  outputHeaders: string[],
  warnings: string[],
  renameAliasLookup?: Map<string, string>,
  headerAliasLookup?: Map<string, string[]>,
  fallbackHeaders: string[] = []
): { col: string; op: string; value?: any; type?: string }[] {
  if (!filterRaw) return [];
  if (!Array.isArray(filterRaw)) {
    warnings.push("table_view_filter_invalid");
    return [];
  }
  const normalized: any[] = [];
  filterRaw.forEach((f) => {
    if (!f || typeof f !== "object") {
      warnings.push("table_view_filter_invalid");
      return;
    }
    const col = (f.col || f.column || f.field || "").toString().trim() || (f.col as any);
    const opRaw = (f.op || f.operator || "").toString();
    const opNorm = normalizeFilterOp(opRaw) || normalizeOpSymbol(opRaw);
    let value = (f as any).value ?? (f as any).val ?? (f as any).values ?? (f as any).vals;
    if ((f as any).type === "number" && typeof value === "string" && value.trim() !== "" && !Number.isNaN(Number(value))) {
      value = Number(value);
    }
    if (opNorm === "in" && typeof value === "string") {
      value = value
        .split(/[;,]/)
        .map((v: string) => v.trim())
        .filter(Boolean);
    }
    const effectiveOp = opNorm === "neq" && (value === undefined || value === null || value === "") ? "notEmpty" : opNorm;
      let colResolved =
        resolveToOutputHeader(col, select, rename, outputHeaders, renameAliasLookup, headerAliasLookup) ||
        resolveToOutputHeader(col, Object.keys(rename), rename, outputHeaders, renameAliasLookup, headerAliasLookup) ||
        (select.length === 0 ? col : null);
      if (!colResolved) {
        const fallbackMatch = resolveHeaderFromList(fallbackHeaders, col, headerAliasLookup);
        if (fallbackMatch) colResolved = fallbackMatch;
      }
      if (!colResolved) {
        warnings.push(`unknown_header:${col}|headers=${outputHeaders.join(",")}`);
        warnings.push("table_view_filter_col_unresolved");
        return;
      }
      const entry: { col: string; op: string; value?: any; type?: string } = { col: colResolved, op: effectiveOp, value };
      const normalizedType = mapFilterType((f as any).type);
      if (normalizedType) entry.type = normalizedType;
      normalized.push(entry);
    });
    return normalized;
  }

function sanitizeSort(
  sortRaw: any,
  select: string[],
  rename: Record<string, string>,
  outputHeaders: string[],
  warnings: string[],
  renameAliasLookup?: Map<string, string>,
  headerAliasLookup?: Map<string, string[]>,
  fallbackHeaders: string[] = [],
  allowTokens: string[] = []
): { col: string; dir: "asc" | "desc" } | undefined {
  if (!sortRaw) return undefined;
  let sortObj: any = sortRaw;
  if (typeof sortRaw === "string") {
    const parts = sortRaw.split(/[ ,;]/).filter(Boolean);
    sortObj = { col: parts[0], dir: parts[1] || "asc" };
  }
  if (typeof sortObj !== "object") {
    warnings.push("table_view_sort_dropped");
    return undefined;
  }
  const colRaw = (sortObj.col || sortObj.column || sortObj.field || "").toString();
  const col = colRaw || (allowTokens.includes("$lastAddedColumn") ? "$lastAddedColumn" : "");
  if (!col) {
    warnings.push("table_view_sort_dropped");
    return undefined;
  }
  const dirRaw = (sortObj.dir || sortObj.direction || "").toString().toLowerCase();
  const dir: "asc" | "desc" = dirRaw === "desc" || dirRaw === "descending" ? "desc" : "asc";
  let match =
    resolveToOutputHeader(col, select, rename, outputHeaders, renameAliasLookup, headerAliasLookup) ||
    resolveToOutputHeader(col, Object.keys(rename), rename, outputHeaders, renameAliasLookup, headerAliasLookup) ||
    (allowTokens.includes(col) ? col : null);
  if (!match) {
    const fallbackMatch = resolveHeaderFromList(fallbackHeaders, col, headerAliasLookup);
    if (fallbackMatch) match = fallbackMatch;
  }
  if (!match) {
    warnings.push(`unknown_header:${col}|headers=${outputHeaders.join(",")}`);
    warnings.push("table_view_sort_col_unresolved");
    return undefined;
  }
  return { col: match, dir };
}

function sanitizeDest(destRaw: any) {
  const dest = destRaw && typeof destRaw === "object" ? { ...destRaw } : {};
  const allowedModes = new Set(["right", "below", "newSheet", "inPlace"]);
  if (!dest.mode || !allowedModes.has(dest.mode)) dest.mode = "newSheet";
  if (dest.anchor && typeof dest.anchor === "object") dest.anchor = { ...dest.anchor };
  return dest;
}

export function sanitizeTableViewParams(
  params: any,
  context?: WorkbookContextSnapshot,
  warningsOut?: string[],
  options?: { allowStyle?: boolean; allowTokens?: string[]; extraHeaders?: string[]; sourceHeaders?: string[]; headerAliases?: Record<string, string[]> }
) {
  const warnings = warningsOut || [];
  const headers =
    (options?.sourceHeaders && options.sourceHeaders.length ? options.sourceHeaders : getSourceHeaders(context, params?.source || {})) || [];
  const headersWithExtras = [...headers, ...(options?.extraHeaders || [])];
  const headerAliasLookup = buildHeaderAliasLookup(options?.headerAliases);
  let select = sanitizeSelect(
    params?.select,
    headersWithExtras,
    warnings,
    headerAliasLookup,
    context,
    options?.allowTokens || []
  );

  const renameRes = sanitizeRename(params?.rename, select, warnings);
  const rename = renameRes.map;
  const aliasLookup = buildRenameAliasLookup(rename);
  const outputHeaders = select.map((h) => (rename[h] ? rename[h] : h));
  const remapAliasToActual = (value?: string) => {
    if (!value) return value;
    const norm = normalizeHeaderText(value);
    if (!norm) return value;
    return aliasLookup.get(norm) || value;
  };
  let filter = sanitizeFilters(
    params?.filter,
    select,
    rename,
    outputHeaders,
    warnings,
    aliasLookup,
    headerAliasLookup,
    headersWithExtras
  );
  let sort = sanitizeSort(
    params?.sort,
    select,
    rename,
    outputHeaders,
    warnings,
    aliasLookup,
    headerAliasLookup,
    headersWithExtras,
    options?.allowTokens || []
  );
  if (aliasLookup && aliasLookup.size) {
    filter = filter.map((entry) => {
      const normalizedCol = remapAliasToActual(entry.col);
      if (normalizedCol && normalizedCol !== entry.col) {
        return { ...entry, col: normalizedCol, column: normalizedCol, field: normalizedCol };
      }
      return entry;
    });
    if (sort) {
      const normalizedSort = remapAliasToActual(sort.col);
      if (normalizedSort && normalizedSort !== sort.col) {
        sort = { ...sort, col: normalizedSort };
      }
    }
  }
  const tokenSet = new Set((options?.allowTokens || []).filter(Boolean));
  const selectSet = new Set<string>();
  select.forEach((col) => {
    const norm = normalizeHeader(col);
    if (norm) selectSet.add(norm);
  });
  const ensureSelectContains = (col?: string) => {
    if (!col) return;
    if (tokenSet.has(col)) return;
    const norm = normalizeHeader(col);
    if (!norm || selectSet.has(norm)) return;
    select.push(col);
    selectSet.add(norm);
  };
  if (sort?.col) ensureSelectContains(sort.col);
  filter.forEach((entry) => ensureSelectContains(entry.col));
  const dest = sanitizeDest(params?.dest);
  const isRenameOnly =
    renameRes.provided &&
    (!params?.filter || (Array.isArray(params.filter) && params.filter.length === 0)) &&
    (!params?.sort || Object.keys(params.sort || {}).length === 0) &&
    (!params?.select || select.length === 0);
  const isSortOnly =
    !renameRes.provided &&
    (!params?.filter || (Array.isArray(params.filter) && params.filter.length === 0)) &&
    !!params?.sort &&
    (!params?.select || select.length === 0);
  if (isRenameOnly || isSortOnly) {
    if (!params?.dest?.mode) dest.mode = "inPlace";
  }
  const optionsObj = params?.options && typeof params.options === "object" ? { ...params.options } : {};
  const outputTableName = params?.outputTableName;

  return {
    source: params?.source || {},
    select,
    rename: renameRes.provided ? rename : Object.keys(rename).length ? rename : undefined,
    filter,
    sort,
    dest,
    outputTableName,
    options: optionsObj,
  };
}

