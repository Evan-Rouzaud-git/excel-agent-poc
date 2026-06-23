import { WorkbookContextSnapshot } from "../context/types";
import {
  AgentMacroName,
  ArtifactRecord,
  JoinMatchStrategy,
  JoinTablesParamsNormalized,
  MacroContext,
  MacroResult,
  ConfirmationRequest,
  ValidateDataParams,
} from "./types";
import {
  AddressBounds,
  BlockArtifactRef,
  findBlock,
  isBlankCell,
  isRangeBlank,
  makeRangeAddress,
  parseA1Address,
  resolveBlockRefInput,
  rowColToA1,
} from "./utils";
import { matchHeaderToken } from "./tableViewUtils";
import { parseDateCell } from "./dateUtils";
import { VALIDATE_DATA_QUESTIONS, VALIDATE_DATA_CONFIRM_CHOICES } from "./validateDataFlow";

type MacroFn = (params: any, ctx: MacroContext) => Promise<MacroResult>;

const DEFAULT_BLANK_AREA = { rows: 6, cols: 4 };

function isWhitespaceChar(ch: string) {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r" || ch === "\f";
}

function collapseWhitespace(str: string) {
  let out = "";
  let inSpace = false;
  for (const ch of str) {
    if (isWhitespaceChar(ch)) {
      if (!inSpace) {
        out += " ";
        inSpace = true;
      }
    } else {
      out += ch;
      inSpace = false;
    }
  }
  return out.trim();
}

function removeDiacritics(str: string) {
  const decomposed = str.normalize("NFD");
  let out = "";
  for (const ch of decomposed) {
    const code = ch.charCodeAt(0);
    if (code < 0x0300 || code > 0x036f) out += ch;
  }
  return out;
}

function normalizeHeader(text?: string) {
  const base = collapseWhitespace((text || "").toString().trim()).toLowerCase();
  return removeDiacritics(base);
}

async function loadAndSync(obj: any, props: string[], ctx: MacroContext) {
  if (!obj || typeof obj.load !== "function") return;
  obj.load(props);
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
}

function buildCandidateNormMap(values: string[]) {
  const map = new Map<string, number>();
  values.forEach((value, idx) => {
    const norm = normalizeHeader(value);
    if (norm && !map.has(norm)) map.set(norm, idx);
  });
  return map;
}

function buildHeaderAliasLookup(aliases?: Record<string, string[]>) {
  const map = new Map<string, string[]>();
  if (!aliases) return map;
  Object.entries(aliases).forEach(([key, values]) => {
    const normKey = normalizeHeader(key);
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

function resolveHeaderWithAliases(
  token: string,
  candidates: string[],
  normMap: Map<string, number>,
  aliasLookup: Map<string, string[]>
): { header: string; index: number } | null {
  if (!token) return null;
  const normalizedToken = token.trim();
  if (!normalizedToken) return null;
  const direct = matchHeaderToken(candidates, normalizedToken);
  if (direct) return { header: direct.header, index: direct.index };
  const norm = normalizeHeader(normalizedToken);
  if (!norm) return null;
  const aliasValues = aliasLookup.get(norm);
  if (!aliasValues?.length) return null;
  for (const alias of aliasValues) {
    const aliasNorm = normalizeHeader(alias);
    if (aliasNorm) {
      const idx = normMap.get(aliasNorm);
      if (typeof idx === "number") {
        const candidate = candidates[idx];
        if (typeof candidate === "string") return { header: candidate, index: idx };
      }
    }
    const directIdx = candidates.findIndex((candidate) => candidate === alias);
    if (directIdx >= 0) {
      const candidate = candidates[directIdx];
      if (typeof candidate === "string") return { header: candidate, index: directIdx };
    }
  }
  return null;
}

function resolveHeaderIndex(headers: string[], target?: string) {
  if (!target) return { idx: null, reason: "missing" };
  const normTarget = normalizeHeader(target);
  const matches = headers
    .map((h, i) => ({ i, norm: normalizeHeader(h) }))
    .filter((h) => h.norm === normTarget);
  if (matches.length === 0) return { idx: null, reason: "not_found" };
  if (matches.length > 1) {
    const first = matches[0]!;
    return { idx: first.i, reason: "ambiguous" };
  }
  const first = matches[0]!;
  return { idx: first.i, reason: null };
}

function findHeaderIndex(headers: any[], label: string): number {
  const target = normalizeHeader(label);
  if (!target || !Array.isArray(headers)) return -1;
  for (let idx = 0; idx < headers.length; idx += 1) {
    const cell = headers[idx];
    const text = cell === null || typeof cell === "undefined" ? "" : `${cell}`;
    if (normalizeHeader(text) === target) return idx;
  }
  return -1;
}

function normalizeFixValue(raw?: any): "apply" | "ignore" {
  const str = raw === null || typeof raw === "undefined" ? "" : `${raw}`;
  const normalized = str.trim().toLowerCase();
  return normalized === "ignore" ? "ignore" : "apply";
}

function resolveDefaultBlockRef(ctx: MacroContext): string | null {
  const activeSheetName = ctx.context.active.sheetName || ctx.context.sheets[0]?.name || null;
  const activeSheet = ctx.context.sheets.find((s) => s.name === activeSheetName) || ctx.context.sheets[0];
  return (
    ctx.context.active.selectionInBlockId ||
    ctx.context.active.nearestBlockId ||
    activeSheet?.blocks[0]?.id ||
    (activeSheet?.usedRange ? `${activeSheet.name}!${activeSheet.usedRange}` : null)
  );
}

function resolveBlockRefOrArtifact(
  ref: BlockArtifactRef | undefined,
  ctx: MacroContext,
  opts: { allowDefault?: boolean; defaultBlockRef?: string } = {}
): { ok: true; blockRef: string; artifact?: any } | { ok: false; reason: string } {
  const resolution = resolveBlockRefInput(ref, ctx.artifacts || []);
  if (resolution.ok) {
    if (ref && !ref.blockRef) ref.blockRef = resolution.blockRef;
    return { ok: true, blockRef: resolution.blockRef, artifact: resolution.artifact };
  }
  if (ref?.artifactRef) return { ok: false, reason: resolution.reason };
  const fallback = opts.defaultBlockRef || (opts.allowDefault ? resolveDefaultBlockRef(ctx) : null);
  if (fallback) {
    if (ref) ref.blockRef = fallback;
    return { ok: true, blockRef: fallback, artifact: undefined };
  }
  if (ref?.blockRef) return { ok: true, blockRef: ref.blockRef, artifact: undefined };
  return { ok: false, reason: resolution.reason };
}

export function convertStructuredToA1(formula: string, headers: string[], rowIndex: number, dataStartCol: number) {
  const missing = new Set<string>();
  let result = "";
  let i = 0;
  const src = formula || "";
  while (i < src.length) {
    const ch = src[i];
    if (ch === "[") {
      const closeIdx = src.indexOf("]", i + 1);
      if (closeIdx > i) {
        const token = src.slice(i + 1, closeIdx).trim();
        if (token) {
          const name = token.startsWith("@") ? token.slice(1).trim() : token;
          const { idx } = resolveHeaderIndex(headers, name);
          if (idx === null || idx === undefined || idx < 0) {
            missing.add(name);
            result += "0";
          } else {
            result += rowColToA1(rowIndex, dataStartCol + idx);
          }
          i = closeIdx + 1;
          continue;
        }
      }
    }
    result += ch;
    i += 1;
  }
  return { formula: result, missing: Array.from(missing) };
}

const SIMPLE_RANGE_DIVISION = /^=\s*\(?\$?([A-Z]+[0-9]+)\)?\s*\/\s*\(?\$?([A-Z]+[0-9]+)\)?$/i;
function wrapDivisionForRange(formula: string) {
  if (!formula) return formula;
  const trimmed = formula.trim();
  const match = trimmed.match(SIMPLE_RANGE_DIVISION);
  if (!match) return formula;
  const left = match[1];
  const right = match[2];
  return `=IFERROR((--${left})/(--${right}),"")`;
}

const HIGHLIGHT_RED_FILL = "#f8d7da";
const HIGHLIGHT_ORANGE_FILL = "#fde9d9";
const HIGHLIGHT_VIOLET_FILL = "#e4d2f5";
type IssueType = "missing" | "duplicate" | "bad_type";
const highlightColorByType: Record<IssueType, string> = {
  missing: HIGHLIGHT_RED_FILL,
  duplicate: HIGHLIGHT_ORANGE_FILL,
  bad_type: HIGHLIGHT_VIOLET_FILL,
};
const highlightPriorityByType: Record<IssueType, number> = {
  missing: 3,
  bad_type: 2,
  duplicate: 1,
};
const issueTypeOrder: Record<IssueType, number> = {
  missing: 0,
  bad_type: 1,
  duplicate: 2,
};
const ISSUE_TYPES: IssueType[] = ["missing", "duplicate", "bad_type"];
const ISSUE_TYPE_MISSING: IssueType = "missing";
const ISSUE_TYPE_DUPLICATE: IssueType = "duplicate";
const ISSUE_TYPE_BAD_TYPE: IssueType = "bad_type";
type IssueCounts = { missing: number; duplicate: number; bad_type: number };
type IssueTableSnapshot = {
  sheet: string;
  headers: any[];
  rows: any[][];
  counts: IssueCounts;
  rangeAddress?: string;
  localRange?: string | null;
  range?: AddressBounds | null;
  reference?: IssuesSheetReference | null;
};
type IssuesSheetReference = {
  sheet: string;
  tableName?: string;
  rangeA1?: string;
  createdAt?: string;
};
const ISSUES_REFERENCE_NAMED_ITEM = "__validate_data_last_issues_ref";

function normalizeIssuesSheetName(raw?: any): string | null {
  if (raw === null || typeof raw === "undefined") return null;
  let value = `${raw}`.trim();
  if (!value) return null;
  if (value.startsWith("=")) value = value.slice(1).trim();
  if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
    value = value.slice(1, -1).trim();
  }
  return value || null;
}

async function persistIssuesSheetReference(ctx: MacroContext, reference: IssuesSheetReference) {
  if (!reference?.sheet) return;
  const workbook: any = (ctx.excelCtx as any)?.workbook;
  const names = workbook?.names;
  if (!names || typeof names.getItemOrNullObject !== "function" || typeof names.add !== "function") {
    ctx.log({
      level: "warn",
      message: "validate_data issues_reference_write_failed",
      data: { error: "names_missing", sheet: reference.sheet },
    });
    return;
  }
  const storedReference = {
    sheet: reference.sheet,
    tableName: typeof reference.tableName === "string" ? reference.tableName : null,
    rangeA1: reference.rangeA1 || null,
    createdAt: reference.createdAt || new Date().toISOString(),
  };
  const payload = JSON.stringify(storedReference);
  try {
    const existing = names.getItemOrNullObject(ISSUES_REFERENCE_NAMED_ITEM);
    await loadAndSync(existing, ["value", "isNullObject"], ctx);
    if (!existing || existing.isNullObject) {
      names.add(ISSUES_REFERENCE_NAMED_ITEM, payload);
    } else if (typeof existing.setValue === "function") {
      existing.setValue(payload);
    } else {
      existing.value = payload;
    }
    if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
  } catch (err: any) {
    ctx.log({
      level: "warn",
      message: "validate_data issues_reference_write_failed",
      data: {
        error: err?.message || err,
        sheet: reference.sheet,
        payload: storedReference,
      },
    });
    return;
  }
  ctx.log({
    level: "info",
    message: "validate_data issues_reference_written",
    data: {
      sheet: reference.sheet,
      tableName: storedReference.tableName,
      rangeA1: storedReference.rangeA1,
    },
  });
}

function parseIssuesReferenceValue(raw?: string): IssuesSheetReference | null {
  if (!raw) return null;
  const normalized = normalizeIssuesSheetName(raw);
  if (!normalized) return null;
  try {
    const parsed = JSON.parse(normalized);
    if (parsed && typeof parsed === "object" && typeof parsed.sheet === "string") {
      const tableName = typeof parsed.tableName === "string" && parsed.tableName ? parsed.tableName : undefined;
      return {
        sheet: parsed.sheet,
        tableName,
        rangeA1: typeof parsed.rangeA1 === "string" ? parsed.rangeA1 : undefined,
        createdAt: typeof parsed.createdAt === "string" ? parsed.createdAt : undefined,
      };
    }
  } catch {
    // ignore parse errors, fallback to sheet string
  }
  if (normalized) {
    return { sheet: normalized, tableName: undefined };
  }
  return null;
}

async function getIssuesSheetReferenceFromNamedItem(ctx: MacroContext): Promise<IssuesSheetReference | null> {
  const workbook: any = (ctx.excelCtx as any)?.workbook;
  const names = workbook?.names;
  if (!names || typeof names.getItemOrNullObject !== "function") return null;
  try {
    const named = names.getItemOrNullObject(ISSUES_REFERENCE_NAMED_ITEM);
    if (!named) {
      ctx.log({ level: "info", message: "validate_data issues_reference_nameditem_missing" });
      return null;
    }
    await loadAndSync(named, ["value", "formula", "name", "type", "isNullObject"], ctx);
    if (named.isNullObject) {
      ctx.log({ level: "info", message: "validate_data issues_reference_nameditem_missing" });
      return null;
    }
    const raw = named.value ?? named.formula ?? "";
    const parsed = parseIssuesReferenceValue(raw);
    return parsed;
  } catch (err: any) {
    ctx.log({
      level: "warn",
      message: "validate_data issues_reference_load_failed",
      data: { error: err?.message || err },
    });
    return null;
  }
}

function extractIssuesScore(name: string): { priority: number; label: string } {
  const digitsMatch = name.match(/(?:_|\\s|\\()+(\\d+)\\)?$/);
  const priority = digitsMatch ? Number(digitsMatch[1]) : 0;
  return { priority: priority || 0, label: name.toLowerCase() };
}

async function scanForIssuesSheetReference(ctx: MacroContext): Promise<IssuesSheetReference | null> {
  const workbook: any = (ctx.excelCtx as any)?.workbook;
  const sheets = workbook?.worksheets;
  if (!sheets) return null;
  if (typeof sheets.load === "function") sheets.load("items/name");
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
  const items: any[] =
    Array.isArray(sheets.items) && sheets.items.length ? sheets.items : typeof sheets.rawItems === "function" ? sheets.rawItems() : [];
  const candidateNames = items
    .map((entry) => entry?.name || entry?._name)
    .filter((name): name is string => typeof name === "string")
    .filter((name) => name.toLowerCase().startsWith("issues"));
  if (!candidateNames.length) return null;
  candidateNames.sort((a, b) => {
    const scoreA = extractIssuesScore(a);
    const scoreB = extractIssuesScore(b);
    if (scoreA.priority !== scoreB.priority) return scoreB.priority - scoreA.priority;
    return b.localeCompare(a);
  });
  const chosen = candidateNames[0];
  if (!chosen) return null;
  let worksheet: any;
  try {
    worksheet = getWorksheet(ctx.excelCtx, chosen);
  } catch {
    return null;
  }
  const tableObj = await resolveIssuesTableOnSheet(ctx, worksheet);
  if (!tableObj) return null;
  const rangeAddress = await resolveTableRangeAddress(tableObj, ctx);
  return {
    sheet: chosen,
    tableName: tableObj.name,
    rangeA1: rangeAddress ? extractLocalRange(rangeAddress) || undefined : undefined,
  };
}

async function resolveIssuesTableOnSheet(ctx: MacroContext, worksheet: any, tableNameHint?: string): Promise<any | null> {
  if (!worksheet) return null;
  const tables = worksheet.tables;
  if (!tables) return null;
  if (typeof tables.load === "function") tables.load("items/name");
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
  if (tableNameHint) {
    try {
      const candidate = tables.getItemOrNullObject(tableNameHint);
      if (candidate) {
        await loadAndSync(candidate, ["isNullObject"], ctx);
      }
      if (candidate && !candidate.isNullObject) {
        return candidate;
      }
    } catch {}
  }
  const list: Excel.Table[] =
    Array.isArray(tables.items) && tables.items.length ? (tables.items as Excel.Table[]) : [];
  return list.find((tbl) => typeof tbl?.name === "string" && tbl.name.toLowerCase().startsWith("issues")) || null;
}

async function resolveTableRangeAddress(table: any, ctx: MacroContext): Promise<string | undefined> {
  if (!table) return undefined;
  const range = table.getRange?.();
  if (!range) return undefined;
  if (typeof range.load === "function") range.load("address");
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
  return typeof range.address === "string" ? range.address : undefined;
}

async function resolveIssuesSheetReference(ctx: MacroContext): Promise<IssuesSheetReference | null> {
  const nameReference = await getIssuesSheetReferenceFromNamedItem(ctx);
  if (nameReference?.sheet && nameReference.tableName) {
    ctx.log({
      level: "info",
      message: "validate_data issues_reference_resolved",
      data: {
        sheet: nameReference.sheet,
        tableName: nameReference.tableName,
        rangeA1: nameReference.rangeA1,
      },
    });
    return nameReference;
  }
  const fallback = await scanForIssuesSheetReference(ctx);
  if (!fallback) return null;
  ctx.log({
    level: "info",
    message: "validate_data issues_reference_fallback_used",
    data: { sheet: fallback.sheet, tableName: fallback.tableName },
  });
  await persistIssuesSheetReference(ctx, fallback);
  ctx.log({
    level: "info",
    message: "validate_data issues_reference_resolved",
    data: {
      sheet: fallback.sheet,
      tableName: fallback.tableName,
      rangeA1: fallback.rangeA1,
    },
  });
  return fallback;
}

function propagateValueToContainingRanges(targetWs: any, row: number, col: number, value: any) {
  if (!targetWs || !targetWs.ranges) return;
  Object.values(targetWs.ranges).forEach((range: any) => {
    const parsed = parseA1Address(range.address);
    if (!parsed) return;
    if (row < parsed.startRow || row > parsed.endRow || col < parsed.startCol || col > parsed.endCol) return;
    const relativeRow = row - parsed.startRow;
    const relativeCol = col - parsed.startCol;
    const data = Array.isArray(range.values) ? range.values.map((r: any[]) => [...r]) : [];
    while (data.length <= relativeRow) data.push([]);
    while (data[relativeRow].length <= relativeCol) data[relativeRow].push("");
    data[relativeRow][relativeCol] = value;
    range.values = data;
  });
}

export function convertStructuredToThisRow(formula: string, headers: string[] = []) {
  const normalizedMap = new Map<string, string>();
  headers.forEach((header) => {
    const norm = normalizeHeader(header);
    if (norm && !normalizedMap.has(norm)) normalizedMap.set(norm, header);
  });
  const missing = new Set<string>();
  if (!formula) return { formula: "", missing: [] };
  const replaced = formula.replace(/\[([^\[\]]+)\]/g, (match, token, offset, full) => {
    if (!token) return match;
    const trimmed = token.trim();
    if (!trimmed) return match;
    if (trimmed.startsWith("#") || trimmed.startsWith("@")) return match;
    if (trimmed.includes("[") || trimmed.includes("]") || trimmed.includes(",") || trimmed.includes("!")) return match;
    const prevChar = offset > 0 ? full[offset - 1] : "";
    if (/[A-Za-z0-9_]/.test(prevChar) || prevChar === "[" || prevChar === "]") return match;
    const normalized = normalizeHeader(trimmed);
    if (!normalized) return match;
    const headerLabel = normalizedMap.get(normalized);
    if (!headerLabel) {
      missing.add(trimmed);
      return match;
    }
    return `[@${headerLabel}]`;
  });
  return { formula: replaced, missing: Array.from(missing) };
}

function hasTableWithName(context: WorkbookContextSnapshot, tableName?: string, sheetName?: string) {
  if (!tableName) return false;
  const normalizedTable = tableName.trim().toLowerCase();
  if (!normalizedTable) return false;
  return (context?.sheets || []).some((sheet) => {
    if (sheetName && sheet.name !== sheetName) return false;
    if (!Array.isArray(sheet.tables)) return false;
    return sheet.tables.some((tbl) => (tbl?.name || "").toLowerCase() === normalizedTable);
  });
}

function getWorksheet(excelCtx: Excel.RequestContext, sheetName: string) {
  const wb: any = (excelCtx as any).workbook;
  const sheets = wb?.worksheets;
  if (sheets?.getItem) return sheets.getItem(sheetName);
  if (sheets?.getItemOrNullObject) {
    const maybe = sheets.getItemOrNullObject(sheetName);
    if (!maybe?.isNullObject) return maybe;
  }
  if (typeof sheets?.rawItems === "function") {
    const safeName = (s: any) => (s && typeof s._name !== "undefined" ? s._name : (() => { try { return s.name; } catch { return undefined; } })());
    const found = sheets.rawItems().find((s: any) => {
      const nm = safeName(s);
      return nm === sheetName;
    });
    if (found) return found;
  }
  if (Array.isArray(sheets?.items)) {
    const safeName = (s: any) => (s && typeof s._name !== "undefined" ? s._name : (() => { try { return s.name; } catch { return undefined; } })());
    const found = sheets.items.find((s: any) => {
      const nm = safeName(s);
      return nm === sheetName;
    });
    if (found) return found;
  }
  throw new Error(`Worksheet ${sheetName} not found`);
}

async function readRangeValues(range: any, ctx: Excel.RequestContext): Promise<any[][]> {
  if (typeof range?.load === "function") range.load("values");
  if (typeof (ctx as any).sync === "function") {
    await (ctx as any).sync();
  }
  return (range as any)?.values ?? [];
}

async function readRangeValuesWithFormats(
  range: any,
  ctx: Excel.RequestContext
): Promise<{ values: any[][]; numberFormat: any[][]; texts: any[][] }> {
  if (typeof range?.load === "function") {
    try {
      range.load(["values", "numberFormat", "text"]);
    } catch {
      // older mock objects may not support array signature
      try {
        range.load("values,numberFormat,text");
      } catch {
        range.load("values");
      }
    }
  }
  if (typeof (ctx as any).sync === "function") {
    await (ctx as any).sync();
  }
  return {
    values: (range as any)?.values ?? [],
    numberFormat: (range as any)?.numberFormat ?? [],
    texts: (range as any)?.text ?? [],
  };
}

async function ensureUniqueSheet(excelCtx: Excel.RequestContext, hint: string): Promise<{ sheet: any; name: string }> {
  const wb: any = (excelCtx as any).workbook;
  const sheets = wb?.worksheets;
  let suffix = 0;
  const base = hint || "Resultats";
  while (suffix < 100) {
    const candidate = suffix === 0 ? base : `${base} (${suffix + 1})`;
    if (sheets?.getItemOrNullObject) {
      const maybe = sheets.getItemOrNullObject(candidate);
      if (typeof maybe.load === "function") maybe.load("isNullObject");
      if (typeof (excelCtx as any).sync === "function") await (excelCtx as any).sync();
      if ((maybe as any).isNullObject) {
        const sheet = sheets.add ? sheets.add(candidate) : { name: candidate, ranges: {} };
        return { sheet, name: candidate };
      }
    } else if (sheets?.getItem) {
      try {
        const existing = sheets.getItem(candidate);
        if (existing && existing.load) {
          existing.load("name");
          if (typeof (excelCtx as any).sync === "function") await (excelCtx as any).sync();
        }
      } catch {
        const sheet = sheets.add ? sheets.add(candidate) : { name: candidate, ranges: {} };
        return { sheet, name: candidate };
      }
    } else {
      const sheet = sheets?.add ? sheets.add(candidate) : { name: candidate, ranges: {} };
      return { sheet, name: candidate };
    }
    suffix += 1;
  }
  throw new Error("Unable to create unique sheet name");
}

function boundsFromBlock(blockRef: string, snapshot: WorkbookContextSnapshot): { sheetName: string; bounds: AddressBounds | null } {
  const parsed = parseA1Address(blockRef);
  const sheetName = parsed?.sheet || blockRef.split("!")[0] || "";
  const found = findBlock(blockRef, snapshot);
  const bounds = parsed || (found.block ? parseA1Address(`${sheetName}!${found.block.address}`) : null);
  return { sheetName, bounds };
}

function confirmationForOverwrite(id: string): ConfirmationRequest {
  return {
    id,
    question: "Zone non vide, que faire ?",
    choices: [
      { id: "newSheet", label: "Nouvelle feuille" },
      { id: "below", label: "Placer en dessous" },
      { id: "right", label: "Placer a droite" },
      { id: "abort", label: "Annuler" },
    ],
  };
}

async function resolvePlacement(params: any, ctx: MacroContext, confirmationId: string): Promise<MacroResult & { address?: string; worksheet?: any }> {
  const mode: "below" | "right" | "newSheet" | "inTableNewColumn" = params.mode || "right";
  const area = params.minBlankArea || DEFAULT_BLANK_AREA;
  const anchor = params.anchor || {};
  let chosenMode = mode;
  const decision = ctx.decisions[confirmationId];
  if (decision) {
    if (decision === "abort") {
      ctx.log({ level: "warn", message: "Placement annule par l'utilisateur", stepId: ctx.step.id, macro: "place_output" });
      return { status: "skipped" };
    }
    chosenMode = decision as any;
  }

  if (chosenMode === "newSheet") {
    const { sheet, name } = await ensureUniqueSheet(ctx.excelCtx, params.newSheetNameHint || "Resultats");
    const address = `${name}!A1:${rowColToA1(area.rows - 1, area.cols - 1)}`;
    return { address, worksheet: sheet, artifacts: [{ type: "range", sheet: name, anchor: "A1", fromStep: ctx.step.id }] };
  }

  const anchorBlock = anchor.blockRef ? boundsFromBlock(anchor.blockRef, ctx.context) : null;
  const targetSheet = anchorBlock?.sheetName || anchor.sheet || ctx.context.active.sheetName || (ctx.context.sheets[0]?.name ?? "Sheet1");
  const bounds = anchorBlock?.bounds || (anchor.cell ? parseA1Address(`${targetSheet}!${anchor.cell}`) : null);
  const baseRow = bounds?.startRow ?? 0;
  const baseCol = bounds?.startCol ?? 0;
  const endRow = bounds?.endRow ?? baseRow;
  const endCol = bounds?.endCol ?? baseCol;

  const startRow = chosenMode === "below" ? endRow + 1 : baseRow;
  const startCol = chosenMode === "right" ? endCol + 1 : baseCol;

  const address = makeRangeAddress(targetSheet, startRow, startCol, area.rows, area.cols);
  const ws = getWorksheet(ctx.excelCtx, targetSheet);
  const localAddr: string = address.includes("!") ? address.split("!")[1] || address : address;
  const range = ws.getRange ? ws.getRange(localAddr) : ws.getRangeByIndexes?.(startRow, startCol, area.rows, area.cols);

  if (params.avoidOverwrite) {
    const values = await readRangeValues(range, ctx.excelCtx);
    const blank = isRangeBlank(values);
    if (!blank && !decision) {
      return { requiresConfirmation: confirmationForOverwrite(confirmationId) };
    }
  }

  return { address, worksheet: ws, artifacts: [{ type: "range", sheet: targetSheet, anchor: localAddr, fromStep: ctx.step.id }] };
}

const macroPlaceOutput: MacroFn = async (params, ctx) => {
  const confirmationId = `${ctx.step.id}:placement`;
  const resolved = await resolvePlacement(params, ctx, confirmationId);
  if (resolved.requiresConfirmation) return { requiresConfirmation: resolved.requiresConfirmation };
  if (!resolved.address) return resolved;
  ctx.log({ level: "info", message: `Zone reservee ${resolved.address}`, macro: "place_output", stepId: ctx.step.id });
  return { artifacts: resolved.artifacts, status: resolved.status };
};

function buildOverwriteConfirmation(id: string, headerName: string): ConfirmationRequest {
  return {
    id,
    question: `Ecraser la colonne ${headerName} ?`,
    choices: [
      { id: "overwrite", label: "Ecraser" },
      { id: "abort", label: "Annuler" },
    ],
  };
}

const macroWriteFormula: MacroFn = async (params, ctx) => {
  const confirmationId = `${ctx.step.id}:overwrite`;
  const target = params.target || {};
  const headerName: string = target.headerName || "Result";
  const ifOverwrite: "ask" | "abort" | "overwrite" = params.ifOverwrite || "ask";
  const fillDown: boolean = params.fillDown !== false;
  const numberFormat: string | undefined = params.numberFormat;

  const hasExplicitRange = !!(target.sheet && target.rangeA1);
  let resolvedBlockRef: string | null = target.blockRef || null;
  let resolvedArtifact: any = null;
  if (!hasExplicitRange) {
    const resolved = resolveBlockRefOrArtifact(target as BlockArtifactRef, ctx, { allowDefault: true });
    if (!resolved.ok) {
      ctx.log({
        level: "error",
        message: `write_formula: cible introuvable (${resolved.reason})`,
        macro: "write_formula",
        stepId: ctx.step.id,
      });
      return { status: "error" };
    }
    resolvedBlockRef = resolved.blockRef;
    resolvedArtifact = resolved.artifact;
  }
  if (!target.blockRef && resolvedBlockRef) target.blockRef = resolvedBlockRef;
  if (resolvedArtifact) {
    target.tableName = target.tableName || (resolvedArtifact as any).tableName;
    target.sheetName = target.sheetName || (resolvedArtifact.sheetName || resolvedArtifact.sheet);
  }
  let sheetName: string | undefined;
  let dataBounds: AddressBounds | null = null;
  let headerRow: number | null = null;
  let blockHeaders: string[] = [];
  let resolvedBlock: any;
  if (resolvedArtifact?.headers && Array.isArray(resolvedArtifact.headers)) {
    blockHeaders = (resolvedArtifact.headers as string[]).slice();
  }

  if (resolvedBlockRef) {
    const { sheet, block } = findBlock(resolvedBlockRef, ctx.context);
    resolvedBlock = block;
    sheetName = sheet?.name;
    if (block?.headers?.length) {
      blockHeaders = block.headers;
    }
    if (block && sheetName) {
      if (block.kind === "range") {
        const parsed = parseA1Address(`${sheetName}!${block.address}`) || parseA1Address(resolvedBlockRef);
        if (parsed) {
          headerRow = parsed.startRow;
          dataBounds = { ...parsed, startRow: Math.min(parsed.endRow, parsed.startRow + 1) };
        }
      } else {
        const tableInfo = sheet?.tables.find(
          (t) => t.address === `${sheetName}!${block.address}` || t.dataBodyAddress === `${sheetName}!${block.address}` || t.name === (block.source as any)?.tableName
        );
        const bodyAddress = tableInfo?.dataBodyAddress || `${sheetName}!${block.address}`;
        dataBounds = parseA1Address(bodyAddress);
        const headerAddress =
          tableInfo?.headerAddress ||
          `${sheetName}!${rowColToA1(Math.max(0, (dataBounds?.startRow ?? 0) - 1), dataBounds?.startCol ?? 0)}:${rowColToA1(
            Math.max(0, (dataBounds?.startRow ?? 0) - 1),
            dataBounds?.endCol ?? 0
          )}`;
        headerRow = parseA1Address(headerAddress)?.startRow ?? null;
      }
    } else if (resolvedBlockRef.includes("!")) {
      const parsed = parseA1Address(resolvedBlockRef);
      sheetName = sheetName || parsed?.sheet;
      blockHeaders = blockHeaders || [];
      if (parsed) {
        headerRow = parsed.startRow;
        dataBounds = { ...parsed, startRow: Math.min(parsed.endRow, parsed.startRow + 1) };
      }
    }
  } else if (hasExplicitRange) {
    sheetName = target.sheet;
    dataBounds = parseA1Address(`${sheetName}!${target.rangeA1}`);
    headerRow = dataBounds?.startRow ?? null;
  }

  if (!sheetName || !dataBounds) {
    ctx.log({ level: "error", message: "Cible introuvable pour write_formula", macro: "write_formula", stepId: ctx.step.id });
    return { status: "error" };
  }

  const ws = getWorksheet(ctx.excelCtx, sheetName);

  // fallback: try reading header row if block headers missing (happens for artifactRef tables not in snapshot)
  if (blockHeaders.length === 0 && headerRow !== null && headerRow !== undefined) {
    try {
      const headerAddrLocal = `${rowColToA1(headerRow, dataBounds.startCol)}:${rowColToA1(headerRow, dataBounds.endCol)}`;
      const headerValues = await readRangeValues(ws.getRange ? ws.getRange(headerAddrLocal) : ws.getRangeByIndexes?.(headerRow, dataBounds.startCol, 1, dataBounds.endCol - dataBounds.startCol + 1), ctx.excelCtx);
      if (Array.isArray(headerValues?.[0])) blockHeaders = headerValues[0] as string[];
    } catch {
      // ignore
    }
  }

  let targetColIndex: number;
  if (target.writeMode === "existingColumn") {
    targetColIndex = blockHeaders.findIndex((h) => h.toLowerCase().trim() === headerName.toLowerCase().trim());
    if (targetColIndex < 0) {
      ctx.log({ level: "warn", message: "Colonne cible introuvable", macro: "write_formula", stepId: ctx.step.id });
      return { status: "skipped" };
    }
  } else {
    targetColIndex = blockHeaders.length || (dataBounds.endCol - dataBounds.startCol + 1);
  }

  const targetColAbsolute = dataBounds.startCol + targetColIndex;
  const originalStartRow = dataBounds.startRow;
  const headerRowIndex = headerRow ?? Math.max(0, originalStartRow - 1);
  const blockSourceType = resolvedBlock?.source?.type;
  const isRangeTarget = resolvedBlock?.kind === "range" || blockSourceType === "range";
  const rangeStartRow = isRangeTarget ? Math.max(headerRowIndex + 1, originalStartRow) : originalStartRow;
  const rangeRowCount = Math.max(1, dataBounds.endRow - rangeStartRow + 1);
  const targetRangeAddress = makeRangeAddress(sheetName, rangeStartRow, targetColAbsolute, rangeRowCount, 1);
  const localAddr = targetRangeAddress.split("!")[1];
  const targetRange = ws.getRange ? ws.getRange(localAddr) : ws.getRangeByIndexes?.(rangeStartRow, targetColAbsolute, rangeRowCount, 1);

  const headerDuplicate = blockHeaders.some((h) => h.toLowerCase().trim() === headerName.toLowerCase().trim());

  if (ifOverwrite === "abort") {
    ctx.log({ level: "warn", message: "Ecriture abandonnee (ifOverwrite=abort)", macro: "write_formula", stepId: ctx.step.id });
    return { status: "skipped" };
  }

  if (ifOverwrite === "ask") {
    const values = await readRangeValues(targetRange, ctx.excelCtx);
    const needConfirm = headerDuplicate || !isRangeBlank(values);
    if (needConfirm && !ctx.decisions[confirmationId]) {
      return { requiresConfirmation: buildOverwriteConfirmation(confirmationId, headerName) };
    }
    if (ctx.decisions[confirmationId] === "abort") {
      ctx.log({ level: "warn", message: "Ecriture annulee par utilisateur", macro: "write_formula", stepId: ctx.step.id });
      return { status: "skipped" };
    }
  }

  const headerCellAddress = `${sheetName}!${rowColToA1(headerRowIndex, targetColAbsolute)}`;
  const headerRange = ws.getRange ? ws.getRange(headerCellAddress.split("!")[1]) : ws.getRangeByIndexes?.(headerRowIndex, targetColAbsolute, 1, 1);

  const candidateTableName =
    target.tableName || (resolvedArtifact as any)?.tableName || (resolvedBlock?.source as any)?.tableName;
  const candidateSheetName = sheetName || target.sheetName || (resolvedArtifact as any)?.sheetName;
  const isTableFromName = hasTableWithName(ctx.context, candidateTableName, candidateSheetName);
  const isTableTarget =
    !isRangeTarget && (resolvedBlock?.kind === "table" || blockSourceType === "table" || isTableFromName);
  const missingHeaders = new Set<string>();
  const formulas: any[][] = [];
  const dataStartRow = rangeStartRow;
  const formulaRowCount = rangeRowCount;
  const wrapDivision = isRangeTarget && typeof params.formula === "string" && params.formula.includes("/");
  const tableFormula =
    isTableTarget && typeof params.formula === "string"
      ? convertStructuredToThisRow(params.formula, blockHeaders)
      : undefined;
  for (let r = 0; r < formulaRowCount; r += 1) {
    if (isRangeTarget) {
      const converted = convertStructuredToA1(params.formula, blockHeaders, dataStartRow + r, dataBounds.startCol);
      converted.missing.forEach((m) => missingHeaders.add(m));
      let formulaText = converted.formula;
      if (wrapDivision) {
        formulaText = wrapDivisionForRange(formulaText);
      }
      formulas.push([fillDown || r === 0 ? formulaText : ""]);
    } else if (isTableTarget) {
      tableFormula?.missing.forEach((m) => missingHeaders.add(m));
      const value = tableFormula?.formula || params.formula || "";
      formulas.push([fillDown || r === 0 ? value : ""]);
    } else {
      formulas.push([fillDown || r === 0 ? params.formula : ""]);
    }
  }

  const missingList = Array.from(missingHeaders);
  if (missingList.length > 0) {
    const logInfo = {
      level: "warn" as const,
      macro: "write_formula" as const,
      stepId: ctx.step.id,
    };
    const message = missingList.join(", ");
    if (isTableTarget) {
      ctx.log({ ...logInfo, message: `structured_ref_missing_at ${message}` });
    } else {
      ctx.log({ ...logInfo, message: `unknown_headers_in_formula ${message}` });
    }
    return { status: "error" };
  }

  if (headerRange) headerRange.values = [[headerName]];
  (targetRange as any).formulas = formulas;

  // In mocked context, precompute values to help downstream steps (sort, view)
  const isMockCtx = typeof (ctx.excelCtx as any)?.trackWrite === "function";
  if (isMockCtx) {
    try {
      const dataRangeAddr = makeRangeAddress(sheetName, rangeStartRow, dataBounds.startCol, rangeRowCount, dataBounds.endCol - dataBounds.startCol + 1);
      const dataRange = ws.getRange ? ws.getRange(dataRangeAddr.split("!")[1]) : ws.getRangeByIndexes?.(rangeStartRow, dataBounds.startCol, rangeRowCount, dataBounds.endCol - dataBounds.startCol + 1);
      const existingValues = await readRangeValues(dataRange, ctx.excelCtx);
      const computeVal = (rowVals: any[]) => {
        const exprRaw = (params.formula || "").toString().replace(/^\s*=/, "");
        if (!exprRaw) return "";
        const replaced = exprRaw.replace(/\[@([^\]]+)\]/g, (_m: string, name: string) => {
          const idx = blockHeaders.findIndex((h) => h && h.toString().toLowerCase().trim() === name.toString().toLowerCase().trim());
          const val = idx >= 0 ? rowVals[idx] : "";
          return Number.isFinite(val) ? `${val}` : `${Number(val)}`;
        });
        // only allow basic arithmetic
        if (!/^[0-9+*\\/().\\-\\s]+$/.test(replaced)) return replaced;
        try {
          // eslint-disable-next-line no-new-func
          const res = Function(`\"use strict\"; return (${replaced});`)();
          return res;
        } catch {
          return replaced;
        }
      };
      const computed = Array.from({ length: rangeRowCount }, (_v, i) => [computeVal(existingValues?.[i] || [])]);
      (targetRange as any).values = computed;

      // also update a combined contiguous range to keep fake grid coherent (helps later reads/sorts)
      const totalCols = Math.max(blockHeaders.length || 0, dataBounds.endCol - dataBounds.startCol + 1, targetColIndex + 1);
      const headerVals = Array.from({ length: totalCols }, (_v, i) => blockHeaders[i] || "");
      headerVals[targetColIndex] = headerName;
      const combinedRows = existingValues.map((r: any[], idx: number) => {
        const row = Array.from({ length: totalCols }, (_v, i) => (typeof r?.[i] !== "undefined" ? r[i] : ""));
        row[targetColIndex] = computed[idx]?.[0] ?? "";
        return row;
      });
      const combinedMatrix = [headerVals, ...combinedRows];
      const combinedAddr = `${rowColToA1(headerRowIndex, dataBounds.startCol)}:${rowColToA1(headerRowIndex + combinedMatrix.length - 1, dataBounds.startCol + totalCols - 1)}`;
      const combinedRange = ws.getRange ? ws.getRange(combinedAddr) : ws.getRangeByIndexes?.(headerRowIndex, dataBounds.startCol, combinedMatrix.length, totalCols);
      if (combinedRange) (combinedRange as any).values = combinedMatrix;
      // propagate artifact metadata if available
      if (resolvedArtifact) {
        resolvedArtifact.blockRef = `${sheetName}!${combinedAddr}`;
        resolvedArtifact.address = combinedAddr.includes("!") ? combinedAddr.split("!")[1] : combinedAddr;
        resolvedArtifact.rows = rangeRowCount;
        resolvedArtifact.cols = totalCols;
        resolvedArtifact.headers = headerVals;
      }
    } catch {
      // ignore mock value computation failures
    }
  }

  if (numberFormat) {
    (targetRange as any).numberFormat = Array.from({ length: rangeRowCount }, () => [numberFormat]);
  }
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();

  ctx.log({ level: "info", message: `Formule ecrite dans ${targetRangeAddress}`, macro: "write_formula", stepId: ctx.step.id });
  return {
    status: "ok",
    createdHeader: headerName,
    blockRef: target.blockRef || resolvedBlockRef || targetRangeAddress.split("!")[0],
    columnIndex: targetColIndex,
  };
};

const macroApplyFormat: MacroFn = async (params, ctx) => {
  const target = params.target;
  let sheetName: string | undefined;
  let rangeAddress: string | undefined;
  let blockRef: string | undefined;
  let headers: string[] = [];
  let targetIsTable = false;
  let tableByName: any = null;
  let resolvedArtifact: any = null;

  if (target?.artifactRef || target?.blockRef) {
    const resolved = resolveBlockRefOrArtifact(target as BlockArtifactRef, ctx, { allowDefault: false });
    if (!resolved.ok) {
      ctx.log({
        level: "error",
        message: `apply_format: cible introuvable (${resolved.reason})`,
        macro: "apply_format",
        stepId: ctx.step.id,
      });
      return { status: "error" };
    }
    blockRef = target.blockRef || resolved.blockRef;
    target.blockRef = blockRef;
    resolvedArtifact = resolved.artifact;
    if (resolvedArtifact) {
      target.tableName = target.tableName || (resolvedArtifact as any).tableName;
      target.sheetName = target.sheetName || (resolvedArtifact.sheetName || resolvedArtifact.sheet);
      rangeAddress = rangeAddress || resolvedArtifact.blockRef || (resolvedArtifact.sheetName && resolvedArtifact.address ? `${resolvedArtifact.sheetName}!${resolvedArtifact.address}` : undefined);
    }
  }
  if (target?.sheetName && target?.tableName) {
    try {
      const ws = getWorksheet(ctx.excelCtx, target.sheetName);
      tableByName = ws.tables?.getItem?.(target.tableName);
      if (tableByName) {
        sheetName = target.sheetName;
        rangeAddress = `${sheetName}!${(tableByName as any).address || target.rangeA1 || "A1"}`;
        targetIsTable = true;
      }
    } catch {
      // fallthrough
    }
  }
  if (!tableByName && target?.blockRef) {
    const { sheet, block } = findBlock(target.blockRef, ctx.context);
    const parsedTarget = parseA1Address(target.blockRef);
    sheetName = sheet?.name || parsedTarget?.sheet || target.sheetName || target.sheet || target.blockRef.split("!")[0];
    headers = block?.headers ?? [];
    const addr = block ? `${sheetName}!${block.address}` : target.blockRef;
    rangeAddress = addr.includes("!") ? addr : sheetName ? `${sheetName}!${addr}` : undefined;
    blockRef = rangeAddress || target.blockRef;
    targetIsTable = block?.kind === "table" || block?.source?.type === "table";
  } else if (target?.sheet && target?.rangeA1) {
    sheetName = target.sheet;
    rangeAddress = `${sheetName}!${target.rangeA1}`;
  }
  if (!sheetName || !rangeAddress) {
    ctx.log({ level: "warn", message: "Cible format introuvable", macro: "apply_format", stepId: ctx.step.id });
    return { status: "skipped" };
  }
  const ws = getWorksheet(ctx.excelCtx, sheetName);
  let resolvedTable = tableByName;
  if (!resolvedTable && target?.tableName && ws?.tables) {
    try {
      if (ws.tables.getItemOrNullObject) {
        const maybe = ws.tables.getItemOrNullObject(target.tableName);
        if (maybe && (maybe as any).isNullObject === false) resolvedTable = maybe;
      } else {
        resolvedTable = ws.tables.getItem(target.tableName);
      }
    } catch {
      resolvedTable = null;
    }
  }
  const bounds = parseA1Address(rangeAddress)!;
  const range = resolvedTable?.getRange
    ? resolvedTable.getRange()
    : ws.getRange
    ? ws.getRange(rangeAddress.split("!")[1])
    : ws.getRangeByIndexes?.(bounds.startRow, bounds.startCol, bounds.endRow - bounds.startRow + 1, bounds.endCol - bounds.startCol + 1);

  const options = params.options || {};
  const applyPreset = options.preset === "corporate_blue";
  const localRangeAddr = rangeAddress.split("!")[1] || rangeAddress;
  let localA1 = localRangeAddr;
  if (localA1.startsWith("'") && localA1.endsWith("'") && localA1.length >= 2) {
    localA1 = localA1.slice(1, -1);
  }
  let targetTable: any = resolvedTable || null;
  const sheetSnapshot = ctx.context.sheets.find((s) => s.name === sheetName);
  try {
    let stage = "preset";
    if (applyPreset) {
      options.header = {
        bold: true,
        background: "#0e2a80",
        fontColor: "#ffffff",
      };
      options.columnWidth = options.columnWidth ?? "auto";
      options.bandedRows = options.bandedRows ?? true;
    }

    // create or fetch table for deterministic banding/extension
    let canCreateTable = applyPreset && !!ws?.tables && !targetIsTable && !tableByName;
    let targetBoundsForTable = parseA1Address(`${sheetName}!${localA1}`);

    // merged cells detection
    if (applyPreset && canCreateTable) {
      try {
        const merged = (range as any)?.getMergedAreasOrNullObject ? (range as any).getMergedAreasOrNullObject() : (range as any)?.mergeAreas;
        if (merged && typeof merged.load === "function") merged.load("address,isNullObject");
        if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
        const hasMerges =
          !!merged &&
          ((merged as any).isNullObject === false ||
            (Array.isArray((merged as any).items) && (merged as any).items.length > 0) ||
            typeof (merged as any).address === "string");
        if (hasMerges) {
          ctx.log({ level: "warn", message: "apply_format: table skipped (merged cells)", macro: "apply_format", stepId: ctx.step.id });
          canCreateTable = false;
        }
      } catch (err: any) {
        ctx.log({ level: "warn", message: `apply_format: merged check failed ${err?.message || err}`, macro: "apply_format", stepId: ctx.step.id });
      }
    }

    // overlap detection against snapshot tables
    if (applyPreset && canCreateTable && targetBoundsForTable && sheetSnapshot?.tables?.length) {
      const overlaps = sheetSnapshot.tables.some((t) => {
        const b = parseA1Address(t.dataBodyAddress || t.address);
        return b ? rangesOverlap(targetBoundsForTable as AddressBounds, b as AddressBounds) : false;
      });
      if (overlaps) {
        ctx.log({ level: "warn", message: "apply_format: table skipped (overlap avec table existante)", macro: "apply_format", stepId: ctx.step.id });
        canCreateTable = false;
      }
    }

    // header sanitation (fill blanks, dedupe) to help table creation succeed
    if (applyPreset && canCreateTable) {
      try {
        const headerRange = ws.getRange ? ws.getRange(localA1) : range;
        const headerFix = await sanitizeHeaderRow(headerRange, targetBoundsForTable, ctx);
        if (headerFix.changed) {
          ctx.log({
            level: "info",
            message: `apply_format: headers normalises (${headerFix.headers.join(", ")})`,
            macro: "apply_format",
            stepId: ctx.step.id,
          });
          if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
        }
      } catch (err: any) {
        ctx.log({ level: "warn", message: `apply_format: header sanitize failed ${err?.message || err}`, macro: "apply_format", stepId: ctx.step.id });
      }
    }

    if (applyPreset && ws?.tables && canCreateTable) {
      const tblInfo =
        sheetSnapshot?.tables.find(
          (t) =>
            t.address === `${sheetName}!${localRangeAddr}` ||
            t.address === rangeAddress ||
            t.dataBodyAddress === `${sheetName}!${localRangeAddr}` ||
            t.dataBodyAddress === rangeAddress ||
            t.name === blockRef
        );
      const tableName = tblInfo?.name;
      try {
        if (tableName && ws.tables.getItem) {
          targetTable = ws.tables.getItem(tableName);
        } else if (ws.tables.getItemOrNullObject) {
          const maybe = ws.tables.getItemOrNullObject(tableName || "");
          if (!maybe?.isNullObject) targetTable = maybe;
        }
      } catch {
        targetTable = null;
      }
      if (!targetTable && ws.tables.add) {
        const rowsOk = bounds && bounds.endRow >= bounds.startRow;
        const colsOk = bounds && bounds.endCol >= bounds.startCol;
        if (!rowsOk || !colsOk) {
          ctx.log({ level: "warn", message: "apply_format: table skipped (range empty)", macro: "apply_format", stepId: ctx.step.id });
        } else {
          const attempts: Array<() => any> = [
            () => ws.tables.add(localA1, true),
            () => {
              const obj = ws.getRange ? ws.getRange(localA1) : range;
              return ws.tables.add(obj, true);
            },
          ];
          const errors: string[] = [];
          for (const attempt of attempts) {
            try {
              targetTable = attempt();
              if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
              ctx.log({
                level: "info",
                message: `apply_format: table created @ ${localA1}`,
                macro: "apply_format",
                stepId: ctx.step.id,
              });
              break;
            } catch (err: any) {
              errors.push(err?.message || String(err));
              targetTable = null;
            }
          }
          if (!targetTable && errors.length) {
            ctx.log({
              level: "warn",
              message: `apply_format: table add failed @ ${localA1}: ${errors.join(" | ")}`,
              macro: "apply_format",
              stepId: ctx.step.id,
            });
          }
        }
      }
    }
    if (applyPreset && targetTable) {
      try {
        targetTable.showBandedRows = false;
        targetTable.showBandedColumns = false as any;
        targetTable.showFilterButton = true as any;
      } catch {}
      try {
        targetTable.showHeaderRow = true;
      } catch {}
      try {
        targetTable.style = "TableStyleLight1";
      } catch {}
      try {
        const body = targetTable.getDataBodyRange ? targetTable.getDataBodyRange() : null;
        if (body?.format?.fill?.clear) body.format.fill.clear();
      } catch {}
    }
    if (applyPreset && range?.format?.fill?.clear) {
      range.format.fill.clear();
    }

    stage = "header";
    if (options.header) {
      const headerRange =
        targetTable?.getHeaderRowRange?.() ||
        (ws.getRange ? ws.getRange(`${rowColToA1(bounds.startRow, bounds.startCol)}:${rowColToA1(bounds.startRow, bounds.endCol)}`) : range);
      if (headerRange?.format?.fill?.clear) headerRange.format.fill.clear();
      if (options.header.bold && headerRange?.format?.font) headerRange.format.font.bold = true;
      if (headerRange?.format?.fill && options.header.background) {
        headerRange.format.fill.color = options.header.background;
      }
      if (headerRange?.format?.font && options.header.fontColor) headerRange.format.font.color = options.header.fontColor;
      if (headerRange?.format) headerRange.format.horizontalAlignment = "Center";
    }

    if (applyPreset && range?.format?.borders) {
      const borders = range.format.borders;
      const setBorder = (idx: any) => {
        try {
          const b = borders.getItem ? borders.getItem(idx) : borders[idx];
          if (!b) return;
          if (b.color !== undefined) b.color = "#000000";
          if (b.style !== undefined) b.style = "Continuous";
          if (b.weight !== undefined) b.weight = "Thin";
        } catch {}
      };
      const BI = (Excel as any)?.BorderIndex || {};
      [BI.edgeTop, BI.edgeBottom, BI.edgeLeft, BI.edgeRight, BI.insideHorizontal, BI.insideVertical].forEach(setBorder);
    }

    stage = "numberFormats";
    if (options.numberFormats) {
      if (!Array.isArray(options.numberFormats)) {
        ctx.log({ level: "warn", message: "numberFormats ignore car non-array", macro: "apply_format", stepId: ctx.step.id });
      } else if (headers?.length) {
        for (const nf of options.numberFormats) {
          const headerHints = Array.isArray((nf as any)?.headerHints) ? (nf as any).headerHints : [];
          const fmt = typeof (nf as any)?.format === "string" ? (nf as any).format : null;
          if (!headerHints.length || !fmt) {
            ctx.log({ level: "warn", message: "numberFormat ignore (headerHints/format manquant)", macro: "apply_format", stepId: ctx.step.id });
            continue;
          }
          const targetCols = headers
            .map((h, idx) => ({ h, idx }))
            .filter((h) => headerHints.some((hint: string) => h.h.toLowerCase().includes((hint || "").toLowerCase().trim())))
            .map((h) => h.idx);
          for (const colIdx of targetCols) {
            const addr = makeRangeAddress(sheetName, bounds.startRow + 1, bounds.startCol + colIdx, bounds.endRow - bounds.startRow, 1);
            const colRange = ws.getRange ? ws.getRange(addr.split("!")[1]) : range;
            (colRange as any).numberFormat = Array.from({ length: bounds.endRow - bounds.startRow }, () => [fmt]);
          }
        }
      }
    }

    stage = "width";
    if (options.columnWidth && range?.format) {
      const minColumnWidth = typeof options.minColumnWidth === "number" ? options.minColumnWidth : 12;
      const colCount =
        headers && headers.length
          ? headers.length
          : bounds && bounds.endCol >= bounds.startCol
          ? bounds.endCol - bounds.startCol + 1
          : 0;
      const targetColCount = Math.max(1, colCount || range?.colCount || 1);
      const columnRanges: any[] = [];

      const collectColumnRanges = () => {
        if (typeof (range as any).getColumn !== "function") return;
        for (let i = 0; i < targetColCount; i += 1) {
          const colRange = (range as any).getColumn(i);
          if (colRange?.format) {
            if (typeof colRange.format.load === "function") colRange.format.load("columnWidth");
            columnRanges.push(colRange);
          }
        }
      };

      if (options.columnWidth === "auto" && typeof range.format.autofitColumns === "function") {
        collectColumnRanges();
        range.format.autofitColumns();
        if (columnRanges.length && typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
        columnRanges.forEach((col) => {
          const width = (col.format as any).columnWidth;
          if (typeof width !== "number" || width < minColumnWidth) col.format.columnWidth = minColumnWidth;
        });
        if (!columnRanges.length && minColumnWidth > 0) {
          const current = (range.format as any).columnWidth as number | undefined;
          if (typeof current !== "number" || current < minColumnWidth) range.format.columnWidth = minColumnWidth;
        }
      } else if (typeof options.columnWidth === "number") {
        range.format.columnWidth = options.columnWidth;
      }
    }

    stage = "freeze";
    if (options.freezeHeaderRow && (ws as any).freezePanes?.freezeRows) {
      (ws as any).freezePanes.freezeRows(1);
    }

    // Center body content for readability
    try {
      const bodyRange =
        targetTable?.getDataBodyRange?.() ||
        (ws.getRange
          ? ws.getRange(`${rowColToA1(bounds.startRow + 1, bounds.startCol)}:${rowColToA1(bounds.endRow, bounds.endCol)}`)
          : range);
      if (bodyRange?.format) bodyRange.format.horizontalAlignment = "Center";
    } catch {}

    if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
    ctx.log({ level: "info", message: `apply_format stage=done target=${rangeAddress}`, macro: "apply_format", stepId: ctx.step.id });
    return {};
  } catch (err: any) {
    ctx.log({ level: "error", message: `apply_format stage=failed target=${rangeAddress} err=${err?.message || err}`, macro: "apply_format", stepId: ctx.step.id });
    return { status: "error" };
  }
};

const macroCreateChart: MacroFn = async (params, ctx) => {
  let stage = "init";
  try {
    stage = "sourceResolve";
    const src = params.source || {};
    if (src.artifactRef || src.blockRef) {
      const resolvedSource = resolveBlockRefOrArtifact(src as BlockArtifactRef, ctx, { allowDefault: true });
      if (!resolvedSource.ok) {
        ctx.log({
          level: "error",
          message: `create_chart: source introuvable (${resolvedSource.reason})`,
          macro: "create_chart",
          stepId: ctx.step.id,
        });
        return { status: "error" };
      }
      src.blockRef = src.blockRef || resolvedSource.blockRef;
      if (resolvedSource.artifact) {
        src.tableName = src.tableName || (resolvedSource.artifact as any).tableName;
        src.sheetName = src.sheetName || (resolvedSource.artifact.sheetName || resolvedSource.artifact.sheet);
      }
    }
    const findTable = (name: string, sheetHint?: string) => {
      for (const s of ctx.context.sheets || []) {
        if (sheetHint && s.name !== sheetHint) continue;
        const t = (s.tables || []).find((tb: any) => tb.name === name);
        if (t) return { sheetName: s.name, address: t.address || t.dataBodyAddress || t.headerAddress, headers: t.headers || [] };
      }
      return null;
    };
    let sheetName: string | undefined;
    let blockAddress: string | undefined;
    let headers: string[] = [];
    let headerRowIndex = 0;
    if (src.tableName) {
      const tbl = findTable(src.tableName, src.sheetName);
      if (tbl) {
        sheetName = tbl.sheetName;
        blockAddress = tbl.address ?? undefined;
        headers = tbl.headers;
        headerRowIndex = 0;
      }
    }
    let sourceRef: string | undefined = src.blockRef ?? undefined;
    if (!blockAddress) {
      sourceRef = sourceRef || (resolveDefaultBlockRef(ctx) ?? undefined);
      if (sourceRef) {
        const fb = findBlock(sourceRef, ctx.context);
        sheetName = fb.sheet?.name;
        blockAddress = fb.block?.address ? `${sheetName}!${fb.block.address}` : sourceRef;
        if (!sheetName) {
          const parsed = parseA1Address(sourceRef);
          sheetName = parsed?.sheet || (sourceRef.includes("!") ? sourceRef.split("!")[0] : undefined);
        }
        headers = fb.block?.headers || [];
        headerRowIndex = fb.block?.headerRowIndex ?? 0;
        if (!fb.block) {
          ctx.log({
            level: "warn",
            message: `Bloc source invalide, fallback sur ${sourceRef}`,
            macro: "create_chart",
            stepId: ctx.step.id,
          });
        }
      }
    }
    if (!sheetName || !blockAddress) {
      ctx.log({ level: "error", message: "create_chart: source introuvable", macro: "create_chart", stepId: ctx.step.id });
      return { status: "error" };
    }
    const safeSheetName = sheetName;
    stage = "bounds";
    const bounds = parseA1Address(blockAddress.includes("!") ? blockAddress : `${safeSheetName}!${blockAddress}`);
    if (!bounds) {
      ctx.log({ level: "error", message: "Adresse source invalide", macro: "create_chart", stepId: ctx.step.id });
      return { status: "error" };
    }
    const totalCols = bounds.endCol - bounds.startCol + 1;
    if (headers.length === 0) {
      try {
        const localBlock = blockAddress.includes("!") ? blockAddress.split("!")[1] || blockAddress : blockAddress;
        const blockRange = getWorksheet(ctx.excelCtx, safeSheetName)?.getRange?.(localBlock);
        const blockValues = await readRangeValues(blockRange, ctx.excelCtx);
        if (Array.isArray(blockValues?.[0])) headers = blockValues[0].slice(0, totalCols);
      } catch {
        // best effort
      }
    }

    const mapping = params.mapping || {};
    let xIdx: number | null = null;
    const warningsLocal: string[] = [];

    const resolveHeaderIdxSafe = (name?: string, fallback?: number | null) => {
      if (!name) return fallback ?? null;
      const r = resolveHeaderIndex(headers, name);
      if (r.idx === null) return null;
      if (r.reason === "ambiguous") warningsLocal.push("chart_mapping_header_ambiguous");
      return r.idx;
    };

    if (mapping.xCol?.header || mapping.xCol?.headerName) {
      xIdx = resolveHeaderIdxSafe(mapping.xCol.header || mapping.xCol.headerName, 0);
      if (xIdx === null) {
        ctx.log({
          level: "error",
          message: `chart_mapping_header_not_found:${mapping.xCol.header || mapping.xCol.headerName}`,
          macro: "create_chart",
          stepId: ctx.step.id,
        });
        return { status: "error" };
      }
    } else {
      const rawX = mapping.xCol?.colIndex ?? 0;
      xIdx = Number.isInteger(rawX) && rawX >= 0 && rawX < totalCols ? rawX : 0;
    }

    if (xIdx === null) {
      ctx.log({ level: "error", message: "Colonne X du chart introuvable", macro: "create_chart", stepId: ctx.step.id });
      return { status: "error" };
    }

    let yIdxs: number[] = [];
    const yCols = Array.isArray(mapping.yCols) ? mapping.yCols : [];
    yCols.forEach((c: any) => {
      const headerName = c?.header || c?.headerName;
      if (headerName) {
        const idx = resolveHeaderIdxSafe(headerName, null);
        if (idx === null) {
          warningsLocal.push(`chart_mapping_header_not_found:${headerName}`);
        } else {
          yIdxs.push(idx);
        }
      } else if (Number.isInteger(c?.colIndex) && c.colIndex >= 0 && c.colIndex < totalCols) {
        yIdxs.push(c.colIndex);
      }
    });

    if (yIdxs.length === 0 && (!mapping.yCols || mapping.yCols.length === 0)) {
      for (let i = 0; i < totalCols; i += 1) if (i !== xIdx) yIdxs.push(i);
    }
    if (warningsLocal.length) {
      warningsLocal.forEach((w) => ctx.log({ level: "warn", message: w, macro: "create_chart", stepId: ctx.step.id }));
    }
    yIdxs = Array.from(new Set(yIdxs)).filter((n) => n !== xIdx);
    if (yIdxs.length === 0) {
      ctx.log({
        level: "error",
        message: "Aucune colonne Y resolue pour le chart (header introuvable ?)",
        macro: "create_chart",
        stepId: ctx.step.id,
      });
      return { status: "error" };
    }

    const headerRowIndexResolved = typeof headerRowIndex === "number" ? headerRowIndex : 0;
    const headerRow = bounds.startRow + headerRowIndexResolved;
    const hasHeaderRow = typeof headerRowIndex === "number" && headerRowIndex >= 0 && headerRow <= bounds.endRow;
    const dataStartRow = hasHeaderRow ? Math.min(bounds.endRow, headerRow + 1) : bounds.startRow;
    const dataRowCount = Math.max(1, bounds.endRow - dataStartRow + 1);

    const sourceSheet = getWorksheet(ctx.excelCtx, safeSheetName);
    const rangeForCol = (colIdx: number) => {
      const colAbs = bounds.startCol + colIdx;
      const addr = makeRangeAddress(safeSheetName, dataStartRow, colAbs, dataRowCount, 1);
      const local = addr.split("!")[1];
      return sourceSheet.getRange ? sourceSheet.getRange(local) : sourceSheet.getRangeByIndexes?.(dataStartRow, colAbs, dataRowCount, 1);
    };
    const headerForCol = (colIdx: number) => {
      const colAbs = bounds.startCol + colIdx;
      const addr = makeRangeAddress(safeSheetName, headerRow, colAbs, 1, 1);
      const local = addr.split("!")[1];
      return sourceSheet.getRange ? sourceSheet.getRange(local) : sourceSheet.getRangeByIndexes?.(headerRow, colAbs, 1, 1);
    };

    const xRange = rangeForCol(xIdx);
    const yRanges = yIdxs.map((i) => rangeForCol(i));
    const firstYIdx = yIdxs[0] as number;
    const firstYRange = yRanges[0];
    const sourceRowStart = dataStartRow;
    const sourceRowCount = Math.max(1, dataRowCount);
    const orderedRelativeYIdxs = Array.from(new Set(yIdxs)).sort((a, b) => a - b);
    const orderedAbsYCols = orderedRelativeYIdxs.map((rel) => bounds.startCol + rel);
    const xAbs = bounds.startCol + xIdx;

    const createRectRange = (startCol: number, width: number) => {
      const fullAddr = makeRangeAddress(safeSheetName, sourceRowStart, startCol, sourceRowCount, width);
      const localAddr = fullAddr.split("!")[1];
      const range =
        sourceSheet.getRange?.(localAddr) ??
        sourceSheet.getRangeByIndexes?.(sourceRowStart, startCol, sourceRowCount, width);
      return { range, address: range ? fullAddr : undefined };
    };

    const createUnionRange = (columns: number[]) => {
      if (typeof sourceSheet?.getRange !== "function" || columns.length === 0) return { range: null, address: undefined };
      const parts = columns.map((colAbs) => makeRangeAddress(safeSheetName, sourceRowStart, colAbs, sourceRowCount, 1).split("!")[1]);
      const local = parts.join(",");
      const fullAddr = `${safeSheetName}!${local}`;
      const range = sourceSheet.getRange(local);
      return { range, address: range ? fullAddr : undefined };
    };

    type SourceMode = "contiguous" | "union" | "fallback";
    const sourceDataInfo: { range: any | null; address?: string; mode?: SourceMode } = { range: null, address: undefined, mode: undefined };
    const columnsForUnion = orderedAbsYCols.slice();
    if (orderedAbsYCols.length > 1) {
      const contiguousSeries = orderedAbsYCols[orderedAbsYCols.length - 1]! - orderedAbsYCols[0]! + 1 === orderedAbsYCols.length;
      if (contiguousSeries) {
        const startCol = orderedAbsYCols[0]!;
        const endCol = orderedAbsYCols[orderedAbsYCols.length - 1]!;
        const rect = createRectRange(startCol, endCol - startCol + 1);
        if (rect.range) {
          sourceDataInfo.range = rect.range;
          sourceDataInfo.address = rect.address;
          sourceDataInfo.mode = "contiguous";
        }
      }
      if (!sourceDataInfo.range) {
        const union = createUnionRange(columnsForUnion);
        if (union.range) {
          sourceDataInfo.range = union.range;
          sourceDataInfo.address = union.address;
          sourceDataInfo.mode = "union";
        }
      }
    } else {
      const singleYAbs = orderedAbsYCols[0];
      const rect = typeof singleYAbs === "number" ? createRectRange(singleYAbs, 1) : { range: null, address: undefined };
      if (rect.range) {
        sourceDataInfo.range = rect.range;
        sourceDataInfo.address = rect.address;
        sourceDataInfo.mode = "contiguous";
      } else {
        const union = createUnionRange(columnsForUnion);
        if (union.range) {
          sourceDataInfo.range = union.range;
          sourceDataInfo.address = union.address;
          sourceDataInfo.mode = "union";
        }
      }
    }

    stage = "placement";
    const destSheetHint = params.dest?.sheetName || params.dest?.sheet;
    const destAnchor = params.dest?.anchor;
    let anchorRef: string | undefined = destAnchor?.blockRef || blockAddress;
    if (destAnchor?.artifactRef) {
      const resolvedAnchor = resolveBlockRefOrArtifact(destAnchor as BlockArtifactRef, ctx, { allowDefault: false });
      if (!resolvedAnchor.ok) {
        ctx.log({
          level: "error",
          message: `create_chart: anchor introuvable (${resolvedAnchor.reason})`,
          macro: "create_chart",
          stepId: ctx.step.id,
        });
        return { status: "error" };
      }
      anchorRef = resolvedAnchor.blockRef;
      if (resolvedAnchor.artifact && !params.dest?.sheetName) {
        params.dest.sheetName = resolvedAnchor.artifact.sheetName || resolvedAnchor.artifact.sheet;
      }
    }
    if (destSheetHint && anchorRef && anchorRef.includes("!")) {
      const anchorSheet = anchorRef.split("!")[0];
      if (anchorSheet !== destSheetHint) {
        const fallbackAnchor = `${destSheetHint}!A1`;
        ctx.log({
          level: "info",
          message: "create_chart: anchor sheet force sur dest.sheetName",
          macro: "create_chart",
          stepId: ctx.step.id,
          data: { from: anchorRef, to: fallbackAnchor },
        });
        anchorRef = fallbackAnchor;
      }
    }
    const anchorObj =
      anchorRef !== undefined
        ? { blockRef: anchorRef }
        : destSheetHint
        ? { sheet: destSheetHint }
        : params.dest?.anchor;

    const placementParams = params.dest
      ? { ...params.dest, anchor: anchorObj, sheet: destSheetHint || params.dest.sheet, newSheetNameHint: destSheetHint || params.dest.newSheetNameHint }
      : { mode: "right", anchor: anchorObj || { blockRef: blockAddress }, sheet: destSheetHint, newSheetNameHint: destSheetHint };
    const confirmationId = `${ctx.step.id}:chartPlacement`;
    const resolved = await resolvePlacement({ ...placementParams, minBlankArea: { rows: 12, cols: 6 } }, ctx, confirmationId);
    if (resolved.requiresConfirmation) return { requiresConfirmation: resolved.requiresConfirmation };

    stage = "destSheet";
    const destSheetName = (resolved.address ? (resolved.address.includes("!") ? resolved.address.split("!")[0] : safeSheetName) : safeSheetName) as string;
    const destWs = resolved.worksheet ?? getWorksheet(ctx.excelCtx, destSheetName);
    const destBounds = resolved.address ? parseA1Address(resolved.address) : null;

    stage = "charts.add";
    const chartType = params.chartType || "columnClustered";
    if (!(destWs as any).charts) {
      (destWs as any).charts = { created: [], add: (t: any, r: any) => ({ type: t, sourceRange: r }) };
    }
    const chartCollection = destWs.charts || (destWs as any).charts || { add: () => ({}) };
    const normalizedChartType = `${chartType}`.toLowerCase();
    const inferredPlotBy = normalizedChartType.includes("line") ? "rows" : "columns";
    const plotBy = params.plotBy || inferredPlotBy;
    const sourceRange = sourceDataInfo.range || firstYRange;
    const sourceAddressForLog =
      sourceDataInfo.address || (firstYRange?.address ? firstYRange.address : blockAddress);
    const sourceMode = sourceDataInfo.range ? sourceDataInfo.mode ?? "contiguous" : "fallback";
    const chart = chartCollection.add
      ? chartCollection.add(chartType as any, sourceRange, plotBy)
      : { type: chartType, sourceRange };
    ctx.log({
      level: "info",
      message: "create_chart: chart source prepared",
      macro: "create_chart",
      stepId: ctx.step.id,
        data: {
          address: sourceAddressForLog,
          plotBy,
          mode: sourceMode,
          xCol: xAbs,
          yCols: orderedAbsYCols,
          includeHeaders: false,
        },
      });

    const createdArr = (destWs as any).charts?.created || (chartCollection as any).created;
    if (createdArr && !createdArr.includes(chart)) {
      createdArr.push(chart);
    } else if (!createdArr && (destWs as any).charts?.push) {
      try {
        (destWs as any).charts.push(chart);
      } catch {
        // ignore
      }
    }
    const assignedName = `AgentChart_${Date.now()}`;
    if ("name" in (chart as any)) {
      (chart as any).name = assignedName;
    }
    if (params.titleHint && (chart as any).title) (chart as any).title.text = params.titleHint;

    stage = "series.load";
    const seriesColl = (chart as any).series;
    try {
      if (seriesColl?.load) seriesColl.load("items");
      if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
    } catch (err: any) {
      const chartSourceLabel = sourceAddressForLog || blockAddress;
      ctx.log({
        level: "error",
        message: `Creation chart params invalides (series.load). source=${chartSourceLabel} chartType=${chartType} x=${xIdx} y=${yIdxs.join(",")} plotBy=${plotBy}`,
        macro: "create_chart",
        stepId: ctx.step.id,
        data: { error: err, mode: sourceMode },
      });
      return { status: "error" };
    }

    const setXAxis = (s: any, rng: any) => {
      if (typeof s?.setXAxisValues === "function") s.setXAxisValues(rng);
      else if (typeof s?.setXValues === "function") s.setXValues(rng);
    };

    stage = "series.configFirst";
    const fallbackToFirstRange = sourceDataInfo.range === null;
    const seriesNameIndexes = orderedRelativeYIdxs.length ? orderedRelativeYIdxs : [firstYIdx];
    const items = seriesColl?.items ?? [];
    if (items.length) {
      items.forEach((series: any, idx: number) => {
        setXAxis(series, xRange);
        const headerIdx = seriesNameIndexes[idx] ?? seriesNameIndexes[0];
        const clamped = typeof headerIdx === "number" ? Math.min(headerIdx, headers.length - 1) : null;
        const headerVal = clamped !== null && clamped >= 0 ? headers[clamped] : undefined;
        if (headerVal && "name" in series) (series as any).name = headerVal;
      });
      if (fallbackToFirstRange) {
        const firstSeries = items[0];
        if (typeof firstSeries.setValues === "function" && firstYRange) firstSeries.setValues(firstYRange);
      }
    }
    stage = "setPosition";
    const topLeft = destBounds ? rowColToA1(destBounds.startRow, destBounds.startCol) : "A1";
    const bottomRight = destBounds ? rowColToA1(destBounds.endRow, destBounds.endCol) : topLeft;
    if (typeof (chart as any).setPosition === "function") {
      const topLeftRange = destWs.getRange ? destWs.getRange(topLeft) : destWs.getRangeByIndexes?.(destBounds?.startRow ?? 0, destBounds?.startCol ?? 0, 1, 1);
      const bottomRightRange = destWs.getRange
        ? destWs.getRange(bottomRight)
        : destWs.getRangeByIndexes?.(destBounds?.endRow ?? 0, destBounds?.endCol ?? 0, 1, 1);
      (chart as any).setPosition(topLeftRange, bottomRightRange);
    }

    stage = "sync-final";
    if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
    ctx.log({
      level: "info",
      message: `Chart cree: ${assignedName} sur ${destSheetName}!${topLeft}`,
      macro: "create_chart",
      stepId: ctx.step.id,
    });
    return {
      artifacts: [
        {
          id: assignedName,
          type: "chart",
          sheet: destSheetName,
          anchor: topLeft,
          fromStep: ctx.step.id,
        },
      ],
    };
  } catch (err) {
    ctx.log({ level: "error", message: `Creation chart a echoue (stage=${stage}): ${err}`, macro: "create_chart", stepId: ctx.step.id });
    return {};
  }
};

const macroTableView: MacroFn = async (params, ctx) => {
  const source = params.source || {};
  const select: string[] = Array.isArray(params.select) ? params.select : [];
  const registry = ctx.headerRegistry || [];
  const resolvedSource = resolveBlockRefOrArtifact(source as BlockArtifactRef, ctx, { allowDefault: false });
  if (!resolvedSource.ok) {
    ctx.log({ level: "error", message: `table_view: source introuvable (${resolvedSource.reason})`, macro: "table_view", stepId: ctx.step.id });
    return { status: "error" };
  }
  const sourceRef = resolvedSource.blockRef;
  const data = await readBlockData(sourceRef, ctx);
  if (data.error) {
    ctx.log({ level: "error", message: "table_view: lecture source echouee", macro: "table_view", stepId: ctx.step.id });
    return { status: "error" };
  }
  const headers = data.headers || [];
  const rows = data.rows || [];
  const rowsText = (data as any).rowsText || [];
  const columnNumberFormats: (string | null)[] = (data as any).numberFormats || [];
  const mode = (params.dest?.mode || "newSheet") as string;
  const headerAliasLookup = buildHeaderAliasLookup((resolvedSource.artifact as any)?.headerAliases);
  const headersNormMap = buildCandidateNormMap(headers);
  ctx.log({
    level: "info",
    message: `table_view_source rows=${rows.length} headers=${headers.join("|")}`,
    macro: "table_view",
    stepId: ctx.step.id,
  });

  const resolveLastAdded = () => {
    if (ctx.lastAddedHeader) return ctx.lastAddedHeader;
    const dsRef = (ctx as any).datasetRef;
    if (dsRef) {
      const latestByDs = registry.slice().reverse().find((r) => r && r.datasetRef === dsRef && r.headerName);
      if (latestByDs?.headerName) return latestByDs.headerName;
    }
    const latest = registry.slice().reverse().find((r) => r && r.headerName);
    if (latest?.headerName) return latest.headerName;
    if (headers.length) return headers[headers.length - 1] || null;
    return null;
  };
  if (mode === "inPlace" && select.length > 0) {
    ctx.log({ level: "error", message: "table_view_inplace_no_select", macro: "table_view", stepId: ctx.step.id });
    return { status: "error" };
  }
  if (mode !== "inPlace" && select.length === 0) {
    ctx.log({ level: "error", message: `table_view_select_required | headers: ${headers.join(", ")}`, macro: "table_view", stepId: ctx.step.id });
    return { status: "error" };
  }
  const selectForResolution = mode === "inPlace" && select.length === 0 ? headers : select;
  const resolvedSelect: { header: string; idx: number }[] = [];
  const unresolvedSelect: string[] = [];
  selectForResolution.forEach((h) => {
    const headerToken = h === "$lastAddedColumn" ? resolveLastAdded() || h : h;
    const match = resolveHeaderWithAliases(headerToken, headers, headersNormMap, headerAliasLookup);
    if (match && resolvedSelect.findIndex((m) => m.idx === match.index) === -1) {
      resolvedSelect.push({ header: match.header, idx: match.index });
    } else if (!match) {
      unresolvedSelect.push(h);
    }
  });
  const selectedIdx: number[] = resolvedSelect.map((m) => m.idx);
  if (mode !== "inPlace" && unresolvedSelect.length > 0) {
    ctx.log({
      level: "error",
      message: `table_view: header introuvable ${unresolvedSelect.join(", ")} | disponibles: ${headers.join(", ")}`,
      macro: "table_view",
      stepId: ctx.step.id,
    });
    return { status: "error" };
  }
  if (mode !== "inPlace" && resolvedSelect.length === 0) {
    ctx.log({
      level: "error",
      message: `table_view_select_required | headers: ${headers.join(", ")}`,
      macro: "table_view",
      stepId: ctx.step.id,
    });
    return { status: "error" };
  }

  const resolvedSelectHeaders = resolvedSelect.map((m) => m.header);
  const resolvedSelectNormMap = buildCandidateNormMap(resolvedSelectHeaders);

  let outputHeaders = resolvedSelect.map((m) => m.header);
  const renameMapRaw: Record<string, string> = params.rename || {};
  const renameMap: Record<string, string> = {};
  const renameSource = mode === "inPlace" ? headers : resolvedSelect.map((m) => m.header);
  renameSource.forEach((h) => {
    const key = Object.keys(renameMapRaw).find((k) => normalizeHeader(k) === normalizeHeader(h));
    if (key && typeof renameMapRaw[key] === "string") renameMap[h] = renameMapRaw[key];
  });
  outputHeaders = outputHeaders.map((h) => renameMap[h] || h);

  const projectedRows = rows.map((r) => selectedIdx.map((idx: number) => (idx >= 0 && idx < r.length ? r[idx] : "")));
  const projectedTexts = rowsText.map((r: any[]) => selectedIdx.map((idx: number) => (idx >= 0 && idx < r.length ? r[idx] : "")));

  const items = projectedRows.map((row, i) => ({ row, text: projectedTexts[i] || [] as any[] }));

  const isExcelSerialDate = (val: any) => typeof val === "number" && val >= 20000 && val <= 80000;
  const excelSerialToDate = (val: number) => {
    const base = new Date(Date.UTC(1899, 11, 30));
    base.setUTCDate(base.getUTCDate() + Math.floor(val));
    return base;
  };
  const dateToExcelSerial = (d: Date) => {
    const base = Date.UTC(1899, 11, 30);
    const msPerDay = 24 * 60 * 60 * 1000;
    const day = Math.floor((Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - base) / msPerDay);
    return day;
  };
  const parseDateVal = (val: any): Date | null => {
    const parsed = parseDateCell(val, "fr-FR", {
      onFallback: (msg: string) => ctx.log({ level: "warn", message: msg, macro: "table_view", stepId: ctx.step.id }),
      onInvalid: (msg: string) => ctx.log({ level: "warn", message: msg, macro: "table_view", stepId: ctx.step.id }),
    });
    return parsed.date;
  };
  const formatLooksDate = (fmt?: string | null) => {
    if (!fmt) return false;
    const lower = fmt.toLowerCase();
    return lower.includes("yy") || (lower.includes("dd") && lower.includes("mm")) || lower.includes("mmm");
  };
  type ColumnMeta = { kind: "date" | "number" | "text"; numberFormat: string | null; sourceIdx: number };
  const headerToIdx = new Map<string, number>();
  outputHeaders.forEach((h, i) => headerToIdx.set(h, i));
	  const getDateCandidate = (item: { row: any[]; text: any[] }, idx: number) => {
	    const value = item.row[idx];
	    const rawText = item.text?.[idx];
	    const text =
	      typeof rawText === "string"
	        ? rawText
	        : rawText === null || typeof rawText === "undefined"
	        ? ""
	        : `${rawText}`;
	    return { text, value };
	  };

		  // Debug automatique (runtime) pour diagnostiquer les écarts Office.js vs FakeExcel.
		  // Limité à quelques lignes (debugDatesCount < 3) pour éviter le spam.
		  const debugDates = true;
	  let debugDatesCount = 0;
	  const maybeParseFrSlashParts = (text: string): { day: number; month: number } | null => {
	    const m = text.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
	    if (!m) return null;
	    const day = Number(m[1]);
	    const month = Number(m[2]);
	    if (!Number.isFinite(day) || !Number.isFinite(month)) return null;
	    return { day, month };
	  };

	  const getDateComparableTs = (
	    item: { row: any[]; text: any[] },
	    idx: number,
	    header?: string,
	    fmt?: string | null,
	    debug?: boolean
	  ): number | null => {
	    const candidate = getDateCandidate(item, idx);
	    const textTrim = (candidate.text || "").trim();
	    let textParsed: ReturnType<typeof parseDateCell> | null = null;
	    if (textTrim) {
	      textParsed = parseDateCell(textTrim, "fr-FR", {
	        onFallback: (msg: string) => ctx.log({ level: "warn", message: msg, macro: "table_view", stepId: ctx.step.id }),
	        onInvalid: (msg: string) => ctx.log({ level: "warn", message: msg, macro: "table_view", stepId: ctx.step.id }),
	      });
		      if (textParsed?.ts !== null) {
		        const v = candidate.value;
		        const parts = maybeParseFrSlashParts(textTrim);
		        const ambiguous = parts ? parts.day >= 1 && parts.day <= 12 && parts.month >= 1 && parts.month <= 12 : false;

		        // Runtime Office.js : `Range.text` peut être en mm/dd même si l’utilisateur voit du jj/mm.
		        // En cas de mismatch sur une date ambiguë, on fait confiance au serial Excel (valeur) qui est univoque.
		        if (ambiguous && isExcelSerialDate(v)) {
		          const serialParsed = parseDateCell(v, "fr-FR");
		          if (serialParsed.ts !== null && serialParsed.ts !== textParsed.ts) {
		            if (debugDates && debug && debugDatesCount < 3) {
		              debugDatesCount += 1;
		              ctx.log({
		                level: "info",
		                message: `table_view_date_debug col=${header || ""} idx=${idx} fmt=${fmt || ""} text=${textTrim} value=${v} ts=${serialParsed.ts} chosen=serial_mismatch`,
		                macro: "table_view",
		                stepId: ctx.step.id,
		              });
		            }
		            return serialParsed.ts;
		          }
		        }

		        if (debugDates && debug && debugDatesCount < 3) {
		          debugDatesCount += 1;
		          ctx.log({
		            level: "info",
		            message: `table_view_date_debug col=${header || ""} idx=${idx} fmt=${fmt || ""} text=${textTrim} value=${candidate.value} ts=${textParsed.ts} chosen=text`,
		            macro: "table_view",
		            stepId: ctx.step.id,
		          });
		        }
		        return textParsed.ts;
		      }
		    }

	    let valueParsed: ReturnType<typeof parseDateCell> | null = null;
	    const v = candidate.value;
	    if (v !== null && typeof v !== "undefined") {
	      if (typeof v === "number" && Number.isFinite(v) && !isExcelSerialDate(v)) {
	        valueParsed = null;
	      } else {
	        valueParsed = parseDateCell(v, "fr-FR", {
	          onFallback: (msg: string) => ctx.log({ level: "warn", message: msg, macro: "table_view", stepId: ctx.step.id }),
	          onInvalid: (msg: string) => ctx.log({ level: "warn", message: msg, macro: "table_view", stepId: ctx.step.id }),
	        });
	      }
	      if (textParsed && valueParsed && textParsed.ts !== null && valueParsed.ts !== null && textParsed.ts !== valueParsed.ts) {
	        ctx.log({
	          level: "info",
	          message: `table_view: date_mismatch_text_vs_value header=${header || ""} text=${textTrim} value=${v} using=text`,
	          macro: "table_view",
	          stepId: ctx.step.id,
	        });
	      }
	      if (!textTrim && valueParsed && valueParsed.ts !== null) {
	        if (debugDates && debug && debugDatesCount < 3) {
	          debugDatesCount += 1;
	          ctx.log({
	            level: "info",
	            message: `table_view_date_debug col=${header || ""} idx=${idx} fmt=${fmt || ""} text=<empty> value=${v} ts=${valueParsed.ts} chosen=value`,
	            macro: "table_view",
	            stepId: ctx.step.id,
	          });
	        }
	        return valueParsed.ts;
	      }
      if (textTrim && textParsed?.ts === null) {
        if (v instanceof Date && !Number.isNaN(v.getTime())) return v.getTime();
        const digitsOnly = typeof textTrim === "string" && isAllDigits(textTrim);
        if (digitsOnly) {
          const headerHint = headerSuggestsDate(header || "");
          const fmtHint = formatLooksDate(fmt);
          const digitsValue = Number(textTrim);
          const digitsLooksLikeSerial = Number.isFinite(digitsValue) && isExcelSerialDate(digitsValue);
          if (headerHint || fmtHint || digitsLooksLikeSerial) {
            const parsedDigits = parseDateCell(digitsValue, "fr-FR", {
              onFallback: (msg: string) => ctx.log({ level: "warn", message: msg, macro: "table_view", stepId: ctx.step.id }),
              onInvalid: (msg: string) => ctx.log({ level: "warn", message: msg, macro: "table_view", stepId: ctx.step.id }),
            });
            if (parsedDigits.ts !== null) return parsedDigits.ts;
          }
        }
      }
    }
    return null;
  };

  const columnMetas: ColumnMeta[] = outputHeaders.map((header, i) => {
    const sourceIdx = selectedIdx[i] ?? i;
    const fmt = columnNumberFormats[sourceIdx] || null;
    const colValues = items.map((it) => it.row[i]).filter((v) => !(v === null || typeof v === "undefined" || v === ""));

	    // sample-based detection text-first
	    const sample = items.slice(0, 15);
	    const parsedTs = sample
	      .map((it) => getDateComparableTs(it, i, header, fmt, false))
	      .filter((t): t is number => t !== null);
	    const successRate = parsedTs.length / (sample.length || 1);

    let numericVotes = 0;
    colValues.forEach((v) => {
      if (typeof v === "number" && Number.isFinite(v)) {
        numericVotes += 1;
        return;
      }
      const maybeNum = Number(v);
      if (Number.isFinite(maybeNum)) numericVotes += 0.3;
    });
    const nonEmpty = colValues.length || 1;
    const headerDateHint = headerSuggestsDate(header);
    const threshold = headerDateHint ? 0.4 : 0.6;

    if (formatLooksDate(fmt) || successRate >= threshold) return { kind: "date", numberFormat: fmt, sourceIdx };
    if (numericVotes / nonEmpty >= 0.6) return { kind: "number", numberFormat: fmt, sourceIdx };
    return { kind: "text", numberFormat: fmt, sourceIdx };
  });

  const startOfYear = (y: number) => new Date(Date.UTC(y, 0, 1));
  const endOfYear = (y: number) => new Date(Date.UTC(y, 11, 31, 23, 59, 59, 999));
  const parseYear = (val: any): number | null => {
    if (typeof val === "number" && Number.isInteger(val) && val >= 1000 && val <= 9999) return val;
    if (typeof val === "string") {
      const trimmed = val.trim();
      if (isAllDigits(trimmed) && trimmed.length === 4) return Number(trimmed);
    }
    return null;
  };
  const normalizeDateFilterValue = (
    val: any,
    op: string
  ): { date?: Date; min?: Date; max?: Date; negate?: boolean } | null => {
    const year = parseYear(val);
    if (year) {
      switch (op) {
        case "gt":
        case ">":
          return { min: startOfYear(year + 1) };
        case "gte":
        case ">=":
          return { min: startOfYear(year) };
        case "lt":
        case "<":
          return { max: endOfYear(year - 1) };
        case "lte":
        case "<=":
          return { max: endOfYear(year) };
        case "equals":
        case "eq":
        case "=":
          return { min: startOfYear(year), max: endOfYear(year) };
        case "neq":
        case "!=":
        case "<>":
          return { min: startOfYear(year), max: endOfYear(year), negate: true };
        default:
          return { date: startOfYear(year) };
      }
    }
    const parsed = parseDateVal(val);
    if (parsed) return { date: parsed };
    return null;
  };
  const compareDateTs = (ta: number | null, tb: number | null) => {
    const a = ta ?? Number.POSITIVE_INFINITY;
    const b = tb ?? Number.POSITIVE_INFINITY;
    return a - b;
  };
  const coerceBetweenBounds = (raw: any): [any, any] | null => {
  if (Array.isArray(raw) && raw.length >= 2) return [raw[0], raw[1]];

  // Support "a,b" (ISO ou FR)
  if (typeof raw === "string") {
    const s = raw.trim();
    // accepte comma ou ';'
    const parts = s.split(/[;,]/).map((p) => p.trim()).filter(Boolean);
    if (parts.length >= 2) return [parts[0], parts[1]];
    return null;
  }

  // Support {min,max} / {start,end}
  if (raw && typeof raw === "object") {
    const a = (raw as any).min ?? (raw as any).start ?? (raw as any).from;
    const b = (raw as any).max ?? (raw as any).end ?? (raw as any).to;
    if (typeof a !== "undefined" && typeof b !== "undefined") return [a, b];
  }

  return null;
  };

  const filtersRaw: any[] = Array.isArray(params.filter) ? params.filter : [];
  if (!Array.isArray(params.filter) && params.filter) {
    ctx.log({ level: "error", message: "table_view: filter invalid shape", macro: "table_view", stepId: ctx.step.id });
    return { status: "error" };
  }
  const filters: { col: string; op: string; value?: any }[] = [];
  for (const f of filtersRaw) {
    if (!f || typeof f !== "object") {
      ctx.log({ level: "error", message: "table_view: filter invalid", macro: "table_view", stepId: ctx.step.id });
      return { status: "error" };
    }
    const colNameRaw = (f.col || (f as any).column || (f as any).field || "").toString();
    const colName = colNameRaw === "$lastAddedColumn" ? resolveLastAdded() || colNameRaw : colNameRaw;
    const op = (f.op || (f as any).operator || "").toString();
    if (!colName || !op) {
      ctx.log({ level: "error", message: "table_view: filter missing col/op", macro: "table_view", stepId: ctx.step.id });
      return { status: "error" };
    }
    if (!["equals", "eq", "neq", "notEmpty", "isEmpty", "contains", "not_contains", "between", "gt", "gte", "lt", "lte", "in"].includes(op)) {
      ctx.log({ level: "error", message: `table_view: filter op invalid ${op}`, macro: "table_view", stepId: ctx.step.id });
      return { status: "error" };
    }
    const resolveCol = () => {
      const matchOutput = matchHeaderToken(outputHeaders, colName);
      if (matchOutput) return outputHeaders[matchOutput.index];
      const matchOriginal = resolveHeaderWithAliases(colName, resolvedSelectHeaders, resolvedSelectNormMap, headerAliasLookup);
      if (matchOriginal) return outputHeaders[matchOriginal.index];
      return null;
    };
    const resolvedCol = resolveCol();
    if (!resolvedCol) {
      ctx.log({ level: "error", message: `table_view: filtre colonne introuvable ${colName} | disponibles: ${outputHeaders.join(", ")}`, macro: "table_view", stepId: ctx.step.id });
      return { status: "error" };
    }
    filters.push({ col: resolvedCol, op: op as any, value:
  (f as any).value ??
  (f as any).values ??
  (f as any).val ??
  ((f as any).from !== undefined || (f as any).to !== undefined
    ? [(f as any).from, (f as any).to]
    : undefined)});
  }

  const sortRaw = params.sort;
	  let sort: { col: string; dir: "asc" | "desc" } | null = null;
	  if (sortRaw && typeof sortRaw === "object") {
    const colNameRaw = (sortRaw.col || (sortRaw as any).column || (sortRaw as any).field || "").toString();
    const colName = colNameRaw === "$lastAddedColumn" ? resolveLastAdded() || colNameRaw : colNameRaw;
    const dirRaw = (sortRaw.dir || (sortRaw as any).direction || "").toString().toLowerCase();
    let resolvedSort = matchHeaderToken(outputHeaders, colName);
    if (!resolvedSort) {
      const aliasSort = resolveHeaderWithAliases(colName, resolvedSelectHeaders, resolvedSelectNormMap, headerAliasLookup);
      if (aliasSort) {
        resolvedSort = { header: aliasSort.header, index: aliasSort.index, score: 1, via: "alias" };
      }
    }
    if (resolvedSort) {
      const colResolved = outputHeaders[resolvedSort.index];
      if (colResolved) sort = { col: colResolved, dir: dirRaw === "desc" ? "desc" : "asc" };
      ctx.log({ level: "info", message: `table_view_resolve_lastAddedColumn sort=${colResolved}`, macro: "table_view", stepId: ctx.step.id });
    } else {
      ctx.log({ level: "error", message: `table_view: sort colonne introuvable ${colName} | disponibles: ${outputHeaders.join(", ")}`, macro: "table_view", stepId: ctx.step.id });
      return { status: "error" };
    }
	  } else if (sortRaw) {
	    ctx.log({ level: "error", message: "table_view: sort invalid shape", macro: "table_view", stepId: ctx.step.id });
	    return { status: "error" };
	  }

	  // Debug (runtime): confirmer résolution sort/filtre et kind final (date/number/text).
	  if (filters.length) {
	    const sample = filters.slice(0, 2).map((f) => {
	      const idx = headerToIdx.get(f.col);
	      const meta = typeof idx === "number" ? columnMetas[idx] : null;
	      return `${f.col}@${idx ?? "?"}:${meta?.kind ?? "?"}:${f.op}`;
	    });
	    ctx.log({
	      level: "info",
	      message: `table_view_filter_branch executed filters=${filters.length} resolved=${sample.join(",")}`,
	      macro: "table_view",
	      stepId: ctx.step.id,
	    });
	  }
	  if (sort) {
	    const idx = headerToIdx.get(sort.col);
	    const meta = typeof idx === "number" ? columnMetas[idx] : null;
	    ctx.log({
	      level: "info",
	      message: `table_view_sort_branch executed col=${sort.col} idx=${idx ?? "?"} kind=${meta?.kind ?? "?"} fmt=${meta?.numberFormat ?? ""} dir=${sort.dir}`,
	      macro: "table_view",
	      stepId: ctx.step.id,
	    });
	  }

	  const filteredItems = items.filter((item) => {
	    for (const f of filters) {
	      const idx = headerToIdx.get(f.col);
	      if (idx === undefined) continue;
	      const meta = columnMetas[idx];
	      const cellTs = getDateComparableTs(item, idx, f.col, meta?.numberFormat ?? null, true);
	      const opNorm = (f.op || "").toString().toLowerCase();
      if (opNorm === "notempty" || f.op === "notEmpty") {
        if (cellTs === null && (item.row[idx] === null || typeof item.row[idx] === "undefined" || item.row[idx] === "") && (item.text?.[idx] === "" || typeof item.text?.[idx] === "undefined"))
          return false;
        continue;
      }
      if (opNorm === "isempty" || f.op === "isEmpty") {
        if (!((cellTs === null || Number.isNaN(cellTs)) && (item.row[idx] === null || typeof item.row[idx] === "undefined" || item.row[idx] === "") && (item.text?.[idx] === "" || typeof item.text?.[idx] === "undefined")))
          return false;
        continue;
      }

      if (meta?.kind === "date") {
        if (cellTs === null) {
          ctx.log({ level: "warn", message: `table_view: date parse failed col=${f.col}`, macro: "table_view", stepId: ctx.step.id });
          return false;
        }
        if (opNorm === "between") {
        const b = coerceBetweenBounds(f.value);
        if (!b) {
          ctx.log({
            level: "warn",
            message: `table_view: date between invalid shape value=${typeof f.value === "string" ? f.value : JSON.stringify(f.value)}`,
            macro: "table_view",
            stepId: ctx.step.id,
          });
          return false;
        }

        const [rawMin, rawMax] = b;

        // bornes inclusives
        const minNorm = normalizeDateFilterValue(rawMin, "gte");
        const maxNorm = normalizeDateFilterValue(rawMax, "lte");

        const min = (minNorm?.min || minNorm?.date) ?? null;
        const max = (maxNorm?.max || maxNorm?.date) ?? null;

        const minTs = min ? min.getTime() : null;
        const maxTs = max ? max.getTime() : null;

        if (minTs === null || maxTs === null) {
          ctx.log({
            level: "warn",
            message: `table_view: date between unparsable bounds min=${String(rawMin)} max=${String(rawMax)}`,
            macro: "table_view",
            stepId: ctx.step.id,
          });
          return false;
        }

        if (cellTs < minTs || cellTs > maxTs) return false;
        continue;
      }

        const normVal = normalizeDateFilterValue(f.value, opNorm);
        if (!normVal) {
          ctx.log({ level: "warn", message: "table_view: date filter incompatible", macro: "table_view", stepId: ctx.step.id });
          return false;
        }
        const refMin = normVal.min || normVal.date || null;
        const refMax = normVal.max || normVal.date || null;
        const refMinTs = refMin ? refMin.getTime() : null;
        const refMaxTs = refMax ? refMax.getTime() : null;
        switch (opNorm) {
          case "equals":
          case "eq":
          case "=":
            if (refMinTs === null || refMaxTs === null) return false;
            if (cellTs < refMinTs || cellTs > refMaxTs) return false;
            break;
          case "neq":
          case "!=":
          case "<>":
            if (refMinTs !== null && refMaxTs !== null && cellTs >= refMinTs && cellTs <= refMaxTs) return false;
            break;
          case "gt":
          case ">":
            if (refMinTs === null || !(cellTs > refMinTs)) return false;
            break;
          case "gte":
          case ">=":
            if (refMinTs === null || !(cellTs >= refMinTs)) return false;
            break;
          case "lt":
          case "<":
            if (refMaxTs === null || !(cellTs < refMaxTs)) return false;
            break;
          case "lte":
          case "<=":
            if (refMaxTs === null || !(cellTs <= refMaxTs)) return false;
            break;
          case "in":
            if (Array.isArray(f.value)) {
              const parsedList = f.value
                .map((v: any) => parseDateCell(v, "fr-FR"))
                .map((p) => p.ts)
                .filter((t): t is number => typeof t === "number" && Number.isFinite(t));
              if (!parsedList.some((t) => t === cellTs)) return false;
            } else {
              ctx.log({ level: "warn", message: "table_view: date in expects array", macro: "table_view", stepId: ctx.step.id });
              return false;
            }
            break;
          default:
            break;
        }
        continue;
      }

      switch (f.op) {
        case "equals":
        case "eq":
          if (item.row[idx] !== f.value && item.text?.[idx] !== f.value) return false;
          break;
        case "neq":
          if (item.row[idx] === f.value || item.text?.[idx] === f.value) return false;
          break;
        case "contains": {
          const source = item.text?.[idx] ?? item.row[idx];
          const c = typeof source === "string" ? source : `${source ?? ""}`;
          if (typeof f.value !== "string") {
            ctx.log({ level: "warn", message: "table_view: contains type incompatible", macro: "table_view", stepId: ctx.step.id });
            continue;
          }
          if (!c.toLowerCase().trim().includes(f.value.toLowerCase().trim())) return false;
          break;
        }
        case "not_contains": {
          const source = item.text?.[idx] ?? item.row[idx];
          const c = typeof source === "string" ? source : `${source ?? ""}`;
          if (typeof f.value !== "string") {
            ctx.log({ level: "warn", message: "table_view: contains type incompatible", macro: "table_view", stepId: ctx.step.id });
            continue;
          }
          if (c.toLowerCase().trim().includes(f.value.toLowerCase().trim())) return false;
          break;
        }
        case "between": {
          const bounds = Array.isArray(f.value) && f.value.length >= 2 ? f.value : [];
          const min = bounds[0];
          const max = bounds[1];
          const source = item.row[idx];
          const num = typeof source === "number" ? source : Number(source);
          const minNum = typeof min === "number" ? min : Number(min);
          const maxNum = typeof max === "number" ? max : Number(max);
          if (!Number.isFinite(num) || !Number.isFinite(minNum) || !Number.isFinite(maxNum)) {
            ctx.log({ level: "warn", message: "table_view: between type incompatible", macro: "table_view", stepId: ctx.step.id });
            continue;
          }
          if (num < minNum || num > maxNum) return false;
          break;
        }
        case "gt":
        case "gte":
        case "lt":
        case "lte": {
          const kind = meta?.kind as ColumnMeta["kind"] | undefined;
          if (kind === "date") {
            const cellDateTs = cellTs;
            const cmpDate = parseDateVal(f.value);
            if (cellDateTs === null || !cmpDate) {
              ctx.log({ level: "warn", message: "table_view: date filter unparsable", macro: "table_view", stepId: ctx.step.id });
              continue;
            }
            const diff = compareDateTs(cellDateTs, cmpDate.getTime());
            if (f.op === "gt" && !(diff > 0)) return false;
            if (f.op === "gte" && !(diff >= 0)) return false;
            if (f.op === "lt" && !(diff < 0)) return false;
            if (f.op === "lte" && !(diff <= 0)) return false;
            break;
          }
          const source = item.row[idx];
          const num = typeof source === "number" ? source : Number(source);
          const val = typeof f.value === "number" ? f.value : Number(f.value);
          if (!Number.isFinite(num) || !Number.isFinite(val)) {
            ctx.log({ level: "warn", message: "table_view: numeric filter incompatible", macro: "table_view", stepId: ctx.step.id });
            continue;
          }
          if (f.op === "gt" && !(num > val)) return false;
          if (f.op === "gte" && !(num >= val)) return false;
          if (f.op === "lt" && !(num < val)) return false;
          if (f.op === "lte" && !(num <= val)) return false;
          break;
        }
        case "in": {
          if (Array.isArray(f.value)) {
            const source = item.row[idx] ?? item.text?.[idx];
            if (!f.value.some((v: any) => v === source)) return false;
          } else {
            ctx.log({ level: "warn", message: "table_view: in value should be array", macro: "table_view", stepId: ctx.step.id });
          }
          break;
        }
        default:
          break;
      }
    }
    return true;
  });
  if (filters.length) {
    ctx.log({
      level: "info",
      message: `table_view_filter_applied count=${filteredItems.length} cols=${filters.map((f) => f.col).join("|")} ops=${filters.map((f) => f.op).join("|")}`,
      macro: "table_view",
      stepId: ctx.step.id,
    });
  }

	  const applySort = (itemsToSort: { row: any[]; text: any[] }[]) => {
	    if (!sort) return;
	    const idx = headerToIdx.get(sort.col);
	    if (idx === undefined) return;
	    const factor = sort.dir === "desc" ? -1 : 1;
		    const meta = columnMetas[idx];
		    ctx.log({
		      level: "info",
		      message: `table_view_sort_branch_apply col=${sort.col} idx=${idx} kind=${meta?.kind ?? "?"} fmt=${meta?.numberFormat ?? ""} dir=${sort.dir}`,
		      macro: "table_view",
		      stepId: ctx.step.id,
		    });
		    itemsToSort.sort((a, b) => {
	      if (meta?.kind === "date") {
	        const ta = getDateComparableTs(a, idx, sort.col, meta?.numberFormat ?? null, true);
	        const tb = getDateComparableTs(b, idx, sort.col, meta?.numberFormat ?? null, true);
	        if (ta === tb) return 0;
	        if (ta === null) return 1;
	        if (tb === null) return -1;
	        return (ta - tb) * factor;
	      }
      const va = a.row[idx];
      const vb = b.row[idx];
      if (va === vb) return 0;
      if (va === undefined || va === null || va === "") return 1 * factor;
      if (vb === undefined || vb === null || vb === "") return -1 * factor;
      if (typeof va === "number" && typeof vb === "number") return (va - vb) * factor;
      return `${va}`.localeCompare(`${vb}`) * factor;
    });
    ctx.log({ level: "info", message: `table_view_sort_applied col=${sort.col} dir=${sort.dir}`, macro: "table_view", stepId: ctx.step.id });
  };

  applySort(filteredItems);

  const normalizedRows = filteredItems.map((item) =>
    item.row.map((cell, idx) => {
      const meta = columnMetas[idx];
	      if (meta?.kind === "date") {
	        const ts = getDateComparableTs(item, idx, outputHeaders[idx], meta?.numberFormat ?? null, false);
	        if (ts !== null) {
	          const d = new Date(ts);
	          return dateToExcelSerial(d);
	        }
        return cell;
      }
      return cell;
    })
  );
  if (columnMetas.some((m) => m.kind === "date")) {
    ctx.log({ level: "info", message: `table_view_date_write_mode mode=serial`, macro: "table_view", stepId: ctx.step.id });
  }

  const outputMatrix = [outputHeaders, ...normalizedRows];
  const rowCount = outputMatrix.length || 1;
  const colCount = outputHeaders.length || 1;
  const dest = params.dest || { mode: "newSheet" };

  const anchorResolved =
    dest.anchor && (dest.anchor as any).artifactRef
      ? resolveBlockRefOrArtifact(dest.anchor as BlockArtifactRef, ctx, { allowDefault: false })
      : null;
  const anchor = dest.anchor
    ? {
        ...dest.anchor,
        blockRef: anchorResolved && anchorResolved.ok ? anchorResolved.blockRef : dest.anchor.blockRef,
      }
    : { blockRef: sourceRef };

  if (dest.mode === "inPlace") {
    // rename (in-place) with already-sorted rows
    const outMatrix = [outputHeaders, ...filteredItems.map((it) => it.row)];
    const inPlaceRowCount = outMatrix.length || 1;
    const inPlaceColCount = outputHeaders.length || 1;
    const ws = getWorksheet(ctx.excelCtx, data.sheetName);
    const headerRow = data.bounds?.startRow ?? 0;
    headers.forEach((h, i) => {
      if (renameMap[h]) {
        const addr = rowColToA1(headerRow, data.bounds?.startCol ? data.bounds.startCol + i : i);
        const cell = ws.getRange ? ws.getRange(addr) : ws.getRangeByIndexes?.(headerRow, (data.bounds?.startCol || 0) + i, 1, 1);
        if (cell) (cell as any).values = [[renameMap[h]]];
      }
    });
    if (data.bounds) {
      const startRow = data.bounds.startRow;
      const startCol = data.bounds.startCol;
      const endRow = startRow + inPlaceRowCount - 1;
      const endCol = startCol + inPlaceColCount - 1;
      const localAddr = `${rowColToA1(startRow, startCol)}:${rowColToA1(endRow, endCol)}`;
      const range = ws.getRange ? ws.getRange(localAddr) : ws.getRangeByIndexes?.(startRow, startCol, inPlaceRowCount, inPlaceColCount);
      if (range) (range as any).values = outMatrix;
    }
    const artifactBlockRef = sourceRef;
    ctx.log({
      level: "info",
      message: `table_view inPlace headers updated on ${artifactBlockRef}`,
      macro: "table_view",
      stepId: ctx.step.id,
    });
    return {
      artifacts: [
        {
          type: "table",
          kind: "table",
          sheet: data.sheetName,
          sheetName: data.sheetName,
          anchor: rowColToA1(data.bounds?.startRow ?? 0, data.bounds?.startCol ?? 0),
          fromStep: ctx.step.id,
          tableName: (resolvedSource.artifact as any)?.tableName,
          blockRef: artifactBlockRef,
          address: data.bounds ? `${rowColToA1(data.bounds.startRow, data.bounds.startCol)}:${rowColToA1(data.bounds.endRow, data.bounds.endCol)}` : "",
          addressA1: data.bounds ? `${rowColToA1(data.bounds.startRow, data.bounds.startCol)}:${rowColToA1(data.bounds.endRow, data.bounds.endCol)}` : "",
          headers: outputHeaders,
          rowCount: inPlaceRowCount,
          colCount: inPlaceColCount,
          rows: inPlaceRowCount > 0 ? inPlaceRowCount - 1 : 0,
          cols: inPlaceColCount,
        },
      ],
    };
  }

  const placement = await resolvePlacement(
    {
      mode: dest.mode || "newSheet",
      anchor,
      sheet: dest.sheetName || dest.sheet,
      minBlankArea: { rows: rowCount, cols: colCount },
      avoidOverwrite: true,
      newSheetNameHint: dest.sheetName || params.outputTableName || `View_${ctx.step.id}`,
    },
    ctx,
    `${ctx.step.id}:table_view_place`
  );
  if (placement.requiresConfirmation) return { requiresConfirmation: placement.requiresConfirmation };
  if (!placement.address) return placement;

  const destSheetName =
    placement.address.split("!")[0] ||
    dest.sheetName ||
    data.sheetName ||
    ctx.context.active.sheetName ||
    (ctx.context.sheets[0]?.name ?? "Sheet1");
  const ws = placement.worksheet || getWorksheet(ctx.excelCtx, destSheetName);
  const placementBounds = parseA1Address(placement.address)!;
  const localAddr = placement.address.includes("!") ? placement.address.split("!")[1] || placement.address : placement.address;
  const range = ws.getRange ? ws.getRange(localAddr) : ws.getRangeByIndexes?.(placementBounds.startRow, placementBounds.startCol, rowCount, colCount);
  const loggedFormats = new Set<string>();
  const applyNumberFormats = (rng: any) => {
    if (!rng) return;
    const nfMatrix = Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => null as any));
    const FR_DATE_FMT = "dd/mm/yyyy";
    columnMetas.forEach((meta, idx) => {
      const fmt = meta.kind === "date" ? FR_DATE_FMT : (meta.numberFormat || null);
      if (fmt) {
        for (let r = 0; r < rowCount; r += 1) nfMatrix[r]![idx] = fmt;
        const logKey = `${outputHeaders[idx]}|${fmt}`;
        if (!loggedFormats.has(logKey)) {
          ctx.log({ level: "info", message: `table_view_format_copied col=${outputHeaders[idx]} numberFormat=${fmt}`, macro: "table_view", stepId: ctx.step.id });
          loggedFormats.add(logKey);
        }
      }
    });
    if (nfMatrix.some((rowFmt) => rowFmt.some((v) => v))) {
      (rng as any).numberFormat = nfMatrix;
    }
  };
  if (range) {
    (range as any).values = outputMatrix;
    applyNumberFormats(range);
  }
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();

  const opts = params.options || {};
  const styleAsTable = opts.styleAsTable !== false;
  let tableName: string | undefined = params.outputTableName || `View_${ctx.step.id}`;
  let tableCreated = false;
  if (styleAsTable && (ws as any).tables?.add) {
    try {
      const tbl = ws.tables.add(localAddr, true);
      tableCreated = true;
      if (tbl) {
        tableName = tableName || ensureTableName(ws, "View");
        (tbl as any).name = tableName;
      }
    } catch (err: any) {
      ctx.log({ level: "warn", message: `table_view: table creation failed ${err?.message || err}`, macro: "table_view", stepId: ctx.step.id });
    }
  }
  // table creation can reset number formats, ensure we reapply
  applyNumberFormats(range);
  if (opts.freezeHeader && (ws as any).freezePanes?.freezeRows) {
    try {
      (ws as any).freezePanes.freezeRows(1);
    } catch {}
  }

  try {
    await macroApplyFormat(
      { target: { blockRef: `${destSheetName}!${localAddr}`, tableName }, options: { preset: "corporate_blue", freezeHeaderRow: true } },
      { ...ctx, step: { ...ctx.step, id: `${ctx.step.id}:corporate_blue` } } as any
    );
    ctx.log({
      level: "info",
      message: `corporate_blue_format_applied table=${tableName || localAddr}`,
      macro: "table_view",
      stepId: ctx.step.id,
    });
    applyNumberFormats(range);
  } catch (err: any) {
    ctx.log({ level: "warn", message: `corporate_blue format failed: ${err?.message || err}`, macro: "table_view", stepId: ctx.step.id });
  }

  const artifactBlockRef = `${destSheetName}!${localAddr}`;
  const artifacts = placement.artifacts ? [...placement.artifacts] : [];
  artifacts.push({
    type: "table",
    kind: "table",
    sheet: destSheetName,
    sheetName: destSheetName,
    anchor: rowColToA1(placementBounds.startRow, placementBounds.startCol),
    fromStep: ctx.step.id,
    tableName,
    blockRef: artifactBlockRef,
    address: localAddr,
    addressA1: localAddr,
    rows: normalizedRows.length,
    cols: outputHeaders.length,
    rowCount: outputMatrix.length,
    colCount: outputHeaders.length,
    headers: outputHeaders.slice(),
    details: { tableCreated },
  });

  ctx.log({
    level: "info",
    message: `table_view output blockRef=${artifactBlockRef} tableName=${tableName || ""}`,
    macro: "table_view",
    stepId: ctx.step.id,
  });

  return { artifacts };
};

const macroValidateData: MacroFn = async (params: ValidateDataParams = {}, ctx: MacroContext) => {
  const clampMaxIssues = (raw?: number) => {
    if (typeof raw !== "number") return 1000;
    const value = Math.floor(raw);
    if (Number.isNaN(value)) return 1000;
    return Math.max(1, Math.min(5000, value));
  };
  const isMissingValue = (value: any) =>
    value === null ||
    typeof value === "undefined" ||
    (typeof value === "string" && value.trim() === "");
  const normalizeDuplicateKey = (value: any) => {
    if (value === null || typeof value === "undefined") return "";
    if (typeof value === "string") return value.trim().toLowerCase();
    return `${value}`.trim().toLowerCase();
  };
  const getRowKeyFromDataRow = (row: any[], headers: string[]): string => {
    if (!Array.isArray(row)) return "";
    const idIdx = headers.findIndex((header) => normalizeHeader(header) === "id");
    if (idIdx < 0) return "";
    const rawId = row[idIdx];
    if (rawId === null || typeof rawId === "undefined") return "";
    const idValue = `${rawId}`.trim();
    return idValue ? `id:${idValue}` : "";
  };
  type RowRef = { excelRow1: number; rowIdx: number; dataRowIdx: number };
  const buildLiveRowIndex = async (
    ctx: MacroContext,
    dataSheetName: string,
    dataBounds: AddressBounds | null,
    headers: string[],
    rangeOverride?: string | null
  ): Promise<{ index: Map<string, RowRef[]>; rows: unknown[][]; bounds: AddressBounds | null }> => {
    const index: Map<string, RowRef[]> = new Map();
    if (!dataBounds && !rangeOverride) {
      return { index, rows: [], bounds: null };
    }
    const rangeAddress =
      rangeOverride ||
      `${dataSheetName}!${rowColToA1(dataBounds!.startRow, dataBounds!.startCol)}:${rowColToA1(
        dataBounds!.endRow,
        dataBounds!.endCol
      )}`;
    const data = await readBlockData(rangeAddress, ctx);
    if (!data?.bounds) {
      return { index, rows: [], bounds: null };
    }
    const liveRows = Array.isArray(data.rows) ? (data.rows as unknown[][]) : [];
    liveRows.forEach((row, idx) => {
      const rowKey = getRowKeyFromDataRow(row, headers);
      if (!rowKey) return;
      const baseRow = data.bounds?.startRow ?? dataBounds?.startRow ?? 0;
      const excelRow1 = baseRow + 2 + idx;
      const entry: RowRef = { excelRow1, rowIdx: baseRow + 1 + idx, dataRowIdx: idx };
      const arr: RowRef[] = index.get(rowKey) ?? [];
      arr.push(entry);
      index.set(rowKey, arr);
    });
    return { index, rows: liveRows, bounds: data.bounds };
  };
  const ensureIdColumn = async (
    ctx: MacroContext,
    sheetName: string,
    bounds: AddressBounds | null,
    headers: string[],
    rows: any[][]
  ): Promise<{ headers: string[]; rows: any[][]; bounds: AddressBounds | null; inserted: boolean }> => {
    if (!bounds) return { headers, rows, bounds, inserted: false };
    const hasId = headers.some((header) => normalizeHeader(header) === "id");
    if (hasId) return { headers, rows, bounds, inserted: false };
    const ws = getWorksheet(ctx.excelCtx, sheetName);
    const totalRows = bounds.endRow - bounds.startRow + 1;
    const dataRows = rows.length;
    const hashString = (text: string) => {
      let h = 0;
      for (let i = 0; i < text.length; i += 1) {
        h = (Math.imul(31, h) + text.charCodeAt(i)) | 0;
      }
      return (h >>> 0).toString(36);
    };
    const serializeCell = (cell: any) => {
      if (cell === null || typeof cell === "undefined") return "";
      if (cell instanceof Date && !Number.isNaN(cell.getTime())) return cell.toISOString();
      if (typeof cell === "object") {
        try {
          return JSON.stringify(cell);
        } catch {
          return `${cell}`;
        }
      }
      return `${cell}`;
    };
    const baseSeed = hashString([sheetName, ...headers].join("\u001e"));
    const seenRowSignatures = new Map<string, number>();
    const makeRowId = (row: any[]) => {
      const signature = row.map((cell) => serializeCell(cell)).join("\u001f");
      const signatureHash = hashString(signature);
      const occurrence = (seenRowSignatures.get(signature) || 0) + 1;
      seenRowSignatures.set(signature, occurrence);
      return `vd_${baseSeed}_${signatureHash}_${occurrence}`;
    };
    const insertRange =
      (typeof ws.getRangeByIndexes === "function" &&
        ws.getRangeByIndexes(bounds.startRow, bounds.startCol, totalRows, 1)) ||
      (typeof ws.getRange === "function" && ws.getRange(rowColToA1(bounds.startRow, bounds.startCol)));
    const canInsert = typeof insertRange?.insert === "function";
    if (canInsert) {
      if (typeof Excel !== "undefined" && Excel.InsertShiftDirection) {
        insertRange.insert(Excel.InsertShiftDirection.right);
      } else {
        insertRange.insert("Right" as any);
      }
    }
    const insertedValues = Array.from({ length: totalRows }, (_v, idx) => {
      if (idx === 0) return ["ID"];
      const row = Array.isArray(rows[idx - 1]) ? (rows[idx - 1] as any[]) : [];
      return [makeRowId(row)];
    });
    const expandedValues = Array.from({ length: totalRows }, (_v, idx) => {
      if (idx === 0) return ["ID", ...headers];
      const rowValues = Array.isArray(rows[idx - 1]) ? (rows[idx - 1] as any[]) : [];
      return [makeRowId(rowValues), ...rowValues];
    });
    const writeWidth = canInsert ? 1 : bounds.endCol - bounds.startCol + 2;
    const writeRange =
      (typeof ws.getRangeByIndexes === "function" &&
        ws.getRangeByIndexes(bounds.startRow, bounds.startCol, totalRows, writeWidth)) ||
      (typeof ws.getRange === "function" &&
        ws.getRange(
          `${rowColToA1(bounds.startRow, bounds.startCol)}:${rowColToA1(
            bounds.startRow + totalRows - 1,
            bounds.startCol + writeWidth - 1
          )}`
        ));
    if (writeRange) {
      writeRange.values = canInsert ? insertedValues : expandedValues;
    }
    if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
    const newBounds: AddressBounds = {
      startRow: bounds.startRow,
      endRow: bounds.endRow,
      startCol: bounds.startCol,
      endCol: bounds.endCol + 1,
    };
    const refreshed = await readBlockData(
      `${sheetName}!${rowColToA1(newBounds.startRow, newBounds.startCol)}:${rowColToA1(newBounds.endRow, newBounds.endCol)}`,
      ctx
    );
    const updatedHeaders = Array.isArray(refreshed.headers) ? refreshed.headers : headers;
    const updatedRows = Array.isArray(refreshed.rows) ? refreshed.rows : rows;
    const updatedBounds = refreshed.bounds || newBounds;
    ctx.log({
      level: "info",
      message: "validate_data id_column_added",
      data: {
        sheet: sheetName,
        oldBounds: bounds,
        newBounds: updatedBounds,
        insertedAtCol: bounds.startCol,
        rowCount: dataRows,
      },
    });
    return { headers: updatedHeaders, rows: updatedRows, bounds: updatedBounds, inserted: canInsert };
  };
  type NumberRepairResult =
    | { kind: "num"; num: number; formatUsed: string; reason?: string }
    | { kind: "unfixable"; reason: string; formatUsed?: string };
  type DateRepairResult =
    | {
        kind: "date";
        serial: number;
        parsedDay: number;
        parsedMonth: number;
        parsedYear: number;
        formatUsed: string;
        reason?: string;
      }
    | { kind: "unfixable"; reason: string; formatUsed?: string };
  type ColumnClassification = "TEXT_MAJORITY" | "NUM_OR_DATE_MAJORITY" | "MIXED";
  type ColumnStats = {
    classification: ColumnClassification;
    countNonEmpty: number;
    countConvertible: number;
    countTextLike: number;
  };
  const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const DATE_SERIAL_POLICY: "never" | "threshold" = "threshold";
  const SERIAL_MIN = 1000;
  const SERIAL_MAX = 60000;

  const normalizeYear = (year: number) => {
    if (!Number.isFinite(year)) return year;
    const rounded = Math.floor(year);
    if (rounded >= 1000) return rounded;
    if (rounded >= 0 && rounded <= 99) {
      return rounded >= 50 ? 1900 + rounded : 2000 + rounded;
    }
    return rounded;
  };

  const validateDatePartsStrict = (year: number, month0: number, day: number) => {
    if (!Number.isFinite(year) || !Number.isFinite(month0) || !Number.isFinite(day)) return false;
    if (!Number.isInteger(year) || !Number.isInteger(month0) || !Number.isInteger(day)) return false;
    if (year < 1900 || year > 9999) return false;
    if (month0 < 0 || month0 > 11) return false;
    if (day < 1) return false;
    const date = new Date(Date.UTC(year, month0, day));
    return (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month0 &&
      date.getUTCDate() === day
    );
  };

  const toExcelSerial = (year: number, month0: number, day: number): number | null => {
    if (!validateDatePartsStrict(year, month0, day)) return null;
    const utc = Date.UTC(year, month0, day);
    const diff = utc - EXCEL_EPOCH_MS;
    if (!Number.isFinite(diff)) return null;
    return Math.round(diff / MS_PER_DAY);
  };

  const numberAsSerial = (value: number): DateRepairResult => {
    if (!Number.isFinite(value)) return { kind: "unfixable", reason: "serial_not_finite" };
    if (DATE_SERIAL_POLICY === "threshold" && (value < SERIAL_MIN || value > SERIAL_MAX)) {
      return { kind: "unfixable", reason: "serial_out_of_threshold" };
    }
    const date = new Date(EXCEL_EPOCH_MS + value * MS_PER_DAY);
    if (!Number.isFinite(date.getTime())) {
      return { kind: "unfixable", reason: "serial_invalid" };
    }
    const year = date.getUTCFullYear();
    const month0 = date.getUTCMonth();
    const day = date.getUTCDate();
    if (!validateDatePartsStrict(year, month0, day)) {
      return { kind: "unfixable", reason: "serial_invalid_parts" };
    }
    return {
      kind: "date",
      serial: value,
      parsedDay: day,
      parsedMonth: month0 + 1,
      parsedYear: year,
      formatUsed: "serial",
      reason: "serial",
    };
  };

  const monthNameLookup = new Map<string, number>();
  ([
    ["janvier", "janv", "jan", "january", "jan"],
    ["février", "fevrier", "fev", "february", "feb"],
    ["mars", "mars", "mar", "march"],
    ["avril", "avr", "apr", "april"],
    ["mai", "mai", "may"],
    ["juin", "jun"],
    ["juillet", "juil", "jul", "july"],
    ["août", "aout", "august", "aug"],
    ["septembre", "sept", "sep", "september"],
    ["octobre", "oct", "october"],
    ["novembre", "nov", "november"],
    ["décembre", "decembre", "dec", "december"],
  ] as string[][]).forEach((variants, idx) => {
    variants.forEach((variant) => {
      const norm = normalizeHeader(variant);
      if (norm) monthNameLookup.set(norm, idx);
    });
  });

  const stripTimePortion = (text: string) => {
    let cleaned = text.trim();
    cleaned = cleaned.replace(/\s+GMT.*$/i, "");
    cleaned = cleaned.replace(/T.*$/i, "");
    cleaned = cleaned.replace(/\s+\d{1,2}:\d{2}(?::\d{2})?(\s*[ap]m)?$/i, "");
    cleaned = cleaned.replace(/\s+\d{1,2}h\d{2}(\s*[ap]m)?$/i, "");
    cleaned = cleaned.replace(/\s+à\s*\d{1,2}h\d{2}$/i, "");
    cleaned = cleaned.replace(/,+$/, "");
    return cleaned.trim();
  };

  const buildDateSuccess = (day: number, month: number, yearRaw: number, formatUsed: string): DateRepairResult | null => {
    const year = normalizeYear(yearRaw);
    const month0 = month - 1;
    const serial = toExcelSerial(year, month0, day);
    if (serial === null) return null;
    return {
      kind: "date",
      serial,
      parsedDay: day,
      parsedMonth: month,
      parsedYear: year,
      formatUsed,
      reason: formatUsed,
    };
  };

  const trySplitDate = (
    text: string,
    separator: RegExp,
    order: "dmy" | "mdy" | "ymd" | "ydm" | "heuristic",
    formatUsed: string
  ): DateRepairResult | null => {
    const parts = text
      .split(separator)
      .map((part) => part.trim())
      .filter((part) => part.length > 0);
    if (parts.length !== 3) return null;
    const numbers = parts.map((part) => Number(part));
    if (numbers.some((num) => Number.isNaN(num))) return null;
    const [first, second, third] = numbers as [number, number, number];
    let day: number;
    let month: number;
    let year: number;
    if (order === "dmy") {
      day = first;
      month = second;
      year = third;
    } else if (order === "mdy") {
      day = second;
      month = first;
      year = third;
    } else if (order === "ymd") {
      day = third;
      month = second;
      year = first;
    } else if (order === "ydm") {
      day = second;
      month = third;
      year = first;
    } else {
      const heurYear = normalizeYear(third);
      if (second > 12) {
        month = first;
        day = second;
      } else if (first > 12) {
        day = first;
        month = second;
      } else {
        day = first;
        month = second;
      }
      return buildDateSuccess(day, month, heurYear, formatUsed);
    }
    return buildDateSuccess(day, month, year, formatUsed);
  };

  const tryMonthName = (text: string): DateRepairResult | null => {
    const cleaned = text
      .replace(/,/g, " ")
      .replace(/-/g, " ")
      .replace(/\./g, " ")
      .trim();
    const tokens = cleaned
      .split(/\s+/)
      .map((token) => token.replace(/(st|nd|rd|th)$/i, "").replace(/[.,]$/, ""))
      .filter((token) => token.length > 0);
    if (tokens.length < 3) return null;
    const tryOrder = (dayIdx: number, monthIdx: number, yearIdx: number, formatUsed: string) => {
      const dayVal = Number(tokens[dayIdx]);
      const yearVal = Number(tokens[yearIdx]);
      const monthNorm = normalizeHeader(tokens[monthIdx]);
      if (!Number.isFinite(dayVal) || !Number.isFinite(yearVal) || !monthNorm) return null;
      const mapped = monthNameLookup.get(monthNorm);
      if (typeof mapped !== "number") return null;
      return buildDateSuccess(dayVal, mapped + 1, yearVal, formatUsed);
    };
    const primary = tryOrder(0, 1, 2, "monthname_dmy");
    if (primary) return primary;
    return tryOrder(1, 0, 2, "monthname_mdy");
  };

  const tryDigitsSerial = (text: string): DateRepairResult | null => {
    const digitsOnly = text.replace(/\D/g, "");
    if (!digitsOnly) return null;
    if (/^\d{8}$/.test(digitsOnly)) {
      const year = Number(digitsOnly.slice(0, 4));
      const month = Number(digitsOnly.slice(4, 6));
      const day = Number(digitsOnly.slice(6));
      const result = buildDateSuccess(day, month, year, "compact_yyyymmdd");
      if (result) return result;
    }
    if (/^\d+$/.test(digitsOnly)) {
      return numberAsSerial(Number(digitsOnly));
    }
    return null;
  };

  const parseNumberRepairFR = (value: any): NumberRepairResult => {
    if (typeof value === "number" && Number.isFinite(value)) {
      return { kind: "num", num: value, formatUsed: "number_native", reason: "number_native" };
    }
    if (value === null || typeof value === "undefined") {
      return { kind: "unfixable", reason: "missing" };
    }
    const raw = `${value}`.trim();
    if (!raw) return { kind: "unfixable", reason: "empty" };
    let text = raw.replace(/\u00A0/g, " ").trim();
    let negative = false;
    if (text.startsWith("(") && text.endsWith(")")) {
      negative = true;
      text = text.slice(1, -1).trim();
    }
    if (!text) return { kind: "unfixable", reason: "empty_after_parenthesis" };
    if (/^[+-]/.test(text)) {
      if (text.startsWith("-")) negative = true;
      text = text.slice(1).trim();
    }
    if (!text) return { kind: "unfixable", reason: "empty_after_sign" };
    if (/[A-Za-z]/.test(text)) return { kind: "unfixable", reason: "contains_letters" };
    if (/[.,]$/.test(text)) return { kind: "unfixable", reason: "trailing_separator" };
    const sanitized = text.replace(/[\s']/g, "");
    const commaCount = (sanitized.match(/,/g) || []).length;
    const dotCount = (sanitized.match(/\./g) || []).length;
    let decimalSeparator: "." | "," | null = null;
    if (commaCount > 0 && dotCount > 0) {
      decimalSeparator = sanitized.lastIndexOf(",") > sanitized.lastIndexOf(".") ? "," : ".";
    } else if (commaCount > 0) {
      const after = sanitized.length - sanitized.lastIndexOf(",") - 1;
      decimalSeparator = after > 0 && after <= 2 ? "," : null;
    } else if (dotCount > 0) {
      const after = sanitized.length - sanitized.lastIndexOf(".") - 1;
      decimalSeparator = after > 0 && after <= 2 ? "." : null;
    }
    let normalized = sanitized;
    if (decimalSeparator) {
      const decimalIdx = normalized.lastIndexOf(decimalSeparator);
      if (decimalIdx >= 0) {
        const prefix = normalized.slice(0, decimalIdx).replace(/[.,]/g, "");
        const postfix = normalized.slice(decimalIdx + 1).replace(/[.,]/g, "");
        normalized = `${prefix}.${postfix}`;
      }
    } else {
      normalized = normalized.replace(/[.,]/g, "");
    }
    if (!normalized || /[^0-9.]/.test(normalized)) return { kind: "unfixable", reason: "normalized_invalid_chars" };
    if (normalized.split(".").length > 2) return { kind: "unfixable", reason: "multiple_decimal_points" };
    const parsed = Number(normalized);
    if (!Number.isFinite(parsed)) return { kind: "unfixable", reason: "parse_failed" };
    const finalValue = negative ? -Math.abs(parsed) : parsed;
    const formatUsed = decimalSeparator === "," ? "fr_decimal" : decimalSeparator === "." ? "us_decimal" : "integer";
    return { kind: "num", num: finalValue, formatUsed, reason: formatUsed };
  };

  const parseDateRepairFR = (value: any): DateRepairResult => {
    if (value === null || typeof value === "undefined") return { kind: "unfixable", reason: "missing" };
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      const serial = numberAsSerial(Math.round((value.getTime() - EXCEL_EPOCH_MS) / MS_PER_DAY));
      if (serial.kind === "date") {
        return { ...serial, formatUsed: "date_native", reason: "date_native" };
      }
      return serial;
    }
    if (typeof value === "number") return numberAsSerial(value);
    if (typeof value !== "string") return { kind: "unfixable", reason: "type_mismatch" };
    const trimmed = stripTimePortion(value);
    if (!trimmed) return { kind: "unfixable", reason: "empty_string" };
    const normalized = trimmed.replace(/\s+/g, " ").trim();
    const patterns: Array<{ sep: RegExp; order: "dmy" | "mdy" | "ymd" | "ydm" | "heuristic"; format: string }> = [
      { sep: /\//, order: "heuristic", format: "heuristic_slash" },
      { sep: /-/, order: "heuristic", format: "heuristic_dash" },
      { sep: /\//, order: "dmy", format: "fr_slash" },
      { sep: /\//, order: "mdy", format: "us_slash" },
      { sep: /\//, order: "ymd", format: "iso_slash" },
      { sep: /\//, order: "ydm", format: "swap_slash" },
      { sep: /-/, order: "ymd", format: "iso_dash" },
      { sep: /-/, order: "dmy", format: "fr_dash" },
      { sep: /-/, order: "ydm", format: "swap_dash" },
      { sep: /\./, order: "dmy", format: "fr_dot" },
      { sep: /\./, order: "ymd", format: "iso_dot" },
      { sep: /\./, order: "ydm", format: "swap_dot" },
      { sep: /\s+/, order: "dmy", format: "space_dmy" },
      { sep: /\s+/, order: "mdy", format: "space_mdy" },
      { sep: /\s+/, order: "ymd", format: "space_ymd" },
    ];
    for (const pattern of patterns) {
      const result = trySplitDate(normalized, pattern.sep, pattern.order, pattern.format);
      if (result) return result;
    }
    const monthResult = tryMonthName(normalized);
    if (monthResult) return monthResult;
    const digitsResult = tryDigitsSerial(normalized);
    if (digitsResult) return digitsResult;
    return { kind: "unfixable", reason: "no_pattern" };
  };

  const isParsableNumber = (value: any) => parseNumberRepairFR(value).kind === "num";
  const isParsableDate = (value: any) => parseDateRepairFR(value).kind === "date";
  type AuditDateCheckResult = { acceptable: boolean; reason: string; raw: string };
  const checkStrictDateComponents = (day: number, month: number, year: number, label: string) => {
    if (!Number.isFinite(day) || !Number.isFinite(month) || !Number.isFinite(year)) {
      return { acceptable: false, reason: `${label}_non_numeric` };
    }
    if (month < 1 || month > 12) {
      return { acceptable: false, reason: `${label}_month_out_of_range` };
    }
    if (!validateDatePartsStrict(year, month - 1, day)) {
      return { acceptable: false, reason: `${label}_invalid_date` };
    }
    return { acceptable: true, reason: `${label}_valid` };
  };
  const evaluateStrictIsoDate = (
    text: string
  ): { acceptable: boolean; reason: string } | null => {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
    if (!match) return null;
    const numbers = match.slice(1).map((part) => Number(part)) as [number, number, number];
    const [year, month, day] = numbers;
    return checkStrictDateComponents(day, month, year, "iso");
  };
  const evaluateStrictSlashDate = (
    text: string
  ): { acceptable: boolean; reason: string } | null => {
    const match = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/.exec(text);
    if (!match) return null;
    const numbers = match.slice(1).map((part) => Number(part)) as [number, number, number];
    const [first, second, third] = numbers;
    const frCheck = checkStrictDateComponents(first, second, third, "fr");
    if (frCheck.acceptable) return frCheck;
    const usCheck = checkStrictDateComponents(second, first, third, "us");
    if (usCheck.acceptable) return usCheck;
    return { acceptable: false, reason: frCheck.reason || usCheck.reason || "slash_invalid" };
  };
  const evaluateAuditDate = (value: any): AuditDateCheckResult => {
    const rawInput =
      typeof value === "string" ? value.trim() : value === null || typeof value === "undefined" ? "" : `${value}`;
    if (typeof value === "number" && Number.isFinite(value)) {
      return { acceptable: true, reason: "numeric_serial", raw: rawInput };
    }
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return { acceptable: true, reason: "date_object", raw: rawInput };
    }
    if (typeof value !== "string") {
      return { acceptable: false, reason: "type_mismatch", raw: rawInput };
    }
    if (!rawInput) {
      return { acceptable: false, reason: "empty_string", raw: rawInput };
    }
    const isoCheck = evaluateStrictIsoDate(rawInput);
    if (isoCheck) {
      return { acceptable: isoCheck.acceptable, reason: isoCheck.reason, raw: rawInput };
    }
    const slashCheck = evaluateStrictSlashDate(rawInput);
    if (slashCheck) {
      return { acceptable: slashCheck.acceptable, reason: slashCheck.reason, raw: rawInput };
    }
    return { acceptable: false, reason: "not_matching_strict_format", raw: rawInput };
  };
  const isDateAcceptableForAudit = (value: any, cached?: AuditDateCheckResult) =>
    (cached ?? evaluateAuditDate(value)).acceptable;
  const formatIssueValue = (value: any) => {
    if (value === null || typeof value === "undefined") return "";
    if (typeof value === "object") {
      try {
        return JSON.stringify(value);
      } catch {
        return `${value}`;
      }
    }
    return `${value}`;
  };
  const inferExpectedTypeFromMessage = (message?: string): "number" | "date" | null => {
    if (!message) return null;
    const lower = message.toLowerCase();
    if (lower.includes("nombre")) return "number";
    if (lower.includes("date")) return "date";
    return null;
  };
  const formatValueForIssues = (value: any, _column?: string) => {
    return formatIssueValue(value);
  };
  const internalState = params.__internal;
  const decisions = internalState?.decisions || {};
  const hasSeenAudit = Boolean(internalState?.phase);
  const detectConfig =
    params.detect && Object.keys(params.detect).length
      ? params.detect
      : { missing: true, duplicates: true, badType: true };
  const runMissing = detectConfig.missing === true;
  const runDuplicates = detectConfig.duplicates === true;
  const runBadType = detectConfig.badType === true;
  const optionsAny = (params as any)?.options ?? {};
  const maxIssues = clampMaxIssues(optionsAny.maxIssues);
  const DEBUG_TYPECAST = optionsAny.debugTypecast === true;
  const DEBUG_AUDIT_DATE = optionsAny.debugAuditDate === true;
  const shouldAskQuestion = (question: typeof VALIDATE_DATA_QUESTIONS[number], counts: IssueCounts) => {
    if (question.detectKey === "missing") return runMissing && counts.missing > 0;
    if (question.detectKey === "duplicates") return runDuplicates && counts.duplicate > 0;
    if (question.detectKey === "badType") return runBadType && counts.bad_type > 0;
    return false;
  };
  const buildConfirmation = (entry: typeof VALIDATE_DATA_QUESTIONS[number]): ConfirmationRequest => ({
    id: entry.id,
    question: entry.question,
    choices: VALIDATE_DATA_CONFIRM_CHOICES,
  });
  const getNextQuestion = (counts: IssueCounts) =>
    VALIDATE_DATA_QUESTIONS.find(
      (entry) => shouldAskQuestion(entry, counts) && typeof decisions[entry.decisionKey] === "undefined"
    );
  let dataSheetName = "";
  let dataBounds: AddressBounds | null = null;
  let dataHeaders: string[] = [];
  let dataRangeAddress: string | null = null;
  let duplicateKeyIndices: number[] = [];
  const resolveIssueTableSnapshot = async (): Promise<IssueTableSnapshot | null> => {
    const reference = await resolveIssuesSheetReference(ctx);
    if (!reference?.sheet) {
      ctx.log({ level: "warn", message: "validate_data issues_sheet_not_found" });
      return null;
    }
    const snapshot = await readIssuesTableFromReference(ctx, reference);
    if (snapshot && snapshot.headers.length > 0 && snapshot.rows.length > 0) return snapshot;
    if (reference.rangeA1) {
      const fallbackSnapshot = await readIssuesSnapshotFromRange(ctx, reference.sheet, reference.rangeA1);
      if (fallbackSnapshot) {
        ctx.log({
          level: "info",
          message: "validate_data issues_table_fallback_to_range",
          data: {
            sheet: reference.sheet,
            tableName: reference.tableName,
            rangeA1: reference.rangeA1,
          },
        });
        fallbackSnapshot.reference = {
          sheet: reference.sheet,
          tableName: reference.tableName,
          rangeA1: reference.rangeA1,
        };
        await tryRecreateIssuesTable(ctx, reference.sheet, reference.rangeA1, reference.tableName);
        return fallbackSnapshot;
      }
    }
    if (snapshot) return snapshot;
    ctx.log({
      level: "warn",
      message: "validate_data issues_table_not_readable",
      data: { sheet: reference.sheet, tableName: reference.tableName, rangeA1: reference.rangeA1 },
    });
    return null;
  };
  const readIssuesSnapshotFromRange = async (
    ctx: MacroContext,
    sheetName: string,
    rangeA1: string
  ): Promise<IssueTableSnapshot | null> => {
    if (!sheetName || !rangeA1) return null;
    let ws: any;
    try {
      ws = getWorksheet(ctx.excelCtx, sheetName);
    } catch {
      ctx.log({
        level: "warn",
        message: "validate_data issues_sheet_not_found",
        data: { sheet: sheetName },
      });
      return null;
    }
    if (!ws) return null;
    let range: any | null = null;
    try {
      range = ws.getRange(rangeA1);
    } catch {
      range = null;
    }
    if (!range) return null;
    await loadAndSync(range, ["values", "address"], ctx);
    const allValues = Array.isArray(range?.values) ? range.values : [];
    if (!allValues.length) return null;
    const headers = Array.isArray(allValues[0]) ? allValues[0] : [];
    if (!headers.length) return null;
    const typeIdx = findHeaderIndex(headers, "type");
    const rowKeyIdx = findHeaderIndex(headers, "rowkey");
    if (typeIdx < 0 || rowKeyIdx < 0) return null;
    const rows = (allValues.slice(1) as any[])
      .map((row) => (Array.isArray(row) ? row : []))
      .filter((row) =>
        row.some((cell) => {
          const text = cell === null || typeof cell === "undefined" ? "" : `${cell}`;
          return text.trim().length > 0;
        })
      );
    const fixIdx = findHeaderIndex(headers, "fix");
    const counts: IssueCounts = { missing: 0, duplicate: 0, bad_type: 0 };
    rows.forEach((row) => {
      const fixValue = fixIdx >= 0 ? normalizeFixValue(row[fixIdx]) : "apply";
      if (fixValue === "ignore") return;
      const type = `${row[typeIdx] ?? ""}`.trim().toLowerCase();
      if (type === "missing") counts.missing += 1;
      if (type === "duplicate") counts.duplicate += 1;
      if (type === "bad_type") counts.bad_type += 1;
    });
    const rangeAddress = typeof range.address === "string" ? range.address : `${sheetName}!${rangeA1}`;
    const localRange = extractLocalRange(rangeAddress) ?? rangeA1;
    return {
      sheet: sheetName,
      headers,
      rows,
      counts,
      rangeAddress,
      localRange,
      range: localRange ? parseA1Address(localRange) : null,
    };
  };
  const tryRecreateIssuesTable = async (
    ctx: MacroContext,
    sheetName: string,
    rangeA1: string,
    tableName?: string
  ) => {
    if (!sheetName || !rangeA1) return;
    let ws: any;
    try {
      ws = getWorksheet(ctx.excelCtx, sheetName);
    } catch {
      return;
    }
    if (!ws || !ws.tables?.add) return;
    try {
      const tbl = ws.tables.add(rangeA1, true);
      if (tbl && tableName) {
        tbl.name = tableName;
      }
    } catch (err: any) {
      ctx.log({
        level: "warn",
        message: "validate_data issues_table_recreate_failed",
        data: {
          sheet: sheetName,
          tableName,
          rangeA1,
          error: err?.message || err,
        },
      });
    }
  };
  const runApplyFixes = async (snapshot: IssueTableSnapshot): Promise<MacroResult> => {
    if (!snapshot.rows.length) {
      ctx.log({ level: "info", message: "validate_data apply_no_data" });
      return { status: "ok" };
    }
    const headers = Array.isArray(snapshot.headers) ? snapshot.headers : [];
    const typeIdx = findHeaderIndex(headers, "type");
    const rowKeyIdx = findHeaderIndex(headers, "rowkey");
    const columnIdx = findHeaderIndex(headers, "colonne");
    const messageIdx = findHeaderIndex(headers, "message");
    const fixIdx = findHeaderIndex(headers, "fix");
    if (typeIdx < 0 || rowKeyIdx < 0) {
      ctx.log({ level: "error", message: "validate_data apply_headers_missing" });
      return { status: "error" };
    }
    if (!dataSheetName || !dataHeaders.length || !dataBounds) {
      ctx.log({ level: "error", message: "validate_data apply_source_missing" });
      return { status: "error" };
    }
    const processedIssuesByType: Record<IssueType, Set<string>> = {
      missing: new Set(),
      duplicate: new Set(),
      bad_type: new Set(),
    };
    const buildProcessedKey = (issueType: IssueType, rowKey: string, columnName?: string) => {
      if (issueType === ISSUE_TYPE_BAD_TYPE && columnName) {
        return `${rowKey}::${normalizeHeader(columnName)}`;
      }
      return rowKey;
    };
    type ParsedIssue = {
      issueRowIdx: number;
      issueType: IssueType;
      rowKey: string;
      columnName: string;
      message: string;
    };
    const toText = (value: any) => (value === null || typeof value === "undefined" ? "" : `${value}`);
    const issues: ParsedIssue[] = [];
    const issueFixStats = { apply: 0, ignore: 0 };
    snapshot.rows.forEach((row, idx) => {
      const issueType = toText(row[typeIdx]).trim().toLowerCase() as IssueType;
      const rowKey = toText(row[rowKeyIdx]).trim();
      const fixValue = normalizeFixValue(fixIdx >= 0 ? row[fixIdx] : undefined);
      issueFixStats[fixValue] = (issueFixStats[fixValue] ?? 0) + 1;
      if (fixValue === "ignore") return;
      if (![ISSUE_TYPE_MISSING, ISSUE_TYPE_DUPLICATE, ISSUE_TYPE_BAD_TYPE].includes(issueType)) return;
      issues.push({
        issueRowIdx: idx,
        issueType,
        rowKey,
        columnName: columnIdx >= 0 ? toText(row[columnIdx]) : "",
        message: messageIdx >= 0 ? toText(row[messageIdx]) : "",
      });
    });
    ctx.log({
      level: "info",
      message: "validate_data apply_start",
      data: { counts: snapshot.counts, fixState: issueFixStats },
    });
    const extractId = (rowKey: string): string | null => {
      const trimmed = (rowKey || "").trim();
      if (!trimmed.toLowerCase().startsWith("id:")) return null;
      const id = trimmed.slice(3);
      return id ? `id:${id}` : null;
    };
    const rebuildLive = async () => {
      const live = await buildLiveRowIndex(ctx, dataSheetName, dataBounds, dataHeaders, dataRangeAddress);
      if (live.bounds) {
        dataBounds = live.bounds;
        dataRangeAddress = `${dataSheetName}!${rowColToA1(live.bounds.startRow, live.bounds.startCol)}:${rowColToA1(
          live.bounds.endRow,
          live.bounds.endCol
        )}`;
      }
      return live;
    };
    let castsApplied = 0;
    let castsSkipped = 0;
    let rowsDeletedMissing = 0;
    let rowsDeletedDuplicate = 0;

    const shouldDeleteMissing = decisions.fixMissing === true;
    if (shouldDeleteMissing) {
      const live = await rebuildLive();
      const deleteTargets = new Map<number, string>();
      issues
        .filter((issue) => issue.issueType === ISSUE_TYPE_MISSING)
        .forEach((issue) => {
          const idKey = extractId(issue.rowKey);
          if (!idKey) return;
          const refs = live.index.get(idKey);
          const ref = refs?.[0];
          if (!ref) return;
          deleteTargets.set(ref.excelRow1, idKey);
        });
      const rowsToDelete = Array.from(deleteTargets.keys()).sort((a, b) => b - a);
      if (rowsToDelete.length) {
        ctx.log({
          level: "info",
          message: "validate_data missing_rows_to_delete",
          data: { ids: Array.from(new Set(deleteTargets.values())), rows: rowsToDelete },
        });
        const ws = getWorksheet(ctx.excelCtx, dataSheetName);
        for (const row1 of rowsToDelete) {
          try {
            const range =
              (typeof ws.getRange === "function" && ws.getRange(`${row1}:${row1}`)) ||
              (typeof ws.getRangeByIndexes === "function" && ws.getRangeByIndexes(row1 - 1, 0, 1, 1));
            if (!range) continue;
            if (typeof range.delete === "function") {
            if (typeof Excel !== "undefined" && Excel.DeleteShiftDirection) {
              range.delete(Excel.DeleteShiftDirection.up);
            } else {
              range.delete();
            }
          }
          rowsDeletedMissing += 1;
          const processedId = deleteTargets.get(row1);
          if (processedId) processedIssuesByType.missing.add(buildProcessedKey(ISSUE_TYPE_MISSING, processedId));
        } catch (err: any) {
          ctx.log({
            level: "warn",
            message: "validate_data apply_deletion_failed",
            data: { row: row1, error: err?.message || err },
            });
          }
        }
        await rebuildLive();
      }
    }

    const shouldDeleteDuplicates = decisions.fixDuplicates === true;
    if (shouldDeleteDuplicates && duplicateKeyIndices.length === 0) {
      ctx.log({ level: "warn", message: "validate_data duplicates_no_key_columns" });
    } else if (shouldDeleteDuplicates) {
      const live = await rebuildLive();
      const duplicateGroups: Array<{ key: string; keep: string; delete: string[] }> = [];
      const keyFirstId = new Map<string, string>();
      const idsToDelete = new Set<string>();
      live.rows.forEach((row, idx) => {
        const rowKey = getRowKeyFromDataRow(row, dataHeaders);
        if (!rowKey) return;
        const keyParts = duplicateKeyIndices.map((colIdx) => normalizeDuplicateKey(row[colIdx]));
        const key = keyParts.join("\u0001");
        if (!key) return;
        if (!keyFirstId.has(key)) {
          keyFirstId.set(key, rowKey);
        } else {
          const keepId = keyFirstId.get(key)!;
          idsToDelete.add(rowKey);
          const group = duplicateGroups.find((g) => g.key === key);
          if (group) {
            group.delete.push(rowKey);
          } else {
            duplicateGroups.push({ key, keep: keepId, delete: [rowKey] });
          }
        }
      });
      if (idsToDelete.size) {
        const liveAfterKeys = await rebuildLive();
        const deleteTargets = new Map<number, string>();
        idsToDelete.forEach((idKey) => {
          const refs = liveAfterKeys.index.get(idKey);
          const ref = refs?.[0];
          if (!ref) return;
          deleteTargets.set(ref.excelRow1, idKey);
        });
        const rowsToDelete = Array.from(deleteTargets.keys()).sort((a, b) => b - a);
        ctx.log({
          level: "info",
          message: "validate_data duplicate_groups",
          data: {
            keepRule: "keep_first_occurrence",
            groups: duplicateGroups.slice(0, 20),
            deleteIds: Array.from(idsToDelete),
            deleteRows: rowsToDelete,
          },
        });
        const ws = getWorksheet(ctx.excelCtx, dataSheetName);
        for (const row1 of rowsToDelete) {
          try {
            const range =
              (typeof ws.getRange === "function" && ws.getRange(`${row1}:${row1}`)) ||
              (typeof ws.getRangeByIndexes === "function" && ws.getRangeByIndexes(row1 - 1, 0, 1, 1));
            if (!range) continue;
            if (typeof range.delete === "function") {
            if (typeof Excel !== "undefined" && Excel.DeleteShiftDirection) {
              range.delete(Excel.DeleteShiftDirection.up);
            } else {
              range.delete();
            }
          }
          rowsDeletedDuplicate += 1;
          const processedId = deleteTargets.get(row1);
          if (processedId) processedIssuesByType.duplicate.add(buildProcessedKey(ISSUE_TYPE_DUPLICATE, processedId));
        } catch (err: any) {
          ctx.log({
            level: "warn",
            message: "validate_data apply_deletion_failed",
            data: { row: row1, error: err?.message || err },
            });
          }
        }
        const postCheck = await rebuildLive();
        const seenKeys = new Set<string>();
        let remaining = 0;
        postCheck.rows.forEach((row) => {
          const keyParts = duplicateKeyIndices.map((colIdx) => normalizeDuplicateKey(row[colIdx]));
          const key = keyParts.join("\u0001");
          if (!key) return;
          if (seenKeys.has(key)) {
            remaining += 1;
          } else {
            seenKeys.add(key);
          }
        });
        if (remaining > 0) {
          ctx.log({
            level: "warn",
            message: "validate_data duplicates_remaining",
            data: { remaining },
          });
        }
      }
    }

    const shouldTypecast = decisions.fixBadType === true;
    if (shouldTypecast) {
      const live = await rebuildLive();
      const castTargets: Array<{
        rowKey: string;
        rowIdx: number;
        colIdx: number;
        expected: "number" | "date";
        columnName: string;
      }> = [];
      issues
        .filter((issue) => issue.issueType === ISSUE_TYPE_BAD_TYPE)
        .forEach((issue) => {
          const idKey = extractId(issue.rowKey);
          if (!idKey) return;
          const refs = live.index.get(idKey);
          const ref = refs?.[0];
          if (!ref) return;
          const colIdx = findHeaderIndex(dataHeaders, issue.columnName);
          if (colIdx < 0) return;
          const expected = inferExpectedTypeFromMessage(issue.message);
          if (!expected) return;
          castTargets.push({ rowKey: idKey, rowIdx: ref.rowIdx, colIdx, expected, columnName: issue.columnName });
        });
      const ws = getWorksheet(ctx.excelCtx, dataSheetName);
      const columnExpected = new Map<number, "number" | "date">();
      castTargets.forEach((target) => {
        if (!columnExpected.has(target.colIdx)) columnExpected.set(target.colIdx, target.expected);
      });
      const columnProfiles = new Map<number, ColumnStats>();
      const sampleLimit = 200;
      const sampledCols = Array.from(new Set(castTargets.map((target) => target.colIdx)));
      sampledCols.forEach((colIdx) => {
        let countNonEmpty = 0;
        let countConvertible = 0;
        let countTextLike = 0;
        let sampled = 0;
        for (let rowIdx = 0; rowIdx < tableRows.length && sampled < sampleLimit; rowIdx += 1) {
          const row = tableRows[rowIdx] || [];
          const value = row[colIdx];
          if (isMissingValue(value)) continue;
          countNonEmpty += 1;
          sampled += 1;
          const textValue = typeof value === "string" ? value.trim() : "";
          const hasLetters = typeof value === "string" && /[A-Za-zÀ-ÿ]/.test(textValue);
          const unusualString =
            typeof value === "string" && textValue.length > 0 && !/^[\d.,\s'()-]+$/.test(textValue);
          if (hasLetters || unusualString) countTextLike += 1;
          const expectedType = columnExpected.get(colIdx);
          if (expectedType === "number" && parseNumberRepairFR(value).kind === "num") {
            countConvertible += 1;
          } else if (expectedType === "date" && parseDateRepairFR(value).kind === "date") {
            countConvertible += 1;
          }
        }
        const denominator = Math.max(1, countNonEmpty);
        const classification: ColumnClassification =
          countTextLike / denominator >= 0.6
            ? "TEXT_MAJORITY"
            : countConvertible / denominator >= 0.6
              ? "NUM_OR_DATE_MAJORITY"
              : "MIXED";
        columnProfiles.set(colIdx, {
          classification,
          countNonEmpty,
          countConvertible,
          countTextLike,
        });
      });
      const logTypecastDebug = (stage: string, data: Record<string, any>) => {
        if (!DEBUG_TYPECAST) return;
        ctx.log({
          level: "info",
          message: `validate_data typecast_debug_${stage}`,
          data,
        });
      };
      const readCellState = (range: any) => ({
        valueAfter:
          Array.isArray(range?.values) && range.values.length ? range.values[0]?.[0] : undefined,
        valueTypeAfter:
          Array.isArray(range?.valueTypes) && range.valueTypes.length ? range.valueTypes[0]?.[0] : undefined,
        numberFormatAfter:
          Array.isArray(range?.numberFormat) && range.numberFormat.length ? range.numberFormat[0]?.[0] : undefined,
      });
      const seenCells = new Set<string>();
      const shouldPropagate = DEBUG_TYPECAST === true;
      for (const target of castTargets) {
        const sheetRowIdx = target.rowIdx;
        const sheetColIdx = (dataBounds?.startCol ?? 0) + target.colIdx;
        const key = `${sheetRowIdx}:${sheetColIdx}`;
        if (seenCells.has(key)) continue;
        seenCells.add(key);
        let cellRange: any = null;
        try {
          cellRange =
            (typeof ws.getRangeByIndexes === "function" && ws.getRangeByIndexes(sheetRowIdx, sheetColIdx, 1, 1)) ||
            (typeof ws.getRange === "function" && ws.getRange(rowColToA1(sheetRowIdx, sheetColIdx)));
        } catch {
          cellRange = null;
        }
        if (!cellRange) {
          castsSkipped += 1;
          continue;
        }
        await loadAndSync(cellRange, ["address", "values", "valueTypes", "numberFormat"], ctx);
        const rawValue =
          Array.isArray(cellRange.values) && cellRange.values.length ? cellRange.values[0]?.[0] : undefined;
        const rawValueType =
          Array.isArray(cellRange.valueTypes) && cellRange.valueTypes.length ? cellRange.valueTypes[0]?.[0] : undefined;
        const numberFormatBefore =
          Array.isArray(cellRange.numberFormat) && cellRange.numberFormat.length
            ? cellRange.numberFormat[0]?.[0]
            : undefined;
        const rawText = rawValue == null ? "" : `${rawValue}`;
        const addressComputed = rowColToA1(sheetRowIdx, sheetColIdx);
        const addressA1 =
          typeof cellRange.address === "string"
            ? cellRange.address
            : `${dataSheetName}!${addressComputed}`;
        const columnStats = columnProfiles.get(target.colIdx);
        const baseDebug = {
          sheetRow0: sheetRowIdx,
          sheetCol0: sheetColIdx,
          addressComputed,
          addressA1,
          columnName: target.columnName,
          rawValue,
          rawValueType,
          rawText,
          numberFormatBefore,
          expected: target.expected,
          columnStats,
        };
        if (columnStats?.classification === "TEXT_MAJORITY") {
          ctx.log({
            level: "warn",
            message: "validate_data bad_type_text_majority",
            data: { ...baseDebug, reason: "column_text_majority" },
          });
          logTypecastDebug("before_write", { ...baseDebug, action: "text_majority" });
          cellRange.numberFormat = [["@"]];
          cellRange.values = [[rawText]];
          await loadAndSync(cellRange, ["values", "valueTypes", "numberFormat"], ctx);
          castsSkipped += 1;
          const afterState = readCellState(cellRange);
          logTypecastDebug("after_write", {
            ...baseDebug,
            repairResult: { kind: "unfixable", reason: "column_text_majority" },
            valueWritten: rawText,
            numberFormatWritten: ["@"],
            ...afterState,
          });
          continue;
        }
        const digitsOnly = rawText.replace(/\s+/g, "");
        const isShortDateAmbiguous =
          columnStats?.classification === "MIXED" &&
          target.expected === "date" &&
          /^[0-9]{1,4}$/.test(digitsOnly);
        if (isShortDateAmbiguous) {
          ctx.log({
            level: "warn",
            message: "validate_data bad_type_date_ambiguous_short",
            data: { ...baseDebug, reason: "short_numeric_mixed" },
          });
          logTypecastDebug("before_write", { ...baseDebug, action: "ambiguous_short" });
          cellRange.numberFormat = [["@"]];
          cellRange.values = [[rawText]];
          await loadAndSync(cellRange, ["values", "valueTypes", "numberFormat"], ctx);
          castsSkipped += 1;
          const afterState = readCellState(cellRange);
          logTypecastDebug("after_write", {
            ...baseDebug,
            repairResult: { kind: "unfixable", reason: "ambiguous_short" },
            valueWritten: rawText,
            numberFormatWritten: ["@"],
            ...afterState,
          });
          continue;
        }
        const repairResult =
          target.expected === "number" ? parseNumberRepairFR(rawValue) : parseDateRepairFR(rawValue);
        logTypecastDebug("before_write", { ...baseDebug, repairResult });
        if (repairResult.kind === "unfixable") {
          ctx.log({
            level: "warn",
            message: `validate_data bad_type_${target.expected}_unfixable`,
            data: { ...baseDebug, reason: repairResult.reason },
          });
          cellRange.numberFormat = [["@"]];
          cellRange.values = [[rawText]];
          await loadAndSync(cellRange, ["values", "valueTypes", "numberFormat"], ctx);
          castsSkipped += 1;
          const afterState = readCellState(cellRange);
          logTypecastDebug("after_write", {
            ...baseDebug,
            repairResult,
            valueWritten: rawText,
            numberFormatWritten: ["@"],
            ...afterState,
          });
          continue;
        }
        if (target.expected === "number" && repairResult.kind !== "num") {
          castsSkipped += 1;
          continue;
        }
        if (target.expected === "date" && repairResult.kind !== "date") {
          castsSkipped += 1;
          continue;
        }
        const finalFormat = target.expected === "number" ? "0,00" : "dd/mm/yyyy";
        const valueWritten = repairResult.kind === "num" ? repairResult.num : repairResult.serial;
        cellRange.numberFormat = [["General"]];
        cellRange.values = [[valueWritten]];
        cellRange.numberFormat = [[finalFormat]];
        await loadAndSync(cellRange, ["values", "valueTypes", "numberFormat"], ctx);
        const afterState = readCellState(cellRange);
        logTypecastDebug("after_write", {
          ...baseDebug,
          repairResult,
          valueWritten,
          numberFormatWritten: ["General", finalFormat],
          ...afterState,
        });
        const { valueAfter, valueTypeAfter } = afterState;
        const accepted =
          typeof valueAfter === "number" &&
          Number.isFinite(valueAfter) &&
          valueTypeAfter !== "String";
        if (!accepted) {
          ctx.log({
            level: "warn",
            message: `validate_data bad_type_${target.expected}_fixed_but_excel_rejected`,
            data: {
              ...baseDebug,
              valueAfter,
              valueTypeAfter,
              numberFormatAfter: afterState.numberFormatAfter,
            },
          });
          cellRange.numberFormat = [["@"]];
          cellRange.values = [[rawText]];
          await loadAndSync(cellRange, ["values", "valueTypes", "numberFormat"], ctx);
          castsSkipped += 1;
          const rollbackState = readCellState(cellRange);
          logTypecastDebug("after_write", {
            ...baseDebug,
            repairResult,
            valueWritten: rawText,
            numberFormatWritten: ["@"],
            ...rollbackState,
          });
          continue;
        }
        castsApplied += 1;
        const processedKey = buildProcessedKey(ISSUE_TYPE_BAD_TYPE, target.rowKey, target.columnName);
        processedIssuesByType.bad_type.add(processedKey);
        if (shouldPropagate) {
          propagateValueToContainingRanges(ws, sheetRowIdx, sheetColIdx, valueAfter);
          await loadAndSync(cellRange, ["values", "valueTypes", "numberFormat"], ctx);
          const propagateState = readCellState(cellRange);
          logTypecastDebug("after_propagate", {
            ...baseDebug,
            repairResult,
            valueWritten,
            numberFormatWritten: ["General", finalFormat],
            ...propagateState,
            propagated: true,
          });
        }
      }
      if (castTargets.length) {
        ctx.log({
          level: "info",
          message: "validate_data bad_type_cast_targets",
          data: {
            count: castTargets.length,
            ids: Array.from(new Set(castTargets.map((c) => c.rowKey))),
            columns: Array.from(new Set(castTargets.map((c) => c.columnName))),
          },
        });
      }
    }

    if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
    ctx.log({
      level: "info",
      message: "validate_data apply_summary",
      data: {
        casts_applied: castsApplied,
        casts_skipped: castsSkipped,
        rows_deleted_missing: rowsDeletedMissing,
        rows_deleted_duplicate: rowsDeletedDuplicate,
      },
    });
    await markIssuesRowsIgnored(ctx, snapshot, processedIssuesByType, fixIdx);
    await clearIssueHighlights(
      ctx,
      snapshot,
      processedIssuesByType,
      dataHeaders,
      dataBounds,
      dataSheetName,
      buildLiveRowIndex,
      dataRangeAddress
    );
    return { status: "ok" };
  };
  const sourceRef: BlockArtifactRef = {
    blockRef: params.source?.blockRef,
    artifactRef: params.source?.artifactRef,
  };
  const resolved = resolveBlockRefOrArtifact(sourceRef, ctx, { allowDefault: true });
  if (!resolved.ok) {
    ctx.log({ level: "error", message: `validate_data source introuvable (${resolved.reason})` });
    return { status: "error" };
  }
  const blockRef = resolved.blockRef;
  const data = await readBlockData(blockRef, ctx);
  if (!data.bounds) {
    ctx.log({ level: "error", message: "validate_data: impossible de lire la source" });
    return { status: "error" };
  }
  dataSheetName = data.sheetName || blockRef.split("!")[0] || ctx.context.active?.sheetName || "Sheet1";
  let dataBoundsLocal = data.bounds;
  let tableRows = Array.isArray(data.rows) ? data.rows : [];
  let headers = Array.isArray(data.headers) ? data.headers : [];
  const idEnsured = await ensureIdColumn(ctx, dataSheetName, dataBoundsLocal, headers, tableRows);
  headers = idEnsured.headers;
  tableRows = idEnsured.rows;
  dataBoundsLocal = idEnsured.bounds;
  dataBounds = dataBoundsLocal;
  dataHeaders = headers;
  dataRangeAddress =
    dataBoundsLocal && dataSheetName
      ? `${dataSheetName}!${rowColToA1(dataBoundsLocal.startRow, dataBoundsLocal.startCol)}:${rowColToA1(
          dataBoundsLocal.endRow,
          dataBoundsLocal.endCol
        )}`
      : null;
  const headerNormMap = new Map<string, number>();
  const headerIndexByName = new Map<string, number>();
  headers.forEach((header, idx) => {
    const norm = normalizeHeader(header);
    if (norm && !headerNormMap.has(norm)) headerNormMap.set(norm, idx);
    if (header && !headerIndexByName.has(header)) headerIndexByName.set(header, idx);
  });
  const normalizedHeaders = headers.map((header) => normalizeHeader(header) || "");
  const findNormalizedHeaderIndex = (predicate: (norm: string) => boolean) => {
    return normalizedHeaders.findIndex((norm) => norm && predicate(norm));
  };
  let maxRowLength = headers.length;
  tableRows.forEach((row) => {
    if (Array.isArray(row)) maxRowLength = Math.max(maxRowLength, row.length);
  });
  const columnCount = Math.max(maxRowLength, 1);
  const resolveColumnIndices = (requested?: string[]) => {
    if (!Array.isArray(requested) || !requested.length) return [];
    const normalized = requested
      .map((col) => {
        if (typeof col !== "string") return null;
        const direct = headerIndexByName.get(col);
        if (typeof direct === "number") return direct;
        const norm = normalizeHeader(col);
        return norm ? headerNormMap.get(norm) : undefined;
      })
      .filter((idx): idx is number => typeof idx === "number")
      .filter((idx, pos, arr) => arr.indexOf(idx) === pos)
      .filter((idx) => idx >= 0 && idx < columnCount);
    return normalized;
  };
  const duplicateCandidateIndices = resolveColumnIndices(params.duplicateKeyColumns);
  const usesDuplicateKeys = runDuplicates && duplicateCandidateIndices.length > 0;
  const fallbackDuplicateKeyColumns = () => {
    const emailIdx = findNormalizedHeaderIndex((norm) => norm.includes("email"));
    if (emailIdx >= 0) return [emailIdx];
    const dateIdx = findNormalizedHeaderIndex((norm) => norm.includes("date"));
    const montantIdx = findNormalizedHeaderIndex((norm) => norm.includes("montant"));
    const clientIdxCandidate = findNormalizedHeaderIndex((norm) => norm.includes("client"));
    const clientIdx =
      clientIdxCandidate >= 0 ? clientIdxCandidate : findNormalizedHeaderIndex((norm) => norm.includes("nom"));
    if (dateIdx >= 0 && montantIdx >= 0 && clientIdx >= 0) {
      return [dateIdx, montantIdx, clientIdx];
    }
    const exceptId = normalizedHeaders
      .map((norm, idx) => ({ idx, norm }))
      .filter((entry) => entry.norm !== "id")
      .map((entry) => entry.idx);
    if (exceptId.length) return exceptId;
    return Array.from({ length: columnCount }, (_v, idx) => idx);
  };
  duplicateKeyIndices = usesDuplicateKeys ? duplicateCandidateIndices : fallbackDuplicateKeyColumns();
  const duplicateKeyNames = duplicateKeyIndices.map((idx) => headers[idx] || `Column${idx + 1}`);
  const duplicateKeyLabel = duplicateKeyNames.join("+");
  const dataStartRow = typeof dataBoundsLocal?.startRow === "number" ? dataBoundsLocal.startRow : 0;
  const dataStartCol = typeof dataBoundsLocal?.startCol === "number" ? dataBoundsLocal.startCol : 0;
  const dataFirstRow = dataStartRow + 1;
  if (hasSeenAudit) {
    const snapshot = await resolveIssueTableSnapshot();
    if (!snapshot) {
      ctx.log({ level: "error", message: "validate_data apply_missing_reference" });
      return { status: "error" };
    }
    if (Object.keys(decisions).length) {
      ctx.log({
        level: "info",
        message: "validate_data confirmation_answered",
        data: { decisions },
      });
    }
    const applyResult = await runApplyFixes(snapshot);
    if (applyResult.status === "error") return applyResult;
    const afterSnapshot = await resolveIssueTableSnapshot();
    if (!afterSnapshot) {
      ctx.log({ level: "error", message: "validate_data apply_missing_reference" });
      return { status: "error" };
    }
    const nextQuestion = getNextQuestion(afterSnapshot.counts);
    if (nextQuestion) {
      ctx.log({
        level: "warn",
        message: "validate_data confirmation_requested",
        data: { id: nextQuestion.id },
      });
      return { requiresConfirmation: buildConfirmation(nextQuestion) };
    }
    return applyResult;
  }
  const missingCandidates = resolveColumnIndices(params.missingColumns);
  const missingIndices =
    runMissing && missingCandidates.length
      ? missingCandidates
      : runMissing
        ? Array.from({ length: columnCount }, (_v, idx) => idx)
        : [];
  const typeRules = Array.isArray(params.typeRules) ? params.typeRules : [];
  const defaultTypeRules: { col: string; type: "number" | "date" }[] = [];
  const addDefaultRule = (keyword: string, type: "number" | "date") => {
    const idx = findNormalizedHeaderIndex((norm) => norm.includes(keyword));
    if (idx >= 0) {
      const name = headers[idx] || `Column${idx + 1}`;
      if (!defaultTypeRules.some((rule) => normalizeHeader(rule.col) === normalizeHeader(name) && rule.type === type)) {
        defaultTypeRules.push({ col: name, type });
      }
    }
  };
  if (!typeRules.length) {
    addDefaultRule("date", "date");
    addDefaultRule("montant", "number");
  }
  const effectiveTypeRules = typeRules.length ? typeRules : defaultTypeRules;
  const ruleMap = new Map<number, "number" | "date">();
  effectiveTypeRules.forEach((rule) => {
    if (!rule || typeof rule.col !== "string") return;
    const norm = normalizeHeader(rule.col);
    const headerIdx = headerIndexByName.get(rule.col) ?? (norm ? headerNormMap.get(norm) : undefined);
    if (headerIdx === undefined) return;
    if (rule.type === "number" || rule.type === "date") {
      ruleMap.set(headerIdx, rule.type);
    }
  });
  const heuristicsTargets: Array<{ idx: number; type: "number" | "date" }> = [];
  if (runBadType && ruleMap.size === 0) {
    for (let colIdx = 0; colIdx < columnCount; colIdx += 1) {
      let nonEmpty = 0;
      let numeric = 0;
      let dated = 0;
      for (let rowIdx = 0; rowIdx < tableRows.length; rowIdx += 1) {
        const row = tableRows[rowIdx] || [];
        const value = row[colIdx];
        if (isMissingValue(value)) continue;
        nonEmpty += 1;
        if (isParsableNumber(value)) numeric += 1;
        if (isParsableDate(value)) dated += 1;
      }
      if (nonEmpty === 0) continue;
      if (numeric / nonEmpty >= 0.8) {
        heuristicsTargets.push({ idx: colIdx, type: "number" });
      } else if (dated / nonEmpty >= 0.8) {
        heuristicsTargets.push({ idx: colIdx, type: "date" });
      }
    }
  }
  const badTypeTargets = new Map<number, "number" | "date">();
  if (runBadType) {
    if (ruleMap.size) {
      ruleMap.forEach((type, idx) => badTypeTargets.set(idx, type));
    } else {
      heuristicsTargets.forEach((entry) => badTypeTargets.set(entry.idx, entry.type));
    }
  }
  const rulesUsedForLog = effectiveTypeRules.map((rule) => `${rule.col}:${rule.type}`);
  type IssueEntry = {
    type: IssueType;
    rowIdx: number;
    colIdx: number;
    rowKey: string;
    column: string;
    value: any;
    message: string;
    highlightCols: number[];
    severity: "error";
    fix: "apply" | "ignore";
  };
  const highlightMap = new Map<string, { row: number; col: number; color: string; priority: number; issueType: IssueType }>();
  const recordHighlight = (rowIdx: number, colIdx: number, issueType: IssueType) => {
    const absRow = dataFirstRow + rowIdx;
    const absCol = dataStartCol + colIdx;
    const key = `${dataSheetName}:${absRow}:${absCol}`;
    const priority = highlightPriorityByType[issueType];
    const existing = highlightMap.get(key);
    if (existing && existing.priority >= priority) return;
    highlightMap.set(key, {
      row: absRow,
      col: absCol,
      color: highlightColorByType[issueType],
      priority,
      issueType,
    });
  };
  const issues: IssueEntry[] = [];
  let truncated = false;
  const addIssue = (issue: IssueEntry) => {
    if (issues.length >= maxIssues) {
      truncated = true;
      return false;
    }
    issues.push(issue);
    issue.highlightCols.forEach((colIdx) => {
      if (colIdx >= 0 && colIdx < columnCount) {
        recordHighlight(issue.rowIdx, colIdx, issue.type);
      }
    });
    return true;
  };
  if (runMissing && missingIndices.length) {
    for (let rowIdx = 0; rowIdx < tableRows.length; rowIdx += 1) {
      if (truncated) break;
        const row = tableRows[rowIdx] || [];
        const rowKey = getRowKeyFromDataRow(row, headers);
        for (const colIdx of missingIndices) {
        if (truncated) break;
        if (colIdx < 0 || colIdx >= columnCount) continue;
        const rawValue = row[colIdx];
        if (!isMissingValue(rawValue)) continue;
        const headerName = headers[colIdx] || `Column${colIdx + 1}`;
          if (
            !addIssue({
              type: "missing",
              rowIdx,
              colIdx,
              rowKey,
              column: headerName,
              value: "",
              message: "Cellule vide",
              highlightCols: [colIdx],
              severity: "error",
              fix: "apply",
            })
          ) {
          break;
        }
      }
    }
  }
  if (runDuplicates && duplicateKeyIndices.length && !truncated) {
    const seenKeys = new Map<string, number>();
    for (let rowIdx = 0; rowIdx < tableRows.length; rowIdx += 1) {
      if (truncated) break;
      const row = tableRows[rowIdx] || [];
      const rowKey = getRowKeyFromDataRow(row, headers);
      const normalizedParts = duplicateKeyIndices.map((colIdx) => normalizeDuplicateKey(row[colIdx]));
      const key = normalizedParts.join("\u0001");
      if (seenKeys.has(key)) {
        const issueCol = duplicateKeyIndices[0];
        if (typeof issueCol !== "number") continue;
        const columnName = headers[issueCol] || `Column${issueCol + 1}`;
        const message = `Doublon sur clé: ${duplicateKeyLabel || duplicateKeyIndices.map((idx) => headers[idx] || `Column${idx + 1}`).join("+")}`;
        const valueParts = duplicateKeyIndices.map((colIdx) => formatIssueValue(row[colIdx]));
        if (
          !addIssue({
            type: "duplicate",
            rowIdx,
            colIdx: issueCol,
            rowKey,
            column: columnName,
            value: valueParts.join(" | "),
            message,
            highlightCols: [...duplicateKeyIndices],
            severity: "error",
            fix: "apply",
          })
        ) {
          break;
        }
      } else {
        seenKeys.set(key, rowIdx);
      }
    }
  }
  if (runBadType && badTypeTargets.size && !truncated) {
    for (const [colIdx, expectedType] of badTypeTargets) {
      if (truncated) break;
      if (colIdx < 0 || colIdx >= columnCount) continue;
      for (let rowIdx = 0; rowIdx < tableRows.length; rowIdx += 1) {
        if (truncated) break;
        const row = tableRows[rowIdx] || [];
        const rowKey = getRowKeyFromDataRow(row, headers);
        const value = row[colIdx];
        if (isMissingValue(value)) continue;
        const auditStatus = expectedType === "date" ? evaluateAuditDate(value) : null;
        const valid =
          expectedType === "number"
            ? typeof value === "number" && Number.isFinite(value)
            : (typeof value === "number" && Number.isFinite(value)) ||
              (value instanceof Date && !Number.isNaN(value.getTime()));
        if (valid) continue;
        const repairResult = expectedType === "date" ? parseDateRepairFR(value) : null;
        const columnName = headers[colIdx] || `Column${colIdx + 1}`;
        if (
          expectedType === "date" &&
          DEBUG_AUDIT_DATE &&
          auditStatus &&
          !auditStatus.acceptable &&
          repairResult?.kind === "date"
        ) {
          ctx.log({
            level: "info",
            message: "validate_data audit_date_repairable_but_flagged",
            data: {
              raw: auditStatus.raw,
              reason: auditStatus.reason,
              rowKey,
              columnName,
            },
          });
        }
        const message = expectedType === "number" ? "Type invalide: attendu nombre" : "Type invalide: attendu date";
        if (
          !addIssue({
            type: "bad_type",
            rowIdx,
            colIdx,
            rowKey,
            column: columnName,
            value,
            message,
            highlightCols: [colIdx],
            severity: "error",
            fix: "apply",
          })
        ) {
          break;
        }
      }
    }
  }
  const uniqueIssues = Array.from(new Map(issues.map((issue) => [`${issue.type}|${issue.rowIdx}|${issue.colIdx}`, issue])).values());
  const sortedIssues = uniqueIssues.sort((a, b) => {
    const orderDiff = issueTypeOrder[a.type] - issueTypeOrder[b.type];
    if (orderDiff !== 0) return orderDiff;
    if (a.rowIdx !== b.rowIdx) return a.rowIdx - b.rowIdx;
    return a.colIdx - b.colIdx;
  });
  const issueCounts = { missing: 0, duplicate: 0, bad_type: 0 };
  sortedIssues.forEach((issue) => {
    issueCounts[issue.type] += 1;
  });
  if (
    typeof issueCounts.missing !== "number" ||
    typeof issueCounts.duplicate !== "number" ||
    typeof issueCounts.bad_type !== "number"
  ) {
    throw new Error("validate_data: issueCounts missing/duplicate/bad_type not computed");
  }
  if (issueCounts.duplicate) {
    ctx.log({
      level: "info",
      message: "validate_data duplicates_found",
      data: { count: issueCounts.duplicate, keyCols: duplicateKeyNames },
    });
  }
  if (issueCounts.bad_type) {
    ctx.log({
      level: "info",
      message: "validate_data bad_type_found",
      data: { count: issueCounts.bad_type, rulesUsed: rulesUsedForLog },
    });
  }
  const issueHeader = ["Type", "RowKey", "Cellule", "Colonne", "Ligne", "Valeur", "Message", "Severity", "Fix"];
  const issueRows = sortedIssues.map((issue) => {
    const absoluteRow = dataFirstRow + issue.rowIdx;
    const absoluteCol = dataStartCol + issue.colIdx;
    const cellAddress = makeRangeAddress(dataSheetName, absoluteRow, absoluteCol, 1, 1);
    return [
      issue.type,
      issue.rowKey,
      cellAddress,
      issue.column,
      issue.rowIdx + 1,
      formatValueForIssues(issue.value, issue.column),
      issue.message,
      issue.severity,
      issue.fix,
    ];
  });
  const applyIssueTypeColumnColors = async () => {
    if (!issueRows.length) return 0;
    if (!parsedRange || typeof parsedRange.startRow !== "number" || typeof parsedRange.startCol !== "number") {
      ctx.log({
        level: "warn",
        message: "validate_data type_column_fill_failed",
        data: { reason: "parse_range", target: `${targetSheetName}!${localRange}` },
      });
      return 0;
    }
    const startTypeRow = parsedRange.startRow + 1;
    const typeColIdx = parsedRange.startCol;
    let appliedCount = 0;
    let rangeIssue = false;
    for (let idx = 0; idx < issueRows.length; idx += 1) {
      const rawType = issueRows[idx]?.[0];
      const typeKey = `${rawType ?? ""}` as IssueType;
      const fillColor = highlightColorByType[typeKey];
      if (!fillColor) continue;
      const cellRow = startTypeRow + idx;
      const cellRef = rowColToA1(cellRow, typeColIdx);
      const cellRange =
        (typeof targetSheet.getRange === "function" && targetSheet.getRange(cellRef)) ||
        (typeof targetSheet.getRangeByIndexes === "function" &&
          targetSheet.getRangeByIndexes(cellRow, typeColIdx, 1, 1));
      if (!cellRange || !cellRange.format || !cellRange.format.fill) {
        rangeIssue = true;
        continue;
      }
      cellRange.format.fill.color = fillColor;
      appliedCount += 1;
    }
    if (rangeIssue) {
      ctx.log({
        level: "warn",
        message: "validate_data type_column_fill_failed",
        data: { reason: "range_access", target: `${targetSheetName}!${localRange}` },
      });
    }
    if (appliedCount) {
      ctx.log({
        level: "info",
        message: "validate_data type_column_fill_applied",
        data: { count: appliedCount },
      });
    }
    return appliedCount;
  };
  const valuesToWrite = [issueHeader, ...issueRows];
  const rowCount = Math.max(1, valuesToWrite.length);
  const colCount = issueHeader.length;
  const applyHighlights = async (columnOffset = 0) => {
    const entries = Array.from(highlightMap.values()).sort((a, b) => (a.row - b.row) || (a.col - b.col));
    if (!entries.length) return { total: 0, counts: { missing: 0, bad_type: 0, duplicate: 0 } };
    let sheetForHighlight: any | null = null;
    try {
      sheetForHighlight = getWorksheet(ctx.excelCtx, dataSheetName);
    } catch (err: any) {
      ctx.log({
        level: "warn",
        message: "validate_data highlight_sheet_not_found",
        data: { sheet: dataSheetName, error: err?.message || err },
      });
      return { total: 0, counts: { missing: 0, bad_type: 0, duplicate: 0 } };
    }
    const counters = { missing: 0, bad_type: 0, duplicate: 0 };
    let highlightedCount = 0;
    for (const entry of entries) {
      const localAddr = rowColToA1(entry.row, entry.col + columnOffset);
      const cellRange =
        (typeof sheetForHighlight.getRange === "function" && sheetForHighlight.getRange(localAddr)) ||
        sheetForHighlight.getRangeByIndexes?.(entry.row, entry.col + columnOffset, 1, 1);
      try {
        if (cellRange?.format?.fill) {
          cellRange.format.fill.color = entry.color;
          highlightedCount += 1;
          counters[entry.issueType] += 1;
        }
      } catch (err: any) {
        ctx.log({
          level: "warn",
          message: "validate_data highlight_failed",
          data: { cell: `${dataSheetName}!${localAddr}`, error: err?.message || err },
        });
      }
    }
    if (highlightedCount) {
      ctx.log({
        level: "info",
        message: "validate_data highlight_applied",
        data: { missing: counters.missing, bad_type: counters.bad_type, duplicate: counters.duplicate },
      });
      if (counters.duplicate) {
        ctx.log({
          level: "info",
          message: "validate_data duplicates_highlight_applied",
          data: { countCells: counters.duplicate },
        });
      }
      if (counters.bad_type) {
        ctx.log({
          level: "info",
          message: "validate_data bad_type_highlight_applied",
          data: { countCells: counters.bad_type },
        });
      }
    }
    return { total: highlightedCount, counts: counters };
  };
  const highlightResult = await applyHighlights();
  const placement = await resolvePlacement(
    {
      mode: "newSheet",
      minBlankArea: { rows: rowCount, cols: colCount },
      avoidOverwrite: true,
      newSheetNameHint: "Issues",
    },
    ctx,
    `${ctx.step.id}:validate_data_place`
  );
  if (placement.requiresConfirmation) return { requiresConfirmation: placement.requiresConfirmation };
  if (!placement.address) return placement;
  const [sheetPart, rangePart] = placement.address.split("!");
  const targetSheetName = sheetPart || ctx.context.active?.sheetName || "Sheet1";
  const localRange = (rangePart || placement.address)!;
  const targetSheet = placement.worksheet || getWorksheet(ctx.excelCtx, targetSheetName);
  const parsedRange = parseA1Address(localRange);
  const writeRange =
    (typeof targetSheet.getRange === "function" && targetSheet.getRange(localRange)) ||
    (parsedRange && targetSheet.getRangeByIndexes?.(parsedRange.startRow, parsedRange.startCol, rowCount, colCount));
  if (writeRange) (writeRange as any).values = valuesToWrite;
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
  let outputTableName = ensureTableName(targetSheet, "Issues");
  let tableCreated = false;
  try {
    if (targetSheet.tables?.add) {
      const tableObj = targetSheet.tables.add(localRange, true);
      tableCreated = true;
      if (tableObj) tableObj.name = outputTableName;
    }
  } catch (err: any) {
    ctx.log({
      level: "warn",
      message: `validate_data table creation failed ${err?.message || err}`,
      macro: "validate_data",
      stepId: ctx.step.id,
    });
  }
  try {
    await macroApplyFormat(
      { target: { blockRef: `${targetSheetName}!${localRange}`, tableName: outputTableName }, options: { preset: "corporate_blue", freezeHeaderRow: true } },
      { ...ctx, step: { ...ctx.step, id: `${ctx.step.id}:corporate_blue` } } as any
    );
    ctx.log({
      level: "info",
      message: `validate_data: corporate_blue format applied table=${outputTableName}`,
      macro: "validate_data",
      stepId: ctx.step.id,
    });
  } catch (err: any) {
    ctx.log({
      level: "warn",
      message: `validate_data corporate_blue format failed ${err?.message || err}`,
      macro: "validate_data",
      stepId: ctx.step.id,
    });
  }
  await applyIssueTypeColumnColors();
  const anchorCell = localRange.includes(":") ? localRange.split(":")[0]! : localRange;
  const artifacts = placement.artifacts ? [...placement.artifacts] : [];
  const finalIssues = issueRows.length;
  const artifact: ArtifactRecord = {
    type: "table",
    kind: "table",
    sheet: targetSheetName,
    sheetName: targetSheetName,
    anchor: anchorCell,
    fromStep: ctx.step.id,
    tableName: outputTableName,
    blockRef: `${targetSheetName}!${localRange}`,
    address: localRange,
    addressA1: localRange,
    headers: issueHeader,
    rowCount,
    colCount,
    rows: Math.max(0, finalIssues),
    cols: colCount,
    counts: {
      totalIssues: finalIssues,
      missing: issueCounts.missing,
      duplicate: issueCounts.duplicate,
      bad_type: issueCounts.bad_type,
    },
      details: {
        truncated,
        highlight: highlightResult.total
          ? { applied: highlightResult.total, counts: highlightResult.counts }
          : undefined,
        tableCreated,
      },
  };
  artifacts.push(artifact);
  ctx.log({
    level: "info",
    message: "validate_data issues_reported",
    data: {
      totalIssues: finalIssues,
      missing: issueCounts.missing,
      duplicate: issueCounts.duplicate,
      bad_type: issueCounts.bad_type,
      truncated,
      maxIssues,
    },
  });
  ctx.log({
    level: "info",
    message: "validate_data counts",
    data: {
      missing: issueCounts.missing,
      duplicate: issueCounts.duplicate,
      bad_type: issueCounts.bad_type,
      total: finalIssues,
      truncated,
      maxIssues,
    },
  });
  if (typeof persistIssuesSheetReference !== "function") {
    throw new Error("persistIssuesSheetReference missing: wiring regression");
  }
  if (!tableCreated) {
    ctx.log({
      level: "warn",
      message: "validate_data issues_reference_persisted_without_table",
      data: {
        sheet: targetSheetName,
        rangeA1: localRange,
        tableName: outputTableName,
      },
    });
  }
  await persistIssuesSheetReference(ctx, {
    sheet: targetSheetName,
    tableName: tableCreated ? outputTableName : undefined,
    rangeA1: localRange,
  });
  const nextQuestion = getNextQuestion(issueCounts);
  if (!nextQuestion) {
    if (issueCounts.missing === 0 && issueCounts.duplicate === 0 && issueCounts.bad_type === 0) {
      ctx.log({ level: "info", message: "validate_data no_fixes_available" });
    }
    return { artifacts };
  }
  ctx.log({ level: "warn", message: "validate_data confirmation_requested", data: { id: nextQuestion.id } });
  return {
    artifacts,
    requiresConfirmation: buildConfirmation(nextQuestion),
  };
};

type JoinKeyMapping = { leftIdx: number; rightIdx: number; strategy: JoinMatchStrategy; leftName: string; rightName: string };

function normalizeKeyValue(val: any, strategy: JoinMatchStrategy) {
  if (strategy === "numeric") {
    const asString = (val ?? "").toString();
    let cleaned = "";
    for (const ch of asString) {
      if (ch !== " " && ch !== "\t" && ch !== "\n" && ch !== "\r") cleaned += ch;
    }
    const num = typeof val === "number" ? val : Number(cleaned);
    return Number.isFinite(num) ? num.toString() : "";
  }
  if (strategy === "exact") return val === null || val === undefined ? "" : `${val}`;
  const str = val === null || val === undefined ? "" : `${val}`;
  return str.trim().toLowerCase();
}

function makeCompositeKey(row: any[], mappings: JoinKeyMapping[], side: "left" | "right") {
  return mappings
    .map((m) => {
      const idx = side === "left" ? m.leftIdx : m.rightIdx;
      const value = idx >= 0 && idx < row.length ? row[idx] : "";
      return normalizeKeyValue(value, m.strategy);
    })
    .join("\u0001");
}

async function readBlockData(blockRef: string, ctx: MacroContext) {
  const parsedRef = parseA1Address(blockRef);
  const { sheet, block } = findBlock(blockRef, ctx.context);
  const sheetName =
    sheet?.name ||
    parsedRef?.sheet ||
    extractSheetName(blockRef) ||
    ctx.context.active?.sheetName ||
    ctx.context.sheets[0]?.name ||
    "";
  const address =
    block && block.address
      ? block.address.includes("!")
        ? block.address
        : `${sheetName}!${block.address}`
      : parsedRef
      ? `${sheetName}!${extractLocalRange(blockRef) ?? blockRef}`
      : blockRef;
  const bounds = parseA1Address(address);
  if (!sheetName || !bounds) {
    return { error: "block_not_found", headers: [] as string[], rows: [] as any[][], columnTypes: [] as string[], sheetName: sheetName || "", bounds: null as any };
  }
  const ws = getWorksheet(ctx.excelCtx, sheetName);
  const localAddr = extractLocalRange(address) ?? address;
  const rowCount = Math.max(1, bounds.endRow - bounds.startRow + 1);
  const colCount = Math.max(1, bounds.endCol - bounds.startCol + 1);
  const range = ws.getRange ? ws.getRange(localAddr) : ws.getRangeByIndexes?.(bounds.startRow, bounds.startCol, rowCount, colCount);
  const { values, numberFormat, texts } = await readRangeValuesWithFormats(range, ctx.excelCtx);
  const width = Math.max(colCount, (values?.[0]?.length as number) || 0, (block?.headers?.length as number) || 0, 1);
  const headerFromSheet = Array.isArray(values?.[0]) ? values[0] : [];
  const blockHeaders = Array.isArray(block?.headers) ? block.headers : [];
  const sheetHeaders = Array.from({ length: width }, (_v, i) => {
    const h = headerFromSheet[i];
    return h === null || h === undefined || `${h}`.trim() === "" ? `Column${i + 1}` : `${h}`;
  });
  const sheetMatchesBlock =
    blockHeaders.length > 0 &&
    headerFromSheet.length === blockHeaders.length &&
    headerFromSheet.every((header, idx) => normalizeHeader(`${header ?? ""}`) === normalizeHeader(blockHeaders[idx] || ""));
  const headers = blockHeaders.length && sheetMatchesBlock ? blockHeaders : sheetHeaders;
  const padRow = (row: any[]) => {
    const arr = Array.from({ length: width }, (_v, i) => (row && typeof row[i] !== "undefined" ? row[i] : ""));
    return arr;
  };
  const rawRows = (values || []).slice(1).map((r: any[]) => padRow(Array.isArray(r) ? r : []));
  const padTextRow = (row: any[]) => {
    const arr = Array.from({ length: width }, (_v, i) => {
      if (row && typeof row[i] !== "undefined") return row[i];
      return "";
    });
    return arr;
  };
  const rawTextRows = (texts && Array.isArray(texts) ? texts : values || []).slice(1).map((r: any[]) => padTextRow(Array.isArray(r) ? r : []));
  const nfMatrix: any[][] = Array.isArray(numberFormat) ? (numberFormat as any[][]) : [];
  const columnNumberFormats: (string | null)[] = Array.from({ length: width }, () => null);
  if (nfMatrix.length > 0) {
    for (let c = 0; c < width; c += 1) {
      // prefer first data-row format (skip header row)
      for (let r = 1; r < nfMatrix.length; r += 1) {
        const rowFmt = nfMatrix[r] || [];
        const fmt = Array.isArray(rowFmt) ? rowFmt[c] : null;
        if (fmt && `${fmt}`.trim() !== "") {
          columnNumberFormats[c] = `${fmt}`;
          break;
        }
      }
      if (!columnNumberFormats[c] && Array.isArray(nfMatrix[0])) {
        const fmt = (nfMatrix[0] as any[])[c];
        if (fmt && `${fmt}`.trim() !== "") columnNumberFormats[c] = `${fmt}`;
      }
    }
  }

  const rows: any[][] = [];
  const rowsText: any[][] = [];
  rawRows.forEach((r, idx) => {
    if (!r.every(isBlankCell)) {
      rows.push(r);
      rowsText.push(rawTextRows[idx] || []);
    }
  });
  return { headers, rows, rowsText, columnTypes: block?.columnTypes || [], sheetName, bounds, numberFormats: columnNumberFormats };
}

function headerSuggestsDate(header: string) {
  const h = (header || "").toLowerCase();
  return h.includes("date") || h.includes("début") || h.includes("debut") || h.includes("fin");
}

function isLikelyDateColumn(header: string, values: any[]): boolean {
  if (headerSuggestsDate(header)) return true;
  const nums = values.filter((v) => typeof v === "number" && Number.isFinite(v));
  if (nums.length === 0) return false;
  const inRange = nums.filter((n) => n >= 30000 && n <= 60000);
  return inRange.length / nums.length >= 0.6;
}

function isAllDigits(s: string | undefined): s is string {
  if (!s) return false;
  if (s.length === 0) return false;
  for (let i = 0; i < s.length; i += 1) {
    const code = s.charCodeAt(i);
    if (code < 48 || code > 57) return false;
  }
  return true;
}

function parseDateLikeString(value: string): Date | null {
  const trimmed = (value || "").trim();
  if (trimmed.includes("/")) {
    const parts = trimmed.split("/");
    if (parts.length === 3) {
      const p0 = parts[0];
      const p1 = parts[1];
      const p2 = parts[2];
      if (isAllDigits(p0) && isAllDigits(p1) && isAllDigits(p2)) {
        const d = Number(p0);
        const m = Number(p1);
        const y = Number(p2);
        if (d >= 1 && d <= 31 && m >= 1 && m <= 12 && y >= 1900) return new Date(y, m - 1, d);
      }
    }
  }
  if (trimmed.includes("-")) {
    const parts = trimmed.split("-");
    if (parts.length === 3) {
      const p0 = parts[0];
      const p1 = parts[1];
      const p2 = parts[2];
      if (isAllDigits(p0) && isAllDigits(p1) && isAllDigits(p2)) {
        const y = Number(p0);
        const m = Number(p1);
        const d = Number(p2);
        if (d >= 1 && d <= 31 && m >= 1 && m <= 12 && y >= 1900) return new Date(y, m - 1, d);
      }
    }
  }
  return null;
}

function extractLocalRange(range?: string): string | null {
  if (!range) return null;
  const bang = range.lastIndexOf("!");
  const local = bang >= 0 ? range.slice(bang + 1) : range;
  const trimmed = local.trim();
  return trimmed || null;
}

function extractSheetName(range?: string): string | null {
  if (!range) return null;
  const bang = range.lastIndexOf("!");
  if (bang < 0) return null;
  let sheet = range.slice(0, bang).trim();
  if ((sheet.startsWith("'") && sheet.endsWith("'")) || (sheet.startsWith('"') && sheet.endsWith('"'))) {
    sheet = sheet.slice(1, -1).trim();
  }
  return sheet || null;
}

async function readIssuesTableFromReference(
  ctx: MacroContext,
  reference: IssuesSheetReference
): Promise<IssueTableSnapshot | null> {
  if (!reference?.sheet || !reference.tableName) return null;
  let sheet: any;
  try {
    sheet = getWorksheet(ctx.excelCtx, reference.sheet);
  } catch {
    ctx.log({
      level: "warn",
      message: "validate_data issues_sheet_not_found",
      data: { sheet: reference.sheet },
    });
    return null;
  }
  if (!sheet || !sheet.tables) {
    ctx.log({
      level: "warn",
      message: "validate_data issues_sheet_tables_missing",
      data: { sheet: reference.sheet },
    });
    return null;
  }
  const tables = sheet.tables;
  let table: any | null = null;
  try {
    if (typeof tables.getItemOrNullObject === "function") {
      table = tables.getItemOrNullObject(reference.tableName);
    } else if (typeof tables.getItem === "function") {
      table = tables.getItem(reference.tableName);
    }
  } catch {
    table = null;
  }
  if (table) {
    await loadAndSync(table, ["isNullObject"], ctx);
  }
  if (!table || table.isNullObject) {
    ctx.log({
      level: "warn",
      message: "validate_data issues_table_not_found",
      data: { sheet: reference.sheet, tableName: reference.tableName },
    });
    return null;
  }

  const headerRange = typeof table.getHeaderRowRange === "function" ? table.getHeaderRowRange() : null;
  const dataRange = typeof table.getDataBodyRange === "function" ? table.getDataBodyRange() : null;
  const tableRange = typeof table.getRange === "function" ? table.getRange() : null;
  if (typeof headerRange?.load === "function") headerRange.load("values");
  if (typeof dataRange?.load === "function") dataRange.load("values");
  if (typeof tableRange?.load === "function") tableRange.load(["values", "address"]);
  if (typeof table.load === "function") table.load("name");
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();

  const headers = Array.isArray(headerRange?.values) && headerRange.values.length ? headerRange.values[0] : [];
  const bodyValues =
    Array.isArray(dataRange?.values) && dataRange.values.length
      ? dataRange.values
      : Array.isArray(tableRange?.values) && tableRange.values.length > 1
        ? tableRange.values.slice(1)
        : [];
  const rows = (bodyValues as any[])
    .map((row) => (Array.isArray(row) ? row : []))
    .filter((row) =>
      row.some((cell) => {
        const text = cell === null || typeof cell === "undefined" ? "" : `${cell}`;
        return text.trim().length > 0;
      })
    );

  const fixIdx = findHeaderIndex(headers, "fix");
  const typeIdx = findHeaderIndex(headers, "type");
  const counts: IssueCounts = { missing: 0, duplicate: 0, bad_type: 0 };
  rows.forEach((row) => {
    const fixValue = fixIdx >= 0 ? normalizeFixValue(row[fixIdx]) : "apply";
    if (fixValue === "ignore") return;
    const type = `${row[typeIdx] ?? ""}`.trim().toLowerCase();
    if (type === "missing") counts.missing += 1;
    if (type === "duplicate") counts.duplicate += 1;
    if (type === "bad_type") counts.bad_type += 1;
  });

  const tableName = typeof table?.name === "string" ? table.name : reference.tableName;
  const rangeAddress =
    typeof tableRange?.address === "string"
      ? tableRange.address
      : reference.rangeA1
        ? `${reference.sheet}!${reference.rangeA1}`
        : undefined;
  const localRange = rangeAddress ? extractLocalRange(rangeAddress) : reference.rangeA1;
  const normalizedRangeA1 = localRange ?? undefined;
  const resolvedReference: IssuesSheetReference = {
    sheet: reference.sheet,
    tableName,
    rangeA1: normalizedRangeA1,
    createdAt: reference.createdAt,
  };
  return {
    sheet: reference.sheet,
    headers,
    rows,
    counts,
    rangeAddress,
    localRange,
    range: localRange ? parseA1Address(localRange) : null,
    reference: resolvedReference,
  };
}

async function markIssuesRowsIgnored(
  ctx: MacroContext,
  snapshot: IssueTableSnapshot,
  processedRowsByType: Record<IssueType, Set<string>>,
  fixIdx: number
) {
  if (fixIdx < 0) return;
  const rangeBounds = snapshot.range || (snapshot.localRange ? parseA1Address(snapshot.localRange) : null);
  if (!rangeBounds) return;
  const rowKeyIdx = findHeaderIndex(snapshot.headers, "rowkey");
  const typeIdx = findHeaderIndex(snapshot.headers, "type");
  const columnIdx = findHeaderIndex(snapshot.headers, "colonne");
  if (rowKeyIdx < 0 || typeIdx < 0) return;
  const rowsToMark = new Set<number>();
  ISSUE_TYPES.forEach((issueType) => {
    const keys = processedRowsByType[issueType];
    if (!keys || !keys.size) return;
    const matchedIdxs: number[] = [];
    snapshot.rows.forEach((row, idx) => {
      const rowKey = `${row[rowKeyIdx] ?? ""}`.trim();
      const rowType = `${row[typeIdx] ?? ""}`.trim().toLowerCase() as IssueType;
      const columnName = columnIdx >= 0 ? `${row[columnIdx] ?? ""}` : "";
      if (rowType !== issueType) return;
      const processedKey =
        issueType === ISSUE_TYPE_BAD_TYPE && columnName ? `${rowKey}::${normalizeHeader(columnName)}` : rowKey;
      if (keys.has(processedKey)) matchedIdxs.push(idx);
    });
    if (matchedIdxs.length) {
      ctx.log({
        level: "info",
        message: "validate_data issues_fix_marked",
        data: {
          issueType,
          processedIssueRowIdxsCount: matchedIdxs.length,
          firstIdxs: matchedIdxs.slice(0, 5),
        },
      });
      matchedIdxs.forEach((idx) => rowsToMark.add(idx));
    }
  });
  if (!rowsToMark.size) return;
  const sheetName = snapshot.reference?.sheet || snapshot.sheet;
  let ws: any;
  try {
    ws = getWorksheet(ctx.excelCtx, sheetName);
  } catch {
    return;
  }
  for (const rowIdx of Array.from(rowsToMark)) {
    if (typeof rowIdx !== "number" || rowIdx < 0) continue;
    const targetRow = rangeBounds.startRow + 1 + rowIdx;
    const targetCol = rangeBounds.startCol + fixIdx;
    let cellRange: any;
    try {
      cellRange =
        (typeof ws.getRangeByIndexes === "function" && ws.getRangeByIndexes(targetRow, targetCol, 1, 1)) ||
        (typeof ws.getRange === "function" && ws.getRange(rowColToA1(targetRow, targetCol)));
    } catch {
      cellRange = null;
    }
    if (!cellRange) continue;
    try {
      cellRange.values = [["ignore"]];
    } catch {}
  }
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
}

async function clearIssueHighlights(
  ctx: MacroContext,
  snapshot: IssueTableSnapshot,
  processedRowsByType: Record<IssueType, Set<string>>,
  dataHeaders: string[],
  dataBounds: AddressBounds | null,
  dataSheetName: string,
  liveIndexBuilder: (
    ctx: MacroContext,
    sheet: string,
    bounds: AddressBounds | null,
    headers: string[],
    rangeOverride?: string | null
  ) => Promise<{
    index: Map<string, { excelRow1: number; rowIdx: number; dataRowIdx: number }[]>;
    rows: unknown[][];
    bounds: AddressBounds | null;
  }>,
  rangeOverride?: string | null
) {
  const typeIdx = findHeaderIndex(snapshot.headers, "type");
  const rowKeyIdx = findHeaderIndex(snapshot.headers, "rowkey");
  const columnIdx = findHeaderIndex(snapshot.headers, "colonne");
  if (typeIdx < 0 || rowKeyIdx < 0 || columnIdx < 0) return;
  const targets: Array<{ rowKey: string; issueType: IssueType; columnName: string }> = [];
  ISSUE_TYPES.forEach((issueType) => {
    const keys = processedRowsByType[issueType];
    if (!keys || !keys.size) return;
    snapshot.rows.forEach((row) => {
      const rowType = `${row[typeIdx] ?? ""}`.trim().toLowerCase() as IssueType;
      const rowKey = `${row[rowKeyIdx] ?? ""}`.trim();
      const columnName = `${row[columnIdx] ?? ""}`;
      if (rowType !== issueType) return;
      const processedKey =
        issueType === ISSUE_TYPE_BAD_TYPE && columnName ? `${rowKey}::${normalizeHeader(columnName)}` : rowKey;
      if (!keys.has(processedKey)) return;
      targets.push({ rowKey, issueType, columnName });
    });
  });
  if (!targets.length) return;
  const live = await liveIndexBuilder(ctx, dataSheetName, dataBounds, dataHeaders, rangeOverride);
  const bounds = live.bounds || dataBounds;
  let ws: any;
  try {
    ws = getWorksheet(ctx.excelCtx, dataSheetName);
  } catch {
    return;
  }
  const cleared = new Set<string>();
  for (const target of targets) {
    const rowRefs = live.index.get(target.rowKey);
    const colIdx = findHeaderIndex(dataHeaders, target.columnName);
    if (colIdx < 0 || !bounds) continue;
    const ref = rowRefs?.[0];
    if (!ref) continue;
    const sheetRowIdx = ref.rowIdx;
    const sheetColIdx = bounds.startCol + colIdx;
    const key = `${sheetRowIdx}:${sheetColIdx}:${target.issueType}`;
    if (cleared.has(key)) continue;
    cleared.add(key);
    let cellRange: any = null;
    try {
      cellRange =
        (typeof ws.getRangeByIndexes === "function" && ws.getRangeByIndexes(sheetRowIdx, sheetColIdx, 1, 1)) ||
        (typeof ws.getRange === "function" && ws.getRange(rowColToA1(sheetRowIdx, sheetColIdx)));
    } catch {
      cellRange = null;
    }
    if (!cellRange) continue;
    try {
      if (cellRange.format && cellRange.format.fill) {
        cellRange.format.fill.clear();
      }
    } catch {}
  }
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
}

function computeSelectionIndexes(selectCfg: any, headers: string[], excludeNorm?: Set<string>) {
  if (!selectCfg || selectCfg.mode === "all") return headers.map((_h, idx) => idx);
  if (selectCfg.mode === "all_except_keys") {
    return headers
      .map((_h, idx) => idx)
      .filter((idx) => {
        const norm = normalizeHeader(headers[idx]);
        return !excludeNorm || !excludeNorm.has(norm);
      });
  }
  if (Array.isArray(selectCfg.columns)) {
    const idxs: number[] = [];
    selectCfg.columns.forEach((name: any) => {
      const { idx } = resolveHeaderIndex(headers, typeof name === "string" ? name : `${name}`);
      if (idx !== null && idx !== undefined && idx >= 0) idxs.push(idx);
    });
    if (idxs.length > 0) return idxs;
    if (selectCfg.columns.length === 0) return [];
  }
  return headers.map((_h, idx) => idx);
}

function ensureTableName(ws: any, hint: string) {
  const base = hint && hint.trim() ? hint.trim() : "Join_Result";
  const candidate = `${base}`;
  const stamp = Date.now();
  return `${candidate}_${stamp}`;
}

const macroJoinTables: MacroFn = async (paramsRaw, ctx) => {
  const params: JoinTablesParamsNormalized = { ...(paramsRaw || {}) } as any;
  params.left = params.left || ({} as any);
  params.right = params.right || ({} as any);
  const leftResolved = resolveBlockRefOrArtifact(params.left as BlockArtifactRef, ctx, { allowDefault: false });
  const rightResolved = resolveBlockRefOrArtifact(params.right as BlockArtifactRef, ctx, { allowDefault: false });
  if (!leftResolved.ok || !rightResolved.ok) {
    const reason = !leftResolved.ok ? `left (${leftResolved.reason})` : `right (${(rightResolved as any).reason})`;
    ctx.log({ level: "error", message: `join_tables: blockRef introuvable ${reason}`, macro: "join_tables", stepId: ctx.step.id });
    return { status: "error" };
  }
  params.left.blockRef = leftResolved.blockRef;
  params.right.blockRef = rightResolved.blockRef;

  if (!params.left?.blockRef || !params.right?.blockRef || !Array.isArray(params.keys) || params.keys.length === 0) {
    ctx.log({ level: "error", message: "join_tables: params incomplets (left/right/keys)", macro: "join_tables", stepId: ctx.step.id });
    return { status: "error" };
  }

  const leftData = await readBlockData(params.left.blockRef, ctx);
  const rightData = await readBlockData(params.right.blockRef, ctx);
  if (leftData.error || rightData.error) {
    ctx.log({ level: "error", message: "join_tables: block introuvable", macro: "join_tables", stepId: ctx.step.id });
    return { status: "error" };
  }

  const joinType: "left" | "inner" | "anti_left" | "full" = (params.joinType as any) || "left";
  const strategyDefault: JoinMatchStrategy = params.match?.defaultStrategy || "case_insensitive_trim";
  const plannedMappings: JoinKeyMapping[] = [];
  for (const k of params.keys) {
    const leftIdxInfo = resolveHeaderIndex(leftData.headers, k.left);
    const rightIdxInfo = resolveHeaderIndex(rightData.headers, k.right);
    if (leftIdxInfo.idx === null || rightIdxInfo.idx === null) {
      ctx.log({
        level: "error",
        message: `join_tables: cle inconnue (${k.left} ou ${k.right})`,
        macro: "join_tables",
        stepId: ctx.step.id,
      });
      return { status: "error" };
    }
    const lType = leftData.columnTypes?.[leftIdxInfo.idx];
    const rType = rightData.columnTypes?.[rightIdxInfo.idx];
    const strategy: JoinMatchStrategy =
      (k.strategy as JoinMatchStrategy) || (lType === "number" && rType === "number" ? "numeric" : strategyDefault);
    plannedMappings.push({ leftIdx: leftIdxInfo.idx, rightIdx: rightIdxInfo.idx, strategy, leftName: k.left, rightName: k.right });
  }

  const conflict = params.conflict || {};
  const conflictMode = (conflict.onDuplicateRightColumns as any) || "suffix";
  const rightSuffix = conflict.rightSuffix || "_r";
  const onMultiple = (conflict.onMultipleMatches as any) || "explode_rows";

  const describeMappings = (maps: JoinKeyMapping[]) =>
    maps.map((m) => `${leftData.headers[m.leftIdx] ?? m.leftName} <-> ${rightData.headers[m.rightIdx] ?? m.rightName}`).join(", ");

  const scoreMappings = (map: JoinKeyMapping[]) => {
    const rightIndex = new Map<string, number>();
    rightData.rows.forEach((row) => {
      const key = makeCompositeKey(row, map, "right");
      rightIndex.set(key, (rightIndex.get(key) || 0) + 1);
    });
    let matchedLeft = 0;
    let pairs = 0;
    leftData.rows.forEach((lrow) => {
      const key = makeCompositeKey(lrow, map, "left");
      const cnt = rightIndex.get(key) || 0;
      if (cnt > 0) {
        matchedLeft += 1;
        pairs += cnt;
      }
    });
    const leftNonEmpty = leftData.rows.length || 1;
    return { matchedLeft, pairs, matchRate: matchedLeft / leftNonEmpty };
  };

  const isExplicitSelection = (sel: any) => sel && sel.mode === "list" && Array.isArray(sel.columns);

  const buildOutputLayout = (map: JoinKeyMapping[]) => {
    const rightKeyNorms = new Set(map.map((m) => normalizeHeader(rightData.headers[m.rightIdx] ?? m.rightName)));
    const leftKeyNorms = new Set(map.map((m) => normalizeHeader(leftData.headers[m.leftIdx] ?? m.leftName)));
    const leftColsRaw = computeSelectionIndexes(params.select?.left, leftData.headers);
    const leftCols = leftColsRaw.length > 0 ? leftColsRaw : leftData.headers.map((_h, idx) => idx);
    const leftSelectedNorms = new Set(leftCols.map((idx) => normalizeHeader(leftData.headers[idx] ?? `Column${idx + 1}`)));
    const rightColsRaw = computeSelectionIndexes(
      params.select?.right,
      rightData.headers,
      params.select?.right?.mode === "all_except_keys" ? rightKeyNorms : undefined
    );
    const rightColsInitial = rightColsRaw.length > 0 ? rightColsRaw : rightData.headers.map((_h, idx) => idx);
    const rightSelectionExplicit = isExplicitSelection(params.select?.right);
    let rightCols: number[] = [];
    if (joinType !== "anti_left") {
      rightCols = rightColsInitial.filter((idx) => {
        const norm = normalizeHeader(rightData.headers[idx] || "");
        const isRightKey = rightKeyNorms.has(norm);
        const isKeyNameClash = isRightKey && leftKeyNorms.has(norm);
        if (isKeyNameClash && params.keepRightKeyColumns !== true && !rightSelectionExplicit) return false;
        if (!rightSelectionExplicit && leftSelectedNorms.has(norm)) return false;
        return true;
      });
      if (rightCols.length === 0 && rightData.headers.length > 0) {
        rightCols = rightData.headers
          .map((_h, i) => i)
          .filter((i) => {
        const norm = normalizeHeader(rightData.headers[i] || "");
        const isRightKey = rightKeyNorms.has(norm);
        const isKeyNameClash = isRightKey && leftKeyNorms.has(norm);
        if (isKeyNameClash && params.keepRightKeyColumns !== true && !rightSelectionExplicit) return false;
        if (!rightSelectionExplicit && leftSelectedNorms.has(norm)) return false;
        return true;
      });
        if (rightCols.length === 0) rightCols = [0];
      }
    }

    const finalHeaders: string[] = [];
    const columnMeta: { from: "left" | "right"; sourceIdx: number; header: string }[] = [];
    const usedNames = new Map<string, number>();
    const leftColMappings = leftCols.map((idx) => {
      const originalName = leftData.headers[idx] ?? `Column${idx + 1}`;
      let name = originalName;
      let norm = normalizeHeader(name);
      leftSelectedNorms.add(norm);
      if (usedNames.has(norm)) {
        let attempt = `${name}${rightSuffix}`;
        let attemptNorm = normalizeHeader(attempt);
        let counter = 1;
        while (usedNames.has(attemptNorm) && counter < 50) {
          counter += 1;
          attempt = `${name}${rightSuffix}${counter}`;
          attemptNorm = normalizeHeader(attempt);
        }
        name = attempt;
        norm = attemptNorm;
      }
      const outIdx = finalHeaders.length;
      finalHeaders.push(name);
      columnMeta[outIdx] = { from: "left", sourceIdx: idx, header: name };
      usedNames.set(norm, outIdx);
      return { sourceIdx: idx, outIdx };
    });

    const rightColMappings: { sourceIdx: number; outIdx: number }[] = [];
    rightCols.forEach((idx) => {
      const baseName = rightData.headers[idx] ?? `Column${idx + 1}`;
      const baseNorm = normalizeHeader(baseName);
      const isRightKey = rightKeyNorms.has(baseNorm);
      const isKeyNameClash = isRightKey && leftKeyNorms.has(baseNorm);
      const conflictsWithLeft = leftSelectedNorms.has(baseNorm) || leftKeyNorms.has(baseNorm);
      if (conflictMode === "skip" && usedNames.has(baseNorm)) return;
      if (!rightSelectionExplicit && params.keepRightKeyColumns !== true && isKeyNameClash) return;
      if (conflictMode === "overwrite_left" && usedNames.has(baseNorm)) {
        rightColMappings.push({ sourceIdx: idx, outIdx: usedNames.get(baseNorm)! });
        return;
      }
      let candidate = baseName;
      let candidateNorm = baseNorm;
      if (usedNames.has(candidateNorm) || conflictsWithLeft) {
        if (conflictMode === "suffix") {
          let attempt = `${baseName}${rightSuffix}`;
          let attemptNorm = normalizeHeader(attempt);
          let counter = 1;
          while (usedNames.has(attemptNorm) && counter < 50) {
            counter += 1;
            attempt = `${baseName}${rightSuffix}${counter}`;
            attemptNorm = normalizeHeader(attempt);
          }
          candidate = attempt;
          candidateNorm = attemptNorm;
        }
      }
      let outIdx = usedNames.get(candidateNorm) ?? finalHeaders.length;
      if (!usedNames.has(candidateNorm)) {
        usedNames.set(candidateNorm, outIdx);
        finalHeaders.push(candidate);
        columnMeta[outIdx] = { from: "right", sourceIdx: idx, header: candidate };
      }
      rightColMappings.push({ sourceIdx: idx, outIdx });
    });

    return { finalHeaders, columnMeta, leftColMappings, rightColMappings, rightKeyNorms };
  };

  const buildIndex = (map: JoinKeyMapping[]) => {
    const idx = new Map<string, { row: any[]; idx: number }[]>();
    rightData.rows.forEach((row, rIdx) => {
      const key = makeCompositeKey(row, map, "right");
      if (!idx.has(key)) idx.set(key, []);
      idx.get(key)!.push({ row, idx: rIdx });
    });
    return idx;
  };

  const runJoin = (map: JoinKeyMapping[], layout: ReturnType<typeof buildOutputLayout>) => {
    const { finalHeaders, leftColMappings, rightColMappings } = layout;
    const rightIndex = buildIndex(map);
    const out: any[][] = [];
    let pairs = 0;
    let matchedLeft = 0;
    const matchedRightIdx = new Set<number>();
    const append = (l: any[] | null, r: any[] | null) => {
      const row = Array(finalHeaders.length).fill("");
      if (l) leftColMappings.forEach((m) => (row[m.outIdx] = l[m.sourceIdx]));
      if (r) rightColMappings.forEach((m) => (row[m.outIdx] = r[m.sourceIdx]));
      out.push(row);
    };
    leftData.rows.forEach((lrow) => {
      const key = makeCompositeKey(lrow, map, "left");
      const matches = rightIndex.get(key) || [];
      if (matches.length > 0) {
        matchedLeft += 1;
        pairs += matches.length;
        if (joinType !== "anti_left") {
          const chosen = onMultiple === "first" ? matches.slice(0, 1) : matches;
          chosen.forEach((m) => {
            matchedRightIdx.add(m.idx);
            append(lrow, m.row);
          });
        }
      } else if (joinType === "left" || joinType === "full" || joinType === "anti_left") {
        append(lrow, null);
      }
    });
    if (joinType === "full") {
      rightData.rows.forEach((rrow, idx) => {
        if (!matchedRightIdx.has(idx)) append(null, rrow);
      });
    }
    return { out, pairs, matchedLeft };
  };

  const initialScore = scoreMappings(plannedMappings);
  let mappingsFinal = [...plannedMappings];
  let fallbackApplied = false;
  let fallbackScore = initialScore.matchRate;
  let fallbackCandidate: JoinKeyMapping[] | null = null;

  if (initialScore.matchRate < 0.2) {
    const sampleLimit = 50;
    const leftSamples = leftData.rows.slice(0, sampleLimit);
    const rightSamples = rightData.rows.slice(0, sampleLimit);
    const candidates: JoinKeyMapping[][] = [];
    leftData.headers.forEach((lh, li) => {
      rightData.headers.forEach((rh, ri) => {
        const lt = leftData.columnTypes?.[li];
        const rt = rightData.columnTypes?.[ri];
        if (lt && rt && lt !== rt) return;
        candidates.push([{ leftIdx: li, rightIdx: ri, strategy: "case_insensitive_trim", leftName: lh, rightName: rh } as any]);
      });
    });
    const scoreCandidate = (maps: JoinKeyMapping[]) => {
      const idx = new Map<string, number>();
      rightSamples.forEach((r) => {
        const key = makeCompositeKey(r, maps as any, "right");
        idx.set(key, (idx.get(key) || 0) + 1);
      });
      let hits = 0;
      leftSamples.forEach((l) => {
        const key = makeCompositeKey(l, maps as any, "left");
        if (idx.has(key)) hits += 1;
      });
      return (hits as number) / (leftSamples.length || 1);
    };
    let bestCand: JoinKeyMapping[] | null = null;
    let bestScore = initialScore.matchRate;
    candidates.forEach((c) => {
      const sc = scoreCandidate(c as any);
      if (sc > bestScore) {
        bestScore = sc;
        bestCand = c as any;
      }
    });
    if (bestCand) {
      fallbackCandidate = bestCand;
      fallbackScore = bestScore;
    }
  }

  if (joinType === "anti_left") {
    fallbackCandidate = null;
  }
  if (fallbackCandidate && fallbackScore > initialScore.matchRate) {
    const confirmationId = `${ctx.step.id}:join_key_fallback`;
    const sig = `join_fallback:${describeMappings(plannedMappings)}->${describeMappings(fallbackCandidate)}`;
    const decision = ctx.decisions[confirmationId];
    const cachedDecision = ctx.decisions[sig];
    const allowFallback = params.allowKeyFallback === true || decision === "use_fallback_keys";
    const keepPlanKeys = decision === "keep_plan_keys";
    const abort = decision === "abort";
    const reuse = cachedDecision === "use_fallback_keys" || cachedDecision === "keep_plan_keys";
    const reuseFallback = cachedDecision === "use_fallback_keys";
    if (abort) {
      ctx.log({
        level: "warn",
        message: "join_tables: fallback refuse, annule",
        macro: "join_tables",
        stepId: ctx.step.id,
      });
      return { status: "skipped" };
    }
    if (!allowFallback && !keepPlanKeys && !reuse) {
      const question = `La cle du plan (${describeMappings(plannedMappings)}) semble incorrecte (taux ${initialScore.matchRate.toFixed(2)}). Utiliser la cle proposee ${describeMappings(fallbackCandidate)} (taux ${fallbackScore.toFixed(2)}) ?`;
      const choices = [
        { id: "use_fallback_keys", label: "Utiliser la cle proposee" },
        { id: "keep_plan_keys", label: "Conserver la cle du plan" },
        { id: "abort", label: "Annuler" },
      ];
      ctx.log({
        level: "warn",
        message: `join_tables: confirmation demandee pour fallback (${describeMappings(plannedMappings)} -> ${describeMappings(fallbackCandidate)})`,
        macro: "join_tables",
        stepId: ctx.step.id,
      });
      return { requiresConfirmation: { id: confirmationId, question, choices } };
    }
    if (allowFallback || reuseFallback) {
      mappingsFinal = fallbackCandidate;
      fallbackApplied = true;
      ctx.decisions[sig] = "use_fallback_keys";
    } else if (keepPlanKeys || reuse) {
      ctx.decisions[sig] = "keep_plan_keys";
    }
  }

  const layout = buildOutputLayout(mappingsFinal);
  const best = runJoin(mappingsFinal, layout);

  const outputRows = best.out;
  const matchPairs = best.pairs;

  if (fallbackApplied) {
    ctx.log({
      level: "warn",
      message: `join_tables: keyFallbackApplied ${describeMappings(plannedMappings)} -> ${describeMappings(mappingsFinal)} matchRate ${initialScore.matchRate.toFixed(2)} -> ${fallbackScore.toFixed(2)}`,
      macro: "join_tables",
      stepId: ctx.step.id,
    });
  }

  const { finalHeaders, columnMeta } = layout;

  const rowCount = Math.max(1, outputRows.length + 1);
  const colCount = Math.max(1, finalHeaders.length);
  const outputMode: "right" | "below" | "newSheet" = (params.output?.mode as any) || "newSheet";
  const placement = await resolvePlacement(
    {
      mode: outputMode,
      anchor: outputMode === "newSheet" ? undefined : params.output?.anchor || { blockRef: params.left.blockRef },
      minBlankArea: { rows: rowCount, cols: colCount },
      newSheetNameHint: params.output?.sheetName || params.output?.tableName || "Join_Result",
      avoidOverwrite: true,
    },
    ctx,
    `${ctx.step.id}:join_place`
  );
  if (placement.requiresConfirmation) return { requiresConfirmation: placement.requiresConfirmation };
  if (!placement.address) return placement;

  const destSheetName =
    placement.address.split("!")[0] ||
    params.output?.sheetName ||
    ctx.context.active.sheetName ||
    leftData.sheetName ||
    (ctx.context.sheets[0]?.name ?? "Sheet1");
  const destWs = placement.worksheet || getWorksheet(ctx.excelCtx, destSheetName);
  const placementBounds = parseA1Address(placement.address) || { startRow: 0, startCol: 0, endRow: rowCount - 1, endCol: colCount - 1 };
  const targetAddress = makeRangeAddress(destSheetName, placementBounds.startRow, placementBounds.startCol, rowCount, colCount);
  const localAddr = targetAddress.split("!")[1] || targetAddress;
  const targetRange = destWs.getRange
    ? destWs.getRange(localAddr)
    : destWs.getRangeByIndexes?.(placementBounds.startRow ?? 0, placementBounds.startCol ?? 0, rowCount, colCount);
  const valuesToWrite = [finalHeaders, ...outputRows];
  if (targetRange) (targetRange as any).values = valuesToWrite;
  if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();

  // Apply date formats heuristically on output columns
  if (targetRange) {
    const totalRows = outputRows.length + 1;
    const totalCols = finalHeaders.length;
    const nfMatrix = Array.from({ length: totalRows }, () => Array.from({ length: totalCols }, () => null as any));
    columnMeta.forEach((meta, colIdx) => {
      const header = finalHeaders[colIdx] || meta.header;
      const colValues = outputRows.map((r) => r[colIdx]).filter((v) => v !== null && v !== undefined && v !== "");
      // attempt to coerce date-like text to Date for formatting
      const parsedValues = colValues.map((v) => {
        if (typeof v === "string") {
          const parsed = parseDateLikeString(v);
          if (parsed) return parsed;
        }
        return v;
      });
      if (isLikelyDateColumn(header, parsedValues)) {
        for (let r = 0; r < totalRows; r += 1) (nfMatrix[r] as any)[colIdx] = "yyyy-mm-dd";
        ctx.log({
          level: "info",
          message: `join_tables: dateFormatApplied ${header}`,
          macro: "join_tables",
          stepId: ctx.step.id,
        });
      }
    });
    if (nfMatrix.some((row) => row.some((v) => v))) {
      (targetRange as any).numberFormat = nfMatrix;
    }
  }

  // auto format corporate preset on output (header bold + bg + autofit)
  try {
    const headerRange = targetRange.getRange
      ? targetRange.getRange(`${rowColToA1(placementBounds.startRow, placementBounds.startCol)}:${rowColToA1(placementBounds.startRow, placementBounds.startCol + colCount - 1)}`)
      : null;
    if (headerRange?.format) {
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "#ffffff";
      headerRange.format.fill.color = "#0e2a80";
    }
    if (targetRange?.format?.autofitColumns) {
      targetRange.format.autofitColumns();
    }
  } catch {
    // best effort
  }

  let tableName = ensureTableName(destWs, params.output?.tableName || params.output?.sheetName || "Join_Result");
  let tableCreated = false;
  try {
    if (destWs.tables?.add) {
      const tbl = destWs.tables.add(localAddr, true);
      if (tbl) {
        (tbl as any).name = tableName;
        tableCreated = true;
      }
    }
  } catch {
    tableName = tableName || "Join_Result";
  }

  const headerAliasGroups: Record<string, Set<string>> = {};
  const addAliasGroup = (...names: (string | undefined)[]) => {
    const normalizedKeys = names
      .map((name) => (typeof name === "string" ? name.trim() : ""))
      .filter(Boolean)
      .map((name) => ({ norm: normalizeHeader(name), label: name }))
      .filter((entry) => entry.norm);
    if (!normalizedKeys.length) return;
    const trimmed = normalizedKeys.map((entry) => entry.label);
    normalizedKeys.forEach((entry) => {
      const existing = headerAliasGroups[entry.norm] || new Set<string>();
      trimmed.forEach((label) => existing.add(label));
      headerAliasGroups[entry.norm] = existing;
    });
  };
  [plannedMappings, mappingsFinal].forEach((maps) => {
    maps.forEach((m) => {
      const leftName = (leftData.headers[m.leftIdx] ?? m.leftName ?? "").toString().trim();
      const rightName = (rightData.headers[m.rightIdx] ?? m.rightName ?? "").toString().trim();
      addAliasGroup(leftName || undefined, rightName || undefined);
    });
  });
  const headerAliases =
    Object.keys(headerAliasGroups).length > 0
      ? Object.fromEntries(
          Object.entries(headerAliasGroups).map(([norm, set]) => [norm, Array.from(set)])
        )
      : undefined;
  const startCell = rowColToA1(placementBounds.startRow, placementBounds.startCol);
  const outputBlockRef = `${destSheetName}!${localAddr}`;
  const artifacts = placement.artifacts ? [...placement.artifacts] : [];
  artifacts.push({
    type: "table",
    kind: "table",
    sheet: destSheetName,
    sheetName: destSheetName,
    anchor: startCell,
    fromStep: ctx.step.id,
    tableName,
    rows: outputRows.length,
    cols: finalHeaders.length,
    rowCount: outputRows.length + 1,
    colCount: finalHeaders.length,
    headers: finalHeaders.slice(),
    headerAliases,
    joinType,
    matchedRows: matchPairs,
    outputRows: outputRows.length,
    blockRef: outputBlockRef,
    address: localAddr,
    addressA1: localAddr,
    details: { tableCreated },
  });

  const keysUsed = mappingsFinal.map((m) => ({
    left: leftData.headers[m.leftIdx] ?? m.leftName,
    right: rightData.headers[m.rightIdx] ?? m.rightName,
    strategy: m.strategy,
  }));

  ctx.log({
    level: "info",
    message: `join_tables: left=${leftData.rows.length} right=${rightData.rows.length} matches=${matchPairs} outputRows=${outputRows.length} joinType=${joinType}`,
    macro: "join_tables",
    stepId: ctx.step.id,
    data: { keysPlan: params.keys, keysUsed, conflict: params.conflict, output: params.output, tableName, tableCreated, fallbackApplied },
  });

  ctx.log({
    level: "info",
    message: `join_tables output blockRef=${outputBlockRef} tableName=${tableName}`,
    macro: "join_tables",
    stepId: ctx.step.id,
  });

  return { artifacts };
};

function rangesOverlap(a: AddressBounds, b: AddressBounds): boolean {
  return !(a.endRow < b.startRow || b.endRow < a.startRow || a.endCol < b.startCol || b.endCol < a.startCol);
}

async function sanitizeHeaderRow(range: any, bounds: AddressBounds | null, ctx: MacroContext) {
  if (!bounds) return { changed: false, headers: [] as string[] };
  try {
    if (typeof range?.load === "function") range.load("values");
    if (typeof (ctx.excelCtx as any).sync === "function") await (ctx.excelCtx as any).sync();
  } catch {
    // best effort, continue
  }
  const values: any[][] = (range as any)?.values || [];
  const width = Math.max(1, bounds.endCol - bounds.startCol + 1, values?.[0]?.length || 0);
  const headerRow = Array.isArray(values?.[0]) ? values[0].slice(0, width) : [];
  const fixed: string[] = [];
  const seen = new Set<string>();
  for (let c = 0; c < width; c += 1) {
    const raw = headerRow[c];
    const base = raw === null || raw === undefined || `${raw}`.trim() === "" ? `Column${c + 1}` : `${raw}`;
    let name = base;
    let idx = 1;
    while (seen.has(name.toLowerCase())) {
      idx += 1;
      name = `${base}_${idx}`;
    }
    seen.add(name.toLowerCase());
    fixed[c] = name;
  }
  const changed = fixed.some((v, i) => v !== headerRow[i]);
  if (changed) {
    const newValues =
      values && values.length ? values.slice() : Array.from({ length: Math.max(1, (range as any)?.rowCount || 1) }, () => [] as any[]);
    newValues[0] = fixed;
    (range as any).values = newValues;
  }
  return { changed, headers: fixed };
}

export const macros: Record<AgentMacroName, MacroFn> = {
  place_output: macroPlaceOutput,
  write_formula: macroWriteFormula,
  apply_format: macroApplyFormat,
  create_chart: macroCreateChart,
  table_view: macroTableView,
  validate_data: macroValidateData,
  join_tables: macroJoinTables,
  summarize_actions: async () => ({}),
};

export const __dateTest__ = { parseDateLikeString, isAllDigits };
