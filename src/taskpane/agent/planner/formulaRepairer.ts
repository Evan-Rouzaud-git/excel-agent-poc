/* global fetch */
import { WorkbookContextSnapshot, BlockSnapshot } from "../../context/types";
import { PlanStep } from "../types";
import { findBlock } from "../utils";

const DEFAULT_HOST =
  (typeof process !== "undefined" && (process as any).env?.OLLAMA_HOST) ||
  (typeof window !== "undefined" && (window as any)?.OLLAMA_HOST) ||
  "http://localhost:11434";
const DEFAULT_MODEL =
  (typeof process !== "undefined" && (process as any).env?.OLLAMA_MODEL) ||
  (typeof window !== "undefined" && (window as any)?.OLLAMA_MODEL) ||
  "qwen3:8b";
const DEFAULT_TIMEOUT =
  Number(
    (typeof process !== "undefined" && (process as any).env?.DEMO_EVAL_TIMEOUT_MS) ||
      (typeof process !== "undefined" && (process as any).env?.AGENT_TIMEOUT_MS) ||
      (typeof window !== "undefined" && (window as any)?.DEMO_EVAL_TIMEOUT_MS) ||
      (typeof window !== "undefined" && (window as any)?.AGENT_TIMEOUT_MS)
  ) || 12000;
const MAX_ATTEMPTS = 2;

export interface RepairInput {
  step: PlanStep;
  context: WorkbookContextSnapshot;
  userPrompt: string;
}

export interface RepairAttemptLog {
  attempt: number;
  rawText?: string;
  parsed?: any;
  error?: string;
  passedChecks?: boolean;
  notesFromModel?: string;
}

export interface RepairLog {
  ok: boolean;
  applied: boolean;
  reason?: string;
  notes?: string;
  originalFormula?: string;
  finalFormula?: string;
  blockRef?: string | null;
  attempts: RepairAttemptLog[];
}

export interface RepairResult {
  patchedStep: PlanStep;
  repairLog: RepairLog;
}

export interface RepairOptions {
  host?: string;
  model?: string;
  timeoutMs?: number;
  maxAttempts?: number;
  callOllama?: (args: { systemPrompt: string; userPrompt: string; model: string; host: string; timeout: number }) => Promise<string>;
  logger?: (level: "info" | "warn" | "error", message: string, data?: any) => void;
}

const FORMULA_SYSTEM_PROMPT = `You are the Formula Repairer for an Excel Office.js agent.
Return ONLY a single JSON object. No markdown. No commentary.
Schema: {"formula": string, "numberFormat"?: string, "headerName"?: string, "fillDown"?: boolean, "notes"?: string}
Rules (hard constraints):
- Output JSON only.
- Formula must be Excel invariant English (Range.formulas friendly).
- Use ONLY the headers provided for the target block; never invent columns.
- If tableName is provided, use structured references ([@Header], TableName[Header], TableName[[#This Row],[Header]]). Never use A1-style refs.
- Allowed edits: formula (required), numberFormat (optional), headerName (optional), fillDown (optional). Do NOT change anything else.
- If a required column is missing, set formula to "" and notes="missing required column: <name>". Do not invent headers or add confirmations.
- Prefer row-wise calculations with [@Header] for ratios/margins on the current row.
- MoM variation: revenue(t) - revenue(t-1) with first row blank.
- MoM growth rate: (t - t-1) / (t-1) with first row blank (""), numberFormat can be percentage.

Few-shot examples:
1) Headers: Revenus, Depenses -> margin percent
{"formula":"=IFERROR([@Revenus]/[@Depenses],0)","numberFormat":"0.0%","fillDown":true}
2) MoM variation on table TableBudget, header Revenus
{"formula":"=IF(ROW()=ROW(TableBudget[#Headers])+1,\"\",[@Revenus]-OFFSET([@[Revenus]],-1,0))","fillDown":true}
3) MoM growth rate on table TableBudget, header Revenus
{"formula":"=IF(ROW()=ROW(TableBudget[#Headers])+1,\"\",IFERROR(([@Revenus]-OFFSET([@[Revenus]],-1,0))/OFFSET([@[Revenus]],-1,0),\"\"))","numberFormat":"0.0%","fillDown":true}`;

function normalizeHeader(text?: string): string {
  const src = (text || "").toString().trim().toLowerCase();
  let out = "";
  let inSpace = false;
  for (const ch of src) {
    const isWs = ch === " " || ch === "\t" || ch === "\n" || ch === "\r" || ch === "\f";
    if (isWs) {
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

function stripFences(text: string): string {
  const trimmed = text.trim();
  if (trimmed.startsWith("```") && trimmed.includes("}")) {
    let withoutStart = trimmed.slice(3);
    const firstNewline = withoutStart.indexOf("\n");
    if (firstNewline >= 0) {
      withoutStart = withoutStart.slice(firstNewline + 1);
    }
    const endIdx = withoutStart.lastIndexOf("```");
    const body = endIdx >= 0 ? withoutStart.slice(0, endIdx) : withoutStart;
    return body.trim();
  }
  return trimmed;
}

function extractFirstJsonObject(text: string): string | null {
  const cleaned = stripFences(text);
  const start = cleaned.indexOf("{");
  if (start < 0) return null;
  let depth = 0;
  for (let i = start; i < cleaned.length; i += 1) {
    const ch = cleaned[i];
    if (ch === "{") depth += 1;
    if (ch === "}") {
      depth -= 1;
      if (depth === 0) return cleaned.slice(start, i + 1);
    }
  }
  return null;
}

function parseJson(text: string): { ok: true; value: any } | { ok: false; error: any } {
  const candidate = extractFirstJsonObject(text) ?? text;
  try {
    return { ok: true, value: JSON.parse(candidate) };
  } catch (err) {
    return { ok: false, error: err };
  }
}

function extractStructuredHeaders(formula: string): string[] {
  const headers: string[] = [];
  let idx = 0;
  while (idx < formula.length) {
    const open = formula.indexOf("[", idx);
    if (open < 0) break;
    const nextChar = formula[open + 1];
    if (nextChar === "#") {
      idx = open + 1;
      continue;
    }
    const close = formula.indexOf("]", open + 1);
    if (close < 0) break;
    const raw = formula.slice(open + 1, close).trim();
    const cleaned = raw.startsWith("@") ? raw.slice(1).trim() : raw;
    if (cleaned) headers.push(cleaned);
    idx = close + 1;
  }
  return headers;
}

function containsA1Reference(formula: string): boolean {
  const upper = formula.toUpperCase();
  for (let i = 0; i < upper.length; i += 1) {
    const ch = upper[i] || "";
    const isLetter = ch >= "A" && ch <= "Z";
    const isDollar = ch === "$";
    if (!isLetter && !isDollar) continue;
    let idx = i;
    if (upper[idx] === "$") idx += 1;
    let letters = "";
    while (idx < upper.length) {
      const current = upper[idx] || "";
      if (!(current >= "A" && current <= "Z") || letters.length >= 3) break;
      letters += current;
      idx += 1;
    }
    if (!letters) continue;
    if (upper[idx] === "$") idx += 1;
    let digits = "";
    while (idx < upper.length) {
      const current = upper[idx] || "";
      if (!(current >= "0" && current <= "9")) break;
      digits += current;
      idx += 1;
    }
    if (digits) {
      const prev = i > 0 ? upper[i - 1] || "" : "";
      const prevIsWord = (prev >= "A" && prev <= "Z") || (prev >= "0" && prev <= "9") || prev === "_";
      if (!prevIsWord) return true;
    }
  }
  return false;
}

function shortPreview(block?: BlockSnapshot) {
  if (!block || !Array.isArray(block.preview)) return [] as any[][];
  return block.preview.slice(0, Math.min(4, block.preview.length));
}

async function callOllamaChat(args: { systemPrompt: string; userPrompt: string; model: string; host: string; timeout: number }): Promise<string> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), args.timeout);
  try {
    const res = await fetch(`${args.host}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Connection: "close" },
      body: JSON.stringify({
        model: args.model,
        stream: false,
        messages: [
          { role: "system", content: args.systemPrompt },
          { role: "user", content: args.userPrompt },
        ],
      }),
      signal: controller.signal,
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const textBody = await res.text();
    let parsed: any;
    try {
      parsed = JSON.parse(textBody);
    } catch {
      parsed = undefined;
    }
    const text = parsed?.message?.content || parsed?.content || textBody || "";
    return typeof text === "string" ? text : JSON.stringify(text);
  } finally {
    clearTimeout(timer);
  }
}

function buildUserPrompt(args: {
  userPrompt: string;
  step: PlanStep;
  blockRef: string | null;
  block: BlockSnapshot | undefined;
  contextJson: string;
  retryNote?: string;
}): string {
  const headers = args.block?.headers || [];
  const tableName = (args.block?.source as any)?.tableName || null;
  const blockInfo = `Target block: ${args.blockRef ?? "unknown"}\nkind=${args.block?.kind ?? "unknown"}\n` +
    `tableName=${tableName ?? "none"}\nheaders=[${headers.join(", ")}]\ncolumnTypes=[${(args.block?.columnTypes || []).join(", ")}]\n` +
    `preview=${JSON.stringify(shortPreview(args.block))}`;
  const retry = args.retryNote ? `\nPrevious attempt invalid: ${args.retryNote}` : "";
  return [
    `User request: ${args.userPrompt}`,
    `Current write_formula step: ${JSON.stringify({ target: args.step.params?.target, formula: args.step.params?.formula, numberFormat: args.step.params?.numberFormat })}`,
    blockInfo,
    `Allowed headers (use only these): ${headers.join(", ") || "<none>"}`,
    `Workbook context JSON: ${args.contextJson}`,
    retry,
  ]
    .filter(Boolean)
    .join("\n\n");
}

function postChecks(result: any, block: BlockSnapshot | undefined): { ok: boolean; reason?: string; notes?: string } {
  if (!result || typeof result !== "object") return { ok: false, reason: "no_object" };
  if (typeof result.formula !== "string") return { ok: false, reason: "missing_formula" };
  const trimmedFormula = result.formula.trim();
  if (!trimmedFormula) return { ok: false, reason: "empty_formula", notes: result.notes };

  const headers = block?.headers || [];
  const normalized = headers.map(normalizeHeader);
  const inFormula = extractStructuredHeaders(trimmedFormula);
  const unknown = inFormula.filter((h) => !normalized.includes(normalizeHeader(h)));
  if (unknown.length > 0) return { ok: false, reason: `unknown_headers:${unknown.join("|")}`, notes: result.notes };

  const isTable = (block?.source as any)?.type === "table";
  if (isTable && containsA1Reference(trimmedFormula)) {
    return { ok: false, reason: "a1_ref_forbidden", notes: result.notes };
  }
  return { ok: true };
}

function applyPatch(step: PlanStep, suggestion: any): PlanStep {
  const target = { ...(step.params?.target || {}) } as any;
  const params = { ...(step.params || {}) } as any;

  params.formula = typeof suggestion.formula === "string" ? suggestion.formula : params.formula;
  if (typeof suggestion.numberFormat === "string") params.numberFormat = suggestion.numberFormat;
  if (typeof suggestion.fillDown === "boolean") params.fillDown = suggestion.fillDown;
  if (suggestion.headerName && typeof suggestion.headerName === "string") {
    target.headerName = suggestion.headerName;
  }
  if (Object.keys(target).length > 0) params.target = target;
  return { ...step, params };
}

function resolveBlockRef(step: PlanStep, context: WorkbookContextSnapshot): string | null {
  const targetRef = (step.params as any)?.target?.blockRef;
  if (typeof targetRef === "string" && targetRef.trim()) return targetRef;
  return (
    context.active.selectionInBlockId ||
    context.active.nearestBlockId ||
    context.sheets[0]?.blocks[0]?.id ||
    null
  );
}

export async function repairWriteFormulaStep(
  step: PlanStep,
  workbookContext: WorkbookContextSnapshot,
  userPrompt: string,
  opts: RepairOptions = {}
): Promise<RepairResult> {
  if ((typeof process !== "undefined" && (process as any).env?.DEMO_EVAL_LOG_FORMULA_REPAIR) === "1") {
    console.log("[formula-corrector] ON");
  }
  const attempts: RepairAttemptLog[] = [];
  const originalFormula = (step.params as any)?.formula;
  const blockRef = resolveBlockRef(step, workbookContext);
  const { block } = blockRef ? findBlock(blockRef, workbookContext) : { block: undefined };
  const host = opts.host || DEFAULT_HOST;
  const model = opts.model || DEFAULT_MODEL;
  const timeout = opts.timeoutMs || DEFAULT_TIMEOUT;
  const maxAttempts = Math.max(1, opts.maxAttempts || MAX_ATTEMPTS);
  const call = opts.callOllama || callOllamaChat;
  const contextJson = JSON.stringify(workbookContext);

  let patchedStep: PlanStep = step;
  let success = false;
  let finalReason: string | undefined;
  let finalNotes: string | undefined;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const retryNote = attempts.length ? attempts[attempts.length - 1]?.error || attempts[attempts.length - 1]?.notesFromModel : undefined;
    const userMessage = buildUserPrompt({
      userPrompt,
      step,
      blockRef,
      block,
      contextJson,
      retryNote,
    });

    let rawText = "";
    let parsed: any;
    try {
      rawText = await call({ systemPrompt: FORMULA_SYSTEM_PROMPT, userPrompt: userMessage, model, host, timeout });
      const parsedJson = parseJson(rawText);
      if (!parsedJson.ok) {
        attempts.push({ attempt, rawText, error: "json_parse_error" });
        finalReason = "json_parse_error";
        continue;
      }
      parsed = parsedJson.value;
      const checks = postChecks(parsed, block);
      if (!checks.ok) {
        attempts.push({ attempt, rawText, parsed, error: checks.reason, notesFromModel: checks.notes });
        finalReason = checks.reason;
        finalNotes = checks.notes;
        continue;
      }
      patchedStep = applyPatch(step, parsed);
      attempts.push({ attempt, rawText, parsed, passedChecks: true, notesFromModel: parsed?.notes });
      success = true;
      finalNotes = parsed?.notes;
      break;
    } catch (err: any) {
      const message = err?.name === "AbortError" ? "timeout" : err?.message || String(err);
      attempts.push({ attempt, rawText, error: message });
      finalReason = message;
      continue;
    }
  }

  const repairLog: RepairLog = {
    ok: success,
    applied: success,
    reason: success ? undefined : finalReason,
    notes: finalNotes,
    originalFormula,
    finalFormula: success ? (patchedStep.params as any)?.formula : originalFormula,
    blockRef,
    attempts,
  };

  if (opts.logger) {
    const level = success ? "info" : "warn";
    opts.logger(level, success ? "formula repair applied" : "formula repair failed", repairLog);
  }

  return { patchedStep, repairLog };
}

export const __private__ = { extractFirstJsonObject, parseJson, postChecks, extractStructuredHeaders, containsA1Reference, FORMULA_SYSTEM_PROMPT };
