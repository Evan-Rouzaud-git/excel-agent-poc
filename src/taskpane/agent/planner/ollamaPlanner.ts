import { validatePlan } from "../planSchema";
import { SYSTEM_PROMPT } from "./systemPrompt";
import { normalizePlan } from "./normalizePlan";
import { sanitizePlan } from "./sanitizePlan";
import { sanitizeTableViewParams } from "../tableViewUtils";
import { canonicalizePlan } from "../canonicalizePlan";
import { normalizeHeader } from "../normalizeHeader";
import { PlanContextTrace } from "../types";
import { validatePlanInvariants } from "../planInvariants";

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
  ) || 90000;
type PlannerResultBase = {
  rawText?: string;
  rawTextRetry?: string;
  sanitizeNotes?: string[];
  parseError?: string | null;
  failureStage?: string | null;
  plan?: any;
  fallbackUsed?: boolean;
  deterministicRepairApplied?: boolean;
};

export type PlannerResult =
  | (PlannerResultBase & {
      status: "ok";
      plan: any;
      fallbackUsed?: boolean;
      deterministicRepairApplied?: boolean;
    })
  | (PlannerResultBase & {
      status: "invalid_plan";
      errors: any[];
      plan?: any;
    })
  | (PlannerResultBase & {
      status: "error";
      error: string;
    });

async function callOllama(prompt: string, model: string, host: string, timeout: number): Promise<string> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    const res = await fetch(`${host}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Connection: "close" },
      body: JSON.stringify({
        model,
        stream: false,
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: prompt },
        ],
      }),
      signal: controller.signal,
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const textBody = await res.text();
    let data: any;
    try {
      data = JSON.parse(textBody);
    } catch {
      data = {};
    }
    const text = data?.message?.content || data?.content || textBody || "";
    return typeof text === "string" ? text : JSON.stringify(text);
  } finally {
    clearTimeout(timer);
  }
}

function stripFences(text: string): string {
  const fence = text.trim();
  if (fence.startsWith("```")) {
    let withoutStart = fence.slice(3);
    const firstNl = withoutStart.indexOf("\n");
    if (firstNl >= 0) {
      withoutStart = withoutStart.slice(firstNl + 1);
    }
    const endIdx = withoutStart.lastIndexOf("```");
    const body = endIdx >= 0 ? withoutStart.slice(0, endIdx) : withoutStart;
    return body.trim();
  }
  return text.trim();
}

export function extractStrictJSONObject(text: string): string | null {
  if (!text) return null;
  const cleaned = stripFences(text);
  const trimmed = cleaned.trim();
  if (!trimmed) return null;
  const start = trimmed.indexOf("{");
  if (start < 0) return null;
  let depth = 0;
  let inString = false;
  let escapeNext = false;
  for (let i = start; i < trimmed.length; i += 1) {
    const ch = trimmed[i];
    if (escapeNext) {
      escapeNext = false;
      continue;
    }
    if (ch === "\\") {
      escapeNext = true;
      continue;
    }
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (!inString) {
      if (ch === "{") {
        depth += 1;
      } else if (ch === "}") {
        depth -= 1;
        if (depth === 0) {
          const candidate = trimmed.slice(start, i + 1).trim();
          if (candidate.startsWith("{") && candidate.endsWith("}")) {
            return candidate;
          }
          return null;
        }
      }
    }
  }
  return null;
}

export function parseJson(text: string): { ok: true; value: any } | { ok: false; error: any } {
  const strictCandidate = extractStrictJSONObject(text);
  const tryParse = (input: string) => {
    try {
      return { ok: true as const, value: JSON.parse(input) };
    } catch (err) {
      return { ok: false as const, error: err };
    }
  };
  if (strictCandidate) {
    const strictResult = tryParse(strictCandidate);
    if (strictResult.ok) return strictResult;
    return strictResult;
  }
  const baseText = stripFences(text) || text || "";
  const first = tryParse(baseText);
  if (first.ok) return first;
  // minimal repair: trim noise before first { and after last }
  const start = baseText.indexOf("{");
  const end = baseText.lastIndexOf("}");
  if (start >= 0 && end > start) {
    const sliced = baseText.slice(start, end + 1);
    const second = tryParse(sliced);
    if (second.ok) return second;
    return second;
  }
  return first;
}

const forbiddenKeys = ["action", "parameters", "plan", "target_block_ref", "targetBlockRef", "steps.action", "steps.parameters"];

function hasForbiddenKeys(obj: any): boolean {
  if (!obj || typeof obj !== "object") return false;
  const keys = Object.keys(obj);
  for (const k of keys) {
    if (forbiddenKeys.includes(k)) return true;
  }
  if (Array.isArray(obj)) return obj.some((v) => hasForbiddenKeys(v));
  return keys.some((k) => hasForbiddenKeys(obj[k]));
}

export function autoRepairFromIntent(plan: any, context?: any, warnings?: string[]) {
  let changed = false;
  const clone = { ...(plan || {}) };
  const safeContext = context || { sheets: [], active: {} };
  const safeWarnings = warnings || [];
  clone.steps = (plan?.steps || []).map((step: any) => {
    if (!step || step.macro !== "table_view") return step;
    const repairedParams = sanitizeTableViewParams(step.params || {}, safeContext, safeWarnings);
    const filterCols = (Array.isArray(repairedParams.filter) ? repairedParams.filter : [])
      .map((f: any) => (typeof f?.col === "string" ? f.col : ""))
      .filter(Boolean);
    if (filterCols.length) {
      const selectCols = Array.isArray(repairedParams.select) ? [...repairedParams.select] : [];
      const missingFilterCols = filterCols.filter((col) => !selectCols.includes(col));
      if (missingFilterCols.length) {
        repairedParams.select = [...selectCols, ...missingFilterCols];
      }
    }
    if (JSON.stringify(repairedParams) !== JSON.stringify(step.params || {})) changed = true;
    return { ...step, params: repairedParams };
  });
  return { plan: clone, changed };
}

function collectBlockRefsFromPlan(plan: any): string[] {
  if (!plan || !Array.isArray(plan.steps)) return [];
  const refs: string[] = [];
  const pushRef = (value?: any) => {
    if (typeof value !== "string") return;
    const trimmed = value.trim();
    if (trimmed) refs.push(trimmed);
  };
  plan.steps.forEach((step: any) => {
    const params = step?.params || {};
    [
      params.source?.blockRef,
      params.target?.blockRef,
      params.left?.blockRef,
      params.right?.blockRef,
      params.dest?.blockRef,
      params.dest?.anchor?.blockRef,
      params.anchor?.blockRef,
    ].forEach(pushRef);
  });
  return refs;
}

function normalizeSheetName(sheet?: string): string {
  if (!sheet) return "";
  let name = sheet.trim();
  if (name.startsWith("'") && name.endsWith("'") && name.length >= 2) {
    name = name.slice(1, -1);
  }
  return name;
}

function collectDestSheetNames(plan: any): Set<string> {
  const sheets = new Set<string>();
  (plan?.steps || []).forEach((step: any) => {
    const params = step?.params || {};
    const dest = params.dest || params.output;
    const addSheet = (value?: any) => {
      if (typeof value === "string" && value.trim()) {
        sheets.add(normalizeSheetName(value));
      }
    };
    addSheet(dest?.sheetName);
    addSheet(dest?.sheet);
    addSheet(dest?.anchor?.sheetName);
  });
  return sheets;
}

function findInvalidBlockRefs(plan: any, allowed: Set<string>): string[] {
  if (!plan) return [];
  const refs = collectBlockRefsFromPlan(plan);
  const extraSheets = collectDestSheetNames(plan);
  return refs.filter((ref) => {
    if (allowed.has(ref)) return false;
    const sheet = normalizeSheetName(ref.split("!")[0]);
    if (sheet && extraSheets.has(sheet)) return false;
    return true;
  });
}

function planHash(obj: any) {
  try {
    const str = JSON.stringify(obj);
    let h = 0;
    for (let i = 0; i < str.length; i += 1) h = (Math.imul(31, h) + str.charCodeAt(i)) | 0;
    return `h${(h >>> 0).toString(16)}`;
  } catch {
    return "h0000";
  }
}

function logPlanTrace(plan: any, stage: string, logger?: (event: string, data?: any) => void) {
  try {
    const hash = planHash(plan);
    (plan?.steps || []).forEach((s: any) => {
      logger?.("plan_trace", {
        stage,
        hash,
        stepId: s?.id,
        macro: s?.macro,
        select: s?.params?.select,
        sort: s?.params?.sort,
        filter: s?.params?.filter,
      });
    });
  } catch {
    // ignore
  }
}

function sortMap(plan: any): Record<string, any> {
  const map: Record<string, any> = {};
  (plan?.steps || []).forEach((s: any) => {
    if (s?.params?.sort) map[s.id] = s.params.sort;
  });
  return map;
}

export function preSchemaRepair(plan: any, context?: any, logger?: (event: string, data?: any) => void) {
  if (!plan || typeof plan !== "object") return plan;
  const copy = JSON.parse(JSON.stringify(plan));
  let tvCount = 1;
  let wfCount = 1;
  (copy.steps || []).forEach((s: any) => {
    if (!s.id) {
      if (s.macro === "table_view") s.id = `tv${tvCount++}`;
      else if (s.macro === "write_formula") s.id = `calc${wfCount++}`;
      else s.id = `step${(tvCount + wfCount + 1)}`;
      logger?.("repair_added_step_id", { stepId: s.id });
    }
    // normalize source shape if {type:"block", address:"A1:B2"}
    const src = s.params?.source;
    if (src && typeof src === "object" && !src.blockRef && src.address) {
      s.params.source = { blockRef: src.address };
    }
    // filter value coercion
    if (Array.isArray(s.params?.filter)) {
      s.params.filter = s.params.filter.map((f: any) => {
        const out = { ...(f || {}) };
        if (out && out.type === "number" && typeof out.value === "string" && out.value.trim() !== "") {
          const n = Number(out.value);
          if (!Number.isNaN(n)) {
            out.value = n;
            logger?.("repair_coerced_filter_value", { stepId: s.id, col: out.col });
          }
        }
        if (typeof out.op === "string") {
          out.op = out.op.trim();
          logger?.("repair_normalized_filter_op", { stepId: s.id, col: out.col, op: out.op });
        }
        return out;
      });
    }
    // select repair
    if (s.macro === "table_view") {
      const destMode = s.params?.dest?.mode || "newSheet";
      const select = Array.isArray(s.params?.select) ? s.params.select : [];
      if (destMode === "inPlace") {
        if (select.length > 0) {
          s.params.select = [];
          logger?.("repair_added_select", { stepId: s.id, reason: "inPlace_removed_select" });
        }
      } else if (!select || select.length === 0) {
        const filters = Array.isArray(s.params?.filter) ? s.params.filter : [];
        const cols: string[] = [];
        const add = (c: any) => {
          if (c && typeof c === "string" && !cols.includes(c)) cols.push(c);
        };
        filters.forEach((f: any) => add(f?.col));
        if (cols.length) {
          s.params.select = cols;
          logger?.("repair_added_select", { stepId: s.id, reason: "missing_select", added: cols });
        } else {
          const fallback = context?.sheets?.[0]?.blocks?.[0]?.headers || [];
          const hasProjet = fallback.find((h: string) => (h || "").toLowerCase().includes("projet"));
          s.params.select = hasProjet ? [hasProjet] : fallback.slice(0, 1);
          if (!s.params.select || s.params.select.length === 0) s.params.select = ["Projet"];
          logger?.("repair_added_select", { stepId: s.id, reason: "missing_select_fallback", added: s.params.select });
        }
      }
    }
  });
  return copy;
}

async function tryParsePlan(rawText: string, context: any, userPrompt: string, warnings?: string[], trace?: PlanContextTrace) {
  let parsed = parseJson(rawText);
  if (!parsed.ok) {
    return { parsed, normalized: null as any, sanitized: null as any, parseError: "json_parse_error" };
  }
  // root allowlist + shape coercion
  const root = parsed.value || {};
  const cleaned: any = {
    version: root.version || "1.0",
    goal: root.goal || "User request",
    confirmations: Array.isArray(root.confirmations) ? root.confirmations : undefined,
  };
  cleaned.steps = Array.isArray(root.steps) ? root.steps.map((s: any) => {
    if (!s || typeof s !== "object") return s;
    const params = { ...(s.params || {}) };
    // migrate legacy
    if (typeof params.source === "string") params.source = { blockRef: params.source };
    if (typeof params.blockRef === "string") {
      params.source = params.source || {};
      params.source.blockRef = params.source.blockRef || params.blockRef;
      delete (params as any).blockRef;
    }
    if (params.output && !params.dest) {
      const out = params.output;
      const mode = out.mode || "newSheet";
      params.dest = { mode, sheetName: out.sheetName || out.sheet };
      delete (params as any).output;
    }
    return { id: s.id, macro: s.macro, params };
  }) : [];

  parsed = { ok: true, value: cleaned } as any;

  const rawPlan = (parsed as any).value;
  if (hasForbiddenKeys(rawPlan)) {
    return { parsed: { ok: false, error: "forbidden_keys" } as any, normalized: null as any, sanitized: null as any, parseError: "forbidden_keys" };
  }
  trace && (trace.rawPlan = rawPlan);
  const canonical = canonicalizePlan(rawPlan, warnings);
  trace && (trace.canonicalPlan = canonical);
  const normalized = normalizePlan(canonical, context, userPrompt, warnings);
  const extraHeaders: Record<string, string[]> = {};
  (normalized?.steps || []).forEach((s: any) => {
    if (s?.macro === "write_formula" && s.params?.target?.artifactRef && s.params?.target?.headerName) {
      const list = extraHeaders[s.params.target.artifactRef] || [];
      list.push(s.params.target.headerName);
      extraHeaders[s.params.target.artifactRef] = list;
    }
    if (s?.macro === "write_formula" && s.params?.target?.blockRef && s.params?.target?.headerName) {
      const list = extraHeaders[s.params.target.blockRef] || [];
      list.push(s.params.target.headerName);
      extraHeaders[s.params.target.blockRef] = list;
    }
  });
  const sanitizedRaw = sanitizePlan(normalized, context, warnings, extraHeaders, userPrompt);
  if (!sanitizedRaw || !Array.isArray(sanitizedRaw.steps) || sanitizedRaw.steps.length === 0) {
    return {
      parsed: { ok: false, error: "sanitize_failed_unrepairable" } as any,
      normalized: null as any,
      sanitized: null as any,
      parseError: "sanitize_failed_unrepairable",
    };
  }
  const sanitized = canonicalizePlan(sanitizedRaw);
  trace && (trace.sanitizedPlan = sanitized);
  return { parsed, normalized, sanitized, parseError: null };
}

export async function planWithOllama(args: {
  context: any;
  userPrompt: string;
  model?: string;
  host?: string;
  timeoutMs?: number;
  logger?: (event: string, data?: any) => void;
}): Promise<PlannerResult> {
  const ollamaDisabled =
    (typeof process !== "undefined" && (process as any).env?.OLLAMA_DISABLED === "1") ||
    (typeof window !== "undefined" && (window as any)?.OLLAMA_DISABLED === "1");
  if (ollamaDisabled) {
    throw new Error("Ollama disabled in unit tests");
  }
  const host = args.host || DEFAULT_HOST;
  const model = args.model || DEFAULT_MODEL;
  const timeout = args.timeoutMs || DEFAULT_TIMEOUT;
  const PARSE_FAIL_MESSAGE = "Je n'ai pas pu structurer le plan. Reformule ta demande simplement, exemple : « Filtre pour ne garder que les dates supérieures à 2027 ».";
  const logEvent = (event: string, data?: any) => {
    try {
      args.logger?.(event, data);
    } catch {
      // ignore logger failures
    }
  };

  const allowedBlockIds = (args.context?.sheets || [])
    .flatMap((s: any) => (s.blocks || []).map((b: any) => b.id))
    .filter((id: string | undefined): id is string => typeof id === "string" && id.length > 0);
  const allowedBlockSet = new Set<string>(allowedBlockIds);
  const tableHints = (args.context?.sheets || [])
    .flatMap((s: any) => (s.tables || []).map((t: any) => `${t.name}@${s.name || "Sheet"}`))
    .filter(Boolean);
  const basePrompt = `Context JSON:\n${JSON.stringify(args.context)}\nTables (prefer tableName + sheetName over blockRef): ${
    tableHints.length ? tableHints.join(", ") : "none"
  }\nAllowed blockRef values (must use only these): ${allowedBlockIds.join(", ")}\nUser request:\n${args.userPrompt}\nProduce a plan JSON.`;

  let sanitizeNotes: string[] = [];
  let sanitizeWarnings: string[] = [];
  let invariantIssues: any[] = [];
  let rawText = "";
  let rawTextRetry: string | undefined;
  let finalParseError: string | null = null;
  try {
    rawText = await callOllama(basePrompt, model, host, timeout);
  } catch (err: any) {
    const msg = err?.name === "AbortError" ? "Ollama request timed out" : err?.message || String(err);
    const hostMsg = msg.includes("fetch failed") || msg.includes("ECONNREFUSED") ? `Ollama not reachable at ${host}` : msg;
    return { status: "error", error: hostMsg };
  }

  const trace: PlanContextTrace = {};
  let { parsed, normalized, sanitized, parseError } = await tryParsePlan(rawText, args.context, args.userPrompt, sanitizeWarnings, trace);
  sanitizeNotes = [...sanitizeWarnings];
  finalParseError = parseError || finalParseError;
  const invalidRefs = parsed?.value ? findInvalidBlockRefs(parsed.value, allowedBlockSet) : [];
  if (invalidRefs.length) {
    const failureStage = "block_ref_invalid";
    return {
      status: "invalid_plan",
      rawText,
      rawTextRetry,
      errors: invalidRefs.map((ref) => `blockref_invalid:${ref}`),
      failureStage,
      sanitizeNotes,
      parseError: failureStage,
    };
  }
  const unsupportedOps = sanitizeWarnings.filter((w) => w.startsWith("unsupported_filter_op"));
  if (unsupportedOps.length) {
    return { status: "invalid_plan", rawText, rawTextRetry, errors: unsupportedOps, failureStage: "filter_op_unsupported", sanitizeNotes };
  }
  const rawSortMap = sortMap(trace.rawPlan || parsed.value || {});
  const throwIfSortDropped = (_stagePlan: any, _stage: string) => {};
  if (!parsed.ok) {
    const repairPrompt = `Return ONLY corrected JSON object, no prose. Errors: ${parsed.error}.`;
    try {
      rawTextRetry = await callOllama(repairPrompt, model, host, timeout);
      const retryWarnings: string[] = [];
      const retryTrace: PlanContextTrace = {};
      const retry = await tryParsePlan(rawTextRetry, args.context, args.userPrompt, retryWarnings, retryTrace);
      Object.assign(trace, retryTrace);
      parsed = retry.parsed;
      normalized = retry.normalized;
      sanitized = retry.sanitized;
      sanitizeNotes = [...retryWarnings];
      sanitizeWarnings = retryWarnings;
      finalParseError = retry.parseError || finalParseError;
      if (retry.parsed.ok && retry.sanitized) {
        rawText = rawTextRetry;
      } else {
        return { status: "invalid_plan", rawText, rawTextRetry, errors: [PARSE_FAIL_MESSAGE], failureStage: "parse_failed", sanitizeNotes };
      }
    } catch (err: any) {
      return { status: "invalid_plan", rawText, rawTextRetry, errors: [PARSE_FAIL_MESSAGE], failureStage: "parse_failed", sanitizeNotes };
    }
  }

  logPlanTrace(trace.rawPlan, "raw", logEvent);
  logPlanTrace(trace.canonicalPlan, "canonical", logEvent);
  logPlanTrace(sanitized, "sanitized", logEvent);
  try {
    throwIfSortDropped(sanitized, "sanitized");
  } catch (err: any) {
    return { status: "invalid_plan", rawText, rawTextRetry, errors: [err.message], failureStage: "sort_dropped", sanitizeNotes };
  }

  if (!parsed.ok || !sanitized) {
    const stage = finalParseError || (!parsed.ok && parsed.error) || "invalid_plan";
    const hardFail = stage === "forbidden_keys" || stage === "sanitize_failed_unrepairable";
    if (hardFail) {
      return { status: "invalid_plan", rawText, rawTextRetry, errors: [stage], failureStage: stage, sanitizeNotes };
    }
    const friendly = stage === "json_parse_error" ? PARSE_FAIL_MESSAGE : stage;
    return { status: "invalid_plan", rawText, rawTextRetry, errors: [friendly], failureStage: stage, sanitizeNotes };
  }

  sanitized = preSchemaRepair(sanitized, args.context, logEvent);
  sanitized.version = sanitized.version || "1.0";
  sanitized.goal = sanitized.goal || args.userPrompt || "User request";

  const validation = validatePlan(sanitized);
  if (validation.valid) {
    const candidatePlan = sanitized;
    logPlanTrace(candidatePlan, "sanitized", logEvent);
    try {
      throwIfSortDropped(candidatePlan, "sanitized");
    } catch (err: any) {
      return { status: "invalid_plan", rawText, rawTextRetry, errors: [err.message], failureStage: "sort_dropped", sanitizeNotes };
    }
    const inv = validatePlanInvariants(candidatePlan as any);
    if (inv.valid) {
      const finalRaw = rawTextRetry || rawText;
      const sanitizedSteps = sanitized?.steps || [];
      const promptText = (args.userPrompt || "").toLowerCase();
      const promptRequiresFilter = /filtr|filter/.test(promptText);
      const hasFilterStep = sanitizedSteps.some(
        (step: any) => step?.macro === "table_view" && Array.isArray(step.params?.filter) && step.params.filter.length > 0
      );
      if (promptRequiresFilter && !hasFilterStep) {
        return { status: "invalid_plan", rawText: finalRaw, rawTextRetry, errors: ["filter_missing"], failureStage: "filter_missing", sanitizeNotes };
      }
      trace.validatedPlan = candidatePlan;
      logEvent("plan_stage", { stage: "validated", hash: planHash(candidatePlan) });
      logPlanTrace(candidatePlan, "validated", logEvent);
      return {
        status: "ok",
        plan: candidatePlan,
        rawText: finalRaw,
        rawTextRetry,
        fallbackUsed: false,
        sanitizeNotes,
        deterministicRepairApplied: false,
      };
    }
    invariantIssues = inv.issues;
    logEvent("plan_invalid", { issues: invariantIssues, stage: "sanitized", hash: planHash(candidatePlan) });
  }

  const validationErrors = validation.valid ? [] : validation.errors || [];
  let repairErrors: string[] = [...validationErrors];
  if (invariantIssues.length) {
    repairErrors = invariantIssues.map((i: any) => i.message || i.code || "plan_invariant_failed");
  }
  const needRetry = repairErrors.length > 0;
  const hasTableViewStep = Array.isArray(sanitized?.steps) && sanitized.steps.some((s: any) => s?.macro === "table_view");
  const hasTableViewError = repairErrors.some((e: string) => e.includes("table_view") || e.includes("filter") || e.includes("select") || e.includes("sort"));
  if (hasTableViewError || hasTableViewStep) {
    const headerHint =
      (args.context?.sheets || [])
        .flatMap((s: any) => (s.blocks || []).flatMap((b: any) => b.headers || []))
        .filter((h: any) => typeof h === "string")
        .slice(0, 8)
        .join(", ") || "none";
    repairErrors = Array.from(
      new Set([
        ...repairErrors,
        `table_view requires: inPlace(no select) or view(select required). Headers: ${headerHint}`,
      ])
    );
    // if a column was just created and should be selected
    const auto = autoRepairFromIntent(sanitized, args.context, sanitizeNotes);
    if (auto.changed) {
      sanitized = auto.plan;
      sanitizeNotes = Array.from(new Set([...sanitizeNotes, "table_view_auto_repair"]));
      const validationAuto = validatePlan(sanitized);
      if (validationAuto.valid) {
        const invAuto = validatePlanInvariants(sanitized as any);
        if (invAuto.valid) {
          const finalRaw = rawTextRetry || rawText;
          return { status: "ok", plan: sanitized, rawText: finalRaw, rawTextRetry, fallbackUsed: true, sanitizeNotes, deterministicRepairApplied: true };
        }
        repairErrors = invAuto.issues.map((i: any) => i.message || i.code);
      } else {
        repairErrors = validationAuto.errors || repairErrors;
      }
    }
  }

  // one repair attempt if schema invalid
  const promptReason = invariantIssues.length ? "Your plan violated required invariants." : "Your JSON was invalid per schema.";
  const repairPrompt = `${promptReason} Fix it. Return ONLY corrected JSON.\nErrors: ${repairErrors.join(
    "; "
  )}\nPrevious JSON:\n${JSON.stringify(parsed.value)}`;
  if (!needRetry) {
    logEvent("planner_retry_failed", { issues: invariantIssues, stage: "sanitized", hash: planHash(sanitized) });
    return { status: "invalid_plan", rawText, rawTextRetry, errors: repairErrors, failureStage: "invariants_failed", sanitizeNotes };
  }

  try {
    logEvent("planner_retry", { errors: repairErrors, stage: "retry", hash: planHash(parsed.value) });
    rawTextRetry = await callOllama(repairPrompt, model, host, timeout);
    const retrySanitizeWarnings: string[] = [];
    const retry = await tryParsePlan(rawTextRetry, args.context, args.userPrompt, retrySanitizeWarnings);
    sanitizeNotes = retrySanitizeWarnings;
    if (retry.parsed.ok) {
      retry.sanitized = preSchemaRepair(retry.sanitized, args.context);
      if (retry.sanitized) {
        retry.sanitized.version = retry.sanitized.version || "1.0";
        retry.sanitized.goal = retry.sanitized.goal || args.userPrompt || "User request";
      }
      const validation2 = validatePlan(retry.sanitized);
      if (validation2.valid) {
        const inv2 = validatePlanInvariants(retry.sanitized as any);
        if (inv2.valid) {
          return { status: "ok", plan: retry.sanitized, rawText: rawTextRetry, rawTextRetry, fallbackUsed: false, sanitizeNotes };
        }
        logEvent("planner_retry_failed", { issues: inv2.issues });
        return { status: "invalid_plan", rawText, rawTextRetry, errors: inv2.issues.map((i) => i.message || i.code), failureStage: "invariants_failed", sanitizeNotes };
      }
      return { status: "invalid_plan", rawText, rawTextRetry, errors: validation2.errors || [], failureStage: "schema_invalid", sanitizeNotes };
    }
    return { status: "invalid_plan", rawText, rawTextRetry, errors: ["repair_failed"], failureStage: "repair_failed", sanitizeNotes };
  } catch (err: any) {
    return { status: "error", error: err?.message || String(err), rawText, rawTextRetry, failureStage: "repair_failed", sanitizeNotes };
  }
}
