import { WorkbookContextSnapshot } from "../context/types";
import { AgentPlan, PlanConfirmationChoice } from "./types";
import { normalizeHeader } from "./normalizeHeader";
import { VALIDATE_DATA_CONFIRM_IDS } from "./validateDataFlow";

export type AutoAnswerMode = "demoEval" | "interactive" | "none";

type AutoAnswerResult = {
  plan: AgentPlan;
  decisions: Record<string, string>;
  autoAnswered: number;
  missingBefore: number;
  warnings: string[];
};

function headersForBlock(context: WorkbookContextSnapshot, blockRef?: string): string[] {
  if (!blockRef) return [];
  const sheetName = blockRef.includes("!") ? blockRef.split("!")[0] : blockRef.split(":")[0];
  const sheet = context.sheets.find((s) => s.name === sheetName);
  const block = sheet?.blocks.find((b) => b.id === blockRef || `${sheetName}!${b.address}` === blockRef);
  return block?.headers || [];
}

function normalize(str?: string) {
  const lower = (str || "").toLowerCase();
  let out = "";
  for (const ch of lower) {
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r" || ch === "\f") continue;
    out += ch;
  }
  return out;
}

function isCodeLike(h: string) {
  const n = normalize(h);
  return n.includes("code") || n.includes("id");
}

function isLabelLike(h: string) {
  const n = normalize(h);
  return n.includes("nom") || n.includes("name") || n.includes("label") || n.includes("libell") || n.includes("region");
}

function normalizeQuestion(text: string) {
  return (text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function matchChoiceByPrompt(choices: PlanConfirmationChoice[], promptNorm: string) {
  if (!promptNorm) return null;
  return choices.find((choice) => {
    const label = normalizeQuestion(choice.label || "");
    return label.length > 0 && promptNorm.includes(label);
  }) || null;
}

function determineSortDirection(promptNorm: string): string | undefined {
  if (!promptNorm) return undefined;
  if (promptNorm.includes("croissant") || promptNorm.includes("asc") || promptNorm.includes("ascending")) return "asc";
  if (promptNorm.includes("decroissant") || promptNorm.includes("desc") || promptNorm.includes("descending")) return "desc";
  return undefined;
}

function isSafeAutoQuestion(text: string) {
  const norm = normalizeQuestion(text);
  return norm.includes("key") || norm.includes("cle") || norm.includes("join") || norm.includes("colonne") || norm.includes("column");
}

function bestChoice(choices: PlanConfirmationChoice[], headers: string[]): string {
  const safeChoices = choices.filter(
    (c) => !(c.id || "").toLowerCase().includes("abort") && !(c.label || "").toLowerCase().includes("annul")
  );
  const pool = safeChoices.length > 0 ? safeChoices : choices;
  let best = pool[0];
  let bestScore = -1;
  pool.forEach((c) => {
    const label = (c.label || "").toLowerCase();
    const score = headers.reduce((acc, h) => (label.includes(h.toLowerCase()) ? acc + 1 : acc), 0);
    if (score > bestScore) {
      bestScore = score;
      best = c;
    }
  });
  return best?.id || pool[0]?.id || choices[0]?.id || "";
}

function decodeJoinKeyChoice(choiceId?: string): { left: string; right: string } | null {
  if (!choiceId || typeof choiceId !== "string") return null;
  const trimmed = choiceId.trim();
  const sep = trimmed.indexOf("|");
  if (sep <= 0 || sep === trimmed.length - 1) return null;
  const left = trimmed.slice(0, sep).trim();
  const right = trimmed.slice(sep + 1).trim();
  if (!left || !right) return null;
  return { left, right };
}

function buildNormalizedHeaderSet(headers: string[]): Set<string> {
  const set = new Set<string>();
  headers.forEach((header) => {
    const normalized = normalizeHeader(header);
    if (normalized) set.add(normalized);
  });
  return set;
}

function headerExists(set: Set<string>, value?: string): boolean {
  if (!value) return false;
  const normalized = normalizeHeader(value);
  return !!normalized && set.has(normalized);
}

function findMatchingJoinKeyChoice(
  choices: PlanConfirmationChoice[],
  leftHeaders: string[],
  rightHeaders: string[]
): PlanConfirmationChoice | undefined {
  const leftSet = buildNormalizedHeaderSet(leftHeaders);
  const rightSet = buildNormalizedHeaderSet(rightHeaders);
  for (const choice of choices) {
    const decoded = decodeJoinKeyChoice(choice.id);
    if (!decoded) continue;
    if (headerExists(leftSet, decoded.left) && headerExists(rightSet, decoded.right)) {
      return choice;
    }
  }
  return undefined;
}

export function autoAnswerConfirmations(
  plan: AgentPlan,
  context: WorkbookContextSnapshot,
  mode: AutoAnswerMode
): AutoAnswerResult {
  if (mode === "none") return { plan, decisions: {}, autoAnswered: 0, missingBefore: plan.confirmations?.length || 0, warnings: [] };
  const confirmations = (plan.confirmations || []).filter((c) => !VALIDATE_DATA_CONFIRM_IDS.has(c.id));
  const decisions: Record<string, string> = {};
  const warnings: string[] = [];
  let autoAnswered = 0;
  const missingBefore = confirmations.filter((c) => c.required).length;

  const joinStep = plan.steps.find((s) => s.macro === "join_tables");
  const joinParams: any = joinStep?.params || {};
  const leftHeaders = headersForBlock(context, joinParams.left?.blockRef);
  const rightHeaders = headersForBlock(context, joinParams.right?.blockRef);
  const keysValid =
    Array.isArray(joinParams.keys) &&
    joinParams.keys.length > 0 &&
    joinParams.keys.every(
      (k: any) => leftHeaders.map((h) => h.toLowerCase()).includes((k.left || "").toLowerCase()) && rightHeaders.map((h) => h.toLowerCase()).includes((k.right || "").toLowerCase())
    );
  const selectRightNonEmpty = Array.isArray(joinParams.select?.right?.columns) && joinParams.select.right.columns.length > 0;

  const promptGlobal = normalizeQuestion(plan.goal || "");
  const filtered = confirmations.filter((c) => {
    const q = c.question.toLowerCase();
    // drop confirmations if already satisfied
    if (q.includes("join type") && joinParams.joinType) return false;
    if (q.includes("key") && keysValid) return false;
    if ((q.includes("colonne") || q.includes("columns")) && selectRightNonEmpty) return false;
    return true;
  });

  const joinKeyDebug =
    typeof process !== "undefined" && !!((process as any).env?.AUTO_ANSWER_JOINKEY_DEBUG);

  filtered.forEach((c) => {
    const safe = isSafeAutoQuestion(c.question);
    if (mode === "interactive" && !safe) return;
    const columnQuestionSuffixes = [":sort_col", ":filter_col", ":unknown_header"];
    const isColumnQuestion = columnQuestionSuffixes.some((suffix) => c.id?.endsWith?.(suffix));
    const isJoinKey = typeof c.id === "string" && c.id.startsWith("joinKey:");
    let choiceId = "";
    let resolvedChoice: PlanConfirmationChoice | undefined;
    if (isJoinKey) {
      const match = findMatchingJoinKeyChoice(c.choices, leftHeaders, rightHeaders);
      if (!match) return;
      choiceId = match.id;
      resolvedChoice = match;
      if (joinKeyDebug) {
        console.debug?.("autoAnswer joinKey", {
          question: c.question,
          choiceId,
          headersLeft: leftHeaders,
          headersRight: rightHeaders,
        });
      }
    } else {
      const headersAll = [...leftHeaders, ...rightHeaders];
      const questionNorm = normalizeQuestion(c.question);
      if (questionNorm.includes("fallback") && questionNorm.includes("cle")) {
        const preferred = c.choices.find((ch) => normalizeQuestion(ch.id).includes("use_fallback") || normalizeQuestion(ch.label).includes("fallback"));
        if (preferred) choiceId = preferred.id;
      }
      if (!choiceId && questionNorm.includes("abort")) {
        const nonAbort = c.choices.find((ch) => !normalizeQuestion(ch.id).includes("abort") && !normalizeQuestion(ch.label).includes("annul"));
        if (nonAbort) choiceId = nonAbort.id;
      }
      if (!choiceId && isColumnQuestion) {
        const promptMatch = matchChoiceByPrompt(c.choices, promptGlobal);
        if (promptMatch) choiceId = promptMatch.id;
      }
      if (!choiceId) choiceId = bestChoice(c.choices, headersAll);
    }
    if (choiceId && !resolvedChoice) {
      resolvedChoice =
        c.choices?.find((ch) => ch.id === choiceId) ?? c.choices?.find((ch) => ch.label === choiceId);
    }
    if (choiceId) {
      const decisionValue = isColumnQuestion && resolvedChoice?.label ? resolvedChoice.label : choiceId;
      decisions[c.id] = decisionValue;
      autoAnswered += 1;
      if (c.id.endsWith(":sort_col")) {
        const sortDirKey = `${c.id.replace(/:sort_col$/, "")}:sort_dir`;
        const dir = determineSortDirection(promptGlobal);
        if (dir) {
          decisions[sortDirKey] = dir;
          autoAnswered += 1;
        }
      }
    }
  });

  const newPlan = { ...plan, confirmations: filtered.filter((c) => !decisions[c.id]) };
  return { plan: newPlan as AgentPlan, decisions, autoAnswered, missingBefore, warnings };
}
