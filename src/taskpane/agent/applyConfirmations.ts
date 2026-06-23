import { WorkbookContextSnapshot } from "../context/types";
import { AgentPlan, PlanConfirmation, ValidateDataInternalState } from "./types";
import { normalizeHeader } from "./normalizeHeader";
import { normalizePlan } from "./planner/normalizePlan";
import { findBlock } from "./utils";
import { VALIDATE_DATA_QUESTIONS } from "./validateDataFlow";

function clone<T>(obj: T): T {
  return JSON.parse(JSON.stringify(obj));
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

function extractQuotedValue(question?: string): string | null {
  if (!question) return null;
  const match = question.match(/['"“”‘’‹›«»](.+?)['"“”‘’‹›«»]/);
  return match?.[1] ?? null;
}

function normalizeBracketToken(token?: string): string | null {
  if (!token) return null;
  const trimmed = token.trim();
  const stripped = trimmed.startsWith("@") ? trimmed.slice(1).trim() : trimmed;
  if (!stripped) return null;
  return normalizeHeader(stripped);
}

function tokensFromFormula(formula: string): string[] {
  if (!formula) return [];
  const tokens: string[] = [];
  const regex = /\[([^\[\]]+)\]/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(formula)) !== null) {
    const token = match[1];
    if (token) {
      const trimmed = token.trim();
      if (trimmed) tokens.push(trimmed);
    }
  }
  return tokens;
}

function escapeRegExp(value: string): string {
  return value.replace(/[-[\]/{}()*+?.\\^$|]/g, "\\$&");
}

function resolveBlockMetadata(target: any, context: WorkbookContextSnapshot) {
  if (!target) return { headers: [] as string[], kind: undefined };
  if (typeof target.blockRef === "string" && target.blockRef) {
    const { block } = findBlock(target.blockRef, context);
    if (block) {
      return {
        headers: Array.isArray(block.headers) ? block.headers : [],
        kind: block.kind,
      };
    }
  }
  if (typeof target.tableName === "string" && target.tableName) {
    const lookup = target.tableName.toLowerCase();
    for (const sheet of context?.sheets || []) {
      const match = (sheet.tables || []).find((tbl) => (tbl?.name || "").toLowerCase() === lookup);
      if (match) {
        return {
          headers: Array.isArray(match.headers) ? match.headers : [],
          kind: "table",
        };
      }
    }
  }
  return { headers: [] as string[], kind: undefined };
}

export function applyConfirmationsToPlan(plan: AgentPlan, answers: Record<string, string>, context: WorkbookContextSnapshot): AgentPlan {
  const draft = clone(plan);
  const normalized = normalizePlan(draft, context);
  const normalizedConfirmations = (normalized.confirmations || []) as PlanConfirmation[];
  const confirmationMap = new Map<string, PlanConfirmation>();
  normalizedConfirmations.forEach((conf) => {
    if (conf?.id) confirmationMap.set(conf.id, conf);
  });
  const formulaConfirmations = normalizedConfirmations.filter(
    (conf) => typeof conf?.id === "string" && conf.id.startsWith("formula_column")
  );
  const handledFormulaConfirmations = new Set<string>();
  const blocks = (context?.sheets || []).flatMap((s: any) => s.blocks || []);
  const fallbackBlock = context?.active?.selectionInBlockId || context?.active?.nearestBlockId || blocks[0]?.id;

  const resolveBlockRef = (ref?: string) => {
    if (ref && blocks.some((b: any) => b.id === ref)) return ref;
    return fallbackBlock;
  };

  draft.steps = normalized.steps.map((step: any) => {
    // apply table_view sort confirmations
    const sortColAnswer = answers[`${step.id}:sort_col`];
    const sortDirAnswer = answers[`${step.id}:sort_dir`];
    if (step.macro === "table_view" && (sortColAnswer || sortDirAnswer)) {
      const nextSort = {
        ...(step.params?.sort || {}),
      };
      if (sortColAnswer) nextSort.col = sortColAnswer;
      if (sortDirAnswer) nextSort.dir = sortDirAnswer;
      if (!nextSort.dir) nextSort.dir = "desc";
      const confirmId = `${step.id}:sort_col`;
      const confirmation = confirmationMap.get(confirmId);
      if (confirmation && sortColAnswer) {
        const choice = (confirmation.choices || []).find(
          (c: any) => c.id === sortColAnswer || c.label === sortColAnswer
        );
        nextSort.col = choice?.label || sortColAnswer;
      }
      step = {
        ...step,
        params: { ...step.params, sort: nextSort },
      };
    }

    // apply table_view filter confirmations (single filter slot)
    const filterCol = answers[`${step.id}:filter_col`];
    const filterOp = answers[`${step.id}:filter_op`];
    const filterValue = answers[`${step.id}:filter_value`];
    if (step.macro === "table_view" && (filterCol || filterOp || typeof filterValue !== "undefined")) {
      const existing = Array.isArray(step.params?.filter) && step.params.filter.length ? step.params.filter[0] : {};
      const nextFilter = {
        ...existing,
      };
      if (filterCol) nextFilter.col = filterCol;
      if (filterOp) nextFilter.op = filterOp;
      if (typeof filterValue !== "undefined") nextFilter.value = filterValue;
      const confirmId = `${step.id}:filter_col`;
      const confirmation = confirmationMap.get(confirmId);
      if (confirmation && filterCol) {
        const choice = (confirmation.choices || []).find(
          (c: any) => c.id === filterCol || c.label === filterCol
        );
        nextFilter.col = choice?.label || filterCol;
      }
      step = {
        ...step,
        params: { ...step.params, filter: [nextFilter] },
      };
    }

    if (step.macro === "join_tables") {
      const joinKeyAnswer = answers[`joinKey:${step.id}`] ?? answers[`${step.id}:join_key`];
      const resolved = decodeJoinKeyChoice(joinKeyAnswer);
      if (resolved) {
        const strategy = step.params?.keys?.[0]?.strategy || "case_insensitive_trim";
        step = {
          ...step,
          params: {
            ...step.params,
            keys: [{ left: resolved.left, right: resolved.right, strategy }],
          },
        };
      }
    }

    if (step.macro === "create_chart") {
      const dest = step.params?.dest || {};
      const answerNoNewSheet = Object.values(answers).includes("no") || answers["conf1"] === "no";
      if (answerNoNewSheet && dest.mode === "newSheet") {
        dest.mode = "right";
        dest.anchor = { blockRef: resolveBlockRef(step.params?.source?.blockRef) };
      }
      step = { ...step, params: { ...step.params, dest } };
    }
    if (step.macro === "write_formula" && formulaConfirmations.length) {
      let formulaText = typeof step.params?.formula === "string" ? step.params.formula : "";
      if (formulaText) {
        const target = (step.params || {}).target || {};
        const blockInfo = resolveBlockMetadata(target, context);
        const headerSet = new Set(
          (blockInfo.headers || [])
            .map((hdr) => normalizeHeader(hdr))
            .filter((hdr): hdr is string => typeof hdr === "string" && hdr.length > 0)
        );
        for (const confirmation of formulaConfirmations) {
          if (!confirmation.id || handledFormulaConfirmations.has(confirmation.id)) continue;
          const decision = answers[confirmation.id];
          if (!decision) continue;
          const tokens = tokensFromFormula(formulaText);
          const missingTokens = tokens.filter((token) => {
            const normalized = normalizeBracketToken(token);
            return !!normalized && !headerSet.has(normalized);
          });
          let placeholderCandidate: string | null = null;
          const questionPlaceholder = extractQuotedValue(confirmation.question);
          if (questionPlaceholder) {
            const normalizedQuestion = normalizeBracketToken(questionPlaceholder);
            if (normalizedQuestion) {
              placeholderCandidate = tokens.find(
                (token) => normalizeBracketToken(token) === normalizedQuestion
              ) || null;
            }
          }
          if (!placeholderCandidate && missingTokens.length) {
            placeholderCandidate = missingTokens[0] ?? null;
          }
          if (!placeholderCandidate) continue;
          const placeholderNorm = normalizeBracketToken(placeholderCandidate);
          if (!placeholderNorm || headerSet.has(placeholderNorm)) continue;
          const choice = (confirmation.choices || []).find(
            (c: any) => c.id === decision || c.label === decision
          );
          const choiceLabel = (choice?.label || decision || "").trim();
          const normalizedChoice = normalizeHeader(choiceLabel);
          if (!normalizedChoice) continue;
          const headerMatch = (blockInfo.headers || []).find(
            (hdr) => normalizeHeader(hdr) === normalizedChoice
          );
          if (!headerMatch) continue;
          const isTableBlock = blockInfo.kind === "table";
          const replacement = isTableBlock ? `[@${headerMatch}]` : `[${headerMatch}]`;
          const quotedPlaceholder = placeholderCandidate.replace(/^@/, "").trim();
          if (!quotedPlaceholder) continue;
          const pattern = new RegExp(`\\[\\s*@?${escapeRegExp(quotedPlaceholder)}\\s*\\]`, "gi");
          const updatedFormula = formulaText.replace(pattern, replacement);
          if (updatedFormula === formulaText) continue;
          formulaText = updatedFormula;
          step = { ...step, params: { ...step.params, formula: formulaText } };
          handledFormulaConfirmations.add(confirmation.id);
        }
      }
    }
    if (step.macro === "validate_data") {
      const confirmationFlow = VALIDATE_DATA_QUESTIONS.map((question) => ({
        id: question.id,
        decisionKey: question.decisionKey,
      }));
      const existingInternal = (step.params?.__internal as ValidateDataInternalState) || {};
      const decisions = { ...(existingInternal.decisions || {}) };
      let hasDecision = false;
      confirmationFlow.forEach((question) => {
        const answer = answers[question.id];
        if (typeof answer === "string") {
          const normalized = answer.trim().toLowerCase();
          if (normalized === "yes" || normalized === "no") {
            decisions[question.decisionKey as keyof typeof decisions] = normalized === "yes";
            hasDecision = true;
          }
        }
      });
      if (hasDecision) {
        const nextExecutions = Math.min(4, (existingInternal.executions || 0) + 1);
        const nextInternal: ValidateDataInternalState = {
          ...existingInternal,
          decisions,
          phase: "confirmations",
          executions: nextExecutions,
        };
        step = {
          ...step,
          params: {
            ...step.params,
            __internal: nextInternal,
          },
        };
      }
    }
    return step;
  });

  return draft as AgentPlan;
}
