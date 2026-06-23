import { WorkbookContextSnapshot } from "../../context/types";
import { AgentPlan, ArtifactRecord, ConfirmationRequest, ExecutionResult, ExecutionStatus } from "../types";
import { applyConfirmationsToPlan } from "../applyConfirmations";
import { autoAnswerConfirmations, AutoAnswerMode } from "../autoConfirm";
import { canonicalizePlan } from "../canonicalizePlan";
import { executePlan } from "../executor";
import { validatePlan } from "../planSchema";
import { validatePlanInvariants } from "../planInvariants";
import { normalizePlan } from "../planner/normalizePlan";
import { preSchemaRepair, planWithOllama, PlannerResult } from "../planner/ollamaPlanner";
import { sanitizePlan } from "../planner/sanitizePlan";
import { collectExtraHeaders } from "../planner/extraHeaders";
import { RepairPlanOptions, repairPlanWriteFormulas } from "../planRepairer";

export type ExcelAdapter = {
  run<T>(cb: (ctx: Excel.RequestContext) => Promise<T>): Promise<T>;
};

export interface RunAgentPipelineOptions {
  context: WorkbookContextSnapshot;
  excelAdapter: ExcelAdapter;
  prompt?: string;
  plan?: AgentPlan;
  plannerModel?: string;
  plannerHost?: string;
  plannerTimeoutMs?: number;
  plannerLogger?: (event: string, data?: any) => void;
  planWithOllamaClient?: typeof planWithOllama;
  autoAnswerMode?: AutoAnswerMode;
  decisions?: Record<string, string>;
  confirmationDecisionHandler?: (confirmations: ConfirmationRequest[]) => Record<string, string>;
  preConfirmDecisionHandler?: (plan: AgentPlan, context: WorkbookContextSnapshot, prompt?: string) => Record<string, string>;
  formulaRepair?: {
    enabled?: boolean;
    options?: RepairPlanOptions;
    hook?: (plan: AgentPlan, context: WorkbookContextSnapshot, userPrompt: string) => Promise<{ plan: AgentPlan; notes?: string[] }>;
  };
  maxAttempts?: number;
  initialArtifacts?: ArtifactRecord[];
}

export interface RunAgentPipelineResult {
  failureStage?: string;
  plannerStage?: {
    status: PlannerResult["status"];
    plan?: AgentPlan;
    rawText?: string | null;
    rawTextRetry?: string | null;
    errors?: any[];
    failureStage?: string | null;
    fallbackUsed?: boolean;
    deterministicRepairApplied?: boolean;
    sanitizeNotes?: string[];
    parseError?: string | null;
    planParsed?: any;
  };
  plan?: AgentPlan | null;
  planNormalized?: AgentPlan | null;
  planSanitized?: AgentPlan | null;
  planFinalExecuted?: AgentPlan | null;
  planRawText?: string | null;
  planRawTextRetry?: string | null;
  planParsed?: any;
  plannerRetryUsed: boolean;
  plannerFallbackUsed: boolean;
  plannerOutputWasNonJson: boolean;
  plannerSanitizeChangedPlan: boolean;
  plannerParseError: string | null;
  plannerSanitizeNotes: string[];
  plannerRetryReason: string | null;
  warnings: string[];
  execution: ExecutionResult;
  decisions: Record<string, string>;
  attempts: number;
  autoAnswerStats: {
    autoAnswered: number;
    missingBefore: number;
    warnings: string[];
  };
  confirmationsRequested?: ExecutionResult["confirmationsRequested"];
}

type FailureParams = {
  errors: string[];
  status?: ExecutionStatus;
  warnings: string[];
  failureStage?: string;
  plannerStage?: RunAgentPipelineResult["plannerStage"];
  planCandidate?: AgentPlan | null;
  planNormalized?: AgentPlan | null;
  planSanitized?: AgentPlan | null;
  planParsed?: any;
  planRawText?: string | null;
  planRawTextRetry?: string | null;
  plannerRetryUsed: boolean;
  plannerFallbackUsed: boolean;
  plannerOutputWasNonJson: boolean;
  plannerSanitizeChangedPlan: boolean;
  plannerParseError: string | null;
  plannerSanitizeNotes: string[];
  plannerRetryReason: string | null;
  decisions?: Record<string, string>;
};

function cloneValue<T>(value: T): T {
  return JSON.parse(JSON.stringify(value));
}

function safeParsePlan(text?: string | null): any | null {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function uniqueStrings(values: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  values.forEach((item) => {
    if (!item) return;
    if (!seen.has(item)) {
      seen.add(item);
      out.push(item);
    }
  });
  return out;
}

function safeFillDecisions(confirmations: ConfirmationRequest[]): Record<string, string> {
  const out: Record<string, string> = {};
  confirmations.forEach((conf) => {
    if (!conf || !Array.isArray(conf.choices) || !conf.choices.length) return;
    const safeChoice = conf.choices.find((choice) => {
      const choiceId = (choice.id || "").toLowerCase();
      const label = (choice.label || "").toLowerCase();
      return !choiceId.includes("abort") && !label.includes("annul");
    });
    const pick = safeChoice || conf.choices[0];
    if (pick?.id) {
      out[conf.id] = pick.id;
    }
  });
  return out;
}

async function defaultFormulaRepairHook(
  plan: AgentPlan,
  context: WorkbookContextSnapshot,
  userPrompt: string,
  opts?: RepairPlanOptions
): Promise<{ plan: AgentPlan; notes?: string[] }> {
  const { repairedPlan, notes } = await repairPlanWriteFormulas(plan, context, userPrompt, opts);
  return { plan: repairedPlan, notes };
}

function createExecutionProblem(errors: string[], status: ExecutionStatus = "error", warnings: string[] = []): ExecutionResult {
  const logs = errors.map((error) => ({
    ts: new Date().toISOString(),
    level: "error" as const,
    message: error,
  }));
  return {
    logs,
    artifacts: [],
    status,
    errors,
    ok: false,
    warnings,
    confirmationsRequested: [],
  };
}

function buildFailureResult(params: FailureParams): RunAgentPipelineResult {
  const warnings = uniqueStrings(params.warnings);
  return {
    failureStage: params.failureStage,
    plannerStage: params.plannerStage,
    plan: params.planCandidate ?? null,
    planNormalized: params.planNormalized ?? null,
    planSanitized: params.planSanitized ?? null,
    planFinalExecuted: params.planSanitized ?? null,
    planRawText: params.planRawText ?? null,
    planRawTextRetry: params.planRawTextRetry ?? null,
    planParsed: params.planParsed,
    plannerRetryUsed: params.plannerRetryUsed,
    plannerFallbackUsed: params.plannerFallbackUsed,
    plannerOutputWasNonJson: params.plannerOutputWasNonJson,
    plannerSanitizeChangedPlan: params.plannerSanitizeChangedPlan,
    plannerParseError: params.plannerParseError,
    plannerSanitizeNotes: uniqueStrings(params.plannerSanitizeNotes.concat(warnings)),
    plannerRetryReason: params.plannerRetryReason,
    warnings,
    execution: createExecutionProblem(params.errors, params.status, warnings),
    decisions: { ...(params.decisions || {}) },
    attempts: 0,
    autoAnswerStats: { autoAnswered: 0, missingBefore: 0, warnings: [] },
    confirmationsRequested: [],
  };
}

export async function runAgentPipeline(options: RunAgentPipelineOptions): Promise<RunAgentPipelineResult> {
  if (!options.prompt && !options.plan) {
    throw new Error("runAgentPipeline requires either prompt or plan");
  }
  const { context, excelAdapter } = options;
  const planWarnings: string[] = [];
  let planCandidate: AgentPlan | null = null;
  let planNormalized: AgentPlan | null = null;
  let planSanitized: AgentPlan | null = null;
  let planRawText: string | null = null;
  let planRawTextRetry: string | null = null;
  let planParsed: any = null;
  let plannerStage: RunAgentPipelineResult["plannerStage"] | undefined;
  let plannerSanitizeNotes: string[] = [];
  let plannerParseError: string | null = null;
  let plannerFallbackUsed = false;
  let plannerRetryUsed = false;
  let plannerRetryReason: string | null = null;
  let plannerOutputWasNonJson = false;
  let plannerSanitizeChangedPlan = false;

  if (options.prompt) {
    const plannerClient = options.planWithOllamaClient ?? planWithOllama;
    const plannerResult = await plannerClient({
      context,
      userPrompt: options.prompt,
      model: options.plannerModel,
      host: options.plannerHost,
      timeoutMs: options.plannerTimeoutMs,
      logger: options.plannerLogger,
    });
    planRawText = plannerResult.rawText || null;
    planRawTextRetry = plannerResult.rawTextRetry || null;
    plannerRetryUsed = !!planRawTextRetry;
    if (plannerRetryUsed) plannerRetryReason = "invalid_json_first_pass";
    plannerFallbackUsed = !!plannerResult.fallbackUsed;
    plannerSanitizeNotes = Array.isArray(plannerResult.sanitizeNotes) ? [...plannerResult.sanitizeNotes] : [];
    plannerParseError = plannerResult.failureStage || plannerResult.parseError || null;
    plannerOutputWasNonJson = plannerResult.status !== "ok";
    planParsed = safeParsePlan(planRawText) ?? (plannerResult.plan ? cloneValue(plannerResult.plan) : null);
    plannerSanitizeChangedPlan = Boolean(planParsed && plannerResult.plan) && JSON.stringify(planParsed) !== JSON.stringify(plannerResult.plan);
    plannerStage = {
      status: plannerResult.status,
      plan: plannerResult.plan as AgentPlan | undefined,
      rawText: planRawText,
      rawTextRetry: planRawTextRetry,
      errors: (plannerResult as any).errors,
      failureStage: plannerResult.failureStage || null,
      fallbackUsed: plannerResult.fallbackUsed,
      deterministicRepairApplied: plannerResult.deterministicRepairApplied,
      sanitizeNotes: plannerSanitizeNotes,
      parseError: plannerResult.parseError ?? plannerResult.failureStage ?? null,
      planParsed,
    };
    if (plannerResult.status !== "ok" || !plannerResult.plan) {
      return buildFailureResult({
        errors:
          plannerResult.status === "error"
            ? [plannerResult.error || "planner_error"]
            : (plannerResult as any).errors || ["planner_failed"],
        warnings: plannerSanitizeNotes,
        plannerStage,
        planCandidate: plannerResult.plan as AgentPlan | undefined ?? null,
        planNormalized: plannerResult.plan as AgentPlan | undefined ?? null,
        planSanitized: plannerResult.plan as AgentPlan | undefined ?? null,
        planParsed,
        planRawText,
        planRawTextRetry,
        plannerRetryUsed,
        plannerFallbackUsed,
        plannerOutputWasNonJson,
        plannerSanitizeChangedPlan,
        plannerParseError,
        plannerSanitizeNotes,
        plannerRetryReason,
        decisions: options.decisions,
        failureStage: plannerResult.failureStage || "planner",
      });
    }
    planSanitized = canonicalizePlan(plannerResult.plan as AgentPlan);
    planCandidate = planSanitized;
    planNormalized = planSanitized;
    planWarnings.push(...plannerSanitizeNotes);
  } else if (options.plan) {
    planParsed = cloneValue(options.plan);
    const canonicalInput = canonicalizePlan(options.plan);
    const normalized = normalizePlan(
      canonicalInput,
      context,
      options.plan.goal || options.prompt || undefined,
      planWarnings
    );
    if (!normalized || !Array.isArray(normalized.steps) || normalized.steps.length === 0) {
      return buildFailureResult({
        errors: ["normalize_failed"],
        warnings: planWarnings,
        plannerStage,
        planCandidate: canonicalInput,
        planNormalized: normalized ?? null,
        planSanitized: null,
        planParsed,
        plannerRetryUsed,
        plannerFallbackUsed,
        plannerOutputWasNonJson,
        plannerSanitizeChangedPlan,
        plannerParseError,
        plannerSanitizeNotes: planWarnings,
        plannerRetryReason,
        decisions: options.decisions,
        failureStage: "normalize",
      });
    }
    planNormalized = normalized as AgentPlan;
    const extraHeaders = collectExtraHeaders(planNormalized);
    const sanitized = sanitizePlan(planNormalized, context, planWarnings, extraHeaders, options.prompt);
    if (!sanitized || !Array.isArray(sanitized.steps) || sanitized.steps.length === 0) {
      return buildFailureResult({
        errors: ["sanitize_failed_unrepairable"],
        warnings: planWarnings,
        plannerStage,
        planCandidate: normalized as AgentPlan,
        planNormalized,
        planSanitized: null,
        planParsed,
        plannerRetryUsed,
        plannerFallbackUsed,
        plannerOutputWasNonJson,
        plannerSanitizeChangedPlan,
        plannerParseError,
        plannerSanitizeNotes: planWarnings,
        plannerRetryReason,
        decisions: options.decisions,
        failureStage: "sanitize",
      });
    }
    const repaired = preSchemaRepair(sanitized, context);
    repaired.version = repaired.version || "1.0";
    repaired.goal = repaired.goal || options.plan.goal || options.prompt || "User request";
    planSanitized = canonicalizePlan(repaired);
    planCandidate = planSanitized;
  }

  const formulaRepairConfig = options.formulaRepair;
  if (planCandidate && formulaRepairConfig?.enabled) {
    const repairPrompt = options.prompt || planCandidate.goal || "User request";
    const repairHook = formulaRepairConfig.hook
      ? formulaRepairConfig.hook
      : (plan: AgentPlan, ctx: WorkbookContextSnapshot, prompt: string) =>
          defaultFormulaRepairHook(plan, ctx, prompt, formulaRepairConfig.options);
    const repairResult = await repairHook(planCandidate, context, repairPrompt);
    if (repairResult && repairResult.plan) {
      planCandidate = canonicalizePlan(repairResult.plan);
      planSanitized = planCandidate;
      planNormalized = planCandidate;
      const repairNotes = repairResult.notes || [];
      if (repairNotes.length) {
        planWarnings.push(...repairNotes);
        plannerSanitizeNotes.push(...repairNotes);
      }
    }
  }

  if (!planCandidate) {
    return buildFailureResult({
      errors: ["plan_missing"],
      warnings: planWarnings,
      plannerStage,
      planCandidate: null,
      planNormalized,
      planSanitized,
      planParsed,
      plannerRetryUsed,
      plannerFallbackUsed,
      plannerOutputWasNonJson,
      plannerSanitizeChangedPlan,
      plannerParseError,
      plannerSanitizeNotes,
      plannerRetryReason,
      decisions: options.decisions,
      failureStage: "plan_missing",
    });
  }

  const validation = validatePlan(planCandidate);
  const validationErrors = validation.valid ? [] : validation.errors || ["plan_schema_invalid"];
  const invariant = validatePlanInvariants(planCandidate);
  const invariantErrors = invariant.valid ? [] : (invariant.issues || []).map((issue) => issue.message || issue.code || "plan_invariant_failed");
  if (validationErrors.length || invariantErrors.length) {
    return buildFailureResult({
      errors: [...validationErrors, ...invariantErrors],
      warnings: planWarnings,
      plannerStage,
      planCandidate,
      planNormalized,
      planSanitized,
      planParsed,
      planRawText,
      planRawTextRetry,
      plannerRetryUsed,
      plannerFallbackUsed,
      plannerOutputWasNonJson,
      plannerSanitizeChangedPlan,
      plannerParseError,
      plannerSanitizeNotes,
      plannerRetryReason,
      decisions: options.decisions,
      failureStage: "validate",
    });
  }

  const confirmationHandler = options.confirmationDecisionHandler || safeFillDecisions;
  const autoMode = options.autoAnswerMode ?? "interactive";
  const maxAttempts = options.maxAttempts ?? 3;
  const decisions: Record<string, string> = { ...(options.decisions || {}) };
  if (planCandidate && options.preConfirmDecisionHandler) {
    Object.assign(decisions, options.preConfirmDecisionHandler(planCandidate, context, options.prompt));
  }
  let planForExecution = canonicalizePlan(planCandidate);
  let execution: ExecutionResult | null = null;
  let attempts = 0;
  let autoAnsweredTotal = 0;
  let missingBefore = 0;
  const autoWarnings: string[] = [];

  while (attempts < maxAttempts) {
    attempts += 1;
    const auto = autoAnswerConfirmations(planForExecution, context, autoMode);
    autoAnsweredTotal += auto.autoAnswered;
    missingBefore = auto.missingBefore;
    autoWarnings.push(...auto.warnings);
    Object.assign(decisions, auto.decisions);
    planForExecution = applyConfirmationsToPlan(auto.plan, decisions, context) || auto.plan;
    execution = await excelAdapter.run((excelCtx) =>
      executePlan(planForExecution as AgentPlan, context, excelCtx, {
        confirmationDecisions: decisions,
        autoAnswerMode: autoMode,
        attempt: attempts,
        initialArtifacts: options.initialArtifacts,
      })
    );
    const needsConfirmation =
      execution.status === "need_user_confirmation" ||
      ((execution.errors || []) as string[]).some((err) => err === "requires_confirmation_unhandled") ||
      (execution.confirmationsRequested?.length || 0) > 0;
    if (!needsConfirmation) break;
    const extra = confirmationHandler(execution.confirmationsRequested || []);
    if (!Object.keys(extra).length) break;
    Object.assign(decisions, extra);
    planForExecution = applyConfirmationsToPlan(planCandidate, decisions, context) || planCandidate;
  }

  if (!execution) {
    execution = createExecutionProblem(["execution_failed"], "error", uniqueStrings([...planWarnings, ...autoWarnings]));
  }

  const finalPlan = canonicalizePlan(planForExecution);
  const warnings = uniqueStrings([...planWarnings, ...autoWarnings]);
  const autoWarningsUnique = uniqueStrings(autoWarnings);
  const sanitizedString = planSanitized ? JSON.stringify(planSanitized) : null;
  const parsedString = planParsed ? JSON.stringify(planParsed) : null;
  const sanitizeChanged = Boolean(sanitizedString && parsedString && sanitizedString !== parsedString);

  return {
    plannerStage,
    plan: planCandidate,
    planNormalized: planNormalized ?? planCandidate,
    planSanitized,
    planFinalExecuted: finalPlan,
    planRawText,
    planRawTextRetry,
    planParsed,
    plannerRetryUsed,
    plannerFallbackUsed,
    plannerOutputWasNonJson,
    plannerSanitizeChangedPlan: sanitizeChanged,
    plannerParseError,
    plannerSanitizeNotes: uniqueStrings(plannerSanitizeNotes.concat(planWarnings)),
    plannerRetryReason,
    warnings,
    execution,
    decisions,
    attempts,
    autoAnswerStats: {
      autoAnswered: autoAnsweredTotal,
      missingBefore,
      warnings: autoWarningsUnique,
    },
    confirmationsRequested: execution.confirmationsRequested,
  };
}
