/* eslint-disable no-console */
import fs from "fs";
import path from "path";
import { afterAll, expect, jest, test } from "@jest/globals";
import promptsDefault from "../prompts/demo_prompts.json";
import { budgetValues, workbook_join_travaux, workbook_simple_sales, workbook_table_view } from "./fixtures/workbooks";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";
import { AgentPlan } from "../src/taskpane/agent/types";
import { mockPlanner } from "../src/taskpane/agent/mockPlanner";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";
import { repairWriteFormulaStep } from "../src/taskpane/agent/planner/formulaRepairer";
import { runAgentPipeline, ExcelAdapter } from "../src/taskpane/agent/pipeline/runAgentPipeline";

export type DemoEvalMode = "mock" | "ollama";
export type DemoEvalSuiteId = "format" | "chart" | "formula" | "join" | "view" | "multi";

type PromptItem = { id: string; prompt: string; expect?: any };

export interface DemoEvalSuiteConfig {
  id: DemoEvalSuiteId;
  prompts: PromptItem[];
  buildSnapshot: () => WorkbookContextSnapshot;
  buildExcelCtx: () => { ctx: FakeContext; ws?: FakeWorksheet; wb?: FakeWorkbook };
}

export interface EvalResult {
  id: string;
  prompt: string;
  suiteId: DemoEvalSuiteId;
  validPlan: boolean;
  executedOk: boolean;
  errors: string[];
  steps: string[];
  artifacts: number;
   artifactsDetail?: any[];
  confirmationsCount?: number;
  warningsCount?: number;
  failureStage?: string;
  failureError?: string;
  failureStack?: string;
  failureLogs?: any[];
  failureStepId?: string;
  failureStepMacro?: string;
  failureStepParams?: any;
  workbookSummary?: any;
  planRawText?: string | null;
  planRawTextRetry?: string | null;
  planParsed?: any;
  planSanitized?: any;
  planFinalExecuted?: any;
  plannerRetryUsed: boolean;
  plannerFallbackUsed: boolean;
  plannerOutputWasNonJson: boolean;
  plannerSanitizeChangedPlan: boolean;
  plannerParseError: string | null;
  plannerSanitizeNotes: string[];
  plannerRetryReason: string | null;
  preSteps?: string[];
}

const DEFAULT_TIMEOUT = Number(process.env.DEMO_EVAL_TIMEOUT_MS || process.env.AGENT_TIMEOUT_MS || 90000);
const DEFAULT_MODE: DemoEvalMode = process.env.DEMO_EVAL_MODE === "mock" ? "mock" : "ollama";
const INCLUDE_PLANS = process.env.DEMO_EVAL_INCLUDE_PLANS !== "0";
const INCLUDE_RAW = process.env.REPORT_INCLUDE_RAW === "1";
const AUTO_CONFIRM = process.env.DEMO_EVAL_AUTO_CONFIRM !== "0"; // default ON
const SUITE_PARAM = (process.env.DEMO_EVAL_SUITE || "all").toLowerCase();
const PLANNER_MODEL = process.env.OLLAMA_MODEL || "qwen3:8b";
const MAX_PROMPTS_GLOBAL = Number(process.env.DEMO_EVAL_MAX_PROMPTS || 0);
const USE_FORMULA_REPAIR = process.env.DEMO_EVAL_USE_FORMULA_REPAIR !== "0";
const MOCK_FORMULA_REPAIR = process.env.DEMO_EVAL_FORMULA_REPAIR_MOCK === "1";

const joinLeft = [
  ["ptvx_id", "charge_aff_code", "tache_nom", "tache_debut_dt", "tache_fin_dt"],
  ["Lyon", "SHO", "Travaux", "13/09/2026", "07/11/2027"],
  ["Paris", "NTE", "Travaux", "10/02/2026", "02/12/2027"],
  ["Bordeaux", "SMA", "Travaux", "19/01/2027", "10/09/2028"],
  ["Marseille", "API", "Réception travaux", "01/10/2028", "01/10/2028"],
];
const joinRight = [
  ["Projet", "code", "Typologie", "m2", "Hab"],
  ["Lyon", "78VER1", "Collectifs", 5020, 20],
  ["Paris", "75PAR1", "Maisons individuelles", 2500, 15],
  ["Bordeaux", "33BOR1", "Collectifs + MI", 5000, 75],
  ["Marseille", "13MAR1", "Grand Collectif", 7000, 80],
];

function buildSalesExcelCtx() {
  const ws = new FakeWorksheet("Sheet1");
  const wb = new FakeWorkbook([ws]);
  const ctx = new FakeContext(wb);
  ws.getRange("A1:I101").values = budgetValues;
  return { ctx, ws, wb };
}

function buildJoinExcelCtx() {
  const wsTrav = new FakeWorksheet("Travaux");
  const wsProj = new FakeWorksheet("Projet");
  const wb = new FakeWorkbook([wsTrav, wsProj]);
  const ctx = new FakeContext(wb);
  wsTrav.getRange("A1:E5").values = joinLeft;
  wsProj.getRange("A1:E5").values = joinRight;
  return { ctx, wb, ws: wsTrav };
}

function buildViewExcelCtx() {
  const ws = new FakeWorksheet("Projets");
  const data = [
    ["Projet", "Ville", "Début EDP", "Fin EDP", "Budget", "Statut"],
    ["PRJ-Alpha", "Marseille", "01/02/2026", "15/05/2026", 120000, "En cours"],
    ["PRJ-Beta", "Lyon", "10/03/2026", "30/06/2026", 90000, ""],
    ["PRJ-Gamma", "Bordeaux", "", "20/07/2026", 110000, "En retard"],
    ["PRJ-Delta", "Marignane", "05/01/2026", "12/04/2026", 80000, "Clôturé"],
    ["PRJ-Epsilon", "", "18/02/2026", "22/05/2026", 70000, "En cours"],
  ];
  ws.getRange("A1:F6").values = data;
  const wb = new FakeWorkbook([ws]);
  const ctx = new FakeContext(wb);
  return { ctx, ws, wb };
}

function hasExactMacros(p: PromptItem, macros: string[]) {
  const m = (p.expect?.hasMacro || []) as string[];
  if (macros.length !== m.length) return false;
  return macros.every((x) => m.includes(x));
}
const hasMacro = (p: PromptItem, macro: string) => ((p.expect?.hasMacro || []) as string[]).includes(macro);

const isFormulaIdAllowed = (id?: string) => {
  if (!id) return false;
  const num = Number(id.replace(/[^0-9]/g, ""));
  return num >= 6 && num <= 12;
};

function promptSets() {
  const base = promptsDefault as PromptItem[];
  return {
    format: base.filter((p) => hasExactMacros(p, ["apply_format"])),
    chart: base.filter((p) => hasExactMacros(p, ["create_chart"])),
    formula: base.filter((p) => hasExactMacros(p, ["write_formula"]) && isFormulaIdAllowed(p.id)),
    multi: base.filter((p) => (p.expect?.hasMacro || []).length > 1),
    join: base.filter((p) => hasExactMacros(p, ["join_tables"])),
    view: base.filter((p) => hasMacro(p, "table_view")),
  };
}

export function defaultDemoEvalSuites(): DemoEvalSuiteConfig[] {
  const sets = promptSets();
  return [
    { id: "format", prompts: sets.format, buildSnapshot: workbook_simple_sales, buildExcelCtx: buildSalesExcelCtx },
    { id: "chart", prompts: sets.chart, buildSnapshot: workbook_simple_sales, buildExcelCtx: buildSalesExcelCtx },
    { id: "formula", prompts: sets.formula, buildSnapshot: workbook_simple_sales, buildExcelCtx: buildSalesExcelCtx },
    { id: "view", prompts: sets.view, buildSnapshot: workbook_table_view, buildExcelCtx: buildViewExcelCtx },
    { id: "multi", prompts: sets.multi, buildSnapshot: workbook_simple_sales, buildExcelCtx: buildSalesExcelCtx },
    { id: "join", prompts: sets.join, buildSnapshot: workbook_join_travaux, buildExcelCtx: buildJoinExcelCtx },
  ];
}

export function pickSuites(filter?: string | string[]): DemoEvalSuiteConfig[] {
  const all = defaultDemoEvalSuites();
  if (!filter || (Array.isArray(filter) && filter.length === 0)) return all;
  const normId = (s: string): DemoEvalSuiteId => {
    const v = s.toLowerCase();
    if (v === "table_view" || v === "tableview") return "view";
    if (v === "join_tables" || v === "jointables") return "join";
    return v as DemoEvalSuiteId;
  };
  const list = Array.isArray(filter) ? filter : filter.split(",").map((s) => s.trim());
  const ids = list.filter(Boolean).map((s) => normId(s));
  const picked = all.filter((s) => ids.includes(s.id));
  return picked.length > 0 ? picked : all;
}


async function mockFormulaRepairHook(plan: any, snapshot: WorkbookContextSnapshot, prompt: string) {
  const steps = (plan.steps || []).map((step: any) => ({ ...step }));
  const notes: string[] = [];
  for (let i = 0; i < steps.length; i += 1) {
    const step = steps[i];
    if (step.macro !== "write_formula") continue;
    const res = await repairWriteFormulaStep(step, snapshot, prompt, {
      callOllama: async () => JSON.stringify({ formula: (step.params || {}).formula }),
    });
    steps[i] = res.patchedStep;
    if (res.repairLog?.notes) notes.push(res.repairLog.notes);
    if (res.repairLog?.reason) notes.push(res.repairLog.reason);
  }
  return { plan: { ...plan, steps }, notes };
}

const asArray = (val: any): any[] => {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  if (typeof val.rawItems === "function") return val.rawItems() || [];
  if (Array.isArray(val.items)) return val.items;
  if (typeof val.items === "function") {
    try {
      const res = val.items();
      if (Array.isArray(res)) return res;
    } catch {
      // ignore
    }
  }
  if (Array.isArray(val._items)) return val._items;
  return [];
};

export function snapshotSummary(snapshot: WorkbookContextSnapshot, ctx?: FakeContext) {
  try {
    const sheetSource = ctx?.workbook?.worksheets || snapshot.sheets || [];
    const sheets = asArray(sheetSource).map((s: any) => {
      const tables = asArray((s as any).tables).map((t: any) => ({ name: t.name, address: t.address }));
      const ranges = Object.keys((s as any).ranges || {});
      return { name: s.name, tables, rangeCount: ranges.length };
    });
    return { sheetCount: sheets.length, sheets };
  } catch (err: any) {
    return { sheetCount: 0, sheets: [], error: err?.message || String(err) };
  }
}

function lastLogs(logs: any[] | undefined, limit = 30) {
  if (!Array.isArray(logs)) return [];
  return logs.slice(-limit);
}

function buildAutoDecisions(plan: any, prompt: string) {
  const decisions: Record<string, string> = {};
  const promptText = (prompt || "").toLowerCase();
  const confirmations = plan?.confirmations || [];
  confirmations.forEach((c: any) => {
    const q = (c.question || "").toLowerCase();
    const safeChoices = (c.choices || []).filter(
      (ch: any) => !(ch.id || "").toLowerCase().includes("abort") && !(ch.label || "").toLowerCase().includes("annul")
    );
    const choice = safeChoices[0] || c.choices?.[0];
    const choice2 = safeChoices[1] || c.choices?.[1];
    let picked = choice;
    const dupHint =
      q.includes("doublon") ||
      q.includes("double") ||
      promptText.includes("doublon") ||
      promptText.includes("double") ||
      promptText.includes("deuxieme") ||
      promptText.includes("deuxième");
    if (dupHint && choice2) picked = choice2;
    if (picked) decisions[c.id] = picked.id;
  });
  return decisions;
}

export function runDemoEvalSuites(suites: DemoEvalSuiteConfig[], opts: { mode?: DemoEvalMode } = {}) {
  const MODE = opts.mode || DEFAULT_MODE;
  const timeoutMs = DEFAULT_TIMEOUT;
  const expectedMacros: Record<DemoEvalSuiteId, string[] | null> = {
    format: ["apply_format"],
    chart: ["create_chart"],
    formula: ["write_formula"],
    join: ["join_tables"],
    view: null,
    multi: null,
  };

  const promptLimit = (count: number) => {
    if (MAX_PROMPTS_GLOBAL && MAX_PROMPTS_GLOBAL > 0) return Math.min(MAX_PROMPTS_GLOBAL, count);
    return count;
  };

  const promptTotal = suites.reduce((acc, s) => acc + promptLimit(s.prompts.length), 0) || 1;
  const globalTimeout = promptTotal * timeoutMs + 60000;
  jest.setTimeout(globalTimeout);

  const outDir = path.resolve(__dirname, "output");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const results: EvalResult[] = [];
  let timedOutCount = 0;
  let totalMs = 0;
  let totalCharts = 0;
  let totalFormulas = 0;
  let totalConfirmations = 0;
  let totalConfirmationsAuto = 0;
  const suiteSummaries: Record<string, { prompts: number; valid: number; executed: number }> = {};
  const recordResult = (res: EvalResult) => {
    results.push(res);
    console.log(
      `[demoEval] prompt done suite=${res.suiteId} id=${res.id} status=${res.validPlan && res.executedOk ? "ok" : "fail"} stage=${res.failureStage || "done"} errors=${(res.errors || []).slice(0, 2).join(" | ")}`
    );
  };


  const runPrompt = async (suite: DemoEvalSuiteConfig, item: PromptItem) => {
    let snapshot = suite.buildSnapshot();
    let preArtifacts: any[] = [];
    let preSteps: string[] | undefined;
    let { ctx } = suite.buildExcelCtx();
    if (item.id === "TV6") {
      const joinSnapshot = workbook_join_travaux();
      const { ctx: joinCtx } = buildJoinExcelCtx();
      const allBlocks = joinSnapshot.sheets.flatMap((s) => s.blocks || []);
      const leftBlockRef = allBlocks[0]?.id;
      const rightBlockRef = allBlocks[1]?.id || allBlocks.find((b) => b.id !== leftBlockRef)?.id;
      const leftHeader = (allBlocks[0]?.headers || [])[0] || "ptvx_id";
      const rightHeader = (allBlocks[1]?.headers || [])[0] || "Projet";
      const joinPlan: AgentPlan = {
        version: "1.0",
        goal: "setup join1",
        steps: [
          {
            id: "join1",
            macro: "join_tables",
            params: {
              left: { blockRef: leftBlockRef },
              right: { blockRef: rightBlockRef },
              keys: [{ left: leftHeader, right: rightHeader }],
              output: { mode: "newSheet" },
            },
          },
        ],
      };
      const joinAdapter: ExcelAdapter = {
        run: async (cb) => cb(joinCtx as any),
      };
      const joinPipeline = await runAgentPipeline({
        context: joinSnapshot,
        excelAdapter: joinAdapter,
        plan: joinPlan,
        maxAttempts: 1,
        autoAnswerMode: "demoEval",
      });
      if (joinPipeline.execution.status !== "ok") {
        throw new Error(
          `TV6 join setup failed: ${joinPipeline.execution.errors?.join(", ") || "plan execution error"}`
        );
      }
      preArtifacts = joinPipeline.execution.artifacts || [];
      preSteps = ["join_tables"];
      snapshot = joinSnapshot;
      ctx = joinCtx;
    }
    const t0 = Date.now();
    const adapter: ExcelAdapter = {
      run: async (cb) => cb(ctx as any),
    };
    const planForMock = MODE === "mock" ? await mockPlanner(snapshot, item.prompt) : undefined;
    const pipelineRes = await runAgentPipeline({
      context: snapshot,
      excelAdapter: adapter,
      prompt: MODE === "ollama" ? item.prompt : undefined,
      plan: MODE === "mock" ? planForMock : undefined,
      plannerModel: PLANNER_MODEL,
      plannerTimeoutMs: timeoutMs,
      autoAnswerMode: AUTO_CONFIRM ? "demoEval" : "interactive",
      maxAttempts: 2,
      initialArtifacts: preArtifacts,
      preConfirmDecisionHandler: AUTO_CONFIRM ? (plan) => buildAutoDecisions(plan, item.prompt) : undefined,
      formulaRepair: USE_FORMULA_REPAIR
        ? {
            enabled: true,
            hook: MOCK_FORMULA_REPAIR ? mockFormulaRepairHook : undefined,
          }
        : undefined,
    });
    const execRes = pipelineRes.execution;
    const planFinalExecuted = pipelineRes.planFinalExecuted;
    const warnings = pipelineRes.warnings || [];
    let failureStage = pipelineRes.failureStage;
    if (!failureStage && (execRes.status === "need_user_confirmation" || execRes.status === "error")) {
      failureStage = "execute";
    }
    if (pipelineRes.failureStage === "timeout" || pipelineRes.plannerParseError === "timeout") {
      timedOutCount += 1;
    }
    const invalidStages = new Set(["planner", "validate", "plan_missing", "normalize", "sanitize", "timeout"]);
    const validPlan = !failureStage || !invalidStages.has(failureStage);
    let executedOk = execRes.status === "ok";
    let errors: string[] = execRes.errors?.slice() || [];
    if (execRes.status === "need_user_confirmation") {
      executedOk = false;
      errors = ["requires_confirmation_unhandled"];
    } else if (!executedOk) {
      errors = errors.length ? errors : (execRes.logs || []).filter((l) => l.level === "error").map((l) => l.message);
      if (!errors.length) errors = ["execution_failed"];
    }
    const failureError = errors[0];
    const failureLogs = lastLogs(execRes.logs, 40);
    let failureStepId: string | undefined;
    let failureStepMacro: string | undefined;
    let failureStepParams: any | undefined;
    const lastErr = (execRes.logs || [])
      .slice()
      .reverse()
      .find((l: any) => l.level === "error" || l.level === "warn" || l.level === "info");
    if (lastErr && planFinalExecuted) {
      failureStepId = lastErr.stepId;
      failureStepMacro = lastErr.macro;
      const step = (planFinalExecuted?.steps || []).find((s: any) => s.id === lastErr.stepId);
      failureStepParams = step?.params;
    }
    const confirmationsCount = planFinalExecuted?.confirmations?.length || 0;
    totalConfirmations += confirmationsCount;
    totalConfirmationsAuto += pipelineRes.autoAnswerStats.autoAnswered;
    const artifacts = execRes.artifacts || [];
    totalCharts += artifacts.filter((a: any) => a.type === "chart").length;
    totalFormulas += (planFinalExecuted?.steps || []).filter((s: any) => s.macro === "write_formula").length;
    totalMs += Date.now() - t0;
    const workbookSummary = snapshotSummary(snapshot, ctx);
    const artifactsDetail = [
      ...(preArtifacts || []).map((a: any) => ({ ...a, _phase: "pre" })),
      ...artifacts.map((a: any) => ({ ...a, _phase: "main" })),
    ];
    recordResult({
      id: item.id,
      suiteId: suite.id,
      prompt: item.prompt,
      validPlan,
      executedOk,
      errors,
      steps: (planFinalExecuted?.steps || []).map((s: any) => s.macro),
      artifacts: artifacts.length,
      confirmationsCount,
      warningsCount: warnings.length,
      planRawText: pipelineRes.planRawText,
      planRawTextRetry: pipelineRes.planRawTextRetry,
      planParsed: pipelineRes.planParsed,
      planSanitized: pipelineRes.planSanitized,
      planFinalExecuted,
      failureStage,
      failureError,
      failureStack: undefined,
      failureLogs,
      failureStepId,
      failureStepMacro,
      failureStepParams,
      workbookSummary,
      plannerRetryUsed: pipelineRes.plannerRetryUsed,
      plannerFallbackUsed: pipelineRes.plannerFallbackUsed,
      plannerOutputWasNonJson: pipelineRes.plannerOutputWasNonJson,
      plannerSanitizeChangedPlan: pipelineRes.plannerSanitizeChangedPlan,
      plannerParseError: pipelineRes.plannerParseError,
      plannerSanitizeNotes: pipelineRes.plannerSanitizeNotes,
      plannerRetryReason: pipelineRes.plannerRetryReason,
      preSteps,
      artifactsDetail,
    });
  };

  suites.forEach((suite) => {
    const promptList = suite.prompts.slice(0, promptLimit(suite.prompts.length));
    const expectMacros = expectedMacros[suite.id] ?? null;
    if (expectMacros) {
      const contamination = promptList.find((p) => {
        const macros = p.expect?.hasMacro || [];
        return expectMacros.some((m) => !macros.includes(m));
      });
      if (contamination) {
        throw new Error(`Suite contamination: prompt ${contamination.id} lacks expected macros for suite ${suite.id}`);
      }
    }
    if (suite.id === "multi") {
      const bad = promptList.find((p) => (p.expect?.hasMacro || []).length <= 1);
      if (bad) throw new Error(`Suite contamination: prompt ${bad.id} not multi-macro`);
    }
    if (suite.id === "formula") {
      const bad = promptList.find((p) => !isFormulaIdAllowed(p.id));
      if (bad) throw new Error(`Suite contamination: formula prompt ${bad.id} not in P06..P12`);
    }
    console.log(
      `[demoEval] suite=${suite.id} prompts=${promptList.length} ids=${promptList.map((p) => p.id).join(",")} expectedMacros=${expectMacros ? expectMacros.join(",") : "mixed"}`
    );
    const suiteTimeout = promptList.length * timeoutMs + 30000;
    suiteSummaries[suite.id] = { prompts: promptList.length, valid: 0, executed: 0 };

    if (MODE === "ollama") {
      test(
        `demoEval ${suite.id} (${promptList.length} prompts) [${MODE}]`,
        async () => {
          for (const item of promptList) {
            try {
              await runPrompt(suite, item);
            } catch (err: any) {
              recordResult({
                id: item.id,
                suiteId: suite.id,
                prompt: item.prompt,
                validPlan: false,
                executedOk: false,
                errors: [err?.message || String(err)],
                steps: [],
                artifacts: 0,
                failureStage: "exception",
                failureError: err?.message,
                failureStack: err?.stack,
                plannerOutputWasNonJson: false,
                plannerRetryUsed: false,
                plannerFallbackUsed: false,
                plannerSanitizeChangedPlan: false,
                plannerParseError: null,
                plannerSanitizeNotes: [],
                plannerRetryReason: null,
              });
            }
          }
          await new Promise((r) => setTimeout(r, 0));
        },
        suiteTimeout
      );
    } else {
      for (const item of promptList) {
        test(
          `demoEval ${suite.id} mock - ${item.id}`,
          async () => {
            try {
              await runPrompt(suite, item);
            } catch (err: any) {
              recordResult({
                id: item.id,
                suiteId: suite.id,
                prompt: item.prompt,
                validPlan: false,
                executedOk: false,
                errors: [err?.message || String(err)],
                steps: [],
                artifacts: 0,
                failureStage: "exception",
                failureError: err?.message,
                failureStack: err?.stack,
                plannerOutputWasNonJson: false,
                plannerRetryUsed: false,
                plannerFallbackUsed: false,
                plannerSanitizeChangedPlan: false,
                plannerParseError: null,
                plannerSanitizeNotes: [],
                plannerRetryReason: null,
              });
            }
          },
          timeoutMs + 1000
        );
      }
    }
  });

  const persistReport = (entry: any) => {
    try {
      const outPath = path.join(outDir, "demo_eval_report.json");
      // overwrite with current run only
      const toWrite = { runs: [entry] };
      const tmp = `${outPath}.tmp-${Date.now()}`;
      fs.writeFileSync(tmp, JSON.stringify(toWrite, null, 2), "utf8");
      fs.renameSync(tmp, outPath);
      console.log(`[demoEval] report written -> ${outPath}`);
    } catch (err) {
      console.error("persistReport failed", err);
    }
  };

  let lastReport: any = null;
  const flushReport = () => {
    if (lastReport) persistReport(lastReport);
  };
  process.on("exit", flushReport);
  process.on("SIGINT", () => {
    flushReport();
    process.exit(1);
  });
  process.on("uncaughtException", (err) => {
    console.error("uncaughtException", err);
    flushReport();
  });
  process.on("unhandledRejection", (err: any) => {
    console.error("unhandledRejection", err);
    flushReport();
  });

  afterAll(async () => {
    const total = results.length || 1;
    const gate1 = results.filter((r) => r.validPlan).length / total;
    const gate2 = results.filter((r) => r.validPlan && r.executedOk).length / total;
    const avgPromptMs = total > 0 ? totalMs / total : 0;
    const failed = results.filter((r) => !r.executedOk || !r.validPlan);
    const totals = {
      chartsCreatedTotal: totalCharts,
      formulasWrittenTotal: totalFormulas,
      confirmationsAskedTotal: totalConfirmations,
      confirmationsAutoAnsweredTotal: totalConfirmationsAuto,
      timeouts: timedOutCount,
      warningsTotal: results.reduce((a, r) => a + (r.warningsCount || 0), 0),
      errorsTotal: results.reduce((a, r) => a + (r.errors?.length || 0), 0),
      retryUsedCount: results.filter((r) => r.plannerRetryUsed).length,
      fallbackUsedCount: results.filter((r) => r.plannerFallbackUsed).length,
      nonJsonCount: results.filter((r) => r.plannerOutputWasNonJson).length,
      sanitizeChangedCount: results.filter((r) => r.plannerSanitizeChangedPlan).length,
    };
    results.forEach((r) => {
      const sum = suiteSummaries[r.suiteId];
      if (!sum) return;
      if (r.validPlan) sum.valid += 1;
      if (r.validPlan && r.executedOk) sum.executed += 1;
    });

    Object.entries(suiteSummaries).forEach(([id, sum]) => {
      const denom = sum.prompts || 1;
      console.log(`Suite ${id}: valid ${(100 * (sum.valid / denom)).toFixed(1)}% | executed ${(100 * (sum.executed / denom)).toFixed(1)}%`);
    });
    console.log(`Demo Eval - Gate1 valid plan: ${(gate1 * 100).toFixed(1)}%`);
    console.log(`Demo Eval - Gate2 executed ok: ${(gate2 * 100).toFixed(1)}%`);
    console.log(
      `Totals: charts=${totals.chartsCreatedTotal}, formulas=${totals.formulasWrittenTotal}, confirmations asked=${totals.confirmationsAskedTotal}, auto=${totals.confirmationsAutoAnsweredTotal}, errors=${totals.errorsTotal}, avgMs=${avgPromptMs.toFixed(
        1
      )}`
    );
    console.log(`Planner: retries=${totals.retryUsedCount}, fallbacks=${totals.fallbackUsedCount}, nonJson=${totals.nonJsonCount}, sanitizeChanged=${totals.sanitizeChangedCount}`);
    if (failed.length > 0) {
      console.log(
        "Failed prompts:",
        failed.map((f) => `${f.suiteId}/${f.id}: ${(f.errors && f.errors[0]) || "unknown"}`).join(" | ")
      );
    }
    if (INCLUDE_PLANS) {
      expect(
        results.every(
          (r) =>
            Object.prototype.hasOwnProperty.call(r, "planRawText") &&
            Object.prototype.hasOwnProperty.call(r, "planRawTextRetry") &&
            Object.prototype.hasOwnProperty.call(r, "planParsed") &&
            Object.prototype.hasOwnProperty.call(r, "planSanitized") &&
            Object.prototype.hasOwnProperty.call(r, "planFinalExecuted") &&
            Object.prototype.hasOwnProperty.call(r, "plannerRetryUsed") &&
            Object.prototype.hasOwnProperty.call(r, "plannerFallbackUsed") &&
            Object.prototype.hasOwnProperty.call(r, "plannerOutputWasNonJson") &&
            Object.prototype.hasOwnProperty.call(r, "plannerSanitizeChangedPlan") &&
            Object.prototype.hasOwnProperty.call(r, "plannerParseError")
        )
      ).toBe(true);
    }
    const outDir = path.join(__dirname, "output");
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
    const cleanedResults = results.map((r) => {
      const { planRawText, planRawTextRetry, ...rest } = r as any;
      if (INCLUDE_RAW) return { ...rest, planRawText, planRawTextRetry };
      return rest;
    });
    const report = {
      mode: MODE,
      model: PLANNER_MODEL,
      suites: Object.keys(suiteSummaries),
      timeoutMs: timeoutMs,
      perPromptMs: timeoutMs,
      timestamp: new Date().toISOString(),
      suiteArg: SUITE_PARAM,
      gate1,
      gate2,
      timedOutCount,
      avgPromptMs,
      results: cleanedResults,
      totals,
      suiteSummaries,
    };
    lastReport = report;
    persistReport(report);
    await new Promise((r) => setTimeout(r, 0));
    jest.clearAllTimers();
  });
}
