import { canonicalizePlan } from "../src/taskpane/agent/canonicalizePlan";
import { repairPlanWriteFormulas } from "../src/taskpane/agent/planRepairer";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { sanitizePlan } from "../src/taskpane/agent/planner/sanitizePlan";
import { validatePlan } from "../src/taskpane/agent/planSchema";
import { runAgentPipeline, ExcelAdapter } from "../src/taskpane/agent/pipeline/runAgentPipeline";
import { workbook_join_travaux } from "./fixtures/workbooks";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";

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

function buildJoinContext() {
  const wsTrav = new FakeWorksheet("Travaux");
  const wsProj = new FakeWorksheet("Projet");
  const wb = new FakeWorkbook([wsTrav, wsProj]);
  const ctx = new FakeContext(wb);
  wsTrav.getRange("A1:E5").values = joinLeft;
  wsProj.getRange("A1:E5").values = joinRight;
  return { ctx, wsTrav, wsProj };
}

const makeExcelAdapter = (ctx: FakeContext): ExcelAdapter => ({
  run: async (cb) => cb(ctx as any),
});

const joinExtraHeaders: Record<string, string[]> = {
  join1: ["Projet", "charge_aff_code", "tache_nom", "tache_debut_dt", "tache_fin_dt", "code", "Typologie", "m2", "Hab"],
};
const defaultJoinKeys = [{ left: "ptvx_id", right: "Projet" }];
const aliasJoinKeys = [
  { left: "ptvx_id", right: "Projet" },
  { left: "charge_aff_code", right: "code" },
];

async function executeJoinPlan(plan: any, snapshot: any, ctx: FakeContext, decisions?: Record<string, string>) {
  const sanitized = sanitizePlan(plan, snapshot, undefined, joinExtraHeaders);
  const normalized = normalizePlan(sanitized, snapshot, plan.goal);
  const validation = validatePlan(normalized);
  expect(validation.valid).toBe(true);
  const { repairedPlan } = await repairPlanWriteFormulas(normalized as any, snapshot, plan.goal);
  const canonical = canonicalizePlan(repairedPlan);
  const pipelineRes = await runAgentPipeline({
    context: snapshot,
    excelAdapter: makeExcelAdapter(ctx),
    plan: canonical as any,
    autoAnswerMode: "demoEval",
    maxAttempts: 1,
    decisions: decisions || {},
  });
  return { pipelineRes, validation };
}

function buildJoinPlan(leftId: string, rightId: string, opts?: { joinType?: string; keys?: any[]; goal?: string }) {
  const keys = opts?.keys || defaultJoinKeys;
  const joinType = opts?.joinType || "anti_left";
  return {
    version: "1.0",
    goal: opts?.goal || "anti_left join and view ptvx_id",
    steps: [
      {
        id: "join1",
        macro: "join_tables",
        params: {
          left: { blockRef: leftId },
          right: { blockRef: rightId },
          keys,
          joinType,
          output: { mode: "newSheet" },
        },
      },
      {
        id: "view1",
        macro: "table_view",
        params: {
          source: { artifactRef: "join1" },
          select: ["ptvx_id"],
        },
      },
    ],
  };
}

test("joinC anti-left preserves ptvx_id for table_view", async () => {
  const snapshot = workbook_join_travaux();
  const leftBlock = snapshot.sheets?.[0]?.blocks?.[0]?.id || "Travaux!A1:E5";
  const rightBlock = snapshot.sheets?.[1]?.blocks?.[0]?.id || "Projet!A1:E5";
  const plan = buildJoinPlan(leftBlock, rightBlock);
  const { ctx } = buildJoinContext();
  const { pipelineRes, validation } = await executeJoinPlan(plan, snapshot, ctx);
  const execRes = pipelineRes.execution;
  expect(execRes.ok).toBe(true);
  const joinLog = (execRes.logs || []).find(
    (log: any) => log?.macro === "join_tables" && Array.isArray(log?.data?.keysUsed)
  );
  expect(joinLog).toBeDefined();
  expect(joinLog?.data?.keysUsed).toEqual(joinLog?.data?.keysPlan);
  const fallbackLog = (execRes.logs || []).find(
    (log: any) => typeof log?.message === "string" && log.message.includes("keyFallbackApplied")
  );
  expect(fallbackLog).toBeUndefined();
  const tableArtifacts = (execRes.artifacts || []).filter((a: any) => a.type === "table");
  expect(tableArtifacts.length).toBeGreaterThanOrEqual(1);
  const joinArtifact = tableArtifacts.find((a) => a.fromStep === "join1");
  expect(joinArtifact).toBeDefined();
  const viewArtifact = tableArtifacts.find((a) => a.fromStep === "view1");
  expect(viewArtifact).toBeDefined();
  expect(viewArtifact?.headers || []).toContain("ptvx_id");
});

test("joinC left join exposes both key headers and header aliases", async () => {
  const snapshot = workbook_join_travaux();
  const leftBlock = snapshot.sheets?.[0]?.blocks?.[0]?.id || "Travaux!A1:E5";
  const rightBlock = snapshot.sheets?.[1]?.blocks?.[0]?.id || "Projet!A1:E5";
  const plan = buildJoinPlan(leftBlock, rightBlock, { joinType: "left", keys: aliasJoinKeys });
  const { ctx } = buildJoinContext();
  const { pipelineRes } = await executeJoinPlan(plan, snapshot, ctx, { "join1:join_key_fallback": "use_fallback_keys" });
  const execRes = pipelineRes.execution;
  expect(execRes.status).toBe("ok");
  const joinArtifact = (execRes.artifacts || []).find(
    (a: any) => a.fromStep === "join1" && Array.isArray(a.headers) && a.headers.length
  );
  expect(joinArtifact).toBeDefined();
  expect(joinArtifact?.headers || []).toEqual(expect.arrayContaining(["ptvx_id", "charge_aff_code", "Projet", "code"]));
  const aliases = (joinArtifact as any).headerAliases || {};
  expect(aliases["ptvx_id"]).toEqual(expect.arrayContaining(["ptvx_id", "Projet"]));
  expect(aliases["projet"]).toEqual(expect.arrayContaining(["ptvx_id", "Projet"]));
  expect(aliases["charge_aff_code"]).toEqual(expect.arrayContaining(["charge_aff_code", "code"]));
  expect(aliases["code"]).toEqual(expect.arrayContaining(["charge_aff_code", "code"]));
});

test("joinC anti-left retains only left headers with alias keys", async () => {
  const snapshot = workbook_join_travaux();
  const leftBlock = snapshot.sheets?.[0]?.blocks?.[0]?.id || "Travaux!A1:E5";
  const rightBlock = snapshot.sheets?.[1]?.blocks?.[0]?.id || "Projet!A1:E5";
  const plan = buildJoinPlan(leftBlock, rightBlock, { joinType: "anti_left", keys: aliasJoinKeys });
  const { ctx } = buildJoinContext();
  const { pipelineRes } = await executeJoinPlan(plan, snapshot, ctx);
  const execRes = pipelineRes.execution;
  expect(execRes.status).toBe("ok");
  const joinArtifact = (execRes.artifacts || []).find(
    (a: any) => a.fromStep === "join1" && Array.isArray(a.headers) && a.headers.length
  );
  expect(joinArtifact).toBeDefined();
  expect(joinArtifact?.headers || []).toEqual(expect.arrayContaining(["ptvx_id", "charge_aff_code"]));
  expect(joinArtifact?.headers || []).not.toContain("Projet");
  expect(joinArtifact?.headers || []).not.toContain("code");
});
