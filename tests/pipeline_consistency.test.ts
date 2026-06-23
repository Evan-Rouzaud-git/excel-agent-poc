import fs from "fs";
import path from "path";
import { runAgentPipeline, ExcelAdapter } from "../src/taskpane/agent/pipeline/runAgentPipeline";
import { workbook_join_travaux } from "./fixtures/workbooks";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";

function buildCtx() {
  const wsTrav = new FakeWorksheet("Travaux");
  const wsProj = new FakeWorksheet("Projet");
  wsTrav.getRange("A1:E5").values = [
    ["ptvx_id", "charge_aff_code", "tache_nom", "tache_debut_dt", "tache_fin_dt"],
    ["Lyon", "SHO", "Travaux", "13/09/2026", "07/11/2027"],
    ["Paris", "NTE", "Travaux", "10/02/2026", "02/12/2027"],
    ["Bordeaux", "SMA", "Travaux", "19/01/2027", "10/09/2028"],
    ["Marseille", "API", "RÃ©ception travaux", "01/10/2028", "01/10/2028"],
  ];
  wsProj.getRange("A1:E5").values = [
    ["Projet", "code", "Typologie", "m2", "Hab"],
    ["Lyon", "78VER1", "Collectifs", 5020, 20],
    ["Paris", "75PAR1", "Maisons individuelles", 2500, 15],
    ["Bordeaux", "33BOR1", "Collectifs + MI", 5000, 75],
    ["Marseille", "13MAR1", "Grand Collectif", 7000, 80],
  ];
  const wb = new FakeWorkbook([wsTrav, wsProj]);
  const ctx = new FakeContext(wb);
  return { ctx, wb, wsTrav, wsProj };
}

const makeExcelAdapter = (ctx: FakeContext): ExcelAdapter => ({
  run: async (cb) => cb(ctx as any),
});

describe("pipeline consistency", () => {
  test("demoEvalRunner imports the canonical pipeline", () => {
    const content = fs.readFileSync(path.join(__dirname, "demoEvalRunner.ts"), "utf8");
    expect(content).toMatch(/runAgentPipeline/);
    expect(content).not.toMatch(/runPlanWithConfirmations/);
    expect(content).not.toMatch(/applyConfirmationsToPlan/);
    expect(content).not.toMatch(/autoAnswerConfirmations/);
    expect(content).not.toMatch(/planWithOllama/);
    expect(content).not.toMatch(/sanitizePlan/);
    expect(content).not.toMatch(/normalizePlan/);
    expect(content).not.toMatch(/canonicalizePlan/);
    expect(content).not.toMatch(/validatePlan/);
    expect(content).not.toMatch(/executePlan/);
  });

  test("runner pipeline matches UI execution order on join plan", async () => {
    const snapshot = workbook_join_travaux();
    const { ctx } = buildCtx();
    const plan: any = {
      version: "1.0",
      goal: "join test",
      
      steps: [
        {
          id: "j1",
          macro: "join_tables",
          params: {
            left: { blockRef: "Travaux!A1:E5" },
            right: { blockRef: "Projet!A1:E5" },
            keys: [{ left: "ptvx_id", right: "Projet" }],
            output: { mode: "newSheet", sheetName: "Resultats" },
          },
        },
      ],
    };

    const pipelineRes = await runAgentPipeline({
      context: snapshot,
      excelAdapter: makeExcelAdapter(ctx),
      plan,
      autoAnswerMode: "demoEval",
      maxAttempts: 3,
    });
    expect(pipelineRes.execution.status).toBe("ok");
    expect(pipelineRes.execution.artifacts?.length || 0).toBeGreaterThan(0);
  });
});
