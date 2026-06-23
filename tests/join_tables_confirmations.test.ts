import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { runAgentPipeline, ExcelAdapter } from "../src/taskpane/agent/pipeline/runAgentPipeline";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

function buildDoubleConfirmPlan() {
  const left = new FakeWorksheet("Left");
  const right = new FakeWorksheet("Right");
  left.getRange("A1:B3").values = [
    ["Key", "ValL"],
    ["A", "L1"],
    ["B", "L2"],
  ];
  right.getRange("A1:B3").values = [
    ["KeyR", "ValR"],
    ["A", "R1"],
    ["B", "R2"],
  ];
  const wb = new FakeWorkbook([left, right]);
  const ctx = new FakeContext(wb);
  const snapshot: WorkbookContextSnapshot = {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "Left", selectionAddress: "A1", selectionInBlockId: "Left!A1:B3", nearestBlockId: "Left!A1:B3" },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: "Left",
        usedRange: "A1:B3",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
        counts: { tables: 1, charts: 0 },
        tables: [],
        blocks: [{ id: "Left!A1:B3", address: "A1:B3", kind: "table", confidence: 1, headerRowIndex: 0, headers: ["Key", "ValL"], columnTypes: ["text", "text"], preview: [], source: { type: "table", tableName: "LeftTbl", tableAddress: "Left!A1:B3" } }],
        charts: [],
        limitations: [],
      },
      {
        name: "Right",
        usedRange: "A1:B3",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
        counts: { tables: 1, charts: 0 },
        tables: [],
        blocks: [{ id: "Right!A1:B3", address: "A1:B3", kind: "table", confidence: 1, headerRowIndex: 0, headers: ["KeyR", "ValR"], columnTypes: ["text", "text"], preview: [], source: { type: "table", tableName: "RightTbl", tableAddress: "Right!A1:B3" } }],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 2, tables: 0, charts: 0, blocks: 2, durationMs: 0 },
  };

  const plan = {
    version: "1.0",
    goal: "double join",
    
    steps: [
      { id: "j1", macro: "join_tables", params: { left: { blockRef: "Left!A1:B3" }, right: { blockRef: "Right!A1:B3" }, keys: [{ left: "Key", right: "KeyR" }] } },
      { id: "j2", macro: "join_tables", params: { left: { blockRef: "Left!A1:B3" }, right: { blockRef: "Right!A1:B3" }, keys: [{ left: "Key", right: "KeyR" }] } },
    ],
  };
  return { plan, snapshot, ctx, wb };
}

const makeExcelAdapter = (ctx: FakeContext): ExcelAdapter => ({
  run: async (cb) => cb(ctx as any),
});

describe("join_tables confirmations chaining", () => {
  test("handles two successive confirmation cycles", async () => {
    const { plan, snapshot, ctx } = buildDoubleConfirmPlan();
    const normalized = normalizePlan(plan as any, snapshot);
    const pipelineRes = await runAgentPipeline({
      context: snapshot,
      excelAdapter: makeExcelAdapter(ctx),
      plan: normalized as any,
      autoAnswerMode: "demoEval",
      maxAttempts: 3,
    });
    expect(pipelineRes.execution.status).toBe("ok");
  });
});

