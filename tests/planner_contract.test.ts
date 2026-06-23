import { canonicalizePlan } from "../src/taskpane/agent/canonicalizePlan";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { sanitizePlan } from "../src/taskpane/agent/planner/sanitizePlan";
import { preSchemaRepair } from "../src/taskpane/agent/planner/ollamaPlanner";
import { validatePlan } from "../src/taskpane/agent/planSchema";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

const snapshot: WorkbookContextSnapshot = {
  workbook: { name: "Book", readOnly: false },
  active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: "Sheet1!A1:B3", nearestBlockId: "Sheet1!A1:B3" },
  capabilities: [],
  limitations: [],
  sheets: [
    {
      name: "Sheet1",
      usedRange: "A1:B3",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
      counts: { tables: 0, charts: 0 },
      tables: [],
      blocks: [
        {
          id: "Sheet1!A1:B3",
          address: "A1:B3",
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["Projet", "Budget"],
          columnTypes: ["text", "number"],
          preview: [],
          source: { type: "range" } as any,
        },
      ],
      charts: [],
      limitations: [],
    },
  ],
  totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
};

function runPipeline(plan: any) {
  const canon = canonicalizePlan(plan);
  const normalized = normalizePlan(canon as any, snapshot, "", []);
  const sanitized = canonicalizePlan(sanitizePlan(normalized as any, snapshot, []));
  const repaired = preSchemaRepair(sanitized as any, snapshot);
  const validation = validatePlan(repaired as any);
  return { plan: repaired, validation };
}

describe("planner contract", () => {
  test("normalizes ge alias to gte", () => {
    const plan = {
      version: "1.0",
      goal: "filter",
      
      steps: [
        {
          id: "view1",
          macro: "table_view",
          params: { source: { blockRef: "Sheet1!A1:B3" }, filter: [{ col: "Budget", op: "ge", value: 10 }], dest: { mode: "newSheet" } },
        },
      ],
    };
    const { plan: finalPlan, validation } = runPipeline(plan);
    expect(validation.valid).toBe(true);
    const op = (finalPlan as any).steps[0].params.filter[0].op;
    expect(op).toBe("gte");
  });

  test("accepts filter type and keeps schema valid", () => {
    const plan = {
      version: "1.0",
      goal: "filter",
      
      steps: [
        {
          id: "view1",
          macro: "table_view",
          params: { source: { blockRef: "Sheet1!A1:B3" }, filter: [{ col: "Budget", op: "gt", value: 100, type: "number" }], dest: { mode: "newSheet" } },
        },
      ],
    };
    const { validation } = runPipeline(plan);
    expect(validation.valid).toBe(true);
  });

  test("fills select when newSheet missing select", () => {
    const plan = {
      version: "1.0",
      goal: "extract",
      
      steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef: "Sheet1!A1:B3" }, dest: { mode: "newSheet" } } }],
    };
    const { plan: finalPlan, validation } = runPipeline(plan);
    expect(validation.valid).toBe(true);
    const sel = (finalPlan as any).steps[0].params.select;
    expect(Array.isArray(sel) && sel.length >= 1).toBe(true);
    expect(sel[0]).toBe("Projet");
  });

  test("date year canonicalized to ISO start", () => {
    const plan = {
      version: "1.0",
      goal: "filter date",
      
      steps: [
        {
          id: "view1",
          macro: "table_view",
          params: { source: { blockRef: "Sheet1!A1:B3" }, filter: [{ col: "Budget", op: ">=", value: 2028, type: "date" }], dest: { mode: "newSheet" } },
        },
      ],
    };
    const { plan: finalPlan, validation } = runPipeline(plan);
    expect(validation.valid).toBe(true);
    const filt = (finalPlan as any).steps[0].params.filter[0];
    expect(filt.op).toBe("gte");
    expect(filt.type).toBe("date");
    expect(filt.value).toBe("2028-01-01");
  });
});
