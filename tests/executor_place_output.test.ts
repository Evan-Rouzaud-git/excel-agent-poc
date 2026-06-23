import { executePlan } from "../src/taskpane/agent/executor";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

const snapshot: WorkbookContextSnapshot = {
  workbook: { name: "Book", readOnly: false },
  active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: null, nearestBlockId: null },
  capabilities: [],
  limitations: [],
  sheets: [
    {
      name: "Sheet1",
      usedRange: "A1:B2",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1, address: "A1:B2" },
      counts: { tables: 1, charts: 0 },
      tables: [
        { name: "Table1", address: "Sheet1!A1:B2", dataBodyAddress: "Sheet1!A2:B2", headerAddress: "Sheet1!A1:B1", headers: ["A", "B"] },
      ],
      blocks: [
        {
          id: "Sheet1!A1:B2",
          address: "A1:B2",
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["A", "B"],
          columnTypes: ["text", "text"],
          preview: [],
          source: { type: "table", tableName: "Table1", tableAddress: "Sheet1!A1:B2" },
        },
      ],
      charts: [],
      limitations: [],
    },
  ],
  totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
};

const basePlan = {
  version: "1.0",
  goal: "test placement",
  
  steps: [
    {
      id: "s1",
      macro: "place_output",
      params: { mode: "right", anchor: { blockRef: "Sheet1!A1:B2" }, avoidOverwrite: true, minBlankArea: { rows: 2, cols: 2 } },
    },
  ],
};

function buildContext() {
  const ws = new FakeWorksheet("Sheet1");
  const wb = new FakeWorkbook([ws]);
  return { ws, ctx: new FakeContext(wb) };
}

describe("executor - place_output", () => {
  test("runs without confirmation when area empty", async () => {
    const { ctx } = buildContext();
    const res = await executePlan(basePlan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    expect(res.confirmationsRequested).toBeUndefined();
  });

  test("requests confirmation when area not blank", async () => {
    const { ctx, ws } = buildContext();
    ws.getRange("C1:D2").values = [
      [1, 2],
      [3, 4],
    ];
    const res = await executePlan(basePlan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("need_user_confirmation");
    expect(res.confirmationsRequested).toBeDefined();
    expect(res.confirmationsRequested?.[0]?.choices.map((c) => c.id)).toContain("newSheet");
  });
});

