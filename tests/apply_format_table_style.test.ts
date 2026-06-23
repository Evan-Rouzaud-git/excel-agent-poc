import { executePlan } from "../src/taskpane/agent/executor";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";

describe("apply_format preset corporate_blue on table", () => {
  test("disables banded rows and sets neutral style", async () => {
    const ws = new FakeWorksheet("Sheet1");
    // pre-create a table in fake worksheet
    const tbl = ws.tables.add("A1:C3", true);
    tbl.name = "MyTbl";
    ws.getRange("A1:C3").values = [
      ["H1", "H2", "H3"],
      [1, 2, 3],
      [4, 5, 6],
    ];
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const snapshot: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: "Sheet1!A1:C3", nearestBlockId: "Sheet1!A1:C3" },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Sheet1",
          usedRange: "A1:C3",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2, address: "A1:C3" },
          counts: { tables: 1, charts: 0 },
          tables: [{ name: "MyTbl", address: "Sheet1!A1:C3", dataBodyAddress: "Sheet1!A2:C3", headerAddress: "Sheet1!A1:C1", headers: ["H1", "H2", "H3"] }],
          blocks: [
            {
              id: "Sheet1!A1:C3",
              address: "A1:C3",
              kind: "table",
              confidence: 1,
              headerRowIndex: 0,
              headers: ["H1", "H2", "H3"],
              columnTypes: ["text", "text", "text"],
              preview: [],
              source: { type: "table", tableName: "MyTbl", tableAddress: "Sheet1!A1:C3" },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
    };

    const plan = {
      version: "1.0",
      goal: "format",
      
      steps: [{ id: "fmt1", macro: "apply_format", params: { target: { sheetName: "Sheet1", tableName: "MyTbl" }, options: { preset: "corporate_blue" } } }],
    };

    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    expect(tbl.showBandedRows).toBe(false);
    expect(tbl.style).toBe("TableStyleLight1");
  });
});

