import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { sanitizePlan } from "../src/taskpane/agent/planner/sanitizePlan";
import { executePlan } from "../src/taskpane/agent/executor";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

const makeSimpleSnapshot = (): { snapshot: WorkbookContextSnapshot; ctx: FakeContext; blockRef: string } => {
  const data = [
    ["Projet", "Ville"],
    ["A", "Paris"],
    ["B", "Lyon"],
  ];
  const ws = new FakeWorksheet("Sheet1");
  ws.getRange("A1:B3").values = data;
  const wb = new FakeWorkbook([ws]);
  const ctx = new FakeContext(wb);
  const blockRef = "Sheet1!A1:B3";
  const snapshot: WorkbookContextSnapshot = {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: blockRef, nearestBlockId: blockRef },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: "Sheet1",
        usedRange: "A1:B3",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
        counts: { tables: 1, charts: 0 },
        tables: [
          { name: "Tbl", address: blockRef, dataBodyAddress: "Sheet1!A2:B3", headerAddress: "Sheet1!A1:B1", headers: data[0] as string[] },
        ],
        blocks: [
          {
            id: blockRef,
            address: "A1:B3",
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: data[0] as string[],
            columnTypes: ["text", "text"],
            preview: [],
            source: { type: "table", tableName: "Tbl", tableAddress: blockRef },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
  };
  return { snapshot, ctx, blockRef };
};

describe("sheetNameHint and artifact enrichment", () => {
  test("sheetNameHint maps to sheetName during normalize", () => {
    const { snapshot, blockRef } = makeSimpleSnapshot();
    const raw = {
      version: "1.0",
      goal: "test",
      
      steps: [
        { id: "s1", macro: "table_view", params: { source: { blockRef }, select: ["Projet"], dest: { mode: "right", sheetNameHint: "ResultSheet" } } },
      ],
    };
    const san = sanitizePlan(raw, snapshot);
    const norm = normalizePlan(san, snapshot);
    const dest = (norm as any).steps[0].params.dest;
    expect(dest.sheetName).toBe("ResultSheet");
  });

  test("table_view artifact contains headers and counts", async () => {
    const { snapshot, ctx, blockRef } = makeSimpleSnapshot();
    const raw = {
      version: "1.0",
      goal: "test",
      
      steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Projet", "Ville"], dest: { mode: "newSheet" } } }],
    };
    const san = sanitizePlan(raw, snapshot);
    const norm = normalizePlan(san, snapshot);
    const exec = await executePlan(norm as any, snapshot, ctx as any, {});
    const art = exec.artifacts.find((a: any) => a.fromStep === "view1" && a.type === "table");
    expect(art?.headers).toEqual(["Projet", "Ville"]);
    expect(art?.rowCount).toBeGreaterThan(0);
    expect(art?.colCount).toBe(2);
  });
});

