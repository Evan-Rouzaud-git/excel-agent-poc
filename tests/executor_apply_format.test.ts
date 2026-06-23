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
      usedRange: "A1:B3",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
      counts: { tables: 1, charts: 0 },
      tables: [
        { name: "Data", address: "Sheet1!A1:B3", dataBodyAddress: "Sheet1!A2:B3", headerAddress: "Sheet1!A1:B1", headers: ["Nom", "Revenus"] },
      ],
      blocks: [
        {
          id: "Sheet1!A1:B3",
          address: "A1:B3",
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["Nom", "Revenus"],
          columnTypes: ["text", "number"],
          preview: [],
          source: { type: "table", tableName: "Data", tableAddress: "Sheet1!A1:B3" },
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
  goal: "format table",
  
  steps: [
    {
      id: "s1",
      macro: "apply_format",
      params: {
        target: { blockRef: "Sheet1!A1:B3" },
        options: {
          header: { bold: true, background: "lightGray", fontColor: "black" },
          numberFormats: [{ headerHints: ["revenus"], format: "#,##0 \"ƒ'ª\"" }],
        },
      },
    },
  ],
};

describe("executor - apply_format", () => {
  test("applies header bold and number format", async () => {
    const ws = new FakeWorksheet("Sheet1");
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    expect(ws.getRange("A1:B1").format.font.bold).toBe(true);
    expect(ws.getRange("B2:B3").numberFormat?.[0]?.[0]).toBe('#,##0 "ƒ\'ª"');
  });

  test("ignores invalid numberFormats without crashing", async () => {
    const ws = new FakeWorksheet("Sheet1");
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const badPlan = {
      ...plan,
      steps: [
        {
          id: "s1",
          macro: "apply_format",
          params: {
            target: { blockRef: "Sheet1!A1:B3" },
            options: {
              numberFormats: { invalid: true },
            },
          },
        },
      ],
    };
    const res = await executePlan(badPlan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
  });

  test("corporate_blue preset creates table once on non-table range", async () => {
    const ws = new FakeWorksheet("Sheet1");
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const nonTableSnapshot: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: null, nearestBlockId: null },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Sheet1",
          usedRange: "G8:I12",
          valueBounds: { firstRow: 7, firstCol: 6, lastRow: 11, lastCol: 8, address: "G8:I12" },
          counts: { tables: 0, charts: 0 },
          tables: [],
          blocks: [
            {
              id: "Sheet1!G8:I12",
              address: "G8:I12",
              kind: "range",
              confidence: 0.6,
              headerRowIndex: 0,
              headers: ["H1", "H2", "H3"],
              columnTypes: ["text", "text", "text"],
              preview: [],
              source: { type: "range" },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
    };

    const improPlan = {
      version: "1.0",
      goal: "format corporate_blue",
      
      steps: [
        {
          id: "fmt1",
          macro: "apply_format",
          params: {
            target: { blockRef: "Sheet1!G8:I12" },
            options: { preset: "corporate_blue" },
          },
        },
      ],
    };

    const res = await executePlan(improPlan as any, nonTableSnapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    expect(ws.tables.addCalls.length).toBe(1);
    const call = ws.tables.addCalls[0]!;
    expect(call.hasHeaders).toBe(true);
    const addr = typeof call.address === "string" ? call.address : (call.address as any).address;
    expect(String(addr).endsWith("G8:I12")).toBe(true);
  });
});

