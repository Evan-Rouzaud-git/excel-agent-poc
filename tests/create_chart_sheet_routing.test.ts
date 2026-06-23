import { executePlan } from "../src/taskpane/agent/executor";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

function buildSnapshot() {
  const leftWs = new FakeWorksheet("Sheet1");
  leftWs.getRange("A1:B3").values = [
    ["Projet", "m2"],
    ["P1", 100],
    ["P2", 200],
  ];
  const joinWs = new FakeWorksheet("Join_Result");
  // table already present from a previous join
  joinWs.tables.add("A1:B3", true).name = "Join_Result_1";
  joinWs.getRange("A1:B3").values = [
    ["Projet", "m2"],
    ["P1", 100],
    ["P2", 200],
  ];

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
        counts: { tables: 1, charts: 0 },
        tables: [{ name: "Sheet1Table", address: "Sheet1!A1:B3", dataBodyAddress: "Sheet1!A2:B3", headerAddress: "Sheet1!A1:B1", headers: ["Projet", "m2"] }],
        blocks: [
          {
            id: "Sheet1!A1:B3",
            address: "A1:B3",
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: ["Projet", "m2"],
            columnTypes: ["text", "number"],
            preview: [],
            source: { type: "table", tableName: "Sheet1Table", tableAddress: "Sheet1!A1:B3" },
          },
        ],
        charts: [],
        limitations: [],
      },
      {
        name: "Join_Result",
        usedRange: "A1:B3",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
        counts: { tables: 1, charts: 0 },
        tables: [{ name: "Join_Result_1", address: "Join_Result!A1:B3", dataBodyAddress: "Join_Result!A2:B3", headerAddress: "Join_Result!A1:B1", headers: ["Projet", "m2"] }],
        blocks: [
          {
            id: "Join_Result!A1:B3",
            address: "A1:B3",
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: ["Projet", "m2"],
            columnTypes: ["text", "number"],
            preview: [],
            source: { type: "table", tableName: "Join_Result_1", tableAddress: "Join_Result!A1:B3" },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 2, tables: 2, charts: 0, blocks: 2, durationMs: 0 },
  };

  const ctx = new FakeContext(new FakeWorkbook([leftWs, joinWs]));
  return { snapshot, ctx };
}

describe("create_chart sheet routing", () => {
  test("chart goes to dest.sheetName using lastProducedTable when source missing", async () => {
    const { snapshot, ctx } = buildSnapshot();
    const plan = {
      version: "1.0",
      goal: "chart on join",
      
      steps: [
        {
          id: "j1",
          macro: "join_tables",
          params: {
            left: { blockRef: "Sheet1!A1:B3" },
            right: { blockRef: "Join_Result!A1:B3" },
            keys: [{ left: "Projet", right: "Projet" }],
            output: { mode: "right", anchor: { blockRef: "Join_Result!A1:B3" } },
          },
        },
        {
          id: "c1",
          macro: "create_chart",
          params: {
            dest: { mode: "right", sheetName: "Join_Result" },
            mapping: { xCol: { header: "Projet" }, yCols: [{ header: "m2" }] },
            source: {},
          },
        },
      ],
    };

    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const chartArtifact = res.artifacts.find((a) => a.type === "chart");
    expect(chartArtifact?.sheet).toBe("Join_Result");
    const joinWs = ctx.workbook.worksheets.getItem("Join_Result");
    expect(joinWs.charts.rawItems().length).toBeGreaterThan(0);
    expect(joinWs.charts.rawItems()[0]?.lastPosition?.sheet || "Join_Result").toBe("Join_Result");
  });
});

