import { executePlan } from "../src/taskpane/agent/executor";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";

function buildJoinScenario() {
  const leftData = [
    ["Code", "Valeur"],
    ["A", 10],
    ["B", 20],
  ];
  const rightData = [
    ["Code", "Nom"],
    ["A", "Alpha"],
    ["B", "Beta"],
  ];

  const leftWs = new FakeWorksheet("Left");
  leftWs.getRange("A1:B3").values = leftData;
  const rightWs = new FakeWorksheet("Right");
  rightWs.getRange("A1:B3").values = rightData;

  const wb = new FakeWorkbook([leftWs, rightWs]);

  const snapshot: WorkbookContextSnapshot = {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "Left", selectionAddress: "A1", selectionInBlockId: null, nearestBlockId: null },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: "Left",
        usedRange: "A1:B3",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
        counts: { tables: 1, charts: 0 },
        tables: [
          { name: "LeftTbl", address: "Left!A1:B3", dataBodyAddress: "Left!A2:B3", headerAddress: "Left!A1:B1", headers: ["Code", "Valeur"] },
        ],
        blocks: [
          {
            id: "Left!A1:B3",
            address: "A1:B3",
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: ["Code", "Valeur"],
            columnTypes: ["text", "number"],
            preview: [],
            source: { type: "table", tableName: "LeftTbl", tableAddress: "Left!A1:B3" },
          },
        ],
        charts: [],
        limitations: [],
      },
      {
        name: "Right",
        usedRange: "A1:B3",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
        counts: { tables: 1, charts: 0 },
        tables: [
          { name: "RightTbl", address: "Right!A1:B3", dataBodyAddress: "Right!A2:B3", headerAddress: "Right!A1:B1", headers: ["Code", "Nom"] },
        ],
        blocks: [
          {
            id: "Right!A1:B3",
            address: "A1:B3",
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: ["Code", "Nom"],
            columnTypes: ["text", "text"],
            preview: [],
            source: { type: "table", tableName: "RightTbl", tableAddress: "Right!A1:B3" },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 2, tables: 2, charts: 0, blocks: 2, durationMs: 0 },
  };

  return { snapshot, wb };
}

describe("artifactRef chaining", () => {
  test("apply_format and create_chart resolve artifactRef from join_tables output", async () => {
    const { snapshot, wb } = buildJoinScenario();
    const ctx = new FakeContext(wb);

    const plan = {
      version: "1.0",
      goal: "join then use artifact",
      
      steps: [
        { id: "join1", macro: "join_tables", params: { left: { blockRef: "Left!A1:B3" }, right: { blockRef: "Right!A1:B3" }, keys: [{ left: "Code", right: "Code" }], output: { sheetName: "Joined" } } },
        { id: "fmt", macro: "apply_format", params: { target: { artifactRef: "join1" }, options: { preset: "corporate_blue" } } },
        {
          id: "chart",
          macro: "create_chart",
          params: {
            source: { artifactRef: "join1" },
            mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 1 }] },
            chartType: "columnClustered",
            dest: { mode: "right", anchor: { artifactRef: "join1" } },
            titleHint: "Join chart",
          },
        },
      ],
    } as any;

    const res = await executePlan(plan, snapshot, ctx as any, {});

    expect(res.status).toBe("ok");
    const tableArtifact = res.artifacts.find((a) => a.type === "table");
    expect(tableArtifact?.blockRef).toMatch(/Joined!/);
    expect(tableArtifact?.sheetName).toMatch(/Joined/i);
    expect(tableArtifact?.rows).toBe(2);

    const joinedSheet = wb.worksheets.rawItems().find((s) => s.name.startsWith("Joined"));
    expect(joinedSheet).toBeTruthy();
    expect(joinedSheet?.charts.created.length).toBe(1);

    const chartArtifact = res.artifacts.find((a) => a.type === "chart");
    expect(chartArtifact?.sheet).toBe(joinedSheet?.name);
    const errorLogs = res.logs.filter((l) => l.level === "error");
    expect(errorLogs.length).toBe(0);
  });
});

