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
      usedRange: "A1:B5",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 1, address: "A1:B5" },
      counts: { tables: 1, charts: 0 },
      tables: [
        { name: "Data", address: "Sheet1!A1:B5", dataBodyAddress: "Sheet1!A2:B5", headerAddress: "Sheet1!A1:B1", headers: ["X", "Y"] },
      ],
      blocks: [
        {
          id: "Sheet1!A1:B5",
          address: "A1:B5",
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["X", "Y"],
          columnTypes: ["number", "number"],
          preview: [],
          source: { type: "table", tableName: "Data", tableAddress: "Sheet1!A1:B5" },
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
  goal: "chart",
  
  steps: [
    {
      id: "chart1",
      macro: "create_chart",
      params: {
        source: { blockRef: "Sheet1!A1:B5" },
        mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 1 }] },
        chartType: "columnClustered",
        dest: { mode: "right", anchor: { blockRef: "Sheet1!A1:B5" } },
        titleHint: "Demo chart",
      },
    },
  ],
};

describe("executor - create_chart", () => {
  test("creates chart to the right with artifact", async () => {
    // Ensure no global Excel object is required
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).Excel = undefined;
    const ws = new FakeWorksheet("Sheet1");
    const wb = new FakeWorkbook([ws]);
    const ctx = new FakeContext(wb);
    ctx.resetWrites();
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const errs = res.logs.filter((l) => l.level === "error");
    expect(errs.length).toBe(0);
    expect(ws.charts.created.length).toBe(1);
    expect(res.artifacts?.[0]?.type).toBe("chart");
    expect(res.artifacts?.[0]?.sheet).toBe("Sheet1");
    expect(ws.charts.created[0].setPosition).toHaveBeenCalled();
    expect(ctx.getWriteCount()).toBe(0);
    expect(ws.charts.created[0].range.address).toBe("Sheet1!B2:B5");
  });

  test("creates chart on new sheet", async () => {
    const ws = new FakeWorksheet("Sheet1");
    const wb = new FakeWorkbook([ws]);
    const ctx = new FakeContext(wb);
    ctx.resetWrites();
    const newSheetPlan = {
      version: "1.0",
      goal: "chart new sheet",
      
      steps: [
        {
          id: "chart1",
          macro: "create_chart",
          params: {
            source: { blockRef: "Sheet1!A1:B5" },
            mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 1 }] },
            chartType: "columnClustered",
            dest: { mode: "newSheet" },
            titleHint: "Demo chart",
          },
        },
      ],
    };
    const res = await executePlan(newSheetPlan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const sheets = wb.worksheets.rawItems();
    expect(sheets.length).toBe(2);
    const destSheet = sheets[1]!;
    const artifactSheet = res.artifacts?.[0]?.sheet;
    expect(artifactSheet).toBeTruthy();
    expect(destSheet.charts.created.length).toBe(1);
    expect(artifactSheet).toBe(destSheet.charts.created[0].lastPosition?.sheet);
    expect(ctx.getWriteCount()).toBe(0);
    expect(destSheet.charts.created[0].range.address).toBe("Sheet1!B2:B5");
  });

  test("uses mapped columns only when creating chart", async () => {
    const ws = new FakeWorksheet("Sheet1");
    const data = [
      ["X", "Y1", "Y2"],
      [1, 10, 100],
      [2, 20, 200],
    ];
    ws.getRange("A1:C3").values = data;

    const wb = new FakeWorkbook([ws]);
    const ctx = new FakeContext(wb);
    ctx.resetWrites();

    const planMapped = {
      version: "1.0",
      goal: "chart mapped",
      
      steps: [
        {
          id: "chart1",
          macro: "create_chart",
          params: {
            source: { blockRef: "Sheet1!A1:C3" },
            mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 2 }] },
            chartType: "columnClustered",
            dest: { mode: "right", anchor: { blockRef: "Sheet1!A1:C3" } },
          },
        },
      ],
    };

    const snapshotMapped: WorkbookContextSnapshot = {
      ...snapshot,
      sheets: [
        {
          ...snapshot.sheets[0],
          name: "Sheet1",
          usedRange: "A1:C3",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2, address: "A1:C3" },
          counts: { tables: 1, charts: 0 },
          charts: [],
          limitations: [],
          tables: [
            {
              name: "Data",
              address: "Sheet1!A1:C3",
              dataBodyAddress: "Sheet1!A2:C3",
              headerAddress: "Sheet1!A1:C1",
              headers: ["X", "Y1", "Y2"],
            },
          ],
          blocks: [
            {
              id: "Sheet1!A1:C3",
              address: "A1:C3",
              kind: "table",
              confidence: 1,
              headerRowIndex: 0,
              headers: ["X", "Y1", "Y2"],
              columnTypes: ["number", "number", "number"],
              preview: [],
              source: { type: "table", tableName: "Data", tableAddress: "Sheet1!A1:C3" },
            },
          ],
        },
      ],
    };

    const res = await executePlan(planMapped as any, snapshotMapped, ctx as any, {});
    expect(res.status).toBe("ok");
    const chart = ws.charts.created[0];
    expect(chart).toBeTruthy();
    expect(chart.range.address).toBe("Sheet1!C2:C3");
    const firstSeries = chart.series.items[0];
    expect(firstSeries.setXAxisValues).toHaveBeenCalled();
    expect(ctx.getWriteCount()).toBe(0);
  });

  test("creates chart with contiguous multiple yCols", async () => {
    const ws = new FakeWorksheet("Sheet1");
    const data = [
      ["X", "Y1", "Y2", "Y3"],
      [1, 10, 100, 1000],
      [2, 20, 200, 2000],
    ];
    ws.getRange("A1:D3").values = data;
    const wb = new FakeWorkbook([ws]);
    const ctx = new FakeContext(wb);

    const planMulti = {
      version: "1.0",
      goal: "chart multi",
      
      steps: [
        {
          id: "chart1",
          macro: "create_chart",
          params: {
            source: { blockRef: "Sheet1!A1:D3" },
            mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 1 }, { colIndex: 2 }] },
            chartType: "columnClustered",
            dest: { mode: "right", anchor: { blockRef: "Sheet1!A1:D3" } },
          },
        },
      ],
    };

    const snapshotMulti = {
      ...snapshot,
      sheets: [
        {
          ...snapshot.sheets[0],
          name: "Sheet1",
          usedRange: "A1:D3",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 3, address: "A1:D3" },
          tables: [
            {
              name: "Data",
              address: "Sheet1!A1:D3",
              dataBodyAddress: "Sheet1!A2:D3",
              headerAddress: "Sheet1!A1:D1",
              headers: ["X", "Y1", "Y2", "Y3"],
            },
          ],
          blocks: [
            {
              id: "Sheet1!A1:D3",
              address: "A1:D3",
              kind: "table",
              confidence: 1,
              headerRowIndex: 0,
              headers: ["X", "Y1", "Y2", "Y3"],
              columnTypes: ["number", "number", "number", "number"],
              preview: [],
              source: { type: "table", tableName: "Data", tableAddress: "Sheet1!A1:D3" },
            },
          ],
        },
      ],
    };

    const res = await executePlan(planMulti as any, snapshotMulti as any, ctx as any, {});
    expect(res.status).toBe("ok");
    const warnLogs = res.logs.filter((l) => l.level === "warn");
    expect(warnLogs.length).toBe(0);
  });

  test("non contiguous yCols degrades to first series", async () => {
    const ws = new FakeWorksheet("Sheet1");
    const data = [
      ["X", "Y1", "Y2", "Y3"],
      [1, 10, 100, 1000],
      [2, 20, 200, 2000],
    ];
    ws.getRange("A1:D3").values = data;
    const wb = new FakeWorkbook([ws]);
    const ctx = new FakeContext(wb);
    const planMulti = {
      version: "1.0",
      goal: "chart multi non contig",
      
      steps: [
        {
          id: "chart1",
          macro: "create_chart",
          params: {
            source: { blockRef: "Sheet1!A1:D3" },
            mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 1 }, { colIndex: 3 }] },
            chartType: "columnClustered",
            dest: { mode: "right", anchor: { blockRef: "Sheet1!A1:D3" } },
          },
        },
      ],
    };
    const snapshotMulti = {
      ...snapshot,
      sheets: [
        {
          ...snapshot.sheets[0],
          name: "Sheet1",
          usedRange: "A1:D3",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 3, address: "A1:D3" },
          tables: [
            {
              name: "Data",
              address: "Sheet1!A1:D3",
              dataBodyAddress: "Sheet1!A2:D3",
              headerAddress: "Sheet1!A1:D1",
              headers: ["X", "Y1", "Y2", "Y3"],
            },
          ],
          blocks: [
            {
              id: "Sheet1!A1:D3",
              address: "A1:D3",
              kind: "table",
              confidence: 1,
              headerRowIndex: 0,
              headers: ["X", "Y1", "Y2", "Y3"],
              columnTypes: ["number", "number", "number", "number"],
              preview: [],
              source: { type: "table", tableName: "Data", tableAddress: "Sheet1!A1:D3" },
            },
          ],
        },
      ],
    };

    const res = await executePlan(planMulti as any, snapshotMulti as any, ctx as any, {});
    expect(res.status).toBe("ok");
    const chart = ws.charts.created[0];
    expect(chart.series.items.length).toBeGreaterThanOrEqual(1);
  });
});

