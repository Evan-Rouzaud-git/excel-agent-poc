import { executePlan } from "../src/taskpane/agent/executor";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";
import { workbook_simple_sales } from "./fixtures/workbooks";

describe("create_chart headerName support", () => {
  const snapshot = workbook_simple_sales();

  test("resolves header names to colIndex", async () => {
    const blockId = snapshot.sheets[0]!.blocks![0]!.id;
    const plan = {
      version: "1.0",
      goal: "chart headers",
      
      steps: [
        {
          id: "c1",
          macro: "create_chart",
          params: {
            source: { blockRef: blockId },
            mapping: { xCol: { headerName: "Mois" }, yCols: [{ headerName: "Depenses" }, { headerName: "Revenus" }] },
            chartType: "columnClustered",
            dest: { mode: "right", anchor: { blockRef: blockId } },
          },
        },
      ],
    } as any;

    const ws = new FakeWorksheet("Sheet1");
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(plan, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const chart = ws.charts.created[0];
    expect(chart).toBeDefined();
    expect(chart.range.address).toContain("B2:C101");
  });

  test("fails clearly when header not found", async () => {
    const blockId = snapshot.sheets[0]!.blocks![0]!.id;
    const plan = {
      version: "1.0",
      goal: "chart headers missing",
      
      steps: [
        {
          id: "c1",
          macro: "create_chart",
          params: {
            source: { blockRef: blockId },
            mapping: { xCol: { headerName: "Mois" }, yCols: [{ headerName: "Inconnue" }] },
            chartType: "columnClustered",
            dest: { mode: "right", anchor: { blockRef: blockId } },
          },
        },
      ],
    } as any;

    const ws = new FakeWorksheet("Sheet1");
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(plan, snapshot, ctx as any, {});
    expect(res.status).toBe("error");
    const warn = res.logs.find((l) => l.level === "warn" && l.message.includes("chart_mapping_header_not_found"));
    expect(warn).toBeDefined();
  });
});

