import { executePlan } from "../src/taskpane/agent/executor";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { workbook_simple_sales } from "./fixtures/workbooks";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";

describe("multi_etapes macro chain", () => {
  test("format + formule + graphique", async () => {
    const snapshot = workbook_simple_sales();
    const blockRef = snapshot.sheets[0]!.blocks![0]!.id;
    const plan = {
      version: "1.0",
      goal: "mise en forme + marge + graphique",
      
      steps: [
        { id: "fmt1", macro: "apply_format", params: { target: { blockRef }, options: { preset: "corporate_blue", freezeHeaderRow: true } } },
        {
          id: "calc1",
          macro: "write_formula",
          params: {
            target: { blockRef, writeMode: "newColumnRight", headerName: "Marge" },
            formula: "=[@Revenus]-[@Depenses]",
            fillDown: true,
            ifOverwrite: "ask",
          },
        },
        {
          id: "chart1",
          macro: "create_chart",
          params: {
            source: { blockRef },
            mapping: { xCol: { headerName: "Mois" }, yCols: [{ headerName: "Revenus" }, { headerName: "Depenses" }] },
            chartType: "line",
            dest: { mode: "right", anchor: { blockRef } },
            titleHint: "Revenus vs Dépenses",
          },
        },
        { id: "sum", macro: "summarize_actions", params: {} },
      ],
    } as any;

    const normalized = normalizePlan(plan, snapshot, plan.goal);
    const ws = new FakeWorksheet("Sheet1");
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(normalized as any, snapshot, ctx as any, { autoAnswerMode: "demoEval" });
    expect(res.status).toBe("ok");
    expect(ws.charts.created.length).toBeGreaterThan(0);
    const chart = ws.charts.created[0];
    expect(chart.range.address).toContain("Sheet1!");
    expect(res.artifacts.some((a: any) => a.type === "chart")).toBe(true);
  });
});

