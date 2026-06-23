import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { validatePlan } from "../src/taskpane/agent/planSchema";

const context = {
  active: { selectionInBlockId: "Sheet1!A1:C3", nearestBlockId: "Sheet1!A1:C3" },
  sheets: [
    {
      name: "Sheet1",
      blocks: [{ id: "Sheet1!A1:C3", headers: ["X", "Y1", "Y2"] }],
    },
  ],
};

describe("normalizePlan", () => {
  test("maps chartType column -> columnClustered", () => {
    const plan = {
      version: "1.0",
      goal: "g",
      
      steps: [
        {
          id: "s1",
          macro: "create_chart",
          params: {
            chartType: "column",
            source: { blockRef: "Sheet1!A1:C3" },
            mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 2 }] },
            dest: { mode: "right", anchor: "topLeft" },
          },
        },
      ],
    };
    const norm = normalizePlan(plan, context);
    expect(norm.steps[0].params.chartType).toBe("columnClustered");
    expect(norm.steps[0].params.dest.anchor.blockRef).toBe("Sheet1!A1:C3");
    const val = validatePlan(norm);
    expect(val.valid).toBe(true);
  });

  test("fallbacks invalid blockRef to nearest", () => {
    const plan = {
      version: "1.0",
      goal: "g",
      
      steps: [
        {
          id: "s1",
          macro: "create_chart",
          params: {
            chartType: "columnClustered",
            source: { blockRef: "Unknown" },
            mapping: { xCol: { colIndex: 10 }, yCols: [{ colIndex: -1 }] },
            dest: { mode: "right" },
          },
        },
      ],
    };
    const norm = normalizePlan(plan, context, "presentation");
    expect(norm.steps[0].params.source.blockRef).toBe("Sheet1!A1:C3");
    const map = norm.steps[0].params.mapping;
    expect(map.xCol.colIndex).toBe(0);
    expect(map.yCols[0].colIndex).toBe(1);
    const val = validatePlan(norm);
    expect(val.valid).toBe(true);
  });
});
  test("forces chart dest to right when user prompt not asking new sheet", () => {
    const plan = {
      version: "1.0",
      goal: "g",
      
      steps: [
        {
          id: "s1",
          macro: "create_chart",
          params: {
            chartType: "columnClustered",
            source: { blockRef: "Sheet1!A1:C3" },
            mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 2 }] },
            dest: { mode: "newSheet" },
          },
        },
      ],
    };
    const norm = normalizePlan(plan, context, "fais un graphique");
    expect(norm.steps[0].params.dest.mode).toBe("right");
  });

