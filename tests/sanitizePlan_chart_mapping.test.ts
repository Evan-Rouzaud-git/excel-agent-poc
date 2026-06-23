import { sanitizePlan } from "../src/taskpane/agent/planner/sanitizePlan";

describe("sanitizePlan chart mapping", () => {
  test("preserves mapping colIndex when valid", () => {
    const plan = {
      version: "1.0",
      goal: "g",
      
      steps: [
        {
          id: "c1",
          macro: "create_chart",
          params: {
            mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 3 }] },
            source: { blockRef: "b1" },
            dest: { mode: "right", anchor: { blockRef: "b1" } },
          },
        },
      ],
    };
    const out = sanitizePlan(plan);
    expect(out.steps[0].params.mapping.yCols[0].colIndex).toBe(3);
    expect(out.steps[0].params.mapping.xCol.colIndex).toBe(0);
  });
});

