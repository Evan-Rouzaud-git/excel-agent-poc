import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { validatePlan } from "../src/taskpane/agent/planSchema";

const context = {
  sheets: [
    {
      name: "S1",
      blocks: [
        { id: "S1!A1:B3", headers: ["A", "B"], columnTypes: ["text", "text"] },
        { id: "S1!D1:E3", headers: ["X", "Y"], columnTypes: ["text", "text"] },
      ],
    },
  ],
  active: { sheetName: "S1", selectionInBlockId: "S1!A1:B3", nearestBlockId: "S1!A1:B3" },
};

describe("normalizePlan join_tables compatibility", () => {
  test("legacy output.name and select.right array shapes are normalized and validate", () => {
    const plan = {
      version: "1.0",
      goal: "join",
      
      steps: [
        {
          id: "j1",
          macro: "join_tables",
          params: {
            left: { blockRef: "S1!A1:B3" },
            right: { blockRef: "S1!D1:E3" },
            keys: [{ left: "A", right: "X" }],
            output: { name: "Join_Result_Legacy" },
            // legacy select.right as array
            "select.right": ["Y"],
          },
        },
      ],
    };

    const normalized = normalizePlan(plan as any, context);
    const stepParams = (normalized.steps[0] as any).params;
    expect(stepParams.output.sheetName).toBe("Join_Result_Legacy");
    expect(stepParams.select.right.mode).toBe("list");
    expect(stepParams.select.right.columns).toEqual(["Y"]);

    const validation = validatePlan(normalized);
    expect(validation.valid).toBe(true);
  });

  test("select.right without mode but with columns array becomes mode=list", () => {
    const plan = {
      version: "1.0",
      goal: "join",
      
      steps: [
        {
          id: "j1",
          macro: "join_tables",
          params: {
            left: { blockRef: "S1!A1:B3" },
            right: { blockRef: "S1!D1:E3" },
            keys: [{ left: "A", right: "X" }],
            select: { right: { columns: ["Y"] } },
          },
        },
      ],
    };
    const normalized = normalizePlan(plan as any, context);
    const stepParams = (normalized.steps[0] as any).params;
    expect(stepParams.select.right.mode).toBe("list");
    expect(stepParams.select.right.columns).toEqual(["Y"]);
    expect(validatePlan(normalized).valid).toBe(true);
  });
});

