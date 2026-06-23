import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";

const baseContext = {
  sheets: [
    {
      name: "Feuil1",
      blocks: [
        { id: "Feuil1!A1:B3", headers: ["A", "B"], columnTypes: ["text", "text"] },
        { id: "Feuil1!D1:E3", headers: ["X", "Y"], columnTypes: ["text", "text"] },
      ],
    },
  ],
  active: { selectionInBlockId: "Feuil1!A1:B3", nearestBlockId: "Feuil1!A1:B3" },
};

describe("normalizePlan selectionPolicy", () => {
  test("defaults to all columns when selectionPolicy absent", () => {
    const plan = {
      version: "1.0",
      goal: "join",
      
      steps: [
        {
          id: "s1",
          macro: "join_tables",
          params: {
            left: { blockRef: "Feuil1!A1:B3" },
            right: { blockRef: "Feuil1!D1:E3" },
            keys: [{ left: "A", right: "X" }],
            select: { right: { mode: "list", columns: [] } },
          },
        },
      ],
    };
    const normalized = normalizePlan(plan as any, baseContext, "");
    const params = (normalized.steps[0] as any).params;
    expect(params.selectionPolicy).toBe("defaultAll");
    expect(params.select.left.mode).toBe("all");
    expect(params.select.right.mode).toBe("all");
    expect(params.select.right.columns.length).toBe(2);
  });

  test("explicit policy keeps list but repairs empty list", () => {
    const plan = {
      version: "1.0",
      goal: "join",
      
      steps: [
        {
          id: "s1",
          macro: "join_tables",
          params: {
            left: { blockRef: "Feuil1!A1:B3" },
            right: { blockRef: "Feuil1!D1:E3" },
            keys: [{ left: "A", right: "X" }],
            selectionPolicy: "explicit",
            select: { right: { mode: "list", columns: [] } },
          },
        },
      ],
    };
    const normalized = normalizePlan(plan as any, baseContext, "");
    const params = (normalized.steps[0] as any).params;
    expect(params.selectionPolicy).toBe("explicit");
    expect(params.select.right.mode).toBe("all"); // repaired because list empty
    expect(params.select.left.mode).toBe("all"); // still all by default
  });
});

