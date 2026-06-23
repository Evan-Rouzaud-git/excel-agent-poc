import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { sanitizePlan } from "../src/taskpane/agent/planner/sanitizePlan";
import { canonicalizePlan } from "../src/taskpane/agent/canonicalizePlan";
import { validatePlan } from "../src/taskpane/agent/planSchema";

const snapshot = {
  sheets: [
    {
      name: "Sheet1",
      blocks: [
        {
          id: "Sheet1!A1:C4",
          address: "A1:C4",
          kind: "table",
          headers: ["Projet", "m2", "Hab"],
          columnTypes: ["text", "number", "number"],
        },
      ],
    },
  ],
  active: {},
};

describe("table_view keeps sort on created column", () => {
  test("sort by created densite is preserved", () => {
    const plan = {
      version: "1.0",
      goal: "densite",
      
      steps: [
        { id: "view1", macro: "table_view", params: { source: { blockRef: "Sheet1!A1:C4" }, select: ["Projet", "m2", "Hab"], dest: { mode: "newSheet" } } },
        {
          id: "f1",
          macro: "write_formula",
          params: { target: { artifactRef: "view1", writeMode: "newColumnRight", headerName: "densite" }, formula: "=[@m2]/[@Hab]", fillDown: true },
        },
        { id: "v2", macro: "table_view", params: { source: { artifactRef: "view1" }, select: ["Projet", "densite"], sort: { col: "densite", dir: "desc" }, dest: { mode: "newSheet" } } },
      ],
    };
    const canonical = canonicalizePlan(plan as any);
    const normalized = normalizePlan(canonical as any, snapshot as any, "densite");
    const sanitized = canonicalizePlan(sanitizePlan(normalized as any, snapshot as any, [], { v2: ["densite"], view1: ["densite"] }));
    const validation = validatePlan(sanitized);
    expect(validation.valid).toBe(true);
    const sort = sanitized.steps[2]?.params?.sort;
    expect(sort).toBeDefined();
    expect(sort.col === "densite" || sort.col === "$lastAddedColumn").toBe(true);
  });
});
