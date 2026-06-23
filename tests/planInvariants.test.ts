import { validatePlanInvariants } from "../src/taskpane/agent/planInvariants";

describe("validatePlanInvariants", () => {
  const basePlan = {
    version: "1.0",
    goal: "demo",
    steps: [],
  } as any;

  test("fails when plan only contains place_output", () => {
    const plan = {
      ...basePlan,
      steps: [{ id: "p1", macro: "place_output", params: { target: { blockRef: "b1", headerName: "A" } } }],
    };
    const res = validatePlanInvariants(plan as any);
    expect(res.valid).toBe(false);
    expect(res.issues?.[0]?.code).toBe("place_output_fallback");
  });

  test("fails when table_view lacks select in newSheet mode", () => {
    const plan = {
      ...basePlan,
      steps: [{ id: "t1", macro: "table_view", params: { source: { blockRef: "b1" }, select: [], dest: { mode: "newSheet" } } }],
    };
    const res = validatePlanInvariants(plan as any);
    expect(res.valid).toBe(false);
    expect(res.issues?.[0]?.code).toBe("table_view_select_required");
  });

  test("passes when table_view provides at least one select entry", () => {
    const plan = {
      ...basePlan,
      steps: [{ id: "t1", macro: "table_view", params: { source: { blockRef: "b1" }, select: ["Projet"], dest: { mode: "newSheet" } } }],
    };
    const res = validatePlanInvariants(plan as any);
    expect(res.valid).toBe(true);
  });

  test("passes when table_view is inPlace without select", () => {
    const plan = {
      ...basePlan,
      steps: [{ id: "t2", macro: "table_view", params: { source: { blockRef: "b1" }, dest: { mode: "inPlace" } } }],
    };
    const res = validatePlanInvariants(plan as any);
    expect(res.valid).toBe(true);
  });
});
