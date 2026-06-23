import { validatePlan } from "../src/taskpane/agent/planSchema";

const basePlan = {
  version: "1.0",
  goal: "Test plan",
  
  steps: [{ id: "s1", macro: "place_output", params: {} }],
};

describe("planSchema", () => {
  test("accepts a valid minimal plan", () => {
    const res = validatePlan(basePlan);
    expect(res.valid).toBe(true);
  });

  test("rejects plan without steps", () => {
    const res = validatePlan({ ...basePlan, steps: [] });
    expect(res.valid).toBe(false);
  });

  test("rejects unknown macro", () => {
    const res = validatePlan({
      ...basePlan,
      steps: [{ id: "s1", macro: "unknown_macro", params: {} }],
    } as any);
    expect(res.valid).toBe(false);
  });

  test("accepts summarize_actions macro with empty params", () => {
    const res = validatePlan({
      version: "1.0",
      goal: "summ",
      
      steps: [{ id: "s1", macro: "summarize_actions", params: {} }],
    });
    expect(res.valid).toBe(true);
  });

  test("accepts join_tables minimal params", () => {
    const res = validatePlan({
      version: "1.0",
      goal: "join",
      
      steps: [
        {
          id: "j1",
          macro: "join_tables",
          params: { left: { blockRef: "Sheet1!A1:B3" }, right: { blockRef: "Sheet2!A1:B3" }, keys: [{ left: "Id", right: "Code" }] },
        },
      ],
    });
    expect(res.valid).toBe(true);
  });
});

