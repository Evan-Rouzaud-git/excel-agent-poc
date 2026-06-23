import { canonicalizePlan } from "../src/taskpane/agent/canonicalizePlan";

describe("canonicalizePlan", () => {
  test("converts string source/target to objects", () => {
    const plan = {
      version: "1.0",
      goal: "demo",
      
      steps: [
        { id: "s1", macro: "table_view", params: { source: "Sheet1!A1:B10", dest: { mode: "newSheet" }, select: ["A"] } },
        { id: "s2", macro: "write_formula", params: { target: "Sheet1!A1:B10", formula: "=1" } },
      ],
    };
    const out = canonicalizePlan(plan as any);
    expect(out.steps?.[0]?.params?.source).toEqual({ blockRef: "Sheet1!A1:B10" });
    expect(out.steps?.[1]?.params?.target).toEqual({ blockRef: "Sheet1!A1:B10" });
  });

  test("canonicalizes sort and filter shapes", () => {
    const plan = {
      version: "1.0",
      goal: "demo",
      
      steps: [
        {
          id: "v1",
          macro: "table_view",
          params: { source: { blockRef: "b1" }, select: ["A"], filter: { col: "A", op: "gt", value: 1 }, sort: ["A", "desc"], dest: { mode: "newSheet" } },
        },
      ],
    };
    const out = canonicalizePlan(plan as any);
    expect(out.steps?.[0]?.params?.sort).toEqual({ col: "A", dir: "desc" });
    expect(out.steps?.[0]?.params?.filter).toEqual([{ col: "A", op: "gt", value: 1 }]);
  });

  test("fuzzy normalizes filter op typos", () => {
    const plan = {
      version: "1.0",
      goal: "demo",
      
      steps: [
        {
          id: "v1",
          macro: "table_view",
          params: { source: { blockRef: "b1" }, select: ["Hab"], filter: [{ col: "Hab", op: "supéreur à", value: 25 }], dest: { mode: "newSheet" } },
        },
      ],
    };
    const out = canonicalizePlan(plan as any);
    expect(out.steps?.[0]?.params?.filter).toEqual([{ col: "Hab", op: "gt", value: 25 }]);
  });

  test("symbol op = stays eq", () => {
    const plan = {
      version: "1.0",
      goal: "demo",
      
      steps: [
        { id: "v1", macro: "table_view", params: { source: { blockRef: "b1" }, select: ["Hab"], filter: [{ col: "Hab", op: "=", value: 1 }], dest: { mode: "newSheet" } } },
      ],
    };
    const out = canonicalizePlan(plan as any);
    expect(out.steps?.[0]?.params?.filter).toEqual([{ col: "Hab", op: "eq", value: 1 }]);
  });

  test("token fuzzy egux a -> eq", () => {
    const plan = {
      version: "1.0",
      goal: "demo",
      
      steps: [
        { id: "v1", macro: "table_view", params: { source: { blockRef: "b1" }, select: ["Hab"], filter: [{ col: "Hab", op: "egux a", value: 1 }], dest: { mode: "newSheet" } } },
      ],
    };
    const out = canonicalizePlan(plan as any);
    expect(out.steps?.[0]?.params?.filter).toEqual([{ col: "Hab", op: "eq", value: 1 }]);
  });
});
