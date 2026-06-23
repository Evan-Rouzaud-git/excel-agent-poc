import { canonicalizePlan } from "../src/taskpane/agent/canonicalizePlan";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { sanitizePlan } from "../src/taskpane/agent/planner/sanitizePlan";

describe("datasetRef chaining and repair", () => {
  const context = {
    sheets: [{ name: "Sheet1", blocks: [{ id: "B1", headers: ["Projet", "code", "surface", "population"] }] }],
    active: {},
  };

  test("adds created column to select and sort after write_formula", () => {
    const plan = {
      version: "1.0",
      goal: "densité",
      
      steps: [
        { id: "calc1", macro: "write_formula", params: { target: { blockRef: "B1", headerName: "densité" }, formula: "=[@surface]/[@population]" } },
        { id: "tv1", macro: "table_view", params: { source: { blockRef: "B1" }, dest: { mode: "newSheet" } } },
      ],
    };
    const normalized = normalizePlan(canonicalizePlan(plan as any), context as any, "prompt");
    const sanitized = sanitizePlan(normalized as any, context as any);
    const tv = sanitized.steps.find((s: any) => s.id === "tv1");
    const sel = tv.params.select as string[];
    const normSel = sel.map((s: string) => s.toLowerCase());
    expect(normSel).toEqual(expect.arrayContaining(["projet", "code"]));
    expect(normSel.some((h) => h === "densité" || h === "$lastaddedcolumn")).toBe(true);
    const sortCol = (tv.params.sort?.col || "").toLowerCase();
    expect(["densité", "$lastaddedcolumn"]).toContain(sortCol);
    expect(tv.params.source.artifactRef || tv.params.source.blockRef).toBeTruthy();
  });

  test("rewrite filter/sort/select to $lastAddedColumn when matching last created header", () => {
    const plan = {
      version: "1.0",
      goal: "densité puis filtre",
      
      steps: [
        { id: "calc1", macro: "write_formula", params: { target: { blockRef: "B1", headerName: "Densité" }, formula: "=[@population]/[@surface]" } },
        {
          id: "view1",
          macro: "table_view",
          params: {
            source: { blockRef: "B1" },
            select: ["Projet", "code", "densité"],
            filter: [{ col: "densité", op: "gt", value: 2 }],
            sort: { col: "densité", dir: "asc" },
            dest: { mode: "newSheet" },
          },
        },
      ],
    };
    const normalized = normalizePlan(canonicalizePlan(plan as any), context as any, "prompt");
    const sanitized = sanitizePlan(normalized as any, context as any);
    const tv = sanitized.steps.find((s: any) => s.id === "view1");
    expect(tv.params.select).toContain("$lastAddedColumn");
    expect(tv.params.sort.col).toBe("$lastAddedColumn");
    expect(tv.params.filter[0].col).toBe("$lastAddedColumn");
  });

  test("resolves source to latest dataset across long chains", () => {
    const longPlan = {
      version: "1.0",
      goal: "long chain",
      
      steps: [
        { id: "tv0", macro: "table_view", params: { source: { blockRef: "B1" }, select: ["Projet", "code"], dest: { mode: "newSheet" } } },
        { id: "calc1", macro: "write_formula", params: { target: { blockRef: "B1", headerName: "dens1" }, formula: "=1" } },
        { id: "s1", macro: "summarize_actions", params: {} },
        { id: "s2", macro: "summarize_actions", params: {} },
        { id: "s3", macro: "summarize_actions", params: {} },
        { id: "s4", macro: "summarize_actions", params: {} },
        { id: "s5", macro: "summarize_actions", params: {} },
        { id: "s6", macro: "summarize_actions", params: {} },
        { id: "s7", macro: "summarize_actions", params: {} },
        { id: "s8", macro: "summarize_actions", params: {} },
        { id: "tv1", macro: "table_view", params: { source: { blockRef: "B1" }, dest: { mode: "newSheet" } } },
      ],
    };
    const normalized = normalizePlan(canonicalizePlan(longPlan as any), context as any, "prompt");
    const sanitized = sanitizePlan(normalized as any, context as any);
    const tv1 = sanitized.steps.find((s: any) => s.id === "tv1");
    expect(tv1.params.source.artifactRef).toBe("tv0");
    const sel = (tv1.params.select as string[]).map((s: string) => s.toLowerCase());
    expect(sel.some((h) => h === "dens1" || h === "$lastaddedcolumn")).toBe(true);
    const sortCol = (tv1.params.sort?.col || "").toLowerCase();
    expect(["dens1", "$lastaddedcolumn"]).toContain(sortCol);
  });
});
