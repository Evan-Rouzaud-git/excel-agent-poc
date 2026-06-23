import { pickSuites } from "./demoEvalRunner";

describe("prompt suite guards", () => {
  test("formula suite uses only P06..P12", () => {
    const suites = pickSuites(["formula"]);
    const formula = suites.find((s) => s.id === "formula");
    expect(formula).toBeTruthy();
    const ids = (formula?.prompts || []).map((p) => p.id);
    expect(ids.length).toBe(7);
    const allowed = new Set(["P06", "P07", "P08", "P09", "P10", "P11", "P12"]);
    ids.forEach((id) => expect(allowed.has(id)).toBe(true));
  });
});
