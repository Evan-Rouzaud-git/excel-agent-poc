import { convertStructuredToA1 } from "../src/taskpane/agent/macros";

describe("convertStructuredToA1", () => {
  test("rewrites plain header tokens into row-specific A1 references", () => {
    const headers = ["Projet", "m2", "Hab"];
    const converted = convertStructuredToA1("=[m2]/[Hab]", headers, 1, 0);
    expect(converted.formula).toBe("=B2/C2");
    expect(converted.missing).toEqual([]);
  });

  test("returns missing headers when a token cannot be resolved", () => {
    const headers = ["Projet", "m2", "Hab"];
    const converted = convertStructuredToA1("=[m2]/[Unknown]", headers, 1, 0);
    expect(converted.formula).toContain("0");
    expect(converted.missing).toContain("Unknown");
  });
});
