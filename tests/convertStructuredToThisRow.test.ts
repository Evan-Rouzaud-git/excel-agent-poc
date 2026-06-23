import { convertStructuredToThisRow } from "../src/taskpane/agent/macros";

describe("convertStructuredToThisRow", () => {
  test("converts simple [Header] tokens to [@Header] when header exists", () => {
    const headers = ["Projet", "m2", "Hab"];
    const result = convertStructuredToThisRow("=[m2]/[Hab]", headers);
    expect(result.formula).toBe("=[@m2]/[@Hab]");
    expect(result.missing).toEqual([]);
  });

  test("ignores tokens with table naming or special markers", () => {
    const headers = ["m2", "Hab"];
    const formula = "=SUM(Table1[Header], [#Headers], [@m2], [Unknown])";
    const result = convertStructuredToThisRow(formula, headers);
    expect(result.formula).toBe(formula);
    expect(result.missing).toEqual(["Unknown"]);
  });
});
