import { checkBlockRefs } from "./helpers/checkBlockRefs";
import { workbook_join_travaux } from "./fixtures/workbooks";

describe("checkBlockRefs", () => {
  test("detects missing blockRef", () => {
    const snapshot = workbook_join_travaux();
    const plan = {
      steps: [
        {
          id: "s1",
          macro: "join_tables",
          params: { left: { blockRef: "Travaux!A1:E5" }, right: { blockRef: "SheetX!A1:B2" }, keys: [{ left: "a", right: "b" }] },
        },
      ],
    };
    const errs = checkBlockRefs(plan, snapshot);
    expect(errs.some((e: string) => e.includes("blockref_not_found:SheetX!A1:B2"))).toBe(true);
  });
});
