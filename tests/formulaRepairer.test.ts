import { repairWriteFormulaStep, __private__ } from "../src/taskpane/agent/planner/formulaRepairer";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

const simpleContext: WorkbookContextSnapshot = {
  workbook: { name: "Book", readOnly: false },
  active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: "Sheet1!A1:C4", nearestBlockId: "Sheet1!A1:C4" },
  capabilities: [],
  limitations: [],
  sheets: [
    {
      name: "Sheet1",
      usedRange: "A1:C4",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 2, address: "A1:C4" },
      counts: { tables: 1, charts: 0 },
      tables: [
        {
          name: "TableBudget",
          address: "Sheet1!A1:C4",
          dataBodyAddress: "Sheet1!A2:C4",
          headerAddress: "Sheet1!A1:C1",
          headers: ["Mois", "Depenses", "Revenus"],
        },
      ],
      blocks: [
        {
          id: "Sheet1!A1:C4",
          address: "A1:C4",
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["Mois", "Depenses", "Revenus"],
          columnTypes: ["text", "number", "number"],
          preview: [
            ["Mois", "Depenses", "Revenus"],
            ["Jan", 100, 120],
            ["Fev", 110, 140],
          ],
          source: { type: "table", tableName: "TableBudget", tableAddress: "Sheet1!A1:C4" },
        },
      ],
      charts: [],
      limitations: [],
    },
  ],
  totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
};

describe("formulaRepairer __private__ helpers", () => {
  test("parseJson extracts first JSON object", () => {
    const text = "```json\n{\"formula\":\"=1\"}\n``` trailing";
    const parsed = __private__.parseJson(text);
    expect(parsed.ok).toBe(true);
    expect((parsed as any).value.formula).toBe("=1");
  });

  test("postChecks rejects unknown headers", () => {
    const block = simpleContext.sheets[0]!.blocks[0]!;
    const res = __private__.postChecks({ formula: "=[@Inconnue]" }, block);
    expect(res.ok).toBe(false);
    expect(res.reason).toContain("unknown_headers");
  });

  test("postChecks rejects A1 refs for tables", () => {
    const block = simpleContext.sheets[0]!.blocks[0]!;
    const res = __private__.postChecks({ formula: "=A2" }, block);
    expect(res.ok).toBe(false);
    expect(res.reason).toBe("a1_ref_forbidden");
  });
});

describe("repairWriteFormulaStep", () => {
  const baseStep = {
    id: "s1",
    macro: "write_formula",
    params: {
      target: { blockRef: "Sheet1!A1:C4", writeMode: "newColumnRight", headerName: "Marge" },
      formula: "=[@Revenus]-[@Depenses]",
      fillDown: true,
    },
  } as any;

  test("applies patched formula from model", async () => {
    const callOllama = jest.fn().mockResolvedValue('{"formula":"=IFERROR([@Revenus]-[@Depenses],0)","numberFormat":"0.0","fillDown":true}');
    const { patchedStep, repairLog } = await repairWriteFormulaStep(baseStep, simpleContext, "calc", { callOllama });
    expect(patchedStep.params.formula).toBe("=IFERROR([@Revenus]-[@Depenses],0)");
    expect(patchedStep.params.numberFormat).toBe("0.0");
    expect(repairLog.ok).toBe(true);
    expect(callOllama).toHaveBeenCalled();
  });

  test("returns original step when headers missing", async () => {
    const callOllama = jest
      .fn()
      .mockResolvedValue('{"formula":"=[@Category]","notes":"missing required column: Category"}');
    const { patchedStep, repairLog } = await repairWriteFormulaStep(baseStep, simpleContext, "calc", { callOllama });
    expect(repairLog.ok).toBe(false);
    expect(patchedStep).toBe(baseStep);
  });
});
