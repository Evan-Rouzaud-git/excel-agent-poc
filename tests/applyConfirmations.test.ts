import { applyConfirmationsToPlan } from "../src/taskpane/agent/applyConfirmations";

const plan = {
  version: "1.0",
  goal: "charts",
  
  steps: [
    {
      id: "c1",
      macro: "create_chart",
      params: {
        source: { blockRef: "Sheet1!A1:B5" },
        chartType: "columnClustered",
        mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 1 }] },
        dest: { mode: "newSheet", anchor: { blockRef: "Sheet1!A1:B5" } },
      },
    },
  ],
};

const context = {
  workbook: { name: "Book", readOnly: false },
  active: {
    sheetName: "Sheet1",
    selectionAddress: "Sheet1!A1:B5",
    selectionInBlockId: "Sheet1!A1:B5",
    nearestBlockId: "Sheet1!A1:B5",
  },
  capabilities: [],
  limitations: [],
  sheets: [
    {
      name: "Sheet1",
      usedRange: "A1:B5",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 1, address: "A1:B5" },
      counts: { tables: 0, charts: 0 },
      tables: [],
          blocks: [
            {
              id: "Sheet1!A1:B5",
              address: "A1:B5",
              kind: "range" as const,
              confidence: 0.9,
              headerRowIndex: 0,
              headers: ["Ville", "Début EDP"],
              columnTypes: ["text", "text"] as ("number" | "date" | "text" | "mixed")[],
              preview: [],
              source: { type: "range" as const },
            },
          ],
      charts: [],
      limitations: [],
    },
  ],
  totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
};

describe("applyConfirmationsToPlan", () => {
  test("forces chart placement to right when confirmation is no", () => {
    const patched = applyConfirmationsToPlan(plan as any, { conf1: "no" }, context);
    const dest = patched.steps?.[0]?.params?.dest;
    expect(dest?.mode).toBe("right");
    expect(dest?.anchor?.blockRef).toBe("Sheet1!A1:B5");
  });

  test("applies column confirmation using labels and keeps params clean", () => {
    const planWithSort = {
      version: "1.0",
      goal: "sort view",
      steps: [
        {
          id: "tv1",
          macro: "table_view",
          params: {
            sort: { col: "needs_confirmation", dir: "desc" },
            source: { blockRef: "Sheet1!A1:B5" },
          },
        },
      ],
      confirmations: [
        {
          id: "tv1:sort_col",
          question: "Choisir la colonne de tri",
          choices: [
            { id: "c0", label: "Ville" },
            { id: "c1", label: "Début EDP" },
          ],
          required: true,
        },
      ],
    };
    const patched = applyConfirmationsToPlan(planWithSort as any, { "tv1:sort_col": "c1" }, context);
    const params = patched.steps?.[0]?.params || {};
    expect(params.sort?.col).toBe("Début EDP");
    expect(params.sort?.dir).toBe("desc");
    expect(Object.keys(params)).not.toContain("confirmationAppliedSort");
  });

  test("formula_column confirmation substitutes placeholder tokens", () => {
    const formulaContext = {
      workbook: { name: "Book", readOnly: false },
      active: {
        sheetName: "Sheet1",
        selectionAddress: null,
        selectionInBlockId: null,
        nearestBlockId: null,
      },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Sheet1",
          usedRange: "A1:B3",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
          counts: { tables: 0, charts: 0 },
          tables: [],
          blocks: [
            {
              id: "Sheet1!A1:B3",
              address: "A1:B3",
              kind: "range" as const,
              confidence: 0.9,
              headerRowIndex: 0,
              headers: ["Revenus", "Depenses"],
              columnTypes: ["number", "number"] as ("number" | "date" | "text" | "mixed")[],
              preview: [],
              source: { type: "range" as const },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
    };
    const planWithFormula = {
      version: "1.0",
      goal: "placeholder replace",
      steps: [
        {
          id: "f1",
          macro: "write_formula",
          params: {
            target: { blockRef: "Sheet1!A1:B3", writeMode: "newColumnRight", headerName: "% Marge" },
            formula: "=[Marge]/[Revenus]",
          },
        },
      ],
      confirmations: [
        {
          id: "formula_column",
          question: "Quelle colonne correspond à 'Marge' ?",
          choices: [
            { id: "Revenus", label: "Revenus" },
            { id: "Depenses", label: "Depenses" },
          ],
          required: true,
        },
      ],
    };
    const patched = applyConfirmationsToPlan(
      planWithFormula as any,
      { formula_column: "Depenses" },
      formulaContext as any
    );
    expect(patched.steps?.[0]?.params?.formula).toBe("=[Depenses]/[Revenus]");
  });
});

