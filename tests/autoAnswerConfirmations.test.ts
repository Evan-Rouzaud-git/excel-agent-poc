import { autoAnswerConfirmations } from "../src/taskpane/agent/autoConfirm";
import { AgentPlan } from "../src/taskpane/agent/types";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";
import { applyConfirmationsToPlan } from "../src/taskpane/agent/applyConfirmations";

const ctxBase: WorkbookContextSnapshot = {
  workbook: { name: "Book", readOnly: false },
  active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: null, nearestBlockId: null },
  capabilities: [],
  limitations: [],
  sheets: [
    {
      name: "Sheet1",
      usedRange: "A1:C4",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 2, address: "A1:C4" },
      counts: { tables: 1, charts: 0 },
      tables: [],
      blocks: [
        {
          id: "Sheet1!A1:C4",
          address: "A1:C4",
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["CodeVille", "VilleNom", "Region"],
          columnTypes: ["text", "text", "text"],
          preview: [],
          source: { type: "table", tableName: "Left", tableAddress: "Sheet1!A1:C4" },
        },
      ],
      charts: [],
      limitations: [],
    },
    {
      name: "Sheet2",
      usedRange: "A1:C4",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 2, address: "A1:C4" },
      counts: { tables: 1, charts: 0 },
      tables: [],
      blocks: [
        {
          id: "Sheet2!A1:C4",
          address: "A1:C4",
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["VilleCode", "VilleNom", "Region"],
          columnTypes: ["text", "text", "text"],
          preview: [],
          source: { type: "table", tableName: "Right", tableAddress: "Sheet2!A1:C4" },
        },
      ],
      charts: [],
      limitations: [],
    },
  ],
  totals: { sheets: 2, tables: 2, charts: 0, blocks: 2, durationMs: 0 },
};

describe("autoAnswerConfirmations", () => {
  test("removes confirmations when keys already valid", () => {
    const plan: AgentPlan = {
      version: "1.0",
      goal: "join test",
      
      steps: [
        { id: "j1", macro: "join_tables", params: { left: { blockRef: "Sheet1!A1:C4" }, right: { blockRef: "Sheet2!A1:C4" }, keys: [{ left: "CodeVille", right: "VilleCode" }], select: { right: { mode: "list", columns: ["VilleNom", "Region"] } } } },
      ],
      confirmations: [{ id: "c1", question: "Confirmer les clés ?", choices: [{ id: "ok", label: "ok" }], required: true }],
    };
    const res = autoAnswerConfirmations(plan, ctxBase, "interactive");
    expect(res.plan.confirmations?.length || 0).toBe(0);
    expect(Object.keys(res.decisions).length).toBe(1);
  });

  test("chooses best choice based on headers", () => {
    const plan: AgentPlan = {
      version: "1.0",
      goal: "join test",
      
      steps: [
        { id: "j1", macro: "join_tables", params: { left: { blockRef: "Sheet1!A1:C4" }, right: { blockRef: "Sheet2!A1:C4" }, keys: [{ left: "CodeVille", right: "VilleCode" }] } },
      ],
      confirmations: [
        {
          id: "c1",
          question: "Choisir les colonnes à garder",
          choices: [
            { id: "c1", label: "Code seulement" },
            { id: "c2", label: "VilleNom et Region" },
          ],
          required: true,
        },
      ],
    };
    const res = autoAnswerConfirmations(plan, ctxBase, "demoEval");
    expect(res.decisions["c1"]).toBe("c2");
  });

  test("does not auto-answer non-safe in interactive", () => {
    const plan: AgentPlan = {
      version: "1.0",
      goal: "join test",
      
      steps: [{ id: "j1", macro: "join_tables", params: { left: { blockRef: "Sheet1!A1:C4" }, right: { blockRef: "Sheet2!A1:C4" } } }],
      confirmations: [{ id: "c1", question: "Choisir ?", choices: [{ id: "a", label: "A" }], required: true }],
    };
    const res = autoAnswerConfirmations(plan, ctxBase, "interactive");
    expect(Object.keys(res.decisions).length).toBe(0);
    expect(res.plan.confirmations?.length || 0).toBe(1);
  });

  test("prefers fallback key and avoids abort", () => {
    const plan: AgentPlan = {
      version: "1.0",
      goal: "join test",
      
      steps: [{ id: "j1", macro: "join_tables", params: { left: { blockRef: "Sheet1!A1:C4" }, right: { blockRef: "Sheet2!A1:C4" }, keys: [{ left: "X", right: "Y" }] } }],
      confirmations: [
        {
          id: "conf1",
          question: "Fallback cle join ?",
          choices: [
            { id: "abort", label: "Annuler" },
            { id: "keep_plan_keys", label: "Garder plan" },
            { id: "use_fallback_keys", label: "Utiliser fallback" },
          ],
          required: true,
        },
      ],
    };
    const res = autoAnswerConfirmations(plan, ctxBase, "demoEval");
    expect(res.decisions["conf1"]).toBe("use_fallback_keys");
  });

  test("auto-answer picks the valid joinKey pair among invalid choices", () => {
    const plan: AgentPlan = {
      version: "1.0",
      goal: "join test",
      
      steps: [
        {
          id: "j1",
          macro: "join_tables",
          params: {
            left: { blockRef: "Sheet1!A1:C4" },
            right: { blockRef: "Sheet2!A1:C4" },
            keys: [{ left: "X", right: "Y" }],
          },
        },
      ],
      confirmations: [
        {
          id: "joinKey:j1",
          question: "Choisir la clé de jointure",
          choices: [
            { id: "missing|missing", label: "mauvaise paire" },
            { id: "CodeVille|VilleCode", label: "CodeVille ↔ VilleCode" },
            { id: "VilleNom|VilleNom", label: "VilleNom ↔ VilleNom" },
          ],
          required: true,
        },
      ],
    };
    const res = autoAnswerConfirmations(plan, ctxBase, "demoEval");
    expect(res.decisions["joinKey:j1"]).toBe("CodeVille|VilleCode");
  });

  test("auto-answer uses prompt-aware column selection and direction", () => {
    const plan: AgentPlan = {
      version: "1.0",
      goal: "Afficher par Début EDP croissant",
      
      steps: [
        {
          id: "tv1",
          macro: "table_view",
          params: { sort: { col: "needs_confirmation" }, source: { blockRef: "Sheet1!A1:C4" } },
        },
      ],
      confirmations: [
        {
          id: "tv1:sort_col",
          question: "Choisir la colonne de tri",
          choices: [
            { id: "c0", label: "Projet" },
            { id: "c1", label: "Début EDP" },
          ],
          required: true,
        },
      ],
    };
    const res = autoAnswerConfirmations(plan, ctxBase, "demoEval");
    expect(res.decisions["tv1:sort_col"]).toBe("Début EDP");
    expect(res.decisions["tv1:sort_dir"]).toBe("asc");
    expect(res.plan.confirmations?.length || 0).toBe(0);
  });
});

describe("auto-answer column labels", () => {
  test("stores column labels so the sort column remains Début EDP", () => {
    const contextWithDebut: WorkbookContextSnapshot = {
      ...ctxBase,
      sheets: ctxBase.sheets.map((sheet) => ({
        ...sheet,
        blocks: sheet.blocks.map((block) => ({
          ...block,
          headers: sheet.name === "Sheet1" ? ["Projet", "Début EDP", "Ville"] : block.headers,
        })),
      })),
    };

    const plan: AgentPlan = {
      version: "1.0",
      goal: "sort by Début EDP",
      
      steps: [
        {
          id: "tv-demo",
          macro: "table_view",
          params: { sort: { col: "needs_confirmation" }, source: { blockRef: "Sheet1!A1:C4" } },
        },
      ],
      confirmations: [
        {
          id: "tv-demo:sort_col",
          question: "Choisir la colonne de tri",
          choices: [
            { id: "c0", label: "Projet" },
            { id: "c1", label: "Début EDP" },
            { id: "c2", label: "Ville" },
          ],
          required: true,
        },
      ],
    };

    const resolved = autoAnswerConfirmations(plan, contextWithDebut, "demoEval");
    const finalPlan = applyConfirmationsToPlan(resolved.plan, resolved.decisions, contextWithDebut);
    const sortCol = finalPlan.steps?.[0]?.params?.sort?.col;
    expect(sortCol).toBe("Début EDP");
  });
});

