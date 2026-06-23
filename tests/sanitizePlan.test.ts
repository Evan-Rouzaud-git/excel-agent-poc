import { sanitizePlan } from "../src/taskpane/agent/planner/sanitizePlan";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";
import { validatePlan } from "../src/taskpane/agent/planSchema";

describe("sanitizePlan", () => {
  test("removes artifacts and extra fields", () => {
    const raw = {
      version: "1.0",
      goal: "g",
      
      steps: [
        {
          id: "s1",
          macro: "create_chart",
          params: { source: { blockRef: "Sheet1!A1:C3" }, mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: 2 }] }, chartType: "columnClustered" },
          extra: "nope",
        },
      ],
      artifacts: [{ id: "bad", type: "chart", sheet: "x", anchor: "A1", fromStep: "s1" }],
      foo: "bar",
    };
    const clean = sanitizePlan(raw);
    expect((clean as any).artifacts).toBeUndefined();
    expect((clean as any).foo).toBeUndefined();
    expect(clean.steps[0].extra).toBeUndefined();
    const val = validatePlan(clean);
    expect(val.valid).toBe(true);
  });

  test("injects validate_data confirmation templates when present", () => {
    const plan = {
      version: "1.0",
      goal: "validate_data metadata",
      steps: [
        {
          id: "validate1",
          macro: "validate_data",
          params: {
            source: { blockRef: "Sheet1!A1:B2" },
          },
        },
      ],
    };
    const sanitized = sanitizePlan(plan as any);
    const ids = (sanitized.confirmations || []).map((c: any) => c.id);
    expect(ids).toEqual(
      expect.arrayContaining([
        "validate_data:fix_missing",
        "validate_data:fix_duplicates",
        "validate_data:fix_bad_type",
      ])
    );
    const missingConfirm = (sanitized.confirmations || []).find((c: any) => c.id === "validate_data:fix_missing");
    expect(missingConfirm?.required).toBe(true);
  });

  test("converts string confirmation choices into id/label objects", () => {
    const plan = {
      version: "1.0",
      goal: "confirm string choices",
      steps: [],
      confirmations: [{ id: "c1", question: "Pick one", choices: ["a", "b"] }],
    };
    const sanitized = sanitizePlan(plan as any);
    const choices = sanitized.confirmations?.[0]?.choices;
    expect(choices).toEqual([
      { id: "c1", label: "a" },
      { id: "c2", label: "b" },
    ]);
  });

  test("keeps table_view filter and sort when valid", () => {
    const plan = {
      version: "1.0",
      goal: "Test plan",
      
      steps: [
        { id: "s1", macro: "table_view", params: { source: { blockRef: "B1" }, select: ["A"], filter: [{ col: "A", op: "gt", value: 1 }], sort: { col: "A", dir: "asc" }, dest: { mode: "newSheet" }, unexpected: "x" } },
      ],
    };
    const context = { sheets: [{ blocks: [{ id: "B1", headers: ["A", "B"] }] }], active: {} };
    const out = sanitizePlan(plan as any, context as any);
    expect(out.steps[0].params.filter?.length).toBe(1);
    expect(out.steps[0].params.sort?.col).toBe("A");
  });
});

describe("sanitizePlan join column alias tolerance", () => {
  const buildPlan = (selectCol: string, filterCol: string, sortCol: string) => ({
    version: "1.0",
    goal: "alias tolerance",
    steps: [
      {
        id: "join1",
        macro: "join_tables",
        params: {
          left: { blockRef: "Left!A1:B2" },
          right: { blockRef: "Right!A1:B2" },
          keys: [{ left: "Projet", right: "ptvx_id" }],
        },
      },
      {
        id: "view1",
        macro: "table_view",
        params: {
          source: { artifactRef: "join1" },
          select: [selectCol, "Ville"],
          filter: [{ col: filterCol, op: "eq", value: "X" }],
          sort: { col: sortCol, dir: "asc" },
          dest: { mode: "newSheet" },
        },
      },
    ],
  });

  const context = { sheets: [], active: {} } as any;

  test("remaps a right-side header alias to the actual join output and keeps the request name", () => {
    const plan = buildPlan("ptvx_id", "ptvx_id", "ptvx_id");
    const sanitized = sanitizePlan(plan as any, context, [], { join1: ["Projet", "Ville"] });
    const viewStep = sanitized.steps.find((s: any) => s.id === "view1");
    expect(viewStep).toBeTruthy();
    const params = viewStep.params;
    expect(params.select).toContain("Projet");
    expect(params.filter?.[0]?.col).toBe("Projet");
    expect(params.sort?.col).toBe("Projet");
    expect(params.rename).toEqual(expect.objectContaining({ Projet: "ptvx_id" }));
  });

  test("remaps the left-side header when the join kept the name and preserves the alias", () => {
    const plan = buildPlan("ptvx_id", "ptvx_id", "ptvx_id");
    const sanitized = sanitizePlan(plan as any, context, [], { join1: ["Projet", "Ville"] });
    const viewStep = sanitized.steps.find((s: any) => s.id === "view1");
    expect(viewStep).toBeTruthy();
    const params = viewStep.params;
    expect(params.select).toContain("Projet");
    expect(params.filter?.[0]?.col).toBe("Projet");
    expect(params.sort?.col).toBe("Projet");
    expect(params.rename).toEqual(expect.objectContaining({ Projet: "ptvx_id" }));
  });
});

describe("sanitizePlan joinKey grounding", () => {
  test("builds confirmation choices from real headers when keys invalid", () => {
    const context: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: {
        sheetName: "Travaux",
        selectionAddress: "A1",
        selectionInBlockId: "Travaux!A1:E5",
        nearestBlockId: "Travaux!A1:E5",
      },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Travaux",
          usedRange: "A1:E5",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4, address: "Travaux!A1:E5" },
          counts: { tables: 1, charts: 0 },
          tables: [
            {
              name: "TravauxTable",
              address: "Travaux!A1:E5",
              dataBodyAddress: "Travaux!A2:E5",
              headerAddress: "Travaux!A1:E5",
              headers: ["ptvx_id", "charge_aff_code", "tache_nom", "m2", "Hab"],
            },
          ],
          blocks: [
            {
              id: "Travaux!A1:E5",
              address: "Travaux!A1:E5",
              kind: "table",
              confidence: 1,
              headerRowIndex: 0,
              headers: ["ptvx_id", "charge_aff_code", "tache_nom", "m2", "Hab"],
              columnTypes: ["text", "text", "text", "number", "number"],
              preview: [],
              source: { type: "table", tableName: "TravauxTable", tableAddress: "Travaux!A1:E5" },
            },
          ],
          charts: [],
          limitations: [],
        },
        {
          name: "Projet",
          usedRange: "A1:E5",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4, address: "Projet!A1:E5" },
          counts: { tables: 1, charts: 0 },
          tables: [
            {
              name: "ProjetTable",
              address: "Projet!A1:E5",
              dataBodyAddress: "Projet!A2:E5",
              headerAddress: "Projet!A1:E5",
              headers: ["Projet", "code", "Typologie", "m2", "Hab"],
            },
          ],
          blocks: [
            {
              id: "Projet!A1:E5",
              address: "Projet!A1:E5",
              kind: "table",
              confidence: 1,
              headerRowIndex: 0,
              headers: ["Projet", "code", "Typologie", "m2", "Hab"],
              columnTypes: ["text", "text", "text", "number", "number"],
              preview: [],
              source: { type: "table", tableName: "ProjetTable", tableAddress: "Projet!A1:E5" },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 2, tables: 2, charts: 0, blocks: 2, durationMs: 0 },
    };

    const plan = {
      version: "1.0",
      goal: "Associe Travaux et Projet",
      steps: [
        {
          id: "join_step",
          macro: "join_tables",
          params: {
            left: { blockRef: "Travaux!A1:E5" },
            right: { blockRef: "Projet!A1:E5" },
            keys: [{ left: "project_id", right: "proj_id" }],
          },
        },
      ],
    };
    const sanitized = sanitizePlan(
      plan as any,
      context,
      [],
      undefined,
      "Associe Travaux et Projet sur ptvx_id vs Projet en gardant m2, Hab"
    );
    const joinStep = sanitized.steps.find((s: any) => s.id === "join_step");
    expect(joinStep?.params?.keys?.[0]).toMatchObject({ left: "ptvx_id", right: "Projet" });
    const joinConfirm = sanitized.confirmations?.find((c: any) => c.id === "joinKey:join_step");
    expect(joinConfirm).toBeTruthy();
    expect(joinConfirm?.choices?.[0]?.id).toBe("ptvx_id|Projet");
  });
});

describe("sanitizePlan join placeholder guard", () => {
    const context = {
      sheets: [
        {
          name: "LeftSheet",
          blocks: [
            {
              id: "LeftSheet!A1:B2",
              address: "LeftSheet!A1:B2",
              kind: "table",
              confidence: 1,
              headers: ["Projet", "Code"],
              columnTypes: ["text", "text"],
              preview: [],
              headerRowIndex: 0,
              source: { type: "table", tableName: "LeftTable", tableAddress: "LeftSheet!A1:B2" },
            },
          ],
          tables: [{ name: "LeftTable", address: "LeftSheet!A1:B2", dataBodyAddress: "LeftSheet!A1:B2", headerAddress: "LeftSheet!A1:B2", headers: ["Projet", "Code"] }],
        charts: [],
        limitations: [],
        usedRange: null,
        valueBounds: { firstRow: null, firstCol: null, lastRow: null, lastCol: null, address: null },
        counts: { tables: 0, charts: 0 },
      },
      {
        name: "RightSheet",
        blocks: [
          {
            id: "RightSheet!A1:B2",
            address: "RightSheet!A1:B2",
            kind: "table",
            confidence: 1,
            headers: ["ptvx_id", "Projet"],
            columnTypes: ["text", "text"],
            preview: [],
            headerRowIndex: 0,
            source: { type: "table", tableName: "RightTable", tableAddress: "RightSheet!A1:B2" },
          },
        ],
        tables: [{ name: "RightTable", address: "RightSheet!A1:B2", dataBodyAddress: "RightSheet!A1:B2", headerAddress: "RightSheet!A1:B2", headers: ["ptvx_id", "Projet"] }],
        charts: [],
        limitations: [],
        usedRange: null,
        valueBounds: { firstRow: null, firstCol: null, lastRow: null, lastCol: null, address: null },
        counts: { tables: 0, charts: 0 },
      },
    ],
    active: { sheetName: null, selectionAddress: null, selectionInBlockId: null, nearestBlockId: null },
  } as any;

  test("replaces placeholder join keys with a fallback pair and injects confirmation", () => {
    const plan = {
      version: "1.0",
      goal: "resolve join placeholder",
      steps: [
          {
            id: "join1",
            macro: "join_tables",
            params: {
              left: { blockRef: "LeftSheet!A1:B2" },
              right: { blockRef: "RightSheet!A1:B2" },
              keys: [{ left: "?join_key?", right: "???" }],
            },
        },
      ],
    };
    const warnings: string[] = [];
    const sanitized = sanitizePlan(plan as any, context, warnings);
    const joinStep = sanitized.steps.find((s: any) => s.id === "join1");
    expect(joinStep.params.keys).toHaveLength(1);
    expect(joinStep.params.keys[0]).toMatchObject({ left: "Projet", right: "ptvx_id" });
    expect(warnings).toContain("join_keys_placeholder_resolved:join1");
    const confirm = sanitized.confirmations?.find((c: any) => c.id === "joinKey:join1");
    expect(confirm).toBeTruthy();
    expect(confirm?.required).toBe(true);
    expect(confirm?.choices?.map((choice: any) => choice.id)).toEqual(
      expect.arrayContaining(["Projet|ptvx_id"])
    );
  });
});
