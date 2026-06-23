import { executePlan } from "../src/taskpane/agent/executor";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";

const colToLetter = (col: number) => {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

function makeSnapshot(sheetName: string, data: any[][], tableName = "Tbl"): { snapshot: WorkbookContextSnapshot; ctx: FakeContext; blockRef: string } {
  const rows = data.length;
  const cols = data[0]?.length || 1;
  const endCell = `${colToLetter(cols - 1)}${rows}`;
  const address = `${sheetName}!A1:${endCell}`;
  const ws = new FakeWorksheet(sheetName);
  ws.getRange(`A1:${endCell}`).values = data;
  const wb = new FakeWorkbook([ws]);
  const ctx = new FakeContext(wb);
  const snapshot: WorkbookContextSnapshot = {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName, selectionAddress: "A1", selectionInBlockId: address, nearestBlockId: address },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: sheetName,
        usedRange: `A1:${endCell}`,
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: rows - 1, lastCol: cols - 1, address: `A1:${endCell}` },
        counts: { tables: 1, charts: 0 },
        tables: [
          {
            name: tableName,
            address,
            dataBodyAddress: `${sheetName}!A2:${colToLetter(cols - 1)}${rows}`,
            headerAddress: `${sheetName}!A1:${colToLetter(cols - 1)}1`,
            headers: data[0] as string[],
          },
        ],
        blocks: [
          {
            id: address,
            address: `A1:${endCell}`,
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: data[0] as string[],
            columnTypes: new Array(cols).fill("text"),
            preview: [],
            source: { type: "table", tableName, tableAddress: address },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
  };
  return { snapshot, ctx, blockRef: address };
}

function rangeFromArtifact(ctx: FakeContext, artifact: any) {
  const ws = ctx.workbook.worksheets.getItem(artifact.sheet);
  const anchor = artifact.anchor || artifact.address || "A1";
  const rows = (artifact.rows || 0) + 1;
  const cols = artifact.cols || 1;
  const matchLetters = anchor.match(/^[A-Z]+/i);
  const matchDigits = anchor.match(/[0-9]+/);
  const startCol = (() => {
    const letters = (matchLetters?.[0] || "A").toUpperCase();
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  })();
  const startRow = matchDigits ? parseInt(matchDigits[0], 10) - 1 : 0;
  const endCol = startCol + cols - 1;
  const endRow = startRow + rows - 1;
  const endCell = `${colToLetter(endCol)}${endRow + 1}`;
  const addr = `${anchor}:${endCell}`;
  return ws.getRange(addr);
}

function expectOk(res: any) {
  if (res.status !== "ok") {
    // eslint-disable-next-line no-console
    console.error("execution failed", res.status, res.errors, res.logs);
  }
  expect(res.status).toBe("ok");
}

describe("table_view macro", () => {
  test("T1 Projection simple (\"Garde uniquement Projet, Ville, Début EDP, Fin EDP\")", async () => {
    const prompt = "Garde uniquement Projet, Ville, Début EDP, Fin EDP";
    const data = [
      ["Projet", "Ville", "Début EDP", "Fin EDP", "Budget"],
      ["P1", "Paris", 1, 3, 100],
      ["P2", "Lyon", 2, 4, 200],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("Sheet1", data, "Src");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: prompt,
        
        steps: [
          { id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Projet", "Ville", "Début EDP", "Fin EDP"], dest: { mode: "newSheet" } } },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expectOk(res);
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    expect(art?.blockRef).toBeTruthy();
    const out = rangeFromArtifact(ctx, art!);
    expect(out.values[0]).toEqual(["Projet", "Ville", "Début EDP", "Fin EDP"]);
    expect(out.values.length).toBe(3);
  });

  test("T2 Rename header", async () => {
    const prompt = "Renomme 'Fin EDP' en 'Fin Etude de Prix' et garde Projet/Ville";
    const data = [
      ["Projet", "Ville", "Fin EDP"],
      ["P1", "Paris", 10],
      ["P2", "Lyon", 20],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("Sheet1", data, "Src");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: prompt,
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Projet", "Ville", "Fin EDP"], rename: { "Fin EDP": "Fin Etude de Prix" }, dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expectOk(res);
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    const out = rangeFromArtifact(ctx, art!);
    expect(out.values[0][2]).toBe("Fin Etude de Prix");
  });

  test("T3 Filter notEmpty", async () => {
    const prompt = "Garde les lignes où Ville n’est pas vide";
    const data = [
      ["Projet", "Ville"],
      ["P1", "Paris"],
      ["P2", ""],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("Sheet1", data, "Src");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: prompt,
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Projet", "Ville"], filter: [{ col: "Ville", op: "notEmpty" }], dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expectOk(res);
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    const out = rangeFromArtifact(ctx, art!);
    expect(out.values.length).toBe(2); // header + 1 row
    expect(out.values[1]).toEqual(["P1", "Paris"]);
  });

  test("T4 Filter contains", async () => {
    const prompt = "Garde les projets dont la Ville contient 'Mar'";
    const data = [
      ["Projet", "Ville"],
      ["P1", "Marseille"],
      ["P2", "Paris"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("Sheet1", data, "Src");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: prompt,
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Projet", "Ville"], filter: [{ col: "Ville", op: "contains", value: "Mar" }], dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expectOk(res);
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    const out = rangeFromArtifact(ctx, art!);
    expect(out.values.length).toBe(2);
    expect(out.values[1][1]).toBe("Marseille");
  });

  test("T4b Filter not_contains case-insensitive trim", async () => {
    const data = [
      ["Projet", "Typologie"],
      ["P1", "MI"],
      ["P2", "Logement"],
      ["P3", "  mi  "],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("SheetNC", data, "Src");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Extrait typologie ne contient pas MI",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Projet", "Typologie"], filter: [{ col: "Typologie", op: "not_contains", value: "mi" }], dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expectOk(res);
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    const out = rangeFromArtifact(ctx, art!);
    expect(out.values.length).toBe(2);
    expect(out.values[1][1]).toBe("Logement");
  });

  test("T5 Sort asc sur Début EDP", async () => {
    const prompt = "Trie par Début EDP croissant";
    const data = [
      ["Projet", "Début EDP"],
      ["P1", 5],
      ["P2", 2],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("Sheet1", data, "Src");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: prompt,
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Projet", "Début EDP"], sort: { col: "Début EDP", dir: "asc" }, dest: { mode: "right" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    const out = rangeFromArtifact(ctx, art!);
    expect(out.values[1][0]).toBe("P2");
    expect(out.values[2][0]).toBe("P1");
  });

  test("T6 Chaînage artifactRef join -> table_view -> write_formula", async () => {
    const prompt = "join puis vue et formule";
    const left = [
      ["Code", "Ville"],
      ["A", "Paris"],
      ["B", "Lyon"],
    ];
    const right = [
      ["Code", "Budget"],
      ["A", 100],
      ["B", 200],
    ];
    const leftSetup = makeSnapshot("Left", left, "LeftTbl");
    const rightSetup = makeSnapshot("Right", right, "RightTbl");
    // merge snapshots into one workbook/context
    const wb = new FakeWorkbook([new FakeWorksheet("Left"), new FakeWorksheet("Right")]);
    wb.worksheets.getItem("Left").getRange("A1:B3").values = left;
    wb.worksheets.getItem("Right").getRange("A1:B3").values = right;
    const ctx = new FakeContext(wb);
    const snapshot: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: { sheetName: "Left", selectionAddress: "A1", selectionInBlockId: leftSetup.blockRef, nearestBlockId: leftSetup.blockRef },
      capabilities: [],
      limitations: [],
      sheets: [...leftSetup.snapshot.sheets, ...rightSetup.snapshot.sheets],
      totals: { sheets: 2, tables: 2, charts: 0, blocks: 2, durationMs: 0 },
    };
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: prompt,
        
        steps: [
          { id: "join1", macro: "join_tables", params: { left: { blockRef: leftSetup.blockRef }, right: { blockRef: rightSetup.blockRef }, keys: [{ left: "Code", right: "Code" }] } },
          { id: "view1", macro: "table_view", params: { source: { artifactRef: "join1" }, select: ["Code", "Ville", "Budget"], dest: { mode: "newSheet" } } },
          {
            id: "formula1",
            macro: "write_formula",
            params: { target: { artifactRef: "view1", writeMode: "newColumnRight", headerName: "Budget*2" }, formula: "=[@Budget]*2", fillDown: true },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expectOk(res);
    const viewArt = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    expect((viewArt as any).headers).toContain("Budget*2");
  });

  test("T7 Rename only happens in place (no new sheet)", async () => {
    const data = [
      ["Code", "Ville"],
      ["A", "Paris"],
      ["B", "Lyon"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("Sheet1", data, "SrcTbl");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Renomme Code en Projet",
        
        steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef }, rename: { Code: "Projet" } } }],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    // no extra sheet created
    expect(ctx.workbook.worksheets.rawItems().length).toBe(1);
    const range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B3");
    expect(range.values[0][0]).toBe("Projet");
    expect(range.values[0][1]).toBe("Ville");
  });

  test("T8 write_formula retargets to latest table_view artifact when target missing", async () => {
    const data = [
      ["Projet", "Ville"],
      ["P1", "Paris"],
      ["P2", "Lyon"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("Src", data, "SrcTbl");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Vue puis formule",
        
        steps: [
          { id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Projet", "Ville"], dest: { mode: "newSheet" } } },
          { id: "f1", macro: "write_formula", params: { target: { writeMode: "newColumnRight", headerName: "VilleUpper" }, formula: "=[@Ville]", fillDown: true } },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const viewArt = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    expect((viewArt as any).headers).toContain("VilleUpper");
  });

  test("T10 filter gt creates view", async () => {
    const data = [
      ["Projet", "Hab"],
      ["P1", 10],
      ["P2", 30],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("SheetF", data, "SrcTbl");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Filtrer Hab > 25",
        
        steps: [
          { id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Projet", "Hab"], filter: [{ col: "Hab", op: "gt", value: 25 }], dest: { mode: "newSheet" } } },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const viewArt = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const range = rangeFromArtifact(ctx, viewArt);
    expect(range.values.length).toBe(2); // header + one row
    expect(range.values[1][0]).toBe("P2");
  });

  test("T11 view without select fails", async () => {
    const data = [
      ["Projet", "Hab"],
      ["P1", 10],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("SheetG", data, "SrcTbl");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Vue sans select",
        
        steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef }, dest: { mode: "newSheet" } } }],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
  });

  test("T12 inPlace with select should error", async () => {
    const data = [
      ["Projet", "Hab"],
      ["P1", 10],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("SheetH", data, "SrcTbl");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Inplace mais select present",
        
        steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Projet"], dest: { mode: "inPlace" } } }],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
  });

  test("Unknown header triggers confirmation (no silent substitution)", async () => {
    const data = [
      ["Projet", "Ville", "m2"],
      ["P1", "Paris", 10],
      ["P2", "Lyon", 20],
    ];
    const { snapshot, blockRef } = makeSnapshot("SheetU", data, "SrcTbl");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Extrait les colonnes ville et m5",
        
        steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Ville", "m5"], dest: { mode: "newSheet" } } }],
      },
      snapshot
    );
    expect(plan.confirmations?.some((c: any) => c.id === "view1:unknown_header")).toBe(true);
    const select = (plan.steps[0] as any).params.select;
    expect(select).toContain("Ville");
    expect(select).not.toContain("m2");
    expect(select).not.toContain("m5");
  });

  test("T9 compute density then sort in place on the derived column", async () => {
    const data = [
      ["Projet", "m2", "Hab"],
      ["P1", 100, 50],
      ["P2", 80, 60],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot("SheetD", data, "SrcTbl");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Extrait puis calcule densité et trie",
        
        steps: [
          { id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Projet", "m2", "Hab"], dest: { mode: "newSheet" } } },
          {
            id: "formula1",
            macro: "write_formula",
            params: { target: { artifactRef: "view1", writeMode: "newColumnRight", headerName: "Densité" }, formula: "=[@Hab]/[@m2]", fillDown: true },
          },
          { id: "sort1", macro: "table_view", params: { source: { artifactRef: "view1" }, sort: { col: "Densité", dir: "desc" }, dest: { mode: "inPlace" } } },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const baseArt = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    expect(baseArt).toBeTruthy();
    expect((baseArt as any)?.headers).toContain("Densité");
  });
});
