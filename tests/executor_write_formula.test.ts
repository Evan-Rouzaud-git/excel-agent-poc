import { executePlan } from "../src/taskpane/agent/executor";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

const snapshot: WorkbookContextSnapshot = {
  workbook: { name: "Book", readOnly: false },
  active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: null, nearestBlockId: null },
  capabilities: [],
  limitations: [],
  sheets: [
    {
      name: "Sheet1",
      usedRange: "A1:B3",
      valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
      counts: { tables: 1, charts: 0 },
      tables: [
        { name: "Ventes", address: "Sheet1!A1:B3", dataBodyAddress: "Sheet1!A2:B3", headerAddress: "Sheet1!A1:B1", headers: ["Prix", "Cout"] },
      ],
      blocks: [
        {
          id: "Sheet1!A1:B3",
          address: "A1:B3",
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["Prix", "Cout"],
          columnTypes: ["number", "number"],
          preview: [],
          source: { type: "table", tableName: "Ventes", tableAddress: "Sheet1!A1:B3" },
        },
      ],
      charts: [],
      limitations: [],
    },
  ],
  totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
};

const plan = {
  version: "1.0",
  goal: "write formula",
  
  steps: [
    {
      id: "s1",
      macro: "write_formula",
      params: {
        target: { blockRef: "Sheet1!A1:B3", writeMode: "newColumnRight", headerName: "Marge" },
        formula: "=[@Prix]-[@Cout]",
      },
    },
  ],
};

describe("executor - write_formula", () => {
  test("adds new column with formulas", async () => {
    const ws = new FakeWorksheet("Sheet1");
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    expect(ws.getRange("C1").values?.[0]?.[0]).toBe("Marge");
    expect(ws.getRange("C2:C3").formulas).toEqual([["=[@Prix]-[@Cout]"], ["=[@Prix]-[@Cout]"]]);
  });

  test("adds formula on range block when no table", async () => {
    const rangeSnapshot: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: { sheetName: "Sheet1", selectionAddress: "Sheet1!G13", selectionInBlockId: null, nearestBlockId: null },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Sheet1",
          usedRange: "E6:F10",
          valueBounds: { firstRow: 5, firstCol: 4, lastRow: 9, lastCol: 5, address: "E6:F10" },
          counts: { tables: 0, charts: 0 },
          tables: [],
          blocks: [
            {
              id: "Sheet1!E6:F10",
              address: "E6:F10",
              kind: "range",
              confidence: 0.65,
              headerRowIndex: 0,
              headers: ["Revenus", "Depenses"],
              columnTypes: ["number", "number"],
              preview: [],
              source: { type: "range" },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
    };

    const rangePlan = {
      version: "1.0",
      goal: "write formula",
      
      steps: [
        {
          id: "s1",
          macro: "write_formula",
          params: {
            target: { headerName: "Marge" },
            formula: "=[@Revenus]-[@Depenses]",
          },
        },
      ],
    };

    const ws = new FakeWorksheet("Sheet1");
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(rangePlan as any, rangeSnapshot, ctx as any, {});

    expect(res.status).toBe("ok");
    expect(ws.getRange("G6").values?.[0]?.[0]).toBe("Marge");
    expect(ws.getRange("G7:G10").formulas).toEqual([["=E7-F7"], ["=E8-F8"], ["=E9-F9"], ["=E10-F10"]]);
  });

  test("write_formula on a range uses relative data row refs and IFERROR division wrapper", async () => {
    const rangeSnapshot: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: { sheetName: "Sheet1", selectionAddress: "Sheet1!C12", selectionInBlockId: "Sheet1!C12:E16", nearestBlockId: "Sheet1!C12:E16" },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Sheet1",
          usedRange: "C12:F16",
          valueBounds: { firstRow: 11, firstCol: 2, lastRow: 15, lastCol: 5, address: "C12:F16" },
          counts: { tables: 0, charts: 0 },
          tables: [],
          blocks: [
            {
              id: "Sheet1!C12:E16",
              address: "C12:E16",
              kind: "range",
              confidence: 0.9,
              headerRowIndex: 11,
              headers: ["Projet", "m2", "Hab"],
              columnTypes: ["text", "number", "number"],
              preview: [],
              source: { type: "range" },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
    };

    const plan = {
      version: "1.0",
      goal: "compute density on range",

      steps: [
        {
          id: "r1",
          macro: "write_formula",
          params: {
            target: { blockRef: "Sheet1!C12:E16", writeMode: "newColumnRight", headerName: "Densite" },
            formula: "=[m2]/[Hab]",
            fillDown: true,
          },
        },
      ],
    };

    const ws = new FakeWorksheet("Sheet1");
    ws.getRange("C12:E16").values = [
      ["Projet", "m2", "Hab"],
      ["P1", 100, 10],
      ["P2", 200, 20],
      ["P3", 150, 30],
      ["P4", 80, 40],
    ];
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(plan as any, rangeSnapshot, ctx as any, {});

    expect(res.status).toBe("ok");
    expect(ws.getRange("F12").values?.[0]?.[0]).toBe("Densite");
    expect(ws.getRange("F13:F16").formulas).toEqual([
      ["=IFERROR((--D13)/(--E13),\"\")"],
      ["=IFERROR((--D14)/(--E14),\"\")"],
      ["=IFERROR((--D15)/(--E15),\"\")"],
      ["=IFERROR((--D16)/(--E16),\"\")"],
    ]);
  });

  test("range formula with missing header logs unknown_headers instead of structured_ref", async () => {
    const rangeSnapshot: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: {
        sheetName: "Sheet1",
        selectionAddress: "Sheet1!E6",
        selectionInBlockId: null,
        nearestBlockId: null,
      },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Sheet1",
          usedRange: "E6:F10",
          valueBounds: { firstRow: 5, firstCol: 4, lastRow: 9, lastCol: 5, address: "E6:F10" },
          counts: { tables: 0, charts: 0 },
          tables: [],
          blocks: [
            {
              id: "Sheet1!E6:F10",
              address: "E6:F10",
              kind: "range",
              confidence: 0.9,
              headerRowIndex: 5,
              headers: ["Revenus", "Depenses"],
              columnTypes: ["number", "number"],
              preview: [],
              source: { type: "range" },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
    };

    const plan = {
      version: "1.0",
      goal: "percentage margin",
      steps: [
        {
          id: "f1",
          macro: "write_formula",
          params: {
            target: { blockRef: "Sheet1!E6:F10", writeMode: "newColumnRight", headerName: "% Marge" },
            formula: "=[Marge]/[Revenus]",
            fillDown: true,
          },
        },
      ],
    };

    const ws = new FakeWorksheet("Sheet1");
    ws.getRange("E6:F10").values = [
      ["Revenus", "Depenses"],
      [100, 80],
      [120, 90],
      [140, 100],
      [160, 110],
    ];
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(plan as any, rangeSnapshot, ctx as any, {});

    expect(res.status).toBe("error");
    expect(res.logs.some((log) => log.message.includes("structured_ref_missing_at"))).toBe(false);
    expect(res.logs.some((log) => log.message.includes("unknown_headers_in_formula Marge"))).toBe(true);
  });

  test("write_formula on a table rewrites structured refs to ThisRow", async () => {
    const tableSnapshot: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: {
        sheetName: "Sheet1",
        selectionAddress: "Sheet1!A1",
        selectionInBlockId: "Sheet1!A1:C3",
        nearestBlockId: "Sheet1!A1:C3",
      },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Sheet1",
          usedRange: "A1:D3",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 3, address: "A1:D3" },
          counts: { tables: 1, charts: 0 },
          tables: [
            {
              name: "Dataset",
              address: "Sheet1!A1:C3",
              dataBodyAddress: "Sheet1!A2:C3",
              headerAddress: "Sheet1!A1:C1",
              headers: ["Projet", "m2", "Hab"],
            },
          ],
          blocks: [
            {
              id: "Sheet1!A1:C3",
              address: "A1:C3",
              kind: "table",
              confidence: 1,
              headerRowIndex: 0,
              headers: ["Projet", "m2", "Hab"],
              columnTypes: ["text", "number", "number"],
              preview: [],
              source: { type: "table", tableName: "Dataset", tableAddress: "Sheet1!A1:C3" },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
    };

    const plan = {
      version: "1.0",
      goal: "write table formula",

      steps: [
        {
          id: "s1",
          macro: "write_formula",
          params: {
            target: { blockRef: "Sheet1!A1:C3", writeMode: "newColumnRight", headerName: "Densite" },
            formula: "=[m2]/[Hab]",
            fillDown: true,
          },
        },
      ],
    };

    const ws = new FakeWorksheet("Sheet1");
    ws.getRange("A1:C3").values = [
      ["Projet", "m2", "Hab"],
      ["P1", 100, 10],
      ["P2", 150, 15],
    ];
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const res = await executePlan(plan as any, tableSnapshot, ctx as any, {});

    expect(res.status).toBe("ok");
    expect(ws.getRange("D2:D3").formulas).toEqual([["=[@m2]/[@Hab]"], ["=[@m2]/[@Hab]"]]);
  });

  test("write_formula skips retarget when explicit blockRef is resolvable", async () => {
    const ws = new FakeWorksheet("Sheet1");
    ws.getRange("A1:B3").values = [
      ["Projet", "Ville"],
      ["P1", "Paris"],
      ["P2", "Lyon"],
    ];
    const ctx = new FakeContext(new FakeWorkbook([ws]));
    const viewSnapshot: WorkbookContextSnapshot = {
      workbook: { name: "Book", readOnly: false },
      active: {
        sheetName: "Sheet1",
        selectionAddress: "Sheet1!A1",
        selectionInBlockId: "Sheet1!A1:B3",
        nearestBlockId: "Sheet1!A1:B3",
      },
      capabilities: [],
      limitations: [],
      sheets: [
        {
          name: "Sheet1",
          usedRange: "A1:B3",
          valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1, address: "A1:B3" },
          counts: { tables: 1, charts: 0 },
          tables: [
            {
              name: "ProjetTbl",
              address: "Sheet1!A1:B3",
              dataBodyAddress: "Sheet1!A2:B3",
              headerAddress: "Sheet1!A1:B1",
              headers: ["Projet", "Ville"],
            },
          ],
          blocks: [
            {
              id: "Sheet1!A1:B3",
              address: "A1:B3",
              kind: "table",
              confidence: 1,
              headerRowIndex: 0,
              headers: ["Projet", "Ville"],
              columnTypes: ["text", "text"],
              preview: [],
              source: { type: "table", tableName: "ProjetTbl", tableAddress: "Sheet1!A1:B3" },
            },
          ],
          charts: [],
          limitations: [],
        },
      ],
      totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
    };
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "write densite to Projet with explicit blockRef",

        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef: "Sheet1!A1:B3" }, select: ["Projet", "Ville"], dest: { mode: "newSheet" } },
          },
          {
            id: "f1",
            macro: "write_formula",
            params: {
              target: { blockRef: "Sheet1!A1:B3", writeMode: "newColumnRight", headerName: "Densite" },
              formula: "=[@Projet]",
              fillDown: true,
            },
          },
        ],
      },
      viewSnapshot
    );
    const res = await executePlan(plan as any, viewSnapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    expect(ws.getRange("C1").values?.[0]?.[0]).toBe("Densite");
    const viewArtifact = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    if (viewArtifact?.sheet) {
      const viewSheet = ctx.workbook.worksheets.getItem(viewArtifact.sheet);
      expect(viewSheet.getRange("C1").values?.[0]?.[0]).not.toBe("Densite");
    }
    expect(
      res.logs.some((log) => log.message === "write_formula retarget skipped because explicit target provided")
    ).toBe(true);
  });
});

