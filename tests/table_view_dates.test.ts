import { executePlan } from "../src/taskpane/agent/executor";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { parseDateCell } from "../src/taskpane/agent/dateUtils";
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

function makeSnapshotWithFormats(
  sheetName: string,
  data: any[][],
  numberFormats?: any[][],
  tableName = "Tbl"
): { snapshot: WorkbookContextSnapshot; ctx: FakeContext; blockRef: string } {
  const rows = data.length;
  const cols = data[0]?.length || 1;
  const endCell = `${colToLetter(cols - 1)}${rows}`;
  const address = `${sheetName}!A1:${endCell}`;
  const ws = new FakeWorksheet(sheetName);
  ws.getRange(`A1:${endCell}`).values = data;
  if (numberFormats) ws.getRange(`A1:${endCell}`).numberFormat = numberFormats;
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

describe("table_view dates", () => {
  const excelSerialToFr = (serial: any, opts?: { pad?: boolean }) => {
    if (typeof serial !== "number" || !Number.isFinite(serial)) return `${serial ?? ""}`;
    const base = Date.UTC(1899, 11, 30);
    const msPerDay = 24 * 60 * 60 * 1000;
    const d = new Date(base + Math.floor(serial) * msPerDay);
    const day = d.getUTCDate();
    const month = d.getUTCMonth() + 1;
    const year = d.getUTCFullYear();
    const pad = (n: number) => (opts?.pad ? String(n).padStart(2, "0") : String(n));
    return `${pad(day)}/${pad(month)}/${year}`;
  };

  test("tri asc sur dates texte FR (d/m/yyyy) est correct", async () => {
    const data = [["Date"], ["2/12/2027"], ["7/11/2027"], ["1/10/2028"], ["10/9/2028"]];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("FrSortAsc", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "trier dates FR",
        
        steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Date"], sort: { col: "Date", dir: "asc" }, dest: { mode: "newSheet" } } }],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = rangeFromArtifact(ctx, art);
    const ordered = out.values.slice(1).map((r: any[]) => excelSerialToFr(r[0]));
    expect(ordered).toEqual(["7/11/2027", "2/12/2027", "10/9/2028", "1/10/2028"]);
  });

  test("runtime: text mm/dd mais values serial + format FR => tri/filtre restent FR", async () => {
    const excelSerial = (y: number, m: number, d: number) => {
      const base = Date.UTC(1899, 11, 30);
      const msPerDay = 24 * 60 * 60 * 1000;
      return Math.floor((Date.UTC(y, m - 1, d) - base) / msPerDay);
    };
    // values sont corrects (serial), mais Office.js renverrait parfois un text en mm/dd
    const data = [["Date"], [excelSerial(2027, 12, 2)], [excelSerial(2027, 11, 7)], [excelSerial(2028, 10, 1)], [excelSerial(2028, 9, 10)]];
    const nf = [["dd/mm/yyyy"], ["dd/mm/yyyy"], ["dd/mm/yyyy"], ["dd/mm/yyyy"], ["dd/mm/yyyy"]];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("RuntimeTextMismatch", data, nf);
    const ws = ctx.workbook.worksheets.getItem("RuntimeTextMismatch");
    ws.getRange("A1:A5").text = [["Date"], ["12/2/2027"], ["11/7/2027"], ["10/1/2028"], ["9/10/2028"]];

    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "trier dates FR malgré text mm/dd",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: {
              source: { blockRef },
              select: ["Date"],
              sort: { col: "Date", dir: "asc" },
              filter: [{ col: "Date", op: "between", value: ["01/01/2027", "31/12/2028"] }],
              dest: { mode: "newSheet" },
            },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = rangeFromArtifact(ctx, art);
    const ordered = out.values.slice(1).map((r: any[]) => excelSerialToFr(r[0]));
    expect(ordered).toEqual(["7/11/2027", "2/12/2027", "10/9/2028", "1/10/2028"]);
  });

  test("runtime-like: numberFormat manquant + text mm/dd => doit quand meme filtrer/ trier en FR via serial", async () => {
    const excelSerial = (y: number, m: number, d: number) => {
      const base = Date.UTC(1899, 11, 30);
      const msPerDay = 24 * 60 * 60 * 1000;
      return Math.floor((Date.UTC(y, m - 1, d) - base) / msPerDay);
    };
    // La date voulue est 12/02/2027 (12 fevrier 2027). Si Office.js renvoie le text en mm/dd => "2/12/2027",
    // un parse FR naïf lit "2/12/2027" comme 2 decembre 2027 et le filtre Jan-Mar devient vide.
    const data = [["Projet", "Date"], ["A", excelSerial(2027, 2, 12)], ["B", excelSerial(2027, 7, 11)], ["C", excelSerial(2028, 10, 9)], ["D", excelSerial(2028, 1, 10)]];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("RuntimeNoNF", data);
    const ws = ctx.workbook.worksheets.getItem("RuntimeNoNF");
    ws.getRange("A1:B5").text = [["Projet", "Date"], ["A", "2/12/2027"], ["B", "7/11/2027"], ["C", "10/9/2028"], ["D", "1/10/2028"]];

    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "between janv-mars 2027 + tri asc",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: {
              source: { blockRef },
              select: ["Projet", "Date"],
              sort: { col: "Date", dir: "asc" },
              filter: [{ col: "Date", op: "between", value: ["01/01/2027", "31/03/2027"] }],
              dest: { mode: "newSheet" },
            },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = rangeFromArtifact(ctx, art);
    const keptProjects = out.values.slice(1).map((r: any[]) => r[0]);
    expect(keptProjects).toEqual(["A"]);
    const keptDates = out.values.slice(1).map((r: any[]) => excelSerialToFr(r[1], { pad: true }));
    expect(keptDates).toEqual(["12/02/2027"]);
  });

  test("tri desc sur dates texte FR (d/m/yyyy) est correct", async () => {
    const data = [["Date"], ["2/12/2027"], ["7/11/2027"], ["1/10/2028"], ["10/9/2028"]];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("FrSortDesc", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "trier dates FR desc",
        
        steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["Date"], sort: { col: "Date", dir: "desc" }, dest: { mode: "newSheet" } } }],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = rangeFromArtifact(ctx, art);
    const ordered = out.values.slice(1).map((r: any[]) => excelSerialToFr(r[0]));
    expect(ordered).toEqual(["1/10/2028", "10/9/2028", "2/12/2027", "7/11/2027"]);
  });

  test('filtre "between janvier 2027 et mars 2027" sur dates texte FR (dd/mm/yyyy) ne garde que 12/02/2027', async () => {
    const data = [["Date"], ["11/07/2027"], ["12/02/2027"], ["09/10/2028"], ["10/01/2028"]];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("FrBetweenMinimal", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "filtre janvier-mars 2027",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Date"], filter: [{ col: "Date", op: "between", value: ["01/01/2027", "31/03/2027"] }], dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = rangeFromArtifact(ctx, art);
    const kept = out.values.slice(1).map((r: any[]) => excelSerialToFr(r[0], { pad: true }));
    expect(kept).toEqual(["12/02/2027"]);
  });

  test("filtre dates > 2027 garde les dates et copie le format", async () => {
    const data = [
      ["Projet", "Date"],
      ["P1", new Date("2026-05-01")],
      ["P2", new Date("2028-06-15")],
      ["P3", new Date("2029-01-10")],
    ];
    const nf = [
      ["", ""],
      ["", "dd/mm/yyyy"],
      ["", "dd/mm/yyyy"],
      ["", "dd/mm/yyyy"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("Sheet1", data, nf, "Src");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Filtre pour ne garder que les dates supérieur à 2027",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Projet", "Date"], filter: [{ col: "Date", op: "gt", value: 2027 }], dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    expect(art).toBeTruthy();
    const out = rangeFromArtifact(ctx, art!);
    expect(out.values.length).toBe(3);
    expect(typeof out.values[1][1] === "number" || out.values[1][1] instanceof Date).toBe(true);
    expect(typeof out.values[2][1] === "number" || out.values[2][1] instanceof Date).toBe(true);
    expect(out.numberFormat[1][1]).toBe("dd/mm/yyyy");
    expect(
      res.logs.some((l) => l.message?.includes("table_view_filter_applied") && l.message?.includes("count=2") && l.stepId === "view1")
    ).toBe(true);
    expect(res.logs.some((l) => l.message?.includes("table_view_format_copied") && l.message?.includes("Date"))).toBe(true);
  });

  test("filtre numeric conserve le format date sur les colonnes sélectionnées", async () => {
    const data = [
      ["Projet", "Hab", "Date"],
      ["P1", 10, new Date("2025-01-01")],
      ["P2", 30, new Date("2027-03-03")],
      ["P3", 40, new Date("2028-04-04")],
    ];
    const nf = [
      ["", "", ""],
      ["", "", "yyyy-mm-dd"],
      ["", "", "yyyy-mm-dd"],
      ["", "", "yyyy-mm-dd"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("Sheet1", data, nf, "Src");
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Extrait les projets dont Hab > 25",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: {
              source: { blockRef },
              select: ["Projet", "Date", "Hab"],
              filter: [{ col: "Hab", op: "gt", value: 25 }],
              dest: { mode: "newSheet" },
            },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table");
    expect(art).toBeTruthy();
    const out = rangeFromArtifact(ctx, art!);
    expect(out.values.length).toBe(3); // header + 2 rows
    expect(typeof out.values[1][1] === "number" || out.values[1][1] instanceof Date).toBe(true);
    expect(out.numberFormat[1][1]).toBe("dd/mm/yyyy");
    expect(res.logs.some((l) => l.message?.includes("table_view_format_copied") && l.message?.includes("Date"))).toBe(true);
  });

  test("dates FR dd/mm/yyyy are parsed in filters and sort", async () => {
    const data = [
      ["Projet", "DateFr"],
      ["P1", "11/07/2027"],
      ["P2", "12/02/2027"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("Fr", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "filtre date fr",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: {
              source: { blockRef },
              select: ["Projet", "DateFr"],
              filter: [{ col: "DateFr", op: "gt", value: "01/01/2027", type: "date" }],
              sort: { col: "DateFr", dir: "asc" },
              dest: { mode: "newSheet" },
            },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B3");
    // sort asc should keep P2 (12/02) then P1 (11/07)
    expect(out.values[1][0]).toBe("P2");
    expect(out.values[2][0]).toBe("P1");
  });

  test("parses Excel serial dates", async () => {
    const serialBase = 45000; // roughly 2023
    const data = [
      ["Date", "Val"],
      [serialBase, 1],
      [serialBase + 1, 2],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("Serial", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "filtre serial",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: {
              source: { blockRef },
              select: ["Date", "Val"],
              filter: [{ col: "Date", op: "gte", value: serialBase }],
              sort: { col: "Date", dir: "desc" },
              dest: { mode: "newSheet" },
            },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B3");
    // desc => second row should be larger serial
    expect(out.values[1][0]).toBe(serialBase + 1);
  });

  test("parses ISO date strings", async () => {
    const data = [
      ["DateIso", "Val"],
      ["2027-07-11", 1],
      ["2027-07-01", 2],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("Iso", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "sort iso",
        
        steps: [{ id: "view1", macro: "table_view", params: { source: { blockRef }, select: ["DateIso", "Val"], sort: { col: "DateIso", dir: "asc" }, dest: { mode: "newSheet" } } }],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B3");
    const ts1 = parseDateCell(out.values[1][0]).ts;
    const ts2 = parseDateCell(out.values[2][0]).ts;
    expect(ts1).not.toBeNull();
    expect(ts2).not.toBeNull();
    expect(ts1! <= ts2!).toBe(true);
  });

  test("mix serial and string dates in comparison", async () => {
    const serial = 45000; // ~2023
    const data = [
      ["DateMix", "Val"],
      [serial, 1],
      ["11/07/2027", 2],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("Mix", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "filtre mix",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["DateMix", "Val"], filter: [{ col: "DateMix", op: "gt", value: "01/01/2023", type: "date" }], sort: { col: "DateMix", dir: "asc" }, dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B3");
    const tsFirst = parseDateCell(out.values[1][0]).ts;
    const tsSecond = parseDateCell(out.values[2][0]).ts;
    expect(tsFirst).not.toBeNull();
    expect(tsSecond).not.toBeNull();
    expect(tsFirst! < tsSecond!).toBe(true);
  });

  test("string dates dd/mm/yyyy sorted FR even without formats", async () => {
    const data = [
      ["Projet", "Date"],
      ["A", "2/12/2027"], // 2 dec 2027
      ["B", "7/11/2027"], // 7 nov 2027
      ["C", "1/10/2028"], // 1 oct 2028
      ["D", "10/9/2028"], // 10 sep 2028
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("FrStr", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "tri fr asc",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Projet", "Date"], sort: { col: "Date", dir: "asc" }, dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B5");
    const ordered = out.values.slice(1).map((r: any[]) => r[0]);
    expect(ordered).toEqual(["B", "A", "D", "C"]);
  });

  test("string dates dd/mm/yyyy sorted FR desc even without formats", async () => {
    const data = [
      ["Projet", "Date"],
      ["A", "2/12/2027"], // 2 dec 2027
      ["B", "7/11/2027"], // 7 nov 2027
      ["C", "1/10/2028"], // 1 oct 2028
      ["D", "10/9/2028"], // 10 sep 2028
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("FrStrDesc", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "tri fr desc",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: { source: { blockRef }, select: ["Projet", "Date"], sort: { col: "Date", dir: "desc" }, dest: { mode: "newSheet" } },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B5");
    const ordered = out.values.slice(1).map((r: any[]) => r[0]);
    expect(ordered).toEqual(["C", "D", "A", "B"]);
  });

  test("between on string dates accepts serial bounds", async () => {
    const data = [
      ["Projet", "Date"],
      ["A", "2/12/2027"], // 2 dec 2027
      ["B", "7/11/2027"], // 7 nov 2027
      ["C", "1/10/2028"], // 1 oct 2028
      ["D", "10/9/2028"], // 10 sep 2028
    ];
    const excelSerial = (y: number, m: number, d: number) => {
      const base = Date.UTC(1899, 11, 30);
      const msPerDay = 24 * 60 * 60 * 1000;
      return Math.floor((Date.UTC(y, m - 1, d) - base) / msPerDay);
    };
    const minSerial = excelSerial(2027, 11, 1);
    const maxSerial = excelSerial(2027, 12, 31);

    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("BetweenSerial", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "filtre between serial",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: {
              source: { blockRef },
              select: ["Projet", "Date"],
              filter: [{ col: "Date", op: "between", value: [minSerial, maxSerial], type: "date" }],
              dest: { mode: "newSheet" },
            },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B3");
    const ordered = out.values.slice(1).map((r: any[]) => r[0]);
    expect(ordered.length).toBe(2);
    expect(ordered).toEqual(expect.arrayContaining(["A", "B"]));
  });

  test("between janvier-mars 2027 sur dates texte FR ne garde que fevrier", async () => {
    const data = [
      ["Projet", "Date"],
      ["A", "11/07/2027"],
      ["B", "12/02/2027"],
      ["C", "09/10/2028"],
      ["D", "10/01/2028"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("FrBetween", data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "filtre janv-mars 2027",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: {
              source: { blockRef },
              select: ["Projet", "Date"],
              filter: [{ col: "Date", op: "between", value: ["01/01/2027", "31/03/2027"] }],
              dest: { mode: "newSheet" },
            },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B5");
    const projects = out.values.slice(1).map((r: any[]) => r[0]);
    expect(projects).toEqual(["B"]);
  });

  test("runtime: serial-first neutralise text mm/dd vs serial (tri + filtre)", async () => {
    const excelSerial = (y: number, m: number, d: number) => {
      const base = Date.UTC(1899, 11, 30);
      const msPerDay = 24 * 60 * 60 * 1000;
      return Math.floor((Date.UTC(y, m - 1, d) - base) / msPerDay);
    };
    // Excel a mal interprété en US : 11/07/2027 => 7 nov 2027, 12/02/2027 => 2 dec 2027
    const data = [
      ["Projet", "Date"],
      ["A", excelSerial(2027, 11, 7)],
      ["B", excelSerial(2027, 12, 2)],
      ["C", excelSerial(2028, 9, 10)],
      ["D", excelSerial(2028, 10, 1)],
    ];
    const text = [
      ["Projet", "Date"],
      ["A", "11/07/2027"],
      ["B", "12/02/2027"],
      ["C", "09/10/2028"],
      ["D", "10/01/2028"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshotWithFormats("Mismatch", data);
    const ws = ctx.workbook.worksheets.getItem("Mismatch");
    const endCell = "B5";
    ws.getRange(`A1:${endCell}`).text = text;

    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "tri et filtre FR strict",
        
        steps: [
          {
            id: "view1",
            macro: "table_view",
            params: {
              source: { blockRef },
              select: ["Projet", "Date"],
              sort: { col: "Date", dir: "asc" },
              filter: [{ col: "Date", op: "between", value: ["01/01/2027", "31/12/2027"] }],
              dest: { mode: "newSheet" },
            },
          },
        ],
      },
      snapshot
    );
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const art = res.artifacts.find((a) => a.fromStep === "view1" && a.type === "table")!;
    const out = ctx.workbook.worksheets.getItem(art.sheet).getRange(art.address || "A1:B5");
    const ordered = out.values.slice(1).map((r: any[]) => r[0]);
    expect(ordered).toEqual(["A", "B"]); // serial (valeur Excel) gagne en cas de mismatch text/value
  });

});
