import { getWorkbookContext } from "../src/taskpane/context/getWorkbookContext";

// ---------- Mock helpers ----------
type RangeInit = {
  address: string;
  rowCount: number;
  columnCount: number;
  rowIndex: number;
  columnIndex: number;
  values: any[][];
  isNullObject?: boolean;
};

class MockRange {
  address: string;
  rowCount: number;
  columnCount: number;
  rowIndex: number;
  columnIndex: number;
  values: any[][];
  isNullObject: boolean;
  constructor(init: RangeInit) {
    this.address = init.address;
    this.rowCount = init.rowCount;
    this.columnCount = init.columnCount;
    this.rowIndex = init.rowIndex;
    this.columnIndex = init.columnIndex;
    this.values = init.values;
    this.isNullObject = init.isNullObject ?? false;
  }
  load(_: string | string[]) {
    return this;
  }
}

class MockTable {
  name: string;
  private tableRange: MockRange;
  private headerRange: MockRange;
  private bodyRange: MockRange;
  constructor(name: string, address: string, headerValues: any[][], bodyValues: any[][] = []) {
    this.name = name;
    const headerRow = headerValues[0] ?? [];
    const fullValues = [headerRow, ...bodyValues];
    this.tableRange = new MockRange({
      address,
      rowCount: fullValues.length,
      columnCount: fullValues[0]?.length ?? 1,
      rowIndex: 0,
      columnIndex: 0,
      values: fullValues,
    });
    this.headerRange = new MockRange({
      address,
      rowCount: 1,
      columnCount: headerRow.length || 1,
      rowIndex: 0,
      columnIndex: 0,
      values: [headerRow],
    });
    this.bodyRange = new MockRange({
      address,
      rowCount: fullValues.length,
      columnCount: fullValues[0]?.length ?? 1,
      rowIndex: 0,
      columnIndex: 0,
      values: fullValues,
    });
  }
  load(_: string | string[]) {
    return this;
  }
  getRange() {
    return this.tableRange;
  }
  getHeaderRowRange() {
    return this.headerRange;
  }
  getDataBodyRange() {
    return this.bodyRange;
  }
}

class MockChart {
  name: string;
  private rawType: string;
  private ready = false;
  constructor(name: string, chartType: string, private pendingList: MockChart[]) {
    this.name = name;
    this.rawType = chartType;
  }
  load(_: string | string[]) {
    this.pendingList.push(this);
    return this;
  }
  markReady() {
    this.ready = true;
  }
  get chartType(): string {
    if (!this.ready) {
      throw new Error("chartType accessed before sync");
    }
    return this.rawType;
  }
}

class MockWorksheet {
  name: string;
  private usedPrimary: MockRange;
  private usedFallback: MockRange;
  tables: { items: MockTable[]; load: (s: string) => void };
  charts: { items: MockChart[]; load: (s: string) => void };

  constructor(
    name: string,
    usedPrimary: MockRange,
    usedFallback?: MockRange,
    tables: MockTable[] = [],
    charts: MockChart[] = []
  ) {
    this.name = name;
    this.usedPrimary = usedPrimary;
    this.usedFallback = usedFallback ?? usedPrimary;
    this.tables = { items: tables, load: () => {} };
    this.charts = { items: charts, load: () => {} };
  }

  load(_: string | string[]) {}

  getUsedRangeOrNullObject() {
    return this.usedPrimary;
  }

  getUsedRange() {
    return this.usedFallback;
  }

  getRangeByIndexes(row: number, col: number, rows: number, cols: number) {
    const vals = this.usedFallback.values.slice(row - this.usedFallback.rowIndex, row - this.usedFallback.rowIndex + rows).map((r) => r.slice(col - this.usedFallback.columnIndex, col - this.usedFallback.columnIndex + cols));
    const addr = `${indexToLetter(col)}${row + 1}:${indexToLetter(col + cols - 1)}${row + rows}`;
    return new MockRange({
      address: addr,
      rowCount: rows,
      columnCount: cols,
      rowIndex: row,
      columnIndex: col,
      values: vals,
    });
  }
}

class MockWorksheets {
  items: MockWorksheet[];
  constructor(items: MockWorksheet[]) {
    this.items = items;
  }
  load(_: string | string[]) {}
  getActiveWorksheet() {
    return this.items[0];
  }
}

class MockWorkbook {
  worksheets: MockWorksheets;
  name = "Book1";
  readOnly = false;
  selectionAddress: string;
  constructor(ws: MockWorksheet[], selectionAddress = "A1") {
    this.worksheets = new MockWorksheets(ws);
    this.selectionAddress = selectionAddress;
  }
  load(_: string | string[]) {}
  getSelectedRange() {
    if (this.selectionAddress === "THROW") throw new Error("not a range");
    return new MockRange({
      address: this.selectionAddress,
      rowCount: 1,
      columnCount: 1,
      rowIndex: 0,
      columnIndex: 0,
      values: [[1]],
    });
  }
}

class MockContext {
  workbook: MockWorkbook;
  syncCalls = 0;
  private pendingCharts: MockChart[];
  constructor(wb: MockWorkbook, pendingCharts: MockChart[]) {
    this.workbook = wb;
    this.pendingCharts = pendingCharts;
  }
  sync() {
    this.pendingCharts.forEach((c) => c.markReady());
    this.pendingCharts.length = 0;
    this.syncCalls += 1;
    return Promise.resolve();
  }
}

function indexToLetter(idx: number): string {
  let n = idx + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ---------- Tests ----------

describe("getWorkbookContext minimal tables/charts V1.2", () => {
  const origExcel = (global as any).Excel;
  afterEach(() => {
    (global as any).Excel = origExcel;
  });

  test("selection unavailable sets limitation and null address", async () => {
    const used = new MockRange({
      address: "Sheet1!A1:A1",
      rowCount: 1,
      columnCount: 1,
      rowIndex: 0,
      columnIndex: 0,
      values: [[1]],
    });
    const ws = new MockWorksheet("Sheet1", used);
    const wb = new MockWorkbook([ws], "THROW");
    const ctx = new MockContext(wb, []);
    (global as any).Excel = { run: (cb: any) => cb(ctx) };
    const snap = await getWorkbookContext();
    expect(snap.active.selectionAddress).toBeNull();
    expect(snap.limitations.join(" ")).toContain("active.selection:not a range");
  });

  test("usedRange fallback uses getUsedRange when orNullObject is empty", async () => {
    const empty = new MockRange({
      address: "Sheet1!A1",
      rowCount: 0,
      columnCount: 0,
      rowIndex: 0,
      columnIndex: 0,
      values: [],
      isNullObject: true,
    });
    const fallback = new MockRange({
      address: "Sheet1!A1:B2",
      rowCount: 2,
      columnCount: 2,
      rowIndex: 0,
      columnIndex: 0,
      values: [
        [1, 2],
        [3, 4],
      ],
    });
    const ws = new MockWorksheet("Sheet1", empty, fallback);
    const wb = new MockWorkbook([ws]);
    const ctx = new MockContext(wb, []);
    (global as any).Excel = { run: (cb: any) => cb(ctx) };
    const snap = await getWorkbookContext();
    expect(snap.sheets[0]?.usedRange).toBe("A1:B2");
    expect(snap.sheets[0]?.valueBounds.address).toBe("A1:B2");
    expect(snap.sheets[0]?.limitations.join(" ")).not.toContain("usedRange failed");
  });

  test("tables create blocks with normalized id and address", async () => {
    const table = new MockTable("T1", "Sheet1!B2:D4", [["H1", "H2", "H3"]], [
      ["a", 1, 2],
      ["b", 3, 4],
    ]);
    const used = new MockRange({
      address: "Sheet1!B2:D4",
      rowCount: 3,
      columnCount: 3,
      rowIndex: 1,
      columnIndex: 1,
      values: [["H1", "H2", "H3"], ["a", 1, 2], ["b", 3, 4]],
    });
    const ws = new MockWorksheet("Sheet1", used, undefined, [table]);
    const wb = new MockWorkbook([ws]);
    const ctx = new MockContext(wb, []);
    (global as any).Excel = { run: (cb: any) => cb(ctx) };
    const snap = await getWorkbookContext();
    const sheet = snap.sheets[0]!;
    expect(sheet.blocks.length).toBe(1);
    const block = sheet.blocks[0]!;
    expect(block.address).toBe("B2:D4");
    expect(block.id).toBe("Sheet1!B2:D4");
    expect(block.source.type).toBe("table");
    expect(block.headers).toEqual(["H1", "H2", "H3"]);
  });

  test("promotes usedRange to range block when headers present", async () => {
    const used = new MockRange({
      address: "Sheet1!E6:F10",
      rowCount: 5,
      columnCount: 2,
      rowIndex: 5,
      columnIndex: 4,
      values: [
        ["Revenus", "Depenses"],
        [100, 40],
        [120, 50],
        [90, 30],
        [200, 120],
      ],
    });
    const ws = new MockWorksheet("Sheet1", used);
    const wb = new MockWorkbook([ws], "Sheet1!G13");
    const ctx = new MockContext(wb, []);
    (global as any).Excel = { run: (cb: any) => cb(ctx) };

    const snap = await getWorkbookContext();
    const sheet = snap.sheets[0]!;
    expect(sheet.blocks.length).toBe(1);
    const block = sheet.blocks[0]!;
    expect(block.kind).toBe("range");
    expect(block.id).toBe("Sheet1!E6:F10");
    expect(block.headers).toEqual(["Revenus", "Depenses"]);
    expect(sheet.counts.tables).toBe(0);
    expect(snap.totals.blocks).toBe(1);
  });

  test("charts detected only after load+sync", async () => {
    const pending: MockChart[] = [];
    const chart = new MockChart("Chart1", "ColumnClustered", pending);
    const used = new MockRange({
      address: "Charts!A1",
      rowCount: 1,
      columnCount: 1,
      rowIndex: 0,
      columnIndex: 0,
      values: [[1]],
    });
    const ws = new MockWorksheet("Charts", used, undefined, [], [chart]);
    const wb = new MockWorkbook([ws]);
    const ctx = new MockContext(wb, pending);
    (global as any).Excel = { run: (cb: any) => cb(ctx) };

    const snap = await getWorkbookContext();
    expect(snap.totals.charts).toBe(1);
    const sheet = snap.sheets[0]!;
    expect(sheet.charts[0]?.chartType).toBe("ColumnClustered");
    expect(pending.length).toBe(0); // all loads flushed via sync
  });
});
