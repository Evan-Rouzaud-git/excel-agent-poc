import { parseA1Address } from "../../src/taskpane/agent/utils";

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

class LoadableCollection<T> {
  protected _items: T[];
  private _loaded = false;
  private _pending = false;
  protected ctx?: FakeContext;
  constructor(items: T[], ctx?: FakeContext) {
    this._items = items;
    this.ctx = ctx;
    if (ctx) ctx.registerLoadable(this);
  }
  attachContext(ctx: FakeContext) {
    this.ctx = ctx;
    ctx.registerLoadable(this);
  }
  load() {
    this._pending = true;
    return this;
  }
  markSynced() {
    if (this._pending) {
      this._loaded = true;
      this._pending = false;
    }
    this._items.forEach((i: any) => {
      if (typeof i?.markSynced === "function") i.markSynced();
    });
  }
  get items(): T[] {
    if (!this._loaded) {
      throw new Error("PropertyNotLoaded: items");
    }
    return this._items;
  }
  rawItems() {
    return this._items;
  }
  push(item: T) {
    this._items.push(item);
  }
  getItem(name: string): any {
    const found = (this._items as any[]).find((i) => i.name === name);
    if (!found) throw new Error("KeyNotFound");
    return found;
  }
  add(...args: any[]): any {
    const first = args[0];
    if (first instanceof FakeWorksheet) {
      this._items.push(first as any);
      return first;
    }
    if (typeof first === "string") {
      const ws = new FakeWorksheet(first, this.ctx);
      if (this.ctx) ws.attachContext(this.ctx);
      this._items.push(ws as any);
      return ws as any;
    }
    this._items.push(first as any);
    return first;
  }
  getItemOrNullObject(name: string): any {
    const found = (this._items as any[]).find((i) => i.name === name);
    return found || { isNullObject: true, load: () => ({}) };
  }
}

class FakeName {
  name: string;
  value: string;
  isNullObject = false;
  constructor(name: string, value: string, private ctx?: FakeContext) {
    this.name = name;
    this.value = value;
  }
  load() {
    return this;
  }
  delete() {
    this.isNullObject = true;
  }
}

class FakeNameCollection extends LoadableCollection<FakeName> {
  constructor(ctx?: FakeContext) {
    super([], ctx);
  }
  add(name: string, value: string) {
    const existing = (this._items as FakeName[]).findIndex((item) => item.name === name);
    if (existing >= 0) {
      (this._items as FakeName[]).splice(existing, 1);
    }
    const entry = new FakeName(name, value, this.ctx);
    this._items.push(entry);
    return entry;
  }
  getItem(name: string) {
    return super.getItem(name);
  }
  getItemOrNullObject(name: string) {
    const found = (this._items as FakeName[]).find((item) => item.name === name);
    if (found) return found;
    return { isNullObject: true, load: () => ({}), name };
  }
}

export class FakeRange {
  address: string;
  private ctx?: FakeContext;
  private _values: any[][];
  private _formulas: any[][];
  private _numberFormat: any[][];
  private _text: any[][];
  private parentRange?: FakeRange;
  private parentRowOffset = 0;
  private parentColOffset = 0;
  colCount: number;
  format: any;
  bandedRows?: boolean;
  constructor(
    address: string,
    values: any[][] = [[]],
    ctx?: FakeContext,
    parentRange?: FakeRange,
    parentRowOffset = 0,
    parentColOffset = 0
  ) {
    this.address = address;
    this.ctx = ctx;
    this._values = values;
    this._text = values.map((row) => row.map((v) => (v === null || typeof v === "undefined" ? "" : `${v}`)));
    this._formulas = [];
    this._numberFormat = [];
    this.parentRange = parentRange;
    this.parentRowOffset = parentRowOffset;
    this.parentColOffset = parentColOffset;
    this.colCount = computeWidth(address, values);
    this.format = {
      font: { bold: false, color: undefined },
      fill: {
        color: "",
        clear: () => {
          this.format.fill.color = "";
        },
      },
      columnWidth: undefined,
      autofitColumns: jest.fn(),
    };
  }
  attachContext(ctx: FakeContext) {
    this.ctx = ctx;
  }
  get values() {
    return this._values;
  }
  set values(v: any[][]) {
    this._values = v;
    this._text = v.map((row) => row.map((val) => (val === null || typeof val === "undefined" ? "" : `${val}`)));
    if (this.parentRange) {
      const parentValues = Array.isArray(this.parentRange.values) ? this.parentRange.values.map((row) => [...row]) : [];
      for (let r = 0; r < v.length; r += 1) {
        const parentRowIdx = this.parentRowOffset + r;
        while (parentValues.length <= parentRowIdx) parentValues.push([]);
        const srcRow = (Array.isArray(v[r]) ? v[r] : []) as any[];
        for (let c = 0; c < srcRow.length; c += 1) {
          const parentColIdx = this.parentColOffset + c;
          const targetRow = parentValues[parentRowIdx] ?? (parentValues[parentRowIdx] = []);
          while (targetRow.length <= parentColIdx) targetRow.push("");
          targetRow[parentColIdx] = srcRow[c];
        }
      }
      this.parentRange.values = parentValues;
    }
    if (this.ctx) this.ctx.trackWrite();
    if (this.ctx?.workbook?.worksheets?.rawItems) {
      this.ctx.workbook.worksheets.rawItems().forEach((ws: any) => {
        if (typeof ws?.applyRangeWrite === "function") ws.applyRangeWrite(this.address, v);
      });
    }
  }
  get text() {
    return this._text;
  }
  set text(v: any[][]) {
    this._text = v;
    if (this.ctx) this.ctx.trackWrite();
  }
  get formulas() {
    return this._formulas;
  }
  set formulas(v: any[][]) {
    this._formulas = v;
    if (this.ctx) this.ctx.trackWrite();
  }
  get numberFormat() {
    return this._numberFormat;
  }
  set numberFormat(v: any[][]) {
    this._numberFormat = v;
    if (this.ctx) this.ctx.trackWrite();
  }
  load() {
    return this;
  }

  getColumn(index: number) {
    const addr = this.address.includes("!")
      ? `${this.address.split("!")[0]}!${colToLetter(index)}:${colToLetter(index)}`
      : `${colToLetter(index)}:${colToLetter(index)}`;
    return new FakeRange(addr, [[]], this.ctx);
  }
  delete(direction?: string) {
    const parsed = parseA1Address(this.address);
    if (!parsed || !this.ctx || !parsed.sheet) return;
    const sheets = this.ctx.workbook?.worksheets;
    if (!sheets) return;
    let ws: any;
    try {
      ws = sheets.getItem(parsed.sheet);
    } catch {
      const candidate = sheets.getItemOrNullObject?.(parsed.sheet);
      if (candidate && !candidate.isNullObject) ws = candidate;
    }
    if (!ws || typeof ws.deleteRow !== "function") return;
    ws.deleteRow(parsed.startRow);
  }
}

function computeWidth(address: string, values: any[][]) {
  if (values && values[0] && Array.isArray(values[0]) && values[0].length) return values[0].length;
  const local = address.includes("!") ? address.split("!")[1] || address : address;
  const segments = local.split(/[,;]/).map((seg) => seg.trim()).filter(Boolean);
  if (!segments.length) return 1;
  const toIdx = (token: string | undefined) => {
    if (!token) return 0;
    let i = 0;
    while (i < token.length && (token[i] === "'" || token[i] === "$")) i += 1;
    let letters = "";
    while (i < token.length) {
      const ch = token.charAt(i);
      const upper = ch.toUpperCase();
      const isLetter = upper >= "A" && upper <= "Z";
      if (!isLetter) break;
      letters += upper;
      i += 1;
    }
    if (!letters) return 0;
    let num = 0;
    for (let j = 0; j < letters.length; j += 1) num = num * 26 + (letters.charCodeAt(j) - 64);
    return num - 1;
  };
  let total = 0;
  segments.forEach((segment) => {
    const parsed = parseA1Address(segment);
    if (parsed) {
      total += parsed.endCol - parsed.startCol + 1;
      return;
    }
    const rangeParts = segment.split(":").map((p) => p.trim());
    if (rangeParts.length === 2) {
      const c1 = toIdx(rangeParts[0]);
      const c2 = toIdx(rangeParts[1]);
      total += Math.abs(c2 - c1) + 1 || 1;
      return;
    }
    total += 1;
  });
  return Math.max(1, total);
}

class FakeChartCollection extends LoadableCollection<any> {
  constructor(ctx?: FakeContext) {
    super([], ctx);
  }
  get created() {
    return (this as any)._items as any[];
  }
  add = (type: string, range: FakeRange, plotBy?: string) => {
    const chart = new FakeChart(type, range, (this as any).ctx);
    if (plotBy) (chart as any).plotBy = plotBy;
    this.push(chart as any);
    return chart as any;
  };
}

class FakeSeries {
  valuesRange?: FakeRange;
  xRange?: FakeRange;
  name?: string;
  setValues = jest.fn((range: FakeRange) => {
    this.valuesRange = range;
  });
  setXAxisValues = jest.fn((range: FakeRange) => {
    this.xRange = range;
  });
  setXValues = this.setXAxisValues;
}

class FakeSeriesCollection extends LoadableCollection<FakeSeries> {
  constructor(items: FakeSeries[], ctx?: FakeContext) {
    super(items, ctx);
  }
  add = (range: FakeRange) => {
    const s = new FakeSeries();
    s.setValues(range);
    this.push(s);
    return s;
  };
}

class FakeChart {
  type: string;
  range: FakeRange;
  plotBy?: string;
  series: FakeSeriesCollection;
  private _name = "";
  private _nameAssigned = false;
  private _loaded = false;
  private _pending = false;
  title = { text: "" };
  lastPosition?: { topLeft: string; bottomRight: string; sheet: string | undefined };
  setPosition = jest.fn((topLeft: FakeRange | string, bottomRight: FakeRange | string) => {
    const tlAddr = typeof topLeft === "string" ? topLeft : topLeft.address;
    const brAddr = typeof bottomRight === "string" ? bottomRight : bottomRight.address;
    const sheet = tlAddr.includes("!") ? tlAddr.split("!")[0] : undefined;
    this.lastPosition = { topLeft: tlAddr, bottomRight: brAddr, sheet };
  });
  constructor(type: string, range: FakeRange, ctx?: FakeContext) {
    this.type = type;
    this.range = range;
    const seriesCount = Math.max(1, range.colCount || 1);
    const initial = Array.from({ length: seriesCount }, () => new FakeSeries());
    this.series = new FakeSeriesCollection(initial, ctx);
  }
  load(arg?: string | string[]) {
    if (arg === "name" || (Array.isArray(arg) && arg.includes("name"))) {
      this._pending = true;
    }
    return this;
  }
  markSynced() {
    if (this._pending) {
      this._loaded = true;
      this._pending = false;
    }
  }
  get name() {
    if (!this._loaded && !this._nameAssigned) {
      throw new Error("PropertyNotLoaded: name");
    }
    return this._name;
  }
  set name(v: string) {
    this._name = v;
    this._nameAssigned = true;
  }
}

class FakeTable {
  name: string;
  address: string;
  showBandedRows = false;
  showHeaderRow = true;
  style: string | undefined;
  private ctx?: FakeContext;
  private sheet?: FakeWorksheet;
  constructor(name: string, address: string, ctx?: FakeContext, sheet?: FakeWorksheet) {
    this.name = name;
    this.address = address;
    this.ctx = ctx;
    this.sheet = sheet;
  }
  private getSourceValues() {
    if (!this.sheet) return [];
    const source = this.sheet.getRange(this.address);
    return Array.isArray(source.values) ? source.values : [];
  }
  private buildRange(localAddress: string, values: any[][]) {
    const sheetPrefix = this.sheet ? `${this.sheet.name}!` : "";
    return new FakeRange(`${sheetPrefix}${localAddress}`, values, this.ctx);
  }
  getHeaderRowRange() {
    const values = this.getSourceValues();
    const header = values.length ? [Array.isArray(values[0]) ? values[0] : []] : [[]];
    const parsed = parseA1Address(this.address);
    const localAddress = parsed ? `${colToLetter(parsed.startCol)}${parsed.startRow + 1}:${colToLetter(parsed.endCol)}${parsed.startRow + 1}` : this.address;
    return this.buildRange(localAddress, header);
  }
  getDataBodyRange() {
    const values = this.getSourceValues();
    const body = values.length > 1 ? values.slice(1) : [];
    const parsed = parseA1Address(this.address);
    const localAddress = parsed
      ? `${colToLetter(parsed.startCol)}${parsed.startRow + 2}:${colToLetter(parsed.endCol)}${parsed.endRow + 1}`
      : this.address;
    return this.buildRange(localAddress, body);
  }
  getRange() {
    const values = this.getSourceValues();
    return this.buildRange(this.address, values);
  }
  load() {
    return this;
  }
  markSynced() {}
}

class FakeTableCollection extends LoadableCollection<FakeTable> {
  addCalls: Array<{ address: string | FakeRange; hasHeaders: boolean }> = [];
  private owner?: FakeWorksheet;
  constructor(ctx?: FakeContext, owner?: FakeWorksheet) {
    super([], ctx);
    this.owner = owner;
  }
  add(address: string | FakeRange, hasHeaders: boolean) {
    const name = `Table${this._items.length + 1}`;
    const addr = typeof address === "string" ? address : (address as any)?.address || "";
    const tbl = new FakeTable(name, addr, this.ctx, this.owner);
    tbl.showHeaderRow = hasHeaders;
    this.push(tbl as any);
    this.addCalls.push({ address, hasHeaders });
    return tbl;
  }
  getItem(name: string) {
    return super.getItem(name) as any;
  }
  getItemOrNullObject(name: string) {
    const found = (this._items as any[]).find((t) => t.name === name);
    return found || { isNullObject: true, load: () => ({}) };
  }
}

export class FakeWorksheet {
  private _name: string;
  private _nameLoaded = true;
  private _pendingName = false;
  ranges: Record<string, FakeRange> = {};
  charts: FakeChartCollection;
  tables: FakeTableCollection;
  freezePanes: { freezeRows: jest.Mock };
  private ctx?: FakeContext;
  constructor(name: string, ctx?: FakeContext) {
    this._name = name;
    this.ctx = ctx;
    this.charts = new FakeChartCollection(ctx);
    this.tables = new FakeTableCollection(ctx, this);
    this.freezePanes = { freezeRows: jest.fn() };
  }
  load(arg?: string | string[]) {
    if (arg === "name" || (Array.isArray(arg) && arg.includes("name"))) {
      this._pendingName = true;
    }
    return this;
  }
  markSynced() {
    if (this._pendingName) {
      this._nameLoaded = true;
      this._pendingName = false;
    }
    this.charts.markSynced();
    this.tables.markSynced();
  }
  get name() {
    return this._name;
  }
  set name(v: string) {
    this._name = v;
    this._nameLoaded = true;
  }
  attachContext(ctx: FakeContext) {
    this.ctx = ctx;
    this.charts.attachContext(ctx);
    this.tables.attachContext(ctx);
    Object.values(this.ranges).forEach((r) => r.attachContext(ctx));
  }
  private findCoveringRange(address: string): FakeRange | null {
    const requested = parseA1Address(address);
    if (!requested) return null;
    let best: FakeRange | null = null;
    let bestArea = -1;
    Object.values(this.ranges).forEach((range) => {
      const parsed = parseA1Address(range.address);
      if (!parsed) return;
      if (requested.sheet && parsed.sheet && requested.sheet !== parsed.sheet) return;
      if (
        requested.startRow < parsed.startRow ||
        requested.endRow > parsed.endRow ||
        requested.startCol < parsed.startCol ||
        requested.endCol > parsed.endCol
      ) {
        return;
      }
      const area = (parsed.endRow - parsed.startRow + 1) * (parsed.endCol - parsed.startCol + 1);
      if (area > bestArea) {
        bestArea = area;
        best = range;
      }
    });
    return best;
  }
  applyRangeWrite(address: string, values: any[][]) {
    const written = parseA1Address(address);
    if (!written) return;
    Object.values(this.ranges).forEach((range) => {
      const parsed = parseA1Address(range.address);
      if (!parsed) return;
      if (written.sheet && parsed.sheet && written.sheet !== parsed.sheet) return;
      const intersects =
        !(written.endRow < parsed.startRow || written.startRow > parsed.endRow || written.endCol < parsed.startCol || written.startCol > parsed.endCol);
      if (!intersects) return;
      const current = Array.isArray((range as any).values) ? (range as any).values.map((row: any[]) => [...row]) : [];
      const targetRowCount = parsed.endRow - parsed.startRow + 1;
      const targetColCount = parsed.endCol - parsed.startCol + 1;
      while (current.length < targetRowCount) current.push([]);
      for (let r = 0; r < targetRowCount; r += 1) {
        const row = current[r] || [];
        while (row.length < targetColCount) row.push("");
        current[r] = row;
      }
      const rowStart = Math.max(parsed.startRow, written.startRow);
      const rowEnd = Math.min(parsed.endRow, written.endRow);
      const colStart = Math.max(parsed.startCol, written.startCol);
      const colEnd = Math.min(parsed.endCol, written.endCol);
      for (let r = rowStart; r <= rowEnd; r += 1) {
        for (let c = colStart; c <= colEnd; c += 1) {
          const sourceRow = values[r - written.startRow] || [];
          const sourceValue = sourceRow[c - written.startCol];
          const targetRowIdx = r - parsed.startRow;
          const targetColIdx = c - parsed.startCol;
          current[targetRowIdx] = current[targetRowIdx] || [];
          current[targetRowIdx][targetColIdx] = sourceValue;
        }
      }
      (range as any)._values = current;
      (range as any)._text = current.map((row: any[]) => row.map((v) => (v === null || typeof v === "undefined" ? "" : `${v}`)));
    });
  }
  deleteRow(rowIndex: number) {
    Object.values(this.ranges).forEach((range) => {
      const parsed = parseA1Address(range.address);
      if (!parsed || typeof parsed.startRow !== "number" || typeof parsed.endRow !== "number") return;
      if (rowIndex < parsed.startRow || rowIndex > parsed.endRow) return;
      const relative = rowIndex - parsed.startRow;
      const currentValues = Array.isArray((range as any).values) ? (range as any).values : [];
      if (relative < 0 || relative >= currentValues.length) return;
      const updated = currentValues.slice();
      updated.splice(relative, 1);
      (range as any).values = updated;
    });
  }
  getRange(address: string) {
    if (!this.ranges[address]) {
      const covering: FakeRange | null = this.findCoveringRange(address);
      const requested = parseA1Address(address);
      const source = covering ? parseA1Address(covering.address) : null;
      if (covering && requested && source) {
        const rowOffset = requested.startRow - source.startRow;
        const colOffset = requested.startCol - source.startCol;
        const slice = (covering.values as any[][])
          .slice(rowOffset, rowOffset + (requested.endRow - requested.startRow + 1))
          .map((row: any[]) => row.slice(colOffset, colOffset + (requested.endCol - requested.startCol + 1)));
        this.ranges[address] = new FakeRange(`${this._name}!${address}`, slice, this.ctx, covering, rowOffset, colOffset);
      } else {
        this.ranges[address] = new FakeRange(`${this._name}!${address}`, [[]], this.ctx);
      }
    }
    return this.ranges[address];
  }
  getRangeByIndexes(row: number, col: number, rows: number, cols: number) {
    const addr = `${colToLetter(col)}${row + 1}:${colToLetter(col + cols - 1)}${row + rows}`;
    return this.getRange(addr);
  }
}

export class FakeWorkbook {
  worksheets: LoadableCollection<FakeWorksheet>;
  names: FakeNameCollection;
  constructor(sheets: FakeWorksheet[], ctx?: FakeContext) {
    this.worksheets = new LoadableCollection<FakeWorksheet>(sheets, ctx);
    this.names = new FakeNameCollection(ctx);
  }
  attachContext(ctx: FakeContext) {
    this.worksheets.attachContext(ctx);
    this.worksheets.rawItems().forEach((ws) => ws.attachContext(ctx));
    this.names.attachContext(ctx);
  }
}

export class FakeContext {
  workbook: FakeWorkbook;
  sync: jest.Mock;
  private loadables: LoadableCollection<any>[] = [];
  private writeCount = 0;
  constructor(wb: FakeWorkbook) {
    this.workbook = wb;
    wb.attachContext(this);
    this.sync = jest.fn(() => {
      this.loadables.forEach((l) => l.markSynced());
      return Promise.resolve();
    });
  }
  registerLoadable(l: LoadableCollection<any>) {
    this.loadables.push(l);
  }
  trackWrite() {
    this.writeCount += 1;
  }
  resetWrites() {
    this.writeCount = 0;
  }
  getWriteCount() {
    return this.writeCount;
  }
}
