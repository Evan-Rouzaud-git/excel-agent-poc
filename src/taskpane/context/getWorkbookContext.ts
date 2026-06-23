/* global Excel */
import {
  BlockSnapshot,
  BoundsSnapshot,
  ChartSnapshot,
  ContextLogger,
  ContextOptions,
  DEFAULT_CONTEXT_OPTIONS,
  SheetSnapshot,
  TableSnapshot,
  WorkbookContextSnapshot,
} from "./types";

type RangeBounds = { startRow: number; startCol: number; endRow: number; endCol: number };

const capabilities = [
  "usedRange.capped",
  "blocks.from.tables",
  "tables.headers",
  "tables.address",
  "charts.detect.simple",
];

const globalLimitations = ["conditionalFormats:not implemented", "validations:not implemented", "overlay:not implemented in V1.2"];

const noopLogger: ContextLogger = { info: () => {}, warn: () => {}, error: () => {} };

export async function getWorkbookContext(options: ContextOptions = {}): Promise<WorkbookContextSnapshot> {
  const opts: Required<ContextOptions> = { ...DEFAULT_CONTEXT_OPTIONS, ...options };
  const log: ContextLogger = (options as any)?.logger ?? noopLogger;
  const started = Date.now();

  return Excel.run<WorkbookContextSnapshot>(async (context) => {
    // Selection (safe)
    let selectionAddress: string | null = null;
    const selectionLimitations: string[] = [];
    try {
      const selection = context.workbook.getSelectedRange();
      selection.load("address");
      await context.sync();
      selectionAddress = typeof (selection as any).address === "string" ? (selection as any).address : null;
    } catch (err) {
      selectionAddress = null;
      selectionLimitations.push("active.selection:not a range");
      log.warn("selection", "selection unavailable", undefined, err);
    }

    // Workbook meta
    let workbookName: string | null = null;
    let readOnly: boolean | null = null;
    try {
      context.workbook.load(["name", "readOnly"]);
      await context.sync();
      workbookName = (context.workbook as any).name ?? null;
      readOnly = (context.workbook as any).readOnly ?? null;
    } catch (err) {
      log.warn("workbook", "metadata unavailable", undefined, err);
    }

    // Sheets
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const activeSheet = sheets.getActiveWorksheet();
    activeSheet.load("name");
    await context.sync();

    const sheetItems = sheets.items.slice(0, opts.maxSheets).sort((a, b) => a.name.localeCompare(b.name));
    const sheetSnapshots: SheetSnapshot[] = [];

    let totalTables = 0;
    let totalBlocks = 0;
    let totalCharts = 0;
    let selectionInBlockId: string | null = null;
    let nearestBlockId: string | null = null;

    const selectionInfo = selectionAddress ? parseAddressWithSheet(selectionAddress) : null;
    const selectionSheet = selectionInfo?.sheet ?? (activeSheet as any)?.name ?? null;

    for (const sheet of sheetItems) {
      const sheetLimitations: string[] = [];

      // Charts (simple)
      const charts = opts.includeCharts ? await detectChartsSimple(context, sheet, sheetLimitations) : [];

      // Used range with fallback
      const usedResult = await getUsedRangeSafe(context, sheet, sheetLimitations);
      if (!usedResult.range) {
        sheetSnapshots.push({
          name: sheet.name,
          usedRange: null,
          valueBounds: emptyBounds(),
          counts: { tables: 0, charts: charts.length },
          tables: [],
          blocks: [],
          charts,
          limitations: sheetLimitations,
        });
        totalCharts += charts.length;
        continue;
      }

      const usedRange = usedResult.range;
      const rowCount = typeof usedRange.rowCount === "number" ? usedRange.rowCount : 0;
      const colCount = typeof usedRange.columnCount === "number" ? usedRange.columnCount : 0;
      const baseRow = typeof usedRange.rowIndex === "number" ? usedRange.rowIndex : 0;
      const baseCol = typeof usedRange.columnIndex === "number" ? usedRange.columnIndex : 0;

      const cappedRows = Math.max(0, Math.min(opts.maxUsedRows, rowCount));
      const cappedCols = Math.max(0, Math.min(opts.maxUsedCols, colCount));

      const readRange =
        cappedRows > 0 && cappedCols > 0
          ? sheet.getRangeByIndexes(baseRow, baseCol, cappedRows, cappedCols)
          : sheet.getRangeByIndexes(baseRow, baseCol, 1, 1);
      readRange.load(["values", "address", "rowCount", "columnCount"]);
      await context.sync();

      const values = (readRange as any).values ?? [];
      const valueBounds = findValueBounds(values, baseRow, baseCol);
      const usedRangeAddress = typeof (usedRange as any).address === "string" ? stripSheetPrefix((usedRange as any).address, sheet.name) : null;

      // Tables -> blocks
      const tableCollection = sheet.tables;
      tableCollection.load("items");
      await context.sync();

      const tableItems = tableCollection.items.slice().sort((a, b) => a.name.localeCompare(b.name));
      const tables: TableSnapshot[] = [];
      const blocks: BlockSnapshot[] = [];

      const tableRanges: Excel.Range[] = [];
      const headerRanges: Excel.Range[] = [];
      const bodyRanges: Excel.Range[] = [];

      tableItems.forEach((t) => {
        t.load("name");
        const tableRange = t.getRange();
        const headerRange = t.getHeaderRowRange();
        const bodyRange = t.getDataBodyRange();
        tableRanges.push(tableRange);
        headerRanges.push(headerRange);
        bodyRanges.push(bodyRange);
        tableRange.load(["address", "values", "rowCount", "columnCount"]);
        headerRange.load(["address", "values"]);
        bodyRange.load(["address"]);
      });
      await context.sync();

      tableItems.forEach((t, idx) => {
        const tableRange = tableRanges[idx];
        const headerRange = headerRanges[idx];
        const bodyRange = bodyRanges[idx];

        const tableAddress = typeof (tableRange as any).address === "string" ? (tableRange as any).address : bodyRange?.address ?? "";
        const headerAddress = typeof (headerRange as any).address === "string" ? (headerRange as any).address : null;
        const dataBodyAddress = typeof (bodyRange as any).address === "string" ? (bodyRange as any).address : null;
        const headerValues = Array.isArray((headerRange as any).values) ? (headerRange as any).values[0] ?? [] : [];
        const previewValuesRaw = Array.isArray((tableRange as any).values) ? (tableRange as any).values : [];
        const preview = previewValuesRaw.slice(0, opts.maxPreviewRows).map((row: any[]) => row.map(serializeCell));

        const headers = headerValues.map((v: any) => (v === null || v === undefined ? "" : String(v)));
        const columnTypes = inferColumnTypes(preview);

        const localAddress = stripSheetPrefix(tableAddress, sheet.name);
        const block: BlockSnapshot = {
          id: `${sheet.name}!${localAddress}`,
          address: localAddress,
          kind: "table",
          confidence: 1,
          headerRowIndex: headers.length > 0 ? 0 : null,
          headers,
          columnTypes,
          preview,
          source: { type: "table", tableName: t.name, tableAddress: tableAddress ?? localAddress },
          duplicateOf: null,
        };

        tables.push({
          name: t.name,
          address: tableAddress,
          headerAddress,
          dataBodyAddress,
          headers,
        });
        blocks.push(block);
      });

      // Promote usedRange to a synthetic range block when no tables are detected
      if (blocks.length === 0) {
        const rangeBlock = promoteRangeBlockFromUsedRange({
          values,
          valueBounds,
          baseRow,
          baseCol,
          usedRangeAddress,
          sheetName: sheet.name,
          maxPreviewRows: opts.maxPreviewRows,
        });
        if (rangeBlock) blocks.push(rangeBlock);
      }

      blocks.sort((a, b) => a.id.localeCompare(b.id));

      // Selection-in-block detection
      if (!selectionInBlockId && selectionAddress && selectionSheet === sheet.name && selectionInfo?.bounds) {
        for (const block of blocks) {
          const blockBounds = parseAddressBounds(block.address);
          if (blockBounds && containsRange(blockBounds, selectionInfo.bounds)) {
            selectionInBlockId = block.id;
            break;
          }
        }
      }

      // Nearest block fallback when no selection
      if (!selectionAddress && sheet.name === (activeSheet as any).name && !nearestBlockId && blocks.length > 0) {
        const firstBlock = blocks[0];
        if (firstBlock) nearestBlockId = firstBlock.id;
      }

      sheetSnapshots.push({
        name: sheet.name,
        usedRange: usedRangeAddress,
        valueBounds,
        counts: { tables: tables.length, charts: charts.length },
        tables,
        blocks,
        charts,
        limitations: sheetLimitations,
      });

      totalTables += tables.length;
      totalBlocks += blocks.length;
      totalCharts += charts.length;
    }

    const durationMs = Date.now() - started;
    log.info("workbook", `workbook processed in ${durationMs}ms`);

    return {
      workbook: { name: workbookName, readOnly },
      active: { sheetName: (activeSheet as any)?.name ?? null, selectionAddress, selectionInBlockId, nearestBlockId },
      capabilities,
      limitations: [...globalLimitations, ...selectionLimitations],
      sheets: sheetSnapshots,
      totals: {
        sheets: sheetSnapshots.length,
        tables: totalTables,
        charts: totalCharts,
        blocks: totalBlocks,
        durationMs,
      },
    };
  });
}

// ----------------- helpers -----------------
function emptyBounds(): BoundsSnapshot {
  return { firstRow: null, firstCol: null, lastRow: null, lastCol: null, address: null };
}

function serializeCell(v: any): any {
  if (v instanceof Date) return v.toISOString();
  return v;
}

function stripSheetPrefix(address: string | null, sheet: string): string {
  if (!address) return "";
  const candidates = [`${sheet}!`, `'${sheet}'!`];
  for (const candidate of candidates) {
    if (address.toLowerCase().startsWith(candidate.toLowerCase())) {
      return address.slice(candidate.length);
    }
  }
  return address;
}

function parseAddressBounds(address: string | null): RangeBounds | null {
  if (!address) return null;
  const trimmed = address.trim();
  const bang = trimmed.lastIndexOf("!");
  const rangePart = bang >= 0 ? trimmed.slice(bang + 1) : trimmed;

  const parseCell = (cell: string): { row: number; col: number } | null => {
    let idx = 0;
    if (cell[idx] === "$") idx += 1;
    let letters = "";
    while (idx < cell.length) {
      const ch = cell[idx] || "";
      const code = ch.charCodeAt(0);
      const isLetter = (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
      if (!isLetter) break;
      letters += ch;
      idx += 1;
    }
    if (!letters) return null;
    if (cell[idx] === "$") idx += 1;
    let digits = "";
    while (idx < cell.length) {
      const ch = cell[idx] || "";
      const isDigit = ch >= "0" && ch <= "9";
      if (!isDigit) return null;
      digits += ch;
      idx += 1;
    }
    if (!digits) return null;
    return { row: Number.parseInt(digits, 10) - 1, col: letterToIndex(letters) };
  };

  const parts = rangePart.split(":");
  const first = parseCell(parts[0] || "");
  const second = parts.length > 1 ? parseCell(parts[1] || "") : null;
  if (!first) return null;
  return {
    startRow: first.row,
    endRow: second ? second.row : first.row,
    startCol: first.col,
    endCol: second ? second.col : first.col,
  };
}

function parseAddressWithSheet(address: string): { sheet: string | null; bounds: RangeBounds | null } | null {
  const trimmed = address.trim();
  const bang = trimmed.lastIndexOf("!");
  let sheet: string | null = null;
  let rangePart = trimmed;
  if (bang >= 0) {
    sheet = trimmed.slice(0, bang);
    rangePart = trimmed.slice(bang + 1);
    if (sheet.startsWith("'") && sheet.endsWith("'") && sheet.length >= 2) {
      sheet = sheet.slice(1, -1);
    }
  }
  const bounds = parseAddressBounds(rangePart);
  return { sheet, bounds };
}

function letterToIndex(letters: string): number {
  return letters
    .toUpperCase()
    .split("")
    .reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
}

function indexToLetter(idx: number): string {
  let n = idx + 1;
  let res = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    res = String.fromCharCode(65 + mod) + res;
    n = Math.floor((n - 1) / 26);
  }
  return res;
}

function makeAddress(startRow: number, startCol: number, endRow: number, endCol: number): string {
  return `${indexToLetter(startCol)}${startRow + 1}:${indexToLetter(endCol)}${endRow + 1}`;
}

function findValueBounds(values: any[][], baseRow: number, baseCol: number): BoundsSnapshot {
  let minR = Number.POSITIVE_INFINITY;
  let minC = Number.POSITIVE_INFINITY;
  let maxR = -1;
  let maxC = -1;
  for (let r = 0; r < values.length; r++) {
    const row = values[r] ?? [];
    for (let c = 0; c < row.length; c++) {
      if (row[c] !== null && row[c] !== undefined && row[c] !== "") {
        minR = Math.min(minR, r);
        minC = Math.min(minC, c);
        maxR = Math.max(maxR, r);
        maxC = Math.max(maxC, c);
      }
    }
  }

  if (maxR === -1 || maxC === -1) return emptyBounds();

  return {
    firstRow: baseRow + minR,
    firstCol: baseCol + minC,
    lastRow: baseRow + maxR,
    lastCol: baseCol + maxC,
    address: makeAddress(baseRow + minR, baseCol + minC, baseRow + maxR, baseCol + maxC),
  };
}

async function getUsedRangeSafe(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  sheetLimitations: string[]
): Promise<{ range: Excel.Range | null }> {
  const metaFields = ["isNullObject", "address", "rowIndex", "columnIndex", "rowCount", "columnCount"];

  const isRangeReady = (range: any): boolean =>
    typeof range?.rowIndex === "number" &&
    typeof range?.columnIndex === "number" &&
    typeof range?.rowCount === "number" &&
    typeof range?.columnCount === "number";

  const tryRange = async (getter: () => Excel.Range, label: string): Promise<Excel.Range | null> => {
    try {
      const range = getter();
      range.load(metaFields as any);
      await context.sync();
      return range;
    } catch (err) {
      sheetLimitations.push(`sheet:${sheet.name}:usedRange failed:${label}:${(err as Error)?.message ?? "unknown"}`);
      return null;
    }
  };

  let range = await tryRange(() => sheet.getUsedRangeOrNullObject(false), "getUsedRangeOrNullObject");
  if (range) {
    const isNull = (range as any).isNullObject === true;
    if (isNull || !isRangeReady(range)) {
      range = await tryRange(() => sheet.getUsedRange(), "getUsedRange");
    }
  } else {
    range = await tryRange(() => sheet.getUsedRange(), "getUsedRange");
  }

  if (!range || !isRangeReady(range)) {
    sheetLimitations.push(`sheet:${sheet.name}:usedRange failed:unavailable`);
    return { range: null };
  }

  return { range };
}

async function detectChartsSimple(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  sheetLimitations: string[]
): Promise<ChartSnapshot[]> {
  const charts: ChartSnapshot[] = [];
  try {
    const chartCollection = sheet.charts;
    chartCollection.load("items");
    await context.sync();
    const items = chartCollection.items ?? [];
    items.forEach((chart) => chart.load(["name", "chartType"] as any));
    await context.sync();
    items
      .slice()
      .sort((a, b) => (a.name || "").localeCompare(b.name || ""))
      .forEach((chart) => {
        charts.push({
          name: chart.name ?? "",
          chartType: (chart as any).chartType ?? "",
        });
      });
  } catch (err) {
    sheetLimitations.push(`charts:unavailable:${(err as Error)?.message ?? "unknown"}`);
    return [];
  }
  return charts;
}

function inferColumnTypes(preview: any[][]): Array<"number" | "date" | "text" | "mixed"> {
  const rows = preview.slice(1); // skip header
  const colCount = rows.reduce((m, r) => Math.max(m, r.length), preview[0]?.length ?? 0);
  const columnTypes: Array<"number" | "date" | "text" | "mixed"> = [];

  for (let c = 0; c < colCount; c++) {
    const seen = new Set<"number" | "date" | "text">();
    for (const row of rows) {
      const val = row?.[c];
      if (val === null || val === undefined || val === "") continue;
      if (typeof val === "number") {
        seen.add("number");
      } else if (val instanceof Date) {
        seen.add("date");
      } else {
        seen.add("text");
      }
    }
    if (seen.size === 0) {
      columnTypes.push("mixed");
    } else if (seen.size === 1) {
      const only = [...seen][0] as "number" | "date" | "text";
      columnTypes.push(only);
    } else {
      columnTypes.push("mixed");
    }
  }

  return columnTypes;
}

type PromoteRangeArgs = {
  values: any[][];
  valueBounds: BoundsSnapshot;
  baseRow: number;
  baseCol: number;
  usedRangeAddress: string | null;
  sheetName: string;
  maxPreviewRows: number;
};

function promoteRangeBlockFromUsedRange(args: PromoteRangeArgs): BlockSnapshot | null {
  const { values, valueBounds, baseRow, baseCol, usedRangeAddress, sheetName, maxPreviewRows } = args;
  if (valueBounds.firstRow === null || valueBounds.firstCol === null || valueBounds.lastRow === null || valueBounds.lastCol === null) {
    return null;
  }
  const trimmed = trimValuesToBounds(values, valueBounds, baseRow, baseCol);
  if (trimmed.length === 0) return null;
  const headerRow = trimmed[0] || [];
  if (!isLikelyHeaderRow(headerRow)) return null;
  const dataRows = trimmed.slice(1);
  const nonEmptyDataRows = dataRows.filter((r) => r.some((v) => v !== null && v !== undefined && v !== ""));
  if (nonEmptyDataRows.length === 0) return null;
  const dataRowsPassing = nonEmptyDataRows.filter((row) => {
    const nonEmpty = row.filter((v) => v !== null && v !== undefined && v !== "");
    if (nonEmpty.length === 0) return false;
    const acceptable = nonEmpty.filter((v) => typeof v === "number" || typeof v === "string");
    return acceptable.length >= Math.ceil(nonEmpty.length * 0.5);
  });
  if (dataRowsPassing.length === 0) return null;

  const headers = headerRow.map((v) => (v === null || v === undefined ? "" : String(v)));
  const preview = trimmed.slice(0, maxPreviewRows).map((row) => row.map(serializeCell));
  const columnTypes = inferColumnTypes(preview.length > 0 ? preview : [headerRow]);
  const address = valueBounds.address || usedRangeAddress || makeAddress(valueBounds.firstRow, valueBounds.firstCol, valueBounds.lastRow, valueBounds.lastCol);

  return {
    id: `${sheetName}!${address}`,
    address,
    kind: "range",
    confidence: 0.65,
    headerRowIndex: 0,
    headers,
    columnTypes,
    preview,
    source: { type: "range" },
    duplicateOf: null,
  };
}

function isLikelyHeaderRow(row: any[]): boolean {
  const nonEmpty = row.filter((v) => v !== null && v !== undefined && v !== "");
  if (nonEmpty.length === 0) return false;
  const strings = nonEmpty.filter((v) => typeof v === "string" && String(v).trim() !== "");
  return strings.length >= Math.ceil(nonEmpty.length * 0.6);
}

function trimValuesToBounds(values: any[][], bounds: BoundsSnapshot, baseRow: number, baseCol: number): any[][] {
  const { firstRow, firstCol, lastRow, lastCol } = bounds;
  if (firstRow === null || firstCol === null || lastRow === null || lastCol === null) return [];
  const startR = Math.max(0, firstRow - baseRow);
  const startC = Math.max(0, firstCol - baseCol);
  const rows = Math.max(0, lastRow - firstRow + 1);
  const cols = Math.max(0, lastCol - firstCol + 1);
  const result: any[][] = [];
  for (let r = 0; r < rows; r += 1) {
    const row = values[startR + r] ?? [];
    result.push(row.slice(startC, startC + cols));
  }
  return result;
}

function containsRange(outer: RangeBounds, inner: RangeBounds): boolean {
  return (
    inner.startRow >= outer.startRow &&
    inner.endRow <= outer.endRow &&
    inner.startCol >= outer.startCol &&
    inner.endCol <= outer.endCol
  );
}

export const __test__ = {
  parseAddressBounds,
  stripSheetPrefix,
  findValueBounds,
  inferColumnTypes,
};
