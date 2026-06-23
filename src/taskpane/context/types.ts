export type BlockKind = "table" | "matrix" | "note" | "title" | "range" | "unknown";

export interface ContextOptions {
  maxSheets?: number;
  maxUsedRows?: number;
  maxUsedCols?: number;
  maxPreviewRows?: number;
  maxBlocksPerSheet?: number;
  logTimings?: boolean;
  includeCharts?: boolean;
  verboseLogs?: boolean;
}

export interface WorkbookContextSnapshot {
  workbook: {
    name: string | null;
    readOnly: boolean | null;
  };
  active: {
    sheetName: string | null;
    selectionAddress: string | null;
    selectionInBlockId: string | null;
    nearestBlockId: string | null;
  };
  capabilities: string[];
  limitations: string[];
  sheets: SheetSnapshot[];
  totals: {
    sheets: number;
    tables: number;
    charts: number;
    blocks: number;
    durationMs: number;
  };
}

export interface SheetSnapshot {
  name: string;
  usedRange: string | null;
  valueBounds: BoundsSnapshot;
  counts: {
    tables: number;
    charts: number;
  };
  tables: TableSnapshot[];
  blocks: BlockSnapshot[];
  charts: ChartSnapshot[];
  limitations: string[];
}

export interface BoundsSnapshot {
  firstRow: number | null;
  firstCol: number | null;
  lastRow: number | null;
  lastCol: number | null;
  address: string | null;
}

export interface TableSnapshot {
  name: string;
  address: string;
  dataBodyAddress: string | null;
  headerAddress: string | null;
  headers: string[];
}

export interface BlockSnapshot {
  id: string;
  address: string;
  kind: BlockKind;
  confidence: number;
  headerRowIndex: number | null;
  headers: string[];
  columnTypes: Array<"number" | "date" | "text" | "mixed">;
  preview: any[][];
  source: { type: "table"; tableName: string; tableAddress: string } | { type: "heuristic" } | { type: "range" };
  priorityScore?: number;
  signature?: string | null;
  duplicateOf?: string | null;
}

export interface ChartSnapshot {
  name: string;
  chartType: string;
}

export const DEFAULT_CONTEXT_OPTIONS: Required<ContextOptions> = {
  maxSheets: 30,
  maxUsedRows: 2000,
  maxUsedCols: 200,
  maxPreviewRows: 5,
  maxBlocksPerSheet: 40,
  logTimings: false,
  includeCharts: true,
  verboseLogs: false,
};

export interface ContextLogger {
  info(step: string, message: string, sheet?: string): void;
  warn(step: string, message: string, sheet?: string, err?: any): void;
  error(step: string, message: string, sheet?: string, err?: any): void;
}
