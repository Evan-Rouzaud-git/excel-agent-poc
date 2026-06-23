import { sanitizeTableViewParams } from "../src/taskpane/agent/tableViewUtils";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

const snapshot: WorkbookContextSnapshot = {
  workbook: { name: "Book", readOnly: false },
  active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: "Sheet1!A1:B3", nearestBlockId: "Sheet1!A1:B3" },
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
          kind: "table",
          confidence: 1,
          headerRowIndex: 0,
          headers: ["Projet", "Budget"],
          columnTypes: ["text", "number"],
          preview: [],
          source: { type: "range" } as any,
        },
      ],
      charts: [],
      limitations: [],
    },
  ],
  totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
};

describe("sanitizeTableViewParams", () => {
  test("drops malformed filter and normalizes sort/default dir", () => {
    const warnings: string[] = [];
    const out = sanitizeTableViewParams(
      {
        source: { blockRef: "Sheet1!A1:B3" },
        select: ["Projet", "Budget"],
        filter: "Ville=Paris",
        sort: { col: "Budget" },
      },
      snapshot,
      warnings
    );
    expect(out.filter).toEqual([]);
    expect(out.sort?.col).toBe("Budget");
    expect(out.sort?.dir).toBe("asc");
    expect(warnings).toContain("table_view_filter_invalid");
  });

  test("invalid select tokens trigger unknown_header warning (no silent substitution)", () => {
    const warnings: string[] = [];
    const out = sanitizeTableViewParams(
      {
        source: { blockRef: "Sheet1!A1:B3" },
        select: ["Z"],
        filter: [],
      },
      snapshot,
      warnings
    );
    expect(out.select).toEqual([]);
    expect(warnings).toContain("table_view_select_unresolved");
    expect(warnings.some((w) => w.startsWith("unknown_header:Z"))).toBe(true);
  });

  test("resolves sort column against extra headers when select is empty", () => {
    const warnings: string[] = [];
    const out = sanitizeTableViewParams(
      {
        source: { blockRef: "Sheet1!A1:B3" },
        sort: { col: "densite", dir: "asc" },
      },
      snapshot,
      warnings,
      {
        allowTokens: ["$lastAddedColumn"],
        extraHeaders: ["densite"],
      }
    );
    expect(out.sort?.col).toBe("densite");
    expect(warnings.some((w) => w.startsWith("unknown_header"))).toBe(false);
  });

  test("adds missing sort/filter columns to select", () => {
    const warnings: string[] = [];
    const out = sanitizeTableViewParams(
      {
        source: { blockRef: "Sheet1!A1:B3" },
        select: ["$lastAddedColumn"],
        sort: { col: "Projet", dir: "asc" },
        filter: [{ col: "Budget", op: "gt", value: 100 }],
      },
      snapshot,
      warnings,
      {
        allowTokens: ["$lastAddedColumn"],
      }
    );
    expect(out.select).toEqual(["$lastAddedColumn", "Projet", "Budget"]);
  });
});
