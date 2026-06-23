import { sanitizeTableViewParams } from "../src/taskpane/agent/tableViewUtils";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

const makeSnapshot = (): { snapshot: WorkbookContextSnapshot; blockRef: string } => {
  const headers = ["Projet", "Ville", "Budget"];
  const blockRef = "Sheet1!A1:C3";
  const snapshot: WorkbookContextSnapshot = {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: blockRef, nearestBlockId: blockRef },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: "Sheet1",
        usedRange: "A1:C3",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2, address: "A1:C3" },
        counts: { tables: 1, charts: 0 },
        tables: [{ name: "Tbl", address: blockRef, dataBodyAddress: "Sheet1!A2:C3", headerAddress: "Sheet1!A1:C1", headers }],
        blocks: [
          {
            id: blockRef,
            address: "A1:C3",
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers,
            columnTypes: ["text", "text", "number"],
            preview: [],
            source: { type: "table", tableName: "Tbl", tableAddress: blockRef },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
  };
  return { snapshot, blockRef };
};

describe("table_view sanitize extras", () => {
  test("select unknown header trimmed with fallback", () => {
    const { snapshot, blockRef } = makeSnapshot();
    const warnings: string[] = [];
    const out = sanitizeTableViewParams(
      { source: { blockRef }, select: ["Projet", "Inconnu"], filter: [], rename: { Projet: "ProjetX" } },
      snapshot,
      warnings
    );
    expect(out.select).toEqual(["Projet"]); // keep Projet, trim Inconnu
    expect(out.rename?.Projet).toBe("ProjetX");
  });

  test("filter op <> mapped to notEmpty", () => {
    const { snapshot, blockRef } = makeSnapshot();
    const out = sanitizeTableViewParams(
      { source: { blockRef }, select: ["Projet", "Ville"], filter: [{ col: "Ville", op: "<>" }] },
      snapshot
    );
    expect(out.filter?.[0]?.op).toBe("notEmpty");
  });

  test("headerAliases option allows alias selection", () => {
    const { snapshot, blockRef } = makeSnapshot();
    const out = sanitizeTableViewParams(
      { source: { blockRef }, select: ["ptvx_id", "Ville"] },
      snapshot,
      [],
      { headerAliases: { ptvx_id: ["ptvx_id", "Projet"], projet: ["ptvx_id", "Projet"] } }
    );
    expect(out.select).toEqual(["Projet", "Ville"]);
  });
});
