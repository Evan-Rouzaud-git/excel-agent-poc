import { snapshotSummary } from "./demoEvalRunner";

describe("snapshotSummary", () => {
  test("handles array sheets and tables", () => {
    const snapshot: any = {
      workbook: { name: "Book", readOnly: false },
      active: { sheetName: "S1", selectionAddress: "A1", selectionInBlockId: null, nearestBlockId: null },
      capabilities: [],
      limitations: [],
      sheets: [{ name: "S1", tables: [{ name: "T1", address: "A1:B2" }], blocks: [], charts: [], usedRange: "", valueBounds: null as any, counts: null as any, limitations: [] }],
      totals: { sheets: 1, tables: 1, charts: 0, blocks: 0, durationMs: 0 },
    };
    const res = snapshotSummary(snapshot);
    expect(res.sheetCount).toBe(1);
    expect(res.sheets[0]?.tables[0]?.name).toBe("T1");
  });

  test("handles collection items()", () => {
    const sheetObj: any = { name: "S1", tables: { items: [{ name: "Tbl", address: "A1" }] } };
    const snapshot: any = { sheets: [sheetObj] };
    const res = snapshotSummary(snapshot as any);
    expect(res.sheetCount).toBe(1);
    expect(res.sheets[0]?.tables.length).toBe(1);
  });

  test("handles rawItems()", () => {
    const sheetObj: any = { name: "S1", tables: { rawItems: () => [{ name: "Tbl2", address: "A1" }] } };
    const res = snapshotSummary({ sheets: [sheetObj] } as any);
    expect(res.sheetCount).toBe(1);
    expect(res.sheets[0]?.tables[0]?.name).toBe("Tbl2");
  });
});
