import { applyConfirmationsToPlan } from "../src/taskpane/agent/applyConfirmations";
import { executePlan } from "../src/taskpane/agent/executor";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { validatePlan } from "../src/taskpane/agent/planSchema";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";
import { parseA1Address, rowColToA1 } from "../src/taskpane/agent/utils";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";

const HIGHLIGHT_RED = "#f8d7da";
const HIGHLIGHT_ORANGE = "#fde9d9";
const HIGHLIGHT_VIOLET = "#e4d2f5";

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

function makeSnapshot(data: any[][], sheetName = "Sheet1"): { snapshot: WorkbookContextSnapshot; ctx: FakeContext; blockRef: string } {
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
            name: "IssuesSrc",
            address,
            dataBodyAddress: `Sheet1!A2:${colToLetter(cols - 1)}${rows}`,
            headerAddress: `Sheet1!A1:${colToLetter(cols - 1)}1`,
            headers: data[0] as string[],
          },
        ],
        blocks: [
          {
            id: address,
            address,
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: data[0] as string[],
            columnTypes: new Array(cols).fill("text"),
            preview: [],
            source: { type: "table", tableName: "IssuesSrc", tableAddress: address },
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

describe("validate_data macro", () => {
  test("détecte missing, duplicate et bad_type sans clés explicites", async () => {
    const data = [
      ["Client", "Montant", "Date", "Email"],
      ["Acme", "1000", "2025-01-01", "contact@acme.com"],
      ["Beta", "2000", "2025-01-02", "contact@beta.com"],
      ["Acme", "1000", "2025-01-01", "contact@acme.com"], // doublon Email + Date/Montant/Client
      ["Gamma", "1.200,5", "2025-01-03", "contact@gamma.com"], // Montant invalide convertible
      ["Delta", "1500", "2026-13-01", "delta@example.com"], // Date invalide
      ["Epsilon", "", "2025-02-02", ""], // missing Montant + Email
      ["Zeta", "500", "31/02/2025", "contact@zeta.com"], // Date invalide
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot(data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Valide les données et génère Issues pro.",
        steps: [
          {
            id: "validate1",
            macro: "validate_data",
            params: {
              source: { blockRef },
              detect: { missing: true, duplicates: true, badType: true },
            },
          },
        ],
      },
      snapshot
    );
    const validation = validatePlan(plan);
    expect(validation.valid).toBe(true);
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("need_user_confirmation");
    expect(res.confirmationsRequested?.some((c) => c.id === "validate_data:fix_missing")).toBe(true);
    const art = res.artifacts.find((a) => a.fromStep === "validate1" && a.type === "table");
    expect(art).toBeTruthy();
    const headers = ["Type", "RowKey", "Cellule", "Colonne", "Ligne", "Valeur", "Message", "Severity", "Fix"];
    expect((art as any).headers).toEqual(headers);
    const counts = (art as any).counts;
    expect(counts?.totalIssues).toBeGreaterThanOrEqual(3);
    expect(counts?.missing).toBeGreaterThanOrEqual(1);
    expect(counts?.duplicate).toBeGreaterThanOrEqual(1);
    expect(counts?.bad_type).toBeGreaterThanOrEqual(1);
    expect(res.logs.some((l) => l.message === "validate_data duplicates_found")).toBe(true);
    expect(res.logs.some((l) => l.message === "validate_data bad_type_found")).toBe(true);
    expect(res.logs.some((l) => l.message === "validate_data highlight_applied")).toBe(true);
    expect(res.logs.some((l) => l.message?.startsWith("validate_data: corporate_blue format applied"))).toBe(true);

    const sheetName = (art as any).sheet;
    const issueSheet = ctx.workbook.worksheets.getItem(sheetName);
    const issueRange = issueSheet.getRange((art as any).address);
    const issueRows = (issueRange.values as any[][]).slice(1);
    const types = issueRows.map((row) => row[0]);
    expect(types).toContain("missing");
    expect(types).toContain("duplicate");
    expect(types).toContain("bad_type");
    const fixes = issueRows.map((row) => row[8]);
    expect(fixes.every((value) => value === "apply")).toBe(true);
    const table = issueSheet.tables.getItem((art as any).tableName);
    expect(table).toBeTruthy();
    expect(issueSheet.freezePanes.freezeRows).toHaveBeenCalledWith(1);

    const dataSheet = ctx.workbook.worksheets.getItem("Sheet1");
    expect(dataSheet.getRange("C7").format.fill.color).toBe(HIGHLIGHT_RED);
    expect(dataSheet.getRange("E4").format.fill.color).toBe(HIGHLIGHT_ORANGE);
    expect(dataSheet.getRange("D6").format.fill.color).toBe(HIGHLIGHT_VIOLET);
    expect(dataSheet.getRange("D8").format.fill.color).toBe(HIGHLIGHT_VIOLET);

    const bounds = parseA1Address((art as any).address);
    expect(bounds).not.toBeNull();
    const startRow = bounds!.startRow;
    const startCol = bounds!.startCol;
    const typeOffsets: Record<string, number> = {};
    issueRows.forEach((row, idx) => {
      const type = row[0];
      if (type && typeOffsets[type] == null) typeOffsets[type] = idx;
    });
    const typeColors: Record<string, string> = {
      missing: HIGHLIGHT_RED,
      duplicate: HIGHLIGHT_ORANGE,
      bad_type: HIGHLIGHT_VIOLET,
    };
    Object.entries(typeColors).forEach(([type, expectedColor]) => {
      const offset = typeOffsets[type];
      expect(offset).not.toBeUndefined();
      if (offset === undefined) return;
      const address = rowColToA1(startRow + 1 + offset, startCol);
    expect(issueSheet.getRange(address).format.fill.color).toBe(expectedColor);
  });
});

  test("enregistre le NamedItem Issues et pose la première question après l'audit", async () => {
    const data = [
      ["Client", "Montant"],
      ["Acme", "1000"],
      ["Beta", ""], // missing Montant
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot(data);
    const plan = normalizePlan(
      {
        version: "1.0",
        goal: "Valider et persister Issues",
        steps: [
          {
            id: "validate1",
            macro: "validate_data",
            params: {
              source: { blockRef },
            },
          },
        ],
      },
      snapshot
    );
    const validation = validatePlan(plan);
    expect(validation.valid).toBe(true);
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("need_user_confirmation");
    expect(res.confirmationsRequested?.some((c) => c.id === "validate_data:fix_missing")).toBe(true);
    const artifact = res.artifacts.find((a) => a.fromStep === "validate1" && a.type === "table");
    expect(artifact).toBeTruthy();
    const sheetName = (artifact as any).sheet;
    const namedItem = ctx.workbook.names.getItem("__validate_data_last_issues_ref");
    expect(namedItem).toBeTruthy();
    const persistedRef = JSON.parse(namedItem.value);
    expect(persistedRef.sheet).toBe(sheetName);
    expect(res.logs.some((log) => log.message === "validate_data issues_reference_written")).toBe(true);
  });

  test("apply_fixes respecte la colonne Fix et supprime les doublons", async () => {
    const data = [
      ["Client", "Montant", "Date", "Email"],
      ["Acme", "1000", "2025-01-01", "contact@acme.com"],
      ["Beta", "2000", "2025-01-02", "contact@beta.com"],
      ["Acme", "1000", "2025-01-01", "contact@acme.com"],
      ["Gamma", "1.200,5", "2025-01-03", "contact@gamma.com"],
      ["Delta", "1500", "2026-13-01", "delta@example.com"],
      ["Epsilon", "", "2025-02-02", ""],
      ["Zeta", "500", "31/02/2025", "contact@zeta.com"],
    ];
    const { snapshot, ctx, blockRef } = makeSnapshot(data);
    const auditPlan = normalizePlan(
      {
        version: "1.0",
        goal: "Audit puis appliquer les issues",
        steps: [
          {
            id: "validate1",
            macro: "validate_data",
            params: {
              source: { blockRef },
              detect: { missing: true, duplicates: true, badType: true },
            },
          },
        ],
      },
      snapshot
    );
    const auditRes = await executePlan(auditPlan as any, snapshot, ctx as any, {});
    expect(auditRes.status).toBe("need_user_confirmation");
    expect(auditRes.confirmationsRequested?.some((c) => c.id === "validate_data:fix_missing")).toBe(true);
    const artifact = auditRes.artifacts.find((a) => a.fromStep === "validate1" && a.type === "table");
    expect(artifact).toBeTruthy();
    const artifactRecord = artifact as any;
    const issueSheet = ctx.workbook.worksheets.getItem(artifactRecord.sheet);
    const issueRange = issueSheet.getRange(artifactRecord.address);
    const issueValues = (issueRange.values as any[][]).slice(1);
    const missingEntries = issueValues
      .map((row, idx) => ({ row, idx }))
      .filter((entry) => entry.row[0] === "missing");
    const fixColIndex = artifactRecord.headers.findIndex((col: string) => col === "Fix");
    expect(fixColIndex).toBeGreaterThanOrEqual(0);
    const issueRangeAfterAudit = issueSheet.getRange(artifactRecord.address);
    const tableValuesAfterAudit = Array.isArray(issueRangeAfterAudit.values)
      ? issueRangeAfterAudit.values.map((row: any[]) => [...row])
      : [];
    missingEntries.forEach((entry) => {
      const targetRow = entry.idx + 1;
      if (tableValuesAfterAudit[targetRow]) {
        tableValuesAfterAudit[targetRow][fixColIndex] = "ignore";
      }
    });
    issueRangeAfterAudit.values = tableValuesAfterAudit;
    const questionIds = ["validate_data:fix_missing", "validate_data:fix_duplicates", "validate_data:fix_bad_type"];
    const decisions: Record<string, string> = {};
    let previousRes = auditRes;
    for (const [idx, questionId] of questionIds.entries()) {
      decisions[questionId] = "yes";
      const nextPlan = applyConfirmationsToPlan(auditPlan as any, decisions, snapshot);
      const res = await executePlan(
        nextPlan as any,
        snapshot,
        ctx as any,
        {
          confirmationDecisions: decisions,
          initialArtifacts: previousRes.artifacts,
        }
      );
      if (idx < questionIds.length - 1) {
        expect(res.status).toBe("need_user_confirmation");
        const expectedNext = questionIds[idx + 1];
        expect(res.confirmationsRequested?.some((c) => c.id === expectedNext)).toBe(true);
      } else {
        expect(res.status).toBe("ok");
        const summaryLog = res.logs.find((log) => log.message === "validate_data apply_summary");
        expect(summaryLog).toBeTruthy();
        const summaryData = summaryLog?.data || {};
        expect(summaryData.rows_deleted_missing).toBe(0);
        expect(summaryData.rows_deleted_duplicate).toBeGreaterThanOrEqual(1);
        expect(summaryData.casts_applied).toBeGreaterThanOrEqual(1);
      }
      previousRes = res;
    }
    const dataSheet = ctx.workbook.worksheets.getItem("Sheet1");
    const allRows = (dataSheet.getRange("A1:D8").values as any[][]) || [];
    const gammaRow = allRows.find((row) => row[1] === "Gamma");
    expect(gammaRow).toBeTruthy();
    if (!gammaRow) return;
    expect(gammaRow[2]).toBe(1200.5);
    expect(allRows.some((row) => row[1] === "Epsilon")).toBe(true);
  });

  test("accepte les blockRef locaux, qualifies et quotes", async () => {
    const data = [
      ["Client", "Montant"],
      ["Acme", "1000"],
      ["Beta", ""],
    ];
    const cases = [
      { sheetName: "Sheet1", blockRef: "A1:B3" },
      { sheetName: "Sheet1", blockRef: "Sheet1!A1:B3" },
      { sheetName: "Sheet 1", blockRef: "'Sheet 1'!A1:B3" },
    ];

    for (const testCase of cases) {
      const { snapshot, ctx } = makeSnapshot(data, testCase.sheetName);
      const plan = normalizePlan(
        {
          version: "1.0",
          goal: "Valide les données",
          steps: [
            {
              id: "validate1",
              macro: "validate_data",
              params: {
                source: { blockRef: testCase.blockRef },
                detect: { missing: true },
              },
            },
          ],
        },
        snapshot
      );
      const res = await executePlan(plan as any, snapshot, ctx as any, {});
      expect(res.status).toBe("need_user_confirmation");
      expect(res.confirmationsRequested?.some((c) => c.id === "validate_data:fix_missing")).toBe(true);
    }
  });
});
