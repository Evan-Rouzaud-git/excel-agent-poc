import { executePlan } from "../src/taskpane/agent/executor";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";

function buildTravauxProjetSnapshot() {
  const leftWs = new FakeWorksheet("Travaux");
  const rightWs = new FakeWorksheet("Projet");
  leftWs.getRange("A1:E5").values = [
    ["ptvx_id", "charge_aff_code", "tache_nom", "tache_debut_dt", "tache_fin_dt"],
    ["Lyon", "SHO", "Travaux", "13/09/2026", "07/11/2027"],
    ["Paris", "NTE", "Travaux", "10/02/2026", "02/12/2027"],
    ["Bordeaux", "SMA", "Travaux", "19/01/2027", "10/09/2028"],
    ["Marseille", "API", "Réception travaux", "01/10/2028", "01/10/2028"],
  ];
  rightWs.getRange("A1:E5").values = [
    ["Projet", "code", "Typologie", "m2", "Hab"],
    ["Lyon", "78VER1", "Collectifs", 5020, 20],
    ["Paris", "75PAR1", "Maisons individuelles", 2500, 15],
    ["Bordeaux", "33BOR1", "Collectifs + MI", 5000, 75],
    ["Marseille", "13MAR1", "Grand Collectif", 7000, 80],
  ];
  const leftId = "Travaux!A1:E5";
  const rightId = "Projet!A1:E5";
  const snapshot: WorkbookContextSnapshot = {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "Travaux", selectionAddress: "A1", selectionInBlockId: leftId, nearestBlockId: leftId },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: "Travaux",
        usedRange: "A1:E5",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4, address: "A1:E5" },
        counts: { tables: 1, charts: 0 },
        tables: [{ name: "TravauxTbl", address: "Travaux!A1:E5", dataBodyAddress: "Travaux!A2:E5", headerAddress: "Travaux!A1:E1", headers: ["ptvx_id", "charge_aff_code", "tache_nom", "tache_debut_dt", "tache_fin_dt"] }],
        blocks: [
          {
            id: leftId,
            address: "A1:E5",
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: ["ptvx_id", "charge_aff_code", "tache_nom", "tache_debut_dt", "tache_fin_dt"],
            columnTypes: ["text", "text", "text", "text", "text"],
            preview: [],
            source: { type: "table", tableName: "TravauxTbl", tableAddress: "Travaux!A1:E5" },
          },
        ],
        charts: [],
        limitations: [],
      },
      {
        name: "Projet",
        usedRange: "A1:E5",
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4, address: "A1:E5" },
        counts: { tables: 1, charts: 0 },
        tables: [{ name: "ProjetTbl", address: "Projet!A1:E5", dataBodyAddress: "Projet!A2:E5", headerAddress: "Projet!A1:E1", headers: ["Projet", "code", "Typologie", "m2", "Hab"] }],
        blocks: [
          {
            id: rightId,
            address: "A1:E5",
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: ["Projet", "code", "Typologie", "m2", "Hab"],
            columnTypes: ["text", "text", "text", "number", "number"],
            preview: [],
            source: { type: "table", tableName: "ProjetTbl", tableAddress: "Projet!A1:E5" },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 2, tables: 2, charts: 0, blocks: 2, durationMs: 0 },
  };
  const ctx = new FakeContext(new FakeWorkbook([leftWs, rightWs]));
  return { snapshot, ctx, leftId, rightId, leftWs };
}

describe("join_tables fallback and defaults", () => {
  test("fallback picks better key when initial match is zero", async () => {
    const { snapshot, ctx, leftId, rightId, leftWs } = buildTravauxProjetSnapshot();
    const plan = {
      version: "1.0",
      goal: "concat travaux projet",
      
      steps: [
        {
          id: "j1",
          macro: "join_tables",
          params: {
            left: { blockRef: leftId },
            right: { blockRef: rightId },
            keys: [{ left: "charge_aff_code", right: "code" }], // mauvaise clé -> 0 match, doit fallback
            allowKeyFallback: true,
          },
        },
      ],
    };
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const artifact = res.artifacts.find((a) => a.type === "table") as any;
    expect(artifact.matchedRows).toBeGreaterThan(0);
    expect(artifact.outputRows).toBeGreaterThan(0);
    const joinSheet = ctx.workbook.worksheets.rawItems().find((ws: any) => ws.name !== "Travaux" && ws.name !== "Projet");
    expect(joinSheet).toBeDefined();
    const ranges = Object.values((joinSheet as any).ranges || {}) as any[];
    const firstRange = ranges.find((r) => Array.isArray(r.values) && r.values.length > 0);
    expect(firstRange).toBeDefined();
    const headers = firstRange?.values?.[0] || [];
    const normHeaders = headers.map((h: string) => (h || "").toLowerCase());
    const idx = headers.findIndex((h: string) => h === "charge_aff_code");
    expect(idx).toBeGreaterThanOrEqual(0);
    expect(firstRange?.values?.[1]?.[idx]).toBeTruthy();
    expect(headers.includes("code")).toBe(true);
    expect(normHeaders.filter((h: string) => h === "projet").length).toBeLessThanOrEqual(1);
  });

  test("default concat includes all columns both sides", async () => {
    const { snapshot, ctx, leftId, rightId, leftWs } = buildTravauxProjetSnapshot();
    const plan = {
      version: "1.0",
      goal: "concat",
      
      steps: [{ id: "j1", macro: "join_tables", params: { left: { blockRef: leftId }, right: { blockRef: rightId }, keys: [{ left: "ptvx_id", right: "Projet" }] } }],
    };
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("ok");
    const artifact = res.artifacts.find((a) => a.type === "table") as any;
    expect(artifact.cols).toBeGreaterThan(7);
  });

  test("fallback without allowKeyFallback returns confirmation request", async () => {
    const { snapshot, ctx, leftId, rightId } = buildTravauxProjetSnapshot();
    const plan = {
      version: "1.0",
      goal: "concat travaux projet",
      
      steps: [
        {
          id: "j1",
          macro: "join_tables",
          params: {
            left: { blockRef: leftId },
            right: { blockRef: rightId },
            keys: [{ left: "charge_aff_code", right: "code" }],
          },
        },
      ],
    };
    const res = await executePlan(plan as any, snapshot, ctx as any, {});
    expect(res.status).toBe("need_user_confirmation");
    const conf = res.confirmationsRequested?.find((c: any) => c.id === "j1:join_key_fallback");
    expect(conf).toBeDefined();
    expect(conf?.choices?.length).toBeGreaterThanOrEqual(3);
  });
});

