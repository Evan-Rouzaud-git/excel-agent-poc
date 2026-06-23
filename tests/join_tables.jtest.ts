import fs from "fs";
import path from "path";
import { runAgentPipeline, ExcelAdapter } from "../src/taskpane/agent/pipeline/runAgentPipeline";
import { normalizePlan } from "../src/taskpane/agent/planner/normalizePlan";
import { FakeContext, FakeWorkbook, FakeWorksheet } from "./mocks/fakeExcel";
import { WorkbookContextSnapshot } from "../src/taskpane/context/types";

const fixtures = (name: string) => fs.readFileSync(path.join(__dirname, "fixtures", name), "utf8");

const tsvToMatrix = (tsv: string): any[][] => {
  const trimmed = tsv.trim();
  const parts = trimmed.split("\n").map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));
  return parts.map((line) => line.split("\t"));
};

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

const letterToCol = (letters: string) => {
  let filtered = "";
  for (const ch of letters) {
    const upper = ch.toUpperCase();
    if (upper >= "A" && upper <= "Z") filtered += upper;
  }
  let col = 0;
  for (let i = 0; i < filtered.length; i += 1) col = col * 26 + (filtered.charCodeAt(i) - 64);
  return col - 1;
};

const rowColToA1 = (row: number, col: number) => `${colToLetter(col)}${row + 1}`;

function rangeFromArtifact(wb: FakeWorkbook, artifact: any) {
  const ws = wb.worksheets.getItem(artifact.sheet as any);
  const anchor = artifact?.anchor || "A1";
  const cols = artifact?.cols || 1;
  const rows = (artifact?.rows || 0) + 1; // include header
  const cellToRowCol = (cell: string) => {
    let letters = "";
    let digits = "";
    for (const ch of cell) {
      if (letters && (ch >= "0" && ch <= "9")) {
        digits += ch;
      } else if (ch.toUpperCase() >= "A" && ch.toUpperCase() <= "Z") {
        letters += ch.toUpperCase();
      } else if (!letters && (ch >= "0" && ch <= "9")) {
        digits += ch;
      }
    }
    return { col: letterToCol(letters || "A"), row: digits ? parseInt(digits, 10) - 1 : 0 };
  };
  const parsed = cellToRowCol(anchor);
  const startCol = parsed.col;
  const startRow = parsed.row;
  const endCol = startCol + cols - 1;
  const endRow = startRow + rows - 1;
  const addr = `${anchor}:${rowColToA1(endRow, endCol)}`;
  return ws.getRange(addr);
}

const makeExcelAdapter = (ctx: FakeContext): ExcelAdapter => ({
  run: async (cb) => cb(ctx as any),
});

async function executeNormalizedPlan(plan: any, snapshot: WorkbookContextSnapshot, ctx: FakeContext) {
  const adapter = makeExcelAdapter(ctx);
  return runAgentPipeline({
    context: snapshot,
    excelAdapter: adapter,
    plan,
    autoAnswerMode: "demoEval",
    maxAttempts: 3,
  });
}

function makeSheetWithData(sheetName: string, startCell: string, data: any[][], columnTypes?: string[]) {
  const cellToRowCol = (cell: string) => {
    let letters = "";
    let digits = "";
    for (const ch of cell) {
      if (letters && ch >= "0" && ch <= "9") {
        digits += ch;
      } else if (ch.toUpperCase() >= "A" && ch.toUpperCase() <= "Z") {
        letters += ch.toUpperCase();
      } else if (!letters && ch >= "0" && ch <= "9") {
        digits += ch;
      }
    }
    return { col: letterToCol(letters || "A"), row: digits ? parseInt(digits, 10) - 1 : 0 };
  };
  const { col: startCol, row: startRow } = cellToRowCol(startCell);
  const rowCount = data.length;
  const colCount = data[0]?.length || 1;
  const endRow = startRow + rowCount - 1;
  const endCol = startCol + colCount - 1;
  const endCell = rowColToA1(endRow, endCol);
  const blockAddress = `${startCell}:${endCell}`;
  const blockId = `${sheetName}!${blockAddress}`;
  const headers = data[0] || [];
  const colTypes = columnTypes && columnTypes.length === headers.length ? columnTypes : headers.map(() => "text");
  const ws = new FakeWorksheet(sheetName);
  ws.getRange(blockAddress).values = data;

  const tableName = `${sheetName}Table`;
  const sheetSnapshot = {
    name: sheetName,
    usedRange: blockAddress,
    valueBounds: { firstRow: startRow, firstCol: startCol, lastRow: endRow, lastCol: endCol, address: `${sheetName}!${blockAddress}` },
    counts: { tables: 1, charts: 0 },
    tables: [
      {
        name: tableName,
        address: `${sheetName}!${blockAddress}`,
        dataBodyAddress: `${sheetName}!${rowColToA1(startRow + 1, startCol)}:${rowColToA1(endRow, endCol)}`,
        headerAddress: `${sheetName}!${rowColToA1(startRow, startCol)}:${rowColToA1(startRow, endCol)}`,
        headers,
      },
    ],
    blocks: [
      {
        id: blockId,
        address: blockAddress,
        kind: "table",
        confidence: 1,
        headerRowIndex: 0,
        headers,
        columnTypes: colTypes,
        preview: [],
        source: { type: "table", tableName, tableAddress: `${sheetName}!${blockAddress}` },
      },
    ],
    charts: [],
    limitations: [],
  };
  return { sheetSnapshot, worksheet: ws, blockId };
}

function buildContext(leftFixture: string, rightFixture: string, opts?: { leftTypes?: string[]; rightTypes?: string[] }) {
  const left = makeSheetWithData("SheetLeft", "A1", tsvToMatrix(fixtures(leftFixture)), opts?.leftTypes);
  const right = makeSheetWithData("SheetRight", "A1", tsvToMatrix(fixtures(rightFixture)), opts?.rightTypes);
  const snapshot: WorkbookContextSnapshot = {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "SheetLeft", selectionAddress: "A1", selectionInBlockId: left.blockId, nearestBlockId: left.blockId },
    capabilities: [],
    limitations: [],
    sheets: [left.sheetSnapshot as any, right.sheetSnapshot as any],
    totals: { sheets: 2, tables: 2, charts: 0, blocks: 2, durationMs: 0 },
  };
  const wb = new FakeWorkbook([left.worksheet, right.worksheet]);
  const ctx = new FakeContext(wb);
  return { snapshot, ctx, leftBlockRef: left.blockId, rightBlockRef: right.blockId, wb };
}

describe("join_tables macro", () => {
  test("A) left join 1:1 with different key names", async () => {
    const { snapshot, ctx, leftBlockRef, rightBlockRef, wb } = buildContext("join_A_left.tsv", "join_A_right.tsv");
    const plan = {
      version: "1.0",
      goal: "join villes",
      
      steps: [{ id: "j1", macro: "join_tables", params: { left: { blockRef: leftBlockRef }, right: { blockRef: rightBlockRef }, keys: [{ left: "CodeVille", right: "VilleCode" }] } }],
    };
    const normalized = normalizePlan(plan as any, snapshot);
    expect(normalized.steps[0].params.select.right.mode).toBe("all");
    expect((normalized.steps[0].params.select.right.columns || [])).toContain("Ville");
    const pipelineRes = await executeNormalizedPlan(normalized as any, snapshot, ctx as any);
    const execRes = pipelineRes.execution;
    expect(execRes.status).toBe("ok");
    const artifact = execRes.artifacts?.find((a: any) => a.type === "table") as any;
    expect(artifact).toBeTruthy();
    const values = rangeFromArtifact(wb, artifact).values;
    expect(values[0]).toEqual(["Projet", "CodeVille", "VilleCode", "Ville"]);
    expect(values.slice(1)).toEqual([
      ["Alpha", "V1", "V1", "Paris"],
      ["Beta", "V2", "", ""],
      ["Gamma", "V3", "V3", "Lyon"],
    ]);
    expect((artifact.sheet || "").toLowerCase()).toContain("join_result");
  });

  test("B) 1:n explode_rows duplication", async () => {
    const { snapshot, ctx, leftBlockRef, rightBlockRef, wb } = buildContext("join_B_left.tsv", "join_B_right.tsv");
    const plan = {
      version: "1.0",
      goal: "join ressources",
      
      steps: [{ id: "j1", macro: "join_tables", params: { left: { blockRef: leftBlockRef }, right: { blockRef: rightBlockRef }, keys: [{ left: "Projet", right: "ProjetCode" }] } }],
    };
    const normalized = normalizePlan(plan as any, snapshot);
    const pipelineRes = await executeNormalizedPlan(normalized as any, snapshot, ctx as any);
    const execRes = pipelineRes.execution;
    expect(execRes.status).toBe("ok");
    const artifact = execRes.artifacts?.find((a: any) => a.type === "table") as any;
    expect(artifact).toBeTruthy();
    const values = rangeFromArtifact(wb, artifact).values;
    expect(values[0]).toEqual(["Projet", "Budget", "ProjetCode", "Ressource"]);
    expect(values.slice(1)).toEqual([
      ["P1", "100", "P1", "R1"],
      ["P1", "100", "P1", "R2"],
      ["P2", "200", "P2", "R3"],
    ]);
  });

  test("C) column conflict uses suffix", async () => {
    const { snapshot, ctx, leftBlockRef, rightBlockRef, wb } = buildContext("join_C_left.tsv", "join_C_right.tsv");
    const plan = {
      version: "1.0",
      goal: "join villes conflict",
      
      steps: [{ id: "j1", macro: "join_tables", params: { left: { blockRef: leftBlockRef }, right: { blockRef: rightBlockRef }, keys: [{ left: "CodeVille", right: "VilleCode" }] } }],
    };
    const normalized = normalizePlan(plan as any, snapshot);
    const pipelineRes = await executeNormalizedPlan(normalized as any, snapshot, ctx as any);
    const execRes = pipelineRes.execution;
    expect(execRes.status).toBe("ok");
    const artifact = execRes.artifacts?.find((a: any) => a.type === "table") as any;
    expect(artifact).toBeTruthy();
    const values = rangeFromArtifact(wb, artifact).values;
    const headers = values[0];
    expect(headers).toEqual(["CodeVille", "Ville", "VilleCode", "Population"]);
    expect(values[1]).toEqual(["V1", "Paris-L", "V1", "2200000"]);
  });

  test("D) joinType inner vs left vs anti_left with no matches", async () => {
    const fixturesPair = { left: "join_D_left.tsv", right: "join_D_right.tsv" };

    const run = async (joinType: "inner" | "left" | "anti_left") => {
      const { snapshot, ctx, leftBlockRef, rightBlockRef, wb } = buildContext(fixturesPair.left, fixturesPair.right);
      const plan = {
        version: "1.0",
        goal: `join ${joinType}`,
        
        steps: [
          { id: "j1", macro: "join_tables", params: { left: { blockRef: leftBlockRef }, right: { blockRef: rightBlockRef }, keys: [{ left: "Code", right: "CodeVille" }], joinType } },
        ],
      };
      const normalized = normalizePlan(plan as any, snapshot);
      const pipelineRes = await executeNormalizedPlan(normalized as any, snapshot, ctx as any);
      const execRes = pipelineRes.execution;
      const artifact = execRes.artifacts?.find((a: any) => a.type === "table") as any;
      const values = rangeFromArtifact(wb, artifact).values;
      return { execution: execRes, values, artifact };
    };

    const inner = await run("inner");
    expect(inner.execution.status).toBe("ok");
    expect(inner.values.length).toBe(1); // header only

    const left = await run("left");
    expect(left.execution.status).toBe("ok");
    expect(left.values.slice(1).length).toBe(2);

    const anti = await run("anti_left");
    expect(anti.execution.status).toBe("ok");
    expect(anti.values.slice(1).length).toBe(2);
  });

  test("E) date columns get formatted", async () => {
    const { snapshot, ctx, leftBlockRef, rightBlockRef, wb } = buildContext("join_E_left.tsv", "join_E_right.tsv", {
      leftTypes: ["text", "number", "number"],
      rightTypes: ["text", "number", "number"],
    });
    const plan = {
      version: "1.0",
      goal: "join dates",
      
      steps: [{ id: "j1", macro: "join_tables", params: { left: { blockRef: leftBlockRef }, right: { blockRef: rightBlockRef }, keys: [{ left: "Projet", right: "ProjetCode" }] } }],
    };
    const normalized = normalizePlan(plan as any, snapshot);
    const pipelineRes = await executeNormalizedPlan(normalized as any, snapshot, ctx as any);
    const execRes = pipelineRes.execution;
    expect(execRes.status).toBe("ok");
    const artifact = execRes.artifacts?.find((a: any) => a.type === "table") as any;
    const range = rangeFromArtifact(wb, artifact);
    const values = range.values;
    expect(values[0]).toEqual(["Projet", "EDP_Debut", "EDP_Fin", "ProjetCode"]);
    const debutFmt = range.numberFormat?.[0]?.[1];
    const finFmt = range.numberFormat?.[0]?.[2];
    expect(debutFmt).toBe("yyyy-mm-dd");
    expect(finFmt).toBe("yyyy-mm-dd");
  });
});

