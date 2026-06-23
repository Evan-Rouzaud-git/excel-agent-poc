import { WorkbookContextSnapshot } from "../context/types";
import { ArtifactRecord } from "./types";

export type AddressBounds = { sheet?: string; startRow: number; startCol: number; endRow: number; endCol: number };

export function colToLetter(col: number): string {
  let n = col + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

export function letterToCol(letter: string): number {
  let col = 0;
  let filtered = "";
  for (const ch of letter) {
    const code = ch.charCodeAt(0);
    const isLetter = (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
    if (isLetter) filtered += ch;
  }
  const normalized = filtered.toUpperCase();
  for (let i = 0; i < normalized.length; i += 1) {
    col = col * 26 + (normalized.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function rowColToA1(row: number, col: number): string {
  return `${colToLetter(col)}${row + 1}`;
}

export function makeRangeAddress(sheet: string, startRow: number, startCol: number, rowCount: number, colCount: number): string {
  const endRow = startRow + rowCount - 1;
  const endCol = startCol + colCount - 1;
  const a1 = `${rowColToA1(startRow, startCol)}:${rowColToA1(endRow, endCol)}`;
  return `${sheet}!${a1}`;
}

export function parseA1Address(address: string): AddressBounds | null {
  if (!address) return null;
  const trimmed = address.trim();
  const bangIdx = trimmed.lastIndexOf("!");
  let sheet: string | undefined;
  let rangePart = trimmed;
  if (bangIdx > -1) {
    sheet = trimmed.slice(0, bangIdx);
    rangePart = trimmed.slice(bangIdx + 1);
    if (sheet.startsWith("'") && sheet.endsWith("'") && sheet.length >= 2) {
      sheet = sheet.slice(1, -1);
    }
  }

  const parseCell = (cell: string): { row: number; col: number } | null => {
    let idx = 0;
    if (cell.charAt(idx) === "$") idx += 1;
    let letters = "";
    while (idx < cell.length) {
      const ch = cell.charAt(idx);
      const code = ch.charCodeAt(0);
      const isLetter = (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
      if (!isLetter) break;
      letters += ch;
      idx += 1;
    }
    if (!letters) return null;
    if (cell.charAt(idx) === "$") idx += 1;
    let digits = "";
    while (idx < cell.length) {
      const ch = cell.charAt(idx);
      const isDigit = ch >= "0" && ch <= "9";
      if (!isDigit) return null;
      digits += ch;
      idx += 1;
    }
    if (!digits) return null;
    return { row: parseInt(digits, 10) - 1, col: letterToCol(letters) };
  };

  const parts = rangePart.split(":");
  const first = parts[0] ? parseCell(parts[0]) : null;
  const second = parts.length > 1 && parts[1] ? parseCell(parts[1]) : null;
  if (!first) return null;
  const startRow = first.row;
  const startCol = first.col;
  const endRow = second ? second.row : startRow;
  const endCol = second ? second.col : startCol;
  return { sheet, startRow, startCol, endRow, endCol };
}

export function isBlankCell(val: any): boolean {
  return val === null || val === undefined || val === "";
}

export function isRangeBlank(values: any[][] | undefined | null): boolean {
  if (!values) return true;
  for (const row of values) {
    for (const cell of row) {
      if (!isBlankCell(cell)) return false;
    }
  }
  return true;
}

export function nowIso(): string {
  return new Date().toISOString();
}

export function findBlock(blockRef: string, snapshot: WorkbookContextSnapshot) {
  const parsed = parseA1Address(blockRef);
  const sheetName = parsed?.sheet || blockRef.split("!")[0];
  const sheet = snapshot.sheets.find((s) => s.name === sheetName);
  if (!sheet) return { sheet: undefined, block: undefined };
  const block = sheet.blocks.find((b) => b.id === blockRef || b.address === blockRef || `${sheetName}!${b.address}` === blockRef);
  return { sheet, block };
}

export type BlockArtifactRef = { blockRef?: string | null; artifactRef?: string | null };

export function resolveBlockRefInput(
  ref: BlockArtifactRef | undefined,
  artifacts: ArtifactRecord[]
): { ok: true; blockRef: string; artifact?: ArtifactRecord } | { ok: false; reason: string } {
  if (!ref) return { ok: false, reason: "missing_ref" };
  const artifactRef = (ref.artifactRef || "").trim();
  const stepId = artifactRef ? (artifactRef.includes(".") ? artifactRef.split(".")[0] || artifactRef : artifactRef) : null;
  if (ref.blockRef && typeof ref.blockRef === "string") {
    const directMatch = artifacts
      .slice()
      .reverse()
      .find((a) => (stepId ? a.fromStep === stepId : true) && (a.blockRef === ref.blockRef || `${a.sheet || a.sheetName}!${a.address || a.addressA1}` === ref.blockRef));
    return { ok: true, blockRef: ref.blockRef, artifact: directMatch };
  }
  if (!artifactRef) return { ok: false, reason: "missing_ref" };
  const matches = artifacts.filter((a) => a.fromStep === stepId);
  if (!matches.length) return { ok: false, reason: "artifact_not_found" };
  const tableArtifacts = matches.filter((a) => a.type === "table");
  const chosen = (tableArtifacts.length ? tableArtifacts[tableArtifacts.length - 1] : matches[matches.length - 1]) as ArtifactRecord;
  const sheetName = chosen.sheetName || chosen.sheet;
  const address = chosen.blockRef || (sheetName && (chosen.address || chosen.addressA1) ? `${sheetName}!${chosen.address || chosen.addressA1}` : null);
  const anchorAddress = !address && sheetName && chosen.anchor ? `${sheetName}!${chosen.anchor}` : null;
  const finalRef = address || anchorAddress || chosen.blockRef;
  if (!finalRef) return { ok: false, reason: "artifact_missing_blockRef" };
  return { ok: true, blockRef: finalRef, artifact: chosen };
}
