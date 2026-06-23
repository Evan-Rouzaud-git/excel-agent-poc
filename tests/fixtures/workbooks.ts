import { WorkbookContextSnapshot } from "../../src/taskpane/context/types";

const headers = ["Mois", "Depenses", "Revenus", "TVA", "Categorie", "Region", "Projet", "Statut", "Commentaire"];

export const budgetValues: any[][] = (() => {
  const months = ["Jan", "Fev", "Mar", "Avr", "Mai", "Juin", "Juil", "Aou", "Sep", "Oct", "Nov", "Dec"];
  const categories = ["Ventes", "Services", "Maintenance", "Support", "Avoir"];
  const regions = ["IDF", "Ouest", "NA"];
  const projets = ["PRJ-1", "PRJ-2", "PRJ-3", "PRJ-4", "PRJ-5"];
  const statuts = ["En cours", "Clôturé", "En retard", "Reporté"];
  const rows: any[][] = [headers];
  for (let i = 0; i < 100; i += 1) {
    const mois = months[i % months.length];
    const depensesBase = 600 + (i % 7) * 35;
    let revenus = 1000 + (i % 9) * 55 + (i % 4 === 0 ? -80 : 0);
    if (i % 12 === 4) revenus = 0;
    if (i % 15 === 0) revenus = 950.5;
    let depenses = depensesBase + (i % 5 - 2) * 40;
    if (i % 18 === 0) depenses = -150.75;
    const tva = Math.round(revenus * 0.2 * 100) / 100;
    const cat = categories[i % categories.length];
    const region = regions[i % regions.length];
    const projet = projets[i % projets.length];
    const statut = statuts[i % statuts.length];
    const commentaire =
      i % 11 === 0 ? "" : i % 9 === 0 ? "Avoir à traiter" : i % 7 === 0 ? "Vérifier facture" : "RAS";
    rows.push([mois, depenses, revenus, tva, cat, region, projet, statut, commentaire]);
  }
  return rows;
})();

export function workbook_table_view(): WorkbookContextSnapshot {
  const data = [
    ["Projet", "Ville", "Début EDP", "Fin EDP", "Budget", "Statut"],
    ["PRJ-Alpha", "Marseille", "01/02/2026", "15/05/2026", 120000, "En cours"],
    ["PRJ-Beta", "Lyon", "10/03/2026", "30/06/2026", 90000, ""],
    ["PRJ-Gamma", "Bordeaux", "", "20/07/2026", 110000, "En retard"],
    ["PRJ-Delta", "Marignane", "05/01/2026", "12/04/2026", 80000, "Clôturé"],
    ["PRJ-Epsilon", "", "18/02/2026", "22/05/2026", 70000, "En cours"],
  ];
  const address = "A1:F6";
  return {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "Projets", selectionAddress: "A1", selectionInBlockId: `Projets!${address}`, nearestBlockId: `Projets!${address}` },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: "Projets",
        usedRange: address,
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 5, lastCol: 5, address },
        counts: { tables: 1, charts: 0 },
        tables: [
          {
            name: "ProjetsTbl",
            address: `Projets!${address}`,
            dataBodyAddress: "Projets!A2:F6",
            headerAddress: "Projets!A1:F1",
            headers: data[0] as string[],
          },
        ],
        blocks: [
          {
            id: `Projets!${address}`,
            address,
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: data[0] as string[],
            columnTypes: ["text", "text", "text", "text", "number", "text"],
            preview: [],
            source: { type: "table", tableName: "ProjetsTbl", tableAddress: `Projets!${address}` },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
  };
}

export function workbook_join_travaux(): WorkbookContextSnapshot {
  const leftData = [
    ["ptvx_id", "charge_aff_code", "tache_nom", "tache_debut_dt", "tache_fin_dt"],
    ["Lyon", "SHO", "Travaux", "13/09/2026", "07/11/2027"],
    ["Paris", "NTE", "Travaux", "10/02/2026", "02/12/2027"],
    ["Bordeaux", "SMA", "Travaux", "19/01/2027", "10/09/2028"],
    ["Marseille", "API", "Réception travaux", "01/10/2028", "01/10/2028"],
  ];
  const rightData = [
    ["Projet", "code", "Typologie", "m2", "Hab"],
    ["Lyon", "78VER1", "Collectifs", 5020, 20],
    ["Paris", "75PAR1", "Maisons individuelles", 2500, 15],
    ["Bordeaux", "33BOR1", "Collectifs + MI", 5000, 75],
    ["Marseille", "13MAR1", "Grand Collectif", 7000, 80],
  ];
  const leftAddr = "A1:E5";
  const rightAddr = "A1:E5";
  const leftId = `Travaux!${leftAddr}`;
  const rightId = `Projet!${rightAddr}`;
  return {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "Travaux", selectionAddress: "A1", selectionInBlockId: leftId, nearestBlockId: leftId },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: "Travaux",
        usedRange: leftAddr,
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4, address: leftAddr },
        counts: { tables: 1, charts: 0 },
        tables: [
          {
            name: "TravauxTbl",
            address: `Travaux!${leftAddr}`,
            dataBodyAddress: "Travaux!A2:E5",
            headerAddress: "Travaux!A1:E1",
            headers: leftData[0] as string[],
          },
        ],
        blocks: [
          {
            id: leftId,
            address: leftAddr,
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: leftData[0] as string[],
            columnTypes: ["text", "text", "text", "text", "text"],
            preview: [],
            source: { type: "table", tableName: "TravauxTbl", tableAddress: `Travaux!${leftAddr}` },
          },
        ],
        charts: [],
        limitations: [],
      },
      {
        name: "Projet",
        usedRange: rightAddr,
        valueBounds: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4, address: rightAddr },
        counts: { tables: 1, charts: 0 },
        tables: [
          {
            name: "ProjetTbl",
            address: `Projet!${rightAddr}`,
            dataBodyAddress: "Projet!A2:E5",
            headerAddress: "Projet!A1:E1",
            headers: rightData[0] as string[],
          },
        ],
        blocks: [
          {
            id: rightId,
            address: rightAddr,
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers: rightData[0] as string[],
            columnTypes: ["text", "text", "text", "number", "number"],
            preview: [],
            source: { type: "table", tableName: "ProjetTbl", tableAddress: `Projet!${rightAddr}` },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 2, tables: 2, charts: 0, blocks: 2, durationMs: 0 },
  };
}

export function workbook_simple_sales(): WorkbookContextSnapshot {
  const lastRow = budgetValues.length - 1; // 100 data rows + header
  const lastCol = headers.length - 1; // 9 columns
  const address = "A1:I101";
  const tableName = "TableBudget";
  const blockId = `Sheet1!${address}`;
  return {
    workbook: { name: "Book", readOnly: false },
    active: { sheetName: "Sheet1", selectionAddress: "A1", selectionInBlockId: blockId, nearestBlockId: blockId },
    capabilities: [],
    limitations: [],
    sheets: [
      {
        name: "Sheet1",
        usedRange: address,
        valueBounds: { firstRow: 0, firstCol: 0, lastRow, lastCol, address },
        counts: { tables: 1, charts: 0 },
        tables: [
          {
            name: tableName,
            address: `Sheet1!${address}`,
            dataBodyAddress: "Sheet1!A2:I101",
            headerAddress: "Sheet1!A1:I1",
            headers,
          },
        ],
        blocks: [
          {
            id: blockId,
            address,
            kind: "table",
            confidence: 1,
            headerRowIndex: 0,
            headers,
            columnTypes: ["text", "number", "number", "number", "text", "text", "text", "text", "text"],
            preview: [],
            source: { type: "table", tableName, tableAddress: `Sheet1!${address}` },
          },
        ],
        charts: [],
        limitations: [],
      },
    ],
    totals: { sheets: 1, tables: 1, charts: 0, blocks: 1, durationMs: 0 },
  };
}
