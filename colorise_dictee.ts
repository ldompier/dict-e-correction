import { promises as fs } from "fs";
import * as path from "path";
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, ShadingType, AlignmentType } from "docx";

/**
 * Utilities
 */
function stripBom(s: string): string {
  return s.replace(/^\uFEFF/, "");
}

function splitCsvSemicolon(line: string): string[] {
  // Simple split for semicolon-separated values (no quotes in your samples)
  return line.split(";").map((x) => x);
}

function normalize(s: string): string {
  return s
    .normalize("NFD")
    .replace(/\p{Mn}/gu, "") // strip accents
    .replace(/’/g, "'")
    .trim();
}

/**
 * Read mapping.csv
 * Line 1: labels (as displayed in header of table)
 * Line 2: codes (e.g., 1,2,3a,3b,4,... in same order as labels)
 */
async function readMapping(mappingPath: string): Promise<{ labels: string[]; codes: string[] }> {
  const raw = stripBom(await fs.readFile(mappingPath, "utf8"));
  const lines = raw
    .split(/\r?\n/)
    .map((l) => l.trimEnd())
    .filter((l) => l.length > 0);
  if (lines.length < 2) throw new Error("mapping.csv doit contenir au moins 2 lignes (labels puis codes).");
  const labels = splitCsvSemicolon(lines[0]).map((x) => x.trim());
  const codes = splitCsvSemicolon(lines[1]).map((x) => x.trim());
  if (labels.length !== codes.length) {
    throw new Error("mapping.csv: le nombre de colonnes des 2 premières lignes doit être identique.");
  }
  return { labels, codes };
}

/**
 * Read compilation_mot.csv
 * Header: eleve;1;2;3a;3b;...
 * Then rows: "<num_eleve>;x;;x;...;"
 */
async function readCompilation(compilPath: string): Promise<{ headerCodes: string[]; rows: Array<{ eleve: number; marks: string[] }> }> {
  const raw = stripBom(await fs.readFile(compilPath, "utf8"));
  const lines = raw
    .split(/\r?\n/)
    .map((l) => l.trimEnd())
    .filter((l) => l.length > 0);

  if (lines.length < 2) throw new Error("compilation_mot.csv vide ou incomplet.");

  const header = splitCsvSemicolon(lines[0]).map((x) => x.trim());
  if (header.length < 2 || normalize(header[0]).toUpperCase() !== "ELEVE") {
    throw new Error("compilation_mot.csv: la première cellule de l'entête doit être 'eleve'.");
  }
  const headerCodes = header.slice(1).map((c) => c.trim());

  const rows: Array<{ eleve: number; marks: string[] }> = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitCsvSemicolon(lines[i]);
    if (!cols.length) continue;
    const eleveStr = (cols[0] ?? "").trim();
    const eleve = Number(eleveStr);
    if (!Number.isFinite(eleve)) continue;
    const marks = cols.slice(1).map((x) => (x || "").trim().toUpperCase()); // normalize 'x'
    rows.push({ eleve, marks });
  }

  return { headerCodes, rows };
}

/**
 * Build code -> column index map from mapping order.
 * We do not try to guess header text in an existing doc; we use mapping order as ground truth.
 */
function buildCodeToColIndex(mappingCodes: string[]): Map<string, number> {
  const map = new Map<string, number>();
  mappingCodes.forEach((code, idx) => {
    if (code) map.set(normalize(code).toUpperCase(), idx);
  });
  return map;
}

/**
 * Create the DOCX with table:
 * - Row 0: empty first cell + mapping labels
 * - Then one row per student: first cell = "1", "2", ... ; cells shaded red where marks == 'X'
 */
async function createDocx(
  outPath: string,
  labels: string[],
  mappingCodes: string[],
  headerCodes: string[],
  studentRows: Array<{ eleve: number; marks: string[] }>
): Promise<void> {
  // Build code->col index mapping from the mapping.csv (authoritative order)
  const codeToCol = buildCodeToColIndex(mappingCodes);

  // Map headerCodes (from compilation) to col indices using codeToCol
  const compColToDocxCol: number[] = headerCodes.map((c) => codeToCol.get(normalize(c).toUpperCase()) ?? -1);

  // Build table rows
  const rows: TableRow[] = [];

  // Header row: first cell blank " " then labels
  const headerCells: TableCell[] = [
    new TableCell({ children: [new Paragraph({ text: "", alignment: AlignmentType.CENTER })] }),
    ...labels.map(
      (lbl) =>
        new TableCell({
          children: [new Paragraph({ text: lbl, alignment: AlignmentType.CENTER })],
        })
    ),
  ];
  rows.push(new TableRow({ children: headerCells }));

  // For each student row: place red shading where marks == 'X'
  // We will allocate cells count = 1 + labels.length
  for (const { eleve, marks } of studentRows) {
    const cellCount = 1 + labels.length;
    const cellArray: TableCell[] = new Array(cellCount).fill(null as any);

    // First cell = student number
    cellArray[0] = new TableCell({ children: [new Paragraph({ text: String(eleve), alignment: AlignmentType.CENTER })] });

    // Initialize all other cells empty
    for (let c = 1; c < cellCount; c++) {
      cellArray[c] = new TableCell({ children: [new Paragraph({ text: "", alignment: AlignmentType.CENTER })] });
    }

    // Shade cells where mark == 'X' and we have a valid mapped column index
    for (let i = 0; i < marks.length; i++) {
      if (marks[i] !== "X") continue;
      const docxCol = compColToDocxCol[i]; // 0-based among labels
      if (docxCol >= 0 && docxCol < labels.length) {
        const targetIndex = 1 + docxCol; // +1 because first column is the student number
        cellArray[targetIndex] = new TableCell({
          children: [new Paragraph({ text: "", alignment: AlignmentType.CENTER })],
          shading: { type: ShadingType.CLEAR, color: "auto", fill: "FF0000" },
        });
      }
    }

    rows.push(new TableRow({ children: cellArray }));
  }

  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows,
  });

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: "Tableau de compilations des résultats de la dictée diagnostique, 4e année" }),
          new Paragraph({ text: "" }),
          table,
        ],
      },
    ],
  });

  const buf = await Packer.toBuffer(doc);
  await fs.writeFile(outPath, buf);
}

/**
 * CLI
 * Usage:
 *   ts-node colorise_dictee.ts --mapping mapping.csv --compil compilation_mot.csv --out sortie.docx
 */
async function main() {
  const args = process.argv.slice(2);
  let mappingPath = "";
  let compilPath = "";
  let outPath = "tableau_colore.docx";

  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === "--mapping") mappingPath = args[++i];
    else if (a === "--compil") compilPath = args[++i];
    else if (a === "--out") outPath = args[++i];
  }

  if (!mappingPath || !compilPath) {
    console.error("Usage: ts-node colorise_dictee.ts --mapping mapping.csv --compil compilation_mot.csv --out tableau_colore.docx");
    process.exit(1);
  }

  const { labels, codes } = await readMapping(mappingPath);
  const { headerCodes, rows } = await readCompilation(compilPath);

  // log helpful diffs if codes differ
  const setA = new Set(codes.map((c) => normalize(c).toUpperCase()));
  const setB = new Set(headerCodes.map((c) => normalize(c).toUpperCase()));
  const missingInMapping: string[] = [];
  const missingInCompil: string[] = [];
  for (const b of setB) if (!setA.has(b)) missingInMapping.push(b);
  for (const a of setA) if (!setB.has(a)) missingInCompil.push(a);
  if (missingInMapping.length) console.warn("⚠️ Codes présents dans compilation_mot.csv mais absents du mapping.csv:", missingInMapping.sort());
  if (missingInCompil.length) console.warn("⚠️ Codes présents dans mapping.csv mais absents du compilation_mot.csv:", missingInCompil.sort());

  await createDocx(path.resolve(outPath), labels, codes, headerCodes, rows);
  console.log(`✅ Fichier généré: ${outPath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
