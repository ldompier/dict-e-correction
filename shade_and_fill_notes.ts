// shade_and_fill_notes.ts
import { promises as fs } from "fs";
import * as path from "path";
import JSZip = require("jszip");
import { DOMParser, XMLSerializer } from "xmldom";

/** WordprocessingML namespace */
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/* ================= Helpers CSV ================= */

function stripBom(s: string) { return s.replace(/^\uFEFF/, ""); }
function splitSemicolon(line: string): string[] { return (line ?? "").split(";"); }
function norm(s: string) {
  return (s ?? "")
    .normalize("NFD").replace(/\p{Mn}/gu, "")
    .replace(/’/g, "'").trim().replace(/\s+/g, " ")
    .toUpperCase();
}

async function readCsvSemicolon(file: string): Promise<string[][]> {
  const raw = stripBom(await fs.readFile(file, "utf8"));
  const lines = raw.split(/\r?\n/).map(l => l.trimEnd()).filter(l => l.length > 0);
  const rows: string[][] = [];
  for (const l of lines) rows.push(splitSemicolon(l));
  return rows;
}

/* ================= mapping_v2 & compilation ================= */
/** mapping_v2.csv
 * row0: labels (affichés)
 * row1: codes  (1,2,3a,3b, …) — même ordre que row0
 * row2: niveaux '3' ou '4'
 */
async function loadMappingV2(mappingPath: string) {
  const rows = await readCsvSemicolon(mappingPath);
  if (rows.length < 3) throw new Error("mapping_v2.csv doit avoir au moins 3 lignes (labels, codes, niveaux).");
  const labels = rows[0].map(x => x ?? "");
  const codes  = rows[1].map(x => x ?? "");
  const levels = rows[2].map(x => (x ?? "").trim());
  if (labels.length !== codes.length || codes.length !== levels.length) {
    throw new Error("mapping_v2.csv: les 3 premières lignes doivent avoir le même nombre de colonnes.");
  }
  const codeToCol = new Map<string, number>();
  const levelByIndex = new Map<number, 3 | 4>();
  let total3 = 0, total4 = 0;
  codes.forEach((c, i) => {
    if (c) codeToCol.set(norm(c), i);
    const lv = levels[i] === "4" ? 4 : 3;
    levelByIndex.set(i, lv);
    if (lv === 3) total3++; else total4++;
  });
  return { labels, codes, codeToCol, levelByIndex, total3, total4 };
}

/** compilation_mot.csv
 * header: eleve;1;2;3a;3b;…
 * rows: "<num_eleve>;x;;x;…"
 */
async function loadCompilation(compilPath: string) {
  const rows = await readCsvSemicolon(compilPath);
  if (rows.length < 2) throw new Error("compilation_mot.csv vide ou incomplet.");
  const header = rows[0];
  if (!header.length || norm(header[0]) !== "ELEVE") {
    throw new Error("compilation_mot.csv: la première cellule du header doit être 'eleve'.");
  }
  const headerCodes = header.slice(1).map(x => x ?? "");
  const data: Array<{ eleve: number; marks: string[] }> = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r.length) continue;
    const n = Number((r[0] ?? "").trim());
    if (!Number.isFinite(n)) continue;
    const marks = r.slice(1).map(x => (x ?? "").trim().toUpperCase()); // 'X' = erreur
    data.push({ eleve: n, marks });
  }
  return { headerCodes, data };
}

/* ================= XML helpers sans prefixes ================= */

function findAllByLocalName(node: Node, local: string, out: Element[] = []): Element[] {
  if ((node as Element).localName === local) out.push(node as Element);
  for (let ch = node.firstChild; ch; ch = ch.nextSibling) findAllByLocalName(ch, local, out);
  return out;
}
function childrenByLocalName(el: Element, local: string): Element[] {
  const res: Element[] = [];
  for (let ch = el.firstChild; ch; ch = ch.nextSibling) {
    if ((ch as Element).localName === local) res.push(ch as Element);
  }
  return res;
}
function firstChildByLocalName(el: Element, local: string): Element | null {
  for (let ch = el.firstChild; ch; ch = ch.nextSibling) {
    if ((ch as Element).localName === local) return ch as Element;
  }
  return null;
}

/* ================= Shading cellule (rouge) ================= */

function ensureTcPr(tc: Element, doc: Document): Element {
  let tcPr = firstChildByLocalName(tc, "tcPr");
  if (!tcPr) {
    tcPr = doc.createElementNS(W_NS, "w:tcPr");
    tc.insertBefore(tcPr, tc.firstChild);
  }
  return tcPr;
}
function clearShading(tcPr: Element) {
  const toRemove = childrenByLocalName(tcPr, "shd");
  for (const n of toRemove) tcPr.removeChild(n);
}
function shadeCellRed(tc: Element, doc: Document) {
  const tcPr = ensureTcPr(tc, doc);
  clearShading(tcPr);
  const shd = doc.createElementNS(W_NS, "w:shd");
  shd.setAttributeNS(W_NS, "w:val", "clear");
  shd.setAttributeNS(W_NS, "w:color", "auto");
  shd.setAttributeNS(W_NS, "w:fill", "FF0000");
  tcPr.appendChild(shd);
}

/* ================= Notes & Moyenne ================= */

type Notes = {
  eleve: number;
  total3e: number;
  total3ePct: number;
  total4e: number;
  total4ePct: number;
  total345: number;
  total345Pct: number;
};

type ScoreMode = "correct" | "errors";

/** Calcule les scores:
 *  - mode "correct": compte les colonnes SANS 'X'
 *  - mode "errors":  compte les colonnes AVEC  'X'
 */
function computeNotesForAll(
  headerCodes: string[],
  students: Array<{ eleve: number; marks: string[] }>,
  mappingCodes: string[],
  levelByIndex: Map<number, 3 | 4>,
  codeToCol: Map<string, number>,
  total3: number,
  total4: number,
  mode: ScoreMode = "correct"
): Notes[] {
  const totalAll = total3 + total4;
  const N = mappingCodes.length;

  // header code -> index dans le mapping
  const headerToMapIdx: number[] = headerCodes.map(c => codeToCol.get(norm(c)) ?? -1);

  const out: Notes[] = [];
  for (const { eleve, marks } of students) {
    // erreurs par colonne du mapping
    const err: boolean[] = new Array(N).fill(false);
    for (let i = 0; i < marks.length; i++) {
      if (marks[i] !== "X") continue;
      const mi = headerToMapIdx[i];
      if (mi >= 0 && mi < N) err[mi] = true;
    }

    let c3 = 0, c4 = 0;
    for (let idx = 0; idx < N; idx++) {
      const lv = levelByIndex.get(idx);
      if (!lv) continue;
      const isErr = err[idx];
      const hit = (mode === "errors") ? isErr : !isErr;
      if (hit) {
        if (lv === 3) c3++;
        else if (lv === 4) c4++;
      }
    }

    const p3 = total3 ? Math.round((c3 * 100) / total3) : 0;
    const p4 = total4 ? Math.round((c4 * 100) / total4) : 0;
    const cAll = c3 + c4;
    const pAll = totalAll ? Math.round((cAll * 100) / totalAll) : 0;

    out.push({ eleve, total3e: c3, total3ePct: p3, total4e: c4, total4ePct: p4, total345: cAll, total345Pct: pAll });
  }
  out.sort((a, b) => a.eleve - b.eleve);
  return out;
}

/** Ligne "Moyenne" (arithmétique, arrondie) */
function computeAverageRow(notes: Notes[]): Notes {
  const n = notes.length || 1;
  const sum = (f: (x: Notes) => number) => notes.reduce((a, b) => a + f(b), 0) / n;
  return {
    eleve: -1,
    total3e: Math.round(sum(x => x.total3e)),
    total3ePct: Math.round(sum(x => x.total3ePct)),
    total4e: Math.round(sum(x => x.total4e)),
    total4ePct: Math.round(sum(x => x.total4ePct)),
    total345: Math.round(sum(x => x.total345)),
    total345Pct: Math.round(sum(x => x.total345Pct)),
  };
}

/* ================= Ecriture cellule texte ================= */

function setCellText(tc: Element, doc: Document, text: string) {
  // remove all <w:p>
  const ps = childrenByLocalName(tc, "p");
  for (const p of ps) tc.removeChild(p);
  // write one simple paragraph
  const p = doc.createElementNS(W_NS, "w:p");
  const r = doc.createElementNS(W_NS, "w:r");
  const t = doc.createElementNS(W_NS, "w:t");
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t); p.appendChild(r); tc.appendChild(p);
}

/* ================= Remplissage du template Notes ================= */

async function fillNotesTemplate(
  templatePath: string,
  outPath: string,
  notes: Notes[],
  totals: { total3: number; total4: number },
  startRowIndex = 1,   // ligne 2 (index 1)
  studentColIndex = 0  // col 0 = numéro d’élève
) {
  const totalAll = totals.total3 + totals.total4;
  const zip = await new JSZip().loadAsync(await fs.readFile(templatePath));
  const docXmlPath = "word/document.xml";
  const xmlBuf = await zip.file(docXmlPath)!.async("nodebuffer");
  const xmlStr = xmlBuf.toString("utf8");

  const dom = new DOMParser().parseFromString(xmlStr, "application/xml");
  const documentElement = dom.documentElement;

  const tbls = findAllByLocalName(documentElement, "tbl");
  if (!tbls.length) throw new Error("Template: aucun tableau trouvé.");
  const tbl = tbls[0];

  const rows = childrenByLocalName(tbl, "tr");

  // Ajoute la ligne Moyenne à la fin
  const avg = computeAverageRow(notes);
  const fullList = [...notes, avg];

  let r = startRowIndex;
  for (const n of fullList) {
    if (r >= rows.length) break;
    const tr = rows[r++];
    const tcs = childrenByLocalName(tr, "tc");

    // libellé 1ère colonne
    const label = n.eleve === -1 ? "Moyenne" : String(n.eleve);
    if (studentColIndex < tcs.length) setCellText(tcs[studentColIndex], dom as unknown as Document, label);

    const cellsWanted = [
      `${n.total3e}/${totals.total3}`,   // 3e /27
      `${n.total3ePct}`,                 // 3e /100
      `${n.total4e}/${totals.total4}`,   // 4e /18
      `${n.total4ePct}`,                 // 4e /100
      `${n.total345}/${totalAll}`,       // 3e+4e /45
      `${n.total345Pct}`,                // 3e+4e /100
    ];
    for (let i = 0; i < cellsWanted.length; i++) {
      const ci = i + 1; // +1 car col 0 = élève
      if (ci < tcs.length) setCellText(tcs[ci], dom as unknown as Document, cellsWanted[i]);
    }
  }

  const updatedXml = new XMLSerializer().serializeToString(dom);
  zip.file(docXmlPath, Buffer.from(updatedXml, "utf8"));
  const outBuf = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(path.resolve(outPath), outBuf);
}

/* ================= Coloration in-place du tableau original ================= */

async function shadeInPlace(
  srcPath: string,
  outPath: string,
  headerRowIndex: number,
  studentRowOffset: number,
  firstDataColIndex: number,
  headerCodes: string[],
  data: Array<{ eleve: number; marks: string[] }>,
  codeToCol: Map<string, number>
) {
  const zip = await new JSZip().loadAsync(await fs.readFile(srcPath));
  const docXmlPath = "word/document.xml";
  const xmlBuf = await zip.file(docXmlPath)!.async("nodebuffer");
  const xmlStr = xmlBuf.toString("utf8");

  const dom = new DOMParser().parseFromString(xmlStr, "application/xml");
  const documentElement = dom.documentElement;

  const tbls = findAllByLocalName(documentElement, "tbl");
  if (!tbls.length) throw new Error("Aucune table trouvée dans le document.");
  const tbl = tbls[0];

  const rows = childrenByLocalName(tbl, "tr");
  if (rows.length <= headerRowIndex) throw new Error("Index de la ligne des mots hors bornes.");

  for (const { eleve, marks } of data) {
    const rowIndex = eleve + studentRowOffset;
    if (rowIndex >= rows.length) continue;
    const tr = rows[rowIndex];
    const tcs = childrenByLocalName(tr, "tc");

    for (let i = 0; i < marks.length; i++) {
      if (marks[i] !== "X") continue;
      const code = headerCodes[i] ?? "";
      const colInMapping = codeToCol.get(norm(code));
      if (colInMapping == null) continue;

      const cellIndex = firstDataColIndex + colInMapping; // décalage si col 0 = numéro élève
      if (cellIndex < tcs.length) shadeCellRed(tcs[cellIndex], dom as unknown as Document);
    }
  }

  const updatedXml = new XMLSerializer().serializeToString(dom);
  zip.file(docXmlPath, Buffer.from(updatedXml, "utf8"));
  const outBuf = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(path.resolve(outPath), outBuf);
}

/* ================= CLI ================= */

async function main() {
  const args = process.argv.slice(2);
  const getArg = (name: string) => { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : undefined; };

  const srcPath    = getArg("--src");                // tableau_original.docx
  const mappingV2  = getArg("--mapping");            // mapping_v2.csv (3 lignes)
  const compilPath = getArg("--compil");             // compilation_mot.csv
  const outTable   = getArg("--out") ?? "Tableau_colore.docx";
  const notesTpl   = getArg("--notes-template") ?? "Notes_eleves.docx";
  const notesOut   = getArg("--notes-out") ?? "Notes_eleves_renseigne.docx";

  // indices/offsets
  const headerRowIndex   = Number(getArg("--header-row-index") ?? "1"); // ligne des mots
  const studentRowOffset = Number(getArg("--student-row-offset") ?? "1"); // élève 1 -> index 2 => +1
  const firstDataColIdx  = Number(getArg("--first-data-col-index") ?? "1"); // col 0 = N° élève, donc 1

  // template notes : la 1re ligne de données commence à la ligne 2 (index 1)
  const notesStartRowIdx = Number(getArg("--notes-start-row-index") ?? "1");
  const notesStudentCol  = Number(getArg("--notes-student-col-index") ?? "0");

  // scores mode: correct (par défaut) ou errors
  const scoresModeArg = (getArg("--scores") ?? "correct").toLowerCase();
  const scoresMode: ScoreMode = (scoresModeArg === "errors") ? "errors" : "correct";

  if (!srcPath || !mappingV2 || !compilPath) {
    console.error("Usage:");
    console.error("  npx ts-node shade_and_fill_notes.ts --src tableau_original.docx --mapping mapping_v2.csv --compil compilation_mot.csv --out Tableau_colore.docx --notes-template Notes_eleves.docx --notes-out Notes_eleves_renseigne.docx [--scores correct|errors]");
    console.error("Options utiles : --header-row-index 1 --student-row-offset 1 --first-data-col-index 1 --notes-start-row-index 1 --notes-student-col-index 0");
    process.exit(1);
  }

  // charger CSV
  const { codes, codeToCol, levelByIndex, total3, total4 } = await loadMappingV2(mappingV2);
  const { headerCodes, data } = await loadCompilation(compilPath);

  // sanity info (codes manquants)
  const setMap = new Set(codes.map(c => norm(c)));
  const setComp = new Set(headerCodes.map(c => norm(c)));
  const missingInMap = [...setComp].filter(c => !setMap.has(c));
  const missingInComp = [...setMap].filter(c => !setComp.has(c));
  if (missingInMap.length) console.warn("⚠️ Codes compilation absents de mapping:", missingInMap.sort());
  if (missingInComp.length) console.warn("⚠️ Codes mapping absents de compilation:", missingInComp.sort());

  // 1) Coloration du tableau original
  await shadeInPlace(
    srcPath, outTable,
    headerRowIndex, studentRowOffset, firstDataColIdx,
    headerCodes, data, codeToCol
  );

  // 2) Calcul & remplissage du template Notes (avec ligne Moyenne)
  const notes = computeNotesForAll(headerCodes, data, codes, levelByIndex, codeToCol, total3, total4, scoresMode);
  await fillNotesTemplate(
    notesTpl, notesOut, notes,
    { total3, total4 },
    notesStartRowIdx, notesStudentCol
  );

  console.log(`✅ Tableau colorié : ${outTable}`);
  console.log(`✅ Notes renseignées : ${notesOut}`);
}

main().catch(e => { console.error(e); process.exit(1); });
