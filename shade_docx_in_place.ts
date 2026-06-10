import { promises as fs } from "fs";
import * as path from "path";
import JSZip = require("jszip");
import { DOMParser, XMLSerializer } from "xmldom";

/** Namespace WordprocessingML (Word) */
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/* -------------------- Helpers CSV -------------------- */

function stripBom(s: string) {
  return s.replace(/^\uFEFF/, "");
}

function splitSemicolon(line: string): string[] {
  // Conserve les vides (fin de ligne avec ';')
  return (line ?? "").split(";");
}

function norm(s: string) {
  return (s ?? "")
    .normalize("NFD")
    .replace(/\p{Mn}/gu, "") // retire accents
    .replace(/’/g, "'")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

async function readCsvSemicolon(file: string): Promise<string[][]> {
  const raw = stripBom(await fs.readFile(file, "utf8"));
  const lines = raw
    .split(/\r?\n/)
    .map((l) => l.trimEnd())
    .filter((l) => l.length > 0);
  const rows: string[][] = [];
  for (const l of lines) rows.push(splitSemicolon(l));
  return rows;
}

/* -------------------- Load mapping & compilation -------------------- */

async function loadMapping(mappingPath: string) {
  const rows = await readCsvSemicolon(mappingPath);
  if (rows.length < 2) {
    throw new Error("mapping.csv doit contenir au moins 2 lignes (labels puis codes).");
  }
  const labels = rows[0].map((x) => x ?? "");
  const codes = rows[1].map((x) => x ?? "");
  if (labels.length !== codes.length) {
    throw new Error("mapping.csv: les 2 premières lignes doivent avoir le même nombre de colonnes.");
  }
  const codeToCol = new Map<string, number>();
  codes.forEach((c, i) => {
    if (c) codeToCol.set(norm(c), i);
  });
  return { labels, codes, codeToCol };
}

async function loadCompilation(compilPath: string) {
  const rows = await readCsvSemicolon(compilPath);
  if (rows.length < 2) throw new Error("compilation_mot.csv vide ou incomplet.");
  const header = rows[0];
  if (!header.length || norm(header[0]) !== "ELEVE") {
    throw new Error("compilation_mot.csv: la première cellule de l'entête doit être 'eleve'.");
  }
  const headerCodes = header.slice(1).map((x) => x ?? "");
  const data: Array<{ eleve: number; marks: string[] }> = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r.length) continue;
    const n = Number((r[0] ?? "").trim());
    if (!Number.isFinite(n)) continue;
    const marks = r.slice(1).map((x) => (x ?? "").trim().toUpperCase()); // 'X' si erreur
    data.push({ eleve: n, marks });
  }
  return { headerCodes, data };
}

/* -------------------- DOM traversal without XPath prefixes -------------------- */

function findAllByLocalName(node: Node, local: string, out: Element[] = []): Element[] {
  if ((node as Element).localName === local) out.push(node as Element);
  for (let ch = node.firstChild; ch; ch = ch.nextSibling) {
    findAllByLocalName(ch, local, out);
  }
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

/* -------------------- Shading -------------------- */

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

/* -------------------- Main -------------------- */

async function main() {
  // CLI
  const args = process.argv.slice(2);
  const getArg = (name: string) => {
    const i = args.indexOf(name);
    return i >= 0 ? args[i + 1] : undefined;
  };

  const srcPath = getArg("--src");
  const outPath = getArg("--out") ?? "tableau_colore.docx";
  const mappingPath = getArg("--mapping");
  const compilPath = getArg("--compil");

  // Indices/offsets selon ta mise en page:
  const tableIndex = Number(getArg("--table-index") ?? "0");         // si plusieurs tableaux
  const headerRowIndex = Number(getArg("--header-row-index") ?? "1");// ligne des mots (souvent 1)
  const studentRowOffset = Number(getArg("--student-row-offset") ?? "1"); // élève 1 = index 2 => offset=1
  const firstDataColIndex = Number(getArg("--first-data-col-index") ?? "1"); // 0=si pas de colonne "num élève"; 1=si première col = numéro

  if (!srcPath || !mappingPath || !compilPath) {
    console.error("Usage:");
    console.error("  npx ts-node shade_docx_in_place.ts --src tableau_original.docx --mapping mapping.csv --compil compilation_mot.csv --out tableau_colore.docx");
    console.error("Options: --table-index 0 --header-row-index 1 --student-row-offset 1 --first-data-col-index 1");
    process.exit(1);
  }

  // CSV
  const { codes: mappingCodes, codeToCol } = await loadMapping(mappingPath);
  const { headerCodes, data } = await loadCompilation(compilPath);

  // Sanity checks (diffs de codes)
  const setA = new Set(mappingCodes.map((c) => norm(c)));
  const setB = new Set(headerCodes.map((c) => norm(c)));
  const missingInMap = [...setB].filter((c) => !setA.has(c));
  const missingInComp = [...setA].filter((c) => !setB.has(c));
  if (missingInMap.length) console.warn("⚠️ Codes dans compilation_mot.csv absents de mapping.csv:", missingInMap.sort());
  if (missingInComp.length) console.warn("⚠️ Codes dans mapping.csv absents de compilation_mot.csv:", missingInComp.sort());

  // Load DOCX (zip)
  const zip = await new JSZip().loadAsync(await fs.readFile(srcPath));
  const docXmlPath = "word/document.xml";
  const xmlBuf = await zip.file(docXmlPath)!.async("nodebuffer");
  const xmlStr = xmlBuf.toString("utf8");

  // Parse XML
  const dom = new DOMParser().parseFromString(xmlStr, "application/xml");
  const documentElement = dom.documentElement;

  // Tables
  const tbls = findAllByLocalName(documentElement, "tbl");
  if (!tbls.length) throw new Error("Aucune table trouvée.");
  const tbl = tbls[tableIndex] ?? tbls[0];

  // Rows of selected table
  const rows = childrenByLocalName(tbl, "tr");
  if (rows.length <= headerRowIndex) throw new Error("Index de la ligne des mots (headerRowIndex) hors bornes.");

  // Apply shading by student row and code mapping
  for (const { eleve, marks } of data) {
    const rowIndex = eleve + studentRowOffset;
    if (rowIndex >= rows.length) {
      console.warn(`⚠️ Ignoré: élève ${eleve} (row index ${rowIndex}) > nb lignes (${rows.length}).`);
      continue;
    }
    const tr = rows[rowIndex];
    const tcs = childrenByLocalName(tr, "tc");

    for (let i = 0; i < marks.length; i++) {
      if (marks[i] !== "X") continue;
      const code = headerCodes[i] ?? "";
      const colInMapping = codeToCol.get(norm(code));
      if (colInMapping == null) continue;

      // Décalage si 1re colonne = numéro d'élève
      const cellIndex = firstDataColIndex + colInMapping;

      if (cellIndex < tcs.length) {
        shadeCellRed(tcs[cellIndex], dom as unknown as Document);
      }
    }
  }

  // Serialize & write back
  const updatedXml = new XMLSerializer().serializeToString(dom);
  zip.file(docXmlPath, Buffer.from(updatedXml, "utf8"));
  const outBuf = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(path.resolve(outPath), outBuf);
  console.log(`✅ Écrit: ${outPath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
