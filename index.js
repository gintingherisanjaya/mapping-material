const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, "data-dource");
const RESULTS_DIR = path.join(ROOT, "results");

const TARGET_GLS = new Set(["51100003", "51100065", "51100106"]);

const GROUPS = [
  "Worm Screw",
  "Press Cage",
  "Shaft",
  "Blade Digester",
  "Gear",
  "Sparepart Press",
  "Bearing",
  "Pelumas",
  "Belt",
  "Hose",
  "Electrode",
  "Bolt",
  "Chain",
  "Cone",
  "Lain-Lain",
];

// keywords: array of substrings (case-insensitive); first matching group wins
const GROUP_RULES = [
  { keywords: ["worm"], group: "Worm Screw" },
  { keywords: ["cage"], group: "Press Cage" },
  { keywords: ["shaft", "mending", "thrust", "cap:protection"], group: "Shaft" },
  { keywords: ["blade", "digester"], group: "Blade Digester" },
  { keywords: ["gear"], group: "Gear" },
  { keywords: ["bearing", "brg"], group: "Bearing" },
  { keywords: ["pelumas", "lube oil", "grease"], group: "Pelumas" },
  { keywords: ["belt"], group: "Belt" },
  { keywords: ["hose"], group: "Hose" },
  { keywords: ["electrode"], group: "Electrode" },
  { keywords: ["bolt", "nut", "baut"], group: "Bolt" },
  { keywords: ["chain"], group: "Chain" },
  { keywords: ["cone"], group: "Cone" },
  { keywords: ["press"], group: "Sparepart Press" },
];

const MONTHS_ID = [
  "Januari",
  "Februari",
  "Maret",
  "April",
  "Mei",
  "Juni",
  "Juli",
  "Agustus",
  "September",
  "Oktober",
  "November",
  "Desember",
];

function assignGroup(materialName) {
  const lower = String(materialName || "").toLowerCase();
  for (const rule of GROUP_RULES) {
    if (rule.keywords.some((kw) => lower.includes(kw))) return rule.group;
  }
  return "Lain-Lain";
}

function normalizeGl(value) {
  if (value === null || value === undefined || value === "") return null;
  const raw = String(value).trim();
  if (/^\d+(\.0+)?$/.test(raw)) {
    const digits = String(Math.trunc(Number(raw)));
    if (/^\d{8}$/.test(digits)) return digits;
  }
  if (/^\d{8}$/.test(raw)) return raw;
  return null;
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value).replace(/,/g, "").trim();
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function cellText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && value.text != null) return String(value.text).trim();
  if (typeof value === "object" && value.richText) {
    return value.richText.map((p) => p.text || "").join("").trim();
  }
  if (typeof value === "object" && value.result != null) return String(value.result).trim();
  return String(value).trim();
}

function getCellValue(row, col) {
  return row.getCell(col).value;
}

function isJumlahRow(materialText) {
  return /^jumlah\s*:/i.test(materialText);
}

function wibTimestampBase() {
  const now = new Date(
    new Date().toLocaleString("en-US", { timeZone: "Asia/Jakarta" })
  );
  const day = String(now.getDate()).padStart(2, "0");
  const month = MONTHS_ID[now.getMonth()];
  const year = now.getFullYear();
  const hour = String(now.getHours()).padStart(2, "0");
  const minute = String(now.getMinutes()).padStart(2, "0");
  return `hasil_mapping_${day}_${month}_${year}_${hour}.${minute}_WIB`;
}

function listWorkbooks() {
  if (!fs.existsSync(DATA_DIR)) {
    throw new Error(`Folder data tidak ditemukan: ${DATA_DIR}`);
  }
  return fs
    .readdirSync(DATA_DIR)
    .filter((f) => /\.xlsx$/i.test(f) && !f.startsWith("~$"))
    .map((f) => path.join(DATA_DIR, f));
}

function extractFromSheet(worksheet) {
  const pks = worksheet.name;
  if (pks.trim().toLowerCase() === "rekap") return [];

  const rows = [];
  let currentGl = null;

  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const colA = getCellValue(row, 1);
    const colB = getCellValue(row, 2);
    const colE = getCellValue(row, 5);
    const colF = getCellValue(row, 6);

    const gl = normalizeGl(colA);
    if (gl) {
      currentGl = TARGET_GLS.has(gl) ? gl : null;
      return;
    }

    if (!currentGl) return;

    const material = cellText(colB);
    if (!material || isJumlahRow(material)) return;

    rows.push({
      gl_account: currentGl,
      material,
      jumlah: toNumber(colE),
      biaya: toNumber(colF) * 1000,
      group: assignGroup(material),
      pks,
    });
  });

  return rows;
}

async function extractAllBySource() {
  const files = listWorkbooks();
  if (files.length === 0) {
    throw new Error(`Tidak ada file .xlsx di ${DATA_DIR}`);
  }

  const bySource = [];
  for (const file of files) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file);
    const rows = [];
    for (const worksheet of workbook.worksheets) {
      rows.push(...extractFromSheet(worksheet));
    }
    bySource.push({
      sourcePath: file,
      sourceName: path.basename(file),
      rows,
    });
  }
  return bySource;
}

function buildSummaryMap(rows) {
  const map = new Map();
  for (const row of rows) {
    if (!map.has(row.pks)) {
      const init = {};
      for (const g of GROUPS) init[g] = { qty: 0, biaya: 0 };
      map.set(row.pks, init);
    }
    const entry = map.get(row.pks)[row.group];
    entry.qty += row.jumlah;
    entry.biaya += row.biaya;
  }
  return map;
}

/** Fill summary/REKAP layout into an existing worksheet (starts at row 3, col B). */
function fillSummarySheet(sheet, rows) {
  const summaryMap = buildSummaryMap(rows);
  const pksList = [...summaryMap.keys()].sort((a, b) =>
    a.localeCompare(b, "id")
  );

  sheet.getCell("B3").value = "PKS";
  sheet.getCell("B3").font = { bold: true };
  sheet.getCell("B3").alignment = { horizontal: "center", vertical: "middle" };

  GROUPS.forEach((group, i) => {
    const startCol = 3 + i * 2;
    const endCol = startCol + 1;
    const startCell = sheet.getCell(3, startCol);
    startCell.value = group;
    startCell.font = { bold: true };
    startCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.mergeCells(3, startCol, 3, endCol);

    const qtyCell = sheet.getCell(4, startCol);
    const biayaCell = sheet.getCell(4, endCol);
    qtyCell.value = "Qty";
    biayaCell.value = "Biaya";
    qtyCell.font = { bold: true };
    biayaCell.font = { bold: true };
    qtyCell.alignment = { horizontal: "center" };
    biayaCell.alignment = { horizontal: "center" };
  });

  sheet.getColumn(2).width = 16;
  GROUPS.forEach((_, i) => {
    sheet.getColumn(3 + i * 2).width = 12;
    sheet.getColumn(4 + i * 2).width = 16;
  });

  pksList.forEach((pks, idx) => {
    const rowNum = 5 + idx;
    sheet.getCell(rowNum, 2).value = pks;
    const data = summaryMap.get(pks);
    GROUPS.forEach((group, i) => {
      const qtyCol = 3 + i * 2;
      const biayaCol = qtyCol + 1;
      const qtyCell = sheet.getCell(rowNum, qtyCol);
      const biayaCell = sheet.getCell(rowNum, biayaCol);
      qtyCell.value = data[group].qty;
      biayaCell.value = data[group].biaya;
      qtyCell.numFmt = "#,##0.###";
      biayaCell.numFmt = "#,##0";
    });
  });
}

function moveSheetToFront(workbook, sheet) {
  sheet.orderNo = 0;
  workbook.worksheets
    .filter((ws) => ws.id !== sheet.id)
    .forEach((ws, i) => {
      ws.orderNo = i + 1;
    });
}

function getOrReplaceRekapSheet(workbook) {
  const existing = workbook.worksheets.find(
    (ws) => ws.name.trim().toLowerCase() === "rekap"
  );
  if (existing) {
    workbook.removeWorksheet(existing.id);
  }
  const rekap = workbook.addWorksheet("REKAP", { state: "visible" });
  moveSheetToFront(workbook, rekap);
  return rekap;
}

async function writeCombinedOutput(allRows, baseName) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "mapping-material";
  workbook.created = new Date();

  const rawSheet = workbook.addWorksheet("raw material");
  rawSheet.columns = [
    { header: "gl_account", key: "gl_account", width: 14 },
    { header: "material", key: "material", width: 60 },
    { header: "jumlah", key: "jumlah", width: 14 },
    { header: "biaya", key: "biaya", width: 18 },
    { header: "group", key: "group", width: 18 },
    { header: "pks", key: "pks", width: 18 },
  ];
  rawSheet.getRow(1).font = { bold: true };
  for (const row of allRows) rawSheet.addRow(row);
  rawSheet.getColumn("jumlah").numFmt = "#,##0.###";
  rawSheet.getColumn("biaya").numFmt = "#,##0";

  const summary = workbook.addWorksheet("summary");
  fillSummarySheet(summary, allRows);

  const outPath = path.join(RESULTS_DIR, `${baseName}.xlsx`);
  await workbook.xlsx.writeFile(outPath);
  return outPath;
}

async function writePerSourceOutputs(bySource, baseName) {
  const folder = path.join(RESULTS_DIR, baseName);
  fs.mkdirSync(folder, { recursive: true });

  const written = [];
  for (const { sourcePath, sourceName, rows } of bySource) {
    const destPath = path.join(folder, sourceName);
    fs.copyFileSync(sourcePath, destPath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(destPath);

    const rekap = getOrReplaceRekapSheet(workbook);
    fillSummarySheet(rekap, rows);

    await workbook.xlsx.writeFile(destPath);
    written.push(destPath);
  }
  return { folder, written };
}

async function main() {
  if (!fs.existsSync(RESULTS_DIR)) {
    fs.mkdirSync(RESULTS_DIR, { recursive: true });
  }

  console.log("Membaca workbook dari data-dource/ ...");
  const bySource = await extractAllBySource();
  const allRows = bySource.flatMap((s) => s.rows);
  console.log(
    `Ditemukan ${allRows.length} baris material dari ${bySource.length} workbook.`
  );

  const baseName = wibTimestampBase();
  const combinedPath = await writeCombinedOutput(allRows, baseName);
  console.log(`Gabungan: ${combinedPath}`);

  const { folder, written } = await writePerSourceOutputs(bySource, baseName);
  console.log(`Per source (${written.length} file) di: ${folder}`);
  for (const p of written) console.log(`  - ${path.basename(p)}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
