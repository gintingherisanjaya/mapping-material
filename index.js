const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const ROOT = __dirname;
const RESULTS_DIR = path.join(ROOT, "results");
const MASTER_PKS_PATH = path.join(ROOT, "master_pks.xlsx");

const TARGET_GLS = new Set(["51100003", "51100065", "51100106"]);
const SIMILARITY_THRESHOLD = 0.55;
const RED_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFF0000" },
};

const PRESS_GROUPS = [
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
const PRESS_GROUP_RULES = [
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

const RAIL_TRACK_GROUPS = ["Rope", "Wire Rope", "Lain-Lain"];

const ROMAN_REGIONAL = {
  i: 1,
  ii: 2,
  iii: 3,
  iv: 4,
  v: 5,
  vi: 6,
  vii: 7,
  viii: 8,
  ix: 9,
  x: 10,
};

/** Runtime mode — di-set di main() dari CLI args. */
let DATA_DIR = path.join(ROOT, "data-dource");
let GROUPS = PRESS_GROUPS;
let GROUP_RULES = PRESS_GROUP_RULES;
let OUTPUT_PREFIX = "hasil_mapping";
let MODE_NAME = "press";

function configureMode(isRailTrack) {
  if (isRailTrack) {
    MODE_NAME = "rail-track";
    DATA_DIR = path.join(ROOT, "rail-track-data-source");
    GROUPS = RAIL_TRACK_GROUPS;
    GROUP_RULES = null;
    OUTPUT_PREFIX = "hasil_mapping_rail_track";
  } else {
    MODE_NAME = "press";
    DATA_DIR = path.join(ROOT, "data-dource");
    GROUPS = PRESS_GROUPS;
    GROUP_RULES = PRESS_GROUP_RULES;
    OUTPUT_PREFIX = "hasil_mapping";
  }
}

function assignGroup(materialName) {
  const lower = String(materialName || "").toLowerCase();

  if (MODE_NAME === "rail-track") {
    // Wire Rope dulu — supaya tidak tertangkap rule Rope
    if (
      lower.includes("wire rope") ||
      lower.includes("rope,wire") ||
      lower.includes("rope wire") ||
      lower.includes("wire,rope") ||
      lower.includes("staal drad") ||
      lower.includes("staal draad") ||
      lower.includes("staal drat")
    ) {
      return "Wire Rope";
    }
    if (
      /\brope\b/.test(lower) ||
      lower.includes("tali") ||
      lower.includes("nilon")
    ) {
      return "Rope";
    }
    return "Lain-Lain";
  }

  for (const rule of GROUP_RULES) {
    if (rule.keywords.some((kw) => lower.includes(kw))) return rule.group;
  }
  return "Lain-Lain";
}

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

function parseJumlahGl(text) {
  const m = String(text || "").match(/^jumlah\s*:\s*(\d{8})\b/i);
  return m ? m[1] : null;
}

function isNumericCell(value) {
  if (value === null || value === undefined || value === "") return false;
  if (typeof value === "number") return Number.isFinite(value);
  const s = String(value).replace(/,/g, "").trim();
  return s !== "" && Number.isFinite(Number(s));
}

/** Heuristik baris material (hindari judul/header ikut ke buffer). */
function isMaterialCandidate(material, colE, colF) {
  if (!material || isJumlahRow(material)) return false;
  if (/^gl\s*account$/i.test(material)) return false;
  if (/stas\d+/i.test(material) && !/^\d{6,}/.test(material)) return false;
  if (isNumericCell(colE) || isNumericCell(colF)) return true;
  if (/^\s*-/.test(material)) return true;
  if (/^\d{6,}/.test(material.trim())) return true;
  return false;
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
  return `${OUTPUT_PREFIX}_${day}_${month}_${year}_${hour}.${minute}_WIB`;
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
  const sheetName = worksheet.name;
  if (sheetName.trim().toLowerCase() === "rekap") return [];

  const rows = [];
  let currentGl = null;
  let pending = []; // materials sebelum header GL (pola Adolina)

  const pushMaterial = (gl, material, jumlah, biaya) => {
    rows.push({
      gl_account: gl,
      material,
      jumlah,
      biaya,
      group: assignGroup(material),
      pks: sheetName,
    });
  };

  const flushPending = (gl) => {
    if (TARGET_GLS.has(gl)) {
      for (const item of pending) {
        pushMaterial(gl, item.material, item.jumlah, item.biaya);
      }
    }
    pending = [];
  };

  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const colA = getCellValue(row, 1);
    const colB = getCellValue(row, 2);
    const colE = getCellValue(row, 5);
    const colF = getCellValue(row, 6);

    const textA = cellText(colA);
    const textB = cellText(colB);

    const glFromA = normalizeGl(colA);
    if (glFromA) {
      pending = [];
      currentGl = TARGET_GLS.has(glFromA) ? glFromA : null;
      return;
    }

    // Jumlah: 51100106 — tutup blok; jika material tanpa header GL, pakai GL dari Jumlah
    const jumlahGl = parseJumlahGl(textB) || parseJumlahGl(textA);
    if (jumlahGl) {
      if (pending.length) flushPending(jumlahGl);
      currentGl = null;
      return;
    }

    if (isJumlahRow(textB) || isJumlahRow(textA)) {
      pending = [];
      currentGl = null;
      return;
    }

    const material = textB;
    if (!material) return;
    if (!isMaterialCandidate(material, colE, colF)) {
      // header/station label → buang buffer yang nyangkut
      if (!currentGl) pending = [];
      return;
    }

    const jumlah = toNumber(colE);
    const biaya = toNumber(colF) * 1000;

    if (currentGl) {
      pushMaterial(currentGl, material, jumlah, biaya);
    } else {
      // buffer dulu; GL bisa baru ketahuan di baris Jumlah berikutnya
      pending.push({ material, jumlah, biaya });
    }
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
      map.set(row.pks, emptyGroupData());
    }
    const entry = map.get(row.pks)[row.group];
    entry.qty += row.jumlah;
    entry.biaya += row.biaya;
  }
  return map;
}

function emptyGroupData() {
  const init = {};
  for (const g of GROUPS) init[g] = { qty: 0, biaya: 0 };
  return init;
}

function normalizePksName(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/kebun\s*&?\s*/g, " ")
    .replace(/\bpks\b/g, " ")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function bigramCounts(s) {
  const m = new Map();
  for (let i = 0; i < s.length - 1; i++) {
    const bg = s.slice(i, i + 2);
    m.set(bg, (m.get(bg) || 0) + 1);
  }
  return m;
}

function diceCoefficient(a, b) {
  if (a === b) return 1;
  if (a.length < 2 || b.length < 2) return 0;
  const A = bigramCounts(a);
  const B = bigramCounts(b);
  let overlap = 0;
  for (const [bg, c] of A) {
    if (B.has(bg)) overlap += Math.min(c, B.get(bg));
  }
  return (2 * overlap) / (a.length - 1 + (b.length - 1));
}

function similarityScore(sheetName, masterName) {
  const a = normalizePksName(sheetName);
  const b = normalizePksName(masterName);
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.includes(b) || b.includes(a)) {
    const shorter = Math.min(a.length, b.length);
    const longer = Math.max(a.length, b.length);
    return 0.85 + 0.14 * (shorter / longer);
  }
  return diceCoefficient(a, b);
}

/** Greedy 1:1 match sheet names → master rows. */
function matchSheetsToMaster(sheetNames, masterList) {
  const candidates = [];
  for (const sheet of sheetNames) {
    for (let i = 0; i < masterList.length; i++) {
      const score = similarityScore(sheet, masterList[i].nama_pks);
      if (score >= SIMILARITY_THRESHOLD) {
        candidates.push({ sheet, masterIdx: i, score });
      }
    }
  }
  candidates.sort((a, b) => b.score - a.score);

  const usedSheets = new Set();
  const usedMaster = new Set();
  const sheetToMaster = new Map();

  for (const c of candidates) {
    if (usedSheets.has(c.sheet) || usedMaster.has(c.masterIdx)) continue;
    usedSheets.add(c.sheet);
    usedMaster.add(c.masterIdx);
    sheetToMaster.set(c.sheet, masterList[c.masterIdx]);
  }
  for (const sheet of sheetNames) {
    if (!sheetToMaster.has(sheet)) sheetToMaster.set(sheet, null);
  }
  return sheetToMaster;
}

async function loadMasterPks() {
  if (!fs.existsSync(MASTER_PKS_PATH)) {
    throw new Error(`Master PKS tidak ditemukan: ${MASTER_PKS_PATH}`);
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(MASTER_PKS_PATH);
  const sheet = workbook.worksheets[0];
  const list = [];
  sheet.eachRow({ includeEmpty: false }, (row, i) => {
    if (i === 1) return;
    const nama = cellText(row.getCell(1).value);
    const kode = cellText(row.getCell(2).value);
    const regional = toNumber(row.getCell(3).value);
    if (!nama) return;
    list.push({
      nama_pks: nama,
      kode_dfarm: kode,
      regional_id: regional,
    });
  });
  return list;
}

function regionalIdFromFilename(filename) {
  const arabic = String(filename).match(/regional\s*(\d+)/i);
  if (arabic) return Number(arabic[1]);
  const roman = String(filename).match(/regional\s*([ivxlcdm]+)/i);
  if (roman) {
    const n = ROMAN_REGIONAL[roman[1].toLowerCase()];
    return n != null ? n : null;
  }
  return null;
}

function buildDisplayRows(rows, masterList) {
  const sheetMap = buildSummaryMap(rows);
  const sheetNames = [...sheetMap.keys()];
  const sheetToMaster = matchSheetsToMaster(sheetNames, masterList);

  const byKode = new Map();
  for (const m of masterList) {
    byKode.set(m.kode_dfarm, {
      sheet_name: "",
      nama_pks: m.nama_pks,
      kode: m.kode_dfarm,
      regional_id: m.regional_id,
      groups: emptyGroupData(),
    });
  }

  const orphans = [];
  for (const [sheetName, groups] of sheetMap) {
    const master = sheetToMaster.get(sheetName);
    if (master && byKode.has(master.kode_dfarm)) {
      const entry = byKode.get(master.kode_dfarm);
      entry.sheet_name = entry.sheet_name
        ? `${entry.sheet_name}; ${sheetName}`
        : sheetName;
      for (const g of GROUPS) {
        entry.groups[g].qty += groups[g].qty;
        entry.groups[g].biaya += groups[g].biaya;
      }
    } else {
      orphans.push({
        sheet_name: sheetName,
        nama_pks: "",
        kode: "",
        regional_id: "",
        groups,
      });
    }
  }

  return [...masterList.map((m) => byKode.get(m.kode_dfarm)), ...orphans];
}

function applyZeroRed(qtyCell, biayaCell, qty) {
  if (qty === 0) {
    qtyCell.fill = RED_FILL;
    biayaCell.fill = RED_FILL;
  }
}

/**
 * Fill summary/REKAP layout.
 * B=Nama Sheet, C=Nama PKS, D=Kode PKS, E=Regional, F..=groups, last=Total
 */
function fillSummarySheet(sheet, rows, masterList) {
  const displayRows = buildDisplayRows(rows, masterList);
  const firstGroupCol = 6; // F
  const totalCol = firstGroupCol + GROUPS.length * 2;

  const headerStyle = { bold: true };
  const center = { horizontal: "center", vertical: "middle" };

  sheet.getCell("B3").value = "Nama Sheet";
  sheet.getCell("C3").value = "Nama PKS";
  sheet.getCell("D3").value = "Kode PKS";
  sheet.getCell("E3").value = "Regional";
  for (const col of [2, 3, 4, 5]) {
    sheet.getCell(3, col).font = headerStyle;
    sheet.getCell(3, col).alignment = center;
    sheet.mergeCells(3, col, 4, col);
  }

  GROUPS.forEach((group, i) => {
    const startCol = firstGroupCol + i * 2;
    const endCol = startCol + 1;
    const startCell = sheet.getCell(3, startCol);
    startCell.value = group;
    startCell.font = headerStyle;
    startCell.alignment = center;
    sheet.mergeCells(3, startCol, 3, endCol);

    const qtyCell = sheet.getCell(4, startCol);
    const biayaCell = sheet.getCell(4, endCol);
    qtyCell.value = "Qty";
    biayaCell.value = "Biaya";
    qtyCell.font = headerStyle;
    biayaCell.font = headerStyle;
    qtyCell.alignment = { horizontal: "center" };
    biayaCell.alignment = { horizontal: "center" };
  });

  const totalHeader = sheet.getCell(3, totalCol);
  totalHeader.value = "Total";
  totalHeader.font = headerStyle;
  totalHeader.alignment = center;
  sheet.mergeCells(3, totalCol, 3, totalCol + 1);
  sheet.getCell(4, totalCol).value = "Qty";
  sheet.getCell(4, totalCol + 1).value = "Biaya";
  sheet.getCell(4, totalCol).font = headerStyle;
  sheet.getCell(4, totalCol + 1).font = headerStyle;
  sheet.getCell(4, totalCol).alignment = { horizontal: "center" };
  sheet.getCell(4, totalCol + 1).alignment = { horizontal: "center" };

  sheet.getColumn(2).width = 28;
  sheet.getColumn(3).width = 28;
  sheet.getColumn(4).width = 12;
  sheet.getColumn(5).width = 10;
  GROUPS.forEach((_, i) => {
    sheet.getColumn(firstGroupCol + i * 2).width = 12;
    sheet.getColumn(firstGroupCol + i * 2 + 1).width = 16;
  });
  sheet.getColumn(totalCol).width = 12;
  sheet.getColumn(totalCol + 1).width = 16;

  displayRows.forEach((row, idx) => {
    const rowNum = 5 + idx;
    sheet.getCell(rowNum, 2).value = row.sheet_name;
    sheet.getCell(rowNum, 3).value = row.nama_pks;
    sheet.getCell(rowNum, 4).value = row.kode;
    sheet.getCell(rowNum, 5).value = row.regional_id;

    let totalQty = 0;
    let totalBiaya = 0;

    GROUPS.forEach((group, i) => {
      const qtyCol = firstGroupCol + i * 2;
      const biayaCol = qtyCol + 1;
      const qtyCell = sheet.getCell(rowNum, qtyCol);
      const biayaCell = sheet.getCell(rowNum, biayaCol);
      const qty = row.groups[group].qty;
      const biaya = row.groups[group].biaya;
      qtyCell.value = qty;
      biayaCell.value = biaya;
      qtyCell.numFmt = "#,##0.###";
      biayaCell.numFmt = "#,##0";
      applyZeroRed(qtyCell, biayaCell, qty);
      totalQty += qty;
      totalBiaya += biaya;
    });

    const tQty = sheet.getCell(rowNum, totalCol);
    const tBiaya = sheet.getCell(rowNum, totalCol + 1);
    tQty.value = totalQty;
    tBiaya.value = totalBiaya;
    tQty.numFmt = "#,##0.###";
    tBiaya.numFmt = "#,##0";
    applyZeroRed(tQty, tBiaya, totalQty);
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

async function writeCombinedOutput(allRows, baseName, masterList) {
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
  fillSummarySheet(summary, allRows, masterList);

  const outPath = path.join(RESULTS_DIR, `${baseName}.xlsx`);
  await workbook.xlsx.writeFile(outPath);
  return outPath;
}

async function writePerSourceOutputs(bySource, baseName, masterList) {
  const folder = path.join(RESULTS_DIR, baseName);
  fs.mkdirSync(folder, { recursive: true });

  const written = [];
  for (const { sourcePath, sourceName, rows } of bySource) {
    const destPath = path.join(folder, sourceName);
    fs.copyFileSync(sourcePath, destPath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(destPath);

    const regionalId = regionalIdFromFilename(sourceName);
    const masterForSource =
      regionalId != null
        ? masterList.filter((m) => Number(m.regional_id) === regionalId)
        : masterList;

    const rekap = getOrReplaceRekapSheet(workbook);
    fillSummarySheet(rekap, rows, masterForSource);

    await workbook.xlsx.writeFile(destPath);
    written.push(destPath);
  }
  return { folder, written };
}

async function main() {
  const isRailTrack = process.argv.includes("--rail-track");
  configureMode(isRailTrack);

  if (!fs.existsSync(RESULTS_DIR)) {
    fs.mkdirSync(RESULTS_DIR, { recursive: true });
  }

  console.log(`Mode: ${MODE_NAME}`);
  console.log(`Sumber data: ${DATA_DIR}`);
  console.log(`Groups: ${GROUPS.join(", ")}`);

  console.log("Membaca master_pks.xlsx ...");
  const masterList = await loadMasterPks();
  console.log(`Master PKS: ${masterList.length} baris.`);

  console.log("Membaca workbook ...");
  const bySource = await extractAllBySource();
  const allRows = bySource.flatMap((s) => s.rows);
  console.log(
    `Ditemukan ${allRows.length} baris material dari ${bySource.length} workbook.`
  );

  const baseName = wibTimestampBase();
  const combinedPath = await writeCombinedOutput(allRows, baseName, masterList);
  console.log(`Gabungan: ${combinedPath}`);

  const { folder, written } = await writePerSourceOutputs(
    bySource,
    baseName,
    masterList
  );
  console.log(`Per source (${written.length} file) di: ${folder}`);
  for (const p of written) console.log(`  - ${path.basename(p)}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
