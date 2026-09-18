import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = "C:/Users/congn/Downloads/employ_phat_giuso.txt";
const outputDir = "C:/Users/congn/Pictures/vnptemploy/outputs/giuso_20260917";
const outputPath = `${outputDir}/employ_phat_giuso.xlsx`;
const previewPath = `${outputDir}/employ_phat_giuso_preview.png`;

const dateKeys = new Set(["ngaytao", "thoigian_giao", "ngay_khoitao"]);
const numericKeys = new Set([
  "hdkh_id", "kenh_toantrinh", "id_loai_sim", "flag_thm", "id_donhang",
  "id_trangthai", "phanvung_id", "phangiao", "id_ungdung", "tien",
  "id_loaidon", "hdtb_id", "trang_thai_khoitao", "loaitb_id", "chitieu_tg",
  "id_trangthai_thanhtoan"
]);

function parseVietnameseDate(value) {
  const match = String(value || "").trim().match(/^(\d{2})\/(\d{2})\/(\d{4})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return value;
  const [, day, month, year, hour, minute, second = "0"] = match;
  return new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
}

function typedValue(key, value) {
  if (value === null || value === undefined || value === "") return null;
  if (dateKeys.has(key)) return parseVietnameseDate(value);
  if (numericKeys.has(key) && typeof value === "number" && Number.isFinite(value)) return value;
  return String(value);
}

function columnLetter(index) {
  let n = index + 1;
  let result = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

const sourceText = await fs.readFile(inputPath, "utf8");
const payload = JSON.parse(sourceText);
const rows = Array.isArray(payload?.data) ? payload.data : [];
if (!rows.length) throw new Error("File không có dữ liệu trong trường data");

// Giữ đúng thứ tự trường của bản ghi đầu tiên để không làm mất cột khi xuất bảng.
const columns = Object.keys(rows[0]);
const matrix = [
  columns,
  ...rows.map(row => columns.map(key => typedValue(key, row?.[key])))
];

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Giữ số");
sheet.showGridLines = false;

const endColumn = columnLetter(columns.length - 1);
const endRow = matrix.length;
const usedRange = sheet.getRange(`A1:${endColumn}${endRow}`);
usedRange.values = matrix;

const header = sheet.getRange(`A1:${endColumn}1`);
header.format = {
  fill: "#1F4E78",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
header.format.rowHeight = 30;

const body = sheet.getRange(`A2:${endColumn}${endRow}`);
body.format = {
  font: { name: "Arial", size: 10, color: "#1F2937" },
  verticalAlignment: "center",
};

const dateColumnIndexes = columns
  .map((key, index) => dateKeys.has(key) ? index : -1)
  .filter(index => index >= 0);
for (const index of dateColumnIndexes) {
  sheet.getRangeByIndexes(1, index, rows.length, 1).format.numberFormat = "dd/mm/yyyy hh:mm:ss";
}
for (const key of ["tien"]) {
  const index = columns.indexOf(key);
  if (index >= 0) sheet.getRangeByIndexes(1, index, rows.length, 1).format.numberFormat = "#,##0";
}

// Autofit trước, sau đó giới hạn các cột văn bản dài để bảng vẫn dễ cuộn.
usedRange.format.autofitColumns();
const widthCaps = {
  ghichu: 28,
  diachi: 42,
  hinhthuc_dktttb: 24,
  ten_kh: 22,
  ma_gd: 20,
  ma_gt: 20,
};
for (const [key, width] of Object.entries(widthCaps)) {
  const index = columns.indexOf(key);
  if (index >= 0) sheet.getRangeByIndexes(0, index, endRow, 1).format.columnWidth = width;
}

sheet.freezePanes.freezeRows(1);
const table = sheet.tables.add(`A1:${endColumn}${endRow}`, true, "GiuSoTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;

workbook.recalculate();
const inspection = await workbook.inspect({
  kind: "table",
  range: `Giữ số!A1:${endColumn}6`,
  include: "values,formulas",
  tableMaxRows: 6,
  tableMaxCols: 12,
  tableMaxCellChars: 80,
});
console.log(inspection.ndjson);

const preview = await workbook.render({ sheetName: "Giữ số", range: `A1:${endColumn}8`, scale: 1, format: "png" });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);

console.log(JSON.stringify({ outputPath, previewPath, rowCount: rows.length, columnCount: columns.length }));
