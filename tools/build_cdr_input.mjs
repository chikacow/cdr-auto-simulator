import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = "/Users/kiencc/Downloads/cdr die.xlsx";
const outputDir = "/Users/kiencc/Downloads/outputs/cdr-simulator-input";
const outputPath = `${outputDir}/cdr-simulator-input.xlsx`;

const sourceFile = await FileBlob.load(inputPath);
const sourceWorkbook = await SpreadsheetFile.importXlsx(sourceFile);
const sourceSheet = sourceWorkbook.worksheets.getItem("retest mercury");
if (!sourceSheet) throw new Error("Không tìm thấy sheet retest mercury.");

const sourceValues = sourceSheet.getRange("C1:E130").values;
const cdrRows = sourceValues
  .map((row) => [String(row[0] ?? "").trim(), String(row[2] ?? "").trim()])
  .filter(([usage, curl]) => usage !== "" && curl !== "");

if (cdrRows.length === 0) throw new Error("Sheet retest mercury không có dòng Usage/Curl hợp lệ.");

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("CDR Input");
sheet.showGridLines = false;
sheet.getRangeByIndexes(0, 0, cdrRows.length + 1, 2).values = [["Usage", "Curl"], ...cdrRows];

const header = sheet.getRange("A1:B1");
header.format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: "#1F4E78" },
};
header.format.rowHeight = 22;

const body = sheet.getRangeByIndexes(1, 0, cdrRows.length, 2);
body.format = {
  verticalAlignment: "top",
  wrapText: false,
  borders: { preset: "outside", style: "thin", color: "#D9E2F3" },
};
sheet.getRange(`A2:A${cdrRows.length + 1}`).format.columnWidth = 42;
sheet.getRange(`B2:B${cdrRows.length + 1}`).format.columnWidth = 110;
sheet.freezePanes.freezeRows(1);

workbook.recalculate();
const check = await workbook.inspect({
  kind: "table",
  range: `CDR Input!A1:B${Math.min(cdrRows.length + 1, 8)}`,
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 2,
});
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 20 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({ sheetName: "CDR Input", range: "A1:B8", scale: 1.5, format: "png" });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(`${outputDir}/cdr-simulator-input-preview.png`, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, rowCount: cdrRows.length }));
