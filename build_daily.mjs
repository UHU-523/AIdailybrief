import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const FONT = "Arial";
const HEADER = ["时间戳", "今日要点", "热门开源项目", "值得关注的讨论/观点", "来源链接", "简要说明"];
const outDir = path.join(process.cwd(), "output");
const targetFile = process.env.DAILY_XLSX || path.join(process.env.USERPROFILE || "", "Desktop", "AI资讯每日汇总.xlsx");

const inputPath = process.argv[2] || path.join(process.cwd(), "input.json");
const input = JSON.parse(await fs.readFile(inputPath, "utf8"));

const row = [
  input.date,
  (input.highlights || []).join("\n"),
  (input.repos || []).join("\n"),
  (input.discussions || []).join("\n"),
  (input.links || []).join("\n"),
  input.brief || "",
];

let wb;
let sheet;
try {
  wb = await SpreadsheetFile.importXlsx(await FileBlob.load(targetFile));
  sheet = wb.worksheets.getItem("AI资讯汇总");
} catch {
  wb = Workbook.create();
  sheet = wb.worksheets.add("AI资讯汇总");
  sheet.getRange("A1:F1").values = [HEADER];
  sheet.getRange("A1:F1").format = {
    fill: "#1F4E78",
    font: { name: FONT, size: 10, bold: true, color: "#FFFFFF" },
    borders: { insideHorizontal: { style: "thin", color: "#FFFFFF" }, insideVertical: { style: "thin", color: "#FFFFFF" } },
  };
  sheet.getRange("A1:F1").format.horizontalAlignment = "center";
  sheet.getRange("A1:F1").format.verticalAlignment = "center";
  sheet.showGridLines = false;
}

const colA = sheet.getRange("A1:A200").values;
let lastRow = 0;
let todayRow = 0;
for (let i = 0; i < colA.length; i++) {
  const v = colA[i] && colA[i][0];
  if (v === null || v === undefined || String(v).trim() === "") continue;
  lastRow = i + 1;
  if (String(v).trim() === input.date) todayRow = i + 1;
}
const targetRow = todayRow || lastRow + 1;

const rng = sheet.getRange(`A${targetRow}:F${targetRow}`);
rng.values = [row];
rng.format.font = { name: FONT, size: 10, color: "#1F1F1F" };
rng.format.verticalAlignment = "top";
rng.format.wrapText = true;
rng.format.rowHeight = 220;
sheet.getRange(`A${targetRow}`).format.horizontalAlignment = "center";

sheet.getRange("A:A").format.columnWidth = 12;
sheet.getRange("B:B").format.columnWidth = 48;
sheet.getRange("C:C").format.columnWidth = 34;
sheet.getRange("D:D").format.columnWidth = 44;
sheet.getRange("E:E").format.columnWidth = 55;
sheet.getRange("F:F").format.columnWidth = 46;
sheet.freezePanes.freezeRows(1);

wb.recalculate();
await fs.mkdir(outDir, { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(path.join(outDir, path.basename(targetFile)));
console.log("done: row=" + targetRow + " -> " + path.join(outDir, path.basename(targetFile)));
