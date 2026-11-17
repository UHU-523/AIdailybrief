import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const FONT = "Arial";
const outDir = path.join(process.cwd(), "output");

// 按需修改为你的技能目录
const home = process.env.USERPROFILE || process.env.HOME || "";
const USER_ROOTS = [
  path.join(home, ".codex", "skills"),
  path.join(home, ".agents", "skills"),
];
const BUNDLED_ROOTS = [
  path.join(home, ".codex", "plugins", "cache", "openai-bundled"),
  path.join(home, ".codex", "plugins", "cache", "openai-primary-runtime"),
];
const BUNDLED_WHITELIST = new Set([
  "computer-use", "visualize", "documents", "pdf", "presentations",
  "spreadsheets", "excel-live-control", "template-creator",
]);
const BLACKLIST = new Set(["nature-shared"]);

function parseFrontmatter(text) {
  const m = text.match(/^---\s*\n([\s\S]*?)\n---/);
  if (!m) return { name: "", description: "" };
  const fm = m[1];
  const name = (fm.match(/^name:\s*["']?([^"'\n]+)/m) || [])[1]?.trim() || "";
  let description = (fm.match(/^description:\s*(.*)$/m) || [])[1]?.trim() || "";
  description = description.replace(/^["']|["']$/g, "");
  if (description.length > 220) description = description.slice(0, 220) + "…";
  return { name, description };
}

async function readSkill(dir) {
  let text;
  try {
    text = await fs.readFile(path.join(dir, "SKILL.md"), "utf8");
  } catch {
    return null;
  }
  const { name, description } = parseFrontmatter(text);
  return { name: name || path.basename(dir), description };
}

async function findSkillMds(dir, out) {
  let entries;
  try {
    entries = await fs.readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const e of entries) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) await findSkillMds(p, out);
    else if (e.isFile() && e.name.toLowerCase() === "skill.md") out.push(p);
  }
}

const seen = new Map();
function add(name, description) {
  if (!name || BLACKLIST.has(name) || seen.has(name)) return;
  seen.set(name, description || "");
}

for (const root of USER_ROOTS) {
  let entries;
  try {
    entries = await fs.readdir(root, { withFileTypes: true });
  } catch {
    continue;
  }
  for (const e of entries) {
    if (!e.isDirectory()) continue;
    if (e.name === ".system") {
      let sub;
      try {
        sub = await fs.readdir(path.join(root, e.name), { withFileTypes: true });
      } catch {
        continue;
      }
      for (const s of sub) {
        if (!s.isDirectory()) continue;
        const info = await readSkill(path.join(root, e.name, s.name));
        if (info) add(info.name, info.description);
      }
    } else {
      const info = await readSkill(path.join(root, e.name));
      if (info) add(info.name, info.description);
    }
  }
}

for (const root of BUNDLED_ROOTS) {
  const files = [];
  await findSkillMds(root, files);
  for (const f of files) {
    let text;
    try {
      text = await fs.readFile(f, "utf8");
    } catch {
      continue;
    }
    const { name, description } = parseFrontmatter(text);
    const finalName = name || path.basename(path.dirname(f));
    if ([...BUNDLED_WHITELIST].some((w) => finalName.toLowerCase() === w.toLowerCase())) {
      add(finalName, description);
    }
  }
}

const rows = [...seen.entries()].sort((a, b) => a[0].localeCompare(b[0], "en"));
const n = rows.length;
const lastRow = 4 + n;

const wb = Workbook.create();
const sheet = wb.worksheets.add("技能清单");
sheet.showGridLines = false;
sheet.getRange("A1").values = [["Codex 技能与功能清单"]];
sheet.getRange("A2").values = [["共 " + n + " 个可用技能"]];
sheet.getRange("A4:C4").values = [["序号", "技能名称", "功能说明"]];
sheet.getRange("A5:C" + lastRow).values = rows.map(([name, desc], i) => [i + 1, name, desc]);

sheet.getRange("A1").format.font = { name: FONT, size: 14, bold: true, color: "#1F1F1F" };
sheet.getRange("A2").format.font = { name: FONT, size: 10, italic: true, color: "#6B7280" };
sheet.getRange("A4:C4").format = {
  fill: "#1F4E78",
  font: { name: FONT, size: 10, bold: true, color: "#FFFFFF" },
  borders: { insideHorizontal: { style: "thin", color: "#FFFFFF" }, insideVertical: { style: "thin", color: "#FFFFFF" } },
};
sheet.getRange("A4:C4").format.horizontalAlignment = "center";
sheet.getRange("A4:C4").format.verticalAlignment = "center";

const body = sheet.getRange("A5:C" + lastRow);
body.format.font = { name: FONT, size: 10, color: "#1F1F1F" };
body.format.verticalAlignment = "center";
body.format.borders = { insideHorizontal: { style: "thin", color: "#E2E8F0" } };
sheet.getRange("A5:A" + lastRow).format.horizontalAlignment = "center";
sheet.getRange("A5:A" + lastRow).format.numberFormat = "0";
sheet.getRange("B5:B" + lastRow).format.font = { name: FONT, size: 10, bold: true, color: "#1F1F1F" };
sheet.getRange("C5:C" + lastRow).format.wrapText = true;
sheet.getRange("C5:C" + lastRow).format.horizontalAlignment = "left";

sheet.getRange("A:A").format.columnWidth = 8;
sheet.getRange("B:B").format.columnWidth = 30;
sheet.getRange("C:C").format.columnWidth = 72;
sheet.freezePanes.freezeRows(4);

wb.recalculate();
await fs.mkdir(outDir, { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(path.join(outDir, "Codex技能与功能清单.xlsx"));
console.log("skills=" + n + " -> " + path.join(outDir, "Codex技能与功能清单.xlsx"));
