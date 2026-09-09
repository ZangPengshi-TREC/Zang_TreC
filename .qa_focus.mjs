import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("/Users/treqd/Downloads/情報システムPMI_MD基幹統合PJ_TRE_課題管理表.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("TRE_課題管理表 MDSCM領域");
const values = sheet.getRange("A1:P224").values;
const headerIndex = values.findIndex((r) => r.some((v) => String(v ?? "").includes("課題・質問概要")));
const headers = values[headerIndex];
const idx = (name) => headers.findIndex((h) => String(h ?? "").includes(name));
const systemIdx = idx("システム");
const statusIdx = idx("ステータス");
const summaryIdx = idx("課題・質問概要");
const answerIdx = idx("回答");
const residualIdx = idx("残課題");
const answerDateIdx = idx("回答日");
const rows = values.slice(headerIndex + 1).map((r, i) => ({
  row: headerIndex + i + 2,
  no: r[1],
  system: String(r[systemIdx] ?? ""),
  summary: String(r[summaryIdx] ?? ""),
  answer: String(r[answerIdx] ?? ""),
  residual: String(r[residualIdx] ?? ""),
  status: String(r[statusIdx] ?? ""),
  answerDate: r[answerDateIdx],
})).filter((x) => x.system.includes("自動補充"));

const open = rows.filter((x) => x.status !== "完了" && x.status !== "中止" && x.summary);
const residual = rows.filter((x) => x.residual.trim());
const byStatus = {};
for (const x of rows) byStatus[x.status || "(blank)"] = (byStatus[x.status || "(blank)"] ?? 0) + 1;
console.log("AUTO_STATUS=" + JSON.stringify(byStatus));
console.log("OPEN=" + JSON.stringify(open, null, 2));
console.log("RESIDUAL=" + JSON.stringify(residual, null, 2));
