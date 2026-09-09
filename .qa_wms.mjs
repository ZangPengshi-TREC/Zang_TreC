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
const detailIdx = idx("詳細説明");
const answerIdx = idx("回答");
const residualIdx = idx("残課題");
const rows = values.slice(headerIndex + 1).map((r, i) => ({
  row: headerIndex + i + 2,
  no: r[1],
  area: r[4],
  system: String(r[systemIdx] ?? ""),
  summary: String(r[summaryIdx] ?? ""),
  detail: String(r[detailIdx] ?? ""),
  answer: String(r[answerIdx] ?? ""),
  residual: String(r[residualIdx] ?? ""),
  status: String(r[statusIdx] ?? ""),
})).filter((x) => x.system === "WMS");

const counts = {};
for (const x of rows) counts[x.status || "(blank)"] = (counts[x.status || "(blank)"] ?? 0) + 1;
console.log("WMS_COUNT=" + rows.length);
console.log("WMS_STATUS=" + JSON.stringify(counts));
console.log(JSON.stringify(rows, null, 2));
