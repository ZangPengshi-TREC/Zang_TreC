import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("/Users/treqd/Downloads/情報システムPMI_MD基幹統合PJ_TRE_課題管理表.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("TRE_課題管理表 MDSCM領域");
const values = sheet.getRange("A1:P224").values;
const headerIndex = values.findIndex((r) => r.some((v) => String(v ?? "").includes("課題・質問概要")));
console.log(`HEADER_INDEX=${headerIndex}`);
console.log(`HEADER=${JSON.stringify(values[headerIndex])}`);
const headers = values[headerIndex];
const idx = (name) => headers.findIndex((h) => String(h ?? "").includes(name));
const systemIdx = idx("システム");
const statusIdx = idx("ステータス");
const summaryIdx = idx("課題・質問概要");
const answerIdx = idx("回答");
const residualIdx = idx("残課題");
console.log(JSON.stringify({ systemIdx, statusIdx, summaryIdx, answerIdx, residualIdx }));
const targets = values.slice(headerIndex + 1).map((r, i) => ({ row: headerIndex + i + 2, r }))
  .filter(({ r }) => String(r[systemIdx] ?? "").includes("自動補充") || String(r[summaryIdx] ?? "").includes("自動補充") || String(r[summaryIdx] ?? "").includes("Master") || String(r[summaryIdx] ?? "").includes("マスタ"));
for (const { row, r } of targets) {
  const out = {
    row,
    no: r[1],
    area: r[4],
    system: r[systemIdx],
    summary: r[summaryIdx],
    answer: r[answerIdx],
    residual: r[residualIdx],
    status: r[statusIdx],
  };
  console.log(JSON.stringify(out));
}

const counts = {};
for (const r of values.slice(headerIndex + 1)) {
  const s = String(r[statusIdx] ?? "").trim();
  if (s) counts[s] = (counts[s] ?? 0) + 1;
}
console.log("STATUS_COUNTS=" + JSON.stringify(counts));
