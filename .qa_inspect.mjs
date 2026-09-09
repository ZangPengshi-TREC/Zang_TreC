import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "/Users/treqd/Downloads/情報システムPMI_MD基幹統合PJ_TRE_課題管理表.xlsx";
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 12000,
  tableMaxRows: 8,
  tableMaxCols: 12,
  tableMaxCellChars: 120,
});
console.log("SUMMARY\n" + summary.ndjson);

const sheets = workbook.worksheets.items;
for (const sheet of sheets) {
  const used = sheet.getUsedRange(true);
  console.log(`SHEET ${sheet.name}`);
  console.log(`USED ${used ? used.address : "none"}`);
  if (used) {
    const region = await workbook.inspect({
      kind: "region",
      sheetId: sheet.name,
      range: used.address,
      maxChars: 30000,
      tableMaxRows: 80,
      tableMaxCols: 30,
      tableMaxCellChars: 300,
    });
    console.log(region.ndjson);
  }
}
