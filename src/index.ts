import ExcelJS from "exceljs";
import { exit } from "process";
import { copyRows } from "./utils/excel_util";

const fs = require("fs").promises;

async function main() {
  const startRow = parseInt(process.argv[2] as string);
  const endRow = parseInt(process.argv[3] as string);
  const targetStartRow = parseInt(process.argv[4] as string);
  const template = await new ExcelJS.Workbook().xlsx.readFile(
    "./templates/template.xlsx"
  );
  const workbook = new ExcelJS.Workbook();
  const targetSheet = workbook.addWorksheet("Target Sheet");
  const sourceSheet = template.worksheets[0];
  if (sourceSheet === undefined) {
    exit(1);
  }

  const buffer = await copyRows(
    targetSheet,
    sourceSheet,
    startRow,
    endRow,
    targetStartRow
  );

  const path = `./results/${Date.now()}.xlsx`;
  await fs.writeFile(path, buffer);
  console.log(`File written to ${path}`);
}

main().catch((error) => console.error(error));
