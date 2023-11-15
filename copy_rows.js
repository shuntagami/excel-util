const ExcelJS = require("exceljs");
const fs = require("fs").promises;

async function copyRows(targetSheet, sourceSheet, from, to) {
  try {
    for (let i = from; i <= to; i++) {
      const sourceRow = sourceSheet.getRow(i);
      const targetRow = targetSheet.getRow(i);

      targetRow.height = sourceRow.height;
      sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const targetCell = targetRow.getCell(colNumber);

        targetCell.style = cell.style;
        targetCell.value = cell.value;

        // merge cell
        const range = `${cell.model.master || cell.address}:${
          targetCell.address
        }`;
        targetSheet.unMergeCells(range);
        targetSheet.mergeCells(range);
      });
    }

    const buffer = await targetSheet.workbook.xlsx.writeBuffer();
    const path = `./results/${Date.now()}.xlsx`;
    await fs.writeFile(path, buffer);

    console.log(`File written to ${path}`);
  } catch (error) {
    console.error("Error:", error);
  }
}

async function main() {
  const startRow = parseInt(process.argv[2]);
  const endRow = parseInt(process.argv[3]);

  const template = await new ExcelJS.Workbook().xlsx.readFile(
    "./templates/template.xlsx"
  );
  const workbook = new ExcelJS.Workbook();
  const targetSheet = workbook.addWorksheet("Target Sheet");
  const sourceSheet = template.worksheets[0];

  await copyRows(targetSheet, sourceSheet, startRow, endRow);
}

main().catch((error) => console.error(error));
