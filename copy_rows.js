const ExcelJS = require("exceljs");
const fs = require("fs").promises;

async function copyRows(targetSheet, sourceSheet, from, to, targetStart) {
  try {
    let targetRowIndex = targetStart;
    for (let i = from; i <= to; i++) {
      const sourceRow = sourceSheet.getRow(i);
      const targetRow = targetSheet.getRow(targetRowIndex + i);

      targetRow.height = sourceRow.height;
      sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const targetCell = targetRow.getCell(colNumber);
        if (i === from) {
          const width = sourceSheet.getColumn(colNumber).width;
          targetSheet.getColumn(colNumber).width = width;
        }

        targetCell.style = cell.style;
        targetCell.value = cell.value;

        // merge cell
        if (cell.isMerged) {
          if (cell.master != cell) return;

          let [colCount, rowCount] = [0, 0];

          for (let c = cell.fullAddress.col; ; c++) {
            const currentCell = cell.worksheet.getCell(cell.row, c);
            if (currentCell.master != cell) break;

            colCount++;
          }

          for (let r = cell.fullAddress.row; ; r++) {
            const currentCell = cell.worksheet.getCell(r, cell.col);
            if (currentCell.master != cell) break;

            rowCount++;
          }

          targetSheet.mergeCells(
            targetRowIndex + i,
            colNumber,
            targetRowIndex + i + rowCount - 1,
            colNumber + colCount - 1
          );
        }
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
  const targetStartRow = parseInt(process.argv[4]);
  const template = await new ExcelJS.Workbook().xlsx.readFile(
    "./templates/template.xlsx"
  );
  const workbook = new ExcelJS.Workbook();
  const targetSheet = workbook.addWorksheet("Target Sheet");
  const sourceSheet = template.worksheets[0];

  await copyRows(targetSheet, sourceSheet, startRow, endRow, targetStartRow);
}

main().catch((error) => console.error(error));
