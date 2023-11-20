import ExcelJS from "exceljs";

/**
 * 指定された範囲の行をソースシートからターゲットシートへコピーします。
 * この関数は、ExcelJSライブラリを使用して、特定の範囲の行をコピーし、
 * 指定された開始位置にそれらの行を挿入します。
 *
 * @param {Worksheet} targetSheet コピー先のワークシートオブジェクト
 * @param {Worksheet} sourceSheet コピー元のワークシートオブジェクト
 * @param {number} from コピーを開始するソースシートの行番号（開始行）
 * @param {number} to コピーを終了するソースシートの行番号（終了行）
 * @param {number} targetStart コピー先のワークシートでの挿入開始行番号
 */
// export async function copyRows(
export const copyRows = async (
  targetSheet: ExcelJS.Worksheet,
  sourceSheet: ExcelJS.Worksheet,
  from: number,
  to: number,
  targetStart: number
) => {
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
};
