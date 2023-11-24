import ExcelJS, { Anchor } from "exceljs";
import sharp from "sharp";

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
      const width =
        sourceSheet.getColumn(colNumber).width ||
        sourceSheet.properties.defaultColWidth;
      targetSheet.getColumn(colNumber).width = width;

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

export const pasteImageWithAspectRatio = (
  worksheet: ExcelJS.Worksheet,
  cell: ExcelJS.Cell,
  /** Workbook.addImage を呼ぶと返される */
  imageId: number,
  /** 画像を貼る範囲のカラム数 */
  colCount: number,
  /** 画像を貼る範囲の行数 */
  rowCount: number,
  // columnWidthP: number,
  // rowHeightP: number,
  imageWidthP: number,
  imageHeightP: number,
  colTl: number,
  rowTl: number
) => {
  const [columnWidthP, rowHeightP] = cellWidthHeightInPixel(cell);
  const [offC, offR] = [
    (columnWidthP - imageWidthP) / 2,
    (rowHeightP - imageHeightP) / 2,
  ];

  worksheet.addImage(imageId, {
    tl: {
      nativeCol: colTl - 1,
      nativeRow: rowTl - 1,
      nativeColOff: Math.trunc(offC * 9525), // 1pixel = 9525EMU
      nativeRowOff: Math.trunc(offR * 9525),
    } as Anchor,
    // brを指定することで必ずセル内に画像が収まるようにする。但しアスペクト比が崩れる場合がある。
    br: {
      nativeCol: colTl - 1 + colCount,
      nativeRow: rowTl - 1 + rowCount,
      nativeColOff: -Math.trunc(offC * 9525), // 1pixel = 9525EMU
      nativeRowOff: -Math.trunc(offR * 9525),
    } as Anchor,
  });
};

/**
 * get the width and height in pixel for specifi cell considering merged cell.
 */
export const cellWidthHeightInPixel = (
  cell: ExcelJS.Cell
): [number, number] => {
  const sheet = cell.worksheet;
  let [width, height] = [0, 0];

  for (let c = cell.fullAddress.col; ; c++) {
    const currentCell = cell.worksheet.getCell(cell.fullAddress.row, c);
    // マージの有無に関わらず、現在のセルの左上（master）が対象のセルである限り値を加算し続ける
    if (currentCell.master != cell) break;

    // 初期値から列の横幅を変えてない場合に、値が取れない可能性があるたワークシート側のプロパティも参照
    width += sheet.getColumn(c).width || sheet.properties.defaultColWidth || 0;
  }

  for (let r = cell.fullAddress.row; ; r++) {
    const currentCell = sheet.getCell(r, cell.col);

    // マージの有無に関わらず、現在のセルの左上（master）が対象のセルである限り値を加算し続ける
    if (currentCell.master != cell) break;

    // 初期値から列の横幅を変えてない場合に、値が取れない可能性があるたワークシート側のプロパティも参照
    height += sheet.getRow(r).height || sheet.properties.defaultRowHeight || 0;
  }
  return [columnWidthInPixel(width), rowHeightInPixel(height)];
};

/**
 * get the column width in pixel.
 */
const columnWidthInPixel = (width: number, fontWidth = 8) => {
  return ((256 * width + 128 / fontWidth) / 256) * fontWidth;
};

/**
 * get the row height in pixel.
 */
const rowHeightInPixel = (height: number) => {
  //  Pixels DPI (96 pixels per inch), Points DPI (72 points per inch)
  return (height * 96) / 72;
};

export const fetchImageAsBuffer = async (
  imageUrl: string
): Promise<Buffer | null> => {
  const response = await fetch(imageUrl);
  if (!response.ok) {
    return null;
  }
  const arrayBuffer = await response.arrayBuffer();
  return Buffer.from(arrayBuffer);
};

export const resizeImage = async (
  buffer: Buffer,
  width: number,
  height: number
) => {
  return await sharp(buffer)
    .resize(Math.trunc(width), Math.trunc(height), { fit: "inside" })
    .jpeg({ mozjpeg: true })
    .toBuffer();
};

// 改ページを挿入する(印刷する時に期待しないところでページが切り替わってしまうのを防ぐために必要)
export const addPageBreak = (
  sheet: ExcelJS.Worksheet,
  rowNum: number
): void => {
  sheet.getRow(rowNum).addPageBreak();
};
