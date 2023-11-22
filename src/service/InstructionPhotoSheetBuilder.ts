import ExcelJS from "exceljs";
import { Blueprint, Sheet } from "./InstructionSheetBuilder";
import {
  cellWidthHeightInPixel,
  copyRows,
  fetchImageAsBuffer,
  pasteImageWithAspectRatio,
  resizeImage,
} from "../utils/excel_util";

export class InstructionPhotoSheetBuilder {
  static readonly VERTICAL_PHOTO_COUNT = 2;
  static readonly HORIZONTAL_PHOTO_COUNT = 3;
  static readonly TOTAL_PHOTO_COUNT_PER_TEMPLATE =
    InstructionPhotoSheetBuilder.VERTICAL_PHOTO_COUNT *
    InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT;
  static readonly COLUMN_COUNT_FOR_PHOTO_CELL = 7;
  static readonly ROW_COUNT_FOR_PHOTO_CELL = 12;

  constructor(
    public readonly workbook: ExcelJS.Workbook,
    public readonly workSheet: ExcelJS.Worksheet,
    private readonly templateSheet: ExcelJS.Worksheet,
    private readonly resources: Blueprint[]
  ) {
    this.workSheet.pageSetup = this.templateSheet.pageSetup;
  }

  async build(
    rowNum: number,
    marginWidth = 100,
    marginHeight = 100
  ): Promise<this> {
    let currentRowNum = rowNum;

    let photoIndex = 0;
    for (const blueprint of this.resources) {
      for (const sheet of blueprint.sheets) {
        // ここでテンプレートの切り替えを行う、シート単位なので
        // 案件名とかも埋める
        for (const instruction of sheet.instructions) {
          for (const photo of instruction.photos) {
            if (
              photoIndex %
                InstructionPhotoSheetBuilder.TOTAL_PHOTO_COUNT_PER_TEMPLATE ===
              0
            ) {
              copyRows(
                this.workSheet,
                this.templateSheet,
                1,
                32,
                currentRowNum - 1
              );
              currentRowNum += 1; // テンプレートの2行目がスタート位置
              this.fillBlueprintContents(currentRowNum, blueprint, sheet);
              currentRowNum += 2; // ヘッダー分2行追加
            }

            // photoを3回はったらcurrentRowNumをインクリメント
            // 6回目でテンプレートの延長が必要
            // photosのカウントが行ったらテンプレートの延長が必要
            this.fillInstructionContents(
              currentRowNum,
              instruction.displayId,
              photoIndex
            );
            await this.pasteInstructionPhoto(
              currentRowNum + 1,
              photo.url,
              marginWidth,
              marginHeight,
              photoIndex
            );
            photoIndex++;
            if (
              photoIndex %
                InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT ===
              0
            ) {
              currentRowNum += 14;
            }
          }
        }
      }
    }
    return this;
  }

  private fillBlueprintContents(
    rowNum: number,
    blueprint: Blueprint,
    sheet: Sheet
  ) {
    const row = this.workSheet.getRow(rowNum);
    row.getCell("A").value = blueprint.orderName;
    row.getCell("I").value = sheet.operationCategory;
    row.getCell("Q").value = blueprint.blueprintName;
  }

  private fillInstructionContents(
    rowNum: number,
    displayId: number,
    photoIndex: number
  ) {
    const columnMapping = ["A", "H", "O"];
    const row = this.workSheet.getRow(rowNum);
    row.getCell(
      columnMapping[
        photoIndex % InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT
      ] as string
    ).value = `${displayId}-${photoIndex + 1}`; // photoIndexは0スタートなので+1
  }

  private async pasteInstructionPhoto(
    rowNum: number,
    url: string,
    marginWidth: number,
    marginHeight: number,
    photoIndex: number
  ) {
    const columnMapping = ["A", "H", "O"];

    const data = await fetchImageAsBuffer(url);
    if (data === null) return;

    const row = this.workSheet.getRow(rowNum);
    const cell = row.getCell(
      columnMapping[
        photoIndex % InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT
      ] as string
    );
    const [cellWidth, cellHeight] = cellWidthHeightInPixel(cell);
    const sharped = await resizeImage(
      data,
      cellWidth - marginWidth,
      cellHeight - marginHeight
    );

    const imageId = this.workbook.addImage({
      buffer: sharped,
      extension: "jpeg",
    });

    pasteImageWithAspectRatio(
      this.workSheet,
      cell,
      imageId,
      InstructionPhotoSheetBuilder.COLUMN_COUNT_FOR_PHOTO_CELL,
      InstructionPhotoSheetBuilder.ROW_COUNT_FOR_PHOTO_CELL,
      cellWidth - marginWidth,
      cellHeight - marginHeight,
      cell.fullAddress.col,
      cell.fullAddress.row
    );
  }
}
