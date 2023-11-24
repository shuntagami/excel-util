import ExcelJS from "exceljs";
import {
  addPageBreak,
  cellWidthHeightInPixel,
  copyRows,
  fetchImageAsBuffer,
  pasteImageWithAspectRatio,
  resizeImage,
} from "../utils/excel_util";
import dayjs from "dayjs";
import { Blueprint, Instruction } from "../types/InstructionResource";

export class InstructionSheetBuilder {
  static readonly INSTRUCTION_TEMPLATE_ROW_SIZE = 32;
  static readonly BLUEPRINT_IMAGE_ROW_SIZE = 29;
  static readonly BLUEPRINT_IMAGE_COLUMN_SIZE = 8;
  static readonly INSTRUCTION_ROW_SIZE = 26;

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
    // TODO: 図面を渡さなくても、orderame, bluerpintName, thumbnailUrlだけ外から渡せば良さそう
    for (const blueprint of this.resources) {
      for (const sheet of blueprint.sheets) {
        let nokori = InstructionSheetBuilder.INSTRUCTION_ROW_SIZE;
        for (const [
          instructionIndex,
          instruction,
        ] of sheet.instructions.entries()) {
          const amari =
            instructionIndex % InstructionSheetBuilder.INSTRUCTION_ROW_SIZE;
          if (amari === 0) {
            if (currentRowNum !== 1) {
              currentRowNum += 3; // 指摘項目の最後の段の空欄分
              addPageBreak(this.workSheet, currentRowNum);
            }

            copyRows(
              this.workSheet,
              this.templateSheet,
              1,
              InstructionSheetBuilder.INSTRUCTION_TEMPLATE_ROW_SIZE,
              currentRowNum - 1
            );
            currentRowNum += 1; // テンプレートの2行目がスタート位置
            this.fillBlueprintContents(
              currentRowNum,
              blueprint.orderName,
              blueprint.blueprintName,
              sheet.operationCategory,
              sheet.sheetName
            );
            currentRowNum += 2; // ヘッダー分2行追加
            await this.pasteBlueprintImage(
              currentRowNum,
              blueprint.thumbnailUrl,
              marginWidth,
              marginHeight
            );
          }
          this.fillInstructionContents(currentRowNum, instruction);
          currentRowNum += 1;
          nokori = InstructionSheetBuilder.INSTRUCTION_ROW_SIZE - amari;
        }
        // シート単位でテンプレートを切り替えるので残った分、currentRowNumに足す
        currentRowNum += nokori;
      }
    }

    return this;
  }

  // テンプレートの以下の項目を埋める
  // 案件名
  // 検査種類
  // 図面名
  // 図面の画像
  private fillBlueprintContents(
    currentRowNum: number,
    orderName: string,
    blueprintName: string,
    operationCategory: string,
    sheetName: string
  ) {
    const currentRow = this.workSheet.getRow(currentRowNum);
    const nextRow = this.workSheet.getRow(currentRowNum + 1);

    currentRow.getCell("A").value = orderName;
    currentRow.getCell("F").value = operationCategory;
    nextRow.getCell("A").value = dayjs(Date.now()).format("YYYY/MM/DD");
    nextRow.getCell("F").value = blueprintName + ":" + sheetName;
  }

  private async pasteBlueprintImage(
    currentRowNum: number,
    url: string,
    marginWidth: number,
    marginHeight: number
  ) {
    const data = await fetchImageAsBuffer(url);
    if (data === null) return;

    const imageCell = this.workSheet.getRow(currentRowNum).getCell("A");
    const [cellWidth, cellHeight] = cellWidthHeightInPixel(imageCell);
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
      imageCell,
      imageId,
      InstructionSheetBuilder.BLUEPRINT_IMAGE_COLUMN_SIZE,
      InstructionSheetBuilder.BLUEPRINT_IMAGE_ROW_SIZE,
      cellWidth - marginWidth,
      cellHeight - marginHeight,
      imageCell.fullAddress.col,
      imageCell.fullAddress.row
    );
  }

  private fillInstructionContents(
    currentRowNum: number,
    instruction: Instruction
  ) {
    const currentRow = this.workSheet.getRow(currentRowNum);
    currentRow.getCell("I").value = instruction.displayId;
    currentRow.getCell("J").value = instruction.room;
    currentRow.getCell("K").value = instruction.part;
    currentRow.getCell("L").value = instruction.finishing;
    currentRow.getCell("M").value = instruction.instruction;
    currentRow.getCell("N").value = instruction.clientNames.join(",");
    currentRow.getCell("O").value = instruction.inspectors.join(",");
    currentRow.getCell("P").value = instruction.createdAt;
    currentRow.getCell("Q").value = instruction.completedAt;
    currentRow.getCell("R").value = instruction.note;
  }
}
