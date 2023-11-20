import ExcelJS from "exceljs";
import { copyRows } from "../utils/excel_util";
import dayjs from "dayjs";

type ExportResource = {
  blueprints: Blueprint[];
  instructionPhotos: InstructionPhoto[];
};

export type ClientData = {
  clientName: string;
  blueprints: Blueprint[];
  instructionPhotos: InstructionPhoto[];
};

export type Blueprint = {
  id: number;
  orderName: string;
  blueprintName: string;
  thumbnailUrl: string;
  sheets: Sheet[];
};

type Sheet = {
  id: number;
  sheetName: string;
  operationCategory: string;
  instructions: Instruction[];
};

type Instruction = {
  id: number;
  displayId: number;
  room: string;
  part: string;
  finishing: string;
  instruction: string;
  note: string;
  inspectors: string;
  createdAt: string;
  completedAt: string;
  clientName: string;
  coordinateGraphics: string;
};

type InstructionPhoto = {
  id: number;
  url: string;
  blueprintId: string;
  blueprintName: string;
  sheetId: string;
  sheetName: string;
  operationCategory: string;
  instructionId: string;
  displayId: string;
};

export class InstructionSheetBuilder {
  static readonly INSTRUCTION_TEMPLATE_ROW_SIZE = 33;
  static readonly INSTRUCTION_ROW_SIZE = 26;

  constructor(
    private readonly workbook: ExcelJS.Workbook,
    private readonly workSheet: ExcelJS.Worksheet,
    private readonly templateSheet: ExcelJS.Worksheet,
    private readonly resources: Blueprint[]
  ) {
    this.workSheet.pageSetup = this.templateSheet.pageSetup;
  }

  async build(rowNum: number): Promise<this> {
    let currentRowNum = rowNum;
    this.resources.forEach((blueprint) => {
      const orderName = blueprint.orderName;
      const blueprintName = blueprint.blueprintName;
      const thumbnailUrl = blueprint.thumbnailUrl;

      blueprint.sheets.forEach((sheet) => {
        let nokori = InstructionSheetBuilder.INSTRUCTION_ROW_SIZE;
        sheet.instructions.forEach(async (instruction, instruction_index) => {
          const amari =
            instruction_index % InstructionSheetBuilder.INSTRUCTION_ROW_SIZE;
          if (amari === 0) {
            if (instruction_index !== 0) {
              currentRowNum += 3; // 指摘項目の最後の段の空欄分
            }
            // テンプレートをコピー
            copyRows(
              this.workSheet,
              this.templateSheet,
              1,
              InstructionSheetBuilder.INSTRUCTION_TEMPLATE_ROW_SIZE,
              currentRowNum - 1
            );
            currentRowNum += 1; // テンプレートの2行目がスタート位置
            this.fillBlueprintContents(currentRowNum, blueprint, sheet);
            currentRowNum += 2; // ヘッダー分2行追加
          }
          this.fillInstructionContents(currentRowNum, instruction);
          currentRowNum += 1;
          nokori = nokori - amari;
        });
        // シート単位でテンプレートを切り替えるので残った分、currentRowNumに足す
        currentRowNum += nokori + 2;
      });
    });

    return this;
  }

  // テンプレートの以下の項目を埋める
  // 案件名
  // 検査種類
  // 図面名
  // 図面の画像
  private fillBlueprintContents(
    currentRowNum: number,
    blueprint: Blueprint,
    sheet: Sheet
  ) {
    const currentRow = this.workSheet.getRow(currentRowNum);
    const nextRow = this.workSheet.getRow(currentRowNum + 1);

    currentRow.getCell("A").value = blueprint.orderName;
    currentRow.getCell("F").value = sheet.operationCategory;
    nextRow.getCell("A").value = dayjs(Date.now()).format("YYYY/MM/DD");
    nextRow.getCell("F").value = blueprint.blueprintName;

    // TODO: 図面の貼り付け
    // const instructions = sheet.instructions;
    // instructionをループして、coordinateGraphicsを使ってsvg string組み立てて貼り付けみたいなことする必要あり。
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
    currentRow.getCell("N").value = instruction.clientName;
    currentRow.getCell("O").value = instruction.inspectors;
    currentRow.getCell("P").value = instruction.createdAt;
    currentRow.getCell("Q").value = instruction.completedAt;
    currentRow.getCell("R").value = instruction.note;
  }
}
