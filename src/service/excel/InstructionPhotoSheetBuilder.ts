import type ExcelJS from 'exceljs'
import { type Blueprint } from '../../types/InstructionResource'
import {
  addPageBreak,
  cellWidthHeightInPixel,
  copyRows,
  fetchImageAsBuffer,
  pasteImageWithAspectRatio,
  resizeImage
} from '../../utils/excel_util'

export class InstructionPhotoSheetBuilder {
  static readonly PHOTO_TEMPLATE_ROW_SIZE = 32
  static readonly VERTICAL_PHOTO_COUNT = 2
  static readonly HORIZONTAL_PHOTO_COUNT = 3
  static readonly TOTAL_PHOTO_COUNT_PER_TEMPLATE =
    InstructionPhotoSheetBuilder.VERTICAL_PHOTO_COUNT *
    InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT

  static readonly COLUMN_COUNT_FOR_PHOTO_CELL = 7
  static readonly ROW_COUNT_FOR_PHOTO_CELL = 12

  constructor (
    private readonly workbook: ExcelJS.Workbook,
    private readonly workSheet: ExcelJS.Worksheet,
    private readonly templateSheet: ExcelJS.Worksheet,
    private readonly resources: Blueprint[]
  ) {
    this.workSheet.pageSetup = this.templateSheet.pageSetup
  }

  async build (
    rowNum = 1,
    marginWidth = 100,
    marginHeight = 100
  ): Promise<this> {
    let currentRowNum = rowNum

    for (const blueprint of this.resources) {
      for (const sheet of blueprint.sheets) {
        let photoIndexBySheet = 0
        for (const instruction of sheet.instructions) {
          for (const photo of instruction.photos) {
            if (
              photoIndexBySheet %
                InstructionPhotoSheetBuilder.TOTAL_PHOTO_COUNT_PER_TEMPLATE ===
              0
            ) {
              copyRows(
                this.workSheet,
                this.templateSheet,
                1,
                InstructionPhotoSheetBuilder.PHOTO_TEMPLATE_ROW_SIZE,
                currentRowNum - 1
              )
              if (currentRowNum !== 1) {
                addPageBreak(this.workSheet, currentRowNum - 1)
              }
              currentRowNum += 1 // テンプレートの2行目がスタート位置
              this.fillBlueprintContents(
                currentRowNum,
                blueprint.orderName,
                blueprint.blueprintName,
                sheet.operationCategory,
                sheet.sheetName
              )
              currentRowNum += 2 // ヘッダー分2行追加
            }

            this.fillInstructionContents(
              currentRowNum,
              instruction.displayId,
              photo.displayId,
              photoIndexBySheet
            )
            await this.pasteInstructionPhoto(
              currentRowNum + 1,
              photo.url,
              marginWidth,
              marginHeight,
              photoIndexBySheet
            )
            photoIndexBySheet++
            if (
              photoIndexBySheet %
                InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT ===
              0
            ) {
              currentRowNum += 14
            }
          }
        }
        const temp =
          photoIndexBySheet %
          InstructionPhotoSheetBuilder.TOTAL_PHOTO_COUNT_PER_TEMPLATE
        if (temp !== 0) {
          if (temp <= InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT) {
            currentRowNum += 28
          } else {
            currentRowNum += 14
          }
        }
      }
    }
    return this
  }

  private fillBlueprintContents (
    rowNum: number,
    orderName: string,
    blueprintName: string,
    operationCategory: string,
    sheetName: string
  ): void {
    const row = this.workSheet.getRow(rowNum)
    row.getCell('A').value = orderName
    row.getCell('I').value = operationCategory
    row.getCell('Q').value = blueprintName + ':' + sheetName
  }

  private fillInstructionContents (
    rowNum: number,
    displayId: number,
    photoDisplayId: number,
    photoIndex: number
  ): void {
    const columnMapping = ['A', 'H', 'O']
    const row = this.workSheet.getRow(rowNum)
    row.getCell(
      columnMapping[
        photoIndex % InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT
      ] as string
    ).value = `${displayId}-${photoDisplayId}`
  }

  private async pasteInstructionPhoto (
    rowNum: number,
    url: string,
    marginWidth: number,
    marginHeight: number,
    photoIndex: number
  ): Promise<void> {
    const columnMapping = ['A', 'H', 'O']

    const data = await fetchImageAsBuffer(url)
    if (data === null) return

    const row = this.workSheet.getRow(rowNum)
    const cell = row.getCell(
      columnMapping[
        photoIndex % InstructionPhotoSheetBuilder.HORIZONTAL_PHOTO_COUNT
      ] as string
    )
    const [cellWidth, cellHeight] = cellWidthHeightInPixel(cell)
    const sharped = await resizeImage(
      data,
      cellWidth - marginWidth,
      cellHeight - marginHeight
    )

    const imageId = this.workbook.addImage({
      buffer: sharped,
      extension: 'jpeg'
    })

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
    )
  }
}
