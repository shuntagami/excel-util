import ExcelJS, { type Anchor } from 'exceljs'
import sharp from 'sharp'
import fs, { readFileSync } from 'fs'
import archiver from 'archiver'
import path = require('node:path')
import {
  type QueueMessage,
  type InstructionResource,
  type InstructionResourceByClient
} from '../types/InstructionResource'
import { InstructionSheetBuilder } from '../service/excel/InstructionSheetBuilder'
import { InstructionPhotoSheetBuilder } from '../service/excel/InstructionPhotoSheetBuilder'
import { exit } from 'process'

export const processInstructionResource = async (
  instructionResource: InstructionResource,
  instructionSheetName = '指摘一覧',
  photoSheetName = '写真一覧'
): Promise<Uint8Array> => {
  const workbook = new ExcelJS.Workbook()
  const targetSheet = workbook.addWorksheet(instructionSheetName, {
    views: [{}]
  })
  const targetPhotoSheet = workbook.addWorksheet(photoSheetName, {
    views: [{}]
  })

  const template = await new ExcelJS.Workbook().xlsx.readFile(
    './templates/instruction.xlsx'
  )
  const sourceSheet = template.worksheets[0]
  const sourcePhotoSheet = template.worksheets[1]
  if (sourceSheet === undefined || sourcePhotoSheet === undefined) {
    exit(1)
  }

  await new InstructionSheetBuilder(
    workbook,
    targetSheet,
    sourceSheet,
    instructionResource.blueprints
  ).build()

  await new InstructionPhotoSheetBuilder(
    workbook,
    targetPhotoSheet,
    sourcePhotoSheet,
    instructionResource.blueprints
  ).build()

  const data = new Uint8Array(await workbook.xlsx.writeBuffer())
  return data
}

export const loadJson = (
  rawData: string
): InstructionResource | InstructionResourceByClient | null => {
  try {
    const data: InstructionResource | InstructionResourceByClient =
      JSON.parse(rawData)
    return data
  } catch (error) {
    console.error('Error reading the file:', error)
    return null
  }
}

export const isInstructionResourceByClient = (
  resource: QueueMessage
): resource is InstructionResourceByClient => {
  return typeof resource === 'object' && resource !== null && 'resources' in resource
}

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
export const copyRows = (
  targetSheet: ExcelJS.Worksheet,
  sourceSheet: ExcelJS.Worksheet,
  from: number,
  to: number,
  targetStart: number
): void => {
  const targetRowIndex = targetStart
  for (let i = from; i <= to; i++) {
    const sourceRow = sourceSheet.getRow(i)
    const targetRow = targetSheet.getRow(targetRowIndex + i)

    targetRow.height = sourceRow.height
    sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const targetCell = targetRow.getCell(colNumber)
      const width = sourceSheet.getColumn(colNumber).width ?? sourceSheet.properties.defaultColWidth
      targetSheet.getColumn(colNumber).width = width

      targetCell.style = cell.style
      targetCell.value = cell.value

      // merge cell
      if (cell.isMerged) {
        if (cell.master !== cell) return

        let [colCount, rowCount] = [0, 0]

        for (let c = cell.fullAddress.col; ; c++) {
          const currentCell = cell.worksheet.getCell(cell.row, c)
          if (currentCell.master !== cell) break

          colCount++
        }

        for (let r = cell.fullAddress.row; ; r++) {
          const currentCell = cell.worksheet.getCell(r, cell.col)
          if (currentCell.master !== cell) break

          rowCount++
        }

        targetSheet.mergeCells(
          targetRowIndex + i,
          colNumber,
          targetRowIndex + i + rowCount - 1,
          colNumber + colCount - 1
        )
      }
    })
  }
}

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
): void => {
  const [columnWidthP, rowHeightP] = cellWidthHeightInPixel(cell)
  const [offC, offR] = [
    (columnWidthP - imageWidthP) / 2,
    (rowHeightP - imageHeightP) / 2
  ]

  worksheet.addImage(imageId, {
    tl: {
      nativeCol: colTl - 1,
      nativeRow: rowTl - 1,
      nativeColOff: Math.trunc(offC * 9525), // 1pixel = 9525EMU
      nativeRowOff: Math.trunc(offR * 9525)
    } as Anchor,
    // brを指定することで必ずセル内に画像が収まるようにする。但しアスペクト比が崩れる場合がある。
    br: {
      nativeCol: colTl - 1 + colCount,
      nativeRow: rowTl - 1 + rowCount,
      nativeColOff: -Math.trunc(offC * 9525), // 1pixel = 9525EMU
      nativeRowOff: -Math.trunc(offR * 9525)
    } as Anchor
  })
}

/**
 * get the width and height in pixel for specifi cell considering merged cell.
 */
export const cellWidthHeightInPixel = (
  cell: ExcelJS.Cell
): [number, number] => {
  const sheet = cell.worksheet
  let [width, height] = [0, 0]

  for (let c = cell.fullAddress.col; ; c++) {
    const currentCell = cell.worksheet.getCell(cell.fullAddress.row, c)
    // マージの有無に関わらず、現在のセルの左上（master）が対象のセルである限り値を加算し続ける
    if (currentCell.master !== cell) break

    // 初期値から列の横幅を変えてない場合に、値が取れない可能性があるたワークシート側のプロパティも参照
    width += sheet.getColumn(c).width ?? sheet.properties.defaultColWidth ?? 0
  }

  for (let r = cell.fullAddress.row; ; r++) {
    const currentCell = sheet.getCell(r, cell.col)

    // マージの有無に関わらず、現在のセルの左上（master）が対象のセルである限り値を加算し続ける
    if (currentCell.master !== cell) break

    // 初期値から列の横幅を変えてない場合に、値が取れない可能性があるたワークシート側のプロパティも参照
    height += sheet.getRow(r).height ?? sheet.properties.defaultRowHeight ?? 0
  }
  return [columnWidthInPixel(width), rowHeightInPixel(height)]
}

/**
 * get the column width in pixel.
 */
const columnWidthInPixel = (width: number, fontWidth = 8): number => {
  return ((256 * width + 128 / fontWidth) / 256) * fontWidth
}

/**
 * get the row height in pixel.
 */
const rowHeightInPixel = (height: number): number => {
  //  Pixels DPI (96 pixels per inch), Points DPI (72 points per inch)
  return (height * 96) / 72
}

export const fetchImageAsBuffer = async (
  imageUrl: string
): Promise<Buffer | null> => {
  const response = await fetch(imageUrl)
  if (!response.ok) {
    return null
  }
  const arrayBuffer = await response.arrayBuffer()
  return Buffer.from(arrayBuffer)
}

export const resizeImage = async (
  buffer: Buffer,
  width: number,
  height: number
): Promise<Buffer> => {
  return await sharp(buffer)
    .resize(Math.trunc(width), Math.trunc(height), { fit: 'inside' })
    .jpeg({ mozjpeg: true })
    .toBuffer()
}

export const addPageBreak = (
  sheet: ExcelJS.Worksheet,
  rowNum: number
): void => {
  sheet.getRow(rowNum).addPageBreak()
}

export const createZip = async (
  zipFileName: string,
  files: string[]
): Promise<void> => {
  await new Promise(async (resolve, reject) => {
    const output = fs.createWriteStream(zipFileName)
    const archive = archiver('zip', {
      zlib: { level: 9 },
      forceLocalTime: true
    })

    output.on('close', () => {
      resolve()
    })

    archive.on('error', (err) => {
      reject(err)
    })

    archive.pipe(output)

    files.forEach((file) => {
      const data = readFileSync(file)
      archive.append(data, { name: path.basename(file) })
    })

    archive.finalize()
  })
}
