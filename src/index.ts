import ExcelJS from "exceljs";
import { createWriteStream } from "fs";
import { execSync } from 'child_process';
import { copyRows } from "./utils/excel_util";
import { exit } from "process";

async function main() {
  const path = `./results/${Date.now()}.xlsx`;
  const stream = createWriteStream(path)
  const options = {
    useStyles: true,
    stream
  }
  const template = await new ExcelJS.Workbook().xlsx.readFile('./templates/template.xlsx')
  const sourceSheet = template.worksheets[0]
  if (sourceSheet === undefined ) {
    exit(1)
  }
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options)
  const stdout = execSync('ulimit -n').toString().trim();
  const ulimitN = parseInt(stdout, 10);
  console.log(ulimitN);

  const targetSheet = workbook.addWorksheet("Target Sheet");
  copyRows(targetSheet, sourceSheet, 1, 20, 1)

  for (let i = 0; i < 1000; i++) {
    const imagePath = `images/image-${i}.jpeg`

    const imageId = workbook.addImage({
      filename: imagePath,
      // buffer: readFileSync(imagePath),
      extension: 'jpeg',
    });

    targetSheet.addImage(imageId, {
      tl: { col: 1, row: (i + 1)*3 },
      ext: { width: 100, height: 100 }
    });

  }
  console.log(`writing data to ${path}`);
  await workbook.commit()
  console.log(`File written to ${path}`);
}

main().catch((error) => console.error(error));
