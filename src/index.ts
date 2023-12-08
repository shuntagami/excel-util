import ExcelJS from "exceljs";
import { readFileSync } from "fs";
import { execSync } from 'child_process';

async function main() {
  const workbook = new ExcelJS.Workbook();
  const stdout = execSync('ulimit -n').toString().trim();
  const ulimitN = parseInt(stdout, 10);
  console.log(ulimitN);

  const targetSheet = workbook.addWorksheet("Target Sheet");

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
  const path = `./results/${Date.now()}.xlsx`;
  console.log(`writing data to ${path}`);

  await workbook.xlsx.writeFile(path)
  console.log(`File written to ${path}`);
}

main().catch((error) => console.error(error));
