import ExcelJS from "exceljs";
import { readFileSync } from "fs";

async function main() {
  const workbook = new ExcelJS.Workbook();
  const targetSheet = workbook.addWorksheet("Target Sheet");

  for (let i = 0; i < 1000; i++) {
    const imagePath = `images/image-${i}.jpeg`
    console.log(imagePath);

    const imageId = workbook.addImage({
      filename: imagePath,
      // buffer: readFileSync(imagePath),
      extension: 'jpeg',
    });
    console.log(`image added to workbook, imageId: ${imageId}`);


    targetSheet.addImage(imageId, {
      tl: { col: 1, row: (i + 1)*3 },
      ext: { width: 100, height: 100 }
    });
    console.log('image added to worksheet');

  }
  const path = `./results/${Date.now()}.xlsx`;
  console.log(`writing data to ${path}`);

  await workbook.xlsx.writeFile(path)
  console.log(`File written to ${path}`);
}

main().catch((error) => console.error(error));
