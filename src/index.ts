import ExcelJS from "exceljs";
import { exit } from "process";
import { InstructionSheetBuilder } from "./service/InstructionSheetBuilder";
import { writeFileSync, readFileSync } from "fs";
import { InstructionPhotoSheetBuilder } from "./service/InstructionPhotoSheetBuilder";
import { ClientData } from "./types/InstructionResource";

const processClientData = async (
  clientData: ClientData,
  sourceSheet: ExcelJS.Worksheet,
  sourcePhotoSheet: ExcelJS.Worksheet
) => {
  const workbook = new ExcelJS.Workbook();
  const targetSheet = workbook.addWorksheet("Target Sheet", { views: [{}] });
  const targetPhotoSheet = workbook.addWorksheet("Target Photo Sheet", {
    views: [{}],
  });

  await new InstructionSheetBuilder(
    workbook,
    targetSheet,
    sourceSheet,
    clientData.blueprints
  ).build(1);

  await new InstructionPhotoSheetBuilder(
    workbook,
    targetPhotoSheet,
    sourcePhotoSheet,
    clientData.blueprints
  ).build(1);

  const data = new Uint8Array(await workbook.xlsx.writeBuffer());
  const dirPath = "./results";
  const path = [dirPath, `${Date.now()}.xlsx`].join("/");
  writeFileSync(path, data);
};

// JSONファイルを読み込んで型にマッピングする関数
const loadJsonFile = (filePath: string): ClientData | ClientData[] | null => {
  try {
    const rawData = readFileSync(filePath, "utf8");
    const data: ClientData | ClientData[] = JSON.parse(rawData);
    return data;
  } catch (error) {
    console.error("Error reading the file:", error);
    return null;
  }
};

async function main() {
  const template = await new ExcelJS.Workbook().xlsx.readFile(
    "./templates/instruction3.xlsx"
  );
  const sourceSheet = template.worksheets[0];
  const sourcePhotoSheet = template.worksheets[1];
  if (sourceSheet === undefined || sourcePhotoSheet === undefined) {
    exit(1);
  }

  // JSONファイルを読み込む
  const resource = loadJsonFile("./templates/resource.json");

  if (resource === null) {
    exit(1);
  }

  if (Array.isArray(resource)) {
    for (const clientData of resource) {
      await processClientData(clientData, sourceSheet, sourcePhotoSheet);
    }
  } else {
    await processClientData(resource, sourceSheet, sourcePhotoSheet);
  }
}

main().catch((error) => console.error(error));
