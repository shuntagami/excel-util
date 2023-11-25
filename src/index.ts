import ExcelJS from "exceljs";
import { exit } from "process";
import { InstructionSheetBuilder } from "./service/InstructionSheetBuilder";
import { writeFileSync, readFileSync } from "fs";
import { InstructionPhotoSheetBuilder } from "./service/InstructionPhotoSheetBuilder";
import {
  InstructionResource,
  InstructionResourceByClient,
} from "./types/InstructionResource";

const processInstructionResource = async (
  instructionResource: InstructionResource,
  instructionSheetName: string,
  photoSheetName: string
) => {
  const workbook = new ExcelJS.Workbook();
  const targetSheet = workbook.addWorksheet(instructionSheetName, {
    views: [{}],
  });
  const targetPhotoSheet = workbook.addWorksheet(photoSheetName, {
    views: [{}],
  });

  const template = await new ExcelJS.Workbook().xlsx.readFile(
    "./templates/instruction.xlsx"
  );
  const sourceSheet = template.worksheets[0];
  const sourcePhotoSheet = template.worksheets[1];
  if (sourceSheet === undefined || sourcePhotoSheet === undefined) {
    exit(1);
  }

  await new InstructionSheetBuilder(
    workbook,
    targetSheet,
    sourceSheet,
    instructionResource.blueprints
  ).build(1);

  await new InstructionPhotoSheetBuilder(
    workbook,
    targetPhotoSheet,
    sourcePhotoSheet,
    instructionResource.blueprints
  ).build(1);

  const data = new Uint8Array(await workbook.xlsx.writeBuffer());
  return data;
};

const loadJson = (
  rawData: string
): InstructionResource | InstructionResourceByClient | null => {
  try {
    const data: InstructionResource | InstructionResourceByClient =
      JSON.parse(rawData);
    return data;
  } catch (error) {
    console.error("Error reading the file:", error);
    return null;
  }
};

const isInstructionResourceByClient = (
  resource: any
): resource is InstructionResourceByClient => {
  return resource && "resources" in resource;
};

const saveToTmpDir = (fileName: string, data: Uint8Array) => {
  const dirPath = "./results";
  const path = [dirPath, fileName].join("/");
  writeFileSync(path, data);
};

async function main() {
  // JSONファイルを読み込む
  const rawJsonData = readFileSync(
    "./templates/resource_by_client.json",
    "utf8"
  );
  const resource: InstructionResource | InstructionResourceByClient | null =
    loadJson(rawJsonData);
  if (resource === null) {
    exit(1);
  }

  if (isInstructionResourceByClient(resource)) {
    for (const instructionResource of resource.resources) {
      const clientName = instructionResource.clientName;
      const data = await processInstructionResource(
        instructionResource,
        "指摘一覧",
        "写真一覧"
      );
      saveToTmpDir(`${clientName}_${Date.now()}.xlsx`, data);
    }
  } else {
    const data = await processInstructionResource(
      resource,
      "指摘一覧",
      "写真一覧"
    );
    saveToTmpDir(`${Date.now()}.xlsx`, data);
  }
}

main().catch((error) => console.error(error));
