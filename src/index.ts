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
  InstructionResource: InstructionResource,
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
    InstructionResource.blueprints
  ).build(1);

  await new InstructionPhotoSheetBuilder(
    workbook,
    targetPhotoSheet,
    sourcePhotoSheet,
    InstructionResource.blueprints
  ).build(1);

  const data = new Uint8Array(await workbook.xlsx.writeBuffer());
  const dirPath = "./results";
  const path = [dirPath, `${Date.now()}.xlsx`].join("/");
  writeFileSync(path, data);
};

const loadJson = (
  filePath: string
): InstructionResource | InstructionResourceByClient | null => {
  try {
    const rawData = readFileSync(filePath, "utf8");
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

async function main() {
  const template = await new ExcelJS.Workbook().xlsx.readFile(
    "./templates/instruction.xlsx"
  );
  const sourceSheet = template.worksheets[0];
  const sourcePhotoSheet = template.worksheets[1];
  if (sourceSheet === undefined || sourcePhotoSheet === undefined) {
    exit(1);
  }

  // JSONファイルを読み込む
  const resource: InstructionResource | InstructionResourceByClient | null =
    loadJson("./templates/resource_by_client.json");

  if (resource === null) {
    exit(1);
  }

  if (isInstructionResourceByClient(resource)) {
    for (const InstructionResource of resource.resources) {
      await processInstructionResource(
        InstructionResource,
        sourceSheet,
        sourcePhotoSheet
      );
    }
  } else {
    await processInstructionResource(resource, sourceSheet, sourcePhotoSheet);
  }
}

main().catch((error) => console.error(error));
