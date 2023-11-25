import ExcelJS from "exceljs";
import { exit } from "process";
import { InstructionSheetBuilder } from "./service/InstructionSheetBuilder";
import { writeFileSync, readFileSync, unlinkSync } from "fs";
import { InstructionPhotoSheetBuilder } from "./service/InstructionPhotoSheetBuilder";
import {
  InstructionResource,
  InstructionResourceByClient,
} from "./types/InstructionResource";
import { createZip } from "./utils/excel_util";
import dayjs = require("dayjs");
import path = require("node:path");

const processInstructionResource = async (
  instructionResource: InstructionResource,
  instructionSheetName = "指摘一覧",
  photoSheetName = "写真一覧"
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

async function main() {
  // JSONファイルを読み込む
  const rawJsonData = readFileSync("./templates/resource.json", "utf8");
  const resource: InstructionResource | InstructionResourceByClient | null =
    loadJson(rawJsonData);
  if (resource === null) {
    exit(1);
  }

  const paths: string[] = [];
  if (isInstructionResourceByClient(resource)) {
    for (const instructionResource of resource.resources) {
      const clientName = instructionResource.clientName;
      const data = await processInstructionResource(instructionResource);
      const tmpPath = path.join(
        "tmp",
        `in_${clientName}_${dayjs().format("YYYYMMDD")}.xlsx`
      );
      paths.push(tmpPath);
      writeFileSync(tmpPath, data);
    }
  } else {
    const data = await processInstructionResource(resource);
    const tmpPath = path.join("tmp", `in_${dayjs().format("YYYYMMDD")}.xlsx`);
    paths.push(tmpPath);
    writeFileSync(tmpPath, data);
  }
  await createZip(
    path.join("tmp", `in_${dayjs().format("YYYYMMDD")}.zip`),
    paths
  );

  paths.forEach((filePath) => {
    unlinkSync(filePath);
  });
}

main().catch((error) => console.error(error));
