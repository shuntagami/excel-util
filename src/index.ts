import ExcelJS from "exceljs";
import { exit } from "process";
import {
  Blueprint,
  ClientData,
  InstructionSheetBuilder,
} from "./service/InstructionSheetBuilder";
import { writeFileSync, readFileSync } from "fs";
import { InstructionPhotoSheetBuilder } from "./service/InstructionPhotoSheetBuilder";

// const fs = require("fs").promises;

async function main() {
  const template = await new ExcelJS.Workbook().xlsx.readFile(
    "./templates/instruction3.xlsx"
  );
  const workbook = new ExcelJS.Workbook();
  const targetSheet = workbook.addWorksheet("Target Sheet", { views: [{}] });
  const targetPhotoSheet = workbook.addWorksheet("Target Photo Sheet", {
    views: [{}],
  });
  const sourceSheet = template.worksheets[0];
  const sourcePhotoSheet = template.worksheets[1];
  if (sourceSheet === undefined || sourcePhotoSheet === undefined) {
    exit(1);
  }

  // JSONファイルを読み込んで型にマッピングする関数
  const loadJsonFile = (filePath: string): ClientData | null => {
    try {
      const rawData = readFileSync(filePath, "utf8");
      const data: ClientData = JSON.parse(rawData);
      return data;
    } catch (error) {
      console.error("Error reading the file:", error);
      return null;
    }
  };

  // JSONファイルを読み込む
  const resource = loadJsonFile("./templates/resource.json");

  if (resource === null) {
    exit(1);
  }
  await new InstructionSheetBuilder(
    workbook,
    targetSheet,
    sourceSheet,
    resource.blueprints
  ).build(1);

  await new InstructionPhotoSheetBuilder(
    workbook,
    targetPhotoSheet,
    sourcePhotoSheet,
    resource.blueprints
  ).build(1);

  const data = new Uint8Array(await workbook.xlsx.writeBuffer());
  const dirPath = "./results";
  const path = [dirPath, `${Date.now()}.xlsx`].join("/");
  writeFileSync(path, data);
}

main().catch((error) => console.error(error));
