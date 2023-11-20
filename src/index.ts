import ExcelJS from "exceljs";
import { exit } from "process";
import {
  Blueprint,
  ClientData,
  InstructionSheetBuilder,
} from "./service/InstructionSheetBuilder";
import { writeFileSync, readFileSync } from "fs";

// const fs = require("fs").promises;

async function main() {
  const template = await new ExcelJS.Workbook().xlsx.readFile(
    "./templates/instruction.xlsx"
  );
  const workbook = new ExcelJS.Workbook();
  const targetSheet = workbook.addWorksheet("Target Sheet");
  const sourceSheet = template.worksheets[0];
  if (sourceSheet === undefined) {
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
  // InstructionSheetBuilder;
  const result = await new InstructionSheetBuilder(
    workbook,
    targetSheet,
    sourceSheet,
    resource.blueprints
  ).build(1);

  const data = new Uint8Array(await workbook.xlsx.writeBuffer());
  const dirPath = "./results";
  const path = [dirPath, `${Date.now()}.xlsx`].join("/");
  writeFileSync(path, data);
}

main().catch((error) => console.error(error));
