import { exit } from "process";
import {
  writeFileSync,
  readFileSync,
  unlinkSync,
  existsSync,
  mkdirSync,
  unlink,
  rmSync,
} from "fs";
import {
  InstructionResource,
  InstructionResourceByClient,
} from "./types/InstructionResource";
import {
  createZip,
  isInstructionResourceByClient,
  loadJson,
  processInstructionResource,
} from "./utils/excel_util";
import dayjs = require("dayjs");
import path = require("node:path");
import { storageService } from "./service/s3";

async function main() {
  const rawJsonData = readFileSync(
    "./templates/resource_by_client.json",
    "utf8"
  );
  const resource: InstructionResource | InstructionResourceByClient | null =
    loadJson(rawJsonData);
  if (resource === null) {
    exit(1);
  }
  const tmpDir = path.join("tmp", `${resource.exportId}`);
  if (!existsSync(tmpDir)) {
    mkdirSync(tmpDir);
  }

  const paths: string[] = [];
  if (isInstructionResourceByClient(resource)) {
    for (const instructionResource of resource.resources) {
      const clientName = instructionResource.clientName;
      const data = await processInstructionResource(instructionResource);
      const tmpPath = path.join(
        tmpDir,
        `in_${clientName}_${dayjs().format("YYYYMMDD")}.xlsx`
      );
      paths.push(tmpPath);
      writeFileSync(tmpPath, data);
    }
  } else {
    const data = await processInstructionResource(resource);
    const tmpPath = path.join(tmpDir, `in_${dayjs().format("YYYYMMDD")}.xlsx`);
    paths.push(tmpPath);
    writeFileSync(tmpPath, data);
  }
  const zipPath = path.join(tmpDir, `in_${dayjs().format("YYYYMMDD")}.zip`);
  await createZip(zipPath, paths);

  const result = await storageService.uploadWithBytes(
    readFileSync(zipPath),
    path.join("export", `${resource.exportId}`, path.basename(zipPath))
  );

  rmSync(tmpDir, { recursive: true, force: true });
}

main().catch((error) => console.error(error));
