import type { SQSEvent } from "aws-lambda/trigger/sqs";
import { storageService } from "../s3";
import {
  createZip,
  isInstructionResourceByClient,
  loadJson,
  processInstructionResource,
} from "../../utils/excel_util";
import { QueueMessage } from "../../types/InstructionResource";
import dayjs = require("dayjs");
import path = require("node:path");
import {
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { blueprintAPIClient } from "../blueprintAPIClient";

export class SQSEventService {
  /**
   * Handles an sqs event by processing every message of it
   */
  async handle(event: SQSEvent) {
    const dequeuedMessages = this.mapEventToDequeuedMessages(event);

    const promises = dequeuedMessages.map(async (message) => {
      await this.processMessage(message);
    });

    await Promise.all(promises);
  }

  private async processMessage(message: QueueMessage) {
    const tmpDir = path.join("tmp", `${message.exportId}`);
    if (!existsSync(tmpDir)) {
      mkdirSync(tmpDir);
    }

    const paths: string[] = [];
    if (isInstructionResourceByClient(message)) {
      for (const instructionResource of message.resources) {
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
      const data = await processInstructionResource(message);
      const tmpPath = path.join(
        tmpDir,
        `in_${dayjs().format("YYYYMMDD")}.xlsx`
      );
      paths.push(tmpPath);
      writeFileSync(tmpPath, data);
    }
    const zipPath = path.join(tmpDir, `in_${dayjs().format("YYYYMMDD")}.zip`);
    await createZip(zipPath, paths);

    const exportId = message.exportId as number
    const orderId = message.orderId as number

    const s3Key = path.join("export", `${exportId}`, path.basename(zipPath))
    await storageService.uploadWithBytes(
      readFileSync(zipPath),
      s3Key
    );

    rmSync(tmpDir, { recursive: true, force: true });

    blueprintAPIClient.updateExportStatus(exportId, orderId, 1, s3Key)
  }

  private mapEventToDequeuedMessages(event: SQSEvent): QueueMessage[] {
    const messages = [];
    for (const record of event.Records) {
      const message = loadJson(record.body);
      if (message !== null) {
        messages.push(message);
      }
    }
    return messages;
  }
}
