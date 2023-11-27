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

dayjs.locale('ja');
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

  /**
  * messageを使った指摘出力のメイン処理
  * zipファイル生成、S3にアップロード、出力完了apiのコールまで行う
  */
  private async processMessage(message: QueueMessage) {
    // /tmp is the only location we are allowed to write to in the Lambda environtment
    const tmpDir = path.join("/tmp", `${message.exportId}`);
    if (!existsSync(tmpDir)) {
      mkdirSync(tmpDir);
    }

    const paths: string[] = [];

    let baseFileName = "" // Excelやzipファイル名に使われる
    if (isInstructionResourceByClient(message)) { // 業者ごとの出力
      baseFileName = `指摘事項一覧(A3)_${dayjs().format("YYYYMMDD_HHmm")}`
      for (const instructionResource of message.resources) {
        const clientName = instructionResource.clientName;
        const data = await processInstructionResource(instructionResource);
        const tmpPath = path.join(
          tmpDir,
          `${clientName}_${baseFileName}.xlsx`
        );
        paths.push(tmpPath);
        writeFileSync(tmpPath, data);
      }
    } else { // not業者ごと(部屋別)
      baseFileName = `部屋別指摘事項一覧(A3)_${dayjs().format("YYYYMMDD_HHmm")}`
      const data = await processInstructionResource(message);
      const tmpPath = path.join(
        tmpDir,
        `${baseFileName}.xlsx`
      );
      paths.push(tmpPath);
      writeFileSync(tmpPath, data);
    }
    const zipPath = path.join(tmpDir, `${baseFileName}.zip`)
    await createZip(zipPath, paths);

    const exportId = message.exportId as number
    const orderId = message.orderId as number

    const s3Key = path.join("export", `${exportId}`, path.basename(zipPath))

    // 完成したzipファイルをS3にアップロード
    await storageService.uploadWithBytes(
      readFileSync(zipPath),
      s3Key
    );

    // 一時ファイル削除
    rmSync(tmpDir, { recursive: true, force: true });

    // blueprint-apiをコールして、export.stateを1(succeed)に更新する
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
