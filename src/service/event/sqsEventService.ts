import type { SQSEvent } from "aws-lambda/trigger/sqs";
import { storageService } from "../s3";
import { Buffer } from "node:buffer";
import { loadJson } from "../../utils/excel_util";
import { QueueMessage } from "../../types/InstructionResource";

export class SQSEventService {
  /**
   * Handles an sqs event by processing every message of it
   */
  async handle(event: SQSEvent) {
    const dequeuedMessages = this.mapEventToDequeuedMessages(event);

    const promises = dequeuedMessages.map(async (message) => {
      try {
        await this.processMessage(message);
      } catch (error) {
        // TODO: エラーハンドリング
      }
    });
    await Promise.all(promises);
  }

  private async processMessage(message: QueueMessage) {
    const jsonString = JSON.stringify(message);
    const buf = Buffer.from(jsonString);
    await storageService.uploadWithBytes(buf, "sample2.json");
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
