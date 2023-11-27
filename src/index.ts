import type { SQSEvent } from 'aws-lambda/trigger/sqs'
import { EventService } from './service/event/index'

/**
 * Entry point for consumer that will be triggered from sqs events
 */
export const handler = async (sqsEvent: SQSEvent) => { await EventService.handle(sqsEvent) }
