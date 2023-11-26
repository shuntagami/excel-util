import { BlueprintAPIClient } from '../../interface/blueprintAPIClient';

export class MockBlueprintAPIClientService implements BlueprintAPIClient {
  async updateExportStatus(id: number, orderId: number, state: number, key: string): Promise<any> {
    return null
  }
}
