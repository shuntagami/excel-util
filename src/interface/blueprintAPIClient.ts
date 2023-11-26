export interface BlueprintAPIClient {
  updateExportStatus(id: number, orderId: number,  state: number, key: string): Promise<any>
}
