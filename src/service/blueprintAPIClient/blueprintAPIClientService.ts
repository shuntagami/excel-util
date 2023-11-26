import fetch from 'node-fetch';
import { BlueprintAPIClient } from '../../interface/blueprintAPIClient';
import { BLUEPRINT_API_BASE_URL, X_API_KEY } from '../../constant/constant';
import path = require("node:path");

export class BlueprintAPIClientService implements BlueprintAPIClient {
  async updateExportStatus(id: number, orderId: number, state: number, key: string): Promise<any> {
    const data = {
      state: state,
      key: key
    }
    const url = path.join(BLUEPRINT_API_BASE_URL, 'api', 'linkage', 'v1', 'orders', `${orderId}`, 'exports', `${id}`)
    const response = await fetch(url, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'X-API-KEY': X_API_KEY
      },
      body: JSON.stringify({data})
    });

    if (!response.ok) {
      throw new Error(`Error: ${response.status}`);
    }

    return response.json();
  }
}
