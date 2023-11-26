import { writeFileSync } from "fs";
import { StorageService } from "../../interface/storageService";
import * as path from "path";

export class MockS3Service implements StorageService {
  async uploadWithBytes(
    data: Uint8Array,
    key: string,
  ): Promise<void> {
    const dirPath = "./results"
    const filePath = path.join(dirPath, path.basename(key));
    writeFileSync(filePath, data);
  }
}
