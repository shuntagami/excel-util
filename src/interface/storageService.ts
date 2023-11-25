export interface StorageService {
  /** uploadWithBytes upload data to bucket and kye specified */
  uploadWithBytes(
    data: Uint8Array,
    key: string,
    bucket?: string
  ): Promise<void>;
}
