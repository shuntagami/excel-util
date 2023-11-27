export type QueueMessage = InstructionResourceByClient | InstructionResource

// 業者ごとの出力resource
export interface InstructionResourceByClient {
  exportId: number
  orderId: number
  resources: InstructionResource[]
}

// 業者ごとでない出力resoure
export interface InstructionResource {
  exportId?: number // 業者ごとじゃない時だけ入る
  orderId?: number // 業者ごとじゃない時だけ入る
  clientName?: string // 業者ごとの出力の時だけ入る
  blueprints: Blueprint[]
}

export interface Blueprint {
  id: number
  orderName: string
  blueprintName: string
  thumbnailUrl: string
  sheets: Sheet[]
}

interface Sheet {
  id: number
  sheetName: string
  operationCategory: string
  instructions: Instruction[]
}

export interface Instruction {
  id: number
  displayId: number
  room: string
  part: string
  finishing: string
  instruction: string
  note: string
  clientNames: string[]
  inspectors: string[]
  createdAt: string
  completedAt: string
  coordinateGraphics: string
  photos: InstructionPhoto[]
}

interface InstructionPhoto {
  displayId: number
  url: string
}
