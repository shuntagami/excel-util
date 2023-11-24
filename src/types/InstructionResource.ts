export type ClientData = {
  clientName: string;
  blueprints: Blueprint[];
};

export type Blueprint = {
  id: number;
  orderName: string;
  blueprintName: string;
  thumbnailUrl: string;
  sheets: Sheet[];
};

type Sheet = {
  id: number;
  sheetName: string;
  operationCategory: string;
  instructions: Instruction[];
};

export type Instruction = {
  id: number;
  displayId: number;
  room: string;
  part: string;
  finishing: string;
  instruction: string;
  note: string;
  clientNames: string[];
  inspectors: string[];
  createdAt: string;
  completedAt: string;
  coordinateGraphics: string;
  photos: InstructionPhoto[];
};

type InstructionPhoto = {
  displayId: number;
  url: string;
};
