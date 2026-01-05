
export enum ItemType {
  SALES = 'SALES',
  INVENTORY = 'INVENTORY',
  PURCHASE = 'PURCHASE'
}

export enum MovementType {
  IN = 'IN',
  OUT = 'OUT',
  TRANSFER = 'TRANSFER'
}

// OITM: Items
export interface Item {
  id: string;
  itemCode: string;
  itemName: string;
  manBatchNum: boolean;
  uom: string;
  isFrozen?: boolean; // New: Frozen status
}

// OWHS: Warehouses
export interface Warehouse {
  whsCode: string;
  whsName: string;
  location: string;
}

// OITW: Item Warehouse Data
export interface ItemWarehouse {
  itemCode: string;
  whsCode: string;
  onHand: number;
  isCommited: number;
  onOrder: number;
  minStock: number; // New: Minimum stock level alert
  avgCost: number; // New: Moving Average Price (MAP)
}

// OIBT: Batches
export interface Batch {
  batchNum: string;
  itemCode: string;
  whsCode: string;
  quantity: number;
  expDate: string;
  inDate: string;
}

// OINM: Inventory Ledger
export interface LedgerEntry {
  transNum: number;
  docType: string;
  docEntry: number;
  itemCode: string;
  whsCode: string;
  quantity: number;
  direction: MovementType;
  batchNum?: string;
  createDate: string;
  balanceAfter: number;
}

// OWTR/WTR1: Stock Transfer
export interface StockTransfer {
  docEntry: number;
  docDate: string;
  fromWhs: string;
  toWhs: string;
  comments: string;
  lines: StockTransferLine[];
}

export interface StockTransferLine {
  lineNum: number;
  itemCode: string;
  quantity: number;
  batchDetails?: { batchNum: string; quantity: number }[];
}

export interface WMSState {
  items: Item[];
  warehouses: Warehouse[];
  inventory: ItemWarehouse[];
  batches: Batch[];
  ledger: LedgerEntry[];
  transfers: StockTransfer[];
}
