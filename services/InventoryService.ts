
import { 
  WMSState, 
  MovementType, 
  ItemWarehouse, 
  Batch, 
  LedgerEntry, 
  StockTransferLine 
} from '../types';

export class InventoryService {
  private state: WMSState;
  private setState: (state: WMSState) => void;

  constructor(state: WMSState, setState: (state: WMSState) => void) {
    this.state = state;
    this.setState = setState;
  }

  public validateQuantity(itemCode: string, whsCode: string, qty: number, batches?: { batchNum: string; quantity: number }[]): boolean {
    const inv = this.state.inventory.find(i => i.itemCode === itemCode && i.whsCode === whsCode);
    if (!inv || inv.onHand < qty) return false;

    if (batches && batches.length > 0) {
      for (const b of batches) {
        const batchRec = this.state.batches.find(br => br.itemCode === itemCode && br.whsCode === whsCode && br.batchNum === b.batchNum);
        if (!batchRec || batchRec.quantity < b.quantity) return false;
      }
    }
    return true;
  }

  /**
   * Goods Receipt (OIGN) - Increase Stock
   */
  public goodsReceipt(whsCode: string, line: StockTransferLine): { success: boolean; error?: string } {
    const newInventory = [...this.state.inventory];
    const newBatches = [...this.state.batches];
    const newLedger = [...this.state.ledger];
    const timestamp = new Date().toISOString();

    let invIdx = newInventory.findIndex(i => i.itemCode === line.itemCode && i.whsCode === whsCode);
    if (invIdx === -1) {
      // Fix: Adding missing properties minStock and avgCost to satisfy ItemWarehouse interface
      newInventory.push({ 
        itemCode: line.itemCode, 
        whsCode: whsCode, 
        onHand: line.quantity, 
        isCommited: 0, 
        onOrder: 0,
        minStock: 0,
        avgCost: 0
      });
      invIdx = newInventory.length - 1;
    } else {
      newInventory[invIdx] = { ...newInventory[invIdx], onHand: newInventory[invIdx].onHand + line.quantity };
    }

    if (line.batchDetails) {
      for (const bd of line.batchDetails) {
        const batchIdx = newBatches.findIndex(b => b.itemCode === line.itemCode && b.whsCode === whsCode && b.batchNum === bd.batchNum);
        if (batchIdx === -1) {
          newBatches.push({ 
            batchNum: bd.batchNum, 
            itemCode: line.itemCode, 
            whsCode: whsCode, 
            quantity: bd.quantity, 
            expDate: new Date(Date.now() + 31536000000).toISOString().split('T')[0], // Default 1 year
            inDate: timestamp 
          });
        } else {
          newBatches[batchIdx] = { ...newBatches[batchIdx], quantity: newBatches[batchIdx].quantity + bd.quantity };
        }
      }
    }

    newLedger.push({
      transNum: newLedger.length + 1,
      docType: 'OIGN',
      docEntry: Math.floor(Math.random() * 1000),
      itemCode: line.itemCode,
      whsCode: whsCode,
      quantity: line.quantity,
      direction: MovementType.IN,
      createDate: timestamp,
      balanceAfter: newInventory[invIdx].onHand
    });

    this.setState({ ...this.state, inventory: newInventory, batches: newBatches, ledger: newLedger });
    return { success: true };
  }

  /**
   * Goods Issue (OIGE) - Decrease Stock
   */
  public goodsIssue(whsCode: string, line: StockTransferLine): { success: boolean; error?: string } {
    if (!this.validateQuantity(line.itemCode, whsCode, line.quantity, line.batchDetails)) {
      return { success: false, error: 'Insufficient stock for Goods Issue.' };
    }

    const newInventory = [...this.state.inventory];
    const newBatches = [...this.state.batches];
    const newLedger = [...this.state.ledger];
    const timestamp = new Date().toISOString();

    const invIdx = newInventory.findIndex(i => i.itemCode === line.itemCode && i.whsCode === whsCode);
    newInventory[invIdx] = { ...newInventory[invIdx], onHand: newInventory[invIdx].onHand - line.quantity };

    if (line.batchDetails) {
      for (const bd of line.batchDetails) {
        const batchIdx = newBatches.findIndex(b => b.itemCode === line.itemCode && b.whsCode === whsCode && b.batchNum === bd.batchNum);
        newBatches[batchIdx] = { ...newBatches[batchIdx], quantity: newBatches[batchIdx].quantity - bd.quantity };
      }
    }

    newLedger.push({
      transNum: newLedger.length + 1,
      docType: 'OIGE',
      docEntry: Math.floor(Math.random() * 1000),
      itemCode: line.itemCode,
      whsCode: whsCode,
      quantity: -line.quantity,
      direction: MovementType.OUT,
      createDate: timestamp,
      balanceAfter: newInventory[invIdx].onHand
    });

    this.setState({ ...this.state, inventory: newInventory, batches: newBatches, ledger: newLedger });
    return { success: true };
  }

  public transferStock(fromWhs: string, toWhs: string, lines: StockTransferLine[]): { success: boolean; error?: string } {
    for (const line of lines) {
      if (!this.validateQuantity(line.itemCode, fromWhs, line.quantity, line.batchDetails)) {
        return { success: false, error: `Insufficient stock for item ${line.itemCode}` };
      }
    }

    const newInventory = [...this.state.inventory];
    const newBatches = [...this.state.batches];
    const newLedger = [...this.state.ledger];
    const timestamp = new Date().toISOString();

    for (const line of lines) {
      const fromInvIdx = newInventory.findIndex(i => i.itemCode === line.itemCode && i.whsCode === fromWhs);
      newInventory[fromInvIdx] = { ...newInventory[fromInvIdx], onHand: newInventory[fromInvIdx].onHand - line.quantity };

      let toInvIdx = newInventory.findIndex(i => i.itemCode === line.itemCode && i.whsCode === toWhs);
      if (toInvIdx === -1) {
        // Fix: Adding missing properties minStock and avgCost to satisfy ItemWarehouse interface
        newInventory.push({ 
          itemCode: line.itemCode, 
          whsCode: toWhs, 
          onHand: line.quantity, 
          isCommited: 0, 
          onOrder: 0,
          minStock: 0,
          avgCost: 0
        });
        toInvIdx = newInventory.length - 1;
      } else {
        newInventory[toInvIdx] = { ...newInventory[toInvIdx], onHand: newInventory[toInvIdx].onHand + line.quantity };
      }

      if (line.batchDetails) {
        for (const bd of line.batchDetails) {
          const sIdx = newBatches.findIndex(b => b.itemCode === line.itemCode && b.whsCode === fromWhs && b.batchNum === bd.batchNum);
          const sourceBatch = newBatches[sIdx];
          newBatches[sIdx] = { ...sourceBatch, quantity: sourceBatch.quantity - bd.quantity };

          const tIdx = newBatches.findIndex(b => b.itemCode === line.itemCode && b.whsCode === toWhs && b.batchNum === bd.batchNum);
          if (tIdx === -1) {
            newBatches.push({ ...sourceBatch, whsCode: toWhs, quantity: bd.quantity, inDate: timestamp });
          } else {
            newBatches[tIdx] = { ...newBatches[tIdx], quantity: newBatches[tIdx].quantity + bd.quantity };
          }
        }
      }

      newLedger.push({
        transNum: newLedger.length + 1,
        docType: 'OWTR',
        docEntry: 1,
        itemCode: line.itemCode,
        whsCode: fromWhs,
        quantity: -line.quantity,
        direction: MovementType.OUT,
        createDate: timestamp,
        balanceAfter: newInventory[fromInvIdx].onHand
      });

      newLedger.push({
        transNum: newLedger.length + 1,
        docType: 'OWTR',
        docEntry: 1,
        itemCode: line.itemCode,
        whsCode: toWhs,
        quantity: line.quantity,
        direction: MovementType.IN,
        createDate: timestamp,
        balanceAfter: newInventory[toInvIdx].onHand
      });
    }

    this.setState({ ...this.state, inventory: newInventory, batches: newBatches, ledger: newLedger });
    return { success: true };
  }
}
