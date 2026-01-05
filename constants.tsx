
import { Item, Warehouse, ItemWarehouse, Batch, LedgerEntry, MovementType } from './types';

export const INITIAL_ITEMS: Item[] = [];

export const INITIAL_WAREHOUSES: Warehouse[] = [];

export const INITIAL_INVENTORY: ItemWarehouse[] = [];

export const INITIAL_BATCHES: Batch[] = [];

export const INITIAL_LEDGER: LedgerEntry[] = [];

export const LARAVEL_CODE_BLUEPRINT = `
/**
 * ERP Inventory System - Laravel 11 Implementation
 * Senior Backend Architecture
 */

// 1. MIGRATION: database/migrations/2024_01_01_000001_create_inventory_tables.php
Schema::create('oitm_items', function (Blueprint $table) {
    $table->id();
    $table->string('item_code')->unique();
    $table->string('item_name');
    $table->boolean('man_batch_num')->default(false);
    $table->string('uom', 20);
    $table->timestamps();
});

Schema::create('oitw_item_warehouse', function (Blueprint $table) {
    $table->id();
    $table->string('item_code');
    $table->string('whs_code');
    $table->decimal('on_hand', 18, 4)->default(0);
    $table->decimal('is_commited', 18, 4)->default(0);
    $table->decimal('avg_cost', 18, 4)->default(0);
    $table->decimal('min_stock', 18, 4)->default(0);
    $table->unique(['item_code', 'whs_code']);
});

Schema::create('oinm_ledger', function (Blueprint $table) {
    $table->id('trans_num');
    $table->string('doc_type', 10); // OIGN, OIGE, OWTR
    $table->unsignedBigInteger('doc_entry');
    $table->string('item_code');
    $table->string('whs_code');
    $table->decimal('qty', 18, 4);
    $table->decimal('balance_after', 18, 4);
    $table->timestamp('created_at')->useCurrent();
});

// 2. MODEL: app/Models/ItemWarehouse.php
namespace App\\Models;
class ItemWarehouse extends Model {
    protected $table = 'oitw_item_warehouse';
    protected $fillable = ['item_code', 'whs_code', 'on_hand', 'avg_cost', 'min_stock'];

    public function scopeLowStock($query) {
        return $query->whereRaw('on_hand < min_stock');
    }
}

// 3. SERVICE: app/Services/InventoryService.php
namespace App\\Services;
use Illuminate\\Support\\Facades\\DB;
use App\\Models\\ItemWarehouse;
use App\\Models\\InventoryLedger;

class InventoryService {
    public function adjustStock(string $itemCode, string $whsCode, float $qty, string $docType, int $docEntry) {
        return DB::transaction(function() use ($itemCode, $whsCode, $qty, $docType, $docEntry) {
            $inv = ItemWarehouse::where('item_code', $itemCode)
                ->where('whs_code', $whsCode)
                ->lockForUpdate()
                ->firstOrFail();

            if ($qty < 0 && ($inv->on_hand + $qty) < 0) {
                throw new \\Exception("Negative Inventory: Insufficient stock for " . $itemCode);
            }

            $inv->on_hand += $qty;
            $inv->save();

            return InventoryLedger::create([
                'doc_type' => $docType,
                'doc_entry' => $docEntry,
                'item_code' => $itemCode,
                'whs_code' => $whsCode,
                'qty' => $qty,
                'balance_after' => $inv->on_hand
            ]);
        });
    }
}
`;
