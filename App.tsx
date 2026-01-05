
import React, { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  INITIAL_ITEMS, 
  INITIAL_WAREHOUSES, 
  INITIAL_INVENTORY, 
  INITIAL_BATCHES, 
  INITIAL_LEDGER,
  LARAVEL_CODE_BLUEPRINT
} from './constants';
import { WMSState, MovementType, Item, Warehouse, LedgerEntry } from './types';
import { InventoryService } from './services/InventoryService';

// Localization Dictionary
const translations: any = {
  en: {
    dashboard: 'Dashboard',
    items: 'Inventory Status',
    masterData: 'Items Master Data',
    warehouses: 'Warehouses',
    receipt: 'Goods Receipt',
    issue: 'Goods Issue',
    transfer: 'Stock Transfer',
    reports: 'Reporting Center',
    ledger: 'Audit Ledger',
    setup: 'System Definition',
    laravel: 'Laravel Blueprint',
    onHand: 'On Hand',
    valuation: 'Valuation',
    alerts: 'Alerts',
    itemCode: 'Item Code',
    itemName: 'Item Name',
    uom: 'UoM',
    batches: 'Batches',
    actions: 'Actions',
    add: 'Add New Item',
    addWhs: 'Add New Warehouse',
    edit: 'Edit',
    delete: 'Delete',
    save: 'Save Changes',
    cancel: 'Cancel',
    confirmDelete: 'Are you sure? Records with history cannot be deleted.',
    success: 'Transaction posted successfully!',
    error: 'Operation failed.',
    lang: 'العربية',
    whs: 'Warehouse',
    whsCode: 'Whs Code',
    whsName: 'Whs Name',
    location: 'Location',
    qty: 'Quantity',
    minStock: 'Min Stock',
    emptyData: 'No data available. Please add master data first.',
    selectItem: 'Select Item',
    selectWhs: 'Select Warehouse',
    frozen: 'Frozen',
    isFrozen: 'Freeze Item (Hide from Reports)',
    status: 'Status',
    active: 'Active',
    filter: 'Filter Results',
    dateFrom: 'Date From',
    dateTo: 'Date To',
    allWhs: 'All Warehouses',
    allDocs: 'All Documents',
    stockStatusReport: 'Stock Levels & Valuation',
    auditReport: 'Inventory Transaction Audit',
    totalQty: 'Total Quantity',
    totalValue: 'Total Assets Value',
    searchItem: 'Search Item Code...',
    reset: 'Reset',
    apply: 'Apply Filters',
    print: 'Print Report',
    export: 'Export to Excel',
    refresh: 'Refresh Data',
    importExcel: 'Import Excel',
    chooseFile: 'Choose Excel File',
    reviewData: 'Review Imported Data',
    confirmImport: 'Confirm & Save All',
    importTemplate: 'Download Template',
    invalidFile: 'Invalid file format. Please use Excel (.xlsx or .xls)',
    openingBalance: 'Opening Balance',
    initialQty: 'Initial Qty',
    targetWhs: 'Target Warehouse',
    optional: '(Optional)'
  },
  ar: {
    dashboard: 'لوحة التحكم',
    items: 'حالة المخزون',
    masterData: 'بيانات الأصناف',
    warehouses: 'المستودعات',
    receipt: 'إضافة مخزنية',
    issue: 'صرف مخزني',
    transfer: 'تحويل مخزني',
    reports: 'مركز التقارير',
    ledger: 'سجل التدقيق',
    setup: 'تعريف النظام',
    laravel: 'كود لارافل',
    onHand: 'متوفر',
    valuation: 'التقييم',
    alerts: 'تنبيهات',
    itemCode: 'كود الصنف',
    itemName: 'اسم الصنف',
    uom: 'وحدة القياس',
    batches: 'شحنات',
    actions: 'إجراءات',
    add: 'إضافة صنف جديد',
    addWhs: 'إضافة مستودع جديد',
    edit: 'تعديل',
    delete: 'حذف',
    save: 'حفظ التغييرات',
    cancel: 'إلغاء',
    confirmDelete: 'هل أنت متأكد؟ لا يمكن حذف السجلات التي لها سجل حركات.',
    success: 'تمت العملية بنجاح!',
    error: 'فشلت العملية.',
    lang: 'English',
    whs: 'المستودع',
    whsCode: 'كود المستودع',
    whsName: 'اسم المستودع',
    location: 'الموقع',
    qty: 'الكمية',
    minStock: 'الحد الأدنى',
    emptyData: 'لا توجد بيانات. يرجى إضافة البيانات الأساسية أولاً.',
    selectItem: 'اختر الصنف',
    selectWhs: 'اختر المستودع',
    frozen: 'مجمد',
    isFrozen: 'تجميد الصنف (إخفاء من التقارير)',
    status: 'الحالة',
    active: 'نشط',
    filter: 'تصفية النتائج',
    dateFrom: 'من تاريخ',
    dateTo: 'إلى تاريخ',
    allWhs: 'كل المستودعات',
    allDocs: 'كل المستندات',
    stockStatusReport: 'تقرير أرصدة المخزون والتقييم',
    auditReport: 'سجل مراجعة حركات المخزون',
    totalQty: 'إجمالي الكمية',
    totalValue: 'إجمالي قيمة المخزون',
    searchItem: 'بحث بكود الصنف...',
    reset: 'إعادة تعيين',
    apply: 'تطبيق/تعيين',
    print: 'طباعة التقرير',
    export: 'تصدير إكسل',
    refresh: 'تحديث البيانات',
    importExcel: 'استيراد من إكسل',
    chooseFile: 'اختر ملف إكسل',
    reviewData: 'مراجعة البيانات المستوردة',
    confirmImport: 'تأكيد وحفظ الكل',
    importTemplate: 'تحميل القالب',
    invalidFile: 'تنسيق ملف غير صالح. يرجى استخدام إكسل (.xlsx أو .xls)',
    openingBalance: 'رصيد أول المدة',
    initialQty: 'كمية أول المدة',
    targetWhs: 'المستودع المستهدف',
    optional: '(اختياري)'
  }
};

const App: React.FC = () => {
  const [lang, setLang] = useState<'en' | 'ar'>('ar');
  const t = (key: string) => translations[lang][key] || key;

  const [state, setState] = useState<WMSState>({
    items: INITIAL_ITEMS,
    warehouses: INITIAL_WAREHOUSES,
    inventory: INITIAL_INVENTORY,
    batches: INITIAL_BATCHES,
    ledger: INITIAL_LEDGER,
    transfers: []
  });

  const [activeTab, setActiveTab] = useState<string>('dashboard');
  const [notification, setNotification] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
  
  // Modal & Import States
  const [isItemModalOpen, setIsItemModalOpen] = useState(false);
  const [editingItem, setEditingItem] = useState<Item | null>(null);
  const [itemForm, setItemForm] = useState({ 
    itemCode: '', 
    itemName: '', 
    uom: 'pcs', 
    manBatchNum: true, 
    isFrozen: false,
    initialQty: 0,
    targetWhs: ''
  });
  const [isWhsModalOpen, setIsWhsModalOpen] = useState(false);
  const [editingWhs, setEditingWhs] = useState<Warehouse | null>(null);
  const [whsForm, setWhsForm] = useState({ whsCode: '', whsName: '', location: '' });
  
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [importType, setImportType] = useState<'items' | 'warehouses'>('items');
  const [previewData, setPreviewData] = useState<any[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const inventoryService = useMemo(() => new InventoryService(state, setState), [state]);

  const showNotification = (type: 'success' | 'error', message: string) => {
    setNotification({ type, message });
    setTimeout(() => setNotification(null), 3000);
  };

  // --- Import Logic ---
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);
        setPreviewData(data);
      } catch (err) {
        showNotification('error', t('invalidFile'));
      }
    };
    reader.readAsBinaryString(file);
  };

  const confirmImport = () => {
    if (previewData.length === 0) return;

    if (importType === 'items') {
      let currentState = { ...state };
      
      previewData.forEach((row: any) => {
        const itemCode = String(row.ItemCode || row['كود الصنف'] || '');
        const itemName = String(row.ItemName || row['اسم الصنف'] || '');
        const qty = Number(row.Quantity || row['الكمية'] || row['كمية أول المدة'] || 0);
        const whs = String(row.WarehouseCode || row['كود المستودع'] || row['المستودع'] || '');

        if (itemCode && itemName && !currentState.items.some(i => i.itemCode === itemCode)) {
          // Add Item
          currentState.items.push({
            id: Date.now().toString() + Math.random(),
            itemCode,
            itemName,
            uom: String(row.UoM || row['الوحدة'] || 'pcs'),
            manBatchNum: row.ManageByBatch === 'Yes' || row['إدارة بالشحنات'] === 'نعم',
            isFrozen: false
          });

          // If Qty > 0, post Goods Receipt (OIGN)
          if (qty > 0 && whs) {
            // Check if whs exists
            if (currentState.warehouses.some(w => w.whsCode === whs)) {
              const service = new InventoryService(currentState, (s) => { currentState = s; });
              service.goodsReceipt(whs, { 
                lineNum: 1, 
                itemCode, 
                quantity: qty,
                batchDetails: row.BatchNum ? [{ batchNum: String(row.BatchNum), quantity: qty }] : undefined
              });
            }
          }
        }
      });
      setState(currentState);
    } else {
      const newWhs = [...state.warehouses];
      previewData.forEach((row: any) => {
        const whsCode = String(row.WhsCode || row['كود المستودع'] || '');
        const whsName = String(row.WhsName || row['اسم المستودع'] || '');
        if (whsCode && whsName && !newWhs.some(w => w.whsCode === whsCode)) {
          newWhs.push({
            whsCode,
            whsName,
            location: String(row.Location || row['الموقع'] || '')
          });
        }
      });
      setState({ ...state, warehouses: newWhs });
    }

    setPreviewData([]);
    setIsImportModalOpen(false);
    showNotification('success', t('success'));
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // --- Item Logic ---
  const handleSaveItem = () => {
    if (!itemForm.itemCode || !itemForm.itemName) return showNotification('error', 'Required fields missing');
    
    let newState = { ...state };
    const isNew = !editingItem;

    if (editingItem) {
      const idx = newState.items.findIndex(i => i.id === editingItem.id);
      newState.items[idx] = { ...editingItem, 
        itemCode: itemForm.itemCode,
        itemName: itemForm.itemName,
        uom: itemForm.uom,
        manBatchNum: itemForm.manBatchNum,
        isFrozen: itemForm.isFrozen 
      };
      setState(newState);
    } else {
      if (newState.items.some(i => i.itemCode === itemForm.itemCode)) return showNotification('error', 'Code must be unique');
      
      const newItem = {
        id: Date.now().toString(),
        itemCode: itemForm.itemCode,
        itemName: itemForm.itemName,
        uom: itemForm.uom,
        manBatchNum: itemForm.manBatchNum,
        isFrozen: itemForm.isFrozen
      };
      newState.items.push(newItem);

      // Handle Initial Quantity for new item
      if (itemForm.initialQty > 0 && itemForm.targetWhs) {
        // Use a temp service to update the state immediately
        const service = new InventoryService(newState, (s) => { newState = s; });
        service.goodsReceipt(itemForm.targetWhs, {
          lineNum: 1,
          itemCode: itemForm.itemCode,
          quantity: itemForm.initialQty
        });
      }
      setState(newState);
    }

    setIsItemModalOpen(false);
    showNotification('success', t('success'));
  };

  const handleSaveWhs = () => {
    if (!whsForm.whsCode || !whsForm.whsName) return showNotification('error', 'Required fields missing');
    const newWhs = [...state.warehouses];
    if (editingWhs) {
      const idx = newWhs.findIndex(w => w.whsCode === editingWhs.whsCode);
      newWhs[idx] = { ...whsForm };
    } else {
      if (newWhs.some(w => w.whsCode === whsForm.whsCode)) return showNotification('error', 'Code must be unique');
      newWhs.push({ ...whsForm });
    }
    setState({ ...state, warehouses: newWhs });
    setIsWhsModalOpen(false);
    showNotification('success', t('success'));
  };

  const handleStockAction = (type: 'receipt' | 'issue' | 'transfer', data: any) => {
    let result: { success: boolean, error?: string };
    if (type === 'receipt') {
      result = inventoryService.goodsReceipt(data.whsCode, { lineNum: 1, itemCode: data.itemCode, quantity: data.qty, batchDetails: data.batch ? [{ batchNum: data.batch, quantity: data.qty }] : undefined });
    } else if (type === 'issue') {
      result = inventoryService.goodsIssue(data.whsCode, { lineNum: 1, itemCode: data.itemCode, quantity: data.qty, batchDetails: data.batch ? [{ batchNum: data.batch, quantity: data.qty }] : undefined });
    } else {
      result = inventoryService.transferStock(data.fromWhs, data.toWhs, [{ lineNum: 1, itemCode: data.itemCode, quantity: data.qty, batchDetails: data.batch ? [{ batchNum: data.batch, quantity: data.qty }] : undefined }]);
    }
    if (result.success) showNotification('success', t('success'));
    else showNotification('error', result.error || t('error'));
  };

  return (
    <div className="flex min-h-screen bg-slate-50" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
      <aside className="w-64 bg-slate-900 text-white flex-shrink-0 print:hidden">
        <div className="p-6 border-b border-slate-800 flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg shadow-lg"><i className="fa-solid fa-warehouse text-xl"></i></div>
          <span className="font-bold text-xl tracking-tight">SAP WMS Pro</span>
        </div>
        <nav className="mt-6 px-3 space-y-1">
          <NavItem icon="chart-line" label={t('dashboard')} active={activeTab === 'dashboard'} onClick={() => setActiveTab('dashboard')} />
          <NavItem icon="boxes-stacked" label={t('items')} active={activeTab === 'items'} onClick={() => setActiveTab('items')} />
          <NavItem icon="folder-tree" label={t('masterData')} active={activeTab === 'masterData'} onClick={() => setActiveTab('masterData')} color="text-indigo-400" />
          <NavItem icon="building-user" label={t('warehouses')} active={activeTab === 'warehouses'} onClick={() => setActiveTab('warehouses')} color="text-amber-400" />
          <div className="pt-4 pb-2 px-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">{lang === 'ar' ? 'العمليات' : 'Transactions'}</div>
          <NavItem icon="circle-plus" label={t('receipt')} active={activeTab === 'receipt'} onClick={() => setActiveTab('receipt')} color="text-emerald-400" />
          <NavItem icon="circle-minus" label={t('issue')} active={activeTab === 'issue'} onClick={() => setActiveTab('issue')} color="text-rose-400" />
          <NavItem icon="right-left" label={t('transfer')} active={activeTab === 'transfer'} onClick={() => setActiveTab('transfer')} color="text-blue-400" />
          <div className="pt-4 pb-2 px-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">{lang === 'ar' ? 'التقارير' : 'Analytics'}</div>
          <NavItem icon="file-invoice-dollar" label={t('reports')} active={activeTab === 'reports'} onClick={() => setActiveTab('reports')} color="text-amber-400" />
          <NavItem icon="list-check" label={t('ledger')} active={activeTab === 'ledger'} onClick={() => setActiveTab('ledger')} />
          <div className="pt-4 pb-2 px-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">{lang === 'ar' ? 'الإعدادات' : 'System'}</div>
          <NavItem icon="gears" label={t('setup')} active={activeTab === 'setup'} onClick={() => setActiveTab('setup')} color="text-indigo-400" />
          <NavItem icon="code" label={t('laravel')} active={activeTab === 'laravel'} onClick={() => setActiveTab('laravel')} />
        </nav>
      </aside>

      <main className="flex-grow overflow-y-auto flex flex-col">
        <header className="bg-white h-16 border-b flex items-center justify-between px-8 sticky top-0 z-10 shadow-sm print:hidden">
          <h1 className="text-xl font-bold text-slate-700 uppercase">{t(activeTab)}</h1>
          <div className="flex items-center gap-4">
             <button onClick={() => setLang(lang === 'en' ? 'ar' : 'en')} className="px-4 py-1.5 bg-slate-100 hover:bg-slate-200 rounded-full text-xs font-bold transition-all border border-slate-200">{t('lang')}</button>
             <div className="w-10 h-10 bg-indigo-600 rounded-full flex items-center justify-center text-white font-bold shadow-md">SR</div>
          </div>
        </header>

        <div className="p-8 max-w-7xl mx-auto w-full animate-fade-in print:p-0">
          {notification && <Notification type={notification.type} message={notification.message} />}
          {activeTab === 'dashboard' && <Dashboard state={state} t={t} />}
          {activeTab === 'items' && <InventoryData state={state} t={t} />}
          {activeTab === 'masterData' && (
            <ItemsMaster 
              state={state} 
              t={t} 
              onOpenModal={(i?: Item) => { 
                if(i) { 
                  setEditingItem(i); 
                  setItemForm({
                    itemCode: i.itemCode, 
                    itemName: i.itemName, 
                    uom: i.uom, 
                    manBatchNum: i.manBatchNum, 
                    isFrozen: !!i.isFrozen,
                    initialQty: 0,
                    targetWhs: ''
                  }); 
                } else { 
                  setEditingItem(null); 
                  setItemForm({ 
                    itemCode: '', 
                    itemName: '', 
                    uom: 'pcs', 
                    manBatchNum: true, 
                    isFrozen: false,
                    initialQty: 0,
                    targetWhs: state.warehouses[0]?.whsCode || ''
                  }); 
                } 
                setIsItemModalOpen(true); 
              }} 
              onOpenImport={() => { setImportType('items'); setIsImportModalOpen(true); }}
              onDelete={(id: string) => setState({...state, items: state.items.filter(i=>i.id!==id)})} 
            />
          )}
          {activeTab === 'warehouses' && (
            <WarehousesMaster 
              state={state} 
              t={t} 
              onOpenModal={(w?: Warehouse) => { if(w) { setEditingWhs(w); setWhsForm({...w}); } else { setEditingWhs(null); setWhsForm({whsCode:'', whsName:'', location:''}); } setIsWhsModalOpen(true); }} 
              onOpenImport={() => { setImportType('warehouses'); setIsImportModalOpen(true); }}
              onDelete={(code: string) => setState({...state, warehouses: state.warehouses.filter(w=>w.whsCode!==code)})} 
            />
          )}
          {activeTab === 'receipt' && <ReceiptForm state={state} t={t} lang={lang} onSubmit={(d: any) => handleStockAction('receipt', d)} />}
          {activeTab === 'issue' && <IssueForm state={state} t={t} lang={lang} onSubmit={(d: any) => handleStockAction('issue', d)} />}
          {activeTab === 'transfer' && <TransferForm state={state} t={t} lang={lang} onSubmit={(d: any) => handleStockAction('transfer', d)} />}
          {activeTab === 'reports' && <ReportsCenter state={state} t={t} lang={lang} />}
          {activeTab === 'ledger' && <InventoryLedger ledger={state.ledger} t={t} />}
          {activeTab === 'setup' && <SystemSetup t={t} />}
          {activeTab === 'laravel' && <LaravelPreview />}
        </div>
      </main>

      {/* Item Modal Enhanced with Opening Balance */}
      {isItemModalOpen && (
        <Modal title={editingItem ? t('edit') : t('add')} onClose={() => setIsItemModalOpen(false)}>
           <div className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <FormInput label={t('itemCode')} value={itemForm.itemCode} onChange={(v:any) => setItemForm({...itemForm, itemCode: v})} disabled={!!editingItem} />
                <FormInput label={t('itemName')} value={itemForm.itemName} onChange={(v:any) => setItemForm({...itemForm, itemName: v})} />
              </div>
              <FormInput label={t('uom')} value={itemForm.uom} onChange={(v:any) => setItemForm({...itemForm, uom: v})} />
              <div className="flex gap-4">
                <Checkbox label={lang === 'ar' ? 'إدارة بالشحنات' : 'Batch Management'} checked={itemForm.manBatchNum} onChange={(v:boolean) => setItemForm({...itemForm, manBatchNum: v})} />
                <Checkbox label={t('isFrozen')} checked={itemForm.isFrozen} onChange={(v:boolean) => setItemForm({...itemForm, isFrozen: v})} color="accent-rose-600" textColor="text-rose-700" />
              </div>
              
              {!editingItem && (
                <div className="mt-6 p-4 bg-blue-50 border border-blue-100 rounded-2xl space-y-4">
                  <h4 className="text-xs font-black text-blue-600 uppercase tracking-widest flex items-center gap-2">
                    <i className="fa-solid fa-box-open"></i> {t('openingBalance')} {t('optional')}
                  </h4>
                  <div className="grid grid-cols-2 gap-4">
                    <FormInput label={t('initialQty')} type="number" value={itemForm.initialQty} onChange={(v:any) => setItemForm({...itemForm, initialQty: v})} />
                    <FormSelect 
                      label={t('targetWhs')} 
                      value={itemForm.targetWhs} 
                      onChange={(v:any) => setItemForm({...itemForm, targetWhs: v})} 
                      options={state.warehouses.map(w => ({ v: w.whsCode, l: w.whsName }))} 
                      placeholder={t('selectWhs')}
                    />
                  </div>
                </div>
              )}

              <button onClick={handleSaveItem} className="w-full bg-indigo-600 text-white py-3 rounded-xl font-bold mt-4 shadow-lg shadow-indigo-100">{t('save')}</button>
           </div>
        </Modal>
      )}

      {/* Warehouse Modal */}
      {isWhsModalOpen && (
        <Modal title={editingWhs ? t('edit') : t('addWhs')} onClose={() => setIsWhsModalOpen(false)}>
           <div className="space-y-4">
              <FormInput label={t('whsCode')} value={whsForm.whsCode} onChange={(v:any) => setWhsForm({...whsForm, whsCode: v})} disabled={!!editingWhs} />
              <FormInput label={t('whsName')} value={whsForm.whsName} onChange={(v:any) => setWhsForm({...whsForm, whsName: v})} />
              <FormInput label={t('location')} value={whsForm.location} onChange={(v:any) => setWhsForm({...whsForm, location: v})} />
              <button onClick={handleSaveWhs} className="w-full bg-amber-600 text-white py-3 rounded-xl font-bold mt-4 shadow-lg shadow-amber-100">{t('save')}</button>
           </div>
        </Modal>
      )}

      {/* Excel Import Modal */}
      {isImportModalOpen && (
        <Modal title={t('importExcel') + " (" + (importType === 'items' ? t('items') : t('warehouses')) + ")"} onClose={() => { setIsImportModalOpen(false); setPreviewData([]); }}>
           <div className="space-y-6">
              <div className="bg-slate-50 border-2 border-dashed border-slate-200 p-8 rounded-2xl text-center">
                 <i className="fa-solid fa-file-excel text-4xl text-emerald-500 mb-4"></i>
                 <p className="text-sm font-bold text-slate-600 mb-4">{t('chooseFile')}</p>
                 <input 
                   type="file" 
                   accept=".xlsx, .xls" 
                   ref={fileInputRef}
                   onChange={handleFileUpload} 
                   className="hidden" 
                   id="excel-file" 
                 />
                 <label htmlFor="excel-file" className="cursor-pointer bg-emerald-600 text-white px-6 py-2 rounded-xl font-bold text-xs hover:bg-emerald-700 transition-all inline-block">
                    {t('chooseFile')}
                 </label>
              </div>

              {previewData.length > 0 && (
                <div className="space-y-4">
                   <h4 className="font-black text-xs uppercase tracking-widest text-slate-400">{t('reviewData')}</h4>
                   <div className="max-h-60 overflow-y-auto border border-slate-100 rounded-xl scrollbar-hide">
                      <table className="w-full text-[10px] text-left rtl:text-right">
                         <thead className="sticky top-0 bg-slate-100">
                            <tr>
                               {Object.keys(previewData[0]).map(k => <th key={k} className="px-3 py-2 border-b">{k}</th>)}
                            </tr>
                         </thead>
                         <tbody className="divide-y divide-slate-50">
                            {previewData.map((row, idx) => (
                               <tr key={idx} className="hover:bg-slate-50">
                                  {Object.values(row).map((v: any, i) => <td key={i} className="px-3 py-2 text-slate-600 font-medium">{String(v)}</td>)}
                               </tr>
                            ))}
                         </tbody>
                      </table>
                   </div>
                   <button onClick={confirmImport} className="w-full bg-indigo-600 text-white py-4 rounded-xl font-bold shadow-lg shadow-indigo-100 flex items-center justify-center gap-2">
                      <i className="fa-solid fa-cloud-arrow-up"></i> {t('confirmImport')} ({previewData.length})
                   </button>
                </div>
              )}

              <div className="pt-4 border-t border-slate-100 space-y-2">
                  <div className="flex justify-between items-center text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                    <span>Columns: ItemCode, ItemName, UoM, Quantity, WarehouseCode</span>
                    <button className="text-indigo-500 hover:underline">{t('importTemplate')}</button>
                  </div>
              </div>
           </div>
        </Modal>
      )}
    </div>
  );
};

// --- Specialized Reporting Module with Export/Print Implementation ---
const ReportsCenter = ({ state, t, lang }: { state: WMSState, t: any, lang: string }) => {
  const [reportType, setReportType] = useState<'status' | 'audit'>('status');
  const [filters, setFilters] = useState({
    itemCode: '',
    whsCode: 'all',
    dateFrom: '',
    dateTo: '',
    docType: 'all'
  });

  const [activeFilters, setActiveFilters] = useState(filters);

  useEffect(() => {
    if (activeFilters.whsCode !== 'all' && !state.warehouses.some(w => w.whsCode === activeFilters.whsCode)) {
      setActiveFilters(prev => ({ ...prev, whsCode: 'all' }));
      setFilters(prev => ({ ...prev, whsCode: 'all' }));
    }
  }, [state.warehouses]);

  const filteredStock = useMemo(() => {
    return state.inventory.filter(inv => {
      const item = state.items.find(i => i.itemCode === inv.itemCode);
      if (item?.isFrozen) return false;
      if (activeFilters.whsCode !== 'all' && inv.whsCode !== activeFilters.whsCode) return false;
      if (activeFilters.itemCode && !inv.itemCode.toLowerCase().includes(activeFilters.itemCode.toLowerCase())) return false;
      return true;
    }).map(inv => ({
      ...inv,
      itemName: state.items.find(i => i.itemCode === inv.itemCode)?.itemName || '',
      totalValue: inv.onHand * inv.avgCost
    }));
  }, [state.inventory, state.items, activeFilters]);

  const filteredAudit = useMemo(() => {
    return state.ledger.filter(entry => {
      const item = state.items.find(i => i.itemCode === entry.itemCode);
      if (item?.isFrozen) return false;
      if (activeFilters.whsCode !== 'all' && entry.whsCode !== activeFilters.whsCode) return false;
      if (activeFilters.itemCode && !entry.itemCode.toLowerCase().includes(activeFilters.itemCode.toLowerCase())) return false;
      if (activeFilters.docType !== 'all' && entry.docType !== activeFilters.docType) return false;
      if (activeFilters.dateFrom && entry.createDate < activeFilters.dateFrom) return false;
      if (activeFilters.dateTo && entry.createDate > activeFilters.dateTo) return false;
      return true;
    });
  }, [state.ledger, state.items, activeFilters]);

  const stats = useMemo(() => {
    if (reportType === 'status') {
      const qty = filteredStock.reduce((s, i) => s + i.onHand, 0);
      const val = filteredStock.reduce((s, i) => s + i.totalValue, 0);
      return { qty, val };
    }
    return null;
  }, [filteredStock, reportType]);

  const handleApplyFilters = () => setActiveFilters({ ...filters });
  const handlePrint = () => window.print();

  const handleExport = () => {
    let exportData: any[] = [];
    let filename = '';

    if (reportType === 'status') {
      filename = `Stock_Report_${new Date().toLocaleDateString()}.xlsx`;
      exportData = filteredStock.map(d => ({
        [t('itemCode')]: d.itemCode,
        [t('itemName')]: d.itemName,
        [t('whs')]: d.whsCode,
        [t('onHand')]: d.onHand,
        'Avg Cost': d.avgCost,
        'Total Value': d.totalValue
      }));
    } else {
      filename = `Audit_Report_${new Date().toLocaleDateString()}.xlsx`;
      exportData = filteredAudit.map(entry => ({
        'Doc Type': entry.docType,
        'Doc Entry': entry.docEntry,
        [t('itemCode')]: entry.itemCode,
        [t('whs')]: entry.whsCode,
        [t('qty')]: entry.quantity,
        'Balance After': entry.balanceAfter,
        'Date': new Date(entry.createDate).toLocaleString()
      }));
    }

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, filename);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 rounded-2xl shadow-sm border border-slate-200 gap-4 print:hidden">
        <div className="flex flex-col gap-3">
          <div className="flex gap-4 p-1 bg-slate-100 rounded-xl w-fit">
            <button onClick={() => setReportType('status')} className={`px-6 py-2 rounded-lg text-sm font-bold transition-all ${reportType === 'status' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-500 hover:text-slate-900'}`}>{t('stockStatusReport')}</button>
            <button onClick={() => setReportType('audit')} className={`px-6 py-2 rounded-lg text-sm font-bold transition-all ${reportType === 'audit' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-500 hover:text-slate-900'}`}>{t('auditReport')}</button>
          </div>
        </div>
        
        <div className="flex items-center gap-3">
          <button onClick={handlePrint} className="px-4 py-2 bg-slate-100 text-slate-700 rounded-xl text-sm font-bold hover:bg-slate-200 flex items-center gap-2 border border-slate-200">
            <i className="fa-solid fa-print"></i> {t('print')}
          </button>
          <button onClick={handleExport} className="px-4 py-2 bg-emerald-50 text-emerald-700 rounded-xl text-sm font-bold hover:bg-emerald-100 flex items-center gap-2 border border-emerald-100">
            <i className="fa-solid fa-file-excel"></i> {t('export')}
          </button>
        </div>
      </div>

      {stats && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 print:grid-cols-2">
           <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-4">
             <div className="w-12 h-12 bg-blue-100 text-blue-600 rounded-xl flex items-center justify-center text-xl"><i className="fa-solid fa-boxes-stacked"></i></div>
             <div><p className="text-[10px] text-slate-400 font-bold uppercase">{t('totalQty')}</p><p className="text-2xl font-black text-slate-800">{stats.qty.toLocaleString()}</p></div>
           </div>
           <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-4">
             <div className="w-12 h-12 bg-emerald-100 text-emerald-600 rounded-xl flex items-center justify-center text-xl"><i className="fa-solid fa-money-bill-trend-up"></i></div>
             <div><p className="text-[10px] text-slate-400 font-bold uppercase">{t('totalValue')}</p><p className="text-2xl font-black text-emerald-600">${stats.val.toLocaleString()}</p></div>
           </div>
        </div>
      )}

      <div className="flex gap-4 overflow-x-auto pb-2 scrollbar-hide print:hidden">
         {state.warehouses.map(w => {
           const stockInWhs = state.inventory
             .filter(inv => inv.whsCode === w.whsCode && !state.items.find(i => i.itemCode === inv.itemCode)?.isFrozen)
             .reduce((sum, inv) => sum + inv.onHand, 0);
           return (
             <button 
               key={w.whsCode}
               onClick={() => setFilters(prev => ({ ...prev, whsCode: filters.whsCode === w.whsCode ? 'all' : w.whsCode }))}
               className={`flex-shrink-0 min-w-[180px] p-4 rounded-2xl border transition-all text-right ${filters.whsCode === w.whsCode ? 'bg-blue-600 border-blue-700 shadow-lg text-white' : 'bg-white border-slate-200 hover:border-blue-300 shadow-sm text-slate-800'}`}
             >
               <p className={`text-[10px] font-black uppercase tracking-tighter mb-1 ${filters.whsCode === w.whsCode ? 'text-blue-100' : 'text-slate-400'}`}>{w.whsCode}</p>
               <p className="font-bold text-sm truncate">{w.whsName}</p>
               <p className={`text-lg font-black mt-2 ${filters.whsCode === w.whsCode ? 'text-white' : 'text-blue-600'}`}>{stockInWhs.toLocaleString()}</p>
             </button>
           );
         })}
      </div>

      <div className="bg-white p-6 rounded-2xl shadow-md border border-slate-200 grid grid-cols-1 md:grid-cols-4 lg:grid-cols-6 gap-4 print:hidden">
        <div className="lg:col-span-1">
          <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">{t('whs')}</label>
          <select value={filters.whsCode} onChange={(e)=>setFilters({...filters, whsCode: e.target.value})} className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold">
            <option value="all">{t('allWhs')}</option>
            {state.warehouses.map(w => <option key={w.whsCode} value={w.whsCode}>{w.whsCode} - {w.whsName}</option>)}
          </select>
        </div>
        <div className="lg:col-span-1">
          <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">{t('itemCode')}</label>
          <input value={filters.itemCode} onChange={(e)=>setFilters({...filters, itemCode: e.target.value})} placeholder={t('searchItem')} className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold" />
        </div>
        {reportType === 'audit' && (
          <>
            <div>
              <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">{t('dateFrom')}</label>
              <input type="date" value={filters.dateFrom} onChange={(e)=>setFilters({...filters, dateFrom: e.target.value})} className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold" />
            </div>
            <div>
              <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">{t('dateTo')}</label>
              <input type="date" value={filters.dateTo} onChange={(e)=>setFilters({...filters, dateTo: e.target.value})} className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold" />
            </div>
          </>
        )}
        <div className="flex items-end gap-2 lg:col-span-2">
          <button onClick={handleApplyFilters} className="px-6 py-3 bg-indigo-600 text-white font-bold rounded-xl text-xs transition-all flex-grow shadow-lg shadow-indigo-100 flex items-center justify-center gap-2">
            <i className="fa-solid fa-check-double"></i> {t('apply')}
          </button>
          <button onClick={() => { setFilters({itemCode:'', whsCode:'all', dateFrom:'', dateTo:'', docType:'all'}); setActiveFilters({itemCode:'', whsCode:'all', dateFrom:'', dateTo:'', docType:'all'}); }} className="px-4 py-3 text-slate-400 font-bold hover:bg-slate-100 rounded-xl text-xs transition-all border border-slate-200">
            {t('reset')}
          </button>
        </div>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden print:border-none print:shadow-none">
        <div className="hidden print:block p-6 border-b text-center">
            <h2 className="text-2xl font-bold">{reportType === 'status' ? t('stockStatusReport') : t('auditReport')}</h2>
            <p className="text-slate-500 text-sm">{new Date().toLocaleString()}</p>
        </div>
        {reportType === 'status' ? (
          <table className="w-full text-left rtl:text-right">
            <thead className="bg-slate-50 border-b border-slate-200 text-[10px] font-black text-slate-500 uppercase tracking-widest print:bg-white">
              <tr>
                <th className="px-6 py-4">{t('itemCode')}</th>
                <th className="px-6 py-4">{t('whs')}</th>
                <th className="px-6 py-4 text-center">{t('onHand')}</th>
                <th className="px-6 py-4 text-right">Avg Cost</th>
                <th className="px-6 py-4 text-right">Total Assets</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredStock.length === 0 ? <tr><td colSpan={5} className="p-10 text-center text-slate-400 italic">{t('emptyData')}</td></tr> : filteredStock.map((d, i) => (
                <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                  <td className="px-6 py-4 font-bold text-slate-700">
                    <div className="flex flex-col">
                       <span>{d.itemCode}</span>
                       <span className="text-[10px] text-slate-400 font-medium truncate max-w-[200px]">{d.itemName}</span>
                    </div>
                  </td>
                  <td className="px-6 py-4"><span className="text-xs font-mono font-bold text-slate-400">{d.whsCode}</span></td>
                  <td className={`px-6 py-4 text-center font-black ${d.onHand < d.minStock ? 'text-rose-600 bg-rose-50/50' : 'text-blue-600'}`}>{d.onHand.toLocaleString()}</td>
                  <td className="px-6 py-4 text-right text-slate-400 font-mono font-bold">${d.avgCost.toFixed(2)}</td>
                  <td className="px-6 py-4 text-right font-black text-emerald-600 font-mono text-lg">${d.totalValue.toLocaleString()}</td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <table className="w-full text-left rtl:text-right">
            <thead className="bg-slate-50 border-b border-slate-200 text-[10px] font-black text-slate-500 uppercase tracking-widest print:bg-white">
              <tr>
                <th className="px-6 py-4">Doc / Ref</th>
                <th className="px-6 py-4">{t('itemCode')}</th>
                <th className="px-6 py-4">{t('whs')}</th>
                <th className="px-6 py-4 text-center">Movement</th>
                <th className="px-6 py-4 text-center">Balance After</th>
                <th className="px-6 py-4 text-right">Time</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredAudit.length === 0 ? <tr><td colSpan={6} className="p-10 text-center text-slate-400 italic">No records found.</td></tr> : filteredAudit.slice().reverse().map((entry, i) => (
                <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                  <td className="px-6 py-4">
                    <div className="flex items-center gap-2">
                      <span className={`w-2 h-2 rounded-full print:hidden ${entry.docType === 'OIGN' ? 'bg-emerald-500' : entry.docType === 'OIGE' ? 'bg-rose-500' : 'bg-blue-500'}`}></span>
                      <span className="font-bold text-slate-800">{entry.docType} <span className="text-slate-300 ml-1">#{entry.docEntry}</span></span>
                    </div>
                  </td>
                  <td className="px-6 py-4 text-sm font-bold text-slate-600">{entry.itemCode}</td>
                  <td className="px-6 py-4 text-xs font-mono font-bold text-slate-400">{entry.whsCode}</td>
                  <td className={`px-6 py-4 text-center font-black ${entry.quantity > 0 ? 'text-emerald-600 bg-emerald-50' : 'text-rose-600 bg-rose-50'}`}>
                    {entry.quantity > 0 ? `+${entry.quantity}` : entry.quantity}
                  </td>
                  <td className="px-6 py-4 text-center font-mono font-bold text-slate-800 bg-slate-50/50">{entry.balanceAfter.toLocaleString()}</td>
                  <td className="px-6 py-4 text-right text-[10px] font-bold text-slate-400">{new Date(entry.createDate).toLocaleString()}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
};

// --- Helper UI Components ---
const NavItem = ({ icon, label, active, onClick, color }: any) => (
  <button onClick={onClick} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${active ? 'bg-blue-600 text-white shadow-lg shadow-blue-900/20' : 'text-slate-400 hover:bg-slate-800 hover:text-white'}`}>
    <i className={`fa-solid fa-${icon} w-5 ${color && !active ? color : ''}`}></i>
    <span className="font-bold text-sm">{label}</span>
  </button>
);

const Notification = ({ type, message }: any) => (
  <div className={`mb-6 p-4 rounded-xl flex items-center gap-3 shadow-lg animate-fade-in ${type === 'success' ? 'bg-emerald-500 text-white' : 'bg-rose-500 text-white'}`}>
    <i className={`fa-solid ${type === 'success' ? 'fa-circle-check' : 'fa-circle-exclamation'}`}></i>
    <span className="font-medium text-sm">{message}</span>
  </div>
);

const Modal = ({ title, onClose, children }: any) => (
  <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4 print:hidden">
    <div className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full p-8 animate-fade-in overflow-hidden relative">
      <button onClick={onClose} className="absolute top-6 right-6 text-slate-300 hover:text-slate-500 transition-colors"><i className="fa-solid fa-xmark text-xl"></i></button>
      <h3 className="text-xl font-black text-slate-800 mb-6">{title}</h3>
      {children}
    </div>
  </div>
);

const FormInput = ({ label, value, onChange, type = "text", disabled = false }: any) => (
  <div className="flex-grow">
    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">{label}</label>
    <input disabled={disabled} type={type} value={value} onChange={(e)=>onChange(type === 'number' ? Number(e.target.value) : e.target.value)} className={`w-full p-4 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700 shadow-sm transition-all ${disabled ? 'opacity-50 cursor-not-allowed' : 'hover:border-slate-300'}`} />
  </div>
);

const FormSelect = ({ label, value, onChange, options, placeholder }: any) => (
  <div>
    <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">{label}</label>
    <div className="relative">
      <select value={value} onChange={(e)=>onChange(e.target.value)} className="w-full p-4 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700 shadow-sm transition-all hover:border-slate-300 appearance-none">
        {!value && <option value="">{placeholder}</option>}
        {options.map((o:any)=><option key={o.v} value={o.v}>{o.l}</option>)}
      </select>
      <i className="fa-solid fa-chevron-down absolute right-4 top-1/2 -translate-y-1/2 text-slate-300 pointer-events-none text-xs"></i>
    </div>
  </div>
);

const Checkbox = ({ label, checked, onChange, color = "accent-indigo-600", textColor = "text-slate-600" }: any) => (
  <div className={`flex items-center gap-3 py-2 px-3 rounded-lg hover:bg-slate-50 cursor-pointer`}>
    <input type="checkbox" id={label} checked={checked} onChange={(e) => onChange(e.target.checked)} className={`w-5 h-5 ${color}`} />
    <label htmlFor={label} className={`text-sm font-bold ${textColor} cursor-pointer`}>{label}</label>
  </div>
);

// --- Content Components ---
const Dashboard = ({ state, t }: any) => {
  const totalOnHand = state.inventory.reduce((sum: number, i: any) => sum + i.onHand, 0);
  const lowStockItems = state.inventory.filter((i: any) => {
    const item = state.items.find((mi: any) => mi.itemCode === i.itemCode);
    return !item?.isFrozen && i.onHand < i.minStock;
  });
  const totalValue = state.inventory.reduce((sum: number, i: any) => {
    const item = state.items.find((mi: any) => mi.itemCode === i.itemCode);
    return !item?.isFrozen ? sum + (i.onHand * i.avgCost) : sum;
  }, 0);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <StatCard title={t('onHand')} value={totalOnHand.toLocaleString()} icon="cubes" color="blue" />
        <StatCard title={t('valuation')} value={`$${totalValue.toLocaleString()}`} icon="sack-dollar" color="emerald" />
        <StatCard title={t('alerts')} value={lowStockItems.length.toString()} icon="triangle-exclamation" color={lowStockItems.length > 0 ? 'rose' : 'indigo'} />
      </div>
    </div>
  );
};

const StatCard = ({ title, value, icon, color }: any) => {
  const colors: any = { blue: 'bg-blue-500 shadow-blue-100', emerald: 'bg-emerald-500 shadow-emerald-100', indigo: 'bg-indigo-500 shadow-indigo-100', rose: 'bg-rose-500 shadow-rose-100' };
  return (
    <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center gap-6">
      <div className={`${colors[color]} w-16 h-16 rounded-2xl flex items-center justify-center text-white text-2xl shadow-lg`}><i className={`fa-solid fa-${icon}`}></i></div>
      <div><p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1">{title}</p><p className="text-3xl font-black text-slate-800 tracking-tight">{value}</p></div>
    </div>
  );
};

const ItemsMaster = ({ state, t, onOpenModal, onOpenImport, onDelete }: any) => (
  <div className="space-y-6">
    <div className="flex justify-between items-center">
      <h2 className="text-2xl font-black text-slate-800">{t('masterData')}</h2>
      <div className="flex gap-2">
        <button onClick={onOpenImport} className="bg-emerald-50 text-emerald-700 px-4 py-2 rounded-xl font-bold hover:bg-emerald-100 transition-all flex items-center gap-2 border border-emerald-100 shadow-sm">
           <i className="fa-solid fa-file-excel"></i> {t('importExcel')}
        </button>
        <button onClick={() => onOpenModal()} className="bg-indigo-600 text-white px-6 py-2 rounded-xl font-bold hover:bg-indigo-700 transition-all flex items-center gap-2 shadow-lg shadow-indigo-100">
           <i className="fa-solid fa-plus"></i> {t('add')}
        </button>
      </div>
    </div>
    <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
      <table className="w-full text-left rtl:text-right">
        <thead className="bg-slate-50 border-b border-slate-200"><tr className="text-[10px] font-black text-slate-500 uppercase tracking-widest"><th className="px-6 py-4">{t('itemCode')}</th><th className="px-6 py-4">{t('itemName')}</th><th className="px-6 py-4 text-center">{t('status')}</th><th className="px-6 py-4 text-center">{t('actions')}</th></tr></thead>
        <tbody className="divide-y divide-slate-100">
          {state.items.length === 0 ? <tr><td colSpan={4} className="p-10 text-center text-slate-400 italic">{t('emptyData')}</td></tr> : state.items.map((item: any) => (
            <tr key={item.id} className={`hover:bg-slate-50/50 ${item.isFrozen ? 'bg-rose-50/30 opacity-70' : ''}`}><td className="px-6 py-4 font-bold text-slate-700">{item.itemCode}</td><td className="px-6 py-4 font-medium text-slate-600">{item.itemName}</td><td className="px-6 py-4 text-center">{item.isFrozen ? <span className="px-2 py-1 bg-rose-100 text-rose-600 rounded text-[10px] font-black">{t('frozen')}</span> : <span className="px-2 py-1 bg-emerald-100 text-emerald-600 rounded text-[10px] font-black">{t('active')}</span>}</td><td className="px-6 py-4"><div className="flex justify-center gap-2"><button onClick={() => onOpenModal(item)} className="p-2 text-indigo-500 hover:bg-indigo-50 rounded-lg transition-all"><i className="fa-solid fa-pen"></i></button><button onClick={() => onDelete(item.id)} className="p-2 text-rose-500 hover:bg-rose-50 rounded-lg transition-all"><i className="fa-solid fa-trash"></i></button></div></td></tr>
          ))}
        </tbody>
      </table>
    </div>
  </div>
);

const WarehousesMaster = ({ state, t, onOpenModal, onOpenImport, onDelete }: any) => (
  <div className="space-y-6">
    <div className="flex justify-between items-center">
      <h2 className="text-2xl font-black text-slate-800">{t('warehouses')}</h2>
      <div className="flex gap-2">
        <button onClick={onOpenImport} className="bg-emerald-50 text-emerald-700 px-4 py-2 rounded-xl font-bold hover:bg-emerald-100 transition-all flex items-center gap-2 border border-emerald-100 shadow-sm">
           <i className="fa-solid fa-file-excel"></i> {t('importExcel')}
        </button>
        <button onClick={() => onOpenModal()} className="bg-amber-600 text-white px-6 py-2 rounded-xl font-bold hover:bg-amber-700 transition-all flex items-center gap-2 shadow-lg shadow-amber-100">
           <i className="fa-solid fa-plus"></i> {t('addWhs')}
        </button>
      </div>
    </div>
    <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
      <table className="w-full text-left rtl:text-right">
        <thead className="bg-slate-50 border-b border-slate-200"><tr className="text-[10px] font-black text-slate-500 uppercase tracking-widest"><th className="px-6 py-4">{t('whsCode')}</th><th className="px-6 py-4">{t('whsName')}</th><th className="px-6 py-4">{t('location')}</th><th className="px-6 py-4 text-center">{t('actions')}</th></tr></thead>
        <tbody className="divide-y divide-slate-100">
          {state.warehouses.length === 0 ? <tr><td colSpan={4} className="p-10 text-center text-slate-400 italic">{t('emptyData')}</td></tr> : state.warehouses.map((whs: any) => (
            <tr key={whs.whsCode} className="hover:bg-slate-50/50"><td className="px-6 py-4 font-bold text-slate-700">{whs.whsCode}</td><td className="px-6 py-4 font-medium text-slate-600">{whs.whsName}</td><td className="px-6 py-4 text-sm text-slate-500">{whs.location}</td><td className="px-6 py-4"><div className="flex justify-center gap-2"><button onClick={() => onOpenModal(whs)} className="p-2 text-indigo-500 hover:bg-indigo-50 rounded-lg transition-all"><i className="fa-solid fa-pen"></i></button><button onClick={() => onDelete(whs.whsCode)} className="p-2 text-rose-500 hover:bg-rose-50 rounded-lg transition-all"><i className="fa-solid fa-trash"></i></button></div></td></tr>
          ))}
        </tbody>
      </table>
    </div>
  </div>
);

const ReceiptForm = ({ state, onSubmit, t, lang }: any) => {
  const [whsCode, setWhsCode] = useState('');
  const [itemCode, setItemCode] = useState('');
  const [qty, setQty] = useState(1);
  const [batch, setBatch] = useState('B-' + Date.now().toString().slice(-4));
  useEffect(() => { if (!whsCode && state.warehouses[0]) setWhsCode(state.warehouses[0].whsCode); if (!itemCode && state.items[0]) setItemCode(state.items[0].itemCode); }, [state]);
  return (
    <div className="bg-white rounded-2xl shadow-xl border border-slate-200 p-8 max-w-2xl mx-auto space-y-6">
      <h2 className="text-2xl font-black text-slate-800 uppercase tracking-tighter">{t('receipt')}</h2>
      <FormSelect label={t('itemCode')} value={itemCode} onChange={setItemCode} options={state.items.filter((i:any)=>!i.isFrozen).map((i:any)=>({v:i.itemCode, l:i.itemCode+' - '+i.itemName}))} placeholder={t('selectItem')} />
      <FormSelect label={t('whs')} value={whsCode} onChange={setWhsCode} options={state.warehouses.map((w:any)=>({v:w.whsCode, l:w.whsCode+' - '+w.whsName}))} placeholder={t('selectWhs')} />
      <div className="grid grid-cols-2 gap-6"><FormInput label={t('qty')} type="number" value={qty} onChange={setQty} /><FormInput label={lang==='ar'?'رقم الشحنة':'Batch No'} value={batch} onChange={setBatch} /></div>
      <button onClick={() => onSubmit({itemCode, whsCode, qty, batch})} className="w-full bg-emerald-600 text-white py-4 rounded-xl font-bold shadow-lg shadow-emerald-100 disabled:opacity-50" disabled={!itemCode || !whsCode}>{t('save')}</button>
    </div>
  );
};

const IssueForm = ({ state, onSubmit, t, lang }: any) => {
  const [whsCode, setWhsCode] = useState('');
  const [itemCode, setItemCode] = useState('');
  const [qty, setQty] = useState(1);
  useEffect(() => { if (!whsCode && state.warehouses[0]) setWhsCode(state.warehouses[0].whsCode); if (!itemCode && state.items[0]) setItemCode(state.items[0].itemCode); }, [state]);
  return (
    <div className="bg-white rounded-2xl shadow-xl border border-slate-200 p-8 max-w-2xl mx-auto space-y-6">
      <h2 className="text-2xl font-black text-slate-800 uppercase tracking-tighter">{t('issue')}</h2>
      <FormSelect label={t('itemCode')} value={itemCode} onChange={setItemCode} options={state.items.filter((i:any)=>!i.isFrozen).map((i:any)=>({v:i.itemCode, l:i.itemName}))} placeholder={t('selectItem')} />
      <FormSelect label={t('whs')} value={whsCode} onChange={setWhsCode} options={state.warehouses.map((w:any)=>({v:w.whsCode, l:w.whsName}))} placeholder={t('selectWhs')} />
      <FormInput label={t('qty')} type="number" value={qty} onChange={setQty} />
      <button onClick={() => onSubmit({itemCode, whsCode, qty})} className="w-full bg-rose-600 text-white py-4 rounded-xl font-bold shadow-lg shadow-rose-100 disabled:opacity-50" disabled={!itemCode || !whsCode}>{t('save')}</button>
    </div>
  );
};

const TransferForm = ({ state, onSubmit, t, lang }: any) => {
  const [fromWhs, setFromWhs] = useState('');
  const [toWhs, setToWhs] = useState('');
  const [itemCode, setItemCode] = useState('');
  const [qty, setQty] = useState(1);
  useEffect(() => { if (!fromWhs && state.warehouses[0]) setFromWhs(state.warehouses[0].whsCode); if (!toWhs && state.warehouses[1]) setToWhs(state.warehouses[1].whsCode); if (!itemCode && state.items[0]) setItemCode(state.items[0].itemCode); }, [state]);
  return (
    <div className="bg-white rounded-2xl shadow-xl border border-slate-200 p-8 max-w-2xl mx-auto space-y-6">
      <h2 className="text-2xl font-black text-slate-800 uppercase tracking-tighter">{t('transfer')}</h2>
      <div className="grid grid-cols-2 gap-4">
        <FormSelect label={lang==='ar'?'من مستودع':'From Whs'} value={fromWhs} onChange={setFromWhs} options={state.warehouses.map((w:any)=>({v:w.whsCode, l:w.whsName}))} placeholder={t('selectWhs')} />
        <FormSelect label={lang==='ar'?'إلى مستودع':'To Whs'} value={toWhs} onChange={setToWhs} options={state.warehouses.map((w:any)=>({v:w.whsCode, l:w.whsName}))} placeholder={t('selectWhs')} />
      </div>
      <FormSelect label={t('itemCode')} value={itemCode} onChange={setItemCode} options={state.items.filter((i:any)=>!i.isFrozen).map((i:any)=>({v:i.itemCode, l:i.itemName}))} placeholder={t('selectItem')} />
      <FormInput label={t('qty')} type="number" value={qty} onChange={setQty} />
      <button onClick={() => onSubmit({fromWhs, toWhs, itemCode, qty})} className="w-full bg-blue-600 text-white py-4 rounded-xl font-bold shadow-lg shadow-blue-100 disabled:opacity-50" disabled={!itemCode || !fromWhs || !toWhs}>{t('save')}</button>
    </div>
  );
};

const InventoryData = ({ state, t }: any) => (
  <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
    <table className="w-full text-left rtl:text-right">
      <thead className="bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-500 uppercase tracking-widest"><tr><th className="px-6 py-5">{t('itemCode')}</th><th className="px-6 py-5">{t('whs')}</th><th className="px-6 py-5 text-center">{t('onHand')}</th><th className="px-6 py-5 text-center">{t('minStock')}</th></tr></thead>
      <tbody className="divide-y divide-slate-100">
        {state.inventory.filter((inv:any)=>!state.items.find(i=>i.itemCode===inv.itemCode)?.isFrozen).length === 0 ? <tr><td colSpan={4} className="p-10 text-center text-slate-400 italic">{t('emptyData')}</td></tr> : state.inventory.filter((inv:any)=>!state.items.find(i=>i.itemCode===inv.itemCode)?.isFrozen).map((inv: any, idx: number) => (
          <tr key={idx} className="hover:bg-slate-50/50 transition-colors"><td className="px-6 py-5 font-bold text-slate-800">{inv.itemCode}</td><td className="px-6 py-5 text-slate-500 font-mono text-xs">{inv.whsCode}</td><td className={`px-6 py-5 text-center font-mono font-bold ${inv.onHand < inv.minStock ? 'text-rose-600' : 'text-blue-600'}`}>{inv.onHand}</td><td className="px-6 py-5 text-center font-mono text-slate-400">{inv.minStock}</td></tr>
        ))}
      </tbody>
    </table>
  </div>
);

const InventoryLedger = ({ ledger, t }: any) => (
  <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
    <table className="w-full text-left rtl:text-right">
      <thead className="bg-slate-50 border-b border-slate-200 text-[10px] font-black text-slate-500 uppercase tracking-widest"><tr><th className="px-6 py-5">Doc</th><th className="px-6 py-5">{t('itemCode')}</th><th className="px-6 py-5">{t('whs')}</th><th className="px-6 py-5 text-center">{t('qty')}</th><th className="px-6 py-5 text-center">Time</th></tr></thead>
      <tbody className="divide-y divide-slate-100 text-sm">
        {ledger.length === 0 ? <tr><td colSpan={5} className="p-10 text-center text-slate-400 italic">No movements recorded.</td></tr> : ledger.slice().reverse().map((l: any, idx: number) => (
          <tr key={idx} className="hover:bg-slate-50/50"><td className="px-6 py-5 font-bold text-slate-800">{l.docType} <span className="text-slate-300">#{l.docEntry}</span></td><td className="px-6 py-5 font-bold text-slate-700">{l.itemCode}</td><td className="px-6 py-5 text-xs font-mono">{l.whsCode}</td><td className={`px-6 py-5 text-center font-black ${l.quantity > 0 ? 'text-emerald-600' : 'text-rose-600'}`}>{l.quantity > 0 ? `+${l.quantity}` : l.quantity}</td><td className="px-6 py-5 text-center text-slate-400 text-[10px] font-bold">{new Date(l.createDate).toLocaleTimeString()}</td></tr>
        ))}
      </tbody>
    </table>
  </div>
);

const SystemSetup = ({ t }: any) => (
  <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-8 max-w-4xl mx-auto space-y-6">
    <h2 className="text-3xl font-black text-slate-800 uppercase tracking-tighter">{t('setup')}</h2>
    <div className="bg-indigo-900 text-indigo-100 p-8 rounded-3xl shadow-xl grid grid-cols-1 md:grid-cols-3 gap-6">
      <div className="p-4 bg-indigo-800 rounded-2xl border border-indigo-700"><i className="fa-solid fa-shield-halved text-2xl mb-3 text-indigo-300"></i><h5 className="font-bold mb-1">Atomic Trans</h5><p className="text-[10px] opacity-70 italic">Ensures movements are all-or-nothing.</p></div>
      <div className="p-4 bg-indigo-800 rounded-2xl border border-indigo-700"><i className="fa-solid fa-list-ol text-2xl mb-3 text-indigo-300"></i><h5 className="font-bold mb-1">FIFO/Batch</h5><p className="text-[10px] opacity-70 italic">Advanced tracking of expiry.</p></div>
      <div className="p-4 bg-indigo-800 rounded-2xl border border-indigo-700"><i className="fa-solid fa-arrows-spin text-2xl mb-3 text-indigo-300"></i><h5 className="font-bold mb-1">Moving Avg</h5><p className="text-[10px] opacity-70 italic">Real-time cost recalculation.</p></div>
    </div>
  </div>
);

const LaravelPreview = () => (
  <div className="bg-slate-900 rounded-2xl overflow-hidden shadow-2xl border border-slate-700">
    <div className="bg-slate-800 px-6 py-4 border-b border-slate-700"><span className="text-xs font-bold text-slate-400 font-mono tracking-widest uppercase italic">app/ERP/BackendStack.php</span></div>
    <pre className="p-8 text-xs font-mono text-indigo-300 overflow-x-auto leading-relaxed"><code>{LARAVEL_CODE_BLUEPRINT}</code></pre>
  </div>
);

export default App;
