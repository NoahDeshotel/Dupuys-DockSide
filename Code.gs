/*************************************************
 * DUPUYS DOCKSIDE V4 – COMPLETE PRODUCTION-READY SYSTEM
 * Ready for fresh Google Sheet initialization
 * 
 * SETUP: Just paste this code and run:
 * Menu > ⚓ Dupuys Dockside V4 > 🔧 Initialize System
 *************************************************/

/******** CONFIG ********/

// ===== SIMPLE LICENSE CHECK (Subterra) =====
// Set these two values per client deployment:
const LICENSE_URL = 'https://script.google.com/macros/s/YOUR_WEB_APP_URL/exec'; // <-- replace with your real URL
const TENANT_ID   = 'client-xyz'; // <-- change per client
const FAIL_OPEN   = false; // fail-closed to avoid simple trigger bypass
const LICENSE_CACHE_MS = 15 * 60 * 1000; // cache license result for 15 minutes

// Contact + suspension messaging (shown when account is locked)
const SUPPORT_CONTACT = 'support@subterra.one';
const LOCK_MESSAGE =
  '🔒 Service is suspended due to non-payment. If this is a misunderstanding or you believe you are current, ' +
  'please contact Subterra at ' + SUPPORT_CONTACT + ' to restore access.';

function isLicensed() {
  const props = PropertiesService.getScriptProperties();
  const k = 'LIC:' + TENANT_ID;

  // 1) Use fresh cached value if available (avoids UrlFetch inside simple triggers)
  const cached = props.getProperty(k);
  if (cached) {
    const parts = cached.split('|');
    const status = parts[0];
    const ts = Number(parts[1] || 0);
    if (Date.now() - ts < LICENSE_CACHE_MS) {
      return status === 'ok';
    }
  }

  // 2) Refresh from server (this may fail in simple triggers; we handle that below)
  try {
    const res = UrlFetchApp.fetch(LICENSE_URL + '?tenant=' + encodeURIComponent(TENANT_ID), {
      muteHttpExceptions: true,
      followRedirects: true,
      timeout: 2000
    });
    const ok = res.getResponseCode() === 200 && (res.getContentText() || '').trim() === 'ok';
    props.setProperty(k, (ok ? 'ok' : 'locked') + '|' + String(Date.now()));
    return ok;
  } catch (e) {
    // 3) On fetch error (e.g., simple trigger auth), fall back to cached status if present; otherwise fail-closed
    if (cached) {
      const status = cached.split('|')[0];
      return status === 'ok';
    }
    return !!FAIL_OPEN; // with FAIL_OPEN=false this will lock on first-run in simple-trigger contexts
  }
}

function isLicensedCachedOnly() {
  const props = PropertiesService.getScriptProperties();
  const k = 'LIC:' + TENANT_ID;
  const cached = props.getProperty(k);
  if (!cached) return null; // unknown
  const parts = cached.split('|');
  const status = parts[0];
  const ts = Number(parts[1] || 0);
  if (Date.now() - ts > LICENSE_CACHE_MS) return null; // stale
  return status === 'ok';
}

function requireLicense(msg) {
  if (isLicensed()) return true;
  SpreadsheetApp.getUi().alert(msg || LOCK_MESSAGE);
  throw new Error('Service locked');
}

function showLocked() {
  SpreadsheetApp.getUi().alert(LOCK_MESSAGE);
}

const SHEET = {
  ORDER_MASTER: 'ORDER_MASTER',
  PRICEBOOK: 'PriceBook',
  CUSTOMERS: 'Customers',
  LOGS: 'Logs',
  ORDER_DATA: '_OrderData'
};

// CLIENT MODE: Set to true for client deployment (simplified menu)
// Set to false for admin/development (full menu access)
const CLIENT_MODE = true;

const STATUS_CHOICES = ['Pending', 'Assigned', 'Shopping', 'Ready for Delivery', 'Out for Delivery', 'Delivered', 'Billed'];
const EXPORT_CHOICES = ['', 'Ready', 'Exported'];
const DRIVE_FOLDER_ID = '';
const CURRENT_SCHEMA_VERSION = 4;
const ORDER_SHEET_PREFIX = 'ORDER_';

// QuickBooks IIF Export Settings
// Account NAMES for IIF export (QuickBooks requires names, not numbers for invoices)
const QB_AR_ACCOUNT = 'Accounts Receivable';  // A/R account name
const QB_INCOME_ACCOUNT = 'Sales Income';  // Income account name (must exist in your Chart of Accounts)

/******** CACHE MANAGER ********/
const CacheManager = (function() {
  const cache = {};
  const CACHE_TTL = 300000;
  
  return {
    get: function(key) {
      const item = cache[key];
      if (!item) return null;
      if (Date.now() - item.timestamp > CACHE_TTL) {
        delete cache[key];
        return null;
      }
      return item.value;
    },
    set: function(key, value) {
      cache[key] = { value: value, timestamp: Date.now() };
    },
    clear: function(key) {
      if (key) {
        delete cache[key];
      } else {
        for (var k in cache) {
          if (cache.hasOwnProperty(k)) delete cache[k];
        }
      }
    }
  };
})();

function clearCache() {
  CacheManager.clear();
  uiToast('✅ Cache cleared');
}

/******** MENU ********/
function onOpen() {
  // Use cached-only status to avoid UrlFetch in simple trigger
  const lic = isLicensedCachedOnly();
  if (lic === false) {
    SpreadsheetApp.getUi()
      .createMenu('⚓ Dupuys')
      .addItem('🔒 Service Locked', 'showLocked')
      .addToUi();
    showLocked(); // immediate clarity
    return;
  }
  // lic is true or unknown -> build menus; actions are still gated by requireLicense()
  if (CLIENT_MODE) {
    buildClientMenu();
  } else {
    buildAdminMenu();
  }
}

/******** CLIENT MENU - Simplified for Daily Use ********/
function buildClientMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('⚓ Dupuys')
    .addItem('📋 Order Master', 'openOrderMaster')
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 QuickBooks Export')
      .addItem('📄 Export Current Order', 'exportCurrentOrderSheet')
      .addItem('📄 Export Ready Batch', 'exportReadyBatch')
      .addSeparator()
      .addItem('📦 Export Items (Optional)', 'exportPriceBookToQuickBooks'))
    .addSubMenu(ui.createMenu('📦 Archive Orders')
      .addItem('Archive Current Order', 'archiveCurrentOrder')
      .addItem('Archive All Exported Orders', 'archiveExported'))
    .addSeparator()
    .addSubMenu(ui.createMenu('👥 Manage Customers')
      .addItem('Add New Customer', 'addCustomerManually')
      .addItem('View PIN Sheet', 'regeneratePinSheet'))
    .addSubMenu(ui.createMenu('🛒 Manage Items')
      .addItem('Add New Item', 'addItemManually'))
    .addSeparator()
    .addItem('🔄 Refresh Data', 'refreshAllDashboards')
    .addItem('🔗 Get Web App URL', 'getWebAppUrl')
    .addSeparator()
    .addItem('ℹ️ Help & Instructions', 'showClientHelp')
    .addToUi();
}

/******** ADMIN MENU - Full Access for Development ********/
function buildAdminMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('⚓ Dupuys Dockside V4 [ADMIN]')
    .addItem('🔧 Initialize System', 'initializeWorkbook')
    .addItem('🌱 Seed Sample Data', 'seedSampleData')
    .addItem('✅ Deployment Checklist', 'runDeploymentChecklist')
    .addSeparator()
    .addSubMenu(ui.createMenu('👥 Customers')
      .addItem('Add Customer Manually', 'addCustomerManually')
      .addItem('Import from QuickBooks CSV', 'importCustomersFromCSV')
      .addItem('Regenerate PIN Sheet', 'regeneratePinSheet'))
    .addSubMenu(ui.createMenu('🛒 Grocery Items')
      .addItem('Add Item Manually', 'addItemManually')
      .addItem('Import Grocery List', 'importGroceryList')
      .addItem('Bulk Update Markup %', 'bulkUpdatePrices'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🌐 Web App')
      .addItem('📝 Deploy Web App Instructions', 'showWebAppDeploymentInstructions')
      .addItem('🔗 Get Web App URL', 'getWebAppUrl')
      .addItem('🔄 Test Web App Connection', 'testWebAppConnection'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Views')
      .addItem('📋 Order Master Index', 'openOrderMaster')
      .addSeparator()
      .addItem('👁️ Show All Order Sheets', 'showAllOrderSheets')
      .addItem('🙈 Hide All Order Sheets', 'hideAllOrderSheets')
      .addSeparator()
      .addItem('📊 Convert to Tables', 'convertAllToTables')
      .addItem('🔄 Refresh All Data', 'refreshAllDashboards')
      .addItem('🔗 Reinstall Edit Sync', 'installOnEditTrigger'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 QuickBooks')
      .addItem('📄 Export Current Order', 'exportCurrentOrderSheet')
      .addItem('📄 Export Ready Batch', 'exportReadyBatch')
      .addSeparator()
      .addItem('📦 Export Items (Optional)', 'exportPriceBookToQuickBooks'))
    .addSubMenu(ui.createMenu('📦 Archive')
      .addItem('Archive Current Order', 'archiveCurrentOrder')
      .addItem('Archive All Exported Orders', 'archiveExported'))
    .addSeparator()
    .addItem('🗑️ Clear Cache', 'clearCache')
    .addItem('🔧 Fix Circular References', 'fixCircularReferenceInOrderSheets')
    .addItem('🔧 Switch to Client Mode', 'switchToClientMode')
    .addToUi();
}

/******** CLIENT HELP ********/
function showClientHelp() {
  const ui = SpreadsheetApp.getUi();
  const webAppUrl = 'Not configured - Ask administrator';
  
  const helpText = '📖 DUPUYS DOCKSIDE - QUICK START GUIDE\n\n' +
    '═══════════════════════════════════\n\n' +
    '📋 DAILY WORKFLOW:\n\n' +
    '1️⃣ View Orders\n' +
    '   • Menu > 📋 Order Master\n' +
    '   • Click "📄 Open Order" links to view details\n' +
    '   • Fill in Base Cost (yellow column) as you shop\n' +
    '   • Update Status as you progress\n\n' +
    '2️⃣ Upload Receipts\n' +
    '   • In order sheet, scroll to Receipt Images section\n' +
    '   • Right-click → Insert → Image → Image in cell\n' +
    '   • Or paste Google Drive links to receipt photos\n\n' +
    '3️⃣ Export to QuickBooks\n' +
    '   • Open an order → Set Export Status = "Ready"\n' +
    '   • Menu > 💰 QuickBooks Export > Export Ready Batch\n' +
    '   • Download CSV/IIF files\n\n' +
    '4️⃣ Archive Orders\n' +
    '   • Menu > 📦 Archive Orders > Archive Current Order\n' +
    '   • Or Archive All Exported Orders\n' +
    '   • Orders saved to "Archived Orders" Drive folder\n\n' +
    '═══════════════════════════════════\n\n' +
    '🌐 WEB APP URL:\n' +
    webAppUrl + '\n\n' +
    'Share this link with boat captains to place orders.\n\n' +
    '═══════════════════════════════════\n\n' +
    '💡 TIPS:\n' +
    '• Orders sync automatically - no manual refresh needed\n' +
    '• Hidden order sheets unhide when you click links\n' +
    '• Use Status dropdown to track progress\n' +
    '• Base Cost column is highlighted yellow\n' +
    '• Upload receipt images immediately after shopping\n\n' +
    '❓ Need Help? Contact your system administrator.';
  
  ui.alert('Help & Instructions', helpText, ui.ButtonSet.OK);
}

/******** PRE-DEPLOYMENT CHECKLIST (Admin Only) ********/
function runDeploymentChecklist() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  let report = '🔍 PRE-DEPLOYMENT CHECKLIST\n\n';
  let allGood = true;
  
  // Check 1: Required sheets exist
  const requiredSheets = [SHEET.ORDER_MASTER, SHEET.PRICEBOOK, SHEET.CUSTOMERS, SHEET.ORDER_DATA];
  let sheetsOk = true;
  for (var i = 0; i < requiredSheets.length; i++) {
    if (!ss.getSheetByName(requiredSheets[i])) {
      sheetsOk = false;
      allGood = false;
    }
  }
  report += sheetsOk ? '✅ All required sheets exist\n' : '❌ Missing required sheets - Run Initialize System\n';
  
  // Check 2: Customers added
  const custSheet = ss.getSheetByName(SHEET.CUSTOMERS);
  const custCount = custSheet ? custSheet.getLastRow() - 1 : 0;
  if (custCount > 0) {
    report += '✅ Customers configured (' + custCount + ' customers)\n';
  } else {
    report += '⚠️ No customers added yet\n';
    allGood = false;
  }
  
  // Check 3: Items added
  const priceSheet = ss.getSheetByName(SHEET.PRICEBOOK);
  const itemCount = priceSheet ? priceSheet.getLastRow() - 1 : 0;
  if (itemCount > 0) {
    report += '✅ Items configured (' + itemCount + ' items)\n';
  } else {
    report += '⚠️ No items added yet\n';
    allGood = false;
  }
  
  // Check 4: Triggers installed
  const triggers = ScriptApp.getProjectTriggers();
  let hasOnEdit = false;
  for (var j = 0; j < triggers.length; j++) {
    if (triggers[j].getHandlerFunction() === 'onEditHandler') {
      hasOnEdit = true;
    }
  }
  report += hasOnEdit ? '✅ Edit sync trigger installed\n' : '⚠️ Edit sync trigger not installed - Run Reinstall Edit Sync\n';
  
  // Check 5: CLIENT_MODE setting
  report += '\n📋 Current Mode: ' + (CLIENT_MODE ? '👥 CLIENT MODE (simplified menu)' : '🔧 ADMIN MODE (full access)') + '\n';
  
  report += '\n═══════════════════════════════\n\n';
  
  if (allGood) {
    report += '🎉 READY FOR DEPLOYMENT!\n\n';
    report += 'Final Steps:\n';
    report += '1. Set CLIENT_MODE = true (if not already)\n';
    report += '2. Test all features\n';
    report += '3. Share with client\n';
  } else {
    report += '⚠️ NEEDS ATTENTION\n\n';
    report += 'Complete the items marked with ❌ or ⚠️ before deploying.';
  }
  
  ui.alert('Deployment Checklist', report, ui.ButtonSet.OK);
}

/******** MODE SWITCHING ********/
function switchToClientMode() {
  SpreadsheetApp.getUi().alert(
    '⚠️ Switch to Client Mode',
    'To switch to Client Mode:\n\n' +
    '1. Open Apps Script Editor (Extensions > Apps Script)\n' +
    '2. Find line: const CLIENT_MODE = false;\n' +
    '3. Change to: const CLIENT_MODE = true;\n' +
    '4. Save (Ctrl+S)\n' +
    '5. Refresh the Google Sheet\n\n' +
    'The menu will show simplified options for daily users.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function switchToAdminMode() {
  SpreadsheetApp.getUi().alert(
    '⚠️ Switch to Admin Mode',
    'To switch to Admin Mode:\n\n' +
    '1. Open Apps Script Editor (Extensions > Apps Script)\n' +
    '2. Find line: const CLIENT_MODE = true;\n' +
    '3. Change to: const CLIENT_MODE = false;\n' +
    '4. Save (Ctrl+S)\n' +
    '5. Refresh the Google Sheet\n\n' +
    'The menu will show all admin/development options.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/******** DATA ACCESS LAYER ********/
const DataLayer = {
  getCustomers: function() {
    const cached = CacheManager.get('customers');
    if (cached) return cached;
    
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.CUSTOMERS);
    if (!sh || sh.getLastRow() < 2) return [];
    
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const customers = vals
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return {
          boatId: String(r[0]).trim(),
          boatName: String(r[1] || '').trim(),
          qbCustomerName: String(r[2] || '').trim(),
          billingEmail: String(r[3] || '').trim(),
          defaultTerms: String(r[4] || '').trim(),
          pin: String(r[5] || '').trim()
        };
      });
    
    CacheManager.set('customers', customers);
    return customers;
  },
  
  getCustomerByBoatId: function(boatId) {
    const customers = this.getCustomers();
    for (var i = 0; i < customers.length; i++) {
      if (customers[i].boatId === boatId) return customers[i];
    }
    return null;
  },
  
  verifyPin: function(boatId, pin) {
    const customer = this.getCustomerByBoatId(boatId);
    if (!customer) return false;
    if (!customer.pin) return true;
    return customer.pin === String(pin).trim();
  },
  
  getPriceBookItems: function() {
    const cached = CacheManager.get('pricebook');
    if (cached) return cached;
    
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.PRICEBOOK);
    if (!sh || sh.getLastRow() < 2) return [];
    
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const items = vals
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return {
          item: String(r[0]).trim(),
          category: String(r[1] || '').trim(),
          unit: String(r[2] || '').trim(),
          basePrice: Number(r[3]) || 0,
          defaultMarkup: Number(r[4]) || 0,
          notes: String(r[5] || '').trim()
        };
      });
    
    CacheManager.set('pricebook', items);
    return items;
  },
  
  getPriceBookItem: function(itemCode) {
    const items = this.getPriceBookItems();
    for (var i = 0; i < items.length; i++) {
      if (items[i].item === itemCode) return items[i];
    }
    return null;
  },
  
  getNextDocNumber: function(boatId) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      
      const props = PropertiesService.getScriptProperties();
      const key = 'COUNTER_' + boatId;
      const current = parseInt(props.getProperty(key) || '0', 10);
      const next = current + 1;
      props.setProperty(key, String(next));
      
      const tz = Session.getScriptTimeZone();
      const ymd = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
      return 'TB-' + boatId + '-' + ymd + '-' + padLeft(next, 4);
    } finally {
      lock.releaseLock();
    }
  }
};

/******** INITIALIZE SYSTEM ********/
function initializeWorkbook() {
  const ss = SpreadsheetApp.getActive();
  
  const orderMaster = ensureSheet(ss, SHEET.ORDER_MASTER);
  const orderData = ensureSheet(ss, SHEET.ORDER_DATA);
  const price = ensureSheet(ss, SHEET.PRICEBOOK);
  const cust = ensureSheet(ss, SHEET.CUSTOMERS);
  const logs = ensureSheet(ss, SHEET.LOGS);
  
  orderData.hideSheet();
  
  const MASTER_COLS = ['Order #', 'DocNumber', 'Date', 'BoatID', 'Boat Name', 'Status', 'Items', 'Total $', 'Assigned To', 'Sheet Link', 'Created', 'Last Updated'];
  const DATA_COLS = ['DocNumber', 'BoatID', 'BoatName', 'Status', 'AssignedTo', 'TxnDate', 'DeliveryLocation', 'Phone', 'Item', 'Category', 'Qty', 'Unit', 'BaseCost', 'Markup%', 'Rate', 'Amount', 'TaxCode', 'Notes', 'ExportStatus', 'CreatedAt'];
  const PRICE_HEADERS = ['Item', 'Category', 'Unit', 'BasePrice', 'DefaultMarkup%', 'Notes'];
  const CUST_HEADERS = ['BoatID', 'BoatName', 'QB_CustomerName', 'BillingEmail', 'DefaultTerms', 'PIN'];
  const LOG_HEADERS = ['Timestamp', 'User', 'Action', 'Details', 'Status'];
  
  initializeSheetHeaders(orderMaster, MASTER_COLS);
  initializeSheetHeaders(orderData, DATA_COLS);
  initializeSheetHeaders(price, PRICE_HEADERS);
  initializeSheetHeaders(cust, CUST_HEADERS);
  initializeSheetHeaders(logs, LOG_HEADERS);
  
  [orderMaster, orderData, price, cust, logs].forEach(function(s) { s.setFrozenRows(1); });
  
  setColumnWidths(orderMaster, 140);
  setColumnWidths(price, 140);
  setColumnWidths(cust, 160);
  
  applyListValidation(orderMaster, 2, 6, STATUS_CHOICES);
  
  // Convert key sheets to Tables for better data management
  convertSheetToTable(orderData, 'OrderDataTable');
  convertSheetToTable(price, 'PriceBookTable');
  convertSheetToTable(cust, 'CustomersTable');
  
  buildOrderMasterSheet(orderMaster);
  
  protectHeaders(orderMaster);
  protectHeaders(price);
  protectHeaders(cust);
  
  // Install trigger
  installOnEditTrigger();
  
  ss.setActiveSheet(orderMaster);
  
  uiToast('✅ System initialized! Next: Add customers → Add items → Deploy web app');
  logAction('Initialize', 'System V4 initialized', 'Success');
}

/******** BUILD ORDER MASTER ********/
function buildOrderMasterSheet(sheet) {
  sheet.clear();
  
  // TITLE ROW - Modern style
  sheet.getRange('A1:L1').merge().setValue('📋 ORDER MASTER INDEX - Click Order Links Below');
  sheet.getRange('A1')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  // INFO ROW - Light blue background
  sheet.getRange('A2:L2').merge().setValue('💡 Click any "📄 Open Order" link to view/edit that order. New orders appear automatically when submitted via form or web app.');
  sheet.getRange('A2')
    .setWrap(true)
    .setBackground('#e8f0fe')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setFontColor('#5f6368');
  
  // STATS ROW - Modern dashboard cards
  sheet.getRange('A4:C4').merge().setBackground('#f1f3f4');
  sheet.getRange('A4').setValue('📊 Total Orders:').setFontWeight('bold');
  sheet.getRange('B4').setFormula('=COUNTA(B6:B)').setFontWeight('bold').setFontSize(14);
  
  sheet.getRange('D4:E4').merge().setBackground('#fef7e0');
  sheet.getRange('D4').setValue('⏳ Pending:').setFontWeight('bold');
  sheet.getRange('E4').setFormula('=COUNTIF(F6:F,"Pending")').setFontWeight('bold').setFontSize(14);
  
  sheet.getRange('G4:H4').merge().setBackground('#e6f4ea');
  sheet.getRange('G4').setValue('🛒 Shopping:').setFontWeight('bold');
  sheet.getRange('H4').setFormula('=COUNTIF(F6:F,"Shopping")').setFontWeight('bold').setFontSize(14);
  
  sheet.getRange('J4:K4').merge().setBackground('#e8f0fe');
  sheet.getRange('J4').setValue('✅ Delivered:').setFontWeight('bold');
  sheet.getRange('K4').setFormula('=COUNTIF(F6:F,"Delivered")').setFontWeight('bold').setFontSize(14);
  
  // TABLE HEADERS - Modern sticky headers
  const headers = ['#', 'DocNumber', 'Date', 'BoatID', 'Boat', 'Status', 'Items', 'Total', 'Assigned', 'Open Order', 'Created', 'Updated'];
  const headerRange = sheet.getRange('A5:L5');
  headerRange
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  
  // Add borders to header
  headerRange.setBorder(
    true, true, true, true, true, true,
    '#ffffff', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // FREEZE ROWS for sticky header effect
  sheet.setFrozenRows(5);
  
  // Set column widths
  sheet.setColumnWidth(1, 50);   // #
  sheet.setColumnWidth(2, 180);  // DocNumber
  sheet.setColumnWidth(3, 120);  // Date
  sheet.setColumnWidth(4, 100);  // BoatID
  sheet.setColumnWidth(5, 150);  // Boat
  sheet.setColumnWidth(6, 130);  // Status
  sheet.setColumnWidth(7, 80);   // Items
  sheet.setColumnWidth(8, 110);  // Total
  sheet.setColumnWidth(9, 120);  // Assigned
  sheet.setColumnWidth(10, 120); // Open Order
  sheet.setColumnWidth(11, 150); // Created
  sheet.setColumnWidth(12, 150); // Updated
  
  // Format currency column
  sheet.getRange('H:H').setNumberFormat('$#,##0.00');
  
  // Apply modern banded rows to data area (starting at row 6)
  try {
    const dataRange = sheet.getRange('A6:L1000');
    const existingBandings = dataRange.getBandings();
    for (var i = 0; i < existingBandings.length; i++) {
      existingBandings[i].remove();
    }
    
    const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.CYAN, false, false);
    banding
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#f8f9fa');
  } catch (e) {
    // Banding optional - continue if it fails
  }
}

/******** CREATE ORDER SHEET ********/
function createOrderSheet(orderInfo) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = ORDER_SHEET_PREFIX + orderInfo.docNumber;
  
  let orderSheet = ss.getSheetByName(sheetName);
  if (orderSheet) {
    logAction('CreateOrder', 'Sheet already exists: ' + sheetName, 'Warning');
    return orderSheet;
  }
  
  orderSheet = ss.insertSheet(sheetName);
  
  // Move sheet to end (far right) instead of beginning
  const allSheets = ss.getSheets();
  ss.setActiveSheet(orderSheet);
  ss.moveActiveSheet(allSheets.length);
  
  buildIndividualOrderSheet(orderSheet, orderInfo);
  addToOrderMaster(orderInfo, sheetName);
  
  // Hide the order sheet by default (access via ORDER_MASTER links)
  orderSheet.hideSheet();
  
  logAction('CreateOrder', 'Created hidden sheet: ' + sheetName, 'Success');
  return orderSheet;
}

/******** BUILD INDIVIDUAL ORDER SHEET ********/
function buildIndividualOrderSheet(sheet, orderInfo) {
  sheet.clear();
  
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  // ========== TITLE SECTION ==========
  sheet.getRange('A1:H1').merge().setValue('🚢 ORDER DETAILS - ' + orderInfo.docNumber);
  sheet.getRange('A1')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  // ========== ORDER INFO SECTION with clean boxes ==========
  // Left column - Boat Info
  sheet.getRange('A3:B3').setBackground('#f8f9fa');
  sheet.getRange('A3').setValue('Doc Number:').setFontWeight('bold');
  sheet.getRange('B3').setValue(orderInfo.docNumber);
  
  sheet.getRange('A4:B4').setBackground('#ffffff');
  sheet.getRange('A4').setValue('Boat ID:').setFontWeight('bold');
  sheet.getRange('B4').setValue(orderInfo.boatId);
  
  sheet.getRange('A5:B5').setBackground('#f8f9fa');
  sheet.getRange('A5').setValue('Boat Name:').setFontWeight('bold');
  sheet.getRange('B5').setValue(orderInfo.boatName);
  
  // Middle column - Status Info
  sheet.getRange('D3:E3').setBackground('#fef7e0');
  sheet.getRange('D3').setValue('Status:').setFontWeight('bold');
  const statusCell = sheet.getRange('E3');
  statusCell.setValue('Pending');
  statusCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS_CHOICES, true).build());
  
  sheet.getRange('D4:E4').setBackground('#e8f0fe');
  sheet.getRange('D4').setValue('Assigned To:').setFontWeight('bold');
  sheet.getRange('E4').setValue('');
  
  sheet.getRange('D5:E5').setBackground('#e6f4ea');
  sheet.getRange('D5').setValue('Delivery Location:').setFontWeight('bold');
  sheet.getRange('E5').setValue(orderInfo.deliveryLocation || '');
  
  // Right column - Date & Customer Info
  sheet.getRange('G3:H3').setBackground('#f8f9fa');
  sheet.getRange('G3').setValue('Order Date:').setFontWeight('bold');
  sheet.getRange('H3').setValue(orderInfo.txnDate);
  
  sheet.getRange('G4:H4').setBackground('#ffffff');
  sheet.getRange('G4').setValue('Phone:').setFontWeight('bold');
  sheet.getRange('H4').setValue(orderInfo.phone || '');
  
  sheet.getRange('G5:H5').setBackground('#f8f9fa');
  sheet.getRange('G5').setValue('QB Customer:').setFontWeight('bold');
  sheet.getRange('H5').setValue(orderInfo.qbCustomerName);
  
  sheet.getRange('G6:H6').setBackground('#ffffff');
  sheet.getRange('G6').setValue('Created:').setFontWeight('bold');
  sheet.getRange('H6').setValue(now);
  
  // Border around info section
  sheet.getRange('A3:H6').setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
  
  // ========== ITEMS TABLE SECTION ==========
  sheet.getRange('A7:H7').merge().setValue('📦 ORDER ITEMS - Fill in Base Cost as you source items');
  sheet.getRange('A7')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('white')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  
  const itemHeaders = ['Item Code', 'Description', 'Category', 'Unit', 'Qty', 'Base Cost', 'Markup %', 'Total'];
  sheet.getRange('A8:H8')
    .setValues([itemHeaders])
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('white')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  
  // Add header borders
  sheet.getRange('A8:H8').setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Items data with alternating colors
  let currentRow = 9;
  for (var i = 0; i < orderInfo.items.length; i++) {
    const item = orderInfo.items[i];
    const priceItem = DataLayer.getPriceBookItem(item.itemCode);
    
    const rowBg = i % 2 === 0 ? '#ffffff' : '#f1f8f4';
    sheet.getRange(currentRow, 1, 1, 8).setBackground(rowBg);
    
    sheet.getRange(currentRow, 1).setValue(item.itemCode);
    sheet.getRange(currentRow, 2).setValue(priceItem ? priceItem.notes : item.itemCode);
    sheet.getRange(currentRow, 3).setValue(item.category);
    sheet.getRange(currentRow, 4).setValue(item.unit);
    sheet.getRange(currentRow, 5).setValue(item.qty);
    sheet.getRange(currentRow, 6).setValue('').setBackground('#fff3cd'); // Highlight for manual entry
    sheet.getRange(currentRow, 7).setValue(priceItem ? priceItem.defaultMarkup : 15);
    sheet.getRange(currentRow, 8).setFormula('=IF(F' + currentRow + '>0, E' + currentRow + '*F' + currentRow + '*(1+G' + currentRow + '/100), "")');
    
    currentRow++;
  }
  
  // Border around items table
  sheet.getRange('A8:H' + (currentRow - 1)).setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
  
  // Totals row with strong styling (using dynamic formulas that adapt to added rows)
  const totalsRow = currentRow + 1;
  sheet.getRange(totalsRow, 1, 1, 4).merge().setValue('💰 TOTAL:').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right').setBackground('#34a853').setFontColor('white');
  // Use SUMIF to dynamically sum all non-empty rows BEFORE the totals row (avoiding circular reference)
  const sumRange = 'A9:A' + (totalsRow - 1);
  const qtyRange = 'E9:E' + (totalsRow - 1);
  const amountRange = 'H9:H' + (totalsRow - 1);
  sheet.getRange(totalsRow, 5).setFormula('=SUMIF(' + sumRange + ',"<>",' + qtyRange + ')').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  sheet.getRange(totalsRow, 6, 1, 2).merge().setBackground('#34a853');
  sheet.getRange(totalsRow, 8).setFormula('=SUMIF(' + sumRange + ',"<>",' + amountRange + ')').setFontWeight('bold').setFontSize(12).setNumberFormat('$#,##0.00').setBackground('#34a853').setFontColor('white');
  sheet.getRange(totalsRow, 1, 1, 8).setBorder(true, true, true, true, false, false, '#34a853', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // ========== NOTES SECTION ==========
  const notesRow = totalsRow + 2;
  sheet.getRange(notesRow, 1, 1, 8).merge().setValue('📝 Notes / Special Instructions:').setFontWeight('bold').setBackground('#e8f0fe').setFontSize(11);
  sheet.getRange(notesRow + 1, 1, 3, 8).merge()
    .setValue(orderInfo.notes || '')
    .setWrap(true)
    .setVerticalAlignment('top')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, false, false, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
  
  // ========== RECEIPT IMAGES SECTION ==========
  const receiptRow = notesRow + 5;
  sheet.getRange(receiptRow, 1, 1, 8).merge().setValue('📸 Receipt Images - Upload or Paste Images Below').setFontWeight('bold').setBackground('#fff3cd').setFontSize(12).setHorizontalAlignment('center');
  
  // Large cell for receipt images with instructions
  const receiptCell = sheet.getRange(receiptRow + 1, 1, 6, 8);
  receiptCell.merge()
    .setValue('📋 INSTRUCTIONS:\n\n' +
      '1. INSERT IMAGE: Right-click here → Insert → Image → Image in cell\n' +
      '2. PASTE DRIVE LINK: Share receipt image from Google Drive and paste link here\n' +
      '3. MULTIPLE RECEIPTS: Insert multiple images or separate links with line breaks\n\n' +
      '💡 TIP: Take photos of receipts immediately after purchase to avoid losing them.')
    .setWrap(true)
    .setVerticalAlignment('top')
    .setBackground('#fffef0')
    .setFontSize(10)
    .setFontColor('#5f6368');
  
  // Make the cell extra tall for images
  for (var rowIdx = receiptRow + 1; rowIdx <= receiptRow + 6; rowIdx++) {
    sheet.setRowHeight(rowIdx, 60);
  }
  
  // Border around receipt section with distinct color
  sheet.getRange(receiptRow, 1, 7, 8).setBorder(true, true, true, true, true, true, '#f9ab00', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // ========== ACTIONS SECTION ==========
  const actionsRow = receiptRow + 8;
  sheet.getRange(actionsRow, 1, 1, 8).merge().setValue('⚙️ Actions & Export').setFontWeight('bold').setFontSize(12).setBackground('#f8f9fa');
  
  sheet.getRange(actionsRow + 1, 1).setValue('Export Status:').setFontWeight('bold').setBackground('#ffffff');
  const exportCell = sheet.getRange(actionsRow + 1, 2, 1, 2);
  exportCell.merge().setValue('').setBackground('#fef7e0');
  exportCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(EXPORT_CHOICES, true).build());
  
  sheet.getRange(actionsRow + 2, 1).setValue('Receipt Link:').setFontWeight('bold').setBackground('#ffffff');
  sheet.getRange(actionsRow + 2, 2, 1, 6).merge().setBackground('#ffffff');
  
  sheet.getRange(actionsRow + 3, 1).setValue('QB Export Link:').setFontWeight('bold').setBackground('#ffffff');
  sheet.getRange(actionsRow + 3, 2, 1, 6).merge().setBackground('#ffffff');
  
  // Border around actions section
  sheet.getRange(actionsRow, 1, 4, 8).setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
  
  // ========== TIP BOX ==========
  sheet.getRange(actionsRow + 5, 1, 2, 8).merge()
    .setValue('💡 TIP: Fill in Base Cost (yellow column) as you source items. Total calculates automatically. Update Status dropdown as you progress. Set Export Status to "Ready" when complete.')
    .setWrap(true)
    .setBackground('#fff3cd')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false, '#f9ab00', SpreadsheetApp.BorderStyle.SOLID);
  
  // ========== FORMATTING ==========
  const lastUsedRow = actionsRow + 6;
  
  // Set column widths
  sheet.setColumnWidth(1, 110);  // Item Code
  sheet.setColumnWidth(2, 200);  // Description
  sheet.setColumnWidth(3, 100);  // Category
  sheet.setColumnWidth(4, 70);   // Unit
  sheet.setColumnWidth(5, 70);   // Qty
  sheet.setColumnWidth(6, 100);  // Base Cost
  sheet.setColumnWidth(7, 90);   // Markup
  sheet.setColumnWidth(8, 110);  // Total
  
  // Currency formatting
  sheet.getRange('F9:F' + (currentRow - 1)).setNumberFormat('$#,##0.00');
  sheet.getRange('H9:H' + (currentRow - 1)).setNumberFormat('$#,##0.00');
  
  // Freeze header rows
  sheet.setFrozenRows(8);
  
  // CLEAN UP: Hide unused rows and columns beyond our content
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  
  if (maxRows > lastUsedRow + 5) {
    sheet.hideRows(lastUsedRow + 1, maxRows - lastUsedRow);
  }
  
  // Hide columns beyond H (except Z which holds metadata)
  if (maxCols > 8) {
    // Hide columns I through Y (9-25)
    if (maxCols >= 25) {
      sheet.hideColumns(9, 17); // I-Y
    } else if (maxCols > 8) {
      sheet.hideColumns(9, maxCols - 8);
    }
  }
  
  // Add outer border around entire used area (creates table effect)
  sheet.getRange(1, 1, lastUsedRow, 8).setBorder(true, true, true, true, null, null, '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // ========== METADATA (hidden in column Z) ==========
  sheet.getRange('Z1').setValue(orderInfo.docNumber);
  sheet.getRange('Z2').setValue(orderInfo.boatId);
  sheet.getRange('Z3').setValue(orderInfo.boatName);
  sheet.getRange('Z4').setValue(currentRow - 1);
  sheet.getRange('Z5').setValue(actionsRow + 1); // Export Status row
  sheet.getRange('Z6').setValue(actionsRow + 3); // QB Export Link row
  
  // Hide the metadata column
  sheet.hideColumns(26); // Column Z
}

/******** ADD TO MASTER INDEX ********/
function addToOrderMaster(orderInfo, sheetName) {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(SHEET.ORDER_MASTER);
  
  const nextRow = master.getLastRow() + 1;
  const orderNum = nextRow - 5;
  
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  const sheetUrl = ss.getUrl() + '#gid=' + ss.getSheetByName(sheetName).getSheetId();
  const linkFormula = '=HYPERLINK("' + sheetUrl + '", "📄 Open Order")';
  
  master.getRange(nextRow, 1).setValue(orderNum);
  master.getRange(nextRow, 2).setValue(orderInfo.docNumber);
  master.getRange(nextRow, 3).setValue(orderInfo.txnDate);
  master.getRange(nextRow, 4).setValue(orderInfo.boatId);
  master.getRange(nextRow, 5).setValue(orderInfo.boatName);
  master.getRange(nextRow, 6).setValue('Pending');
  master.getRange(nextRow, 7).setValue(orderInfo.items.length);
  master.getRange(nextRow, 8).setValue('');
  master.getRange(nextRow, 9).setValue('');
  master.getRange(nextRow, 10).setFormula(linkFormula);
  master.getRange(nextRow, 11).setValue(now);
  master.getRange(nextRow, 12).setValue(now);
}

/******** AUTO-UNHIDE ON SELECTION ********/
/**
 * SIMPLE TRIGGER - Works automatically without installation
 * Automatically unhides order sheets when user clicks "Open Order" link in ORDER_MASTER
 */
function onSelectionChange(e) {
  const licCached = isLicensedCachedOnly();
  if (licCached === false) return; // locked -> quietly no-op
  if (!e || !e.range) return;
  
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    
    // If we're in ORDER_MASTER and clicking the "Open Order" column
    if (sheetName === SHEET.ORDER_MASTER && e.range.getColumn() === 10) {
      const row = e.range.getRow();
      if (row > 5) { // Skip header rows
        const docNumber = sheet.getRange(row, 2).getValue();
        if (docNumber) {
          const ss = SpreadsheetApp.getActive();
          const orderSheet = ss.getSheetByName(ORDER_SHEET_PREFIX + docNumber);
          
          // Auto-unhide the order sheet when link is clicked
          if (orderSheet && orderSheet.isSheetHidden()) {
            orderSheet.showSheet();
            logAction('AutoUnhide', 'Auto-showed sheet for ' + docNumber, 'Info');
          }
        }
      }
    }
  } catch (err) {
    // Silently fail - don't interrupt user workflow
  }
}

/******** SYNC ORDER DATA ********/
function syncOrderToDataSheet(docNumber) {
  const ss = SpreadsheetApp.getActive();
  const orderSheet = ss.getSheetByName(ORDER_SHEET_PREFIX + docNumber);
  const dataSheet = ss.getSheetByName(SHEET.ORDER_DATA);
  
  if (!orderSheet || !dataSheet) return;
  
  const boatId = orderSheet.getRange('Z2').getValue();
  const boatName = orderSheet.getRange('Z3').getValue();
  
  // Calculate dynamic positions based on current sheet content
  const positions = calculateOrderSheetPositions(orderSheet);
  const lastItemRow = positions.lastItemRow;
  const exportStatusRow = positions.exportStatusRow;
  
  const status = orderSheet.getRange('E3').getValue();
  const assignedTo = orderSheet.getRange('E4').getValue();
  const deliveryLocation = orderSheet.getRange('E5').getValue();
  const txnDate = orderSheet.getRange('H3').getValue();
  const phone = orderSheet.getRange('H4').getValue();
  const exportStatus = orderSheet.getRange(exportStatusRow, 2).getValue();
  
  const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const idx = makeHeaderIndex(dataHeaders);
  const allData = dataSheet.getLastRow() > 1 ? dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues() : [];
  
  const filteredData = allData.filter(function(row) {
    return String(row[idx['DocNumber']]).trim() !== docNumber;
  });
  
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  let totalAmount = 0;
  let itemCount = 0; // Track items for THIS order only
  
  for (var currentRow = 9; currentRow <= lastItemRow; currentRow++) {
    const itemCode = orderSheet.getRange(currentRow, 1).getValue();
    if (!itemCode) continue;
    
    const category = orderSheet.getRange(currentRow, 3).getValue();
    const unit = orderSheet.getRange(currentRow, 4).getValue();
    const qty = orderSheet.getRange(currentRow, 5).getValue();
    const baseCost = orderSheet.getRange(currentRow, 6).getValue();
    const markup = orderSheet.getRange(currentRow, 7).getValue();
    const amount = orderSheet.getRange(currentRow, 8).getValue();
    
    totalAmount += Number(amount || 0);
    itemCount++; // Count each item for THIS order
    
    const row = new Array(dataHeaders.length).fill('');
    row[idx['DocNumber']] = docNumber;
    row[idx['BoatID']] = boatId;
    row[idx['BoatName']] = boatName;
    row[idx['Status']] = status;
    row[idx['AssignedTo']] = assignedTo || '';
    row[idx['TxnDate']] = txnDate;
    row[idx['DeliveryLocation']] = deliveryLocation;
    row[idx['Phone']] = phone || '';
    row[idx['Item']] = itemCode;
    row[idx['Category']] = category;
    row[idx['Qty']] = qty;
    row[idx['Unit']] = unit;
    row[idx['BaseCost']] = baseCost;
    row[idx['Markup%']] = markup;
    row[idx['Rate']] = baseCost ? baseCost * (1 + markup / 100) : '';
    row[idx['Amount']] = amount;
    row[idx['TaxCode']] = 'NON';
    row[idx['Notes']] = '';
    row[idx['ExportStatus']] = exportStatus;
    row[idx['CreatedAt']] = now;
    
    filteredData.push(row);
  }
  
  dataSheet.clearContents();
  dataSheet.getRange(1, 1, 1, dataHeaders.length).setValues([dataHeaders]);
  if (filteredData.length > 0) {
    dataSheet.getRange(2, 1, filteredData.length, dataHeaders.length).setValues(filteredData);
  }
  
  updateMasterIndex(docNumber, status, itemCount, totalAmount, assignedTo);
}

/******** UPDATE MASTER INDEX ********/
function updateMasterIndex(docNumber, status, itemCount, total, assignedTo) {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(SHEET.ORDER_MASTER);
  const orderSheet = ss.getSheetByName(ORDER_SHEET_PREFIX + docNumber);
  
  const data = master.getLastRow() > 5 ? master.getRange(6, 1, master.getLastRow() - 5, master.getLastColumn()).getValues() : [];
  
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim() === docNumber) {
      const row = i + 6;
      const tz = Session.getScriptTimeZone();
      const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
      
      // Get order date from order sheet
      const orderDate = orderSheet ? orderSheet.getRange('H3').getValue() : data[i][2];
      
      master.getRange(row, 3).setValue(orderDate); // Date column
      master.getRange(row, 6).setValue(status); // Status column
      master.getRange(row, 7).setValue(itemCount); // Items column
      master.getRange(row, 8).setValue(total); // Total column
      master.getRange(row, 9).setValue(assignedTo || ''); // Assigned To column
      master.getRange(row, 12).setValue(now); // Last Updated column
      break;
    }
  }
}

/******** FORM SUBMISSION ********/
function onFormSubmit(e) {
  requireLicense('🔒 Submissions are disabled because this account is past due. If this is a misunderstanding, contact Subterra at ' + SUPPORT_CONTACT + '.');
  try {
    const named = e.namedValues || {};
    
    const boatRaw = first(named['Boat (BoatID)']);
    const pin = first(named['PIN']);
    const deliveryLocation = first(named['Delivery Location']) || first(named['Delivery Dock / Location']) || '';
    const phone = first(named['Phone Number']) || '';
    const reqDate = first(named['Date']) || first(named['Requested Delivery Date']);
    const notes = first(named['Notes / Special Instructions']) || '';
    const additionalNotes = first(named['Additional Notes or Substitutions']) || '';
    const finalNotes = [notes, additionalNotes].filter(Boolean).join(' | ');
    
    if (!boatRaw || !pin) throw new Error('Missing boat or PIN');
    
    const parts = String(boatRaw).split('—');
    const boatId = parts[0].trim();
    
    if (!DataLayer.verifyPin(boatId, pin)) {
      logAction('FormAuth', 'Failed PIN for ' + boatId, 'Failed');
      throw new Error('Invalid PIN for ' + boatId);
    }
    
    const customer = DataLayer.getCustomerByBoatId(boatId);
    if (!customer) throw new Error('Customer not found');
    
    const docNumber = DataLayer.getNextDocNumber(boatId);
    const tz = Session.getScriptTimeZone();
    const txnDate = normalizeDateYMD(reqDate) || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    
    const allItems = DataLayer.getPriceBookItems();
    const items = [];
    
    for (var i = 0; i < allItems.length; i++) {
      const itemCode = allItems[i].item;
      const qtyStr = first(named[itemCode]);
      if (!qtyStr) continue;
      
      const qty = parseFloat(String(qtyStr).trim());
      if (!isFinite(qty) || qty <= 0) continue;
      
      items.push({
        itemCode: itemCode,
        category: allItems[i].category,
        unit: allItems[i].unit,
        qty: qty
      });
    }
    
    if (items.length === 0) throw new Error('No items with quantities entered');
    
    const orderInfo = {
      docNumber: docNumber,
      boatId: boatId,
      boatName: customer.boatName,
      qbCustomerName: customer.qbCustomerName || customer.boatName,
      txnDate: txnDate,
      deliveryLocation: deliveryLocation,
      phone: phone,
      notes: finalNotes,
      items: items
    };
    
    createOrderSheet(orderInfo);
    syncOrderToDataSheet(docNumber);
    
    logAction('FormSubmit', 'Created order ' + docNumber, 'Success');
    uiToast('✅ Order created: ' + docNumber + ' (' + items.length + ' items)');
    
  } catch (err) {
    logAction('FormError', String(err), 'Failed');
    uiToast('❌ Order failed: ' + String(err));
  }
}

/******** WEB APP DEPLOYMENT INFO ********/
function showWebAppDeploymentInstructions() {
  requireLicense();
  const ui = SpreadsheetApp.getUi();
  
  const instructions = 'WEB APP DEPLOYMENT INSTRUCTIONS:\n\n' +
    '1. In Apps Script editor, click "Deploy" > "New deployment"\n' +
    '2. Click gear icon ⚙️ next to "Select type"\n' +
    '3. Choose "Web app"\n' +
    '4. Settings:\n' +
    '   - Description: Dupuys Order Form\n' +
    '   - Execute as: Me\n' +
    '   - Who has access: Anyone\n' +
    '5. Click "Deploy"\n' +
    '6. Copy the Web app URL\n' +
    '7. Share it with boat captains\n\n' +
    'Note: Store the URL somewhere safe for reference.';
  
  ui.alert('Web App Deployment', instructions, ui.ButtonSet.OK);
}

function getWebAppUrl() {
  requireLicense();
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Web App URL',
    'To get your Web App URL:\n\n' +
    '1. Extensions > Apps Script\n' +
    '2. Click "Deploy" > "Manage deployments"\n' +
    '3. Click on the active web app deployment\n' +
    '4. Copy the "Web app URL"\n\n' +
    'Share that URL with boat captains to place orders.',
    ui.ButtonSet.OK
  );
}

/******** FIX CIRCULAR REFERENCE IN ORDER SHEETS ********/
function fixCircularReferenceInOrderSheets() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  let fixedCount = 0;
  
  for (var i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    
    // Only process order sheets
    if (sheetName.indexOf(ORDER_SHEET_PREFIX) !== 0) continue;
    
    try {
      // Find the totals row
      const positions = calculateOrderSheetPositions(sheet);
      const totalsRow = positions.totalsRow;
      
      // Fix the formulas to exclude the totals row itself
      const sumRange = 'A9:A' + (totalsRow - 1);
      const qtyRange = 'E9:E' + (totalsRow - 1);
      const amountRange = 'H9:H' + (totalsRow - 1);
      
      sheet.getRange(totalsRow, 5).setFormula('=SUMIF(' + sumRange + ',"<>",' + qtyRange + ')');
      sheet.getRange(totalsRow, 8).setFormula('=SUMIF(' + sumRange + ',"<>",' + amountRange + ')');
      
      fixedCount++;
      
    } catch (err) {
      logAction('FixCircular', 'Failed to fix ' + sheetName + ': ' + String(err), 'Warning');
    }
  }
  
  uiToast('✅ Fixed ' + fixedCount + ' order sheet(s)');
  SpreadsheetApp.getUi().alert(
    '✅ Circular Reference Fixed',
    'Updated ' + fixedCount + ' order sheet(s).\n\n' +
    'The totals formulas have been corrected to avoid circular dependencies.\n\n' +
    'New orders will automatically use the correct formula.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/******** NAVIGATION ********/
function openOrderMaster() {
  SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(SHEET.ORDER_MASTER));
}

function refreshAllDashboards() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    if (sheetName.indexOf(ORDER_SHEET_PREFIX) === 0) {
      const docNumber = sheetName.replace(ORDER_SHEET_PREFIX, '');
      syncOrderToDataSheet(docNumber);
    }
  }
  
  uiToast('✅ All data refreshed');
}

/******** SHEET VISIBILITY MANAGEMENT ********/
function showAllOrderSheets() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  var count = 0;
  
  for (var i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    if (sheetName.indexOf(ORDER_SHEET_PREFIX) === 0) {
      if (sheets[i].isSheetHidden()) {
        sheets[i].showSheet();
        count++;
      }
    }
  }
  
  uiToast('✅ Showed ' + count + ' order sheet(s)');
  logAction('ShowSheets', 'Showed ' + count + ' order sheets', 'Success');
}

function hideAllOrderSheets() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  var count = 0;
  
  for (var i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    if (sheetName.indexOf(ORDER_SHEET_PREFIX) === 0) {
      if (!sheets[i].isSheetHidden()) {
        sheets[i].hideSheet();
        count++;
      }
    }
  }
  
  uiToast('✅ Hid ' + count + ' order sheet(s). Access via ORDER_MASTER links.');
  logAction('HideSheets', 'Hid ' + count + ' order sheets', 'Success');
}

function convertAllToTables() {
  const ss = SpreadsheetApp.getActive();
  
  const orderData = ss.getSheetByName(SHEET.ORDER_DATA);
  const price = ss.getSheetByName(SHEET.PRICEBOOK);
  const cust = ss.getSheetByName(SHEET.CUSTOMERS);
  const logs = ss.getSheetByName(SHEET.LOGS);
  
  convertSheetToTable(orderData, 'OrderDataTable');
  convertSheetToTable(price, 'PriceBookTable');
  convertSheetToTable(cust, 'CustomersTable');
  convertSheetToTable(logs, 'LogsTable');
  
  uiToast('✅ Converted data sheets to table format with filters and banding');
  logAction('ConvertTables', 'Converted sheets to table format', 'Success');
}

/******** CUSTOMER MANAGEMENT ********/
function addCustomerManually() {
  const ui = SpreadsheetApp.getUi();
  
  const boatNameResp = ui.prompt('Add Customer', 'Enter Boat/Company Name:', ui.ButtonSet.OK_CANCEL);
  if (boatNameResp.getSelectedButton() !== ui.Button.OK) return;
  const boatName = boatNameResp.getResponseText().trim();
  
  const qbNameResp = ui.prompt('Add Customer', 'Enter QuickBooks Customer Name:', ui.ButtonSet.OK_CANCEL);
  if (qbNameResp.getSelectedButton() !== ui.Button.OK) return;
  const qbName = qbNameResp.getResponseText().trim();
  
  const emailResp = ui.prompt('Add Customer', 'Enter Billing Email:', ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  const email = emailResp.getResponseText().trim();
  
  const pinResp = ui.prompt('Add Customer', 'Enter 4-6 digit PIN:', ui.ButtonSet.OK_CANCEL);
  if (pinResp.getSelectedButton() !== ui.Button.OK) return;
  const pin = pinResp.getResponseText().trim();
  
  if (!boatName || !qbName || !pin) {
    ui.alert('Error: All fields required');
    return;
  }
  
  const existingIds = getExistingBoatIds();
  const boatId = generateBoatId(boatName, existingIds);
  
  const ss = SpreadsheetApp.getActive();
  const custSheet = ss.getSheetByName(SHEET.CUSTOMERS);
  custSheet.appendRow([boatId, boatName, qbName, email, 'Net 7', pin]);
  
  CacheManager.clear();
  ui.alert('✅ Customer Added', 'BoatID: ' + boatId + '\nBoat: ' + boatName + '\nPIN: ' + pin, ui.ButtonSet.OK);
  logAction('CustomerAdd', 'Added ' + boatName, 'Success');
}

function generateBoatId(name, existingIds) {
  const cleaned = name.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  const prefix = cleaned.substring(0, 4);
  var counter = 1;
  var boatId = prefix + padLeft(counter, 3);
  while (existingIds.includes(boatId)) {
    counter++;
    boatId = prefix + padLeft(counter, 3);
  }
  return boatId;
}

function getExistingBoatIds() {
  const ss = SpreadsheetApp.getActive();
  const custSheet = ss.getSheetByName(SHEET.CUSTOMERS);
  if (!custSheet || custSheet.getLastRow() < 2) return [];
  const values = custSheet.getRange(2, 1, custSheet.getLastRow() - 1, 1).getValues();
  return values.map(function(row) { return String(row[0]).trim(); }).filter(Boolean);
}

function regeneratePinSheet() {
  const ss = SpreadsheetApp.getActive();
  const custSheet = ss.getSheetByName(SHEET.CUSTOMERS);
  if (!custSheet || custSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No customers');
    return;
  }
  const data = custSheet.getRange(2, 1, custSheet.getLastRow() - 1, 6).getValues();
  const oldSheet = ss.getSheetByName('Customer_PINs');
  if (oldSheet) ss.deleteSheet(oldSheet);
  const pinSheet = ss.insertSheet('Customer_PINs');
  pinSheet.getRange(1, 1, 1, 3).setValues([['BoatID', 'Boat Name', 'PIN']]).setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
  const pinData = data.map(function(c) { return [c[0], c[1], c[5]]; });
  pinSheet.getRange(2, 1, pinData.length, 3).setValues(pinData);
  pinSheet.setColumnWidths(1, 3, 150);
  SpreadsheetApp.getUi().alert('✅ PIN sheet regenerated');
}

/******** ITEM MANAGEMENT ********/
function addItemManually() {
  const ui = SpreadsheetApp.getUi();
  
  const codeResp = ui.prompt('Add Item', 'Enter Item Code (e.g., MILK-2PCT):', ui.ButtonSet.OK_CANCEL);
  if (codeResp.getSelectedButton() !== ui.Button.OK) return;
  const code = codeResp.getResponseText().trim().toUpperCase();
  
  const catResp = ui.prompt('Add Item', 'Enter Category (e.g., Dairy):', ui.ButtonSet.OK_CANCEL);
  if (catResp.getSelectedButton() !== ui.Button.OK) return;
  const category = catResp.getResponseText().trim();
  
  const unitResp = ui.prompt('Add Item', 'Enter Unit (e.g., gallon, dozen, lb):', ui.ButtonSet.OK_CANCEL);
  if (unitResp.getSelectedButton() !== ui.Button.OK) return;
  const unit = unitResp.getResponseText().trim();
  
  const priceResp = ui.prompt('Add Item', 'Enter Base Price (optional):', ui.ButtonSet.OK_CANCEL);
  if (priceResp.getSelectedButton() !== ui.Button.OK) return;
  const price = parseFloat(priceResp.getResponseText()) || 0;
  
  const markupResp = ui.prompt('Add Item', 'Enter Default Markup % (e.g., 15):', ui.ButtonSet.OK_CANCEL);
  if (markupResp.getSelectedButton() !== ui.Button.OK) return;
  const markup = parseFloat(markupResp.getResponseText()) || 15;
  
  const notesResp = ui.prompt('Add Item', 'Enter Description/Notes:', ui.ButtonSet.OK_CANCEL);
  if (notesResp.getSelectedButton() !== ui.Button.OK) return;
  const notes = notesResp.getResponseText().trim();
  
  if (!code) {
    ui.alert('Error: Item Code required');
    return;
  }
  
  const ss = SpreadsheetApp.getActive();
  const priceSheet = ss.getSheetByName(SHEET.PRICEBOOK);
  priceSheet.appendRow([code, category, unit, price, markup, notes]);
  
  CacheManager.clear();
  ui.alert('✅ Item Added', 'Code: ' + code + '\nCategory: ' + category, ui.ButtonSet.OK);
  logAction('ItemAdd', 'Added ' + code, 'Success');
}

function seedSampleData() {
  const ss = SpreadsheetApp.getActive();
  const price = ss.getSheetByName(SHEET.PRICEBOOK);
  const cust = ss.getSheetByName(SHEET.CUSTOMERS);
  
  if (cust.getLastRow() < 2) {
    const boats = [
      ['B001', 'Boat Alpha', 'Boat Alpha LLC', 'alpha@fleet.example', 'Net 7', '1234'],
      ['B002', 'Boat Bravo', 'Boat Bravo Inc', 'bravo@fleet.example', 'Net 7', '5678'],
      ['B003', 'Boat Charlie', 'Boat Charlie Co', 'charlie@fleet.example', 'Net 7', '9012']
    ];
    cust.getRange(2, 1, boats.length, boats[0].length).setValues(boats);
  }
  
  if (price.getLastRow() < 2) {
    const items = [
      ['MILK-2PCT', 'Dairy', 'gallon', 4.25, 15, 'Milk 2%'],
      ['EGGS-DOZ', 'Dairy', 'dozen', 3.60, 12, 'Eggs dozen'],
      ['BREAD-LOAF', 'Bakery', 'loaf', 2.90, 12, 'White bread'],
      ['RICE-5LB', 'Staples', 'bag', 6.75, 10, 'Rice 5lb'],
      ['CHICKEN-5LB', 'Meat', 'pack', 13.50, 18, 'Chicken thighs'],
      ['WATER-CASE', 'Beverage', 'case', 5.95, 15, '24pk water'],
      ['COFFEE-2LB', 'Beverage', 'bag', 13.90, 18, 'Ground coffee']
    ];
    price.getRange(2, 1, items.length, items[0].length).setValues(items);
  }
  
  CacheManager.clear();
  uiToast('✅ Sample data seeded. PINs: B001=1234, B002=5678, B003=9012');
}

/******** INSTALL ON EDIT TRIGGER ********/
function installOnEditTrigger() {
  requireLicense();
  const triggers = ScriptApp.getProjectTriggers();
  
  // Remove existing onEdit triggers
  for (var i = 0; i < triggers.length; i++) {
    const funcName = triggers[i].getHandlerFunction();
    if (funcName === 'onEditHandler') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Install onEdit trigger for bidirectional sync
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  
  // Note: onSelectionChange is a simple trigger and works automatically (no installation needed)
  
  uiToast('✅ Edit sync trigger installed (auto-unhide works automatically)');
  logAction('Trigger', 'Installed onEdit handler', 'Success');
}

/******** ON EDIT HANDLER - BIDIRECTIONAL SYNC ********/
/**
 * AUTOMATIC SYNC BETWEEN ORDER_MASTER AND ORDER SHEETS
 * 
 * ORDER_MASTER → Order Sheet:
 *   - Date (Col 3) → H3 (Order Date)
 *   - Status (Col 6) → E3 (Status dropdown)
 *   - Assigned (Col 9) → E4 (Assigned To)
 * 
 * Order Sheet → ORDER_MASTER + _OrderData:
 *   - E3 (Status) → Col 6 + syncs full order data
 *   - E4 (Assigned To) → Col 9 + syncs full order data
 *   - E5 (Delivery Location) → syncs to _OrderData
 *   - H3 (Order Date) → Col 3 + syncs to _OrderData
 *   - H4 (Phone) → syncs to _OrderData
 *   - F (Base Cost) → recalculates total, syncs to _OrderData
 *   - E (Qty) → recalculates total, syncs to _OrderData
 *   - G (Markup%) → recalculates total, syncs to _OrderData
 *   - Export Status → syncs to _OrderData
 * 
 * All syncs update "Last Updated" timestamp automatically
 */
function onEditHandler(e) {
  if (!isLicensed()) return; // silently no-op when locked
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Handle ORDER_MASTER edits
  if (sheetName === SHEET.ORDER_MASTER && row > 5) {
    handleOrderMasterEdit(sheet, row, col, e.value);
    return;
  }
  
  // Handle individual order sheet edits
  if (sheetName.indexOf(ORDER_SHEET_PREFIX) === 0) {
    handleOrderSheetEdit(sheet, sheetName, row, col, e.value);
    return;
  }
}

/******** HANDLE ORDER_MASTER EDITS ********/
function handleOrderMasterEdit(masterSheet, row, col, newValue) {
  try {
    const headers = masterSheet.getRange(5, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    const idx = makeHeaderIndex(headers);
    
    const docNumber = String(masterSheet.getRange(row, idx['DocNumber'] + 1).getValue()).trim();
    if (!docNumber) return;
    
    const ss = SpreadsheetApp.getActive();
    const orderSheet = ss.getSheetByName(ORDER_SHEET_PREFIX + docNumber);
    if (!orderSheet) return;
    
    const colName = headers[col - 1];
    const tz = Session.getScriptTimeZone();
    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    
    // Sync Date (Column 3 in master = C) 
    if (colName === 'Date') {
      orderSheet.getRange('H3').setValue(newValue || '');
      masterSheet.getRange(row, idx['Updated'] + 1).setValue(now);
      syncOrderToDataSheet(docNumber);
      logAction('Sync', 'Master→Sheet: Date for ' + docNumber, 'Success');
    }
    
    // Sync Status (Column 6 in master = F)
    if (colName === 'Status') {
      orderSheet.getRange('E3').setValue(newValue || 'Pending');
      masterSheet.getRange(row, idx['Updated'] + 1).setValue(now);
      syncOrderToDataSheet(docNumber);
      logAction('Sync', 'Master→Sheet: Status for ' + docNumber, 'Success');
    }
    
    // Sync Assigned To (Column 9 in master = I)
    if (colName === 'Assigned') {
      orderSheet.getRange('E4').setValue(newValue || '');
      masterSheet.getRange(row, idx['Updated'] + 1).setValue(now);
      syncOrderToDataSheet(docNumber);
      logAction('Sync', 'Master→Sheet: AssignedTo for ' + docNumber, 'Success');
    }
    
  } catch (err) {
    logAction('SyncError', 'handleOrderMasterEdit: ' + String(err), 'Failed');
  }
}

/******** HANDLE ORDER SHEET EDITS ********/
function handleOrderSheetEdit(orderSheet, sheetName, row, col, newValue) {
  try {
    const docNumber = sheetName.replace(ORDER_SHEET_PREFIX, '');
    const ss = SpreadsheetApp.getActive();
    const masterSheet = ss.getSheetByName(SHEET.ORDER_MASTER);
    if (!masterSheet) return;
    
    const tz = Session.getScriptTimeZone();
    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    
    // Calculate dynamic positions based on current sheet content
    const positions = calculateOrderSheetPositions(orderSheet);
    const exportStatusRow = positions.exportStatusRow;
    
    let shouldSync = false;
    
    // Status changed (E3)
    if (row === 3 && col === 5) {
      shouldSync = true;
      updateMasterField(masterSheet, docNumber, 'Status', newValue || 'Pending');
      logAction('Sync', 'Sheet→Master: Status for ' + docNumber, 'Success');
    }
    
    // Assigned To changed (E4)
    if (row === 4 && col === 5) {
      shouldSync = true;
      updateMasterField(masterSheet, docNumber, 'Assigned', newValue || '');
      logAction('Sync', 'Sheet→Master: AssignedTo for ' + docNumber, 'Success');
    }
    
    // Export Status changed (dynamic row, column B)
    if (row === exportStatusRow && col === 2) {
      shouldSync = true;
      logAction('Sync', 'Sheet→Data: ExportStatus for ' + docNumber, 'Success');
    }
    
    // Any item data changed (rows 9+, columns E=Qty, F=Base Cost, G=Markup%, H=Total)
    if (row >= 9 && (col === 5 || col === 6 || col === 7 || col === 8)) {
      shouldSync = true;
      logAction('Sync', 'Sheet→Master: Item data updated for ' + docNumber, 'Info');
    }
    
    // Order Date changed (H3)
    if (row === 3 && col === 8) {
      shouldSync = true;
      logAction('Sync', 'Sheet→Master: Order date for ' + docNumber, 'Info');
    }
    
    // Phone changed (H4)
    if (row === 4 && col === 8) {
      shouldSync = true;
      logAction('Sync', 'Sheet→Data: Phone for ' + docNumber, 'Info');
    }
    
    // Delivery Location changed (E5)
    if (row === 5 && col === 5) {
      shouldSync = true;
      logAction('Sync', 'Sheet→Data: Delivery location for ' + docNumber, 'Info');
    }
    
    if (shouldSync) {
      syncOrderToDataSheet(docNumber);
    }
    
  } catch (err) {
    logAction('SyncError', 'handleOrderSheetEdit: ' + String(err), 'Failed');
  }
}

/******** UPDATE MASTER FIELD ********/
function updateMasterField(masterSheet, docNumber, fieldName, value) {
  const headers = masterSheet.getRange(5, 1, 1, masterSheet.getLastColumn()).getValues()[0];
  const idx = makeHeaderIndex(headers);
  
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 6) return;
  
  const data = masterSheet.getRange(6, 1, lastRow - 5, masterSheet.getLastColumn()).getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idx['DocNumber']]).trim() === docNumber) {
      const row = i + 6;
      const tz = Session.getScriptTimeZone();
      const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
      
      masterSheet.getRange(row, idx[fieldName] + 1).setValue(value);
      masterSheet.getRange(row, idx['Updated'] + 1).setValue(now);
      break;
    }
  }
}

/******** EXPORT CURRENT ORDER SHEET ********/
function exportCurrentOrderSheet() {
  requireLicense();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (sheetName.indexOf(ORDER_SHEET_PREFIX) !== 0) {
    SpreadsheetApp.getUi().alert('⚠️ Please open an order sheet first (ORDER_TB-...)');
    return;
  }
  
  const docNumber = sheetName.replace(ORDER_SHEET_PREFIX, '');
  
  // Calculate dynamic positions based on current sheet content
  const positions = calculateOrderSheetPositions(sheet);
  const lastItemRow = positions.lastItemRow;
  
  // Check if order has items with costs
  let hasCosts = false;
  
  for (var row = 9; row <= lastItemRow; row++) {
    const itemCode = sheet.getRange(row, 1).getValue();
    const baseCost = sheet.getRange(row, 6).getValue(); // Column F (Base Cost)
    if (itemCode && baseCost > 0) {
      hasCosts = true;
      break;
    }
  }
  
  if (!hasCosts) {
    SpreadsheetApp.getUi().alert('⚠️ Please fill in Base Cost (column F) for items before exporting.\n\nYou need to enter the actual cost you paid for each item.');
    return;
  }
  
  exportOrders('single', docNumber);
}

/******** EXPORT READY BATCH ********/
function exportReadyBatch() {
  requireLicense('🔒 Exports disabled. Account is past due. If this is a misunderstanding, contact Subterra at ' + SUPPORT_CONTACT + '.');
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ss.getSheetByName(SHEET.ORDER_DATA);
  
  if (!dataSheet || dataSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('⚠️ No orders to export');
    return;
  }
  
  const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const idx = makeHeaderIndex(headers);
  
  const readyOrders = {};
  for (var i = 0; i < data.length; i++) {
    const exportStatus = String(data[i][idx['ExportStatus']]).trim();
    if (exportStatus === 'Ready') {
      const docNum = String(data[i][idx['DocNumber']]).trim();
      readyOrders[docNum] = true;
    }
  }
  
  const orderCount = Object.keys(readyOrders).length;
  if (orderCount === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No orders with Export Status = "Ready"\n\nOpen order sheets and set Export Status to "Ready" first.');
    return;
  }
  
  const response = SpreadsheetApp.getUi().alert(
    'Export Ready Batch',
    'Found ' + orderCount + ' order(s) ready to export.\n\nContinue?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response !== SpreadsheetApp.getUi().Button.YES) return;
  
  exportOrders('batch', null);
}

/******** MAIN EXPORT FUNCTION ********/
function exportOrders(mode, singleDocNumber) {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ss.getSheetByName(SHEET.ORDER_DATA);
  
  if (!dataSheet || dataSheet.getLastRow() < 2) {
    uiToast('❌ No order data found');
    return;
  }
  
  const allData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const idx = makeHeaderIndex(headers);
  
  // Filter lines to export
  const linesToExport = [];
  for (var i = 0; i < allData.length; i++) {
    const row = allData[i];
    const docNum = String(row[idx['DocNumber']]).trim();
    const exportStatus = String(row[idx['ExportStatus']]).trim();
    const baseCost = Number(row[idx['BaseCost']]) || 0;
    
    if (!docNum) continue;
    
    // Skip items without costs
    if (baseCost <= 0) continue;
    
    if (mode === 'single') {
      if (docNum === singleDocNumber) {
        linesToExport.push({
          docNumber: docNum,
          boatId: String(row[idx['BoatID']]).trim(),
          boatName: String(row[idx['BoatName']]).trim(),
          status: String(row[idx['Status']]).trim(),
          txnDate: String(row[idx['TxnDate']]).trim(),
          deliveryLocation: String(row[idx['DeliveryLocation']]).trim(),
          item: String(row[idx['Item']]).trim(),
          qty: Number(row[idx['Qty']]) || 0,
          baseCost: baseCost,
          markup: Number(row[idx['Markup%']]) || 0,
          rate: Number(row[idx['Rate']]) || 0,
          amount: Number(row[idx['Amount']]) || 0,
          taxCode: String(row[idx['TaxCode']]) || 'NON',
          notes: String(row[idx['Notes']]).trim()
        });
      }
    } else { // batch
      if (exportStatus === 'Ready') {
        linesToExport.push({
          docNumber: docNum,
          boatId: String(row[idx['BoatID']]).trim(),
          boatName: String(row[idx['BoatName']]).trim(),
          status: String(row[idx['Status']]).trim(),
          txnDate: String(row[idx['TxnDate']]).trim(),
          deliveryLocation: String(row[idx['DeliveryLocation']]).trim(),
          item: String(row[idx['Item']]).trim(),
          qty: Number(row[idx['Qty']]) || 0,
          baseCost: baseCost,
          markup: Number(row[idx['Markup%']]) || 0,
          rate: Number(row[idx['Rate']]) || 0,
          amount: Number(row[idx['Amount']]) || 0,
          taxCode: String(row[idx['TaxCode']]) || 'NON',
          notes: String(row[idx['Notes']]).trim()
        });
      }
    }
  }
  
  if (linesToExport.length === 0) {
    uiToast('❌ No items with costs found to export');
    return;
  }
  
  // Group by DocNumber
  const invoicesByDoc = {};
  for (var j = 0; j < linesToExport.length; j++) {
    const line = linesToExport[j];
    const docNum = line.docNumber;
    
    if (!invoicesByDoc[docNum]) {
      // Get customer info
      const customer = DataLayer.getCustomerByBoatId(line.boatId);
      
      invoicesByDoc[docNum] = {
        docNumber: docNum,
        customer: customer ? customer.qbCustomerName : line.boatName,
        txnDate: line.txnDate,
        terms: customer ? customer.defaultTerms : 'Net 7',
        memo: line.notes,
        class: line.deliveryLocation,
        lines: []
      };
    }
    
    // Sanitize line: force positive Qty/Rate, recompute Amount
    const _qty  = Math.abs(Number(line.qty)  || 0);
    const _rate = Math.abs(Number(line.rate) || 0);
    const _amt  = round2(_qty * _rate);

    invoicesByDoc[docNum].lines.push({
      item: line.item,
      description: line.item,
      qty: _qty,
      rate: _rate,
      amount: _amt,
      taxCode: line.taxCode
    });
  }
  
  const invoices = Object.values(invoicesByDoc);
  
  // Generate export files
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  const qboCsv = buildQboCsv(invoices);
  const qbdIif = buildQbdIif(invoices);
  
  const folder = getDriveFolder();
  const qboFile = folder.createFile('Dupuys_QBO_' + mode + '_' + timestamp + '.csv', qboCsv, MimeType.CSV);
  const qbdFile = folder.createFile('Dupuys_QBD_' + mode + '_' + timestamp + '.iif', qbdIif, MimeType.PLAIN_TEXT);
  
  const qboUrl = qboFile.getUrl();
  const qbdUrl = qbdFile.getUrl();
  
  // Update order sheets
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  
  for (var k = 0; k < invoices.length; k++) {
    const invoice = invoices[k];
    const orderSheet = ss.getSheetByName(ORDER_SHEET_PREFIX + invoice.docNumber);
    
    if (orderSheet) {
      // Calculate dynamic row positions based on current sheet content
      const positions = calculateOrderSheetPositions(orderSheet);
      
      // Update QB Export Link (just set value, cells already merged)
      orderSheet.getRange(positions.qbExportLinkRow, 2).setValue(qboUrl);
      
      // Update Export Status to "Exported" (just set value, cells already merged)
      orderSheet.getRange(positions.exportStatusRow, 2).setValue('Exported');
      
      // Add export note
      orderSheet.getRange(positions.qbExportLinkRow + 1, 1).setValue('Last Exported:').setFontWeight('bold');
      orderSheet.getRange(positions.qbExportLinkRow + 1, 2).setValue(now);
    }
  }
  
  // Refresh data sheet
  for (var m = 0; m < invoices.length; m++) {
    syncOrderToDataSheet(invoices[m].docNumber);
  }
  
  logAction('Export', mode + ': ' + invoices.length + ' invoice(s), ' + linesToExport.length + ' line(s)', 'Success');
  
  SpreadsheetApp.getUi().alert(
    '✅ Export Complete!',
    'Exported ' + invoices.length + ' invoice(s) with ' + linesToExport.length + ' line items.\n\n' +
    '📄 QuickBooks Online CSV: ' + qboFile.getName() + '\n' +
    '📄 QuickBooks Desktop IIF: ' + qbdFile.getName() + '\n\n' +
    'Files saved to: ' + folder.getName() + '\n\n' +
    'ℹ️ Items referenced in invoices will be auto-created as SERV (service) items by QuickBooks.\n' +
    'No need to pre-import items!\n\n' +
    'Download links added to order sheets.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/******** EXPORT PRICEBOOK ITEMS TO QUICKBOOKS (OPTIONAL) ********/
// This export is OPTIONAL - QuickBooks will auto-create SERV (service) items when importing invoices
// Use this only if you want to pre-define item descriptions or default prices
function exportPriceBookToQuickBooks() {
  requireLicense();
  const items = DataLayer.getPriceBookItems();
  
  if (items.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No items in PriceBook', 'Add items to your PriceBook first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Build IIF for non-inventory items
  const lines = [];
  lines.push('!INVITEM\tNAME\tINVITEMTYPE\tDESC\tACCNT\tPRICE');
  
  for (var i = 0; i < items.length; i++) {
    const item = items[i];
    const sellingPrice = item.basePrice + (item.basePrice * item.defaultMarkup / 100);
    
    // Ensure all fields have values (IIF format requires no empty fields)
    const itemName = sanitizeTab(item.item) || 'ITEM';
    const itemDesc = sanitizeTab(item.notes || item.item) || itemName;
    const itemPrice = numStr(sellingPrice) || '0.00';
    
    lines.push([
      'INVITEM',
      itemName,                                  // NAME - item code (required)
      'SERV',                                    // INVITEMTYPE - service item
      itemDesc,                                  // DESC - description (required)
      QB_INCOME_ACCOUNT || 'Sales Income',       // ACCNT - income account name (required)
      itemPrice                                  // PRICE - selling price (required)
    ].join('\t'));
  }
  
  const iifContent = lines.join('\n');
  
  // Save to Drive
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  const folder = getDriveFolder();
  const file = folder.createFile('Dupuys_Items_' + timestamp + '.iif', iifContent, MimeType.PLAIN_TEXT);
  
  logAction('ExportItems', 'Exported ' + items.length + ' items to IIF', 'Success');
  
  SpreadsheetApp.getUi().alert(
    '✅ PriceBook Items Exported!',
    'Exported ' + items.length + ' non-inventory item(s) to:\n\n' +
    '📄 ' + file.getName() + '\n\n' +
    '📁 Location: ' + folder.getName() + '\n\n' +
    'ℹ️ NOTE: This export is OPTIONAL!\n\n' +
    'QuickBooks will automatically create items as SERV (service type) ' +
    'when you import invoices that reference them.\n\n' +
    'Use this export only if you want to:\n' +
    '• Pre-define item descriptions\n' +
    '• Set specific default prices\n' +
    '• Control item setup before invoicing\n\n' +
    'Otherwise, just import your invoices directly!\n\n' +
    'Items will post to: ' + QB_INCOME_ACCOUNT,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/******** BUILD QUICKBOOKS ONLINE CSV ********/
function buildQboCsv(invoices) {
  // Force positive Qty & Rate; compute Amount = Qty * Rate
  // Columns are simple so you can map them easily in QBO.
  const headers = ['*Qty', '*ItemCode', '*PriceEach', '*AmountTotal'];
  const rows = [headers.join(',')];

  function numStr(n) {
    return (Number(n) || 0).toFixed(2);
  }
  function esc(v) {
    var s = String(v == null ? '' : v);
    if (s.indexOf('"') !== -1 || s.indexOf(',') !== -1 || s.indexOf('\n') !== -1) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  for (var i = 0; i < invoices.length; i++) {
    var inv = invoices[i];

    for (var j = 0; j < inv.lines.length; j++) {
      var line = inv.lines[j];

      var qty = Math.abs(Number(line.qty) || 0);
      var rate = Math.abs(Number(line.rate) || 0);
      var amount = round2(qty * rate);

      var record = [
        qty,                  // leave quantity unformatted for mapper
        esc(line.item),
        numStr(rate),
        numStr(amount)
      ];

      rows.push(record.join(','));
    }
  }

  return rows.join('\n');
}

/******** BUILD QUICKBOOKS DESKTOP IIF ********/
// QuickBooks will auto-create items as SERV (service type) when referenced in invoices
// No need to pre-import items - they're created on-the-fly during invoice import
function buildQbdIif(invoices) {
  const lines = [];
  lines.push('!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tDOCNUM\tTERMS\tAMOUNT\tMEMO');
  lines.push('!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tINVITEM\tQNTY\tPRICE\tAMOUNT\tMEMO');
  lines.push('!ENDTRNS');

  for (var i = 0; i < invoices.length; i++) {
    const inv = invoices[i];

    // --- Date -> MM/DD/YYYY ---
    let date = inv.txnDate || todayYMD();
    if (date instanceof Date) {
      date = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    } else {
      const d = new Date(date);
      if (!isNaN(d.getTime())) date = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    }

    // --- Compute positive total from positive qty*rate ---
    var total = 0;
    for (var j = 0; j < inv.lines.length; j++) {
      const q = Math.abs(Number(inv.lines[j].qty) || 0);
      const r = Math.abs(Number(inv.lines[j].rate) || 0);
      total += round2(q * r);
    }
    const invoiceTotal = round2(total);

    // Optional: keep DOCNUM short for Desktop
    let doc = String(inv.docNumber || '').trim();
    if (doc.length > 11) doc = doc.substring(0, 11);

    // --- Header (TRNS) - POSITIVE amount ---
    lines.push([
      'TRNS',
      'INVOICE',
      date,
      QB_AR_ACCOUNT,
      sanitizeTab(inv.customer),
      sanitizeTab(inv.class),
      sanitizeTab(doc),
      sanitizeTab(inv.terms || 'Net 7'),
      numStr(invoiceTotal),        // POSITIVE
      sanitizeTab(inv.memo || '')
    ].join('\t'));

    // --- Lines (SPL) - NEGATIVE qty (QB Desktop convention), POSITIVE rate, AMOUNT blank ---
    for (var k = 0; k < inv.lines.length; k++) {
      const L = inv.lines[k];
      const qty  = Math.abs(Number(L.qty)  || 0);
      const rate = Math.abs(Number(L.rate) || 0);
      const memo = sanitizeTab(L.description || L.item);

      lines.push([
        'SPL',
        'INVOICE',
        date,
        QB_INCOME_ACCOUNT,                            // Income account (e.g., "Sales Income")
        sanitizeTab(inv.customer),
        sanitizeTab(inv.class),
        sanitizeTab(L.item),
        numStr(-qty),                                 // NEGATIVE qty (lets QB compute a negative line amount)
        numStr(rate),                                 // POSITIVE rate
        '',                                           // AMOUNT left blank so QB computes from qty*rate
        memo
      ].join('\t'));
    }

    lines.push('ENDTRNS');
  }

  return lines.join('\n');
}

/******** ARCHIVE HELPER - GET OR CREATE ARCHIVE FOLDER ********/
function getArchiveFolder() {
  const ss = SpreadsheetApp.getActive();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolders = ssFile.getParents();
  
  // Get the parent folder of the spreadsheet (or root if none)
  const parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  
  // Look for existing Archive folder
  const existingFolders = parentFolder.getFoldersByName('Archived Orders');
  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }
  
  // Create new Archive folder
  return parentFolder.createFolder('Archived Orders');
}

/******** ARCHIVE SINGLE ORDER TO DRIVE ********/
function archiveOrderToDrive(docNumber) {
  const ss = SpreadsheetApp.getActive();
  const orderSheet = ss.getSheetByName(ORDER_SHEET_PREFIX + docNumber);
  
  if (!orderSheet) {
    logAction('Archive', 'Order sheet not found: ' + docNumber, 'Failed');
    return false;
  }
  
  try {
    // Get or create archive folder
    const archiveFolder = getArchiveFolder();
    
    // Create a new spreadsheet for this order
    const newSS = SpreadsheetApp.create(docNumber);
    const newFile = DriveApp.getFileById(newSS.getId());
    
    // Move to archive folder
    archiveFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);
    
    // Copy the order sheet to the new spreadsheet using copyTo method
    const copiedSheet = orderSheet.copyTo(newSS);
    copiedSheet.setName(docNumber);
    
    // Delete the default sheet
    const defaultSheet = newSS.getSheets()[0];
    if (defaultSheet.getName() !== docNumber) {
      newSS.deleteSheet(defaultSheet);
    }
    
    // Flush changes to ensure everything is saved
    SpreadsheetApp.flush();
    
    logAction('Archive', 'Archived ' + docNumber + ' to Drive: ' + newFile.getUrl(), 'Success');
    return newFile.getUrl();
    
  } catch (err) {
    logAction('Archive', 'Failed to archive ' + docNumber + ': ' + String(err), 'Failed');
    return false;
  }
}

/******** ARCHIVE CURRENT ORDER ********/
function archiveCurrentOrder() {
  requireLicense();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (sheetName.indexOf(ORDER_SHEET_PREFIX) !== 0) {
    SpreadsheetApp.getUi().alert('⚠️ Please open an order sheet first (ORDER_TB-...)');
    return;
  }
  
  const docNumber = sheetName.replace(ORDER_SHEET_PREFIX, '');
  
  const response = SpreadsheetApp.getUi().alert(
    'Archive Current Order',
    'Archive order: ' + docNumber + '?\n\nThis will:\n' +
    '1. Create a new Google Sheet in "Archived Orders" folder\n' +
    '2. Copy all order data to the new sheet\n' +
    '3. Delete from this workbook\n' +
    '4. Remove from ORDER_MASTER\n\n' +
    'Continue?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response !== SpreadsheetApp.getUi().Button.YES) return;
  
  // Archive to Drive
  const archiveUrl = archiveOrderToDrive(docNumber);
  
  if (!archiveUrl) {
    SpreadsheetApp.getUi().alert('❌ Failed to archive order. Check Logs sheet for details.');
    return;
  }
  
  // Delete from current workbook
  ss.deleteSheet(sheet);
  
  // Remove from ORDER_MASTER
  removeFromOrderMaster([docNumber]);
  
  // Remove from _OrderData
  removeFromOrderData([docNumber]);
  
  SpreadsheetApp.getUi().alert(
    '✅ Order Archived!',
    'Order ' + docNumber + ' has been archived.\n\n' +
    'Archived file location:\n' +
    archiveUrl + '\n\n' +
    'The order has been removed from this workbook.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/******** ARCHIVE ALL EXPORTED ORDERS ********/
function archiveExported() {
  requireLicense();
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ss.getSheetByName(SHEET.ORDER_DATA);
  
  if (!dataSheet || dataSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('⚠️ No orders to archive');
    return;
  }
  
  const allData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const idx = makeHeaderIndex(headers);
  
  // Find exported orders
  const exportedDocs = [];
  for (var i = 0; i < allData.length; i++) {
    const row = allData[i];
    const exportStatus = String(row[idx['ExportStatus']]).trim();
    if (exportStatus === 'Exported') {
      const docNum = String(row[idx['DocNumber']]).trim();
      if (docNum && exportedDocs.indexOf(docNum) === -1) {
        exportedDocs.push(docNum);
      }
    }
  }
  
  if (exportedDocs.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No exported orders to archive');
    return;
  }
  
  const response = SpreadsheetApp.getUi().alert(
    'Archive All Exported Orders',
    'Found ' + exportedDocs.length + ' exported order(s).\n\nThis will:\n' +
    '1. Create new Google Sheets in "Archived Orders" folder\n' +
    '2. Copy each order to its own new sheet\n' +
    '3. Delete from this workbook\n' +
    '4. Remove from ORDER_MASTER\n\n' +
    'Continue?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response !== SpreadsheetApp.getUi().Button.YES) return;
  
  // Archive each order to Drive
  var archivedCount = 0;
  var failedCount = 0;
  var archiveUrls = [];
  
  for (var i = 0; i < exportedDocs.length; i++) {
    const docNum = exportedDocs[i];
    const archiveUrl = archiveOrderToDrive(docNum);
    
    if (archiveUrl) {
      archiveUrls.push(docNum + ': ' + archiveUrl);
      archivedCount++;
      
      // Delete order sheet from workbook
      const orderSheet = ss.getSheetByName(ORDER_SHEET_PREFIX + docNum);
      if (orderSheet) {
        ss.deleteSheet(orderSheet);
      }
    } else {
      failedCount++;
    }
  }
  
  // Remove archived orders from ORDER_MASTER and _OrderData
  if (archivedCount > 0) {
    removeFromOrderMaster(exportedDocs);
    removeFromOrderData(exportedDocs);
  }
  
  logAction('Archive', 'Archived ' + archivedCount + ' order(s), ' + failedCount + ' failed', 'Success');
  
  var message = '✅ Archive Complete!\n\n' +
    'Successfully archived: ' + archivedCount + ' order(s)\n';
  
  if (failedCount > 0) {
    message += 'Failed: ' + failedCount + ' order(s)\n';
  }
  
  message += '\n📁 Orders saved to "Archived Orders" folder\n' +
    '🗑️ Orders removed from this workbook\n\n';
  
  if (archiveUrls.length > 0 && archiveUrls.length <= 5) {
    message += 'Archived files:\n' + archiveUrls.join('\n');
  } else if (archiveUrls.length > 5) {
    message += 'View archived files in the "Archived Orders" folder.';
  }
  
  SpreadsheetApp.getUi().alert('Archive Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/******** ORDER_MASTER ROW DELETE (PRESERVE FORMATTING/LINKS) ********/
function removeFromOrderMaster(docNumbers) {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(SHEET.ORDER_MASTER);
  if (!master) return;

  // Header row is 5 in buildOrderMasterSheet()
  const HEADER_ROW = 5;
  const headers = master.getRange(HEADER_ROW, 1, 1, master.getLastColumn()).getValues()[0];
  const idx = makeHeaderIndex(headers);

  if (master.getLastRow() <= HEADER_ROW) return;

  const firstDataRow = HEADER_ROW + 1;
  const data = master.getRange(firstDataRow, 1, master.getLastRow() - HEADER_ROW, master.getLastColumn()).getValues();

  // Map docNumbers for fast lookup
  const toDelete = {};
  (docNumbers || []).forEach(d => { if (d) toDelete[String(d).trim()] = true; });

  // Collect absolute sheet row numbers to delete
  const rowsToDelete = [];
  for (var r = 0; r < data.length; r++) {
    const doc = String(data[r][idx['DocNumber']]).trim();
    if (toDelete[doc]) rowsToDelete.push(firstDataRow + r);
  }
  if (rowsToDelete.length === 0) return;

  // Delete from bottom to top (preserves formatting of remaining rows)
  rowsToDelete.sort((a,b) => b - a).forEach(rowNum => master.deleteRow(rowNum));

  // Renumber the "#" column only
  const lastRow = master.getLastRow();
  if (lastRow >= firstDataRow) {
    const n = lastRow - HEADER_ROW;
    const numbers = Array.from({length: n}, (_, i) => [i + 1]);
    master.getRange(firstDataRow, 1, n, 1).setValues(numbers);
  }

  // Re-apply banding to data area only (don’t touch header/title)
  try {
    const dataArea = master.getRange(firstDataRow, 1, Math.max(0, master.getLastRow() - HEADER_ROW), master.getLastColumn());
    const bandings = dataArea.getBandings();
    bandings.forEach(b => b.remove());
    if (dataArea.getNumRows() > 0) {
      const b = dataArea.applyRowBanding(SpreadsheetApp.BandingTheme.CYAN, false, false);
      b.setFirstRowColor('#ffffff').setSecondRowColor('#f8f9fa');
    }
  } catch (_e) { /* optional */ }

  // Keep the stat formula intact
  try { master.getRange('B4').setFormula('=COUNTA(B6:B)'); } catch (_e) {}
}

/******** REMOVE FROM ORDER DATA ********/
function removeFromOrderData(docNumbers) {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ss.getSheetByName(SHEET.ORDER_DATA);
  
  if (!dataSheet || dataSheet.getLastRow() < 2) return;
  
  const allData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const idx = makeHeaderIndex(headers);
  
  const keepData = allData.filter(function(row) {
    const docNum = String(row[idx['DocNumber']]).trim();
    return docNumbers.indexOf(docNum) === -1;
  });
  
  dataSheet.clearContents();
  dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (keepData.length > 0) {
    dataSheet.getRange(2, 1, keepData.length, headers.length).setValues(keepData);
  }
}

/******** EXPORT HELPER FUNCTIONS ********/
function csvEscape(value) {
  const str = String(value == null ? '' : value);
  return '"' + str.replace(/"/g, '""') + '"';
}

function sanitizeTab(value) {
  return String(value == null ? '' : value)
    .replace(/\t/g, ' ')
    .replace(/\r?\n/g, ' ')
    .trim();
}

function numStr(num) {
  if (num == null || isNaN(num)) return '';
  return Number(num).toFixed(2);
}

function round2(num) {
  return Math.round((Number(num) + Number.EPSILON) * 100) / 100;
}

function getDriveFolder() {
  // Get the spreadsheet file's location in Google Drive
  const ss = SpreadsheetApp.getActive();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolders = ssFile.getParents();
  
  // Get the parent folder of the spreadsheet (or root if none)
  const parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  
  // Look for existing QuickBooks_Exports folder in the same location
  const existingFolders = parentFolder.getFoldersByName('QuickBooks_Exports');
  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }
  
  // Create new QuickBooks_Exports folder in the same location as the spreadsheet
  return parentFolder.createFolder('QuickBooks_Exports');
}

/******** WEB APP ********/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('order_form_webapp').setTitle('Dupuys Dockside Order Form');
}

function getItemsForWebApp() {
  return DataLayer.getPriceBookItems().map(function(item) {
    return { code: item.item, name: item.notes || item.item, category: item.category, unit: item.unit, price: item.basePrice };
  });
}

function getBoatsForWebApp() {
  return DataLayer.getCustomers().map(function(c) {
    return { id: c.boatId, name: c.boatName };
  });
}

function verifyPinForWebApp(boatId, pin) {
  try {
    if (!boatId || !pin) return { success: false, error: 'Missing boat ID or PIN' };
    const isValid = DataLayer.verifyPin(boatId, pin);
    if (!isValid) return { success: false, error: 'Invalid PIN for this boat' };
    return { success: true };
  } catch (e) {
    return { success: false, error: 'Verification error: ' + e.message };
  }
}

function submitWebAppOrder(orderData) {
  try {
    if (!orderData || !orderData.boatId || !orderData.pin) return { success: false, error: 'Missing data' };
    if (!DataLayer.verifyPin(orderData.boatId, orderData.pin)) return { success: false, error: 'Invalid PIN' };
    
    const customer = DataLayer.getCustomerByBoatId(orderData.boatId);
    if (!customer) return { success: false, error: 'Customer not found' };
    
    const docNumber = DataLayer.getNextDocNumber(orderData.boatId);
    const items = (orderData.items || []).filter(it => it && it.code && Number(it.qty) > 0).map(function(item) {
      const priceItem = DataLayer.getPriceBookItem(item.code);
      return { itemCode: item.code, category: priceItem ? priceItem.category : '', unit: priceItem ? priceItem.unit : 'ea', qty: Number(item.qty) };
    });
    
    if (items.length === 0) return { success: false, error: 'No valid items' };
    
    const orderInfo = {
      docNumber: docNumber,
      boatId: orderData.boatId,
      boatName: customer.boatName,
      qbCustomerName: customer.qbCustomerName || customer.boatName,
      txnDate: orderData.orderDate || todayYMD(),
      deliveryLocation: orderData.deliveryLocation || '',
      phone: orderData.phone || '',
      notes: orderData.notes || '',
      items: items
    };
    
    createOrderSheet(orderInfo);
    syncOrderToDataSheet(docNumber);
    
    logAction('WebAppOrder', 'Created order ' + docNumber, 'Success');
    return { success: true, docNumber: docNumber };
    
  } catch (err) {
    logAction('WebAppError', String(err), 'Failed');
    return { success: false, error: String(err) };
  }
}

/******** CONVERT SHEET TO TABLE ********/
function convertSheetToTable(sheet, tableName) {
  if (!sheet) return;
  
  try {
    // Check if sheet has headers
    if (sheet.getLastRow() < 1) return;
    
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    
    const lastRow = Math.max(2, sheet.getLastRow()); // Need at least 2 rows for table
    
    // Create the table range (including headers)
    const tableRange = sheet.getRange(1, 1, lastRow, lastCol);
    
    // Try to use native Google Sheets Table API
    try {
      // First, remove any existing tables in this range
      const existingTables = sheet.getTables();
      for (var i = 0; i < existingTables.length; i++) {
        existingTables[i].remove();
      }
      
      // Create a new native Google Sheets table
      const table = sheet.addTable(tableRange);
      
      // Configure table settings
      table.setHeaderRowIndex(0); // First row is header
      
      logAction('TableConvert', 'Created native table for ' + sheet.getName() + ': ' + tableName, 'Success');
      
    } catch (tableErr) {
      // If native tables not available, fall back to styled range
      logAction('TableConvert', 'Native tables not available, using styled format: ' + String(tableErr), 'Info');
      
      // STICKY HEADER - Freeze row 1
      sheet.setFrozenRows(1);
      
      // Get the header row
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      
      // BEAUTIFUL HEADER FORMATTING (matches native table style)
      headerRange
        .setFontWeight('bold')
        .setBackground('#1a73e8')
        .setFontColor('#ffffff')
        .setFontSize(11)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('left');
      
      // Add borders around header
      headerRange.setBorder(
        true, true, true, true, true, true,
        '#ffffff', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
      
      // Add filter to header row
      if (sheet.getFilter()) {
        sheet.getFilter().remove();
      }
      const dataRange = sheet.getRange(1, 1, sheet.getMaxRows(), lastCol);
      dataRange.createFilter();
      
      // MODERN BANDED ROWS
      const bandedRange = sheet.getRange(1, 1, sheet.getMaxRows(), lastCol);
      const existingBandings = bandedRange.getBandings();
      for (var j = 0; j < existingBandings.length; j++) {
        existingBandings[j].remove();
      }
      
      // Apply modern color scheme
      const banding = bandedRange.applyRowBanding(SpreadsheetApp.BandingTheme.CYAN, false, false);
      banding
        .setHeaderRowColor('#1a73e8')
        .setFirstRowColor('#ffffff')
        .setSecondRowColor('#e8f0fe')
        .setFooterRowColor(null);
      
      // Set column widths for readability
      for (var col = 1; col <= lastCol; col++) {
        sheet.setColumnWidth(col, 130);
      }
      
      // Add subtle grid lines
      dataRange.setBorder(
        false, false, false, false, true, true,
        '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID
      );
      
      logAction('TableConvert', 'Applied styled table format to ' + sheet.getName(), 'Success');
    }
    
  } catch (err) {
    logAction('TableConvert', 'Could not convert ' + sheet.getName() + ': ' + String(err), 'Warning');
  }
}

/******** CALCULATE ORDER SHEET POSITIONS ********/
function calculateOrderSheetPositions(orderSheet) {
  // Dynamically find row positions in order sheet based on current content
  // This handles orders with any number of items and manually added items
  
  try {
    const maxRows = Math.min(orderSheet.getLastRow(), 200); // Limit search range
    
    // Find the last row with item data (look for items starting at row 9)
    let lastItemRow = 9;
    for (var row = 9; row <= maxRows; row++) {
      const itemCode = orderSheet.getRange(row, 1).getValue();
      const cellValue = String(itemCode).trim();
      
      // Stop if we hit a section marker or empty after items
      if (cellValue.indexOf('💰 TOTAL') === 0 || cellValue.indexOf('TOTAL') === 0) {
        lastItemRow = row - 1;
        break;
      }
      if (cellValue && row > lastItemRow) {
        lastItemRow = row;
      }
    }
    
    // Search for actual section markers instead of calculating positions
    let totalsRow = lastItemRow + 2;
    let actionsRow = 0;
    let exportStatusRow = 0;
    let qbExportLinkRow = 0;
    
    // Find Actions section by searching for marker text
    for (var searchRow = lastItemRow; searchRow <= maxRows; searchRow++) {
      const cellA = String(orderSheet.getRange(searchRow, 1).getValue()).trim();
      
      // Look for "⚙️ Actions" or totals marker
      if (cellA.indexOf('💰 TOTAL') === 0) {
        totalsRow = searchRow;
      }
      if (cellA.indexOf('⚙️ Actions') === 0 || cellA.indexOf('Actions & Export') === 0) {
        actionsRow = searchRow;
      }
      if (cellA.indexOf('Export Status') === 0) {
        exportStatusRow = searchRow;
      }
      if (cellA.indexOf('QB Export Link') === 0) {
        qbExportLinkRow = searchRow;
      }
    }
    
    // If markers not found, calculate based on last item
    if (actionsRow === 0) actionsRow = lastItemRow + 15;
    if (exportStatusRow === 0) exportStatusRow = actionsRow + 1;
    if (qbExportLinkRow === 0) qbExportLinkRow = actionsRow + 3;
    
    const notesRow = totalsRow + 2;
    const receiptSectionRow = notesRow + 5;
    const receiptLinkRow = actionsRow + 2;
    
    return {
      lastItemRow: lastItemRow,
      totalsRow: totalsRow,
      notesRow: notesRow,
      receiptSectionRow: receiptSectionRow,
      actionsRow: actionsRow,
      exportStatusRow: exportStatusRow,
      receiptLinkRow: receiptLinkRow,
      qbExportLinkRow: qbExportLinkRow
    };
    
  } catch (err) {
    // Fallback to default positions if calculation fails
    logAction('PositionCalc', 'Failed to calculate positions, using fallback: ' + String(err), 'Warning');
    return {
      lastItemRow: 20,
      totalsRow: 22,
      notesRow: 24,
      receiptSectionRow: 29,
      actionsRow: 37,
      exportStatusRow: 38,
      receiptLinkRow: 39,
      qbExportLinkRow: 40
    };
  }
}

/******** HELPERS ********/
function ensureSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function initializeSheetHeaders(sheet, headers) {
  if (!sheet || sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#eef2f7');
}

function setColumnWidths(sheet, width) {
  if (!sheet) return;
  sheet.setColumnWidths(1, sheet.getLastColumn() || 1, width);
}

function applyListValidation(sheet, startRow, col, listValues) {
  if (!sheet) return;
  sheet.getRange(startRow, col, sheet.getMaxRows() - startRow + 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(listValues, true).build());
}

function protectHeaders(sheet) {
  if (!sheet) return;
  const protection = sheet.getRange(1, 1, 1, sheet.getLastColumn()).protect();
  protection.setDescription(sheet.getName() + ' header');
  protection.setWarningOnly(true);
}

function makeHeaderIndex(headers) {
  const idx = {};
  for (var i = 0; i < headers.length; i++) idx[String(headers[i]).trim()] = i;
  return idx;
}

function first(arrOrVal) {
  return Array.isArray(arrOrVal) ? arrOrVal[0] : arrOrVal;
}

function normalizeDateYMD(dateInput) {
  if (!dateInput) return '';
  const d = new Date(dateInput);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function todayYMD() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function padLeft(num, length) {
  var str = String(num);
  while (str.length < length) str = '0' + str;
  return str;
}

function uiToast(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, '⚓ Dupuys Dockside V4', 4);
}

function logAction(action, details, status) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.LOGS);
    if (!sh) return;
    sh.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details, status || 'Success']);
  } catch (e) {}
}

function testWebAppConnection() {
  const boats = DataLayer.getCustomers();
  const items = DataLayer.getPriceBookItems();
  
  if (boats.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No customers found. Add customers first.');
    return;
  }
  
  if (items.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No items found. Add items first.');
    return;
  }
  
  SpreadsheetApp.getUi().alert('✅ Web App Ready!\n\n' +
    'Customers: ' + boats.length + '\n' +
    'Items: ' + items.length + '\n\n' +
    'Your web app can now accept orders.');
}

/******** WEB APP DEPLOYMENT INFO ********/
function showWebAppDeploymentInstructions() {
  requireLicense();
  const ui = SpreadsheetApp.getUi();
  
  const instructions = 'WEB APP DEPLOYMENT INSTRUCTIONS:\n\n' +
    '1. In Apps Script editor, click "Deploy" > "New deployment"\n' +
    '2. Click gear icon ⚙️ next to "Select type"\n' +
    '3. Choose "Web app"\n' +
    '4. Settings:\n' +
    '   - Description: Dupuys Order Form\n' +
    '   - Execute as: Me\n' +
    '   - Who has access: Anyone\n' +
    '5. Click "Deploy"\n' +
    '6. Copy the Web app URL\n' +
    '7. Share it with boat captains\n\n' +
    'Note: Store the URL somewhere safe for reference.';
  
  ui.alert('Web App Deployment', instructions, ui.ButtonSet.OK);
}

function getWebAppUrl() {
  requireLicense();
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Web App URL',
    'To get your Web App URL:\n\n' +
    '1. Extensions > Apps Script\n' +
    '2. Click "Deploy" > "Manage deployments"\n' +
    '3. Click on the active web app deployment\n' +
    '4. Copy the "Web app URL"\n\n' +
    'Share that URL with boat captains to place orders.',
    ui.ButtonSet.OK
  );
}

/*************************************************
 * DUPUYS DOCKSIDE V4 - COMPLETE DEPLOYMENT GUIDE
 *************************************************/

/*************************************************
 * PART 1: INITIAL SETUP (Admin/Developer)
 *************************************************/
// 1. Create new Google Sheet
// 2. Extensions > Apps Script
// 3. Delete any existing code
// 4. Paste this ENTIRE script
// 5. Save (Ctrl+S or Cmd+S)
// 6. Close Apps Script tab
// 7. Refresh your Google Sheet
// 8. You'll see "⚓ Dupuys Dockside V4 [ADMIN]" menu appear
// 9. Run: Menu > 🔧 Initialize System
// 10. Wait 10-15 seconds (creates all sheets with modern tables)

/*************************************************
 * PART 2: CONFIGURATION (Before Client Deployment)
 *************************************************/
// 11. Add Customers: Menu > 👥 Customers > Add Customer Manually
//     - Or import from QuickBooks CSV
//     - Each customer gets a BoatID and PIN
//
// 12. Add Items: Menu > 🛒 Grocery Items > Add Item Manually
//     - Or import grocery list
//     - Set base prices and default markup %
//
// 13. Deploy Web App (Optional):
//     - Menu > 🌐 Web App > Deploy Web App Instructions
//     - Follow the deployment steps
//     - Save the Web App URL in Config
//
// 14. Install Triggers: Menu > 📊 Views > Reinstall Edit Sync
//     - Installs automatic sync between sheets
//
// 15. Run Checklist: Menu > ✅ Deployment Checklist
//     - Verifies all setup is complete
//     - Shows any remaining issues

/*************************************************
 * PART 3: CLIENT DEPLOYMENT
 *************************************************/
// 16. Switch to Client Mode:
//     - In Apps Script editor (line 24)
//     - Change: const CLIENT_MODE = false;
//     - To: const CLIENT_MODE = true;
//     - Save and refresh sheet
//
// 17. Menu will change to "⚓ Dupuys" (simplified)
//     - Clients see only necessary operations
//     - Technical functions hidden
//
// 18. Share with client:
//     - Give Editor access to spreadsheet
//     - Share Web App URL with boat captains
//     - Provide quick start guide (Menu > ℹ️ Help)

/*************************************************
 * CLIENT MODE FEATURES (Simplified Menu):
 *************************************************/
// ⚓ Dupuys Menu:
// - 📋 Order Master (view all orders)
// - 🎯 CEO Dashboard (real-time metrics)
// - 🛒 Shopping List (grouped by category)
// - 💰 QuickBooks Export (export & archive)
// - 👥 Manage Customers (add customers, PINs)
// - 🛒 Manage Items (add grocery items)
// - 🔄 Refresh Data (manual sync if needed)
// - 🔗 Get Web App URL (share link)
// - ℹ️ Help & Instructions (full guide)

/*************************************************
 * ADMIN MODE FEATURES (Full Access):
 *************************************************/
// ⚓ Dupuys Dockside V4 [ADMIN] Menu:
// - All client features PLUS:
// - 🔧 Initialize System
// - 🌱 Seed Sample Data
// - ✅ Deployment Checklist
// - Import/Export tools
// - Web App deployment
// - Show/Hide order sheets
// - Convert to Tables
// - Clear Cache
// - Trigger management

/*************************************************
 * KEY FEATURES:
 *************************************************/
// ✅ Auto-hidden order sheets (cleaner workbook)
// ✅ Native Google Sheets Tables with filters
// ✅ Sticky headers on all tables
// ✅ Full bidirectional sync (Master ↔ Orders)
// ✅ Dynamic row positioning (any # of items)
// ✅ Modern color-coded design
// ✅ Yellow-highlighted entry fields
// ✅ Automatic calculations
// ✅ Receipt image upload section
// ✅ QB export (Online & Desktop)
// ✅ Web App for boat captains
// ✅ CEO Dashboard with metrics
// ✅ Shopping List by category
// ✅ PIN-based security
// ✅ Comprehensive logging

/*************************************************
 * DAILY CLIENT WORKFLOW:
 *************************************************/
// 1. Orders arrive → Hidden sheets created automatically
// 2. View in Order Master → Click "📄 Open Order" links
// 3. Fill in Base Cost (yellow column) as you shop
// 4. Upload receipt images in Receipt Images section
// 5. Update Status dropdown (Pending → Shopping → Delivered)
// 6. Changes sync instantly to Master & Dashboard
// 7. Set Export Status = "Ready" when complete
// 8. Export to QuickBooks (batch or individual)
// 9. Archive orders to Drive folder (current or all exported)

/*************************************************
 * ADMIN TASKS (Via Apps Script Editor):
 *************************************************/
// - Initialize System: Run initializeWorkbook()
// - Seed Test Data: Run seedSampleData()
// - Clear Cache: Run clearCache()
// - Reinstall Triggers: Run installOnEditTrigger()
// - Deployment Check: Run runDeploymentChecklist()
// - Convert Tables: Run convertAllToTables()
// - Switch Modes: Change CLIENT_MODE constant

/*************************************************
 * TROUBLESHOOTING:
 *************************************************/
// - If sync not working: Menu > 🔗 Reinstall Edit Sync
// - If data stale: Menu > 🔄 Refresh Data
// - If QB link wrong spot: System recalculates dynamically
// - If sheets messy: Menu > 🙈 Hide All Order Sheets
// - If tables look plain: Menu > 📊 Convert to Tables
// - Check logs: View the "Logs" sheet

/*************************************************
 * SUPPORT:
 *************************************************/
// - Clients: Use Menu > ℹ️ Help & Instructions
// - Admins: Review this documentation
// - Issues: Check Logs sheet for errors
// - Updates: Re-paste updated code, refresh sheet
 /*************************************************/