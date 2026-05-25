/* Rahaza Digital Premium Webapps Builder by ddsuparman */

// ID Spreadsheet yang ditargetkan
const SPREADSHEET_ID = "1aRClSDk_R5_JJiJgS3FAISljUOI3ZACKgJr5R19ckQM";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('KeuanganKu - Native App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Handle External POST Requests (API/Webhook) - Mendukung PWA Standalone Mode
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    
    const action = payload.action;
    const args = payload.arguments || [];
    
    let result;
    if (action === "loginUser") {
      result = loginUser(args[0], args[1]);
    } else if (action === "getDashboardData") {
      result = getDashboardData();
    } else if (action === "getTransactions") {
      result = getTransactions();
    } else if (action === "getInstallments") {
      result = getInstallments();
    } else if (action === "addTransaction") {
      result = addTransaction(args[0]);
    } else if (action === "deleteTransaction") {
      result = deleteTransaction(args[0]);
    } else if (action === "addInstallment") {
      result = addInstallment(args[0]);
    } else if (action === "deleteInstallment") {
      result = deleteInstallment(args[0]);
    } else if (action === "resetDatabase") {
      result = resetDatabase();
    } else if (action === "setupDatabase") {
      result = setupDatabase();
    } else {
      // Fallback jika memanggil webhook luar dengan payload format addTransaction
      if (action === "addTransaction" && payload.data) {
        result = addTransaction(payload.data);
      } else {
        result = createResponse(true, null, "POST request diterima dengan baik.");
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Menangani permintaan OPTIONS preflight (CORS) jika diperlukan browser
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Standardized API Response Contract
 */
function createResponse(success, data, message) {
  return { "success": success, "data": data, "message": message };
}

/**
 * Helper: Safely serialize a spreadsheet row value.
 */
function serializeValue(val) {
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'number') return val;
  return String(val).trim();
}

/**
 * Setup Database: Inisialisasi sheet USERS, TRANSACTIONS, dan INSTALLMENTS jika belum ada.
 */
function setupDatabase() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Tab USERS
    let sheetUsers = ss.getSheetByName("USERS");
    if (!sheetUsers) {
      sheetUsers = ss.insertSheet("USERS");
      sheetUsers.appendRow(["id", "username", "password", "role", "status", "created_at"]);
      // Akun default Active admin / admin123
      sheetUsers.appendRow(["usr-1", "admin", "admin123", "Administrator", "Active", new Date()]);
    }
    
    // Tab TRANSACTIONS
    let sheetTrx = ss.getSheetByName("TRANSACTIONS");
    if (!sheetTrx) {
      sheetTrx = ss.insertSheet("TRANSACTIONS");
      sheetTrx.appendRow(["id", "date", "type", "category", "amount", "description", "status", "timestamp"]);
    }
    
    // Tab INSTALLMENTS
    let sheetInst = ss.getSheetByName("INSTALLMENTS");
    if (!sheetInst) {
      sheetInst = ss.insertSheet("INSTALLMENTS");
      sheetInst.appendRow(["id", "name", "creditor", "total_amount", "paid_amount", "remaining", "start_date", "status", "description", "timestamp"]);
    }
    
    return createResponse(true, null, "Database berhasil disiapkan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/**
 * Reset Database: Mengosongkan data transaksi dan cicilan (meninggalkan header).
 */
function resetDatabase() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    const sheetTrx = ss.getSheetByName("TRANSACTIONS");
    if (sheetTrx) {
      sheetTrx.clear();
      sheetTrx.appendRow(["id", "date", "type", "category", "amount", "description", "status", "timestamp"]);
    }
    
    const sheetInst = ss.getSheetByName("INSTALLMENTS");
    if (sheetInst) {
      sheetInst.clear();
      sheetInst.appendRow(["id", "name", "creditor", "total_amount", "paid_amount", "remaining", "start_date", "status", "description", "timestamp"]);
    }
    
    return createResponse(true, null, "Database transaksi dan cicilan berhasil dikosongkan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   AUTH — LOGIN / LOGOUT
   ================================================================ */

/**
 * Login: cek username & password di sheet USERS.
 * Mengembalikan data user (tanpa password) jika berhasil.
 */
function loginUser(username, password) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Auto-setup database jika USERS belum terbentuk
    let sheet = ss.getSheetByName("USERS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("USERS");
    }

    const data = sheet.getDataRange().getValues();
    // Header: id | username | password | role | status
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[1]).trim() === String(username).trim() &&
          String(row[2]).trim() === String(password).trim()) {

        if (String(row[4]).trim() !== 'Active') {
          return createResponse(false, null, "Akun Anda tidak aktif. Hubungi administrator.");
        }
        return createResponse(true, {
          id:       String(row[0]),
          username: String(row[1]),
          role:     String(row[3]),
          status:   String(row[4])
        }, "Login berhasil.");
      }
    }
    return createResponse(false, null, "Username atau password salah.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   DASHBOARD
   ================================================================ */

function getDashboardData() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Auto-setup jika sheet TRANSACTIONS tidak ditemukan
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const data = sheet.getDataRange().getValues();
    let totalIncome = 0, totalExpense = 0, totalDebt = 0, totalReceivable = 0, totalInstallmentPaid = 0;

    for (let i = 1; i < data.length; i++) {
      const type   = String(data[i][2]).trim();
      const amount = parseFloat(data[i][4]) || 0;
      if      (type === "Pendapatan")  totalIncome   += amount;
      else if (type === "Pengeluaran") totalExpense  += amount;
      else if (type === "Utang")       totalDebt     += amount;
      else if (type === "Piutang")     totalReceivable += amount;
      else if (type === "Cicilan")     totalInstallmentPaid += amount;
    }

    // Ringkasan cicilan dari sheet INSTALLMENTS
    const installSheet = ss.getSheetByName("INSTALLMENTS");
    let totalInstallmentOutstanding = 0;
    if (installSheet) {
      const iData = installSheet.getDataRange().getValues();
      for (let i = 1; i < iData.length; i++) {
        if (String(iData[i][7]).trim() !== 'Lunas') {
          totalInstallmentOutstanding += parseFloat(iData[i][5]) || 0; // kolom sisa
        }
      }
    }

    const balance = (totalIncome + totalDebt) - (totalExpense + totalReceivable + totalInstallmentPaid);

    return createResponse(true, {
      income:                   totalIncome,
      expense:                  totalExpense + totalInstallmentPaid,
      debt:                     totalDebt,
      receivable:               totalReceivable,
      balance:                  balance,
      installmentPaid:          totalInstallmentPaid,
      installmentOutstanding:   totalInstallmentOutstanding
    }, "Data dashboard berhasil dimuat.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   TRANSACTIONS — CRUD
   ================================================================ */

function getTransactions() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const transactions = [];

    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      if (!row[0] || String(row[0]).trim() === '') continue;

      const trx = {};
      for (let j = 0; j < headers.length; j++) {
        trx[headers[j]] = serializeValue(row[j]);
      }
      trx['amount'] = parseFloat(row[4]) || 0;

      const rawDate = row[1];
      if (rawDate instanceof Date) {
        trx['date'] = rawDate.toISOString();
      } else if (rawDate && String(rawDate).trim() !== '') {
        const parsed = new Date(String(rawDate).trim());
        trx['date'] = isNaN(parsed.getTime()) ? String(rawDate).trim() : parsed.toISOString();
      } else {
        trx['date'] = '';
      }
      transactions.push(trx);
    }

    return createResponse(true, transactions, "Data transaksi berhasil dimuat.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function addTransaction(trxData) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const newId = "TRX-" + new Date().getTime();
    sheet.appendRow([
      newId,
      trxData.date,
      trxData.type,
      trxData.category,
      parseFloat(trxData.amount) || 0,
      trxData.description || '',
      trxData.status || 'Selesai',
      new Date()
    ]);

    // Jika tipe Cicilan dan ada installment_id, catat setoran ke sheet INSTALLMENTS
    if (trxData.type === "Cicilan" && trxData.installment_id) {
      recordInstallmentPayment(trxData.installment_id, parseFloat(trxData.amount) || 0);
    }

    return createResponse(true, { id: newId }, "Transaksi berhasil disimpan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function deleteTransaction(id) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        return createResponse(true, null, "Transaksi berhasil dihapus.");
      }
    }
    return createResponse(false, null, "ID Transaksi tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   INSTALLMENTS — Cicilan (Total / Setor / Sisa)
   Sheet kolom: id | name | creditor | total_amount | paid_amount | remaining | start_date | status | description | timestamp
   ================================================================ */

function getInstallments() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const list    = [];

    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      if (!row[0] || String(row[0]).trim() === '') continue;

      const item = {};
      for (let j = 0; j < headers.length; j++) {
        item[headers[j]] = serializeValue(row[j]);
      }
      item['total_amount'] = parseFloat(row[3]) || 0;
      item['paid_amount']  = parseFloat(row[4]) || 0;
      item['remaining']    = parseFloat(row[5]) || 0;

      const rawDate = row[6];
      if (rawDate instanceof Date) item['start_date'] = rawDate.toISOString();
      else item['start_date'] = rawDate ? String(rawDate).trim() : '';

      list.push(item);
    }

    return createResponse(true, list, "Data cicilan berhasil dimuat.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function addInstallment(trxData) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const newId      = "INST-" + new Date().getTime();
    const total      = parseFloat(trxData.total_amount) || 0;
    const paid       = 0;
    const remaining  = total;

    sheet.appendRow([
      newId,
      trxData.name        || '',
      trxData.creditor    || '',
      total,
      paid,
      remaining,
      trxData.start_date  || new Date(),
      'Berjalan',          // status
      trxData.description || '',
      new Date()           // timestamp
    ]);

    return createResponse(true, { id: newId }, "Cicilan baru berhasil ditambahkan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function recordInstallmentPayment(installmentId, payAmount) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(installmentId).trim()) {
        const rowNum     = i + 1;
        const total      = parseFloat(data[i][3]) || 0;
        const oldPaid    = parseFloat(data[i][4]) || 0;
        const newPaid    = oldPaid + (parseFloat(payAmount) || 0);
        const remaining  = Math.max(0, total - newPaid);
        const status     = remaining <= 0 ? 'Lunas' : 'Berjalan';

        sheet.getRange(rowNum, 5).setValue(newPaid);    // paid_amount
        sheet.getRange(rowNum, 6).setValue(remaining);  // remaining
        sheet.getRange(rowNum, 8).setValue(status);     // status

        return createResponse(true, { paid: newPaid, remaining: remaining, status: status },
          "Pembayaran cicilan berhasil dicatat.");
      }
    }
    return createResponse(false, null, "ID Cicilan tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function deleteInstallment(id) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        return createResponse(true, null, "Cicilan berhasil dihapus.");
      }
    }
    return createResponse(false, null, "ID Cicilan tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}
