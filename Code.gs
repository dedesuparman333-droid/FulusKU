/**
 * FulusKUpro — Google Sheets Backend API Script
 * 
 * SCRIPT INI DIINSTAL DI SPREADSHEET GOOGLE APPS SCRIPT:
 * 1. Buka spreadsheet ID: 1aRClSDk_R5_JJiJgS3FAISljUOI3ZACKgJr5R19ckQM
 * 2. Masuk ke Extensions > Apps Script
 * 3. Hapus kode bawaan, lalu paste seluruh kode ini
 * 4. Deploy sebagai Web App:
 *    - Click "Deploy" > "New deployment"
 *    - Select type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 5. Salin URL Web App hasil deploy ke pengaturan PWA Anda!
 */

// ID Spreadsheet Default
var DEFAULT_SPREADSHEET_ID = "1aRClSDk_R5_JJiJgS3FAISljUOI3ZACKgJr5R19ckQM";

// --- CORS & HTTP ENTRY POINT ---
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    var args = payload.arguments || [];
    var spreadsheetId = payload.spreadsheetId || DEFAULT_SPREADSHEET_ID;
    
    // Buka spreadsheet berdasarkan ID yang dikirim client
    var ss = SpreadsheetApp.openById(spreadsheetId);
    
    var result;
    if (action === "loginUser") {
      result = loginUser(args[0], args[1], ss);
    } else if (action === "getDashboardData") {
      result = getDashboardData(ss);
    } else if (action === "getTransactions") {
      result = getTransactions(ss);
    } else if (action === "getInstallments") {
      result = getInstallments(ss);
    } else if (action === "addTransaction") {
      result = addTransaction(args[0], ss);
    } else if (action === "deleteTransaction") {
      result = deleteTransaction(args[0], ss);
    } else if (action === "addInstallment") {
      result = addInstallment(args[0], ss);
    } else if (action === "deleteInstallment") {
      result = deleteInstallment(args[0], ss);
    } else if (action === "resetDatabase") {
      result = resetDatabase(ss);
    } else {
      throw new Error("Aksi tidak dikenal: " + action);
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Gagal memproses permintaan server: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Menangani permintaan OPTIONS preflight (CORS) jika diperlukan browser
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

// --- HELPER UNTUK KONTEKS SPREADSHEET ---
function getActiveSs(ss) {
  if (ss) return ss;
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    return SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
  }
}

// --- INITIALIZATION SHEET & HEADER ---
function initSheets(ss) {
  var activeSs = getActiveSs(ss);
  
  // Tab Users
  var sheetUsers = activeSs.getSheetByName("Users");
  if (!sheetUsers) {
    sheetUsers = activeSs.insertSheet("Users");
    sheetUsers.appendRow(["id", "username", "password", "created_at"]);
    // Tambahkan default user admin / admin123
    sheetUsers.appendRow(["usr_1", "admin", "admin123", new Date().toISOString()]);
  }
  
  // Tab Transactions
  var sheetTrx = activeSs.getSheetByName("Transactions");
  if (!sheetTrx) {
    sheetTrx = activeSs.insertSheet("Transactions");
    sheetTrx.appendRow(["id", "date", "type", "category", "description", "amount", "installment_id", "created_at"]);
  }
  
  // Tab Installments
  var sheetInst = activeSs.getSheetByName("Installments");
  if (!sheetInst) {
    sheetInst = activeSs.insertSheet("Installments");
    sheetInst.appendRow(["id", "name", "creditor", "total_amount", "start_date", "description", "created_at"]);
  }
}

// --- CORE ENDPOINTS ---

/**
 * 1. Login User
 */
function loginUser(username, password, ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  var sheet = activeSs.getSheetByName("Users");
  var values = sheet.getDataRange().getValues();
  
  for (var i = 1; i < values.length; i++) {
    var dbUser = values[i][1];
    var dbPass = values[i][2];
    if (dbUser === username && dbPass === password) {
      return {
        success: true,
        data: { id: values[i][0], username: dbUser }
      };
    }
  }
  return { success: false, message: "Username atau password salah!" };
}

/**
 * 2. Get Dashboard Data
 */
function getDashboardData(ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  var sheetTrx = activeSs.getSheetByName("Transactions");
  var trxValues = sheetTrx.getDataRange().getValues();
  
  var income = 0;
  var expense = 0;
  var debt = 0;
  var receivable = 0;
  var paidInstallments = 0;
  
  // Hitung akumulasi tipe transaksi
  for (var i = 1; i < trxValues.length; i++) {
    var type = trxValues[i][2];
    var amount = parseFloat(trxValues[i][5]) || 0;
    
    if (type === "Pendapatan") income += amount;
    else if (type === "Pengeluaran") expense += amount;
    else if (type === "Utang") debt += amount;
    else if (type === "Piutang") receivable += amount;
    else if (type === "Cicilan") paidInstallments += amount;
  }
  
  // Dapatkan data sisa cicilan aktif
  var installmentsRes = getInstallments(activeSs);
  var installments = installmentsRes.success ? installmentsRes.data : [];
  var installmentOutstanding = 0;
  
  for (var j = 0; j < installments.length; j++) {
    if (installments[j].status !== "Lunas") {
      installmentOutstanding += installments[j].remaining;
    }
  }
  
  // Saldo bersih cash = Pendapatan + Utang (menerima uang) - Pengeluaran - Piutang (memberi uang) - Cicilan (bayar cicilan)
  var balance = income + debt - expense - receivable - paidInstallments;
  
  return {
    success: true,
    data: {
      balance: balance,
      income: income,
      expense: expense,
      debt: debt,
      receivable: receivable,
      installmentOutstanding: installmentOutstanding
    }
  };
}

/**
 * 3. Get Transactions List
 */
function getTransactions(ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  var sheet = activeSs.getSheetByName("Transactions");
  var values = sheet.getDataRange().getValues();
  var transactions = [];
  
  // Format menjadi array of objects
  for (var i = values.length - 1; i >= 1; i--) { // Reverse order agar yang terbaru tampil di atas
    if (!values[i][0]) continue; // Skip jika baris kosong
    transactions.push({
      id: values[i][0].toString(),
      date: values[i][1],
      type: values[i][2],
      category: values[i][3],
      description: values[i][4],
      amount: parseFloat(values[i][5]) || 0,
      installment_id: values[i][6] ? values[i][6].toString() : null
    });
  }
  
  return { success: true, data: transactions };
}

/**
 * 4. Get Installments List
 */
function getInstallments(ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  var sheetInst = activeSs.getSheetByName("Installments");
  var instValues = sheetInst.getDataRange().getValues();
  
  var sheetTrx = activeSs.getSheetByName("Transactions");
  var trxValues = sheetTrx.getDataRange().getValues();
  
  // Hitung total cicilan terbayar per ID
  var paidMap = {};
  for (var i = 1; i < trxValues.length; i++) {
    var type = trxValues[i][2];
    var instId = trxValues[i][6];
    var amount = parseFloat(trxValues[i][5]) || 0;
    
    if (type === "Cicilan" && instId) {
      var idStr = instId.toString();
      if (!paidMap[idStr]) paidMap[idStr] = 0;
      paidMap[idStr] += amount;
    }
  }
  
  var installments = [];
  for (var j = 1; j < instValues.length; j++) {
    if (!instValues[j][0]) continue;
    
    var id = instValues[j][0].toString();
    var name = instValues[j][1];
    var creditor = instValues[j][2];
    var totalAmount = parseFloat(instValues[j][3]) || 0;
    var startDate = instValues[j][4];
    var description = instValues[j][5];
    
    var paidAmount = paidMap[id] || 0;
    var remaining = totalAmount - paidAmount;
    if (remaining < 0) remaining = 0; // Guard
    var status = remaining <= 0 ? "Lunas" : "Berjalan";
    
    installments.push({
      id: id,
      name: name,
      creditor: creditor,
      total_amount: totalAmount,
      paid_amount: paidAmount,
      remaining: remaining,
      start_date: startDate,
      status: status,
      description: description
    });
  }
  
  return { success: true, data: installments };
}

/**
 * 5. Add Transaction
 */
function addTransaction(data, ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  var sheet = activeSs.getSheetByName("Transactions");
  var id = "trx_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
  
  // Format data
  var date = data.date || new Date().toISOString().substring(0, 10);
  var type = data.type;
  var category = data.category;
  var description = data.description || "";
  var amount = parseFloat(data.amount) || 0;
  var installmentId = data.installment_id || "";
  var createdAt = new Date().toISOString();
  
  sheet.appendRow([id, date, type, category, description, amount, installmentId, createdAt]);
  
  return { success: true, id: id };
}

/**
 * 6. Delete Transaction
 */
function deleteTransaction(id, ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  var sheet = activeSs.getSheetByName("Transactions");
  var values = sheet.getDataRange().getValues();
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0].toString() === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: "Transaksi tidak ditemukan." };
}

/**
 * 7. Add Installment
 */
function addInstallment(data, ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  var sheet = activeSs.getSheetByName("Installments");
  var id = "inst_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
  
  var name = data.name;
  var creditor = data.creditor;
  var totalAmount = parseFloat(data.total_amount) || 0;
  var startDate = data.start_date || new Date().toISOString().substring(0, 10);
  var description = data.description || "";
  var createdAt = new Date().toISOString();
  
  sheet.appendRow([id, name, creditor, totalAmount, startDate, description, createdAt]);
  
  return { success: true, id: id };
}

/**
 * 8. Delete Installment
 */
function deleteInstallment(id, ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  var sheet = activeSs.getSheetByName("Installments");
  var values = sheet.getDataRange().getValues();
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0].toString() === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: "Data cicilan tidak ditemukan." };
}

/**
 * 9. Reset Database (Transactions & Installments)
 */
function resetDatabase(ss) {
  var activeSs = getActiveSs(ss);
  initSheets(activeSs);
  
  // Reset Transactions
  var sheetTrx = activeSs.getSheetByName("Transactions");
  if (sheetTrx) {
    sheetTrx.clear();
    sheetTrx.appendRow(["id", "date", "type", "category", "description", "amount", "installment_id", "created_at"]);
  }
  
  // Reset Installments
  var sheetInst = activeSs.getSheetByName("Installments");
  if (sheetInst) {
    sheetInst.clear();
    sheetInst.appendRow(["id", "name", "creditor", "total_amount", "start_date", "description", "created_at"]);
  }
  
  return { success: true };
}
