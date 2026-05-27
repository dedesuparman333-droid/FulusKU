/* Rahaza Digital Premium Webapps Builder by ddsuparman */

/**
 * Run this function ONCE from the Apps Script Editor to initialize the database.
 * It will create all necessary sheets, insert default data, and update columns if needed.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let message = "Setup Database Completed:\n";

  // 1. Setup USERS Sheet
  let usersSheet = ss.getSheetByName("USERS");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("USERS");
    usersSheet.appendRow(["id", "username", "password", "role", "status", "created_at"]);
    usersSheet.appendRow(["usr-1", "admin", "admin123", "Administrator", "Active", new Date()]);
    usersSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    message += "- USERS sheet created (Default Login: admin / admin123)\n";
  } else {
    message += "- USERS sheet already exists\n";
  }

  // 2. Setup TRANSACTIONS Sheet
  let trxSheet = ss.getSheetByName("TRANSACTIONS");
  if (!trxSheet) {
    trxSheet = ss.insertSheet("TRANSACTIONS");
    trxSheet.appendRow(["id", "date", "type", "category", "amount", "description", "status", "timestamp"]);
    trxSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    message += "- TRANSACTIONS sheet created\n";
  } else {
    message += "- TRANSACTIONS sheet already exists\n";
  }

  // 3. Setup INSTALLMENTS Sheet
  let instSheet = ss.getSheetByName("INSTALLMENTS");
  const instHeaders = ["id", "name", "creditor", "total_amount", "paid_amount", "remaining", "start_date", "due_date", "status", "description", "timestamp"];
  
  if (!instSheet) {
    instSheet = ss.insertSheet("INSTALLMENTS");
    instSheet.appendRow(instHeaders);
    instSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    message += "- INSTALLMENTS sheet created with 'due_date' column\n";
  } else {
    // Upgrade existing sheet safely if needed
    const currentHeaders = instSheet.getRange(1, 1, 1, instSheet.getLastColumn()).getValues()[0];
    const hasDueDate = currentHeaders.some(h => String(h).trim().toLowerCase() === "due_date");
    
    if (!hasDueDate) {
      // Find where start_date is (index 6, column 7)
      let startDateIdx = currentHeaders.findIndex(h => String(h).trim().toLowerCase() === "start_date");
      if (startDateIdx === -1) startDateIdx = 6; // default fallback
      
      const insertCol = startDateIdx + 2; // Insert after start_date (column index is idx + 1, so insert after it is idx + 2)
      instSheet.insertColumnAfter(startDateIdx + 1);
      instSheet.getRange(1, insertCol).setValue("due_date");
      instSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
      message += "- INSTALLMENTS sheet upgraded: added 'due_date' column successfully\n";
    } else {
      message += "- INSTALLMENTS sheet already has 'due_date' column\n";
    }
  }

  Logger.log(message);
  Browser.msgBox("Database Setup Result", message.replace(/\n/g, "\\n"), Browser.Buttons.OK);
}
