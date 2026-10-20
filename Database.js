// =================================================================
// Database.js — תשתית גיליונות ונתונים
// =================================================================

function getSheet(name) {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(name);
}

/**
 * מחזיר מנהל לפי email, או null אם לא נמצא
 */
function getManagerByEmail(email) {
  const sheet = getSheet("Managers");
  const rows = sheet.getDataRange().getValues().slice(1);
  for (const row of rows) {
    if (row[0] === email) {
      return { email: row[0], displayName: row[1], telegramUsername: row[2], color: row[3] };
    }
  }
  return null;
}

/**
 * האם כתובת המייל שייכת למנהל (כולל מייל המיזם)
 */
function isManagerEmail(email) {
  if (email === CONFIG.MIZAM_EMAIL) return true;
  return getManagerByEmail(email) !== null;
}

/**
 * מחזיר את ה-msgId-ים הקיימים בגיליון Messages
 */
function getExistingMsgIds() {
  const sheet = getSheet("Messages");
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
}

/**
 * הקמת גיליונות ראשונית — הרץ פעם אחת בלבד
 */
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const sheets = {
    "Messages": ["msgId", "threadId", "fromEmail", "toEmail", "direction", "status", "handledBy", "telegramMsgId", "date", "subject", "body"],
    "Threads":  ["threadId", "displaySubject", "contactEmail", "label", "notes", "lastActivity"],
    "Managers": ["email", "displayName", "telegramUsername", "color"],
    "Labels":   ["labelName", "color"],
    "Contacts": ["email", "displayName", "role", "createdAt"]
  };

  for (const [name, headers] of Object.entries(sheets)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clearContents();
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#121b22")
      .setFontColor("#ffffff");
  }

  // נתוני מנהלים ראשוניים (דוגמה)
  const mgr = ss.getSheetByName("Managers");
  mgr.getRange(2, 1, 3, 4).setValues([
    [CONFIG.MIZAM_EMAIL || "manager1@example.com", CONFIG.DEFAULT_MANAGER || "מנהל ראשי", "", "#00a884"],
    ["manager2@example.com", "מנהל 2", "", "#7B68EE"],
    ["manager3@example.com", "מנהל 3", "", "#FF8C00"]
  ]);

  // תוויות ראשוניות
  const lbl = ss.getSheetByName("Labels");
  lbl.getRange(2, 1, 6, 2).setValues([
    ["חדש",     "#3498db"],
    ["בטיפול",  "#f39c12"],
    ["תורם",    "#9b59b6"],
    ["זוכה",    "#27ae60"],
    ["ספאם",    "#e74c3c"],
    ["סגור",    "#95a5a6"]
  ]);

  Logger.log("✅ גיליונות הוקמו בהצלחה!");
}

/**
 * עדכון גיליונות — יצירת גיליונות חסרים
 */
function addMissingSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  const toAdd = {
    "Notes": ["noteId", "threadId", "body", "type", "targetManager", "createdBy", "createdAt", "doneAt"]
  };

  for (const [name, headers] of Object.entries(toAdd)) {
    if (ss.getSheetByName(name)) {
      Logger.log("גיליון " + name + " כבר קיים — לא נגע בו");
      continue;
    }
    const sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#121b22")
      .setFontColor("#ffffff");
    Logger.log("✅ גיליון " + name + " נוצר");
  }
}

function upsertContact(email, name) {
  const sheet = getSheet("Contacts");
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    if (emails.includes(email)) return;
  }
  sheet.appendRow([email, name, "Client", new Date()]);
}

function upsertThread(threadId, subject, contactEmail) {
  const sheet = getSheet("Threads");
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    if (ids.includes(threadId)) {
      // עדכון lastActivity
      const rowIndex = ids.indexOf(threadId) + 2;
      sheet.getRange(rowIndex, 6).setValue(new Date());
      return;
    }
  }
  sheet.appendRow([threadId, subject, contactEmail, "חדש", "", new Date()]);
}

function updateMessageStatus(msgId, status, handledBy) {
  const sheet = getSheet("Messages");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === msgId) {
      sheet.getRange(i + 1, 6).setValue(status);
      sheet.getRange(i + 1, 7).setValue(handledBy);
      return;
    }
  }
}

function updateThreadLabel(threadId, label, handledBy) {
  const sheet = getSheet("Threads");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === threadId) {
      sheet.getRange(i + 1, 4).setValue(label);
      sheet.getRange(i + 1, 6).setValue(new Date()); // lastActivity
      return;
    }
  }
}

function updateThreadDisplaySubject(threadId, newSubject) {
  const sheet = getSheet("Threads");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === threadId) {
      sheet.getRange(i + 1, 2).setValue(newSubject);
      return true;
    }
  }
  return false;
}

function updateThreadNotes(threadId, notes) {
  const sheet = getSheet("Threads");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === threadId) {
      sheet.getRange(i + 1, 5).setValue(notes);
      return true;
    }
  }
  return false;
}
