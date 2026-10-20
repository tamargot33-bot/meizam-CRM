// =================================================================
// WebApp.js — ממשק ה-Web App וה-CRM
// =================================================================

function doGet(e) {
  const param = e && e.parameter ? e.parameter : {};

  // API: טעינת נתונים
  if (param.fetch === "data") {
    return ContentService
      .createTextOutput(loadCrmData(param.manager || ""))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // דף ראשי
  const template = HtmlService.createTemplateFromFile("index");
  template.SUPER_MANAGER = CONFIG.SUPER_MANAGER || CONFIG.DEFAULT_MANAGER || "";
  template.ONESIGNAL_APP_ID = CONFIG.ONESIGNAL_APP_ID || "";
  template.ONESIGNAL_SAFARI_ID = CONFIG.ONESIGNAL_SAFARI_ID || "";
  const html = template.evaluate();
  html.setTitle("Mizam CRM");
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function loadCrmData(currentManagerName) {
  try {
    const SUPER_MANAGER = CONFIG.SUPER_MANAGER || CONFIG.DEFAULT_MANAGER || "מנהל ראשי"; // המנהל/ת המיוחד/ת שרואה גם הודעות ושרשורים מחוקים
    
    const contacts = getSheet("Contacts").getDataRange().getValues().slice(1);
    let messages   = getSheet("Messages").getDataRange().getValues().slice(1);
    let threads    = getSheet("Threads").getDataRange().getValues().slice(1);
    const labels   = getSheet("Labels").getDataRange().getValues().slice(1);
    const notesRaw = getSheet("Notes") ? getSheet("Notes").getDataRange().getValues().slice(1) : [];
    
    // מסנן הערות: מציג אם זה לכולם, מיועד למנהל הנוכחי, או נכתב ע"י המנהל הנוכחי
    const notes = notesRaw.filter(function(r) {
      return !r[4] || r[4] === currentManagerName || r[5] === currentManagerName;
    });

    // מנהלים — שולחים רק displayName + color, ללא מייל
    const managers = getSheet("Managers").getDataRange().getValues().slice(1)
      .map(r => [r[1], r[3]]); // [displayName, color]

    // מנגנון סינון מחיקה רכה
    if (currentManagerName !== SUPER_MANAGER) {
      // ב-Threads: עמודה 7 (אינדקס 6) שומרת את מי שמחק. אם היא ריקה — השרשור פעיל.
      threads = threads.filter(r => !r[6]); 
      
      // ב-Messages: עמודה 12 (אינדקס 11) שומרת את מי שמחק. אם היא ריקה — ההודעה פעילה.
      messages = messages.filter(r => !r[11]);
    }

    return JSON.stringify({ contacts, messages, threads, managers, labels, notes });
  } 
  catch (err) {
    Logger.log("loadCrmData error: " + err.toString());
    return JSON.stringify({ contacts: [], messages: [], threads: [], managers: [], labels: [], notes: [] });
  }
}

function deleteThreadFromWeb(threadId, managerName, deleteFromGmail) {
  // מחיקה רכה — כותב שם המוחק בעמודה 7
  // אם deleteFromGmail=true — מעביר גם לאשפה בג'ימייל (רק למנהל/ת ראשי/ת)
  const sheet = getSheet("Threads");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === threadId) {
      sheet.getRange(i + 1, 7).setValue(managerName);
      if (deleteFromGmail) {
        try {
          const thread = GmailApp.getThreadById(threadId);
          if (thread) thread.moveToTrash();
        } catch(e) {
          Logger.log("Gmail trash error: " + e.toString());
        }
      }
      return true;
    }
  }
  return false;
}

function deleteMessageFromWeb(msgId, managerName, deleteFromGmail) {
  // מחיקה רכה — כותב שם המוחק בעמודה 12
  // אם deleteFromGmail=true — מעביר גם לאשפה בג'ימייל (רק למנהל/ת ראשי/ת)
  const sheet = getSheet("Messages");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === msgId) {
      sheet.getRange(i + 1, 12).setValue(managerName);
      if (deleteFromGmail) {
        try {
          const msg = GmailApp.getMessageById(msgId);
          if (msg) msg.moveToTrash();
        } catch(e) {
          Logger.log("Gmail trash error: " + e.toString());
        }
      }
      return true;
    }
  }
  return false;
}

function undeleteThread(threadId) {
  // ביטול מחיקה — מנקה את עמודת deletedBy
  const sheet = getSheet("Threads");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === threadId) {
      sheet.getRange(i + 1, 7).setValue("");
      return true;
    }
  }
  return false;
}

function undeleteMessage(msgId) {
  // ביטול מחיקה — מנקה את עמודת deletedBy
  const sheet = getSheet("Messages");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === msgId) {
      sheet.getRange(i + 1, 12).setValue("");
      return true;
    }
  }
  return false;
}
