// =================================================================
// MIZAM CRM — Code.gs
// =================================================================

// =================================================================
// תשתית — גיליונות ונתונים
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

// =================================================================
// הקמת גיליונות ראשונית — הרץ פעם אחת בלבד
// =================================================================

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
// =================================================================
// עדכון גליונות
// =================================================================
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
// =================================================================
// 1. סורק מיילים אוטומטי
// =================================================================

function checkAndSendEmails() {
  const threads = GmailApp.search("is:unread category:primary");
  const msgSheet = getSheet("Messages");
  const threadSheet = getSheet("Threads");
  const existingIds = getExistingMsgIds();

  // מיילים של מנהלים — לסינון לופ
  const managerSheet = getSheet("Managers");
  const managerEmails = managerSheet.getDataRange().getValues().slice(1).map(r => r[0]);
  managerEmails.push(CONFIG.MIZAM_EMAIL);

  for (let i = threads.length - 1; i >= 0; i--) {
    const thread = threads[i];
    const messages = thread.getMessages();

    for (const msg of messages) {
      const msgId = msg.getId();
      if (existingIds.includes(msgId)) continue;

      const fromRaw = msg.getFrom();
      const emailMatch = fromRaw.match(/<([^>]+)>/) || [null, fromRaw.trim()];
      const senderEmail = emailMatch[1].trim();

      // סינון לופ — מיילים שיצאו מאיתנו
      if (managerEmails.includes(senderEmail)) continue;

      const senderName = fromRaw.replace(/<[^>]+>/, "").replace(/"/g, "").trim() || "לקוח ללא שם";
      const subject = msg.getSubject() || "(ללא נושא)";
      //const body = msg.getPlainBody().substring(0, 1500);
      const body = cleanEmailBody(msg.getPlainBody()).substring(0, 1500);
      const threadId = thread.getId();
      const date = msg.getDate();

      // שמירה ב-Contacts
      upsertContact(senderEmail, senderName);

      // שמירה ב-Threads (רק אם לא קיים)
      upsertThread(threadId, subject, senderEmail);

      // שליחה לטלגרם
      const telegramMsgId = sendEmailToTelegram(senderName, senderEmail, subject, body, msgId, threadId);
      // שליחת התרעת דחיפה מיידית לכל המחשבים דרך OneSignal
        sendPushNotification(senderName, subject);


      // שמירה ב-Messages
      msgSheet.appendRow([
        msgId, threadId, senderEmail, CONFIG.MIZAM_EMAIL,
        "Incoming", "חדש", "", telegramMsgId || "",
        date, subject, body
      ]);

      existingIds.push(msgId); // מניעת כפילות בריצה
    }
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
// =================================================================
 //התראות
 // =================================================================
 // שולח התראת דחיפה לכל המחשבים הרשומים ב-OneSignal
 // @param {string} senderName שם השולח
 // @param {string} subject נושא המייל

function sendPushNotification(senderName, subject) {
  try {
    UrlFetchApp.fetch("https://onesignal.com", {
      method: "POST",
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Authorization": "Basic " + CONFIG.ONESIGNAL_API_KEY
      }, 
      muteHttpExceptions: true,
      payload: JSON.stringify({ 
        app_id: CONFIG.ONESIGNAL_APP_ID, 
        included_segments: ["Subscribed Users"], // שולח לכל המחשבים הרשומים 
        headings: { "en": "🔔 פנייה חדשה במערכת!" },
        contents: { "en": "מאת: " + senderName + "\nנושא: " + subject } 
      })
    }); 
  } catch (err) {
    Logger.log("שגיאה בשליחת התראה ל-OneSignal: " + err);
  }
}
//בדיקת התראות
function testMyPushToSpecificComputer() {
  try {
    // ה-Subscription ID מעמוד ה-Subscription ב-OneSignal (לא ה-App ID!)
    const targetComputerId = CONFIG.TEST_SUBSCRIPTION_ID || "PASTE_YOUR_SUBSCRIPTION_ID_HERE";

    const response = UrlFetchApp.fetch("https://api.onesignal.com/notifications", {
      method: "POST",
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Authorization": "Key " + CONFIG.ONESIGNAL_API_KEY
      },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        app_id: CONFIG.ONESIGNAL_APP_ID,
        include_subscription_ids: [targetComputerId],
        headings: { "en": "בדיקה פנימית למחשב שלך" },
        contents: { "en": "ההתראה עובדת ישירות מגוגל לאלקטרון." }
      })
    });

    Logger.log("קוד תגובה: " + response.getResponseCode());
    Logger.log("תגובה: " + response.getContentText());
  } catch (err) {
    Logger.log("שגיאה: " + err);
  }
}


// =================================================================
// 2. Telegram — שליחה
// =================================================================

function sendEmailToTelegram(name, email, subject, body, msgId, threadId) {
  const text =
    "📬 *מייל נכנס חדש*\n" +
    "👤 *שולח:* " + escapeMd(name) + " \\(" + escapeMd(email) + "\\)\n" +
    "📌 *נושא:* " + escapeMd(subject) + "\n" +
    "🏷️ *סטטוס:* 🔵 חדש\n" +
    "──────────────────\n" +
    escapeMd(body.substring(0, 800));

  const keyboard = {
    inline_keyboard: [
      [
        { text: "✍️ מענה מטלגרם", callback_data: "reply_" + msgId },
        { text: "💻 פתח בווב", url: CONFIG.WEB_APP_URL + "?thread=" + threadId }
      ],
      [
        { text: "🏷️ שנה תווית", callback_data: "label_thread_" + threadId }
      ]
    ]
  };

  const res = telegramPost("sendMessage", {
    chat_id: CONFIG.TELEGRAM_CHAT_ID,
    text: text,
    parse_mode: "MarkdownV2",
    reply_markup: JSON.stringify(keyboard)
  });

  return res && res.ok ? res.result.message_id : null;
}

function escapeMd(text) {
  if (!text) return "";
  return String(text).replace(/[_*[\]()~`>#+\-=|{}.!\\]/g, "\\$&");
}

// =================================================================
// 3. Telegram — Webhook
// =================================================================

function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    if (update.callback_query) handleCallbackQuery(update.callback_query);
    if (update.message)        handleTelegramMessage(update.message);
  } catch (err) {
    Logger.log("doPost error: " + err.toString());
  }
}

function handleCallbackQuery(cb) {
  const data        = cb.data;
  const chatId      = cb.message.chat.id;
  const tgMsgId     = cb.message.message_id;
  const cbId        = cb.id;
  const managerName = CONFIG.DEFAULT_MANAGER;

  // — מענה מטלגרם: שלב 1 — בקשת טקסט
  if (data.startsWith("reply_")) {
    const msgId = data.replace("reply_", "");
    PropertiesService.getScriptProperties().setProperty("reply_" + chatId, msgId);
    telegramPost("sendMessage", {
      chat_id: chatId,
      text: "✍️ כתבי את התשובה שלך. שלחי /cancel לביטול."
    });
    answerCb(cbId, "ממתין לתשובה...");
    return;
  }

  // — בחירת תווית לשרשור
  if (data.startsWith("label_thread_")) {
    const threadId = data.replace("label_thread_", "");
    const labels = getSheet("Labels").getDataRange().getValues().slice(1);
    const keyboard = buildLabelKeyboard(labels, "set_thread_label_" + threadId + "_");
    keyboard.push([{ text: "↩️ ביטול", callback_data: "back_" + threadId + "_" + tgMsgId }]);
    editMarkup(chatId, tgMsgId, { inline_keyboard: keyboard });
    answerCb(cbId, "בחרי תווית...");
    return;
  }

  // — קביעת תווית לשרשור
  if (data.startsWith("set_thread_label_")) {
    const withoutPrefix = data.replace("set_thread_label_", "");
    const underscoreIdx = withoutPrefix.indexOf("_");
    const threadId = withoutPrefix.substring(0, underscoreIdx);
    const label    = withoutPrefix.substring(underscoreIdx + 1);

    updateThreadLabel(threadId, label, managerName);

    const defaultKb = buildDefaultKeyboard(threadId, tgMsgId);
    const originalText = cb.message.text || "";
    const newText = "🏷️ *תווית עודכנה:* " + escapeMd(label) + " ע\"י " + escapeMd(managerName) + "\n" + originalText;
    editText(chatId, tgMsgId, newText, defaultKb);
    answerCb(cbId, "תווית עודכנה ✅");
    return;
  }

  // — חזרה
  if (data.startsWith("back_")) {
    const parts    = data.split("_");
    const threadId = parts[1];
    editMarkup(chatId, tgMsgId, buildDefaultKeyboard(threadId, tgMsgId));
    answerCb(cbId, "בוטל");
    return;
  }
}

function handleTelegramMessage(message) {
  if (!message || !message.text) return;
  const text   = message.text;
  const chatId = message.chat.id;
  const props  = PropertiesService.getScriptProperties();
  const managerName = CONFIG.DEFAULT_MANAGER;

  if (text === "/cancel") {
    props.deleteProperty("reply_" + chatId);
    telegramPost("sendMessage", { chat_id: chatId, text: "❌ הפעולה בוטלה." });
    return;
  }

  const activeMsgId = props.getProperty("reply_" + chatId);
  if (!activeMsgId) return;

  try {
    const originalMsg = GmailApp.getMessageById(activeMsgId);
    const thread      = originalMsg.getThread();
    originalMsg.reply(text, {
  name: "מיזם השבת",
  htmlBody: '<div dir="rtl" style="text-align:right;font-family:Arial">' +
            text.replace(/\n/g, '<br>') + '</div>'
});
  
    // שומר את התשובה כהודעה יוצאת חדשה בשיטס
    const updatedMsgs = thread.getMessages();
    const sentMsg     = updatedMsgs[updatedMsgs.length - 1];
    getSheet("Messages").appendRow([
      sentMsg.getId(), thread.getId(),
      CONFIG.MIZAM_EMAIL, extractEmail(originalMsg.getFrom()),
      "Outgoing", "נענה ע\"י " + managerName, managerName,
      "", new Date(), "Re: " + originalMsg.getSubject(), text, ""
    ]);

    // עדכון סטטוס ב-Messages + Threads
    updateMessageStatus(activeMsgId, "נענה ע\"י " + managerName, managerName);
    updateThreadLabel(thread.getId(), "נענה", managerName);

 // עדכון הודעת הטלגרם בקבוצה
    // שולפים פרטי ההודעה המקורית מהשיטס לפורמט מלא
    const rows = getSheet("Messages").getDataRange().getValues().slice(1);
    for (const row of rows) {
      if (row[0] === activeMsgId && row[7]) {
        // row: msgId, threadId, fromEmail, toEmail, direction, status, handledBy, telegramMsgId, date, subject, body
        const senderEmail = row[2] || "";
        const subject     = row[9] || "";
        const body        = String(row[10] || "").substring(0, 600);
        const threadId    = row[1];
        const tgMsgId     = row[7];
 
        // שולפים שם איש קשר אם קיים
        const contacts    = getSheet("Contacts").getDataRange().getValues().slice(1);
        const contact     = contacts.find(c => c[0] === senderEmail);
        const senderName  = contact ? contact[1] : senderEmail;
 
        const updatedText =
          "👤 *" + escapeMd(senderName) + "* \\(" + escapeMd(senderEmail) + "\\)\n" +
          "📌 *" + escapeMd(subject) + "*\n" +
          "──────────────────\n" +
          escapeMd(body) + "\n" +
          "──────────────────\n" +
          "✍️ *תשובה:* " + "\\(" + escapeMd(managerName) + "\\)\n" +
          "──────────────────\n" +
          escapeMd(text) + "\n" +
          
          "🏷️ *סטטוס:* נענה";
 
        editText(
          CONFIG.TELEGRAM_CHAT_ID, tgMsgId, updatedText,
          buildDefaultKeyboard(threadId, tgMsgId)
        );
        break;
      }
    }

    props.deleteProperty("reply_" + chatId);
    telegramPost("sendMessage", { chat_id: chatId, text: "📧 המייל נשלח בהצלחה!" });

  }
   catch (err) {
    telegramPost("sendMessage", { chat_id: chatId, text: "❌ שגיאה: " + err.toString() });
  }
}

// עזרי טלגרם
function buildLabelKeyboard(labels, prefix) {
  const keyboard = [];
  let row = [];
  labels.forEach((label, i) => {
    if (!label[0]) return;
    row.push({ text: label[0], callback_data: prefix + label[0] });
    if (row.length === 2 || i === labels.length - 1) { keyboard.push(row); row = []; }
  });
  return keyboard;
}

function buildDefaultKeyboard(threadId, tgMsgId) {
  return {
    inline_keyboard: [
      [
        { text: "✍️ מענה", callback_data: "reply_" + tgMsgId },
        { text: "💻 פתח בווב", url: CONFIG.WEB_APP_URL + "?thread=" + threadId }
      ],
      [{ text: "🏷️ שנה תווית", callback_data: "label_thread_" + threadId }]
    ]
  };
}

function telegramPost(method, payload) {
  try {
    const res = UrlFetchApp.fetch(
      "https://api.telegram.org/bot" + CONFIG.TELEGRAM_BOT_TOKEN + "/" + method,
      { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true }
    );
    return JSON.parse(res.getContentText());
  } catch (e) {
    Logger.log("Telegram error (" + method + "): " + e.toString());
    return null;
  }
}

function editMarkup(chatId, msgId, markup) {
  telegramPost("editMessageReplyMarkup", { chat_id: chatId, message_id: msgId, reply_markup: JSON.stringify(markup) });
}

function editText(chatId, msgId, text, markup) {
  telegramPost("editMessageText", { chat_id: chatId, message_id: msgId, text: text, parse_mode: "MarkdownV2", reply_markup: JSON.stringify(markup) });
}

function answerCb(id, text) {
  telegramPost("answerCallbackQuery", { callback_query_id: id, text: text });
}

// =================================================================
// 4. עדכוני נתונים
// =================================================================

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

// =================================================================
// 5. Web App
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


function sendCrmReplyFromWeb(threadId, replyText, managerName) {
  managerName = managerName || "מנהל";

  const thread   = GmailApp.getThreadById(threadId);
  const messages = thread.getMessages();
  if (!messages.length) throw new Error("שרשור ריק");

//קובעים את השולח והמקבל (לקוח) לפי הודעה אפס בלבד
  const firstMsg    = messages[0];
  const firstFrom   = extractEmail(firstMsg.getFrom());
  const clientEmail = firstFrom === CONFIG.MIZAM_EMAIL 
    ? extractEmail(firstMsg.getTo()) 
    : firstFrom;

  // 1. מחפש את הודעת הלקוח האחרונה
  let lastClientMsg = null;
  for (let i = messages.length - 1; i >= 0; i--) {
    if (extractEmail(messages[i].getFrom()) === clientEmail) {
      lastClientMsg = messages[i];
      break;
    }
  }

  // הגדרת עיצוב המייל 
  const mailOptions = {
    name: "מיזם השבת",
    htmlBody: '<div dir="rtl" style="text-align:right;font-family:Arial">' + 
              replyText.replace(/\n/g, '<br>') + '</div>'
  };

  // 2. בדיקה ושליחה
  if (lastClientMsg) {
    // אם נמצאה הודעת לקוח - רפלי רגיל עליה
    lastClientMsg.reply(replyText, mailOptions);
  } else {
    // אם לא נמצאה - replyAll על הודעה 0, ומגדירים אותה כעוגן בשביל השיטס
    firstMsg.replyAll(replyText, mailOptions);
    lastClientMsg = firstMsg;
  }

  // 3. רישום בשיטס (רץ רק אחרי שהשליחה הצליחה)
  const sentMsgs = thread.getMessages();
  const sentMsg  = sentMsgs[sentMsgs.length - 1];

  getSheet("Messages").appendRow([
    sentMsg.getId(),
    threadId,
    CONFIG.MIZAM_EMAIL,
    clientEmail,
    "Outgoing",
    "נענה ע\"י " + managerName,
    managerName,
    "",
    new Date(),
    "Re: " + firstMsg.getSubject(),
    replyText
  ]);

  updateThreadLabel(threadId, "נענה", managerName);
}


function sendNewEmailFromWeb(toEmail, subject, body, managerName) {
  managerName = managerName || "מנהל";

  // שליחת מייל חדש
 GmailApp.sendEmail(toEmail, subject, body, {
  name: "מיזם השבת",
  htmlBody: '<div dir="rtl" style="text-align:right;font-family:Arial">' + 
            body.replace(/\n/g, '<br>') + '</div>'
});

  // מחפשים את השרשור שנוצר
  Utilities.sleep(2000); // המתנה קצרה כדי שג'ימייל יעבד
  const sent    = GmailApp.search("to:" + toEmail + " subject:\"" + subject + "\" in:sent", 0, 1);
  const thread  = sent.length ? sent[0] : null;
  const threadId = thread ? thread.getId() : "new_" + Date.now();
  const sentMsg  = thread ? thread.getMessages()[thread.getMessages().length - 1] : null;
  const msgId    = sentMsg ? sentMsg.getId() : "msg_" + Date.now();

  // Contacts — אם לא קיים
  const toName = toEmail.split("@")[0];
  upsertContact(toEmail, toName);

  // Thread
  upsertThread(threadId, subject, toEmail);

  // Message
  getSheet("Messages").appendRow([
    msgId, threadId,
    CONFIG.MIZAM_EMAIL, toEmail,
    "Outgoing", "נשלח ע\"י " + managerName, managerName,
    "", new Date(), subject, body
  ]);

  updateThreadLabel(threadId, "בטיפול", managerName);
  return threadId;
}

function extractEmail(raw) {
  if (!raw) return "";
  const match = raw.match(/<([^>]+)>/);
  return match ? match[1].trim() : raw.trim();
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

function addNote(threadId, body, type, targetManager, createdBy) {
  const noteId = "note_" + Date.now();
  getSheet("Notes").appendRow([
    noteId, threadId, body, type,
    targetManager || "", createdBy, new Date(), ""
  ]);
  return noteId;
}

function completeTask(noteId) {
  const sheet = getSheet("Notes");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === noteId) {
      sheet.getRange(i + 1, 4).setValue("note");   // הופך להערה רגילה
      sheet.getRange(i + 1, 8).setValue(new Date()); // doneAt
      return true;
    }
  }
  return false;
}

function editNote(noteId, newBody, requestedBy) {
  const sheet = getSheet("Notes");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === noteId && rows[i][5] === requestedBy) {
      sheet.getRange(i + 1, 3).setValue(newBody);
      return true;
    }
  }
  return false; // לא נמצא או לא בעלים
}

function deleteNote(noteId, requestedBy) {
  const sheet = getSheet("Notes");
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === noteId && rows[i][5] === requestedBy) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
//-----
// שייך לדאטה או ליד אדפט ת'ירד ומסג'
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


// =================================================================
// 6. הגדרות ועזרים
// =================================================================
//פונקצית עזר לניקוי היסטוריה בלי בדיקה אם ההודעה קיימת
function cleanEmailBody(body) {
  if (!body) return "";
  
  // נשארו רק העוגנים ההכרחיים: תגובות אימוג'י, ציטוט בעברית, וציטוט באנגלית
  const regex = /(?:התגובה של[\s\S]+?בוצעה דרך Gmail|בתאריך[\s\S]+?מאת[\s\S]+?<[\s\S]+?>:|On[\s\S]+?wrote:)[\s\S]*/i;
  
  // חותך את מה שנמצא ומנקה רווחים
  return body.replace(regex, '').trim();
}

function setTelegramWebhook() {
  const res = telegramPost("setWebhook", { url: CONFIG.WEB_APP_URL });
  Logger.log(res.ok ? "✅ Webhook חובר: " + CONFIG.WEB_APP_URL : "❌ שגיאה: " + JSON.stringify(res));
}

/**
 * backfill — משיכת שרשורים ישנים
 * שנה את maxThreads לפי הצורך
 */
function backfillOldThreads() {
  const maxThreads = 50;
  const threads = GmailApp.search("in:inbox category:primary", 0, maxThreads);
  const existingIds = getExistingMsgIds();
  const managerEmails = getSheet("Managers").getDataRange().getValues().slice(1).map(r => r[0]);
  managerEmails.push(CONFIG.MIZAM_EMAIL);

  let added = 0;
  for (let i = threads.length - 1; i >= 0; i--) {
    const thread   = threads[i];
    const messages = thread.getMessages();
    for (const msg of messages) {
      const msgId = msg.getId();
      if (existingIds.includes(msgId)) continue;
      const fromRaw    = msg.getFrom();
      const emailMatch = fromRaw.match(/<([^>]+)>/) || [null, fromRaw.trim()];
      const senderEmail = emailMatch[1].trim();
      if (managerEmails.includes(senderEmail)) continue;

      const senderName = fromRaw.replace(/<[^>]+>/, "").replace(/"/g, "").trim() || "לקוח ללא שם";
      const subject    = msg.getSubject() || "(ללא נושא)";
      const body       = msg.getPlainBody().substring(0, 1500);
      const threadId   = thread.getId();

      upsertContact(senderEmail, senderName);
      upsertThread(threadId, subject, senderEmail);
      getSheet("Messages").appendRow([
        msgId, threadId, senderEmail, CONFIG.MIZAM_EMAIL,
        "Incoming", "ארכיון", "", "",
        msg.getDate(), subject, body
      ]);
      existingIds.push(msgId);
      added++;
    }
  }
  Logger.log("✅ backfill הוסיף " + added + " הודעות.");
  }
function checkWebhook() {
  const res = telegramPost("getWebhookInfo", {});
  Logger.log(JSON.stringify(res));
}


 function testCreateTelegramTopic() {
  const chatId = CONFIG.TELEGRAM_CHAT_ID;
  
  // 1. בקשה מטלגרם לפתוח נושא חדש
  Logger.log("מנסה לפתוח נושא חדש...");
  const createRes = telegramPost("createForumTopic", {
    chat_id: chatId,
    name: "🧪 טסט: לקוח חדש",
    icon_custom_emoji_id: "" // אפשר להשאיר ריק
  });
  
  Logger.log("תשובת יצירת נושא: " + JSON.stringify(createRes));
  
  // 2. אם הנושא נפתח בהצלחה, נשלח אליו הודעה
  if (createRes && createRes.ok) {
    const topicThreadId = createRes.result.message_thread_id;
    Logger.log("✅ נושא נפתח בהצלחה! מזהה הנושא: " + topicThreadId);
    
    // שליחת הודעה לתוך התיקייה הספציפית
    const sendRes = telegramPost("sendMessage", {
      chat_id: chatId,
      message_thread_id: topicThreadId, // זה הפרמטר שמנתב את ההודעה לנושא הנכון
      text: "שלום! 🚀\nזהו שרשור ניסיון שנפתח ישירות מ-Apps Script."
    });
    
    Logger.log("תשובת שליחת הודעה: " + JSON.stringify(sendRes));
  } else {
    Logger.log("❌ שגיאה בפתיחת הנושא. האם הפעלת 'Topics' בהגדרות הקבוצה?");
  }
}
 
/**
 * סנכרון שרשורים קיימים לקבוצת טלגרם מבוססת-נושאים (Topics)
 */
function backfillThreadsToTelegramTopics() {
  const threadSheet = getSheet("Threads");
  const msgSheet = getSheet("Messages");
  const contactSheet = getSheet("Contacts");
  const noteSheet = getSheet("Notes"); // נניח שקיים, אחרת יחזיר null
  
  const threadRows = threadSheet.getDataRange().getValues();
  if (threadRows.length <= 1) {
    Logger.log("אין שרשורים לסנכרן.");
    return;
  }
  
  const allMessages = msgSheet.getDataRange().getValues().slice(1);
  const allContacts = contactSheet.getDataRange().getValues().slice(1);
  const allNotes = noteSheet ? noteSheet.getDataRange().getValues().slice(1) : [];
  
  for (let i = 1; i < threadRows.length; i++) {
    const threadId = threadRows[i][0];
    const displaySubject = threadRows[i][1] || "(ללא נושא)";
    const contactEmail = threadRows[i][2] || "לקוח לא ידוע";
    const deletedBy = threadRows[i][6]; 
    const existingTopicId = threadRows[i][7]; // עמודה H
    
    // סינון שרשורים שנמחקו או שכבר סונכרנו
    if (deletedBy || existingTopicId) continue;
    
    // נקודה 3: חילוץ איש קשר במקום מייל לשם הנושא
    let contactName = contactEmail;
    const contactRow = allContacts.find(c => c[0] === contactEmail);
    if (contactRow && contactRow[1]) {
      contactName = contactRow[1];
    }
    
    Logger.log("מנסה לפתוח נושא עבור: " + contactName);
    
    // פתיחת הנושא בטלגרם (מוגבל ל-128 תווים, נחתוך ל-80 ליתר ביטחון)
    const topicName = (contactName + " | " + displaySubject).substring(0, 80);
    const createRes = telegramPost("createForumTopic", {
      chat_id: CONFIG.TELEGRAM_CHAT_ID,
      name: topicName
    });
    
    if (createRes && createRes.ok) {
      const topicId = createRes.result.message_thread_id;
      Logger.log("✅ נושא נפתח: " + topicId);
      
      // שמירת המזהה בשיטס (עמודה 8)
      threadSheet.getRange(i + 1, 8).setValue(topicId);
      
      // איסוף הודעות (נקודה 1) והערות (נקודה 5) לאותו שרשור
      let allItems = [];
      
      allMessages.filter(m => m[1] === threadId).forEach(m => {
        allItems.push({ time: new Date(m[8]), type: 'msg', data: m });
      });
      
      allNotes.filter(n => n[1] === threadId).forEach(n => {
        allItems.push({ time: new Date(n[6]), type: 'note', data: n });
      });
      
      // מיון כרונולוגי לפי זמן (נקודה 6)
      allItems.sort((a, b) => a.time - b.time);
      
      // שליחת כל פריט לטלגרם לפי הסדר
      for (const item of allItems) {
        let textToSend = "";
        const formattedDate = Utilities.formatDate(item.time, "Asia/Jerusalem", "dd/MM/yyyy HH:mm"); // נקודה 2
        
        if (item.type === 'msg') {
          const msgData = item.data;
          const direction = msgData[4]; // Incoming / Outgoing
          const handledBy = msgData[6] || "מנהל"; // אינדקס 6 הוא מי טיפל
          const bodyPreview = String(msgData[10]).substring(0, 800); // חיתוך מגבלת טלגרם
          
          if (direction === "Outgoing") {
            // נקודה 4: הוספת שם המנהל בהודעות יוצאות
            textToSend = 
              "📤 *הודעה יוצאת*\n" +
              "✍️ מאת: *" + escapeMd(handledBy) + "*\n" +
              "🕒 " + escapeMd(formattedDate) + "\n" +
              "──────────────────\n" +
              escapeMd(bodyPreview);
          } else {
            textToSend = 
              "📥 *הודעה נכנסת*\n" +
              "👤 מאת: *" + escapeMd(contactName) + "*\n" +
              "🕒 " + escapeMd(formattedDate) + "\n" +
              "──────────────────\n" +
              escapeMd(bodyPreview);
          }
        } 
        else if (item.type === 'note') {
          const noteData = item.data;
          const noteBody = noteData[2];
          const noteType = noteData[3] === 'task' ? '✅ משימה' : '📝 הערה';
          const createdBy = noteData[5] || "מערכת";
          const isDone = noteData[7] ? " (בוצע)" : "";
          
          textToSend = 
            escapeMd(noteType) + isDone + "\n" +
            "👤 מאת: *" + escapeMd(createdBy) + "*\n" +
            "🕒 " + escapeMd(formattedDate) + "\n" +
            "──────────────────\n" +
            "_" + escapeMd(noteBody) + "_";
        }
        
        // שליחה לטלגרם לתוך התיקייה הספציפית
        telegramPost("sendMessage", {
          chat_id: CONFIG.TELEGRAM_CHAT_ID,
          message_thread_id: topicId,
          text: textToSend,
          parse_mode: "MarkdownV2"
        });
        
        // השהייה של חצי שנייה בין הודעה להודעה כדי שטלגרם לא יחסום אותנו על הצפה (Rate Limit)
        Utilities.sleep(500); 
      }
      
    } else {
      Logger.log("❌ שגיאה ביצירת נושא לשרשור " + threadId + ": " + JSON.stringify(createRes));
    }
  }
  Logger.log("סיום סנכרון שרשורים מורחב.");
}
