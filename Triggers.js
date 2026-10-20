// =================================================================
// Triggers.js — הגדרות טלגרם, בדיקות וסנכרון היסטוריה
// =================================================================

function setTelegramWebhook() {
  const res = telegramPost("setWebhook", { url: CONFIG.WEB_APP_URL });
  Logger.log(res.ok ? "✅ Webhook חובר: " + CONFIG.WEB_APP_URL : "❌ שגיאה: " + JSON.stringify(res));
}

function checkWebhook() {
  const res = telegramPost("getWebhookInfo", {});
  Logger.log(JSON.stringify(res));
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
    
    // חילוץ איש קשר במקום מייל לשם הנושא
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
      
      // איסוף הודעות והערות לאותו שרשור
      let allItems = [];
      
      allMessages.filter(m => m[1] === threadId).forEach(m => {
        allItems.push({ time: new Date(m[8]), type: 'msg', data: m });
      });
      
      allNotes.filter(n => n[1] === threadId).forEach(n => {
        allItems.push({ time: new Date(n[6]), type: 'note', data: n });
      });
      
      // מיון כרונולוגי לפי זמן
      allItems.sort((a, b) => a.time - b.time);
      
      // שליחת כל פריט לטלגרם לפי הסדר
      for (const item of allItems) {
        let textToSend = "";
        const formattedDate = Utilities.formatDate(item.time, "Asia/Jerusalem", "dd/MM/yyyy HH:mm");
        
        if (item.type === 'msg') {
          const msgData = item.data;
          const direction = msgData[4]; // Incoming / Outgoing
          const handledBy = msgData[6] || "מנהל"; // אינדקס 6 הוא מי טיפל
          const bodyPreview = String(msgData[10]).substring(0, 800); // חיתוך מגבלת טלגרם
          
          if (direction === "Outgoing") {
            // הוספת שם המנהל בהודעות יוצאות
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
