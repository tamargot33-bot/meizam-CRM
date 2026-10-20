// =================================================================
// Incoming.js — סריקת מיילים נכנסים והתראות
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

/**
 * שולח התראת דחיפה לכל המחשבים הרשומים ב-OneSignal
 * @param {string} senderName שם השולח
 * @param {string} subject נושא המייל
 */
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

/**
 * בדיקת התראות למחשב ספציפי
 */
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
