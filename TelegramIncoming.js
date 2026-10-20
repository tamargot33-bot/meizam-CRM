// =================================================================
// TelegramIncoming.js — קבלת אירועים ומענה מטלגרם (Webhook)
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

  } catch (err) {
    telegramPost("sendMessage", { chat_id: chatId, text: "❌ שגיאה: " + err.toString() });
  }
}
