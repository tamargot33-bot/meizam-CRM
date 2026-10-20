// =================================================================
// TelegramOutgoing.js — שליחת הודעות לטלגרם וכפתורים
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
