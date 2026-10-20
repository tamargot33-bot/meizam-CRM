// =================================================================
// Outgoing.js — שליחת מיילים יוצאים מג'ימייל
// =================================================================

function sendCrmReplyFromWeb(threadId, replyText, managerName) {
  managerName = managerName || "מנהל";

  const thread   = GmailApp.getThreadById(threadId);
  const messages = thread.getMessages();
  if (!messages.length) throw new Error("שרשור ריק");

  // קובעים את השולח והמקבל (לקוח) לפי הודעה אפס בלבד
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

/**
 * פונקציית עזר לניקוי היסטוריה מההודעה
 */
function cleanEmailBody(body) {
  if (!body) return "";
  
  // נשארו רק העוגנים ההכרחיים: תגובות אימוג'י, ציטוט בעברית, וציטוט באנגלית
  const regex = /(?:התגובה של[\s\S]+?בוצעה דרך Gmail|בתאריך[\s\S]+?מאת[\s\S]+?<[\s\S]+?>:|On[\s\S]+?wrote:)[\s\S]*/i;
  
  // חותך את מה שנמצא ומנקה רווחים
  return body.replace(regex, '').trim();
}
