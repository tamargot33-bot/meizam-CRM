# תוכן העניינים של פרויקט Mizam CRM (Table of Contents)

מערכת CRM אינטגרטיבית המבוססת על Google Apps Script, המחברת בין Gmail, קבוצת טלגרם מבוססת בוט/נושאים, Google Sheets וממשק ווב עשיר.

---

## מפת הקבצים הכללית

| קובץ | תפקיד עיקרי | מועלה ל-GitHub? | נדחף ל-Apps Script? |
| :--- | :--- | :---: | :---: |
| **`Config.js`** | מפתחות, טוקנים, מזהי Sheets והגדרות סודיות | ❌ לא (ב-`.gitignore`) | ✅ כן |
| **`Config.example.js`** | תבנית הגדרות ציבורית עם ערכים ריקים | ✅ כן | ❌ לא (ב-`.claspignore`) |
| **`Database.js`** | גישה לגוגל שיטס, הקמת טבלאות ועדכוני רשומות | ✅ כן | ✅ כן |
| **`Incoming.js`** | סורק מיילים נכנסים והפצת התראות Push | ✅ כן | ✅ כן |
| **`Outgoing.js`** | מענה ומשלוח מיילים חדשים מג'ימייל | ✅ כן | ✅ כן |
| **`WebApp.js`** | שרת ה-Web App, טעינת ה-CRM ומחיקות/שחזורים | ✅ כן | ✅ כן |
| **`Notes.js`** | ניהול פתקים, הערות ומשימות פנימיות | ✅ כן | ✅ כן |
| **`TelegramOutgoing.js`** | שליחת הודעות וכפתורים אינטראקטיביים לטלגרם | ✅ כן | ✅ כן |
| **`TelegramIncoming.js`** | Webhook של טלגרם, טיפול בלחיצות כפתור ומענה | ✅ כן | ✅ כן |
| **`Triggers.js`** | חיבור Webhook, בדיקות וסנכרון היסטוריה | ✅ כן | ✅ כן |
| **`index.html`** | ממשק המשתמש (צד לקוח) של ה-CRM | ✅ כן | ✅ כן |
| **`appsscript.json`** | מניפסט והרשאות Google Apps Script | ✅ כן | ✅ כן |
| **`backup/`** | תיקיית גיבוי לקוד המקורי השלם | ❌ לא | ❌ לא |

---

## פירוט פונקציות לפי קבצים

### 1. `Database.js` — תשתית גיליונות ונתונים
| שם פונקציה | פרמטרים | תיאור |
| :--- | :--- | :--- |
| `getSheet(name)` | `name` | פותחת ומחזירה גיליון ספציפי מתוך ה-Spreadsheet המוגדר ב-CONFIG. |
| `getManagerByEmail(email)` | `email` | מאתרת מנהל לפי כתובת המייל שלו מתוך גיליון "Managers" ומחזירה אובייקט עם פרטיו. |
| `isManagerEmail(email)` | `email` | בודקת האם כתובת מייל שייכת לאחד המנהלים או למייל הארגון. |
| `getExistingMsgIds()` | — | מחזירה רשימת `msgId` הקיימים בגיליון "Messages" למניעת כפילויות. |
| `setupSheets()` | — | בונה ומעצבת את כל הגיליונות הנדרשים למערכת אם אינם קיימים, ומזינה נתוני פתיחה. |
| `addMissingSheets()` | — | מוסיפה גיליונות חדשים (כגון "Notes") מבלי לגעת בגיליונות קיימים. |
| `upsertContact(email, name)` | `email`, `name` | מוסיפה איש קשר חדש לגיליון "Contacts" במידה ואינו קיים. |
| `upsertThread(threadId, subject, contactEmail)` | `threadId`, `subject`, `contactEmail` | רושמת שרשור חדש ב-"Threads" או מעדכנת את תאריך הפעילות האחרונה שלו. |
| `updateMessageStatus(msgId, status, handledBy)` | `msgId`, `status`, `handledBy` | מעדכנת סטטוס הודעה ושם המטפל בגיליון "Messages". |
| `updateThreadLabel(threadId, label, handledBy)` | `threadId`, `label`, `handledBy` | מעדכנת את התווית של השרשור ותאריך הפעילות שלו בגיליון "Threads". |
| `updateThreadDisplaySubject(threadId, newSubject)` | `threadId`, `newSubject` | מעדכנת נושא תצוגה מותאם אישית לשרשור. |
| `updateThreadNotes(threadId, notes)` | `threadId`, `notes` | שומרת הערות פנימיות כלליות עבור השרשור. |

---

### 2. `Incoming.js` — סריקת מיילים והתראות
| שם פונקציה | פרמטרים | תיאור |
| :--- | :--- | :--- |
| `checkAndSendEmails()` | — | **פונקציית הטריגר הראשית**: סורקת מיילים חדשים שלא נקראו בג'ימייל, מסננת מיילים פנימיים, מתעדת בשיטס, שולחת לטלגרם ומפעילה התראת פוש. |
| `sendPushNotification(senderName, subject)` | `senderName`, `subject` | שולחת התראת דחיפה (Push) דרך OneSignal לכל מחשבי המנהלים המחוברים. |
| `testMyPushToSpecificComputer()` | — | פונקציית בדיקה לשליחת התראת פוש למכשיר ספציפי לפי `TEST_SUBSCRIPTION_ID`. |

---

### 3. `Outgoing.js` — שליחת מיילים יוצאים
| שם פונקציה | פרמטרים | תיאור |
| :--- | :--- | :--- |
| `sendCrmReplyFromWeb(threadId, replyText, managerName)` | `threadId`, `replyText`, `managerName` | שולחת מענה ישיר מהממשק להודעת הלקוח האחרונה בג'ימייל, מעדכנת תוויות ומתעדת בשיטס. |
| `sendNewEmailFromWeb(toEmail, subject, body, managerName)` | `toEmail`, `subject`, `body`, `managerName` | פותחת שרשור מייל חדש ללקוח ישירות מהממשק, מוסיפה לאנשי קשר ומתעדת בשיטס. |
| `extractEmail(raw)` | `raw` | מחלצת כתובת אימייל נקייה מתוך שורת שולח/נמען (מסירה שמות וסוגריים משולשים). |
| `cleanEmailBody(body)` | `body` | מסירה היסטוריית תכתובות וציטוטים ישנים מגוף המייל כדי להציג רק את התגובה החדשה. |

---

### 4. `WebApp.js` — ממשק הווב וה-CRM
| שם פונקציה | פרמטרים | תיאור |
| :--- | :--- | :--- |
| `doGet(e)` | `e` | נקודת הכניסה לדפדפן: מחזירה נתוני JSON (אם `?fetch=data`) או מגישה את `index.html` עם משתני קונפיגורציה מוזרקים. |
| `loadCrmData(currentManagerName)` | `currentManagerName` | שולפת ומסננת את כל המידע ל-CRM (הודעות, שרשורים, מנהלים, תוויות, הערות). רק מנהל-על רואה פריטים מחוקים. |
| `deleteThreadFromWeb(threadId, managerName, deleteFromGmail)` | `threadId`, `managerName`, `deleteFromGmail` | מחיקה רכה של שרשור בשיטס, ואפשרות למחיקה מוחלטת לאשפת ג'ימייל (למנהל ראשי). |
| `deleteMessageFromWeb(msgId, managerName, deleteFromGmail)` | `msgId`, `managerName`, `deleteFromGmail` | מחיקה רכה של הודעה בודדת בשיטס, עם אפשרות העברה לאשפה בג'ימייל. |
| `undeleteThread(threadId)` | `threadId` | שחזור וביטול מחיקה של שרשור. |
| `undeleteMessage(msgId)` | `msgId` | שחזור וביטול מחיקה של הודעה. |

---

### 5. `Notes.js` — ניהול הערות ומשימות
| שם פונקציה | פרמטרים | תיאור |
| :--- | :--- | :--- |
| `addNote(threadId, body, type, targetManager, createdBy)` | `threadId`, `body`, `type`, `targetManager`, `createdBy` | יוצרת פתק או משימה המשויכת לשרשור ושומרת בגיליון "Notes". |
| `completeTask(noteId)` | `noteId` | מסמנת משימה כהושלמה ורושמת חותמת זמן של הביצוע. |
| `editNote(noteId, newBody, requestedBy)` | `noteId`, `newBody`, `requestedBy` | עורכת תוכן של פתק (מורשה אך ורק למי שיצר אותו). |
| `deleteNote(noteId, requestedBy)` | `noteId`, `requestedBy` | מוחקת פתק מהמערכת (מורשה אך ורק למי שיצר אותו). |

---

### 6. `TelegramOutgoing.js` — שליחה לטלגרם וכפתורים
| שם פונקציה | פרמטרים | תיאור |
| :--- | :--- | :--- |
| `sendEmailToTelegram(name, email, subject, body, msgId, threadId)` | `name`, `email`, `subject`, `body`, `msgId`, `threadId` | שולחת הודעה לקבוצת הטלגרם עם תמצית המייל וכפתורי פעולה אינטראקטיביים. |
| `escapeMd(text)` | `text` | מנטרלת תווים שמורים של Telegram MarkdownV2 כדי למנוע קריסת הודעות. |
| `buildLabelKeyboard(labels, prefix)` | `labels`, `prefix` | יוצרת מקלדת Inline בטלגרם לבחירת תווית. |
| `buildDefaultKeyboard(threadId, tgMsgId)` | `threadId`, `tgMsgId` | בונה את מקלדת ברירת המחדל להודעה (מענה, פתיחה בווב, שינוי תווית). |
| `telegramPost(method, payload)` | `method`, `payload` | עטיפה לקריאות HTTP POST ישירות מול Telegram Bot API. |
| `editMarkup(chatId, msgId, markup)` | `chatId`, `msgId`, `markup` | מעדכנת את הכפתורים של הודעה קיימת בטלגרם. |
| `editText(chatId, msgId, text, markup)` | `chatId`, `msgId`, `text`, `markup` | מעדכנת טקסט וכפתורים של הודעה קיימת בטלגרם. |
| `answerCb(id, text)` | `id`, `text` | מחזירה מענה ל-Callback Query של טלגרם (סגירת אינדיקציית טעינה / הודעת Toast קופצת). |

---

### 7. `TelegramIncoming.js` — קבלת אירועים ומענה מטלגרם
| שם פונקציה | פרמטרים | תיאור |
| :--- | :--- | :--- |
| `doPost(e)` | `e` | **נקודת הכניסה של ה-Webhook של טלגרם**: מפענחת את ה-Update ומנתבת לכפתור או להודעה. |
| `handleCallbackQuery(cb)` | `cb` | מטפלת בלחיצה על כפתור Inline: מענה למייל, החלפת תווית, או ביטול. |
| `handleTelegramMessage(message)` | `message` | מטפלת בהודעת טקסט שנשלחה בתגובה: שליחת המייל ללקוח בג'ימייל, תיעוד בשיטס ועדכון הודעת הטלגרם. |

---

### 8. `Triggers.js` — הגדרות טלגרם, בדיקות וסנכרון
| שם פונקציה | פרמטרים | תיאור |
| :--- | :--- | :--- |
| `setTelegramWebhook()` | — | מחברת את כתובת ה-Web App של הפרויקט כ-Webhook של בוט הטלגרם. |
| `checkWebhook()` | — | בודקת את סטטוס ה-Webhook הנוכחי בטלגרם ומציגה מידע ושגיאות בלוג. |
| `backfillOldThreads()` | — | מייבאת היסטוריית מיילים ישנים מג'ימייל לתוך ה-CRM ומסמנת אותם כארכיון. |
| `testCreateTelegramTopic()` | — | בדיקה ליצירת נושא חדש (Forum Topic) בקבוצת טלגרם תומכת נושאים ושליחת הודעה לתוכו. |
| `backfillThreadsToTelegramTopics()` | — | סנכרון מלא של כל השרשורים, ההודעות וההערות לתוך נושאים נפרדים בטלגרם. |
