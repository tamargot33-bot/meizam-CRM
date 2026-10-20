// =================================================================
// Notes.js — ניהול פתקים, הערות ומשימות ב-CRM
// =================================================================

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
      sheet.getRange(i + 1, 4).setValue("note");     // הופך להערה רגילה
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
