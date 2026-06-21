// Google Apps Script — העתק את כל הקוד הזה ל-script.google.com
// פרסם כ-Web App: Execute as Me, Anyone can access

const SHEET_ID = 'GOOGLE_SHEET_ID_HERE'; // החלף ב-ID של הגיליון

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();

    // הוסף שורה חדשה
    sheet.appendRow([
      data.date || new Date().toLocaleString('he-IL'),
      data.name || '',
      data.phone || '',
      '', // אימייל — ריק כרגע
      'לא טופל 🟡',
      'נרשם דרך אתגר 11 ימים',
      'לא'
    ]);

    // מייל התראה
    MailApp.sendEmail(
      'aviamira5@gmail.com',
      '🔥 נרשם חדש לאתגר 11 ימים',
      `שם: ${data.name}\nטלפון: ${data.phone}\nתאריך: ${data.date}`
    );

    return ContentService
      .createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error',msg:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Challenge 11 — Webhook Active');
}
