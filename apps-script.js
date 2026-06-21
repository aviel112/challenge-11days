// ============================================================
// Google Apps Script — אתגר חיטוב 11 ימים
// ============================================================
// הוראות:
// 1. פתח script.google.com → New Project
// 2. הדבק את כל הקוד הזה
// 3. הרץ את setup() פעם אחת — זה יצור גיליון חדש אוטומטית
// 4. העתק את ה-SHEET_ID שמוצג ב-Logger ושים אותו בשורה הבאה:
// 5. Deploy → Web App → Execute as Me → Anyone → Copy URL
// 6. שים את ה-URL בתוך index.html במקום APPS_SCRIPT_URL_HERE
// ============================================================

let SHEET_ID = ''; // ימולא אוטומטי אחרי setup()

// מריץ פעם אחת — יוצר גיליון חדש
function setup() {
  const ss = SpreadsheetApp.create('אתגר חיטוב 11 ימים — נרשמים');
  const sheet = ss.getActiveSheet();
  sheet.setName('נרשמים');
  sheet.setRightToLeft(true);

  // כותרות
  const headers = ['תאריך', 'שם', 'טלפון', 'סטטוס תשלום', 'הערות'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // עיצוב כותרות
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a1a2e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);

  // רוחב עמודות
  sheet.setColumnWidth(1, 160); // תאריך
  sheet.setColumnWidth(2, 140); // שם
  sheet.setColumnWidth(3, 130); // טלפון
  sheet.setColumnWidth(4, 130); // סטטוס
  sheet.setColumnWidth(5, 200); // הערות

  // הקפאת שורת כותרות
  sheet.setFrozenRows(1);

  Logger.log('✅ גיליון נוצר! SHEET_ID: ' + ss.getId());
  Logger.log('🔗 קישור: ' + ss.getUrl());
}

// מקבל נרשם חדש מהאתר
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('נרשמים');

    sheet.appendRow([
      data.date || new Date().toLocaleString('he-IL'),
      data.name || '',
      data.phone || '',
      data.age || '',
      data.height || '',
      data.weight || '',
      data.bf ? data.bf + '%' : '',
      data.tdee || '',
      data.protein || '',
      'ממתין לתשלום 🟡',
      ''
    ]);

    // מייל התראה
    MailApp.sendEmail(
      'aviamira5@gmail.com',
      '🔥 נרשם חדש — אתגר 11 ימים',
      `שם: ${data.name}\nטלפון: ${data.phone}\nגיל: ${data.age} | גובה: ${data.height} | משקל: ${data.weight}\nאחוז שומן: ${data.bf}%\nTDEE: ${data.tdee} קל׳ | חלבון: ${data.protein}ג\nתאריך: ${data.date}\n\nבדוק אם שלח אישור תשלום בווטסאפ.`
    );

    return ContentService
      .createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error', msg:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Challenge 11 — Webhook Active ✅');
}
