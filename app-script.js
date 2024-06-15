// Original code from https://github.com/jamiewilson/form-to-google-sheets
// Updated for 2021 and ES6 standards

const sheetName = 'Sheet1';
const scriptProp = PropertiesService.getScriptProperties();
const recaptchaSecret = 'your-recaptcha-secret-key'; // Replace it with your reCAPTCHA secret key

function initialSetup() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  scriptProp.setProperty('key', activeSpreadsheet.getId());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const recaptchaResponse = e.parameter['g-recaptcha-response'];
    
    // Verify Recaptcha
    const recaptchaVerification = UrlFetchApp.fetch('https://www.google.com/recaptcha/api/siteverify', {
      method: 'post',
      payload: {
        secret: recaptchaSecret,
        response: recaptchaResponse
      }
    });
    
    const recaptchaData = JSON.parse(recaptchaVerification.getContentText());

    if (recaptchaData.success) {
      // Recaptcha verification successful, proceed with form data processing
      const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
      const sheet = doc.getSheetByName(sheetName);

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const nextRow = sheet.getLastRow() + 1;

      const newRow = headers.map(function(header) {
        return header === 'Date' ? new Date() : e.parameter[header];
      });

      sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

      return ContentService
        .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      // Recaptcha verification failed
      return ContentService
        .createTextOutput(JSON.stringify({ 'result': 'error', 'error': 'Recaptcha verification failed' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
