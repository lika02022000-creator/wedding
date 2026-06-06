function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    // Map form data
    const fullname = data.fullname || '';
    const drinks = (data.drinks || []).join(', ');
    const dish = data.dish || '';
    const transport = data.transport || '';
    const stay = data.stay || '';
    const extra = data.extra || '';
    const timestamp = new Date();

    sheet.appendRow([timestamp, fullname, drinks, dish, transport, stay, extra]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}
