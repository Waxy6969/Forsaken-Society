const SHEET_NAME = 'Request Tracker';
const SPREADSHEET_ID = '1vF7H7Yp7MrHOKe4j6HRkjjYrpxTtEh5ugEQ_OpKDaYU';
const UPLOAD_FOLDER_ID = '17Elym_RLgFLL2EPOOgS2FKhwNA3ikd-f';
const SECRET = 'change-this-secret';

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (SECRET && payload.secret !== SECRET) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ ok: false, error: `Missing sheet: ${SHEET_NAME}` });
    }

    const requestId = nextRequestId_(sheet);
    const values = payload.values || [];
    const fileLinks = saveFiles_(payload.files || [], requestId);
    values[0] = requestId;
    if (fileLinks.length) values[20] = fileLinks.join('\n');
    while (values.length < 24) values.push('');

    sheet.appendRow(values.slice(0, 24));
    return jsonResponse({ ok: true, request_id: requestId });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function saveFiles_(files, requestId) {
  if (!files.length) return [];
  const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  return files.map((file, index) => {
    const bytes = Utilities.base64Decode(file.data || '');
    const safeName = String(file.name || `upload-${index + 1}`).replace(/[\\/:*?"<>|]/g, '-');
    const blob = Utilities.newBlob(bytes, file.type || 'application/octet-stream', `${requestId}-${safeName}`);
    const created = folder.createFile(blob);
    created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return created.getUrl();
  });
}

function nextRequestId_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const ids = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  let highest = 0;
  ids.forEach((value) => {
    const match = String(value || '').match(/^CRP-(\d+)$/);
    if (match) highest = Math.max(highest, Number(match[1]));
  });
  return `CRP-${String(highest + 1).padStart(4, '0')}`;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
