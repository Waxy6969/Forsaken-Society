const SHEET_NAME = 'Request Tracker';
const SPREADSHEET_ID = '1vF7H7Yp7MrHOKe4j6HRkjjYrpxTtEh5ugEQ_OpKDaYU';
const UPLOAD_FOLDER_ID = '17Elym_RLgFLL2EPOOgS2FKhwNA3ikd-f';
const SECRET = 'change-this-secret';
const ADMIN_DASHBOARD_VERSION = '2026-05-23-admin-v2';

function doGet(e) {
  try {
    const params = e.parameter || {};
    if (SECRET && params.secret !== SECRET) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }

    if (params.action === 'version') {
      return jsonResponse({ ok: true, admin_dashboard: true, supports_delete: true, version: ADMIN_DASHBOARD_VERSION });
    }

    if (params.action === 'listRequests') {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = spreadsheet.getSheetByName(SHEET_NAME);
      if (!sheet) {
        return jsonResponse({ ok: false, error: `Missing sheet: ${SHEET_NAME}` });
      }
      return jsonResponse({ ok: true, admin_dashboard: true, supports_delete: true, version: ADMIN_DASHBOARD_VERSION, requests: listRequests_(sheet) });
    }

    return jsonResponse({ ok: true, admin_dashboard: true, supports_delete: true, version: ADMIN_DASHBOARD_VERSION });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

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

    if (payload.action === 'listRequests') {
      return jsonResponse({ ok: true, requests: listRequests_(sheet) });
    }

    if (payload.action === 'updateRequest') {
      updateRequest_(sheet, payload.request_id, payload.updates || {});
      return jsonResponse({ ok: true });
    }

    if (payload.action === 'deleteRequest') {
      deleteRequest_(sheet, payload.request_id);
      return jsonResponse({ ok: true });
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

function listRequests_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, 24).getDisplayValues();
  return rows
    .filter((row) => row[0])
    .map((row) => ({
      request_id: row[0],
      submitted_at: row[1],
      member_name: row[2],
      team: row[3],
      email: row[4],
      project_name: row[5],
      design_type: row[6],
      description: row[7],
      requested_deadline: row[8],
      priority: row[9],
      rush_option: row[10],
      rush_fee: row[11],
      request_status: row[12],
      seen_status: row[13],
      seen_by: row[14],
      approval_status: row[15],
      assigned_designer: row[16],
      design_start_date: row[17],
      design_due_date: row[18],
      drive_folder_link: row[19],
      uploaded_files_link: row[20],
      final_deliverables_link: row[21],
      last_updated: row[22],
      admin_notes: row[23],
    }))
    .reverse();
}

function updateRequest_(sheet, requestId, updates) {
  const row = findRequestRow_(sheet, requestId);
  const columns = {
    request_status: 13,
    seen_status: 14,
    seen_by: 15,
    approval_status: 16,
    assigned_designer: 17,
    design_start_date: 18,
    design_due_date: 19,
    final_deliverables_link: 22,
    admin_notes: 24,
  };
  Object.keys(columns).forEach((key) => {
    if (updates[key] !== undefined) {
      sheet.getRange(row, columns[key]).setValue(updates[key]);
    }
  });
  sheet.getRange(row, 23).setValue(new Date());
}

function deleteRequest_(sheet, requestId) {
  const row = findRequestRow_(sheet, requestId);
  sheet.deleteRow(row);
}

function findRequestRow_(sheet, requestId) {
  if (!requestId) throw new Error('Missing request ID');
  const lastRow = sheet.getLastRow();
  const ids = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat() : [];
  const index = ids.findIndex((id) => String(id) === String(requestId));
  if (index < 0) throw new Error(`Request not found: ${requestId}`);
  return index + 2;
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
