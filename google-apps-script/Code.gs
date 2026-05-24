const SHEET_NAME = 'Request Tracker';
const APPROVED_SHEET_NAME = 'Approved';
const DISAPPROVED_SHEET_NAME = 'Disapproved';
const DESIGNERS_SHEET_NAME = 'Designers';
const SPREADSHEET_ID = '1vF7H7Yp7MrHOKe4j6HRkjjYrpxTtEh5ugEQ_OpKDaYU';
const UPLOAD_FOLDER_ID = '17Elym_RLgFLL2EPOOgS2FKhwNA3ikd-f';
const SECRET = 'change-this-secret';
const ADMIN_DASHBOARD_VERSION = '2026-05-24-admin-v3';

function doGet(e) {
  try {
    const params = e.parameter || {};
    if (SECRET && params.secret !== SECRET) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }

    if (params.action === 'version') {
      return jsonResponse({ ok: true, admin_dashboard: true, supports_delete: true, supports_designers: true, version: ADMIN_DASHBOARD_VERSION });
    }

    if (params.action === 'listRequests') {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = spreadsheet.getSheetByName(SHEET_NAME);
      if (!sheet) {
        return jsonResponse({ ok: false, error: `Missing sheet: ${SHEET_NAME}` });
      }
      return jsonResponse({ ok: true, admin_dashboard: true, supports_delete: true, supports_designers: true, version: ADMIN_DASHBOARD_VERSION, requests: listAllRequests_(spreadsheet) });
    }

    if (params.action === 'listDesigners') {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      return jsonResponse({ ok: true, designers: listDesigners_(spreadsheet) });
    }

    return jsonResponse({ ok: true, admin_dashboard: true, supports_delete: true, supports_designers: true, version: ADMIN_DASHBOARD_VERSION });
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
      updateRequest_(spreadsheet, payload.request_id, payload.updates || {});
      return jsonResponse({ ok: true });
    }

    if (payload.action === 'deleteRequest') {
      deleteRequest_(spreadsheet, payload.request_id);
      return jsonResponse({ ok: true });
    }

    if (payload.action === 'addDesigner') {
      addDesigner_(spreadsheet, payload.designer || {});
      return jsonResponse({ ok: true });
    }

    if (payload.action === 'deleteDesigner') {
      deleteDesigner_(spreadsheet, payload.email || '');
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
      source_tab: sheet.getName(),
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

function listAllRequests_(spreadsheet) {
  const tabs = [SHEET_NAME, APPROVED_SHEET_NAME, DISAPPROVED_SHEET_NAME];
  return tabs.flatMap((tabName) => {
    const sheet = spreadsheet.getSheetByName(tabName);
    return sheet ? listRequests_(sheet) : [];
  });
}

function listDesigners_(spreadsheet) {
  const sheet = ensureDesignersSheet_(spreadsheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues()
    .filter((row) => row[0] && row[2] !== 'No')
    .map((row) => ({
      name: row[0],
      email: row[1],
      active: row[2] || 'Yes',
      added_at: row[3],
    }));
}

function addDesigner_(spreadsheet, designer) {
  const sheet = ensureDesignersSheet_(spreadsheet);
  const name = String(designer.name || '').trim();
  const email = String(designer.email || '').trim();
  if (!name) throw new Error('Designer name is required');
  if (!email) throw new Error('Designer email is required');
  const existingRow = findDesignerRow_(sheet, email);
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, 4).setValues([[name, email, 'Yes', new Date()]]);
  } else {
    sheet.appendRow([name, email, 'Yes', new Date()]);
  }
}

function deleteDesigner_(spreadsheet, email) {
  const sheet = ensureDesignersSheet_(spreadsheet);
  const row = findDesignerRow_(sheet, email);
  if (!row) throw new Error(`Designer not found: ${email}`);
  sheet.getRange(row, 3).setValue('No');
}

function ensureDesignersSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(DESIGNERS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DESIGNERS_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([['Designer Name', 'Email', 'Active', 'Added At']]);
  }
  return sheet;
}

function findDesignerRow_(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const emails = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues().flat();
  const index = emails.findIndex((value) => String(value).toLowerCase() === String(email).toLowerCase());
  return index < 0 ? null : index + 2;
}

function updateRequest_(spreadsheet, requestId, updates) {
  const location = findRequestLocation_(spreadsheet, requestId);
  const sheet = location.sheet;
  const row = location.row;
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

  const approvalStatus = String(updates.approval_status || '');
  if (approvalStatus === 'Approved') {
    moveRequestToTab_(sheet, row, APPROVED_SHEET_NAME);
  } else if (approvalStatus === 'Not Approved' || approvalStatus === 'Needs Info') {
    moveRequestToTab_(sheet, row, DISAPPROVED_SHEET_NAME);
  }
}

function moveRequestToTab_(sourceSheet, sourceRow, targetSheetName) {
  const spreadsheet = sourceSheet.getParent();
  const targetSheet = ensureRequestTab_(spreadsheet, targetSheetName, sourceSheet);
  const rowValues = sourceSheet.getRange(sourceRow, 1, 1, 24).getValues()[0];
  const requestId = String(rowValues[0] || '');
  if (!requestId) return;

  removeExistingRequest_(targetSheet, requestId);
  targetSheet.appendRow(rowValues);
  sourceSheet.deleteRow(sourceRow);
}

function ensureRequestTab_(spreadsheet, tabName, sourceSheet) {
  let targetSheet = spreadsheet.getSheetByName(tabName);
  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet(tabName);
  }
  if (targetSheet.getLastRow() === 0) {
    const headers = sourceSheet.getRange(1, 1, 1, 24).getValues();
    targetSheet.getRange(1, 1, 1, 24).setValues(headers);
  }
  return targetSheet;
}

function removeExistingRequest_(sheet, requestId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
  for (let index = ids.length - 1; index >= 0; index -= 1) {
    if (String(ids[index]) === String(requestId)) {
      sheet.deleteRow(index + 2);
    }
  }
}

function deleteRequest_(spreadsheet, requestId) {
  const location = findRequestLocation_(spreadsheet, requestId);
  location.sheet.deleteRow(location.row);
}

function findRequestRow_(sheet, requestId) {
  if (!requestId) throw new Error('Missing request ID');
  const lastRow = sheet.getLastRow();
  const ids = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat() : [];
  const index = ids.findIndex((id) => String(id) === String(requestId));
  if (index < 0) throw new Error(`Request not found: ${requestId}`);
  return index + 2;
}

function findRequestLocation_(spreadsheet, requestId) {
  const tabs = [SHEET_NAME, APPROVED_SHEET_NAME, DISAPPROVED_SHEET_NAME];
  for (const tabName of tabs) {
    const sheet = spreadsheet.getSheetByName(tabName);
    if (!sheet) continue;
    try {
      return { sheet, row: findRequestRow_(sheet, requestId) };
    } catch (error) {
      // Keep searching the other request tabs.
    }
  }
  throw new Error(`Request not found: ${requestId}`);
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
