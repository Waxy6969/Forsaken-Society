const SHEET_NAME = 'Request Tracker';
const APPROVED_SHEET_NAME = 'Approved';
const DISAPPROVED_SHEET_NAME = 'Disapproved';
const DESIGNERS_SHEET_NAME = 'Designers';
const SPREADSHEET_ID = '1FRv0kfJc10hZLygUnApg5kbbFQ1Bnvxt2aE3-HQn5jQ';
const UPLOAD_FOLDER_ID = '17Elym_RLgFLL2EPOOgS2FKhwNA3ikd-f';
const SECRET = 'change-this-secret';
const ADMIN_DASHBOARD_VERSION = '2026-05-24-admin-v5-drive-upload';
const REQUEST_HEADERS = [
  'Request ID',
  'Date Submitted',
  'Member Name',
  'Company / Team',
  'Email',
  'Project Name',
  'Design Type',
  'Description',
  'Requested Deadline',
  'Priority',
  'Rush Option',
  'Estimated Rush Fee',
  'Status',
  'Seen Status',
  'Viewed Timestamp',
  'Approval Status',
  'Assigned Designer',
  'Estimated Completion',
  'Member Feedback Needed',
  'Google Drive Request Folder',
  'Uploaded Files Link',
  'Final Files Link',
  'Last Updated',
  'Admin Notes',
];

function doGet(e) {
  try {
    const params = e.parameter || {};
    if (params.view === 'simpleForm') {
      return simpleFormHtml_();
    }

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
    const params = e.parameter || {};
    if (params.action === 'submitSimpleForm') {
      return handleSimpleFormSubmit_(params);
    }

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

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const requestId = nextRequestId_(sheet);
      const values = payload.values || [];
      const uploadResult = saveFilesSafely_(payload.files || [], requestId);
      values[0] = requestId;
      if (uploadResult.links.length) values[20] = uploadResult.links.join('\n');
      while (values.length < 24) values.push('');
      if (uploadResult.error) {
        values[23] = [values[23], uploadResult.error].filter(Boolean).join('\n');
      }

      sheet.appendRow(values.slice(0, 24));
      return jsonResponse({ ok: true, request_id: requestId, upload_error: uploadResult.error });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function handleSimpleFormSubmit_(params) {
  const required = ['member_name', 'email', 'design_type', 'description'];
  const missing = required.filter((key) => !String(params[key] || '').trim());
  if (missing.length) {
    return simpleFormHtml_(`Missing required field: ${missing.join(', ')}`, true);
  }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(params.email || ''))) {
    return simpleFormHtml_('Enter a valid email address.', true);
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ensureRequestTab_(spreadsheet, SHEET_NAME, null);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const requestId = nextRequestId_(sheet);
    const submitted = new Date();
    const designType = String(params.design_type || '').trim();
    const projectName = `${designType || 'Design'} Request`;
    const deadline = 'Flexible';
    const rushOption = params.rush_requested ? 'Rush Order +$20' : 'No Rush';
    const rushFee = params.rush_requested ? '$20 rush fee' : 'No additional fee';
    const notes = [
      String(params.notes || '').trim(),
      designType === 'Other' && params.other_design_type ? `Other design type: ${params.other_design_type}` : '',
      customCartNotes_(params.custom_cart_items),
    ].filter(Boolean).join('\n');

    sheet.appendRow([
      requestId,
      submitted,
      String(params.member_name || '').trim(),
      'Forsaken',
      String(params.email || '').trim(),
      projectName,
      designType,
      String(params.description || '').trim(),
      deadline,
      'Standard',
      rushOption,
      rushFee,
      'Submitted',
      'Not Seen',
      '',
      'Pending Review',
      '',
      '',
      '',
      '',
      String(params.uploaded_files_link || '').trim(),
      '',
      submitted,
      notes,
    ]);

    return simpleFormHtml_(`I got your request. Request ID: ${requestId}`, false);
  } finally {
    lock.releaseLock();
  }
}

function customCartNotes_(rawItems) {
  if (!rawItems) return '';
  let items = [];
  try {
    items = JSON.parse(rawItems);
  } catch (error) {
    return '';
  }
  const lines = items
    .map((item) => {
      const name = String(item.name || '').trim();
      const price = String(item.price || '').trim();
      return name ? `- ${name}: ${price || '$0'}` : '';
    })
    .filter(Boolean);
  return lines.length ? `Custom cart items:\n${lines.join('\n')}` : '';
}

function simpleFormHtml_(status, isError) {
  const statusHtml = status
    ? `<div class="status ${isError ? 'error' : ''}">${escapeHtml_(status)}</div>`
    : '';
  return HtmlService
    .createHtmlOutput(`<!doctype html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Forsaken Simple Request Form</title>
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      font-family: Arial, Helvetica, sans-serif;
      color: #171717;
      background: #111;
    }
    main {
      width: min(720px, calc(100% - 28px));
      background: #fff;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 18px 60px rgba(0,0,0,.35);
    }
    header {
      padding: 22px 24px;
      background: #161616;
      color: #fff;
      border-bottom: 4px solid #e74719;
    }
    h1 { margin: 0; font-size: 1.55rem; text-transform: uppercase; }
    form {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
      padding: 24px;
    }
    label { display: grid; gap: 7px; font-weight: 800; }
    input, select, textarea {
      width: 100%;
      min-height: 44px;
      border: 1px solid #d2ccc2;
      border-radius: 6px;
      background: #f8f5ef;
      padding: 11px 12px;
      font: inherit;
    }
    .rush-check {
      display: flex;
      align-items: center;
      gap: 10px;
      min-height: 44px;
      border: 1px solid #d2ccc2;
      border-radius: 6px;
      background: #f8f5ef;
      padding: 10px 12px;
    }
    .rush-check input { width: 18px; min-height: 18px; }
    textarea { min-height: 132px; resize: vertical; }
    .full, .status, button { grid-column: 1 / -1; }
    .order-cart {
      grid-column: 1 / -1;
      border: 1px solid #d8d1c8;
      border-radius: 8px;
      background: #fffdf8;
      padding: 16px;
    }
    .order-cart h2 {
      margin: 0 0 12px;
      font-size: 1rem;
      text-transform: uppercase;
    }
    .cart-line, .cart-total {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 14px;
      margin-top: 8px;
    }
    .cart-line > span:first-child {
      min-width: 0;
      overflow-wrap: anywhere;
    }
    .cart-line span:last-child { color: #b93412; font-weight: 900; white-space: nowrap; }
    .cart-price { display: inline-flex; align-items: center; gap: 8px; white-space: nowrap; }
    .cart-total {
      border-top: 1px solid #e4e0da;
      padding-top: 12px;
      font-size: 1.1rem;
      font-weight: 900;
    }
    .cart-note {
      display: block;
      margin-top: 8px;
      color: #6d6d6d;
      font-size: .82rem;
    }
    .status {
      border: 1px solid #bbf7d0;
      border-radius: 6px;
      background: #f0fdf4;
      color: #166534;
      padding: 12px;
      font-weight: 800;
    }
    .status.error { border-color: #fed7aa; background: #fff7ed; color: #a24112; }
    button {
      min-height: 46px;
      border: 0;
      border-radius: 6px;
      background: #e74719;
      color: #fff;
      font-weight: 900;
      text-transform: uppercase;
      cursor: pointer;
    }
    @media (max-width: 620px) { form { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
  <main>
    <header><h1>Forsaken Simple Request Form</h1></header>
    <form method="post">
      ${statusHtml}
      <input type="hidden" name="action" value="submitSimpleForm">
      <label>Member Name
        <input name="member_name" autocomplete="name" required>
      </label>
      <label>Email
        <input name="email" type="email" autocomplete="email" required>
      </label>
      <label class="full">Design Type
        <select id="design_type" name="design_type" required>
          <option>AVIs</option>
          <option value="Twitter Headers">Twitter Headers - $25</option>
          <option value="YouTube Headers">YouTube Headers - $15</option>
          <option value="TikTok Banners">TikTok Banners - $15</option>
          <option value="Social Media Packages">Social Media Packages - $100</option>
          <option value="3D Animations">3D Animations - $95-$170</option>
          <option value="Intros or Outros">Intros or Outros - $150</option>
          <option value="Simple Editing">Simple Editing - $3/clip</option>
          <option value="Boss Editing">Boss Editing - $8/clip</option>
          <option value="Other">Other - Free</option>
        </select>
      </label>
      <label class="full">Describe What You Need
        <textarea name="description" required placeholder="Include style, colors, text, platform, and reference links."></textarea>
      </label>
      <label class="full">Rush Order
        <span class="rush-check"><input id="rush_requested" name="rush_requested" type="checkbox" value="Rush Order +$20"> Add rush fee (+$20)</span>
      </label>
      <label class="full">File or URL Link
        <input name="uploaded_files_link" type="url" placeholder="Optional Google Drive, Canva, Dropbox, or reference link">
      </label>
      <label class="full">Notes for Designer
        <textarea name="notes" placeholder="Optional"></textarea>
      </label>
      <section class="order-cart" aria-live="polite">
        <h2>Cart Summary</h2>
        <div id="cart_items"></div>
        <div class="cart-total">
          <span>Total Due</span>
          <span id="cart_total">$0</span>
        </div>
        <span id="cart_note" class="cart-note"></span>
      </section>
      <button type="submit">Submit Request</button>
    </form>
  </main>
  <script>
    const servicePrices = {
      "AVIs": { label: "AVI / Profile Picture", display: "Free", min: 0, max: 0 },
      "Twitter Headers": { label: "Twitter / X Header", display: "$25", min: 25, max: 25 },
      "YouTube Headers": { label: "YouTube Banner", display: "$15", min: 15, max: 15 },
      "TikTok Banners": { label: "TikTok Cover Graphic", display: "$15", min: 15, max: 15 },
      "Social Media Packages": { label: "Social Media Package", display: "$100", min: 100, max: 100 },
      "3D Animations": { label: "Custom 3D Intro / Outro", display: "$95-$170", min: 95, max: 170 },
      "Intros or Outros": { label: "Intro or Outro", display: "$150", min: 150, max: 150 },
      "Simple Editing": { label: "Simple Editing", display: "$3/clip", min: 3, max: 3, unit: "clip" },
      "Boss Editing": { label: "Boss Editing", display: "$8/clip", min: 8, max: 8, unit: "clip" },
      "Other": { label: "Other", display: "Free", min: 0, max: 0 },
    };
    const select = document.getElementById("design_type");
    const rushCheckbox = document.getElementById("rush_requested");
    const cartItems = document.getElementById("cart_items");
    const cartTotal = document.getElementById("cart_total");
    const cartNote = document.getElementById("cart_note");
    function money(value) {
      return "$" + Number(value || 0).toLocaleString();
    }
    function escapeText(value) {
      return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }
    function syncCart() {
      const service = servicePrices[select.value] || servicePrices.Other;
      const lines = [
        { name: service.label, price: service.display, min: service.min || 0, max: service.max || 0, quote: Boolean(service.quote) },
      ];
      if (rushCheckbox && rushCheckbox.checked) {
        lines.push({ name: 'Rush Order Fee', price: '$20', min: 20, max: 20 });
      }
      cartItems.innerHTML = lines.map((line) => (
        '<div class="cart-line"><span>' + escapeText(line.name) + '</span><span class="cart-price">' + line.price + '</span></div>'
      )).join('');
      const rushTotal = rushCheckbox && rushCheckbox.checked ? 20 : 0;
      if (service.quote) {
        cartTotal.textContent = "Custom Quote";
        cartNote.textContent = "Final pricing will be confirmed after review.";
      } else if (service.max && service.max !== service.min) {
        cartTotal.textContent = money(service.min + rushTotal) + "-" + money(service.max + rushTotal);
        cartNote.textContent = "Total is an estimate based on the published price range.";
      } else {
        cartTotal.textContent = money(service.min + rushTotal);
        cartNote.textContent = service.unit ? "Total shown is per " + service.unit + ". Final total depends on clip count." : "Total due based on selected service.";
      }
    }
    select.addEventListener("change", syncCart);
    if (rushCheckbox) rushCheckbox.addEventListener("change", syncCart);
    syncCart();
  </script>
</body>
</html>`)
    .setTitle('Forsaken Simple Request Form');
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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
    const headers = sourceSheet ? sourceSheet.getRange(1, 1, 1, 24).getValues() : [REQUEST_HEADERS];
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

function saveFilesSafely_(files, requestId) {
  if (!files.length) return { links: [], error: '' };
  try {
    return { links: saveFiles_(files, requestId), error: '' };
  } catch (error) {
    const message = String(error && error.message ? error.message : error);
    const names = files.map((file) => file.name).filter(Boolean).join(', ');
    const detail = names ? ` Attempted files: ${names}.` : '';
    return {
      links: [],
      error: `Upload files were not saved because Google Drive upload permission needs to be authorized in Apps Script.${detail} Apps Script error: ${message}`,
    };
  }
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
