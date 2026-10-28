/**
 * Receipts — snap a photo, AI files it in Google Drive
 * Google Apps Script + Claude API
 *
 * ┌──────────────────────────────────────────────────────────┐
 * │  TO CUSTOMIZE: edit only the MY_CONFIG block below.       │
 * │  Everything else works on its own.                        │
 * └──────────────────────────────────────────────────────────┘
 */

const MY_CONFIG = {
  COMPANY:   'My Company Inc.',        // your business name (the AI uses it as context)
  LOCATION:  'Ontario, Canada',        // province/country (helps recognize merchants and dates)
  FOLDER:    'Receipts',               // main folder name in your Google Drive
  APP_TITLE: 'Receipts'                // title shown in the phone app
};

const CONFIG = {
  ROOT_NAME: MY_CONFIG.FOLDER,
  MODEL: 'claude-sonnet-5-5',       // switch to 'claude-haiku-5-5' to lower cost
  MAX_RUN_MS: 5 * 60 * 1000,        // Apps Script stops at 6 min; leave a margin
  MIN_CONFIDENCE: 0.6,              // below this, the receipt goes to _Review
  MAX_IMAGE_BYTES: 4.5 * 1024 * 1024
};

// Business expense categories (CRA-aligned). The key is what appears in the file name.
const CATEGORIES = {
  meals:        'Meals & entertainment (restaurants, cafés, coffee shops, takeout)',
  fuel:         'Fuel / gas stations',
  vehicle:      'Vehicle maintenance, repairs, tires, car wash',
  parking:      'Parking and tolls (incl. 407 ETR)',
  office:       'Office supplies and stationery (Staples, etc.)',
  equipment:    'Equipment and hardware (computers, electronics, furniture)',
  software:     'Software and subscriptions',
  phone:        'Phone and internet',
  travel:       'Travel: flights, hotels, transit, taxis, Uber',
  training:     'Training, courses, books, certifications',
  professional: 'Professional fees (accountant, lawyer, consultants)',
  advertising:  'Advertising and marketing',
  bank:         'Bank and payment fees',
  other:        'Other business expense that fits none of the above'
};

const MONTHS = ['January','February','March','April','May','June','July',
                'August','September','October','November','December'];

/* ---------- SETUP (run once) ---------- */

function setup() {
  const props = PropertiesService.getScriptProperties();
  // Optional: PARENT_FOLDER_ID = ID of a folder or Shared Drive.
  // If not set, uses "My Drive" of the account running the script.
  const parentId = props.getProperty('PARENT_FOLDER_ID');
  const parent = parentId ? DriveApp.getFolderById(parentId) : DriveApp.getRootFolder();
  const root   = getOrCreateFolder_(parent, CONFIG.ROOT_NAME);
  const inbox  = getOrCreateFolder_(root, '_Inbox');
  const review = getOrCreateFolder_(root, '_Review');

  let sheetId = props.getProperty('SHEET_ID');
  if (!sheetId) {
    const ss = SpreadsheetApp.create(CONFIG.ROOT_NAME + ' - Log');
    ss.getSheets()[0].appendRow(['Processed', 'Receipt date', 'Merchant', 'Category',
                                 'Total', 'Currency', 'Confidence', 'Status', 'File', 'Link']);
    ensureInstructions_(ss.getSheets()[0]);
    DriveApp.getFileById(ss.getId()).moveTo(root);
    sheetId = ss.getId();
  }

  if (!props.getProperty('UPLOAD_KEY')) {
    props.setProperty('UPLOAD_KEY', Utilities.getUuid().replace(/-/g, ''));
  }
  props.setProperties({
    ROOT_ID: root.getId(), INBOX_ID: inbox.getId(),
    REVIEW_ID: review.getId(), SHEET_ID: sheetId
  });

  // Hourly trigger (if the inbox is empty, it exits immediately)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processInbox')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processInbox').timeBased().everyHours(1).create();

  Logger.log('✅ Setup done. Folder created in Drive: ' + root.getUrl());
  Logger.log('Next: Deploy → New deployment (Web app). Then run "myLink".');
}

// Run AFTER deploying: prints the link for your phone.
// Note: from the editor, Apps Script may return the /dev (test) URL, which only
// works for the owner. The shareable URL always ends in /exec.
function myLink() {
  const key = PropertiesService.getScriptProperties().getProperty('UPLOAD_KEY');
  if (!key) return Logger.log('❌ Run "setup" first.');
  const url = ScriptApp.getService().getUrl() || '';
  if (/\/exec$/.test(url)) {
    Logger.log('📱 Your phone link (keep it private):\n' + url + '?k=' + key);
  } else {
    Logger.log('📱 Build your phone link (keep it private):\n' +
      '1) Deploy → Manage deployments → copy the Web app URL (it ends in /exec)\n' +
      '2) Add this to the end:  ?k=' + key);
  }
}

/* ---------- WEB APP ---------- */

function doGet(e) {
  const key = PropertiesService.getScriptProperties().getProperty('UPLOAD_KEY');
  if (!e || !e.parameter || e.parameter.k !== key) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:-apple-system,system-ui,sans-serif;text-align:center;padding:80px 24px;color:#86868b">' +
      '<div style="font-size:44px">🔒</div><p>This link is not valid.</p></div>')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const t = HtmlService.createTemplateFromFile('Index');
  t.k = key;
  t.title = MY_CONFIG.APP_TITLE;
  return t.evaluate()
    .setTitle(MY_CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function uploadReceipt(k, base64, mimeType) {
  const props = PropertiesService.getScriptProperties();
  if (k !== props.getProperty('UPLOAD_KEY')) throw new Error('Unauthorized');
  if (!props.getProperty('INBOX_ID')) throw new Error('Setup not finished: run "setup" in the editor');

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  const name = stamp + '_' + Math.random().toString(36).slice(2, 6) + '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || 'image/jpeg', name);
  DriveApp.getFolderById(props.getProperty('INBOX_ID')).createFile(blob);
  return { ok: true, name: name };
}

/* ---------- BATCH PROCESSING ---------- */

function processInbox() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // another run is in progress

  try {
    const start = Date.now();
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('ANTHROPIC_API_KEY');
    if (!apiKey) throw new Error('Missing ANTHROPIC_API_KEY in Script Properties');

    const root   = DriveApp.getFolderById(props.getProperty('ROOT_ID'));
    const review = DriveApp.getFolderById(props.getProperty('REVIEW_ID'));
    const inbox  = DriveApp.getFolderById(props.getProperty('INBOX_ID'));
    const sheet  = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheets()[0];

    // 1) Corrections typed in the log (Status = FIX). No AI call needed.
    ensureInstructions_(sheet);
    applyFixes_(sheet, root, review);

    // 2) New photos in the inbox.
    const files = inbox.getFiles();
    while (files.hasNext() && Date.now() - start < CONFIG.MAX_RUN_MS) {
      const file = files.next();
      try {
        const r = extract_(file, apiKey);
        const ok = isValidDate_(r.date) && CATEGORIES[r.category] && Number(r.confidence) >= CONFIG.MIN_CONFIDENCE;
        const newName = fileReceipt_(file, r, ok, root, review);
        const status = ok ? 'OK' + (r.note ? ' (' + r.note + ')' : '')
                          : 'REVIEW' + (r.note ? ' (' + r.note + ')' : '');
        sheet.appendRow([new Date(), r.date || '', r.merchant || '', r.category || '',
                         r.total == null ? '' : r.total, r.currency || '',
                         r.confidence == null ? '' : r.confidence,
                         status, newName, file.getUrl()]);
      } catch (err) {
        // Stays in the inbox and is retried on the next run
        console.error(file.getName() + ': ' + err);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// Renames and moves a receipt. ok=true → Year/MM-Month, otherwise → _Review.
function fileReceipt_(file, r, ok, root, review) {
  const merchant = slug_(r.merchant || 'unknown');
  let dest, base;
  if (ok) {
    const [y, m] = r.date.split('-');
    dest = getOrCreateFolder_(getOrCreateFolder_(root, y), m + '-' + MONTHS[Number(m) - 1]);
    base = r.date + '_' + r.category + '_' + merchant;
  } else {
    dest = review;
    base = (isValidDate_(r.date) ? r.date : 'no-date') + '_' +
           (CATEGORIES[r.category] ? r.category : 'no-category') + '_' + merchant;
  }
  const newName = uniqueName_(dest, base, ext_(file));
  file.setName(newName);
  file.moveTo(dest);
  return newName;
}

/*
 * FIXING A RECEIPT FROM THE LOG
 * In the "Log" sheet, on the receipt's row: correct Receipt date (YYYY-MM-DD), Merchant or
 * Category, then type FIX in Status. The next run files it with your values (no AI call).
 * Run "processInbox" from the editor to apply right away instead of waiting for the hour.
 */
function applyFixes_(sheet, root, review) {
  const first = headerRow_(sheet) + 1;
  const last = sheet.getLastRow();
  if (last < first) return;
  const rows = sheet.getRange(first, 1, last - first + 1, 10).getValues();
  const tz = Session.getScriptTimeZone();

  rows.forEach((row, i) => {
    const status = String(row[7]).trim().toUpperCase();
    if (status !== 'FIX') return;
    const rowNum = i + first;
    const id = String(row[9]).match(/[-\w]{25,}/);
    if (!id) { sheet.getRange(rowNum, 8).setValue('REVIEW (file link missing)'); return; }
    let file;
    try { file = DriveApp.getFileById(id[0]); }
    catch (e) { sheet.getRange(rowNum, 8).setValue('REVIEW (file not found)'); return; }

    // Sheets turns typed dates into Date objects; accept both that and plain text.
    const isDate = Object.prototype.toString.call(row[1]) === '[object Date]' && !isNaN(row[1]);
    const date = isDate ? Utilities.formatDate(row[1], tz, 'yyyy-MM-dd') : String(row[1]).trim();
    const r = { date: date, merchant: String(row[2]).trim(), category: String(row[3]).trim().toLowerCase() };
    const problems = [];
    if (!isValidDate_(r.date)) problems.push('date must be YYYY-MM-DD');
    if (!CATEGORIES[r.category]) problems.push('category must be one of: ' + Object.keys(CATEGORIES).join(', '));
    if (problems.length) {
      sheet.getRange(rowNum, 8).setValue('REVIEW (' + problems.join('; ') + ')');
      return;
    }
    const newName = fileReceipt_(file, r, true, root, review);
    sheet.getRange(rowNum, 2).setNumberFormat('@').setValue(r.date);
    sheet.getRange(rowNum, 4).setValue(r.category);
    sheet.getRange(rowNum, 8).setValue('OK (fixed)');
    sheet.getRange(rowNum, 9).setValue(newName);
  });
}

const INSTRUCTIONS = '✏️ To fix a receipt: if anything is wrong (or Status says REVIEW), correct its ' +
  'Receipt date (YYYY-MM-DD), Merchant or Category in that row, then type FIX in Status. ' +
  'It is re-filed with your values within the hour.';

// Adds the instruction banner above the header once (idempotent).
function ensureInstructions_(sheet) {
  if (String(sheet.getRange(1, 1).getValue()).indexOf('✏️ To fix') === 0) return;
  sheet.insertRowBefore(1);
  const banner = sheet.getRange(1, 1, 1, 10);
  banner.merge();
  banner.setValue(INSTRUCTIONS).setWrap(true).setVerticalAlignment('middle')
        .setBackground('#fff4cc').setFontWeight('bold').setFontColor('#5c4200');
  sheet.setRowHeight(1, 52);
  sheet.setFrozenRows(2);
}

// Row number of the "Processed / Receipt date / …" header (1 or 2).
function headerRow_(sheet) {
  return String(sheet.getRange(1, 8).getValue()) === 'Status' ? 1 : 2;
}

function isValidDate_(s) {
  return typeof s === 'string' && /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(s);
}

function extract_(file, apiKey) {
  const blob = file.getBlob();
  const mime = blob.getContentType();
  if (!/^image\/(jpeg|png|webp|gif)$/.test(mime)) {
    return { confidence: 0, note: 'unsupported format' };
  }
  const bytes = blob.getBytes();
  if (bytes.length > CONFIG.MAX_IMAGE_BYTES) {
    return { confidence: 0, note: 'image too large' };
  }

  const catList = Object.keys(CATEGORIES).map(k => '- ' + k + ': ' + CATEGORIES[k]).join('\n');
  const prompt =
    'You are a bookkeeping assistant for ' + MY_CONFIG.COMPANY + ', a small business in ' + MY_CONFIG.LOCATION + '.\n' +
    'Read this receipt/invoice and return ONLY a JSON object, no other text:\n' +
    '{"date":"YYYY-MM-DD","merchant":"short merchant name","category":"<key>","total":number,' +
    '"currency":"3-letter ISO code","confidence":0.0-1.0,"note":"short note or null"}\n\n' +
    'Rules:\n' +
    '- The photo may show SEVERAL slips of the SAME purchase (e.g. the itemized bill plus the card ' +
    'payment slip where a tip was added). Treat them as ONE receipt. In that case total = the final amount ' +
    'actually charged (the card slip including tip, normally the higher amount) and note = "bill + card slip with tip".\n' +
    '- If the slips are clearly DIFFERENT purchases (different merchants or dates), set confidence < 0.6 ' +
    'and note = "multiple receipts in one photo".\n' +
    '- date = the transaction date printed on the receipt (NOT today). Use null if unreadable and note = "date unreadable".\n' +
    '- Choose the category from the MERCHANT type, using exactly one of these keys:\n' + catList + '\n' +
    '- total = final amount paid including tax and tip. null if unreadable.\n' +
    '- confidence = how sure you are about date AND category together. ' +
    'Use < 0.6 if the image is blurry, not a receipt, or the merchant type is ambiguous.';

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CONFIG.MODEL,
      max_tokens: 400,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mime, data: Utilities.base64Encode(bytes) } },
          { type: 'text', text: prompt }
        ]
      }]
    })
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('API ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  }
  const text = JSON.parse(res.getContentText()).content.map(c => c.text || '').join('');
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) return { confidence: 0, note: 'unreadable response' };
  return JSON.parse(match[0]);
}

/* ---------- HELPERS ---------- */

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function slug_(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 30) || 'unknown';
}

function ext_(file) {
  const mime = file.getMimeType();
  if (mime === 'image/png') return '.png';
  if (mime === 'image/webp') return '.webp';
  return '.jpg';
}

function uniqueName_(folder, base, ext) {
  let name = base + ext, i = 2;
  while (folder.getFilesByName(name).hasNext()) name = base + '_' + (i++) + ext;
  return name;
}
