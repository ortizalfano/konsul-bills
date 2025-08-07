/**
 * @OnlyCurrentDoc
 * @AuthMode(ScriptApp.AuthMode.FULL)
 * @Scope https://www.googleapis.com/auth/spreadsheets
 * @Scope https://www.googleapis.com/auth/gmail.send
 * @Scope https://www.googleapis.com/auth/gmail.modify
 * @Scope https://www.googleapis.com/auth/script.external_request
 * @Scope https://www.googleapis.com/auth/userinfo.email
 */

// Constantes de nombres de hoja
const QUOTES_SHEET_NAME   = 'Quotes';
const INVOICES_SHEET_NAME = 'Invoices';
const BILLING_SHEET_NAME  = 'Billing';

// =========================
// SETUP INICIAL
// =========================
function initialSetup() {
  const props = PropertiesService.getUserProperties();
  let ssId = props.getProperty('spreadsheetId');
  if (ssId) {
    try { SpreadsheetApp.openById(ssId); return; }
    catch(e) { ssId = null; }
  }
  // Crear nueva hoja
  const ss = SpreadsheetApp.create('Kônsul - Datos de Facturación');
  props.setProperty('spreadsheetId', ss.getId());
  ensureSheets_();
}

function ensureSheets_() {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return;
  const ss = SpreadsheetApp.openById(ssId);
  if (!ss.getSheetByName(QUOTES_SHEET_NAME)) {
    ss.insertSheet(QUOTES_SHEET_NAME)
      .getRange('A1:H1')
      .setValues([[
        'quoteID','clientName','description','total','status','dateCreated','invoiceID','threadId'
      ]]);
  }
  if (!ss.getSheetByName(INVOICES_SHEET_NAME)) {
    ss.insertSheet(INVOICES_SHEET_NAME)
      .getRange('A1:F1')
      .setValues([[
        'invoiceID','quoteID','clientEmail','amount','status','dateCreated'
      ]]);
  }
  if (!ss.getSheetByName(BILLING_SHEET_NAME)) {
    ss.insertSheet(BILLING_SHEET_NAME)
      .getRange('A1:F1')
      .setValues([[
        'id','type','description','amount','status','email'
      ]]);
  }
}

// =========================
// ENTRY POINT WEB APP
// =========================
function doGet(e) {
  initialSetup();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Kônsul Billing')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

// =========================
// API CLIENTE
// =========================
function getBillingRecords() {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) throw new Error('Spreadsheet no configurado');
  const rows = SpreadsheetApp.openById(ssId)
    .getSheetByName(BILLING_SHEET_NAME)
    .getDataRange()
    .getValues();
  return rows.slice(1).map(r => ({
    id: r[0],
    type: r[1],
    description: r[2],
    amount: r[3],
    status: r[4],
    email: r[5]
  }));
}

// =========================
// CREAR COTIZACIÓN DESDE GEMINI
// =========================
function createQuoteFromNotes(notes) {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return { success: false };
  ensureSheets_();

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY no configurada');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent?key=' + apiKey;
  const prompt =
    'Eres un asistente que extrae datos de cotización. ' +
    'Del texto proporcionado, devuelve EXACTAMENTE un JSON con esta estructura:\n' +
    '{"clientName":"","items":[{"description":"","price":0}],"total":0,"summary":""}\n' +
    'Texto:\n"""' + notes + '"""';

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const textResponse = response.getContentText();
  Logger.log('Gemini raw response: ' + textResponse);

  let parsed;
  try {
    const candidate = JSON.parse(textResponse).candidates[0].content.parts[0].text.trim();
    Logger.log('Gemini candidate JSON: ' + candidate);
    parsed = JSON.parse(candidate);
  } catch (e) {
    Logger.log('Error parsing Gemini output: ' + e);
    parsed = { clientName: '', items: [], total: 0, summary: notes };
  }

  const { clientName, items, total, summary } = parsed;
  const ss = SpreadsheetApp.openById(ssId);

  // Guardar en Quotes
  const sheetQ = ss.getSheetByName(QUOTES_SHEET_NAME);
  const quoteID = 'quote_' + Date.now();
  sheetQ.appendRow([quoteID, clientName, summary, total, 'Draft', new Date(), '', '']);

  // Guardar en Billing (para UI)
  const sheetB = ss.getSheetByName(BILLING_SHEET_NAME);
  const desc = items.map(i => i.description + ': $' + i.price).join('; ') || summary;
  sheetB.appendRow([quoteID, 'Quote', desc, total, 'Draft', clientName]);

  return { success: true, quoteID: quoteID };
}

// =========================
// ENVÍO DE COTIZACIÓN MANUAL
// =========================
function sendQuote(id) {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return { success: false };
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(BILLING_SHEET_NAME);
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === id && vals[i][1] === 'Quote') {
      GmailApp.sendEmail(vals[i][5], 'Cotización',
        'Estimado ' + vals[i][5] + ', adjunto tu cotización. Monto: $' + vals[i][3]
      );
      sheet.getRange(i + 1, 5).setValue('Sent');
      return { success: true };
    }
  }
  return { success: false };
}

// =========================
// MARCAR FACTURA PAGADA
// =========================
function markInvoicePaid(id) {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return { success: false };
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(BILLING_SHEET_NAME);
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === id && vals[i][1] === 'Invoice') {
      sheet.getRange(i + 1, 5).setValue('Paid');
      return { success: true };
    }
  }
  return { success: false };
}

// =========================
// SEGUIMIENTO AUTOMÁTICO
// =========================
function followUpQuotesAndInvoices() {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return;
  const ss = SpreadsheetApp.openById(ssId);
  const today = new Date();

  ss.getSheetByName(QUOTES_SHEET_NAME).getDataRange().getValues().slice(1)
    .forEach(r => {
      const diff = Math.floor((today - new Date(r[5])) / (1000*60*60*24));
      if (r[4] === 'Sent' && diff > 3 && diff % 3 === 0) {
        GmailApp.sendEmail(r[1], 'Seguimiento cotización', 'Revisa tu cotización.');
      }
    });

  ss.getSheetByName(INVOICES_SHEET_NAME).getDataRange().getValues().slice(1)
    .forEach(r => {
      const diff = Math.floor((today - new Date(r[5])) / (1000*60*60*24));
      if (r[4] === 'Unpaid' && diff > 7 && diff % 7 === 0) {
        GmailApp.sendEmail(r[2], 'Recordatorio factura', 'Factura ' + r[0] + ' pendiente.');
      }
    });
}
