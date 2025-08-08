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
  const QUOTES_HEADERS   = ['quoteID','clientName','clientEmail','subject','item','total','status','quoteDate','invoiceID','threadId'];
  const INVOICES_HEADERS = ['invoiceID','quoteID','clientEmail','subject','item','amount','status','quoteDate','dateCreated'];

  let sheet = ss.getSheetByName(QUOTES_SHEET_NAME);
  if (!sheet) {
    ss.insertSheet(QUOTES_SHEET_NAME)
      .getRange(1,1,1,QUOTES_HEADERS.length)
      .setValues([QUOTES_HEADERS]);
  } else {
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const needsMigration = headers.length !== QUOTES_HEADERS.length || headers.some((h,i) => h !== QUOTES_HEADERS[i]);
    if (needsMigration) {
      const data = sheet.getDataRange().getValues().slice(1);
      let newData;
      if (headers.includes('description')) {
        // Migrar desde versión antigua
        newData = data.map(r => [
          r[0], // quoteID
          r[1], // clientName
          '',   // clientEmail no disponible
          '',   // subject no disponible
          r[2], // item desde description
          r[3], // total
          r[4], // status
          r[5], // quoteDate desde dateCreated
          r[6] || '',
          r[7] || ''
        ]);
      } else {
        newData = data.map(r => [
          r[0] || '',
          r[1] || '',
          '',
          r[2] || '',
          r[3] || '',
          r[4] || '',
          r[5] || '',
          r[6] || '',
          r[7] || '',
          r[8] || ''
        ]);
      }
      sheet.clear();
      sheet.getRange(1,1,1,QUOTES_HEADERS.length).setValues([QUOTES_HEADERS]);
      if (newData.length)
        sheet.getRange(2,1,newData.length,QUOTES_HEADERS.length).setValues(newData);
    }
  }
  sheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  if (!sheet) {
    ss.insertSheet(INVOICES_SHEET_NAME)
      .getRange(1,1,1,INVOICES_HEADERS.length)
      .setValues([INVOICES_HEADERS]);
  } else {
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const needsMigration = headers.length !== INVOICES_HEADERS.length || headers.some((h,i) => h !== INVOICES_HEADERS[i]);
    if (needsMigration) {
      const data = sheet.getDataRange().getValues().slice(1);
      let newData;
      if (headers.length === 6 && headers[3] === 'amount') {
        // Migrar desde versión antigua
        newData = data.map(r => [
          r[0], // invoiceID
          r[1], // quoteID
          r[2], // clientEmail
          '',   // subject
          '',   // item
          r[3], // amount
          r[4], // status
          '',   // quoteDate
          r[5]  // dateCreated
        ]);
      } else {
          newData = data.map(r => [
            r[0] || '', r[1] || '', r[2] || '', r[3] || '', r[4] || '', r[5] || '', r[6] || '', r[7] || '', r[8] || ''
          ]);
      }
      sheet.clear();
      sheet.getRange(1,1,1,INVOICES_HEADERS.length).setValues([INVOICES_HEADERS]);
      if (newData.length)
        sheet.getRange(2,1,newData.length,INVOICES_HEADERS.length).setValues(newData);
    }
  }

  if (!ss.getSheetByName(BILLING_SHEET_NAME)) {
    ss.insertSheet(BILLING_SHEET_NAME)
      .getRange('A1:H1')
      .setValues([[
         'id','type','description','amount','status','clientName','clientEmail','subject'
      ]]);
       } else {
    const sheetB = ss.getSheetByName(BILLING_SHEET_NAME);
    const headers = sheetB.getRange(1, 1, 1, sheetB.getLastColumn()).getValues()[0];
    let nameCol = headers.indexOf('clientName') + 1;
    if (!nameCol) {
      const emailCol = headers.indexOf('email') + 1;
      if (emailCol) {
        sheetB.getRange(1, emailCol).setValue('clientName');
        nameCol = emailCol;
      } else {
        sheetB.insertColumnAfter(5);
        sheetB.getRange(1, 6).setValue('clientName');
        nameCol = 6;
      }
    }
    if (headers.indexOf('clientEmail') === -1) {
      sheetB.insertColumnAfter(nameCol);
      sheetB.getRange(1, nameCol + 1).setValue('clientEmail');
    }
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
   const ss = SpreadsheetApp.openById(ssId);

  const billingRows = ss.getSheetByName(BILLING_SHEET_NAME)
    .getDataRange()
    .getValues();
   // Build a map of quote details to enrich billing records
  const quotesSheet = ss.getSheetByName(QUOTES_SHEET_NAME);
  const quotesData = quotesSheet ? quotesSheet.getDataRange().getValues() : [];
  const quoteHeaders = quotesData[0] || [];
  const qIdIdx = quoteHeaders.indexOf('quoteID');
  const qSubjectIdx = quoteHeaders.indexOf('subject');
  const qItemIdx = quoteHeaders.indexOf('item');
  const qDateIdx = quoteHeaders.indexOf('quoteDate');

  const quoteMap = {};
  for (let i = 1; i < quotesData.length; i++) {
    const row = quotesData[i];
    const id = qIdIdx >= 0 ? row[qIdIdx] : null;
    if (!id) continue;
    quoteMap[id] = {
      subject: qSubjectIdx >= 0 ? row[qSubjectIdx] : '',
      item: qItemIdx >= 0 ? row[qItemIdx] : '',
      quoteDate: qDateIdx >= 0 ? row[qDateIdx] : null
    };
  }

  return billingRows.slice(1).map(r => {
    const details = quoteMap[r[0]] || {};
    return {
      id: r[0],
      type: r[1],
      description: r[2],
      amount: r[3],
      status: r[4],
      clientName: r[5],
      clientEmail: r[6],
      subject: details.subject || '',
      item: details.item || '',
      quoteDate: details.quoteDate ? Utilities.formatDate(new Date(details.quoteDate), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
    };
  });
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
    '{"clientName":"","clientEmail":"","subject":"","date":"","item":"","amount":0}\n' +
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
     parsed = {};
  }

  const clientName = typeof parsed.clientName === 'string' ? parsed.clientName : '';
  const clientEmail = typeof parsed.clientEmail === 'string' ? parsed.clientEmail : '';
  const subject = typeof parsed.subject === 'string' ? parsed.subject : '';
  const dateStr = typeof parsed.date === 'string' ? parsed.date : '';
  const item = typeof parsed.item === 'string' ? parsed.item : '';
  const amount = typeof parsed.amount === 'number' ? parsed.amount : 0;
  let quoteDate = new Date(dateStr);
  if (!dateStr || isNaN(quoteDate.getTime())) quoteDate = new Date();

  const ss = SpreadsheetApp.openById(ssId);

  // Guardar en Quotes
  const sheetQ = ss.getSheetByName(QUOTES_SHEET_NAME);
  const quoteID = 'quote_' + Date.now();
  sheetQ.appendRow([quoteID, clientName, clientEmail, subject, item, amount, 'Draft', quoteDate, '', '']);

  // Guardar en Billing (para UI)
  const sheetB = ss.getSheetByName(BILLING_SHEET_NAME);
  const desc = item || subject;
  sheetB.appendRow([quoteID, 'Quote', desc, amount, 'Draft', clientName, clientEmail, subject]);

  return { success: true, quoteID: quoteID, clientName, clientEmail, subject, date: dateStr, item, amount };
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
      const email = vals[i][6];
      const name = vals[i][5];
      const subject = vals[i][7] || 'Cotización';
      GmailApp.sendEmail(
        email,
        subject,
        'Estimado ' + name + ', adjunto tu cotización. Monto: $' + vals[i][3]
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

  const billingSheet = ss.getSheetByName(BILLING_SHEET_NAME);
  const billingRows = billingSheet ? billingSheet.getDataRange().getValues() : [];
  const emailMap = {};
  for (let i = 1; i < billingRows.length; i++) {
    const r = billingRows[i];
    emailMap[r[0]] = r[6];
  }

 const qSheet = ss.getSheetByName(QUOTES_SHEET_NAME);
  if (qSheet) {
    const qData = qSheet.getDataRange().getValues();
    const qHeaders = qData[0] || [];
    const qIdIdx = qHeaders.indexOf('quoteID');
    const qStatusIdx = qHeaders.indexOf('status');
    const qDateIdx = qHeaders.indexOf('quoteDate');
    const qEmailIdx = qHeaders.indexOf('clientEmail');

    qData.slice(1).forEach(r => {
      const diff = Math.floor((today - new Date(r[qDateIdx])) / (1000 * 60 * 60 * 24));
      if (r[qStatusIdx] === 'Sent' && diff > 3 && diff % 3 === 0) {
        let clientEmail = qEmailIdx >= 0 ? r[qEmailIdx] : '';
        if (!clientEmail) clientEmail = emailMap[r[qIdIdx]] || '';
        if (clientEmail)
          GmailApp.sendEmail(clientEmail, 'Seguimiento cotización', 'Revisa tu cotización.');
      }
    });
    }

  const iSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  if (iSheet) {
    iSheet.getDataRange().getValues().slice(1)
      .forEach(r => {
        const diff = Math.floor((today - new Date(r[8])) / (1000 * 60 * 60 * 24));
        if (r[6] === 'Unpaid' && diff > 7 && diff % 7 === 0) {
          const clientEmail = r[2] || emailMap[r[0]] || '';
          if (clientEmail)
            GmailApp.sendEmail(clientEmail, 'Recordatorio factura', 'Factura ' + r[0] + ' pendiente.');
        }
      });
  }
}
