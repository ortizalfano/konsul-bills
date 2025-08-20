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

// Estados permitidos
const QUOTE_STATUSES = ['Draft','Sent'];
const INVOICE_STATUSES = ['Draft','Sent','Paid','Cancelled'];

// =========================
// SETUP INICIAL
// =========================
function initialSetup() {
  const props = PropertiesService.getUserProperties();
  let ssId = props.getProperty('spreadsheetId');
  if (ssId) {
    try { SpreadsheetApp.openById(ssId); }
    catch(e) { ssId = null; }
  }
  if (!ssId) {
    const ss = SpreadsheetApp.create('Kônsul - Datos de Facturación');
    props.setProperty('spreadsheetId', ss.getId());
    ensureSheets_();
  }
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(t =>
    t.getHandlerFunction() === 'processGmailMessages' &&
    t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
  );
  if (!hasTrigger) setupTriggers();
}

function setupTriggers() {
  ScriptApp.newTrigger('processGmailMessages')
    .timeBased()
    .everyMinutes(30)
    .create();
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

function ensureDriveFolders() {
  const props = PropertiesService.getUserProperties();

  const getOrCreate = (propName, names) => {
    let id = props.getProperty(propName);
    if (id) {
      try {
        DriveApp.getFolderById(id);
        return id;
      } catch (e) {
        id = null;
      }
    }
    let folder = null;
    for (const name of names) {
      const it = DriveApp.getFoldersByName(name);
      if (it.hasNext()) {
        folder = it.next();
        break;
      }
    }
    if (!folder) {
      folder = DriveApp.createFolder(names[0]);
    }
    id = folder.getId();
    props.setProperty(propName, id);
    return id;
  };

  const quotesFolderId = getOrCreate('quotesFolderId', ['Cotizaciones', 'Quotes']);
  const invoicesFolderId = getOrCreate('invoicesFolderId', ['Facturas', 'Invoices']);
  return { quotesFolderId, invoicesFolderId };
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
    let status = r[4];
    if (r[1] === 'Quote') {
      if (!QUOTE_STATUSES.includes(status)) status = 'Draft';
    } else if (r[1] === 'Invoice') {
      if (!INVOICE_STATUSES.includes(status)) status = 'Draft';
    }
    return {
      id: r[0],
      type: r[1],
      description: r[2],
      amount: r[3],
      status: status,
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
  Logger.log('Iniciando creación de cotización desde notas');
  if (typeof notes !== 'string' || notes.trim() === '') {
    Logger.log('Notas vacías o no provistas');
    return { success: false, error: 'Notas vacías' };
  }
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

  Logger.log('Enviando petición a Gemini');
  const response = UrlFetchApp.fetch(url, options);
    const status = response.getResponseCode();
  Logger.log('Código de respuesta: ' + status);
  if (status !== 200) {
    Logger.log('Error en llamada a Gemini: ' + response.getContentText());
    return { success: false, error: 'Error en API Gemini', status: status };
  }


  const textResponse = response.getContentText();
  Logger.log('Gemini raw response: ' + textResponse);

  let parsed = {};
  try {
     const candidate = JSON.parse(textResponse).candidates[0].content.parts[0].text || '';
    Logger.log('Gemini candidate: ' + candidate);
    let cleaned = candidate.replace(/```json|```/gi, '').trim();
    try {
      parsed = JSON.parse(cleaned);
    } catch (err) {
      Logger.log('Error parseando candidato, intentando limpiar: ' + err);
      const start = cleaned.indexOf('{');
      const end = cleaned.lastIndexOf('}');
      if (start >= 0 && end >= 0) {
        const sub = cleaned.substring(start, end + 1);
        try { parsed = JSON.parse(sub); }
        catch (err2) { Logger.log('Fallback parse fallido: ' + err2); }
      }
    }
  
  } catch (e) {
    Logger.log('Error general parsing Gemini output: ' + e);
  }

  const fields = { clientName: '', clientEmail: '', subject: '', date: '', item: '', amount: 0 };
  Object.keys(fields).forEach(k => {
    const v = parsed[k];
    if (k === 'amount') {
      if (typeof v === 'number') fields[k] = v;
    } else if (typeof v === 'string') {
      fields[k] = v;
    }
  });

  if (!fields.clientEmail) {
    const m = notes.match(/[\w.%+-]+@[\w.-]+\.[A-Za-z]{2,}/);
    if (m) fields.clientEmail = m[0];
  }
  if (!fields.amount) {
    const m = notes.match(/\d+(?:[.,]\d+)?/);
    if (m) fields.amount = parseFloat(m[0].replace(',', '.'));
  }
  if (!fields.date) {
    const m = notes.match(/\d{4}-\d{2}-\d{2}/) || notes.match(/\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}/);
    if (m) fields.date = m[0];
  }
  if (!fields.clientName) {
    const m = notes.match(/(?:client|cliente)[:\s]+([\w\s]+)/i);
    if (m) fields.clientName = m[1].trim();
  }
  if (!fields.subject) {
    const m = notes.match(/(?:subject|asunto)[:\s]+(.+)/i);
    fields.subject = m ? m[1].split('\n')[0].trim() : notes.split('\n')[0].trim();
  }
  if (!fields.item) {
    const m = notes.match(/(?:item|concepto|producto)[:\s]+(.+)/i);
    if (m) fields.item = m[1].split('\n')[0].trim();
  }

  if (!fields.clientName)  fields.clientName  = 'Por Actualizar';
  if (!fields.clientEmail) fields.clientEmail = 'N/A';

  fields.amount = parseFloat(fields.amount) || 0;
  let quoteDate = new Date(fields.date);
  if (isNaN(quoteDate.getTime())) quoteDate = new Date();

  Object.keys(fields).forEach(k => {
    if (fields[k] === '' || fields[k] === 0) Logger.log('Campo faltante o vacío: ' + k);
  });

  const ss = SpreadsheetApp.openById(ssId);

  // Guardar en Quotes
  const sheetQ = ss.getSheetByName(QUOTES_SHEET_NAME);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, error: 'No se pudo generar ID' };
  }
  let quoteID;
  try {
    const last = sheetQ.getRange(sheetQ.getLastRow(), 1).getValue();
    const next = (parseInt(String(last).replace('COT', '')) + 1 || 1)
      .toString()
      .padStart(5, '0');
    quoteID = 'COT' + next;
  } finally {
    lock.releaseLock();
  }
  const clientName = fields.clientName;
  const clientEmail = fields.clientEmail;
  sheetQ.appendRow([
    quoteID,
    clientName,
    clientEmail,
    fields.subject,
    fields.item,
    fields.amount,
    'Draft',
    quoteDate,
    '',
    ''
  ]);

  // Guardar en Billing (para UI)
  const sheetB = ss.getSheetByName(BILLING_SHEET_NAME);
  const desc = fields.item || fields.subject;
  sheetB.appendRow([
    quoteID,
    'Quote',
    desc,
    fields.amount,
    'Draft',
    clientName,
    clientEmail,
    fields.subject
  ]);

  Logger.log('Cotización creada con ID: ' + quoteID);
  return { success: true, quoteID: quoteID, ...fields };
}

function createInvoiceFromNotes(notes) {
  Logger.log('Iniciando creación de factura desde notas');
  if (typeof notes !== 'string' || notes.trim() === '') {
    Logger.log('Notas vacías o no provistas');
    return { success: false, error: 'Notas vacías' };
  }
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return { success: false };
  ensureSheets_();

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY no configurada');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent?key=' + apiKey;
  const prompt =
    'Eres un asistente que extrae datos de factura. ' +
    'Del texto proporcionado, devuelve EXACTAMENTE un JSON con esta estructura:\n' +
    '{"clientName":"","clientEmail":"","subject":"","date":"","item":"","amount":0}\n' +
    'Texto:\n"""' + notes + '"""';

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  };

  Logger.log('Enviando petición a Gemini');
  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  Logger.log('Código de respuesta: ' + status);
  if (status !== 200) {
    Logger.log('Error en llamada a Gemini: ' + response.getContentText());
    return { success: false, error: 'Error en API Gemini', status: status };
  }

  const textResponse = response.getContentText();
  Logger.log('Gemini raw response: ' + textResponse);

  let parsed = {};
  try {
    const candidate = JSON.parse(textResponse).candidates[0].content.parts[0].text || '';
    Logger.log('Gemini candidate: ' + candidate);
    let cleaned = candidate.replace(/```json|```/gi, '').trim();
    try {
      parsed = JSON.parse(cleaned);
    } catch (err) {
      Logger.log('Error parseando candidato, intentando limpiar: ' + err);
      const start = cleaned.indexOf('{');
      const end = cleaned.lastIndexOf('}');
      if (start >= 0 && end >= 0) {
        const sub = cleaned.substring(start, end + 1);
        try { parsed = JSON.parse(sub); }
        catch (err2) { Logger.log('Fallback parse fallido: ' + err2); }
      }
    }

  } catch (e) {
    Logger.log('Error general parsing Gemini output: ' + e);
  }

  const fields = { clientName: '', clientEmail: '', subject: '', date: '', item: '', amount: 0 };
  Object.keys(fields).forEach(k => {
    const v = parsed[k];
    if (k === 'amount') {
      if (typeof v === 'number') fields[k] = v;
    } else if (typeof v === 'string') {
      fields[k] = v;
    }
  });

  if (!fields.clientEmail) {
    const m = notes.match(/[\w.%+-]+@[\w.-]+\.[A-Za-z]{2,}/);
    if (m) fields.clientEmail = m[0];
  }
  if (!fields.amount) {
    const m = notes.match(/\d+(?:[.,]\d+)?/);
    if (m) fields.amount = parseFloat(m[0].replace(',', '.'));
  }
  if (!fields.date) {
    const m = notes.match(/\d{4}-\d{2}-\d{2}/) || notes.match(/\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}/);
    if (m) fields.date = m[0];
  }
  if (!fields.clientName) {
    const m = notes.match(/(?:client|cliente)[:\s]+([\w\s]+)/i);
    if (m) fields.clientName = m[1].trim();
  }
  if (!fields.subject) {
    const m = notes.match(/(?:subject|asunto)[:\s]+(.+)/i);
    fields.subject = m ? m[1].split('\n')[0].trim() : notes.split('\n')[0].trim();
  }
  if (!fields.item) {
    const m = notes.match(/(?:item|concepto|producto)[:\s]+(.+)/i);
    if (m) fields.item = m[1].split('\n')[0].trim();
  }

  if (!fields.clientName)  fields.clientName  = 'Por Actualizar';
  if (!fields.clientEmail) fields.clientEmail = 'N/A';

  fields.amount = parseFloat(fields.amount) || 0;
  let invoiceDate = new Date(fields.date);
  if (isNaN(invoiceDate.getTime())) invoiceDate = new Date();

  Object.keys(fields).forEach(k => {
    if (fields[k] === '' || fields[k] === 0) Logger.log('Campo faltante o vacío: ' + k);
  });

  const ss = SpreadsheetApp.openById(ssId);

  // Guardar en Invoices
  const sheetI = ss.getSheetByName(INVOICES_SHEET_NAME);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, error: 'No se pudo generar ID' };
  }
  let invoiceID;
  try {
    const last = sheetI.getRange(sheetI.getLastRow(), 1).getValue();
    const next = (parseInt(String(last).replace('INV', '')) + 1 || 1)
      .toString()
      .padStart(3, '0');
    invoiceID = 'INV' + next;
  } finally {
    lock.releaseLock();
  }
  sheetI.appendRow([
    invoiceID,
    '',
    fields.clientEmail,
    fields.subject,
    fields.item,
    fields.amount,
    'Draft',
    invoiceDate,
    new Date()
  ]);

  // Guardar en Billing (para UI)
  const sheetB = ss.getSheetByName(BILLING_SHEET_NAME);
  const desc = fields.item || fields.subject;
  sheetB.appendRow([
    invoiceID,
    'Invoice',
    desc,
    fields.amount,
    'Draft',
    fields.clientName,
    fields.clientEmail,
    fields.subject
  ]);

  // Generar PDF opcional y mover a carpeta si existe
  try {
    const file = generateInvoicePdf(invoiceID);
    const folderId = PropertiesService.getScriptProperties().getProperty('INVOICES_FOLDER_ID');
    if (folderId && file) {
      DriveApp.getFolderById(folderId).addFile(file);
    }
  } catch (err) {
    Logger.log('Error generando PDF: ' + err);
  }

  Logger.log('Factura creada con ID: ' + invoiceID);
  return { success: true, invoiceID: invoiceID, ...fields };
}

// =========================
// ACTUALIZAR COTIZACIÓN
// =========================
function updateQuote(data) {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return { success:false };
  const ss = SpreadsheetApp.openById(ssId);
  const qSheet = ss.getSheetByName(QUOTES_SHEET_NAME);
  const bSheet = ss.getSheetByName(BILLING_SHEET_NAME);
  const qVals = qSheet.getDataRange().getValues();
  const bVals = bSheet.getDataRange().getValues();

  if (!QUOTE_STATUSES.includes(data.status)) data.status = 'Draft';

  for (let i = 1; i < qVals.length; i++) {
    if (qVals[i][0] === data.id) {
      qSheet.getRange(i+1,2,1,6).setValues([[data.clientName, data.clientEmail,
        data.subject, data.item, data.amount, data.status]]);
      break;
    }
  }
  for (let j = 1; j < bVals.length; j++) {
    if (bVals[j][0] === data.id) {
      bSheet.getRange(j+1,3,1,5).setValues([[data.item || data.subject,
        data.amount, data.status, data.clientName, data.clientEmail]]);
      break;
    }
  }
  return { success:true };
}

// =========================
// TRANSFORMAR COTIZACIÓN A FACTURA
// =========================
function transformQuoteToInvoice(id) {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return { success: false };
  const ss = SpreadsheetApp.openById(ssId);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const qSheet = ss.getSheetByName(QUOTES_SHEET_NAME);
    const iSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
    const bSheet = ss.getSheetByName(BILLING_SHEET_NAME);
    if (!qSheet || !iSheet || !bSheet) return { success: false };

    const qData = qSheet.getDataRange().getValues();
    let qRow = null, qIndex = -1;
    for (let i = 1; i < qData.length; i++) {
      if (qData[i][0] === id) {
        qRow = qData[i];
        qIndex = i;
        break;
      }
    }
    if (!qRow) return { success: false };

    const iData = iSheet.getDataRange().getValues();
    let maxNum = 0;
    for (let i = 1; i < iData.length; i++) {
      const match = String(iData[i][0]).match(/^INV(\d+)$/);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
    const invoiceID = 'INV' + Utilities.formatString('%03d', maxNum + 1);

    iSheet.appendRow([
      invoiceID,
      id,
      qRow[2],
      qRow[3],
      qRow[4],
      qRow[5],
      'Draft',
      qRow[7],
      new Date()
    ]);

    const bData = bSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      if (bData[i][0] === id && bData[i][1] === 'Quote') {
        bSheet.getRange(i + 1, 1, 1, 5).setValues([[invoiceID, 'Invoice', bData[i][2], qRow[5], 'Draft']]);
        break;
      }
    }

    if (qIndex > 0) {
      qSheet.getRange(qIndex + 1, 9).setValue(invoiceID);
    }

    return { success: true, invoiceID: invoiceID };
  } finally {
    lock.releaseLock();
  }
}


// =========================

// =========================  
// ENVÍO DE COTIZACIÓN MANUAL
// =========================
function sendQuote(id) {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return { success: false };
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(BILLING_SHEET_NAME);
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

      const qSheet = ss.getSheetByName(QUOTES_SHEET_NAME);
      const qData  = qSheet.getDataRange().getValues();
      for (let q = 1; q < qData.length; q++) {
        if (qData[q][0] === id) {
          qSheet.getRange(q + 1, 7).setValue('Sent');
          break;
        }
      }
      return { success: true };
    }
  }
  return { success: false };
}

// =========================
// ENVÍO DE FACTURA MANUAL
// =========================
function sendInvoice(id) {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!ssId) return { success: false };
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(BILLING_SHEET_NAME);
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === id && vals[i][1] === 'Invoice') {
      const email = vals[i][6];
      const name = vals[i][5];
      const subject = vals[i][7] || 'Factura';
      GmailApp.sendEmail(
        email,
        subject,
        'Estimado ' + name + ', adjunto tu factura. Monto: $' + vals[i][3]
      );
      sheet.getRange(i + 1, 5).setValue('Sent');

      const iSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
      const iData = iSheet.getDataRange().getValues();
      for (let j = 1; j < iData.length; j++) {
        if (iData[j][0] === id) {
          iSheet.getRange(j + 1, 7).setValue('Sent');
          break;
        }
      }
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
  const ss = SpreadsheetApp.openById(ssId);
  const bSheet = ss.getSheetByName(BILLING_SHEET_NAME);
  const bVals = bSheet.getDataRange().getValues();
  for (let i = 1; i < bVals.length; i++) {
    if (bVals[i][0] === id && bVals[i][1] === 'Invoice') {
      bSheet.getRange(i + 1, 5).setValue('Paid');
      const iSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
      if (iSheet) {
        const iVals = iSheet.getDataRange().getValues();
        for (let j = 1; j < iVals.length; j++) {
          if (iVals[j][0] === id) {
            iSheet.getRange(j + 1, 7).setValue('Paid');
            break;
          }
        }
      }
      return { success: true };
    }
  }
  return { success: false };
}

// =========================
// GENERAR PDF DE FACTURA
// =========================
function generateInvoicePdf(invoiceId) {
  const ssId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  const ss    = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data.find(r => r[0] === invoiceId);
  if (!row) return null;

  const tmpl = HtmlService.createTemplateFromFile('InvoicePdf');
  headers.forEach((h, i) => tmpl[h] = row[i]);
  const html = tmpl.evaluate().getContent();
  const blob = Utilities.newBlob(html, 'text/html')
    .getAs('application/pdf')
    .setName(invoiceId + '.pdf');
  return DriveApp.createFile(blob); // o devolver blob si se prefiere
}

// =========================
// PROCESAR MENSAJES GMAIL
// =========================
function processGmailMessages() {
  const labelNames = ['Cotizaciones', 'Quotes', 'Facturas', 'Invoices'];
  const processedLabel = GmailApp.getUserLabelByName('Processed') || GmailApp.createLabel('Processed');
  labelNames.forEach(name => {
    const label = GmailApp.getUserLabelByName(name);
    if (!label) return;
    const threads = label.getUnreadThreads();
    threads.forEach(thread => {
      const messages = thread.getMessages();
      if (!messages.length) return;
      const body = messages[0].getBody();
      if (name === 'Cotizaciones' || name === 'Quotes') {
        createQuoteFromEmail(body);
      } else if (name === 'Facturas' || name === 'Invoices') {
        createInvoiceFromEmail(body);
      }
      thread.markRead();
      thread.addLabel(processedLabel);
      thread.removeLabel(label);
    });
  });
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
      const qDate = new Date(r[qDateIdx]);
      if (isNaN(qDate)) {
        Logger.log('Invalid quoteDate for ID: ' + r[qIdIdx]);
        return;
      }
      const diff = Math.floor((today - qDate) / (1000 * 60 * 60 * 24));
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
        const iDate = new Date(r[8]);
        if (isNaN(iDate)) {
          Logger.log('Invalid invoice date for ID: ' + r[0]);
          return;
        }
        const diff = Math.floor((today - iDate) / (1000 * 60 * 60 * 24));
        if (r[6] === 'Sent' && diff > 7 && diff % 7 === 0) {
          const clientEmail = r[2] || emailMap[r[0]] || '';
          if (clientEmail)
            GmailApp.sendEmail(clientEmail, 'Recordatorio factura', 'Factura ' + r[0] + ' pendiente.');
        }
      });
  }
}

function processGmailMessages() {
  const labelName = 'konsul-billing';
  const threads = GmailApp.search('label:' + labelName + ' is:unread');
  const processedLabel = GmailApp.getUserLabelByName(labelName + '-processed') || GmailApp.createLabel(labelName + '-processed');

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      if (!msg.isUnread()) return;
      const body = msg.getPlainBody();
      const subject = msg.getSubject().toLowerCase();
      let res;
      if (subject.includes('invoice') || subject.includes('factura')) {
        res = createInvoiceFromNotes(body);
      } else {
        res = createQuoteFromNotes(body);
      }
      msg.markRead();
      processedLabel.addToThread(thread);
      Logger.log('Processed email with result: ' + JSON.stringify(res));
    });
  });
}
