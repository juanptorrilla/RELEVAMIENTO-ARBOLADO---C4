// Relevamiento Comuna 4 — backend V7 ampliado
// Mantiene la escritura de datos del árbol en I:R y limita la actualización de altura a la columna C.
const TOKEN = 'C4-ARBOLADO-2026-PLANILLA-4';
const SPREADSHEET_ID = '1l_58opct0coaz2suIsez7UDtB5Wa1RiF';
const SHEET_GID = 1419998981;
const BASE_GENERAL_ID = '1DxbjbjFwOvhhBYI6i7_CjlrepAOxU7cSvr2BCMn_s_w';
const BASE_GENERAL_SHEET = 'BASE_RECLAMOS';

function doPost(e) {
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (p.token !== TOKEN) return json_({ok:false,error:'Token de acceso inválido.'});
    if (p.action === 'savePdf') return json_(savePdf_(p));
    if (p.action === 'listActas') return json_(listActas_());
    if (p.action === 'getActa') return json_(getActa_(p));
    if (p.action === 'saveTree') return json_(saveTree_(p.record || {}));
    if (p.action === 'denyClaim') return json_(denyClaim_(p));
    return json_({ok:false,error:'Acción no reconocida.'});
  } catch (err) { return json_({ok:false,error:String(err && err.message || err)}); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function savePdf_(p) {
  if (!p.data || !p.fileName) throw new Error('Faltan datos del PDF o nombre del archivo.');
  const bytes = Utilities.base64Decode(p.data);
  const blob = Utilities.newBlob(bytes, p.mimeType || 'application/pdf', p.fileName);
  const pdf = DriveApp.createFile(blob);
  const savedAt = new Date().toISOString();
  const sidecar = {app:'Relevamiento Comuna 4',pdfFileId:pdf.getId(),fileName:pdf.getName(),savedAt:savedAt,snapshot:p.snapshot || null};
  const record = DriveApp.createFile(p.fileName + '.acta.json', JSON.stringify(sidecar), MimeType.PLAIN_TEXT);
  return {ok:true,message:'PDF y datos de recuperación guardados en Google Drive.',fileId:pdf.getId(),recordId:record.getId(),fileName:pdf.getName()};
}
function listActas_() {
  const files = DriveApp.getFiles();
  const items = [];
  while (files.hasNext() && items.length < 100) {
    const f = files.next();
    if (!f.getName().endsWith('.acta.json') || f.isTrashed()) continue;
    try {
      const data = JSON.parse(f.getBlob().getDataAsString());
      if (data && data.app === 'Relevamiento Comuna 4' && data.snapshot) {
        items.push({fileId:f.getId(),fileName:data.fileName || f.getName().replace(/\.acta\.json$/,''),savedAt:data.savedAt || ''});
      }
    } catch (_) {}
  }
  items.sort((a,b)=>String(b.savedAt).localeCompare(String(a.savedAt)));
  return {ok:true,items:items};
}
function getActa_(p) {
  if (!p.fileId) throw new Error('Falta identificar el relevamiento de Drive.');
  const file = DriveApp.getFileById(p.fileId);
  const data = JSON.parse(file.getBlob().getDataAsString());
  if (!data || data.app !== 'Relevamiento Comuna 4' || !data.snapshot) throw new Error('El archivo no contiene datos recuperables de la aplicación.');
  return {ok:true,snapshot:data.snapshot,fileName:data.fileName || ''};
}

function saveTree_(r) {
  const aviso = String(r.aviso || '').trim();
  const submissionId = String(r.submissionId || '').trim();
  if (!aviso) throw new Error('El número de reclamo (AVISO) está vacío. No se modificó la planilla.');
  if (!submissionId) throw new Error('Falta el identificador del envío.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const sentKey = 'sent_' + Utilities.base64EncodeWebSafe(submissionId).slice(0,180);
    if (props.getProperty(sentKey)) {
      return {ok:true,message:'Este envío ya había sido registrado; no se duplicó.'};
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheets().find(s => s.getSheetId() === SHEET_GID);
    if (!sh) throw new Error('No se encontró la hoja configurada. No se modificó la planilla.');
    const last = sh.getLastRow();
    if (last < 2) throw new Error('La planilla no contiene números de reclamo.');

    // Buscar todas las filas que tengan el mismo número de reclamo en la columna A.
    const avisos = sh.getRange(2,1,last-1,1).getDisplayValues().flat().map(v => String(v).trim());
    const rows = [];
    avisos.forEach((v,i) => { if (v === aviso) rows.push(i+2); });
    if (!rows.length) {
      throw new Error('No se encontró el reclamo ' + aviso + ' en la columna A. No se escribió en ninguna fila.');
    }

    // Si hay una fila del reclamo sin datos del árbol en I:R, completar esa fila.
    let target = rows.find(row =>
      sh.getRange(row,9,1,10).getDisplayValues()[0].every(v => String(v).trim() === '')
    );
    let addedAdditionalTree = false;

    // Si todas las filas del reclamo ya tienen datos del árbol, agregar una fila
    // inmediatamente debajo de la última fila existente para ese mismo reclamo.
    if (!target) {
      const lastClaimRow = rows[rows.length - 1];
      sh.insertRowAfter(lastClaimRow);
      target = lastClaimRow + 1;
      // Copiar la estructura de la fila del mismo reclamo para conservar formato
      // y los datos generales existentes, sin escribir identificadores en la columna B.
      sh.getRange(lastClaimRow,1,1,24).copyTo(
        sh.getRange(target,1,1,24),
        {contentsOnly:false}
      );
      addedAdditionalTree = true;
    }

    const values = [[
      r.fecha || '',
      r.intervencion || '',
      r.especie || '',
      r.dap || '',
      r.altura || '',
      r.m2 || '',
      r.codigoBaldosa || '',
      r.riesgo || '',
      r.observacion || '',
      r.inspector || ''
    ]];

    // Los datos del árbol se escriben exclusivamente en I:R.
    // La columna B no se utiliza ni se modifica.
    const numeroRaw = r.numero != null ? r.numero : r['N°'];
    const numero = numeroRaw == null ? '' : String(numeroRaw);
    let chapaActualizada = false;
    if (numero.trim() !== '') {
      const cell = sh.getRange(target,3);
      const current = String(cell.getDisplayValue() == null ? '' : cell.getDisplayValue());
      if (current !== numero) {
        cell.setNumberFormat('@');
        cell.setValue(numero);
        chapaActualizada = true;
      }
    }

    sh.getRange(target,9,1,10).setValues(values);
    SpreadsheetApp.flush();
    props.setProperty(sentKey, new Date().toISOString());

    return {
      ok:true,
      message:(addedAdditionalTree
        ? 'Se agregó una línea para un árbol adicional'
        : 'Se completaron los datos del árbol') +
        ' del reclamo ' + aviso + ' (fila ' + target + ').' +
        (chapaActualizada ? ' Se actualizó la CHAPA CATASTRAL (columna C).' : ''),
      row:target,
      addedAdditionalTree:addedAdditionalTree,
      chapaActualizada:chapaActualizada
    };
  } finally {
    lock.releaseLock();
  }
}

function denyClaim_(p) {
  const aviso = String(p.aviso || '').trim();
  const reason = String(p.reason || '').trim();
  const extra = String(p.extra || '').trim().slice(0,200);
  if (!aviso || !reason) throw new Error('Falta el número de reclamo o el motivo de denegación.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(BASE_GENERAL_ID);
    const sh = ss.getSheetByName(BASE_GENERAL_SHEET);
    if (!sh) throw new Error('No se encontró la hoja BASE_RECLAMOS de la base general.');
    const last = sh.getLastRow();
    const avisos = sh.getRange(1,1,last,1).getDisplayValues().flat().map(v=>String(v).trim());
    const index = avisos.findIndex(v=>v===aviso);
    if (index < 0) throw new Error('No se encontró el reclamo ' + aviso + ' en la BASE GENERAL.');
    const row = index + 1;
    const email = String(sh.getRange(row,7).getDisplayValue() || '').trim();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('El reclamo no tiene una dirección de correo válida en la columna G.');
    const prior = String(sh.getRange(row,13).getDisplayValue() || '').toUpperCase();
    if (prior.indexOf('DENEGACIÓN ENVIADA') >= 0 || prior.indexOf('DENEGACION ENVIADA') >= 0) return {ok:true,message:'La notificación de denegación ya figura enviada. No se volvió a enviar.'};
    const subject = 'Respuesta a su reclamo de arbolado — ' + aviso;
    const body = 'Estimado/a vecino/a:\n\nEn relación con el reclamo ' + aviso + ', le informamos que no corresponde realizar la intervención solicitada por el siguiente motivo:\n\n' + reason + (extra ? '\n\n' + extra : '') + '\n\nSaludos cordiales,\nÁrea de Arbolado\nComuna 4';
    MailApp.sendEmail({to:email,subject:subject,body:body,name:'Arbolado Comuna 4'});
    sh.getRange(row,11).setValue('DENEGADO');
    sh.getRange(row,13).setValue('DENEGACIÓN ENVIADA');
    sh.getRange(row,14).setValue(new Date());
    SpreadsheetApp.flush();
    return {ok:true,message:'Correo de denegación enviado y notificación registrada en la BASE GENERAL.'};
  } finally { lock.releaseLock(); }
}
