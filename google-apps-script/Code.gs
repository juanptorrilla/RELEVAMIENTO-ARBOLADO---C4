/**
 * Relevamiento Arbolado - Comuna 4
 * Completa exclusivamente las columnas I:R de la pestaña de reclamos.
 * Si el reclamo ya tiene un árbol cargado, agrega una fila nueva copiando A:E
 * del reclamo original; nunca modifica las filas existentes ni las columnas
 * de fórmulas/controles F:H y S:X.
 */
const SPREADSHEET_ID = '1l_58opct0coaz2suIsez7UDtB5Wa1RiF';
const TARGET_GID = 1068617970;
const SHARED_TOKEN = 'C4-ARBOLADO-2026-PLANILLA-4';

function doGet(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || '');
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput('/* callback inválido */').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  let result;
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error('No se recibieron datos.');
    result = saveTreeRecord_(JSON.parse(raw));
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function saveTreeRecord_(payload) {
  if (!payload || payload.token !== SHARED_TOKEN) throw new Error('Token de conexión inválido.');
  if (payload.spreadsheetId !== SPREADSHEET_ID || Number(payload.gid) !== TARGET_GID) throw new Error('La planilla o pestaña de destino no coincide.');
  const record = payload.record || {};
  const claim = String(record.aviso || '').trim();
  const submissionId = String(payload.submissionId || '').trim();
  if (!claim) throw new Error('Falta el número de reclamo / AVISO.');
  if (!submissionId) throw new Error('Falta el identificador del envío.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const propertyKey = 'submission_' + submissionId.replace(/[^A-Za-z0-9_-]/g, '_');
    const priorRow = props.getProperty(propertyKey);
    if (priorRow) return { ok: true, duplicate: true, row: Number(priorRow) };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets().find(s => s.getSheetId() === TARGET_GID);
    if (!sheet) throw new Error('No se encontró la pestaña de destino (gid ' + TARGET_GID + ').');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('La pestaña no contiene reclamos cargados para localizar.');
    // A:E son los datos originales del reclamo. Se leen, nunca se editan en su fila existente.
    const sourceValues = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
    const normalizeClaim = v => String(v || '').replace(/\s+/g, '').toUpperCase();
    const matches = [];
    sourceValues.forEach((row, i) => {
      if (normalizeClaim(row[0]) === normalizeClaim(claim)) matches.push({ rowNumber: i + 2, values: row });
    });
    if (!matches.length) throw new Error('No se encontró ese número de reclamo en la columna A. No se modificó la planilla.');

    // I:R: fecha, intervención del inspector, especie, DAP, altura, M2,
    // código de baldosa, riesgo, observación del inspector e inspector.
    const values = [
      record.fecha || '', record.intervencion || '', record.especie || '', record.dap || '',
      record.alturaArbol || '', record.superficie || '', record.codigoBaldosa || '',
      record.riesgo || '', record.observaciones || '', record.inspector || ''
    ];

    // Reintento idempotente: si ya existe exactamente esta ficha para el reclamo,
    // no vuelve a agregarla. La comparación se limita a I:R.
    for (const match of matches) {
      const existing = sheet.getRange(match.rowNumber, 9, 1, 10).getDisplayValues()[0];
      if (sameValues_(existing, values)) {
        props.setProperty(propertyKey, String(match.rowNumber));
        return { ok: true, duplicate: true, row: match.rowNumber };
      }
    }

    // Si el reclamo tiene una fila sin datos de relevamiento, se utiliza esa fila.
    let targetRow = null;
    for (const match of matches) {
      const existing = sheet.getRange(match.rowNumber, 9, 1, 10).getDisplayValues()[0];
      if (existing.every(v => String(v).trim() === '')) { targetRow = match.rowNumber; break; }
    }

    if (targetRow === null) {
      // Árbol adicional: duplicar la fila del reclamo como base para conservar
      // también las fórmulas, validaciones y datos auxiliares de las demás
      // columnas. La fila original nunca se modifica. Luego se reemplaza
      // únicamente I:R con los datos del nuevo ejemplar.
      targetRow = Math.max(sheet.getLastRow() + 1, 2);
      const sourceRow = matches[0].rowNumber;
      sheet.getRange(sourceRow, 1, 1, 24).copyTo(
        sheet.getRange(targetRow, 1, 1, 24),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );
      // Reafirmar A:E como valores visibles del reclamo original, sin alterar
      // el contenido de la fila de origen.
      sheet.getRange(targetRow, 1, 1, 5).setValues([matches[0].values]);
    }
    // En la fila destinada a este árbol se completan exclusivamente I:R.
    // A:E y F:H / S:X quedan copiados desde la fila original si fue necesario
    // crear una línea adicional; la fila original permanece intacta.
    sheet.getRange(targetRow, 9, 1, 10).setValues([values]);
    props.setProperty(propertyKey, String(targetRow));
    return { ok: true, duplicate: false, row: targetRow, sheet: sheet.getName() };
  } finally {
    lock.releaseLock();
  }
}

function sameValues_(existing, incoming) {
  return existing.length === incoming.length && existing.every((v, i) => String(v || '').trim() === String(incoming[i] == null ? '' : incoming[i]).trim());
}
