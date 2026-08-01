/** SITI rápido: Apps Script se usa sólo para guardar o cerrar tickets. */
const CONFIG = {
  spreadsheetId: '12XIxmGGvFk1QE5qgfA2ULO1fYGW6ZNM5ucJmrFrMA7g',
  timeZone: 'America/Monterrey',
  ticketsSheet: 'TICKETS',
  appUrl: 'https://siti-prodiesa.vercel.app',
  driveFolderName: 'SITI - Tickets de Mantenimiento'
};

let spreadsheetCache;
let folderCache;

function doGet(e) {
  try {
    const action = String(e.parameter.action || '').toLowerCase();
    if (action === 'ticket') return json_(getTicket_(e.parameter.folio || ''));
    return json_({ ok: false, message: 'Acción no reconocida.' });
  } catch (error) {
    return json_({ ok: false, message: error.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').toLowerCase();
    if (action === 'open') return json_(openTicket_(body));
    if (action === 'close') return json_(closeTicket_(body));
    return json_({ ok: false, message: 'Acción no reconocida.' });
  } catch (error) {
    return json_({ ok: false, message: error.message });
  }
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  if (!spreadsheetCache) spreadsheetCache = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  return spreadsheetCache;
}

function tickets_() {
  const sheet = ss_().getSheetByName(CONFIG.ticketsSheet);
  if (!sheet) throw new Error('No existe la hoja TICKETS.');
  return sheet;
}

function headers_(sheet) {
  const names = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return names.reduce((map, name, index) => {
    map[String(name).trim()] = index;
    return map;
  }, {});
}

function dateInfo_(date) {
  const value = date || new Date();
  const day = Utilities.formatDate(value, CONFIG.timeZone, 'dd');
  const monthNumber = Utilities.formatDate(value, CONFIG.timeZone, 'MM');
  const year = Utilities.formatDate(value, CONFIG.timeZone, 'yyyy');
  const time = Utilities.formatDate(value, CONFIG.timeZone, 'HH:mm');
  const month = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'][Number(monthNumber) - 1];
  return { year, date: `${day}/${monthNumber}/${year}`, time, month, dateTime: `${day}/${monthNumber}/${year} ${time}` };
}

function nextFolio_(year) {
  const properties = PropertiesService.getScriptProperties();
  const key = `siti_folio_${year}`;
  let last = Number(properties.getProperty(key) || 0);

  if (!last) {
    const sheet = tickets_();
    const rows = Math.max(0, sheet.getLastRow() - 1);
    if (rows) {
      const values = sheet.getRange(2, 1, rows, 3).getValues();
      last = values.reduce((max, row) => String(row[1]) === String(year) ? Math.max(max, Number(row[2]) || 0) : max, 0);
    }
  }

  const consecutive = last + 1;
  properties.setProperty(key, String(consecutive));
  return { consecutive, folio: `MTTO-${year}-${String(consecutive).padStart(5, '0')}` };
}

function folder_() {
  if (folderCache) return folderCache;
  const found = DriveApp.getFoldersByName(CONFIG.driveFolderName);
  folderCache = found.hasNext() ? found.next() : DriveApp.createFolder(CONFIG.driveFolderName);
  return folderCache;
}

function savePhoto_(base64, name) {
  if (!base64) return '';
  const clean = String(base64).replace(/^data:image\/[a-zA-Z+]+;base64,/, '');
  const blob = Utilities.newBlob(Utilities.base64Decode(clean), 'image/jpeg', name);
  return folder_().createFile(blob).getUrl();
}

function set_(row, headers, name, value) {
  if (headers[name] !== undefined) row[headers[name]] = value;
}

function openTicket_(body) {
  const required = [
    ['machineId', 'máquina'],
    ['operatorId', 'operador'],
    ['description', 'descripción de la falla']
  ];
  const missing = required
    .filter(([field]) => !String(body[field] || '').trim())
    .map(([, label]) => label);
  if (missing.length) throw new Error(`Faltan datos obligatorios: ${missing.join(', ')}.`);

  // La máquina llega desde el QR. Estos respaldos evitan rechazar un ticket
  // si una liga antigua no trae el nombre o el área, sin perder el folio.
  const machineName = String(body.machineName || body.machineId).trim();
  const area = String(body.area || 'Área por confirmar').trim();
  const operatorName = String(body.operatorName || body.operatorId).trim();
  const impact = String(body.priority || 'Sin clasificar').trim();
  const productionStop = /Paro por calidad|Producción detenida|Producción parcial/i.test(impact) ? 'Sí' : 'No';
  const qualityStop = /Paro por calidad/i.test(impact) ? 'Sí' : 'No';

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = dateInfo_();
    const ticket = nextFolio_(now.year);
    const link = `${CONFIG.appUrl}?ticket=${encodeURIComponent(ticket.folio)}&maquina=${encodeURIComponent(body.machineId)}&area=${encodeURIComponent(area)}&nombre=${encodeURIComponent(machineName)}`;
    const photo = savePhoto_(body.photoBase64, `${ticket.folio}-apertura.jpg`);
    const sheet = tickets_();
    const headers = headers_(sheet);
    const row = Array(sheet.getLastColumn()).fill('');

    set_(row, headers, 'Folio', ticket.folio); set_(row, headers, 'Año', Number(now.year)); set_(row, headers, 'Consecutivo', ticket.consecutive); set_(row, headers, 'Estatus', 'Abierto');
    set_(row, headers, 'Fecha apertura', now.date); set_(row, headers, 'Mes', now.month); set_(row, headers, 'Hora apertura', now.time);
    set_(row, headers, 'ID operador', body.operatorId); set_(row, headers, 'Operador', operatorName);
    set_(row, headers, 'ID máquina', body.machineId); set_(row, headers, 'Área', area); set_(row, headers, 'Máquina', machineName);
    set_(row, headers, 'Prioridad', impact); set_(row, headers, '¿Paro de producción?', productionStop); set_(row, headers, '¿Paro por calidad?', qualityStop);
    set_(row, headers, 'Parte que falla', body.part || ''); set_(row, headers, 'Tipo de falla', body.failureType || ''); set_(row, headers, 'Descripción de falla', body.description);
    set_(row, headers, 'Foto apertura (Drive)', photo); set_(row, headers, 'Liga de ticket', link);
    sheet.appendRow(row);

    return {
      ok: true,
      folio: ticket.folio,
      ticketUrl: link,
      shareText: `🚨 Ticket abierto ${ticket.folio}\nMotivo: ${body.description}\nÁrea: ${area}\nMáquina: ${machineName}\nFecha y hora: ${now.dateTime}\nCerrar ticket: ${link}`
    };
  } finally {
    lock.releaseLock();
  }
}

function parseDate_(dateText, timeText) {
  const parts = String(dateText || '').split('/').map(Number);
  const time = String(timeText || '00:00').split(':').map(Number);
  return parts.length === 3 ? new Date(parts[2], parts[1] - 1, parts[0], time[0] || 0, time[1] || 0) : null;
}

function mins_(from, to) {
  return from && to ? Math.max(0, Math.round((to - from) / 60000)) : '';
}

function getTicket_(folio) {
  const sheet = tickets_();
  const headers = headers_(sheet);
  const data = sheet.getDataRange().getDisplayValues();
  const index = data.findIndex((row, rowIndex) => rowIndex && row[headers.Folio] === folio);
  if (index < 0) return { ok: false, message: 'No se encontró ese ticket.' };
  const ticket = data[index].reduce((object, value, column) => {
    const header = Object.keys(headers).find(key => headers[key] === column);
    object[header] = value;
    return object;
  }, {});
  return { ok: true, ticket };
}

function closeTicket_(body) {
  const required = [
    ['folio', 'folio'],
    ['technicianId', 'técnico'],
    ['technicianName', 'nombre del técnico'],
    ['workType', 'tipo de trabajo'],
    ['correction', 'corrección realizada'],
    ['receiverOperatorId', 'operador que recibe la máquina'],
    ['receiverOperatorName', 'nombre del operador que recibe la máquina']
  ];
  const missing = required
    .filter(([field]) => !String(body[field] || '').trim())
    .map(([, label]) => label);
  if (String(body.machineReleased || '') !== 'Sí') missing.push('confirmación de máquina liberada y probada');
  if (missing.length) throw new Error(`Faltan datos obligatorios para cerrar el ticket: ${missing.join(', ')}.`);

  const sheet = tickets_();
  const headers = headers_(sheet);
  const data = sheet.getDataRange().getDisplayValues();
  const dataIndex = data.findIndex((row, index) => index && row[headers.Folio] === body.folio);
  if (dataIndex < 0) throw new Error('No se encontró el ticket.');

  const row = [...data[dataIndex]];
  if (row[headers.Estatus] === 'Cerrado') throw new Error('Este ticket ya está cerrado.');

  const nowDate = new Date();
  const now = dateInfo_(nowDate);
  const receivedAt = body.receivedAt ? new Date(body.receivedAt) : nowDate;
  if (isNaN(receivedAt.getTime())) throw new Error('La hora de recibido no es válida.');
  const repairStart = body.repairStart ? new Date(body.repairStart) : receivedAt;
  if (isNaN(repairStart.getTime())) throw new Error('El inicio de reparación no es válido.');
  if (repairStart > nowDate) throw new Error('El inicio de reparación no puede ser posterior al cierre.');
  const repair = dateInfo_(repairStart);
  const received = dateInfo_(receivedAt);
  const openDate = parseDate_(row[headers['Fecha apertura']], row[headers['Hora apertura']]);
  const photo = savePhoto_(body.photoBase64, `${body.folio}-cierre.jpg`);

  set_(row, headers, 'Estatus', 'Cerrado'); set_(row, headers, 'Fecha/hora atención', received.dateTime); set_(row, headers, 'ID técnico asignado', body.technicianId); set_(row, headers, 'Técnico asignado', body.technicianName);
  set_(row, headers, 'Inicio de reparación', repair.dateTime); set_(row, headers, 'Tipo de trabajo', body.workType || 'Correctivo'); set_(row, headers, 'Causa raíz', body.rootCause || '');
  set_(row, headers, 'Corrección realizada', body.correction); set_(row, headers, 'Refacciones utilizadas', body.partsUsed || ''); set_(row, headers, 'Comentarios / recomendaciones', body.recommendations || '');
  set_(row, headers, 'Fecha/hora máquina lista / probada', now.dateTime); set_(row, headers, 'Fecha cierre', now.date); set_(row, headers, 'Hora cierre', now.time);
  set_(row, headers, 'Tiempo de respuesta (min)', mins_(openDate, receivedAt)); set_(row, headers, 'Tiempo de reparación (min)', mins_(repairStart, nowDate)); set_(row, headers, 'Tiempo detenido (min)', mins_(openDate, nowDate)); set_(row, headers, 'Foto cierre (Drive)', photo);
  set_(row, headers, '¿Máquina liberada / probada?', 'Sí'); set_(row, headers, 'ID operador recibe', body.receiverOperatorId); set_(row, headers, 'Operador recibe máquina', body.receiverOperatorName); set_(row, headers, 'Validación de liberación', body.releaseValidation || '');

  sheet.getRange(dataIndex + 1, 1, 1, sheet.getLastColumn()).setValues([row]);
  const link = row[headers['Liga de ticket']] || `${CONFIG.appUrl}?ticket=${encodeURIComponent(body.folio)}`;
  return {
    ok: true,
    folio: body.folio,
    ticketUrl: link,
    shareText: `✅ Ticket cerrado ${body.folio}\nÁrea: ${row[headers.Área]}\nMáquina: ${row[headers.Máquina]}\nCorrección: ${body.correction}\nMáquina liberada y probada: Sí\nCerrado: ${now.dateTime}`
  };
}
