/** Sistema de Tickets de Mantenimiento - Prodiesa
 * Pegar este archivo en un proyecto de Apps Script VINCULADO al Google Sheet.
 * Después desplegar como Web App: ejecutar como tú / acceso: cualquier persona.
 */
const CONFIG = {
  spreadsheetId: '12XIxmGGvFk1QE5qgfA2ULO1fYGW6ZNM5ucJmrFrMA7g',
  timeZone: 'America/Monterrey',
  sheets: { tickets: 'TICKETS', machines: 'Máquinas-LI', operators: 'Operadores-LI', technicians: 'Técnicos-LI', catalogs: 'Catálogos' },
  appUrl: 'https://TU-PROYECTO.vercel.app', // cambiar al publicar en Vercel
  driveFolderName: 'SITI - Tickets de Mantenimiento'
};

function doGet(e) {
  try {
    const action = (e.parameter.action || 'bootstrap').toLowerCase();
    if (action === 'bootstrap') return json_(bootstrap_(e.parameter.machineId || ''));
    if (action === 'ticket') return json_(getTicket_(e.parameter.folio || ''));
    return json_({ ok: false, message: 'Acción no reconocida.' });
  } catch (error) { return json_({ ok: false, message: error.message }); }
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    let result;
    if (body.action === 'open') result = openTicket_(body);
    else if (body.action === 'close') result = closeTicket_(body);
    else result = { ok: false, message: 'Acción no reconocida.' };
    return json_(result);
  } catch (error) { return json_({ ok: false, message: error.message }); }
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.openById(CONFIG.spreadsheetId); }
function sheet_(name) { const s = ss_().getSheetByName(name); if (!s) throw new Error(`No existe la hoja ${name}.`); return s; }
function headers_(sheet) {
  const values = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return values.reduce((map, name, i) => { map[name.trim()] = i; return map; }, {});
}
function records_(sheetName) {
  const sh = sheet_(sheetName), values = sh.getDataRange().getDisplayValues();
  const headers = values.shift().map(String);
  return values.filter(row => row.some(Boolean)).map(row => headers.reduce((o, h, i) => (o[h] = row[i], o), {}));
}
function now_() { return new Date(); }
function f_(date, pattern) { return Utilities.formatDate(date, CONFIG.timeZone, pattern); }
function monthName_(date) { return ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'][date.getMonth()]; }
function folder_() {
  const it = DriveApp.getFoldersByName(CONFIG.driveFolderName);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.driveFolderName);
}
function validatePerson_(role, id, pin) {
  const source = role === 'operator' ? CONFIG.sheets.operators : CONFIG.sheets.technicians;
  const label = role === 'operator' ? 'Nombre del operador' : 'Nombre del técnico';
  const person = records_(source).find(r => String(r.ID).trim() === String(id).trim() && String(r.PIN).trim() === String(pin).trim() && r.Activo === 'Sí');
  if (!person) throw new Error('ID o PIN no válido.');
  return { id: person.ID, name: person[label] };
}
function machine_(id) {
  const machine = records_(CONFIG.sheets.machines).find(r => String(r.ID).trim() === String(id).trim() && r.Activa === 'Sí');
  if (!machine) throw new Error('La máquina no está registrada o no está activa.');
  return machine;
}
function bootstrap_(machineId) {
  const machine = machineId ? machine_(machineId) : null;
  const catalog = records_(CONFIG.sheets.catalogs);
  return {
    ok: true, machine,
    machines: records_(CONFIG.sheets.machines).filter(r => r.Activa === 'Sí'),
    operators: records_(CONFIG.sheets.operators).filter(r => r.Activo === 'Sí').map(r => ({ id:r.ID, name:r['Nombre del operador'] })),
    technicians: records_(CONFIG.sheets.technicians).filter(r => r.Activo === 'Sí').map(r => ({ id:r.ID, name:r['Nombre del técnico'] })),
    catalogs: {
      priorities: catalog.map(r => r.Prioridad).filter(Boolean),
      failureTypes: catalog.map(r => r['Tipo de falla']).filter(Boolean),
      workTypes: catalog.map(r => r['Tipo de trabajo']).filter(Boolean),
      rootCauses: catalog.map(r => r['Causa raíz']).filter(Boolean)
    }
  };
}
function nextFolio_(year) {
  const rows = records_(CONFIG.sheets.tickets);
  const max = rows.filter(r => String(r.Año) === String(year)).reduce((m, r) => Math.max(m, Number(r.Consecutivo) || 0), 0);
  const consecutive = max + 1;
  return { consecutive, folio: `MTTO-${year}-${String(consecutive).padStart(5, '0')}` };
}
function storeImage_(base64, fileName) {
  if (!base64) return '';
  const clean = base64.replace(/^data:image\/[a-zA-Z+]+;base64,/, '');
  const blob = Utilities.newBlob(Utilities.base64Decode(clean), 'image/jpeg', fileName);
  const file = folder_().createFile(blob);
  return file.getUrl();
}
function card_(title, lines, fileName) {
  const safe = value => String(value || '').replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
  const text = lines.map((line, i) => `<text x="48" y="${128 + i * 50}" fill="#172033" font-family="Arial" font-size="${i === 0 ? 30 : 22}" font-weight="${i === 0 ? '700':'400'}">${safe(line)}</text>`).join('');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="700"><rect width="1200" height="700" fill="#f8fafc"/><rect width="1200" height="76" fill="#123f67"/><text x="48" y="49" fill="white" font-family="Arial" font-size="29" font-weight="700">PRODIESA · MANTENIMIENTO</text><text x="48" y="108" fill="#64748b" font-family="Arial" font-size="22">${safe(title)}</text>${text}</svg>`;
  const file = folder_().createFile(Utilities.newBlob(svg, 'image/svg+xml', fileName));
  return file.getUrl();
}
function writeRow_(sheet, values) {
  const headers = headers_(sheet);
  const row = Array(sheet.getLastColumn()).fill('');
  Object.keys(values).forEach(key => { if (headers[key] !== undefined) row[headers[key]] = values[key]; });
  sheet.appendRow(row);
  return sheet.getLastRow();
}
function openTicket_(body) {
  if (!body.machineId || !body.operatorId || !body.pin || !body.description) throw new Error('Faltan datos obligatorios para abrir el ticket.');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const machine = machine_(body.machineId), operator = validatePerson_('operator', body.operatorId, body.pin), date = now_();
    const year = f_(date, 'yyyy'), ticket = nextFolio_(year);
    const link = `${CONFIG.appUrl}?ticket=${encodeURIComponent(ticket.folio)}`;
    const photo = storeImage_(body.photoBase64, `${ticket.folio}-apertura.jpg`);
    const card = card_('TICKET ABIERTO', [ticket.folio, `${machine.Área} · ${machine['Nombre de máquina']}`, body.priority || 'Sin prioridad', body.description, f_(date, 'dd/MM/yyyy HH:mm')], `${ticket.folio}-apertura.svg`);
    writeRow_(sheet_(CONFIG.sheets.tickets), {
      'Folio': ticket.folio, 'Año': Number(year), 'Consecutivo': ticket.consecutive, 'Estatus': 'Abierto',
      'Fecha apertura': f_(date, 'dd/MM/yyyy'), 'Mes': monthName_(date), 'Hora apertura': f_(date, 'HH:mm'),
      'ID operador': operator.id, 'Operador': operator.name, 'ID máquina': machine.ID, 'Área': machine.Área, 'Máquina': machine['Nombre de máquina'],
      'Prioridad': body.priority || '', '¿Paro de producción?': body.productionStop || 'No', '¿Paro por calidad?': body.qualityStop || 'No',
      'Parte que falla': body.part || '', 'Tipo de falla': body.failureType || '', 'Descripción de falla': body.description, 'Foto apertura (Drive)': photo,
      'Liga de ticket': link
    });
    return { ok:true, folio:ticket.folio, cardUrl:card, ticketUrl:link, shareText:`🚨 Ticket abierto ${ticket.folio}\nMotivo: ${body.description}\nÁrea: ${machine.Área}\nMáquina: ${machine['Nombre de máquina']}\nFecha y hora: ${f_(date, 'dd/MM/yyyy HH:mm')}\nCerrar ticket: ${link}` };
  } finally { lock.releaseLock(); }
}
function parseDateTime_(dateText, timeText) {
  const parts = String(dateText || '').split('/').map(Number), time = String(timeText || '00:00').split(':').map(Number);
  return parts.length === 3 ? new Date(parts[2], parts[1] - 1, parts[0], time[0] || 0, time[1] || 0) : null;
}
function mins_(from, to) { return from && to ? Math.max(0, Math.round((to - from) / 60000)) : ''; }
function getTicket_(folio) {
  const sh = sheet_(CONFIG.sheets.tickets), headers = headers_(sh), rows = sh.getDataRange().getDisplayValues();
  const i = rows.findIndex((r, ix) => ix && r[headers.Folio] === folio);
  if (i < 0) return { ok:false, message:'No se encontró ese ticket.' };
  return { ok:true, ticket: rows[i].reduce((o, value, ix) => (o[Object.keys(headers).find(k => headers[k] === ix)] = value, o), {}) };
}
function closeTicket_(body) {
  if (!body.folio || !body.technicianId || !body.pin || !body.correction) throw new Error('Faltan datos obligatorios para cerrar el ticket.');
  const sh = sheet_(CONFIG.sheets.tickets), headers = headers_(sh), data = sh.getDataRange().getDisplayValues();
  const rowIndex = data.findIndex((r, i) => i && r[headers.Folio] === body.folio);
  if (rowIndex < 0) throw new Error('No se encontró el ticket.');
  const current = data[rowIndex], status = current[headers.Estatus];
  if (status === 'Cerrado') throw new Error('Este ticket ya está cerrado.');
  const tech = validatePerson_('technician', body.technicianId, body.pin), date = now_();
  const openDate = parseDateTime_(current[headers['Fecha apertura']], current[headers['Hora apertura']]);
  const repairStart = body.repairStart ? new Date(body.repairStart) : date;
  const photo = storeImage_(body.photoBase64, `${body.folio}-cierre.jpg`);
  const card = card_('TICKET CERRADO', [body.folio, `${current[headers.Área]} · ${current[headers.Máquina]}`, body.correction, `Cerrado: ${f_(date, 'dd/MM/yyyy HH:mm')}`], `${body.folio}-cierre.svg`);
  const updates = {
    'Estatus':'Cerrado', 'Fecha/hora atención': f_(repairStart, 'dd/MM/yyyy HH:mm'), 'ID técnico asignado': tech.id, 'Técnico asignado':tech.name,
    'Inicio de reparación': f_(repairStart, 'dd/MM/yyyy HH:mm'), 'Tipo de trabajo':body.workType || 'Correctivo',
    'Causa raíz':body.rootCause || '', 'Corrección realizada':body.correction, 'Refacciones utilizadas':body.partsUsed || '',
    'Comentarios / recomendaciones':body.recommendations || '', 'Fecha/hora máquina lista / probada':f_(date, 'dd/MM/yyyy HH:mm'),
    'Fecha cierre':f_(date, 'dd/MM/yyyy'), 'Hora cierre':f_(date, 'HH:mm'), 'Tiempo de respuesta (min)':mins_(openDate, repairStart),
    'Tiempo de reparación (min)':mins_(repairStart, date), 'Tiempo detenido (min)':mins_(openDate, date), 'Foto cierre (Drive)':photo
  };
  Object.keys(updates).forEach(key => { if (headers[key] !== undefined) sh.getRange(rowIndex + 1, headers[key] + 1).setValue(updates[key]); });
  const link = current[headers['Liga de ticket']];
  return { ok:true, folio:body.folio, cardUrl:card, ticketUrl:link, shareText:`✅ Ticket cerrado ${body.folio}\nÁrea: ${current[headers.Área]}\nMáquina: ${current[headers.Máquina]}\nCorrección: ${body.correction}\nCerrado: ${f_(date, 'dd/MM/yyyy HH:mm')}` };
}
