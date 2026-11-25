const SPREADSHEET_ID_2026 = '1eZ2t1dVZqAiZTflLigA9y8sHzbXSVdISgmlCt8MeOyk';
const SPREADSHEET_ID_2027 = '1ZXO7JxQZ6u9uge8QSMYn8sNmzZ7zW7b-KXsJ_64WZe0';
const SPREADSHEET_ID_UTENTI = '1dUkCiIWNuNRksmBd6IbVo616CzghI_pjOKMfqGI5KPA';
const TOKEN_TTL = 15; // Tempo espresso in giorni

const spreadsheetIds = { 2026: SPREADSHEET_ID_2026, 2027: SPREADSHEET_ID_2027 };

function doGet(e) {

  const action = e && e.parameter ? e.parameter.action : null;
  if (!isAuthorizedRequest(e)) return unauthorizedResponse();

  switch (action) {
    case 'leggi':
      return leggiPrenotazione(e);
    case 'leggiSpese':
      return leggiSpese(e);
    default:
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Azione GET non supportata"
      })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {

  const action = getActionFromRequest(e);
  if (action !== 'login' && !isAuthorizedRequest(e)) return unauthorizedResponse();

  switch (action) {
    case 'savePrenotazione':
      return savePrenotazione(e);
    case 'cancellaPrenotazione':
      return cancellaPrenotazione(e);
    case 'modificaPrenotazione':
      return modificaPrenotazione(e);
    case 'login':
      return login(e); 
    case 'saveSpesa':
      return saveSpesa(e);

    default:
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Azione POST non supportata"
      })).setMimeType(ContentService.MimeType.JSON);
  }
}

function isAuthorizedRequest(e) {
  const token = getTokenFromRequest(e);
  return Boolean(token) && checkTokenTTL(token);
}

function unauthorizedResponse() {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'error',
    code: 401,
    message: 'Token mancante o non valido.'
  })).setMimeType(ContentService.MimeType.JSON);
}

function getTokenFromRequest(e) {
  const queryToken = e && e.parameter ? e.parameter.token : null;
  if (queryToken) {
    return queryToken;
  }

  try {
    if (e && e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      return body && body.token ? body.token : null;
    }
  } catch (error) {
    // Ignora body non JSON
  }

  return null;
}

function getActionFromRequest(e) {
  const queryAction = e && e.parameter ? e.parameter.action : null;
  if (queryAction) {
    return queryAction;
  }

  try {
    if (e && e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      return body && body.action ? body.action : null;
    }
  } catch (error) {
    // Ignora body non JSON
  }

  return null;
}


function saveSpesa(e) {
  let lock;
  try {
    const request = JSON.parse(e.postData.contents);
    const year = requiredBookingText(request.anno, 'Anno');
    if (!Object.prototype.hasOwnProperty.call(spreadsheetIds, year)) throw new Error('Anno non supportato.');
    const date = parseBookingDate(request.data);
    const valuesByHeader = {
      Data: date.serial,
      Categoria: bookingSheetText(requiredBookingText(request.categoria, 'Categoria')),
      Descrizione: bookingSheetText(requiredBookingText(request.descrizione, 'Descrizione')),
      Importo: parseBookingAmount(request.importo),
      Note: bookingSheetText(String(request.note || '').trim())
    };

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Salvataggio occupato. Riprova tra poco.');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetIds[year]);
    const sheet = spreadsheet.getSheetByName('Expenses');
    if (!sheet) throw new Error('Foglio non trovato: Expenses');
    const data = sheet.getDataRange().getValues();
    const headers = (data[0] || []).map(header => String(header).trim().toLowerCase());
    const entries = Object.keys(valuesByHeader).map(header => {
      const column = headers.indexOf(header.toLowerCase());
      if (column === -1) throw new Error('Colonna mancante nel foglio Expenses: ' + header);
      return { header: header, column: column + 1, value: valuesByHeader[header] };
    });
    const rowNumber = data.length + 1;
    entries.forEach(entry => {
      const cell = sheet.getRange(rowNumber, entry.column);
      cell.setValue(entry.value);
      if (entry.header === 'Data') cell.setNumberFormat('dd/MM/yyyy');
      if (entry.header === 'Importo') cell.setNumberFormat('#,##0.00 "\u20ac"');
    });
    SpreadsheetApp.flush();

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      year: Number(year),
      row: rowNumber
    })).setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.TEXT);
  } finally {
    if (lock && lock.hasLock()) lock.releaseLock();
  }
}

function leggiSpese(e) {
  const year = String(e && e.parameter && e.parameter.anno || '').trim();
  let result;
  if (!Object.prototype.hasOwnProperty.call(spreadsheetIds, year)) {
    result = { code: 400, message: 'Parametro anno mancante o non supportato.', data: [] };
  } else {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetIds[year]);
      const sheet = spreadsheet.getSheetByName('Expenses');
      if (!sheet) throw new Error('Foglio non trovato: Expenses');
      const rows = sheet.getDataRange().getValues();
      const headers = ['Data', 'Categoria', 'Descrizione', 'Importo', 'Note'];
      const data = [];

      rows.slice(1).forEach(row => {
        if (!row.slice(0, headers.length).some(value => value !== '' && value !== null && value !== undefined)) return;
        const entry = {};
        headers.forEach((header, column) => {
          const value = row[column];
          entry[header] = header === 'Data' && value instanceof Date
            ? Utilities.formatDate(value, 'Europe/Rome', 'dd/MM/yyyy') : value;
        });
        data.push(entry);
      });

      result = { code: 200, message: 'Lettura completata.', data: data };
    } catch (error) {
      result = { code: 500, message: error.toString(), data: [] };
    }
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


function login(e) {
  let lock;
  try {
    let request;
    try {
      request = JSON.parse(e.postData.contents);
    } catch (error) {
      request = e && e.parameter ? e.parameter : {};
    }
    const username = requiredBookingText(request.username, 'Username');
    const password = requiredBookingText(request.password, 'Password');

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Accesso occupato. Riprova tra poco.');
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID_UTENTI).getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex((row, index) => index > 0 &&
      String(row[0] ?? '').trim() === username && String(row[1] ?? '') === password);
    if (rowIndex === -1) throw new Error('Utente non esistente o password errata.');

    const now = new Date();
    const token = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, username + now.getTime() + Utilities.getUuid())
      .map(byte => ((byte & 0xff) + 0x100).toString(16).slice(1))
      .join('');
    const dateText = Utilities.formatDate(now, 'Europe/Rome', 'dd/MM/yyyy HH:mm:ss');
    sheet.getRange(rowIndex + 1, 3).setValue(token);
    sheet.getRange(rowIndex + 1, 4).setValue(dateText);
    SpreadsheetApp.flush();

    return ContentService.createTextOutput(JSON.stringify({
      code: 200,
      message: 'Login completato.',
      token: token
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      code: 500,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (lock && lock.hasLock()) lock.releaseLock();
  }
}

function checkTokenTTL(token) {
  const tokenText = String(token ?? '').trim();
  if (!tokenText) return false;

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID_UTENTI).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const row = data.find((candidate, index) => index > 0 && String(candidate[2] ?? '').trim() === tokenText);
  if (!row) return false;

  const tokenDate = parseTokenDate(row[3]);
  if (!tokenDate) return false;
  return Date.now() - tokenDate.getTime() <= TOKEN_TTL * 24 * 60 * 60 * 1000;
}

function parseTokenDate(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  const parts = /^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/.exec(String(value ?? '').trim());
  if (!parts) return null;
  return new Date(Number(parts[3]), Number(parts[2]) - 1, Number(parts[1]),
    Number(parts[4]), Number(parts[5]), Number(parts[6]));
}

function modificaPrenotazione(e) {
  return writePrenotazione(e, true);
}

function cancellaPrenotazione(e) {
  let lock;
  try {
    const request = JSON.parse(e.postData.contents);
    const year = requiredBookingText(request.anno, 'Anno');
    if (!Object.prototype.hasOwnProperty.call(spreadsheetIds, year)) {
      throw new Error('Anno non supportato.');
    }
    const channels = { 'booking.com': 'Booking.com', booking: 'Booking.com', airbnb: 'Airbnb', black: 'Black' };
    const channelKey = requiredBookingText(request.channel, 'Canale').toLowerCase();
    if (!Object.prototype.hasOwnProperty.call(channels, channelKey)) throw new Error('Canale non valido.');
    const channel = channels[channelKey];
    const nominativo = requiredBookingText(request.nominativo, 'Nominativo');

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Salvataggio occupato. Riprova tra poco.');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetIds[year]);
    const sheet = spreadsheet.getSheets().find(candidate => candidate.getName().toLowerCase() === channel.toLowerCase());
    if (!sheet) throw new Error('Foglio non trovato: ' + channel);
    const data = sheet.getDataRange().getValues();
    const headers = (data[0] || []).map(header => String(header).trim().toLowerCase());
    const nameColumn = headers.indexOf('nominativo');
    const statusColumn = headers.indexOf('stato prenotazione');
    if (nameColumn === -1 || statusColumn === -1) {
      throw new Error('Colonna Nominativo o Stato prenotazione mancante nel foglio ' + channel);
    }
    const rowIndex = data.findIndex((row, index) => index > 0 && String(row[nameColumn] ?? '').trim() === nominativo);
    if (rowIndex === -1) throw new Error('Prenotazione non trovata.');
    sheet.getRange(rowIndex + 1, statusColumn + 1).setValue('cancellata');
    SpreadsheetApp.flush();

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      year: Number(year),
      channel: channel,
      row: rowIndex + 1
    })).setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.TEXT);
  } finally {
    if (lock && lock.hasLock()) lock.releaseLock();
  }
}


function leggiPrenotazione(e) {
  
  const year = String(e && e.parameter && e.parameter.anno || '').trim();
  let result;
  if (!Object.prototype.hasOwnProperty.call(spreadsheetIds, year)) {
    result = { code: 400, message: 'Parametro anno mancante o non supportato.', data: [] };
  } else {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetIds[year]);
      const sheets = spreadsheet.getSheets();
      const timeZone = 'Europe/Rome';
      const data = [];
      const channels = ['booking', 'airbnb', 'black'];

      [1, 2, 3].forEach(index => {
        const sheet = sheets[index];
        if (!sheet) throw new Error('Sheet con indice ' + index + ' non trovato.');
        const rows = sheet.getDataRange().getValues();
        if (!rows.length) return;
        const headers = rows[0].map(header => String(header).trim());
        const nameColumn = headers.indexOf('Nominativo');
        const statusColumn = headers.indexOf('Stato prenotazione');

        rows.slice(1).forEach(row => {
          if (!String(row[nameColumn] ?? '').trim()) return;
          if (String(row[statusColumn] ?? '').trim().toLowerCase() === 'cancellata') return;
          const entry = {};
          let hasData = false;
          headers.forEach((header, column) => {
            if (!header) return;
            const value = row[column];
            entry[header] = (header === 'Check-in' || header === 'Check-out') && value instanceof Date
              ? Utilities.formatDate(value, timeZone, 'dd/MM/yyyy') : value;
            if (header === 'Cellulare' && typeof value === 'string') {
              entry[header] = value.replace(/^\\+/, '+');
            }
            if (value !== '' && value !== null && value !== undefined) hasData = true;
          });
          entry.channel = channels[index - 1];
          if (hasData) data.push(entry);
        });
      });

      result = { code: 200, message: 'Lettura completata.', data: data };
    } catch (error) {
      result = { code: 500, message: error.toString(), data: [] };
    }
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}





function savePrenotazione(e) {
  return writePrenotazione(e, false);
}

function writePrenotazione(e, updating) {
  let lock;
  try {
    const request = JSON.parse(e.postData.contents);
    const checkin = parseBookingDate(request.checkin);
    const checkout = parseBookingDate(request.checkout);
    if (checkout.serial <= checkin.serial) throw new Error('Il check-out deve essere successivo al check-in.');

    const year = updating ? requiredBookingText(request.anno, 'Anno') : String(checkin.year);
    if (!Object.prototype.hasOwnProperty.call(spreadsheetIds, year)) throw new Error('Anno non supportato.');
    const spreadsheetId = spreadsheetIds[year];

    const channels = { 'booking.com': 'Booking.com', booking: 'Booking.com', airbnb: 'Airbnb', black: 'Black' };
    const channelKey = requiredBookingText(request.channel, 'Canale').toLowerCase();
    if (!Object.prototype.hasOwnProperty.call(channels, channelKey)) throw new Error('Canale non valido.');
    const channel = channels[channelKey];

    const nominativo = requiredBookingText(request.nominativo, 'Nominativo');
    const cellulare = String(request.cellulare ?? '').trim();
    if (cellulare && !/^[0-9 +]+$/.test(cellulare)) throw new Error('Cellulare: sono ammessi solo numeri, spazi e +.');
    const numeroOspiti = requiredBookingText(request.numeroOspiti, 'Numero Ospiti');
    let taxGuests = 0;
    if (channel !== 'Black') {
      const taxGuestsText = requiredBookingText(request.ospitiSoggettiATassa, 'Ospiti soggetti a tassa di soggiorno');
      taxGuests = Number(taxGuestsText);
      if (!/^\d+$/.test(taxGuestsText) || !Number.isSafeInteger(taxGuests)) {
        throw new Error('Gli ospiti soggetti a tassa devono essere un intero non negativo.');
      }
    }

    const valuesByHeader = {
      Nominativo: bookingSheetText(nominativo),
      Cellulare: bookingSheetText(cellulare.replace(/\+/g, '\\')),
      'Check-in': checkin.serial,
      'Check-out': checkout.serial,
      'Numero Ospiti': bookingSheetText(numeroOspiti),
      Note: bookingSheetText(String(request.note || '').trim())
    };
    if (!updating) valuesByHeader['Stato prenotazione'] = 'prenotata';
    if (channel === 'Airbnb') valuesByHeader['Costi di pulizia'] = parseBookingAmount(request.costiPulizia);
    if (channel === 'Airbnb' || channel === 'Booking.com' ) {
      valuesByHeader['Guadagno Lordo'] = parseBookingAmount(request.guadagnoLordo);
      valuesByHeader['Ospiti soggetti a tassa di soggiorno'] = taxGuests;
    }
    if (channel === 'Black') valuesByHeader['Guadagno Netto'] = parseBookingAmount(request.guadagnoLordo);

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Salvataggio occupato. Riprova tra poco.');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheets().find(candidate => candidate.getName().toLowerCase() === channel.toLowerCase());
    if (!sheet) throw new Error('Foglio non trovato: ' + channel);
    const data = sheet.getDataRange().getValues();
    const headers = (data[0] || []).map(header => String(header).trim().toLowerCase());
    const entries = Object.keys(valuesByHeader).map(header => {
      const column = headers.indexOf(header.toLowerCase());
      if (column === -1) throw new Error('Colonna mancante nel foglio ' + channel + ': ' + header);
      return { header: header, column: column + 1, value: valuesByHeader[header] };
    });
    const nameColumn = headers.indexOf('nominativo');
    const rowIndex = data.findIndex((row, index) => index > 0 &&
      String(row[nameColumn] ?? '').trim() === (updating ? nominativo : ''));
    if (updating && rowIndex === -1) throw new Error('Prenotazione non trovata.');
    const rowNumber = rowIndex === -1 ? data.length + 1 : rowIndex + 1;

    entries.forEach(entry => {
      const cell = sheet.getRange(rowNumber, entry.column);
      cell.setValue(entry.value);
      if (entry.header === 'Check-in' || entry.header === 'Check-out') cell.setNumberFormat('dd/MM/yyyy');
      if (entry.header === 'Guadagno Lordo' || entry.header === 'Guadagno Netto') cell.setNumberFormat('#,##0.00 "\u20ac"');
      if (entry.header === 'Costi di pulizia') cell.setNumberFormat('#,##0.00 "\u20ac"');
    });
    const nightsColumn = headers.indexOf('notti');
    if (nightsColumn !== -1) {
      const nightsCell = sheet.getRange(rowNumber, nightsColumn + 1);
      if (!nightsCell.getFormula()) nightsCell.setValue(checkout.serial - checkin.serial);
    }
    SpreadsheetApp.flush();

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      year: Number(year),
      channel: channel,
      row: rowNumber
    })).setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.TEXT);
  } finally {
    if (lock && lock.hasLock()) lock.releaseLock();
  }
}

function requiredBookingText(value, label) {
  if (value === null || value === undefined || !String(value).trim()) {
    throw new Error('Campo obbligatorio: ' + label);
  }
  return String(value).trim();
}

function parseBookingDate(value) {
  const text = requiredBookingText(value, 'Data');
  const parts = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(text);
  if (!parts) throw new Error('Data non valida: usa gg/mm/aaaa.');
  const day = Number(parts[1]);
  const month = Number(parts[2]);
  const year = Number(parts[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error('Data non valida.');
  }
  return { year: year, serial: date.getTime() / 86400000 + 25569 };
}

function parseBookingAmount(value) {
  const text = requiredBookingText(value, 'Importo');
  const amount = Number(text.replace(',', '.'));
  if (!/^\d+([.,]\d{1,2})?$/.test(text) || !Number.isFinite(amount)) {
    throw new Error('Importo non valido: usa un valore non negativo con massimo due decimali.');
  }
  return amount;
}

function bookingSheetText(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}


