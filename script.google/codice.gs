const SPREADSHEET_ID_BOOKINGS = '1ZXO7JxQZ6u9uge8QSMYn8sNmzZ7zW7b-KXsJ_64WZe0';
const BOOKINGS_SHEET_NAME = 'Bookings';
const EXPENSES_SHEET_NAME = 'Expenses';
const ALL_YEARS = 'tutti';
const SPREADSHEET_ID_UTENTI = '1dUkCiIWNuNRksmBd6IbVo616CzghI_pjOKMfqGI5KPA';
const TOKEN_TTL = 15; // Tempo espresso in giorni

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
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID_BOOKINGS);
    const sheet = spreadsheet.getSheetByName(EXPENSES_SHEET_NAME);
    if (!sheet) throw new Error('Foglio non trovato: ' + EXPENSES_SHEET_NAME);
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
      year: date.year,
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
  let result;
  try {
    const year = parseYearFilter(e && e.parameter && e.parameter.anno);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID_BOOKINGS);
    const sheet = spreadsheet.getSheetByName(EXPENSES_SHEET_NAME);
    if (!sheet) throw new Error('Foglio non trovato: ' + EXPENSES_SHEET_NAME);
    const rows = sheet.getDataRange().getValues();
    const headers = ['Data', 'Categoria', 'Descrizione', 'Importo', 'Note'];
    const data = [];

    rows.slice(1).forEach(row => {
      if (!row.slice(0, headers.length).some(value => value !== '' && value !== null && value !== undefined)) return;
      if (year !== null && sheetDateYear(row[0]) !== year) return;
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
    result = { code: error.badRequest ? 400 : 500, message: error.toString(), data: [] };
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

const BOOKING_CHANNELS = { 'booking.com': 'Booking.com', booking: 'Booking.com', airbnb: 'Airbnb', black: 'Black' };
const BOOKING_CHANNEL_KEYS = { 'Booking.com': 'booking', Airbnb: 'airbnb', Black: 'black' };
const BOOKING_HEADERS = ['ID', 'Nominativo', 'Cellulare', 'Check-in', 'Check-out', 'Numero Ospiti',
  'Ospiti soggetti a tassa di soggiorno', 'Guadagno Lordo', 'Costi di pulizia', 'Tassa di soggiorno', 'Guadagno Netto',
  'Stato prenotazione', 'Note', 'Channel'];

function modificaPrenotazione(e) {
  return writePrenotazione(e, true);
}

function cancellaPrenotazione(e) {
  let lock;
  try {
    const request = JSON.parse(e.postData.contents);
    const id = requiredBookingId(request.id);

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Salvataggio occupato. Riprova tra poco.');
    const bookings = openBookingsSheet();
    const rowIndex = findBookingRow(bookings, id);
    bookings.sheet.getRange(rowIndex + 1, bookings.columns['Stato prenotazione'] + 1).setValue('cancellata');
    SpreadsheetApp.flush();

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      id: id,
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
  let result;
  try {
    const year = parseYearFilter(e && e.parameter && e.parameter.anno);
    let bookings = openBookingsSheet();
    if (bookingRowsWithoutId(bookings).length) {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) throw new Error('Foglio occupato. Riprova tra poco.');
      try {
        bookings = openBookingsSheet();
        assignMissingBookingIds(bookings);
      } finally {
        lock.releaseLock();
      }
    }
    const columns = bookings.columns;
    const timeZone = 'Europe/Rome';
    const data = [];

    bookings.data.slice(1).forEach(row => {
      if (!String(row[columns.Nominativo] ?? '').trim()) return;
      if (String(row[columns['Stato prenotazione']] ?? '').trim().toLowerCase() === 'cancellata') return;
      if (year !== null && sheetDateYear(row[columns['Check-in']]) !== year) return;
      const channel = BOOKING_CHANNELS[String(row[columns.Channel] ?? '').trim().toLowerCase()];
      if (!channel) return;

      const entry = {};
      BOOKING_HEADERS.forEach(header => {
        const value = row[columns[header]];
        entry[header] = (header === 'Check-in' || header === 'Check-out') && value instanceof Date
          ? Utilities.formatDate(value, timeZone, 'dd/MM/yyyy') : value;
        if (header === 'Cellulare' && typeof value === 'string') {
          entry[header] = value.replace(/^\\+/, '+');
        }
      });
      entry.channel = BOOKING_CHANNEL_KEYS[channel];
      data.push(entry);
    });

    result = { code: 200, message: 'Lettura completata.', data: data };
  } catch (error) {
    result = { code: error.badRequest ? 400 : 500, message: error.toString(), data: [] };
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

    const id = updating ? requiredBookingId(request.id) : null;
    const channel = requiredBookingChannel(request.channel);
    const nominativo = requiredBookingText(request.nominativo, 'Nominativo');
    const cellulare = String(request.cellulare ?? '').trim();
    if (cellulare && !/^[0-9 +]+$/.test(cellulare)) throw new Error('Cellulare: sono ammessi solo numeri, spazi e +.');
    const numeroOspiti = requiredBookingText(request.numeroOspiti, 'Numero Ospiti');

    const valuesByHeader = {
      Nominativo: bookingSheetText(nominativo),
      Cellulare: bookingSheetText(cellulare.replace(/\+/g, '\\')),
      'Check-in': checkin.serial,
      'Check-out': checkout.serial,
      'Numero Ospiti': bookingSheetText(numeroOspiti),
      'Guadagno Lordo': parseBookingAmount(request.guadagnoLordo),
      'Costi di pulizia': parseBookingAmount(request.costiPulizia),
      Note: bookingSheetText(String(request.note || '').trim()),
      Channel: channel
    };
    if (!updating) valuesByHeader['Stato prenotazione'] = 'prenotata';
    if (channel === 'Black') {
      valuesByHeader['Ospiti soggetti a tassa di soggiorno'] = 0;
    } else {
      const taxGuestsText = requiredBookingText(request.ospitiSoggettiATassa, 'Ospiti soggetti a tassa di soggiorno');
      const taxGuests = Number(taxGuestsText);
      if (!/^\d+$/.test(taxGuestsText) || !Number.isSafeInteger(taxGuests)) {
        throw new Error('Gli ospiti soggetti a tassa devono essere un intero non negativo.');
      }
      valuesByHeader['Ospiti soggetti a tassa di soggiorno'] = taxGuests;
    }

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Salvataggio occupato. Riprova tra poco.');
    const bookings = openBookingsSheet();
    let rowIndex;
    let bookingId = id;
    if (updating) {
      rowIndex = findBookingRow(bookings, id);
    } else {
      assignMissingBookingIds(bookings);
      bookingId = nextBookingId(bookings);
      valuesByHeader.ID = bookingId;
      // Le righe con Nominativo vuoto contengono già le formule (Tassa di soggiorno, Guadagno Netto)
      rowIndex = bookings.data.findIndex((row, index) => index > 0 &&
        !String(row[bookings.columns.Nominativo] ?? '').trim());
    }
    const rowNumber = rowIndex === -1 ? bookings.data.length + 1 : rowIndex + 1;

    Object.keys(valuesByHeader).forEach(header => {
      const cell = bookings.sheet.getRange(rowNumber, bookings.columns[header] + 1);
      cell.setValue(valuesByHeader[header]);
      if (header === 'Check-in' || header === 'Check-out') cell.setNumberFormat('dd/MM/yyyy');
      if (header === 'Guadagno Lordo' || header === 'Costi di pulizia') cell.setNumberFormat('#,##0.00 "\u20ac"');
    });
    SpreadsheetApp.flush();

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      id: bookingId,
      year: checkin.year,
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

function openBookingsSheet() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID_BOOKINGS).getSheetByName(BOOKINGS_SHEET_NAME);
  if (!sheet) throw new Error('Foglio non trovato: ' + BOOKINGS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = (data[0] || []).map(header => String(header).trim().toLowerCase());
  const columns = {};
  BOOKING_HEADERS.forEach(header => {
    const column = headers.indexOf(header.toLowerCase());
    if (column === -1) throw new Error('Colonna mancante nel foglio ' + BOOKINGS_SHEET_NAME + ': ' + header);
    columns[header] = column;
  });
  return { sheet: sheet, data: data, columns: columns };
}

function findBookingRow(bookings, id) {
  const columns = bookings.columns;
  const matches = [];
  bookings.data.forEach((row, index) => {
    if (index > 0 && bookingIdValue(row[columns.ID]) === id &&
      String(row[columns['Stato prenotazione']] ?? '').trim().toLowerCase() !== 'cancellata') {
      matches.push(index);
    }
  });
  if (!matches.length) throw new Error('Prenotazione non trovata.');
  if (matches.length > 1) throw new Error('ID ' + id + ' duplicato nel foglio ' + BOOKINGS_SHEET_NAME + '.');
  return matches[0];
}

function bookingIdValue(value) {
  const text = String(value ?? '').trim();
  return /^\d+$/.test(text) ? Number(text) : null;
}

function nextBookingId(bookings) {
  return bookings.data.slice(1).reduce((max, row) =>
    Math.max(max, bookingIdValue(row[bookings.columns.ID]) || 0), 0) + 1;
}

function bookingRowsWithoutId(bookings) {
  const columns = bookings.columns;
  const rows = [];
  bookings.data.forEach((row, index) => {
    if (index > 0 && String(row[columns.Nominativo] ?? '').trim() && bookingIdValue(row[columns.ID]) === null) {
      rows.push(index);
    }
  });
  return rows;
}

// Va chiamata sotto lock: assegna un ID incrementale alle prenotazioni che ne sono prive
function assignMissingBookingIds(bookings) {
  const rows = bookingRowsWithoutId(bookings);
  if (!rows.length) return;
  let id = nextBookingId(bookings);
  rows.forEach(index => {
    bookings.sheet.getRange(index + 1, bookings.columns.ID + 1).setValue(id);
    bookings.data[index][bookings.columns.ID] = id;
    id++;
  });
  SpreadsheetApp.flush();
}

function requiredBookingId(value) {
  const id = bookingIdValue(value);
  if (!id) throw new Error('ID prenotazione mancante o non valido.');
  return id;
}

function sheetDateYear(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : Number(Utilities.formatDate(value, 'Europe/Rome', 'yyyy'));
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return new Date((value - 25569) * 86400000).getUTCFullYear();
  }
  const parts = /^\d{1,2}\/\d{1,2}\/(\d{4})$/.exec(String(value ?? '').trim());
  return parts ? Number(parts[1]) : null;
}

// Restituisce null quando anno = "tutti", cioè nessun filtro per anno
function parseYearFilter(value) {
  return String(value ?? '').trim().toLowerCase() === ALL_YEARS ? null : requiredBookingYear(value);
}

function requiredBookingYear(value) {
  const text = String(value ?? '').trim();
  if (!/^\d{4}$/.test(text)) {
    const error = new Error('Parametro anno mancante o non valido.');
    error.badRequest = true;
    throw error;
  }
  return Number(text);
}

function requiredBookingChannel(value) {
  const channel = BOOKING_CHANNELS[requiredBookingText(value, 'Canale').toLowerCase()];
  if (!channel) throw new Error('Canale non valido.');
  return channel;
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


