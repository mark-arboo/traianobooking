const SPREADSHEET_ID_2026 = '1eZ2t1dVZqAiZTflLigA9y8sHzbXSVdISgmlCt8MeOyk';
const SPREADSHEET_ID_2027 = '1ZXO7JxQZ6u9uge8QSMYn8sNmzZ7zW7b-KXsJ_64WZe0';
const spreadsheetIds = { 2026: SPREADSHEET_ID_2026, 2027: SPREADSHEET_ID_2027 };

function doGet(e) {
  return leggiPrenotazione(e);
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



function doPost(e) {
  return savePrenotazione(e);
}

function savePrenotazione(e) {
  let lock;
  try {
    const request = JSON.parse(e.postData.contents);
    const checkin = parseBookingDate(request.checkin);
    const checkout = parseBookingDate(request.checkout);
    if (checkout.serial <= checkin.serial) throw new Error('Il check-out deve essere successivo al check-in.');

    const spreadsheetId = spreadsheetIds[checkin.year];
    if (!spreadsheetId) throw new Error('Anno del check-in non supportato.');

    const channels = { 'booking.com': 'Booking.com', booking: 'Booking.com', airbnb: 'Airbnb', black: 'Black' };
    const channel = channels[String(request.channel).toLowerCase()];
    if (!channel) throw new Error('Canale non valido.');

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
      Cellulare: bookingSheetText(cellulare),
      'Check-in': checkin.serial,
      'Check-out': checkout.serial,
      'Numero Ospiti': bookingSheetText(numeroOspiti),
      Note: bookingSheetText(String(request.note || '').trim()),
      'Stato prenotazione': 'prenotata'
    };
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
    const headers = data[0].map(header => String(header).trim().toLowerCase());
    const entries = Object.keys(valuesByHeader).map(header => {
      const column = headers.indexOf(header.toLowerCase());
      if (column === -1) throw new Error('Colonna mancante nel foglio ' + channel + ': ' + header);
      return { header: header, column: column + 1, value: valuesByHeader[header] };
    });
    const nameColumn = headers.indexOf('nominativo');
    const emptyRow = data.findIndex((row, index) => index > 0 && String(row[nameColumn]).trim() === '');
    const rowNumber = emptyRow === -1 ? data.length + 1 : emptyRow + 1;

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


