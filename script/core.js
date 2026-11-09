// URL applicazione Script google
const bookingScriptURL = 'https://script.google.com/macros/s/AKfycbwLmckTZoLxkMf6cys52o2CbvI6FRHpXdC0TZQuxQ_xYjhdr4qurelNIFraEf3XdNKKQg/exec';
// URL del foglio Google Sheets esportato in formato XLSX
let googleSheetURL_2026 = "https://docs.google.com/spreadsheets/d/1eZ2t1dVZqAiZTflLigA9y8sHzbXSVdISgmlCt8MeOyk/export?format=xlsx";
let googleSheetURL_2027 = "https://docs.google.com/spreadsheets/d/1ZXO7JxQZ6u9uge8QSMYn8sNmzZ7zW7b-KXsJ_64WZe0/export?format=xlsx";
const spreadsheetURL = { 2026: googleSheetURL_2026, 2027: googleSheetURL_2027 };

let arrayBooking = [];
let selectedBookingYear = 2027;
let currentMonth = new Date().getMonth();
let currentYear = new Date().getFullYear();
let cache = false;

/**
 * Controlla se è la prima volta che l'applicazione viene caricata
 * o se è un refresh di pagina
 * @returns {boolean} true se è la prima volta, false se è un refresh
 */
function isFirstLoad() {
    // Verifica usando Navigation API moderna
    const navigationEntries = performance.getEntriesByType('navigation');
    let isRefresh = false;
    
    if (navigationEntries.length > 0) {
        const navEntry = navigationEntries[0];
        isRefresh = navEntry.type === 'reload';
    } else {
        // Fallback per browser più vecchi
        isRefresh = performance.navigation && performance.navigation.type === performance.navigation.TYPE_RELOAD;
    }
    
    // Verifica anche il sessionStorage per distinguere tra nuova sessione e refresh
    const hasSessionData = sessionStorage.getItem('appLoaded');
    
    if (!hasSessionData && !isRefresh) {
        // Prima volta in questa sessione
        sessionStorage.setItem('appLoaded', 'true');
        return true;
    }
    
    return false;
}




/**
 * Logica eseguita al primo caricamento
 */
function onFirstLoad() {

    if (localStorage.getItem('arrayBooking') && cache &&
        localStorage.getItem('arrayBookingYear') === String(selectedBookingYear)) {
        console.log('onFirstLoad - Caricamento dati da localStorage');

        arrayBooking = JSON.parse(localStorage.getItem('arrayBooking'));
        sessionStorage.setItem('arrayBooking', localStorage.getItem('arrayBooking'));
        sessionStorage.setItem('arrayBookingYear', String(selectedBookingYear));
    
        showListView(); // Mostra la vista lista all'avvio dell'app
        
    } else {
        loadFromRemote(googleSheetURL_2027);
    }
}

function changeBookingYear(year) {

    const sheetURL = spreadsheetURL[year];
    return loadFromRemote(sheetURL);
}

function loadFromRemote(URL = selectedBookingYear === 2026 ? googleSheetURL_2026 : googleSheetURL_2027) {
    const bookingYear = URL === googleSheetURL_2026 ? 2026 : 2027;
    const yearSelect = document.getElementById('bookingYearSelect');
    yearSelect.disabled = true;
    const listView = document.getElementById('listView');
    listView.setAttribute('aria-busy', 'true');
    listView.innerHTML = '<img src="./img/loading.gif" alt="Caricamento..." id="loadingIndicator" width="50" height="50">';
    console.log('loadFromRemote - Caricamento dati da URL');
    // legge i primi 3 sheet di un file excel letto da una URL 
    return fetch(URL)
        .then(response => {
            if (!response.ok) throw new Error('Caricamento non riuscito (HTTP ' + response.status + ')');
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const bookingSheetName = workbook.SheetNames[1];
            const airbnbSheetName = workbook.SheetNames[2];
            const blackSheetName = workbook.SheetNames[3];
            const bookingWorksheet = workbook.Sheets[bookingSheetName];
            const airbnbWorksheet = workbook.Sheets[airbnbSheetName];
            const blackWorksheet = workbook.Sheets[blackSheetName];
            const bookingData = XLSX.utils.sheet_to_json(bookingWorksheet).map(booking => ({ ...booking, channel: 'booking' }));
            const airbnbData = XLSX.utils.sheet_to_json(airbnbWorksheet).map(booking => ({ ...booking, channel: 'airbnb' }));
            const blackData = XLSX.utils.sheet_to_json(blackWorksheet).map(booking => ({ ...booking, channel: 'black' }));

            if (arrayBooking.length > 0) {
                // Svuoto l'arrayBooking prima di concatenare i nuovi dati
                arrayBooking = [];
            }
            arrayBooking = arrayBooking.concat(bookingData);
            arrayBooking = arrayBooking.concat(airbnbData);
            arrayBooking = arrayBooking.concat(blackData);

            // Dell'arrayBooking voglio solo le prenotazioni che hanno il campo nominativo valorizzato e il campo "Stato prenotazione", se esiste, diversa da "cancellata"
            arrayBooking = arrayBooking.filter(booking => booking.Nominativo && (!booking.hasOwnProperty('Stato prenotazione') || booking['Stato prenotazione'].toLowerCase() != 'cancellata'));

            // Converte i numeri seriali Excel in date leggibili
            arrayBooking.forEach(booking => {
                if (booking['Check-in']) {
                    booking['Check-in'] = excelDateToJSDate(booking['Check-in']);
                }
                if (booking['Check-out']) {
                    booking['Check-out'] = excelDateToJSDate(booking['Check-out']);
                }
            });

            // ordina l'arrayBooking per data di arrivo (campo Check-in)
            arrayBooking.sort((a, b) => parseItalianDate(a['Check-in']) - parseItalianDate(b['Check-in']));

            selectedBookingYear = bookingYear;
            sessionStorage.setItem('arrayBooking', JSON.stringify(arrayBooking));
            sessionStorage.setItem('arrayBookingYear', String(selectedBookingYear));

            if (cache) {
                localStorage.setItem('arrayBooking', JSON.stringify(arrayBooking));
                localStorage.setItem('arrayBookingYear', String(selectedBookingYear));
            }

            showListView(); // Mostra la vista lista all'avvio dell'app
        })
        .catch(error => {
            yearSelect.value = String(selectedBookingYear);
            showListView();
            console.error('Errore durante il caricamento del file Excel:', error);
        })
        .finally(() => {
            listView.setAttribute('aria-busy', 'false');
            yearSelect.disabled = false;
        });
}

/**
 * Logica eseguita durante il refresh
 */
function onPageRefresh() {
    const storedYear = sessionStorage.getItem('arrayBookingYear');
    if (sessionStorage.getItem('arrayBooking') && (storedYear === '2026' || storedYear === '2027')) {
        arrayBooking = JSON.parse(sessionStorage.getItem('arrayBooking'));
        selectedBookingYear = Number(storedYear);
        showListView();
    } else {
        onFirstLoad();
    }
}




function initializeApp() {

    if (isFirstLoad()) {
        onFirstLoad();
    } else {
        onPageRefresh();
    }

}


function showGraphicView() {

    sessionStorage.setItem('lastActivePanel', 'grafico');

    // Rimuove la classe active da tutti i pulsanti
    document.getElementById('listViewBtn').classList.remove('active');
    document.getElementById('calendarViewBtn').classList.remove('active');
    // Aggiunge la classe active al pulsante Grafico
    document.getElementById('graphicViewBtn').classList.add('active');

    // Aggiorna il titolo
    updateViewHeader('Report', 'grafico');

    renderGraphicView();
  
    // Mostra la vista grafico e nasconde le altre viste
    document.getElementById('graphicView').classList.add('active');
    document.getElementById('listView').classList.remove('active');
    document.getElementById('calendarView').classList.remove('active');

}

function renderGraphicView() {
    const graphicContainer = document.getElementById('graphicView');
    graphicContainer.innerHTML = ''; // Pulisce il contenitore
    
    // Crea l'header con navigazione
    const header = document.createElement('div');
    header.className = 'calendar-header';
    const monthNames = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                        'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    header.innerHTML = `
        <button class="nav-button" onclick="previousMonthGraphic()">◀</button>
        <h2>${monthNames[currentMonth]} ${currentYear}</h2>
        <button class="nav-button" onclick="nextMonthGraphic()">▶</button>
    `;
    graphicContainer.appendChild(header);

    // Contenitore per le statistiche
    const statsContainer = document.createElement('div');
    statsContainer.className = 'stats-container';

    // Filtra le prenotazioni per il mese corrente in base alla data di check-in
    const monthStart = new Date(currentYear, currentMonth, 1);
    const monthEnd = new Date(currentYear, currentMonth + 1, 0);
    monthEnd.setHours(23, 59, 59, 999);

    const bookingsInMonth = arrayBooking.filter(booking => {
        const checkin = parseItalianDate(booking['Check-in']);
        return checkin >= monthStart && checkin <= monthEnd;
    });

    // Calcola le somme per channel sull'intero costo della prenotazione
    const revenueByChannel = {
        booking: 0,
        airbnb: 0,
        black: 0
    };

    bookingsInMonth.forEach(booking => {
        const guadagnoNetto = booking['Guadagno Netto'] || 0;
        const channel = booking.channel || 'unknown';
        
        if (revenueByChannel.hasOwnProperty(channel)) {
            revenueByChannel[channel] += guadagnoNetto;
        }
    });

    // Calcola il totale mensile
    const totalRevenue = revenueByChannel.booking + revenueByChannel.airbnb + revenueByChannel.black;

    // Calcola la revenue annuale
    const yearlyRevenue = calculateYearlyRevenue(currentYear);

    // Crea la visualizzazione con due card separate
    statsContainer.innerHTML = `
        <div class="stats-card">
            <h3>Revenue per Canale</h3>
            <div class="revenue-item">
                <img src="./img/booking.png" alt="Booking" class="channel-logo">
                <div class="channel-info">
                    <p class="channel-name">Booking.com</p>
                    <p class="channel-revenue">${formatEuro(revenueByChannel.booking)} €</p>
                </div>
            </div>
            <div class="revenue-item">
                <img src="./img/airbnb.png" alt="Airbnb" class="channel-logo">
                <div class="channel-info">
                    <p class="channel-name">Airbnb</p>
                    <p class="channel-revenue">${formatEuro(revenueByChannel.airbnb)} €</p>
                </div>
            </div>
            <div class="revenue-item">
                <img src="./img/black.png" alt="Black" class="channel-logo">
                <div class="channel-info">
                    <p class="channel-name">Altro</p>
                    <p class="channel-revenue">${formatEuro(revenueByChannel.black)} €</p>
                </div>
            </div>
            <div class="revenue-total">
                <p class="total-label">Totale Mensile:</p>
                <p class="total-revenue">${formatEuro(totalRevenue)} €</p>
            </div>
            <div class="booking-count">
                <p>Numero prenotazioni: <strong>${bookingsInMonth.length}</strong></p>
            </div>
        </div>
        
        <div class="stats-card yearly-card">
            <div class="yearly-revenue">
                <i class="fa fa-line-chart" aria-hidden="true"></i>
                <div class="yearly-info">
                    <p class="yearly-label">Revenue Annuale ${currentYear}</p>
                    <p class="yearly-amount">${formatEuro(yearlyRevenue)} €</p>
                </div>
            </div>
        </div>
    `;

    graphicContainer.appendChild(statsContainer);
}

function previousMonthGraphic() {
    currentMonth--;
    if (currentMonth < 0) {
        currentMonth = 11;
        currentYear--;
    }
    renderGraphicView();
}

function nextMonthGraphic() {
    currentMonth++;
    if (currentMonth > 11) {
        currentMonth = 0;
        currentYear++;
    }
    renderGraphicView();
}

function calculateYearlyRevenue(year) {
    let totalYearRevenue = 0;
    
    // Itera su tutti i 12 mesi dell'anno
    for (let month = 0; month < 12; month++) {
        const monthStart = new Date(year, month, 1);
        const monthEnd = new Date(year, month + 1, 0);
        monthEnd.setHours(23, 59, 59, 999);
        
        // Filtra le prenotazioni per il mese in base alla data di check-in
        const bookingsInMonth = arrayBooking.filter(booking => {
            const checkin = parseItalianDate(booking['Check-in']);
            return checkin >= monthStart && checkin <= monthEnd;
        });
        
        // Somma l'intero costo di ogni prenotazione con check-in nel mese
        bookingsInMonth.forEach(booking => {
            totalYearRevenue += booking['Guadagno Netto'] || 0;
        });
    }
    
    return totalYearRevenue;
}


function updateViewHeader(title, panel) {
    document.querySelector('h1').textContent = title;
    const yearSelect = document.getElementById('bookingYearSelect');
    yearSelect.hidden = panel !== 'lista';
    yearSelect.value = String(selectedBookingYear);
    document.getElementById('insertBookingBtn').hidden = panel !== 'lista';
    document.getElementById('backToListBtn').hidden = panel !== 'inserisci';
    document.getElementById('insertBookingView').classList.toggle('active', panel === 'inserisci');
}

function showInsertBooking() {
    sessionStorage.setItem('lastActivePanel', 'inserisci');
    document.querySelectorAll('.view, .tabButton').forEach(element => {
        element.classList.remove('active');
    });
    document.getElementById('listViewBtn').classList.add('active');
    updateViewHeader('Nuova Prenotazione', 'inserisci');

    const container = document.getElementById('insertBookingView');
    container.innerHTML = `
        <form id="bookingForm" class="booking-form">
            <div class="booking-field">
                <label for="bookingChannel">Canale</label>
                <select id="bookingChannel" name="channel" required>
                    <option value="Booking.com">Booking.com</option>
                    <option value="Airbnb">Airbnb</option>
                    <option value="Black">Black</option>
                </select>
            </div>
            <div class="booking-field">
                <label for="bookingName">Nominativo</label>
                <input id="bookingName" name="Nominativo" type="text" autocomplete="name" required>
            </div>
            <div class="booking-field booking-field-wide">
                <label for="bookingPhone">Cellulare</label>
                <input id="bookingPhone" name="Cellulare" type="tel" autocomplete="tel" title="Sono ammessi solo numeri, spazi e il carattere +" required>
            </div>
            <div class="booking-field">
                <label for="bookingCheckin">Check-in (gg/mm/aaaa)</label>
                <input id="bookingCheckin" name="Check-in" type="text" inputmode="numeric" placeholder="gg/mm/aaaa" pattern="[0-9]{2}/[0-9]{2}/[0-9]{4}" maxlength="10" required>
            </div>
            <div class="booking-field">
                <label for="bookingCheckout">Check-out (gg/mm/aaaa)</label>
                <input id="bookingCheckout" name="Check-out" type="text" inputmode="numeric" placeholder="gg/mm/aaaa" pattern="[0-9]{2}/[0-9]{2}/[0-9]{4}" maxlength="10" required>
            </div>
            <div class="booking-field">
                <label for="bookingGuests">Numero ospiti</label>
                <input id="bookingGuests" name="Numero Ospiti" type="text" required>
            </div>
            <div class="booking-field">
                <label for="bookingTaxGuests">Ospiti soggetti a tassa di soggiorno</label>
                <input id="bookingTaxGuests" name="Ospiti soggetti a tassa di soggiorno" type="number" min="0" step="1" required>
            </div>
            <div class="booking-field booking-field-wide">
                <label for="bookingGross">Lordo (euro)</label>
                <input id="bookingGross" name="Guadagno Lordo" type="number" inputmode="decimal" min="0" step="0.01" required>
            </div>
            <div id="bookingCleaningField" class="booking-field booking-field-wide" hidden>
                <label for="bookingCleaning">Costi di pulizia (euro)</label>
                <input id="bookingCleaning" name="Costi di pulizia" type="number" inputmode="decimal" min="0" step="0.01" disabled>
            </div>
            <div class="booking-field booking-field-wide">
                <label for="bookingNotes">Note</label>
                <textarea id="bookingNotes" name="Note" rows="4"></textarea>
            </div>
            <button type="submit" class="booking-submit booking-field-wide">
                Inserisci prenotazione
            </button>
            <p id="bookingSaveStatus" class="booking-field-wide" role="status" hidden></p>
        </form>
    `;

    const form = document.getElementById('bookingForm');
    updateBookingCleaningField(form);
    form.elements.namedItem('channel').addEventListener('change', () => updateBookingCleaningField(form));
    form.addEventListener('submit', submitBooking);
    form.addEventListener('focusout', event => {
        if (!event.target.matches('input, select, textarea')) return;
        event.target.dataset.touched = 'true';
        validateBookingForm(form);
        if (!event.target.validity.valid) event.target.reportValidity();
    });
    form.addEventListener('input', event => {
        if (!event.target.matches('input, select, textarea')) return;
        event.target.setCustomValidity('');
        event.target.removeAttribute('aria-invalid');
    });
    form.addEventListener('invalid', event => {
        event.target.dataset.touched = 'true';
        updateBookingFieldError(event.target);
    }, true);
}

function updateBookingCleaningField(form) {
    const isAirbnb = form.elements.namedItem('channel').value === 'Airbnb';
    const field = form.elements.namedItem('Costi di pulizia');
    document.getElementById('bookingCleaningField').hidden = !isAirbnb;
    field.disabled = !isAirbnb;
    field.required = isAirbnb;
    if (!isAirbnb) {
        field.setCustomValidity('');
        field.removeAttribute('aria-invalid');
        delete field.dataset.touched;
    }
}

function updateBookingFieldError(field) {
    const invalid = !field.validity.valid;
    field.setAttribute('aria-invalid', String(invalid));
    if (!invalid) return;

    const formatHints = {
        channel: 'Scegli un canale tra Booking.com, Airbnb e Black.',
        Nominativo: 'Inserisci il nominativo.',
        Cellulare: 'Usa solo numeri, spazi e il carattere +, ad esempio +39 333 1234567.',
        'Check-in': 'Inserisci una data valida in formato gg/mm/aaaa, ad esempio 09/11/2026.',
        'Check-out': 'Inserisci una data valida in formato gg/mm/aaaa, successiva al check-in.',
        'Numero Ospiti': 'Inserisci il numero ospiti, anche come testo: ad esempio 2 adulti e 1 bambino.',
        'Ospiti soggetti a tassa di soggiorno': 'Inserisci un numero intero maggiore o uguale a 0.',
        'Guadagno Lordo': 'Inserisci un importo maggiore o uguale a 0 con massimo due decimali, ad esempio 250,75.',
        'Costi di pulizia': 'Inserisci un importo maggiore o uguale a 0 con massimo due decimali, ad esempio 250,75.'
    };

    const message = field.validity.valueMissing
        ? 'Campo obbligatorio: inserisci un valore.'
        : field.validity.customError
            ? field.validationMessage : formatHints[field.name] || field.validationMessage;
    field.setCustomValidity(message);
}

function validateBookingForm(form) {
    const fields = form.elements;
    const data = new FormData(form);

    form.querySelectorAll('input, select, textarea').forEach(field => {
        field.setCustomValidity('');
    });

    ['Check-in', 'Check-out'].forEach(name => {
        const value = data.get(name);
        const date = parseItalianDate(value);
        const formattedDate = `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
        fields.namedItem(name).setCustomValidity(
            /^[0-9]{2}\/[0-9]{2}\/[0-9]{4}$/.test(value) && formattedDate === value
                ? '' : 'Inserisci una data valida in formato gg/mm/aaaa.'
        );
    });

    if (fields.namedItem('Check-in').validity.valid && fields.namedItem('Check-out').validity.valid &&
        parseItalianDate(data.get('Check-out')) <= parseItalianDate(data.get('Check-in'))) {
        fields.namedItem('Check-out').setCustomValidity('Il check-out deve essere successivo al check-in.');
    }

    fields.namedItem('Nominativo').setCustomValidity(
        data.get('Nominativo').trim() ? '' : 'Inserisci il nominativo.'
    );
    fields.namedItem('Cellulare').setCustomValidity(
        data.get('Cellulare').trim() && /^[0-9 +]+$/.test(data.get('Cellulare'))
            ? '' : 'Inserisci il cellulare usando solo numeri, spazi e il carattere +.'
    );

    form.querySelectorAll('input, select, textarea').forEach(field => {
        if (field.disabled) return;
        if (field.required && !field.value.trim()) {
            field.setCustomValidity('Campo obbligatorio: inserisci un valore.');
        }
        if (field.dataset.touched === 'true') {
            updateBookingFieldError(field);
        }
    });
}

async function submitBooking(event) {
    event.preventDefault();
    const form = event.currentTarget;
    if (form.dataset.saving === 'true') return;
    const bookingData = collectBookingData(event);
    if (!bookingData) return;

    const status = form.querySelector('#bookingSaveStatus');
    status.hidden = false;
    if (!bookingScriptURL) {
        status.textContent = 'Configura bookingScriptURL con l\'URL /exec della Web App Google Apps Script.';
        return;
    }

    const controls = Array.from(form.querySelectorAll('input, select, textarea, button'));
    const disabledStates = controls.map(control => control.disabled);
    const button = form.querySelector('button[type="submit"]');
    form.dataset.saving = 'true';
    controls.forEach(control => { control.disabled = true; });
    button.textContent = 'Salvataggio...';
    status.textContent = 'Salvataggio in corso...';

    try {
        const response = await fetch(bookingScriptURL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'savePrenotazione', ...bookingData })
        });
        if (!response.ok) throw new Error('Risposta HTTP ' + response.status);
        const result = await response.json();
        if (result.status !== 'success') {
            status.textContent = result.message || 'Salvataggio non riuscito.';
            return;
        }

        status.textContent = 'Prenotazione salvata.';
        form.reset();
        await changeBookingYear(String(result.year));
    } catch (error) {
        status.textContent = 'Salvataggio non confermato. Verifica il foglio prima di riprovare. ' + error.message;
    } finally {
        controls.forEach((control, index) => { control.disabled = disabledStates[index]; });
        button.textContent = 'Inserisci prenotazione';
        form.dataset.saving = 'false';
        updateBookingCleaningField(form);
    }
}

function collectBookingData(event) {
    event.preventDefault();
    const form = event.currentTarget;
    validateBookingForm(form);
    if (!form.reportValidity()) return null;

    const data = new FormData(form);
    const guests = data.get('Numero Ospiti').trim();
    const taxGuests = Number(data.get('Ospiti soggetti a tassa di soggiorno'));
    const euroFormatter = new Intl.NumberFormat('it-IT', {
        useGrouping: false,
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
    const bookingData = {
        channel: data.get('channel'),
        nominativo: data.get('Nominativo').trim(),
        cellulare: data.get('Cellulare').trim(),
        checkin: data.get('Check-in'),
        checkout: data.get('Check-out'),
        numeroOspiti: guests,
        ospitiSoggettiATassa: taxGuests,
        note: data.get('Note').trim(),
        guadagnoLordo: euroFormatter.format(Number(data.get('Guadagno Lordo').replace(',', '.')))
    };

    if (bookingData.channel === 'Airbnb') {
        bookingData.costiPulizia = euroFormatter.format(Number(data.get('Costi di pulizia').replace(',', '.')));
    }

    return bookingData;
}


function showListView() {

    sessionStorage.setItem('lastActivePanel', 'lista');

    // Rimuove la classe active da tutti i pulsanti
    document.getElementById('calendarViewBtn').classList.remove('active');
    document.getElementById('graphicViewBtn').classList.remove('active');
    // Aggiunge la classe active al pulsante Lista
    document.getElementById('listViewBtn').classList.add('active');

    // Aggiorna il titolo
    updateViewHeader('Prenotazioni', 'lista');
 
    // svuota la visualizzazione precedente
    const listView = document.getElementById('listView');
    listView.innerHTML = ''; 
    
    // Filtra solo le prenotazioni con Check-in o Check-out >= data odierna
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Reset ore per confronto solo data

    const filteredBookings = arrayBooking.filter(booking => {
        const checkinDate = parseItalianDate(booking['Check-in']);
        const checkoutDate = parseItalianDate(booking['Check-out']);
        return checkinDate >= today || checkoutDate >= today;
    });

    // Inserisce i controlli della lista: aggiornamento remoto e ricerca rapida.
    // La ricerca parte ad ogni lettera digitata e filtra le card in base al nominativo.
    listView.innerHTML = `
        <div class="list-actions">
            <button type="button" class="reload-remote-btn" onclick="loadFromRemote()">
                <i class="fa fa-refresh" aria-hidden="true"></i>
                Aggiorna da remoto
            </button>
            <input type="text" class="search-input" placeholder="Cerca prenotazione" onkeyup="filterBookings(this)">
        </div>
    `; 

    // visualizza una div per ogni prenotazione con i campi Nominativo, Check-in, Check-out, Notti, Numero Ospiti
    filteredBookings.forEach(booking => {
        const bookingDiv = document.createElement('div');
        bookingDiv.classList.add('booking-item');
           
        printCard(bookingDiv, booking);
 
        document.getElementById('listView').appendChild(bookingDiv);
    });

    // Mostra la vista lista e nasconde le altre viste  
    document.getElementById('listView').classList.add('active');
    document.getElementById('calendarView').classList.remove('active');
    document.getElementById('graphicView').classList.remove('active');


    
}

function filterBookings(input) {
    const filter = input.value.toLowerCase();
    const bookingItems = document.querySelectorAll('.booking-item');
    bookingItems.forEach(item => {
        const nominativo = item.querySelector('p strong').textContent.toLowerCase();
        if (nominativo.includes(filter)) {
            item.style.display = '';
        } else {
            item.style.display = 'none';
        }   
    }); 
}   


function showCalendarView() {
    sessionStorage.setItem('lastActivePanel', 'calendario');

    // Rimuove la classe active da tutti i pulsanti
    document.getElementById('listViewBtn').classList.remove('active');
    document.getElementById('graphicViewBtn').classList.remove('active');
    // Aggiunge la classe active al pulsante Calendario
    document.getElementById('calendarViewBtn').classList.add('active');

    // Aggiorna il titolo
    updateViewHeader('Disponibilità', 'calendario');
    
    renderCalendar();

    // Mostra la vista calendario e nasconde le altre viste
    document.getElementById('calendarView').classList.add('active');
    document.getElementById('listView').classList.remove('active');
    document.getElementById('graphicView').classList.remove('active');

    
}

function renderCalendar() {
    const calendarContainer = document.getElementById('calendarView');
    calendarContainer.innerHTML = ''; // Pulisce il contenitore
    
    // Crea l'header del calendario con navigazione
    const header = document.createElement('div');
    header.className = 'calendar-header';
    const monthNames = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                        'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    header.innerHTML = `
        <button class="nav-button" onclick="previousMonth()">◀</button>
        <h2>${monthNames[currentMonth]} ${currentYear}</h2>
        <button class="nav-button" onclick="nextMonth()">▶</button>
    `;
    calendarContainer.appendChild(header);

    // Crea la griglia del calendario
    const calendarGrid = document.createElement('div');
    calendarGrid.className = 'calendar-grid';

    // Intestazioni giorni della settimana
    const dayNames = ['Lun', 'Mar', 'Mer', 'Gio', 'Ven', 'Sab', 'Dom'];
    dayNames.forEach(day => {
        const dayHeader = document.createElement('div');
        dayHeader.className = 'day-header';
        dayHeader.textContent = day;
        calendarGrid.appendChild(dayHeader);
    });

    // Calcola il primo giorno del mese e il numero di giorni
    const firstDay = new Date(currentYear, currentMonth, 1);
    const lastDay = new Date(currentYear, currentMonth + 1, 0);
    const daysInMonth = lastDay.getDate();
    
    // Ottiene il giorno della settimana (0=domenica, 1=lunedì, ...)
    let firstDayOfWeek = firstDay.getDay();
    // Converte da 0=domenica a 0=lunedì
    firstDayOfWeek = firstDayOfWeek === 0 ? 6 : firstDayOfWeek - 1;

    // Aggiungi celle vuote prima del primo giorno
    for (let i = 0; i < firstDayOfWeek; i++) {
        const emptyCell = document.createElement('div');
        emptyCell.className = 'calendar-day empty';
        calendarGrid.appendChild(emptyCell);
    }

    // Crea le celle per ogni giorno del mese
    for (let day = 1; day <= daysInMonth; day++) {
        const dayCell = document.createElement('div');
        dayCell.className = 'calendar-day';
        
        const currentDate = new Date(currentYear, currentMonth, day);
        const dateString = `${String(day).padStart(2, '0')}/${String(currentMonth + 1).padStart(2, '0')}/${currentYear}`;
        
        // Verifica se è il giorno corrente
        const today = new Date();
        if (currentDate.getDate() === today.getDate() && 
            currentDate.getMonth() === today.getMonth() && 
            currentDate.getFullYear() === today.getFullYear()) {
            dayCell.classList.add('today');
        }
        
        // Determina lo stato del giorno
        const bookingsForDay = getBookingsForDate(currentDate, dateString);
        const hasCheckOut = hasCheckOutOnDate(currentDate, dateString);
        
        if (hasCheckOut && bookingsForDay.length > 0) {
            // Checkout + prenotazione attiva = gradiente arancione-rosso
            dayCell.classList.add('checkout-occupied');
        } else if (hasCheckOut && bookingsForDay.length === 0) {
            // Checkout + nessuna prenotazione = gradiente arancione-verde
            dayCell.classList.add('checkout-free');
        } else if (bookingsForDay.length > 0) {
            // Solo prenotazione attiva = rosso normale
            dayCell.classList.add('occupied');
        } else {
            // Giorno libero = verde normale
            dayCell.classList.add('free');
        }
        
        dayCell.innerHTML = `
            <div class="day-number">${day}</div>
        `;
        
        // Aggiungi event listener per il click
        dayCell.addEventListener('click', () => showDayDetails(dateString, bookingsForDay));
        
        calendarGrid.appendChild(dayCell);
    }

    calendarContainer.appendChild(calendarGrid);

    // Aggiungi la sezione per i dettagli del giorno
    const detailsSection = document.createElement('div');
    detailsSection.id = 'day-details';
    detailsSection.className = 'day-details';
    calendarContainer.appendChild(detailsSection);
}

function getBookingsForDate(date, dateString) {
    return arrayBooking.filter(booking => {
        const checkin = parseItalianDate(booking['Check-in']);
        const checkout = parseItalianDate(booking['Check-out']);
        
        // La prenotazione è attiva se la data è >= check-in e < check-out
        return date >= checkin && date < checkout;
    });
}

function hasCheckOutOnDate(date, dateString) {
    return arrayBooking.some(booking => {
        const checkout = parseItalianDate(booking['Check-out']);
        return date.getTime() === checkout.getTime();
    });
}

function showDayDetails(dateString, bookings) {
    const detailsContainer = document.getElementById('day-details');
    
    if (bookings.length === 0) {
        detailsContainer.innerHTML = `<p class="no-bookings">Nessuna prenotazione per il ${dateString}</p>`;
        // Scroll al div dei dettagli
        detailsContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
        return;
    }
    
    detailsContainer.innerHTML = `<h3>Prenotazioni per il ${dateString}</h3>`;
    
    bookings.forEach(booking => {
        const bookingCard = document.createElement('div');
        bookingCard.className = 'booking-card';
        
        printCard(bookingCard, booking);

        detailsContainer.appendChild(bookingCard);
    });
    
    // Scroll al div dei dettagli
    detailsContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function printCard(component, booking) {
        component.innerHTML = `
            <p><strong>${booking.Nominativo}</strong></p>
            <p><i class="fa fa-calendar-o" aria-hidden="true"></i> ${booking['Check-in']} - ${booking['Check-out']}</p>
            <p><i class="fa fa-moon-o" aria-hidden="true"></i> ${booking.Notti} Notti</p>
            <p><i class="fa fa-user-o" aria-hidden="true"></i> ${booking['Numero Ospiti']} ${booking['Numero Ospiti'].toString().length === 1 ? 'Ospiti' : ''}</p>
            <p><i class="fa fa-percent" aria-hidden="true"></i> ${booking['Tassa di soggiorno'] && booking.hasOwnProperty('Tassa di soggiorno') ? formatEuro(booking['Tassa di soggiorno']) + ' € Tassa di soggiorno' : 'No Tax'}</p>
            <p><i class="fa fa-phone" aria-hidden="true"></i> ${booking.Cellulare && booking.hasOwnProperty("Cellulare") ? booking.Cellulare.replace(/\\/g, '+') : ' - '}</p>
            <p><i class="fa fa-sticky-note-o" aria-hidden="true"></i> ${booking.Note && booking.hasOwnProperty("Note") ? booking.Note : ' - '}</p>
            <div class="logo"><img src="./img/${booking.channel}.png" alt="Logo" width="40"></div>
            <div class="pay">
                <span class="lordo" style="display:none;">${booking['Guadagno Lordo'] && booking.hasOwnProperty('Guadagno Lordo') ? formatEuro(booking['Guadagno Lordo']) + ' € /' : ''}</span>
                <span class="netto" style="display:none;">${booking['Guadagno Netto'] && booking.hasOwnProperty('Guadagno Netto') ? formatEuro(booking['Guadagno Netto']) + ' €' : ''}</span>
                <span class="noLordo" style="display:inline;"> -- /</span>
                <span class="noNetto" style="display:inline;"> -- €</span>
            </div>
        `;
        
        // Aggiungi l'event listener dopo aver creato l'HTML
        const payElement = component.querySelector('.pay');
        payElement.addEventListener('click', function() {
            showHidePayDetails();
        });
 
}

function showHidePayDetails() {

    // Verificare se le classi .noLordo sono attualmente nascoste o visibili
    const lordoElements = document.querySelectorAll('.lordo');
    const isCurrentlyHidden = lordoElements.length > 0 && lordoElements[0].style.display === 'none';

     if (isCurrentlyHidden) {
        // Mostra tutti i dati - visualizzando gli elementi con classi .lordo e .netto e nascondendo quelli con classi .noLordo e .noNetto
        document.querySelectorAll('.lordo').forEach(el => el.style.display = 'inline');
        document.querySelectorAll('.netto').forEach(el => el.style.display = 'inline');
        document.querySelectorAll('.noLordo').forEach(el => el.style.display = 'none');
        document.querySelectorAll('.noNetto').forEach(el => el.style.display = 'none');
    } else {
        // Nascondi tutti i dati - visualizzando gli elementi con classi .noLordo e .noNetto e nascondendo quelli con classi .lordo e .netto
        document.querySelectorAll('.lordo').forEach(el => el.style.display = 'none');
        document.querySelectorAll('.netto').forEach(el => el.style.display = 'none');
        document.querySelectorAll('.noLordo').forEach(el => el.style.display = 'inline');
        document.querySelectorAll('.noNetto').forEach(el => el.style.display = 'inline');
    }

}


function previousMonth() {
    currentMonth--;
    if (currentMonth < 0) {
        currentMonth = 11;
        currentYear--;
    }
    renderCalendar();
}

function nextMonth() {
    currentMonth++;
    if (currentMonth > 11) {
        currentMonth = 0;
        currentYear++;
    }
    renderCalendar();
}




// Funzione per convertire i numeri seriali di Excel in date leggibili
function excelDateToJSDate(serial) {
    if (!serial || typeof serial !== 'number') return serial;
    // Excel memorizza le date come giorni dal 1° gennaio 1900
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    
    // Formatta la data come gg/mm/aaaa
    const day = String(date_info.getDate()).padStart(2, '0');
    const month = String(date_info.getMonth() + 1).padStart(2, '0');
    const year = date_info.getFullYear();
    
    return `${day}/${month}/${year}`;
}

// Funzione per convertire stringa "gg/mm/aaaa" in oggetto Date per ordinamento
function parseItalianDate(dateString) {
    if (!dateString || typeof dateString !== 'string') return new Date(0);
    const parts = dateString.split('/');
    if (parts.length !== 3) return new Date(0);
    // new Date(year, month, day) - month è 0-based
    return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
}

// Funzione per formattare numeri in formato euro (es: 1.234,56)
function formatEuro(number) {
    if (!number || isNaN(number)) return '0,00';
    return new Intl.NumberFormat('it-IT', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(number);
}