// URL applicazione Script google
const bookingScriptURL = 'https://script.google.com/macros/s/AKfycbxFq5h6XHTyXlFMf8gx_ifdY_VuM1hq1oB4C-IqOIiJWys4yuaqCQRxnP023vqucooa3Q/exec';

const ALL_YEARS = 'tutti';
let arrayBooking = [];
let selectedBookingYear = new Date().getFullYear();
let pendingBookingYear = null;
let expensesRequestId = 0;
let graphicRequestId = 0;
let expensesCacheVersion = 0;
let currentMonth = new Date().getMonth();
let currentYear = new Date().getFullYear();
let cache = false;
const pendingBookingDeletions = new WeakSet();

function normalizeBookingYear(year) {
    return String(year).trim().toLowerCase() === ALL_YEARS ? ALL_YEARS : Number(year);
}

function titleWithYear(title) {
    return selectedBookingYear === ALL_YEARS ? title : `${title} ${selectedBookingYear}`;
}

function changeBookingYear(year) {
    pendingBookingYear = normalizeBookingYear(year);
    updateBookingYearSelection();
    return loadFromRemote(year).finally(() => {
        pendingBookingYear = null;
        updateBookingYearSelection();
    });
}

function loadFromRemote(year) {
    if (year == null || String(year).trim() === '') {
        year = new Date().getFullYear();
    }
    if (!bookingScriptURL) {
        console.error('URL della Web App Google Apps Script non configurato.');
        return Promise.resolve();
    }

    const bookingYear = normalizeBookingYear(year);
    const yearButtons = document.querySelectorAll('#bookingYearSelector .booking-year-button');
    yearButtons.forEach(button => { button.disabled = true; });
    const listView = document.getElementById('listView');
    document.getElementById('bookingItems').innerHTML = '';
    document.getElementById('loadingIndicator').hidden = false;
    listView.setAttribute('aria-busy', 'true');
    console.log('loadFromRemote - Caricamento dati da URL');
    const parameters = new URLSearchParams({ anno: String(bookingYear), action: 'leggi', token: localStorage.getItem('token') ?? '' });
    const requestURL = bookingScriptURL + (bookingScriptURL.includes('?') ? '&' : '?') + parameters.toString();

    return fetch(requestURL)
        .then(response => {
            if (!response.ok) throw new Error('Caricamento non riuscito (HTTP ' + response.status + ')');
            return response.json();
        })
        .then(result => {
            if (result.code === 401) {
                handleUnauthorized();
                return;
            }
            if (result.code !== 200) throw new Error(result.message || 'Caricamento prenotazioni non riuscito.');
            const bookings = typeof result.data === 'string' ? JSON.parse(result.data) : result.data;
            if (!Array.isArray(bookings)) throw new Error('Formato dati prenotazioni non valido.');

            if (arrayBooking.length > 0) {
                // Svuoto l'arrayBooking prima di concatenare i nuovi dati
                arrayBooking = [];
            }

            arrayBooking = bookings;

            // Dell'arrayBooking voglio solo le prenotazioni che hanno il campo nominativo valorizzato e il campo "Stato prenotazione", se esiste, diversa da "cancellata"
            //arrayBooking = arrayBooking.filter(booking => booking.Nominativo && (!booking.hasOwnProperty('Stato prenotazione') || booking['Stato prenotazione'].toLowerCase() != 'cancellata'));

            // ordina l'arrayBooking per data di arrivo (campo Check-in)
            arrayBooking.sort((a, b) => parseItalianDate(a['Check-in']) - parseItalianDate(b['Check-in']));

            selectedBookingYear = bookingYear;
            sessionStorage.setItem('arrayBooking', JSON.stringify(arrayBooking));
            sessionStorage.setItem('arrayBookingYear', String(selectedBookingYear));

            if (cache) {
                localStorage.setItem('arrayBooking', JSON.stringify(arrayBooking));
                localStorage.setItem('arrayBookingYear', String(selectedBookingYear));
            }

            showListView();
        })
        .catch(error => {
            updateBookingYearSelection();
            showListView();
            console.error('Errore durante il caricamento del file Excel:', error);
        })
        .finally(() => {
            listView.setAttribute('aria-busy', 'false');
            document.getElementById('loadingIndicator').hidden = true;
            document.querySelectorAll('#bookingYearSelector .booking-year-button').forEach(button => { button.disabled = false; });
        });

}

/**
 * Gestisce una risposta 401 (token mancante o scaduto): rimuove il token e torna alla schermata di login
 */
function handleUnauthorized() {
    localStorage.removeItem('token');
    showLoginView(() => loadFromRemote(selectedBookingYear));
}

function initializeApp() {
    const storedYear = sessionStorage.getItem('arrayBookingYear');
    if (document.querySelector(`#bookingYearSelector .booking-year-button[data-year="${storedYear}"]`)) {
        selectedBookingYear = normalizeBookingYear(storedYear);
    }
    updateBookingYearSelection();

    const startApp = () => {
        const lastPanel = sessionStorage.getItem('lastActivePanel');
        const cachedBookings = readCachedBookings();
        if (lastPanel && lastPanel !== 'lista' && cachedBookings) {
            arrayBooking = cachedBookings;
            restoreLastActivePanel(lastPanel);
            return;
        }
        showListView();
        loadFromRemote(selectedBookingYear);
    };

    showLoginView(startApp);
    if (localStorage.getItem('token')) {
        hideLoginView();
        startApp();
    }
}

function readCachedBookings() {
    if (sessionStorage.getItem('arrayBookingYear') !== String(selectedBookingYear)) return null;
    try {
        const bookings = JSON.parse(sessionStorage.getItem('arrayBooking'));
        return Array.isArray(bookings) ? bookings : null;
    } catch (error) {
        return null;
    }
}

function restoreLastActivePanel(panel) {
    if (!localStorage.getItem('token') || document.getElementById('loginView').classList.contains('active')) return;
    switch (panel) {
        case 'calendario':
            showCalendarView();
            break;
        case 'spese':
            showListExpenses();
            break;
        case 'grafico':
            showGraphicView();
            break;
        case 'inserisci':
            showInsertBooking();
            break;
        case 'inserisciSpesa':
            showInsertExpense();
            break;
        default:
            showListView();
    }
}

function hideLoginView() {
    document.getElementById('loginView').classList.remove('active');
    document.querySelector('.page-header').hidden = false;
    document.querySelector('.tabbar').hidden = false;
}

/**
 * Mostra la schermata di login e, dopo un accesso riuscito, esegue onLoginSuccess
 * @param {Function} onLoginSuccess callback da eseguire dopo il login
 */
function showLoginView(onLoginSuccess) {
    document.querySelector('.page-header').hidden = true;
    document.querySelector('.tabbar').hidden = true;
    document.querySelectorAll('.view').forEach(view => view.classList.remove('active'));
    document.getElementById('loginView').classList.add('active');
    toggleLoginPasswordVisibility(false);

    const form = document.getElementById('loginForm');
    form.onsubmit = event => {
        event.preventDefault();
        submitLogin(form, onLoginSuccess);
    };
}

function toggleLoginPasswordVisibility(visible = document.getElementById('loginPassword').type === 'password') {
    const input = document.getElementById('loginPassword');
    const button = document.getElementById('loginPasswordToggle');
    input.type = visible ? 'text' : 'password';
    const label = visible ? 'Nascondi password' : 'Mostra password';
    button.setAttribute('aria-label', label);
    button.setAttribute('aria-pressed', String(visible));
    button.title = label;
    const icon = button.querySelector('i');
    icon.classList.toggle('fa-eye', !visible);
    icon.classList.toggle('fa-eye-slash', visible);
}

async function submitLogin(form, onLoginSuccess) {
    const status = document.getElementById('loginStatus');
    const button = form.querySelector('button[type="submit"]');
    const passwordInput = form.elements.namedItem('password');
    const username = form.elements.namedItem('username').value.trim();
    const password = passwordInput.value;

    button.disabled = true;
    status.textContent = 'Accesso in corso...';

    try {
        const response = await fetch(bookingScriptURL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'login', username: username, password: password })
        });
        if (!response.ok) throw new Error('Risposta HTTP ' + response.status);
        const result = await response.json();
        if (result.code !== 200) throw new Error(result.message || 'Credenziali non valide.');

        localStorage.setItem('token', result.token);
    } catch (error) {
        status.textContent = 'Accesso non riuscito. ' + error.message;
        passwordInput.value = '';
        toggleLoginPasswordVisibility(false);
        button.disabled = false;
        return;
    }

    status.textContent = '';
    form.reset();
    toggleLoginPasswordVisibility(false);
    button.disabled = false;
    hideLoginView();
    onLoginSuccess();
}


function showGraphicView() {

    sessionStorage.setItem('lastActivePanel', 'grafico');
    currentYear = selectedBookingYear === ALL_YEARS ? new Date().getFullYear() : selectedBookingYear;

    // Rimuove la classe active da tutti i pulsanti
    document.getElementById('listViewBtn').classList.remove('active');
    document.getElementById('calendarViewBtn').classList.remove('active');
    // Aggiunge la classe active al pulsante Grafico
    document.getElementById('graphicViewBtn').classList.add('active');

    // Aggiorna il titolo
    updateViewHeader(titleWithYear('Profitto'), 'grafico');

    const rendering = renderGraphicView();
  
    // Mostra la vista grafico e nasconde le altre viste
    document.getElementById('graphicView').classList.add('active');
    document.getElementById('listView').classList.remove('active');
    document.getElementById('calendarView').classList.remove('active');

    return rendering;
}

async function renderGraphicView() {
    const requestId = ++graphicRequestId;
    const year = currentYear;
    const month = currentMonth;
    const graphicContainer = document.getElementById('graphicView');
    graphicContainer.innerHTML = ''; // Pulisce il contenitore
    graphicContainer.setAttribute('aria-busy', 'true');
    if (graphicContainer.classList.contains('active')) updateViewHeader(titleWithYear('Profitto'), 'grafico');
    
    // Crea l'header con navigazione
    const header = document.createElement('div');
    header.className = 'calendar-header';
    const monthNames = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                        'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    header.innerHTML = `
        <button class="nav-button" onclick="previousMonthGraphic()">◀</button>
        <h2>${monthNames[month]} ${year}</h2>
        <button class="nav-button" onclick="nextMonthGraphic()">▶</button>
    `;
    graphicContainer.appendChild(header);

    // Contenitore per le statistiche
    const statsContainer = document.createElement('div');
    statsContainer.className = 'stats-container';

    // Filtra le prenotazioni per il mese corrente in base alla data di check-in
    const monthStart = new Date(year, month, 1);
    const monthEnd = new Date(year, month + 1, 0);
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
    const yearlyRevenue = calculateYearlyRevenue(year);

    // Crea la visualizzazione con due card separate
    statsContainer.innerHTML = `
        <div class="stats-card">
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
                    <p class="channel-name">Diretto</p>
                    <p class="channel-revenue">${formatEuro(revenueByChannel.black)} €</p>
                </div>
            </div>
            <div class="revenue-item profit-expenses">
                <i class="fa fa-credit-card expense-icon" aria-hidden="true"></i>
                <div class="channel-info">
                    <p class="channel-name">Spese</p>
                    <p class="expense-amount monthly-expenses-amount">Caricamento...</p>
                </div>
            </div>
            <div class="revenue-total">
                <p class="total-label">Totale Mensile:</p>
                <p class="total-revenue">Caricamento...</p>
            </div>
            <div class="booking-count">
                <p>Numero prenotazioni: <strong>${bookingsInMonth.length}</strong></p>
            </div>
        </div>
        
        <div class="stats-card yearly-card">
            <div class="yearly-revenue">
                <i class="fa fa-line-chart" aria-hidden="true"></i>
                <div class="yearly-info">
                    <p class="yearly-label">Revenue Annuale ${year}</p>
                    <p class="yearly-amount">${formatEuro(yearlyRevenue)} €</p>
                </div>
            </div>
            <div class="yearly-revenue yearly-expenses">
                <i class="fa fa-credit-card" aria-hidden="true"></i>
                <div class="yearly-info">
                    <p class="yearly-label">Spese Annuali ${year}</p>
                    <p class="yearly-amount expense-amount annual-expenses-amount">Caricamento...</p>
                </div>
            </div>
        </div>
    `;

    graphicContainer.appendChild(statsContainer);
    const status = document.createElement('p');
    status.className = 'expenses-status expenses-error';
    status.setAttribute('role', 'status');
    status.hidden = true;
    graphicContainer.appendChild(status);

    try {
        const expenses = await loadExpenses(selectedBookingYear);
        if (requestId !== graphicRequestId) return;
        const monthlyTotal = expenses.reduce((sum, expense) => {
            const date = parseItalianDate(expense.Data);
            return date >= monthStart && date <= monthEnd ? sum + (expense.Importo || 0) : sum;
        }, 0);
        const annualTotal = expenses.reduce((sum, expense) =>
            parseItalianDate(expense.Data).getFullYear() === year ? sum + (expense.Importo || 0) : sum, 0);
        statsContainer.querySelector('.monthly-expenses-amount').textContent = formatEuro(monthlyTotal) + ' \u20ac';
        statsContainer.querySelector('.annual-expenses-amount').textContent = formatEuro(annualTotal) + ' \u20ac';
        statsContainer.querySelector('.total-revenue').textContent = formatEuro(totalRevenue - monthlyTotal) + ' \u20ac';
    } catch (error) {
        if (requestId !== graphicRequestId) return;
        if (error.code === 401) {
            localStorage.removeItem('token');
            showLoginView(() => showGraphicView());
            return;
        }
        statsContainer.querySelectorAll('.monthly-expenses-amount, .annual-expenses-amount, .total-revenue')
            .forEach(amount => { amount.textContent = 'Non disponibile'; });
        status.textContent = 'Caricamento spese non riuscito. ' + error.message;
        status.hidden = false;
        console.error('Errore durante il caricamento delle spese per il profitto:', error);
    } finally {
        if (requestId === graphicRequestId) graphicContainer.setAttribute('aria-busy', 'false');
    }
}

function previousMonthGraphic() {
    currentMonth--;
    if (currentMonth < 0) {
        currentMonth = 11;
        currentYear--;
    }
    return renderGraphicView();
}

function nextMonthGraphic() {
    currentMonth++;
    if (currentMonth > 11) {
        currentMonth = 0;
        currentYear++;
    }
    return renderGraphicView();
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
    updateBookingYearSelection();
    document.getElementById('expensesView').classList.toggle('active', panel === 'spese');
    document.getElementById('expensesViewBtn').classList.toggle('active', panel === 'spese' || panel === 'inserisciSpesa');
    document.getElementById('insertBookingBtn').hidden = panel !== 'lista';
    document.getElementById('insertExpenseBtn').hidden = panel !== 'spese';
    document.getElementById('backToListBtn').hidden = panel !== 'inserisci';
    document.getElementById('backToExpensesBtn').hidden = panel !== 'inserisciSpesa';
    document.getElementById('insertBookingView').classList.toggle('active', panel === 'inserisci');
    document.getElementById('insertExpenseView').classList.toggle('active', panel === 'inserisciSpesa');
}

function updateBookingYearSelection() {
    document.querySelectorAll('#bookingYearSelector .booking-year-button').forEach(button => {
        const isSelected = button.dataset.year === String(pendingBookingYear ?? selectedBookingYear);
        button.classList.toggle('active', isSelected);
        button.setAttribute('aria-pressed', String(isSelected));
    });
}

function showInsertBooking(booking = null) {
    sessionStorage.setItem('lastActivePanel', 'inserisci');
    document.querySelectorAll('.view, .tabButton').forEach(element => {
        element.classList.remove('active');
    });
    document.getElementById('listViewBtn').classList.add('active');
    updateViewHeader(booking ? 'Modifica Prenotazione' : 'Nuova Prenotazione', 'inserisci');

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
                <input id="bookingPhone" name="Cellulare" type="tel" autocomplete="tel" title="Sono ammessi solo numeri, spazi e il carattere +">
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
            <div id="bookingTaxGuestsField" class="booking-field">
                <label for="bookingTaxGuests">Ospiti soggetti a tassa di soggiorno</label>
                <input id="bookingTaxGuests" name="Ospiti soggetti a tassa di soggiorno" type="number" inputmode="numeric" min="0" step="1" required>
            </div>
            <div class="booking-field booking-field-wide">
                <label for="bookingGross">Guadagno Lordo (euro)</label>
                <input id="bookingGross" name="Guadagno Lordo" type="number" inputmode="decimal" min="0" step="0.01" required>
            </div>
            <div id="bookingCleaningField" class="booking-field booking-field-wide">
                <label for="bookingCleaning">Costi di pulizia (euro)</label>
                <input id="bookingCleaning" name="Costi di pulizia" type="number" inputmode="decimal" min="0" step="0.01" value="0" required>
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
    form.dataset.action = booking ? 'modificaPrenotazione' : 'savePrenotazione';
    if (booking) {
        form.dataset.id = String(booking.ID ?? '');
        const channels = { booking: 'Booking.com', 'booking.com': 'Booking.com', airbnb: 'Airbnb', black: 'Black' };
        const channel = channels[String(booking.channel).toLowerCase()];
        const values = {
            channel: channel,
            Nominativo: booking.Nominativo,
            Cellulare: String(booking.Cellulare ?? '').replace(/\\/g, '+'),
            'Check-in': booking['Check-in'],
            'Check-out': booking['Check-out'],
            'Numero Ospiti': booking['Numero Ospiti'],
            'Ospiti soggetti a tassa di soggiorno': booking['Ospiti soggetti a tassa di soggiorno'],
            'Guadagno Lordo': booking['Guadagno Lordo'],
            'Costi di pulizia': booking['Costi di pulizia'] === '' || booking['Costi di pulizia'] == null ? 0 : booking['Costi di pulizia'],
            Note: booking.Note
        };
        Object.entries(values).forEach(([name, value]) => {
            form.elements.namedItem(name).value = value ?? '';
        });
        form.querySelector('button[type="submit"]').textContent = 'Salva modifiche';
    }
    updateBookingChannelFields(form);
    form.elements.namedItem('channel').addEventListener('change', () => updateBookingChannelFields(form));
    form.addEventListener('submit', submitBooking);
    form.addEventListener('focusout', event => {
        if (!event.target.matches('input, select, textarea')) return;
        event.target.dataset.touched = 'true';
        validateBookingForm(form);
        if (!event.target.validity.valid) event.target.reportValidity();
    });
    form.addEventListener('input', event => {
        if (!event.target.matches('input, select, textarea')) return;
        if (event.target.id === 'bookingCheckin' || event.target.id === 'bookingCheckout') {
            formatBookingDateInput(event.target);
        }
        event.target.setCustomValidity('');
        event.target.removeAttribute('aria-invalid');
    });
    form.addEventListener('invalid', event => {
        event.target.dataset.touched = 'true';
        updateBookingFieldError(event.target);
    }, true);
}

function formatBookingDateInput(input) {
    const cursor = input.selectionStart ?? input.value.length;
    const digits = input.value.replace(/\D/g, '').slice(0, 8);
    const digitsBeforeCursor = Math.min(input.value.slice(0, cursor).replace(/\D/g, '').length, digits.length);
    let formatted = digits.slice(0, 2);
    if (digits.length > 2) formatted += '/' + digits.slice(2, 4);
    if (digits.length > 4) formatted += '/' + digits.slice(4);

    input.value = formatted;
    const nextCursor = digitsBeforeCursor + (digitsBeforeCursor > 2 ? 1 : 0) + (digitsBeforeCursor > 4 ? 1 : 0);
    input.setSelectionRange(nextCursor, nextCursor);
}

function updateBookingChannelFields(form) {
    const showField = form.elements.namedItem('channel').value !== 'Black';
    [['bookingTaxGuestsField', 'Ospiti soggetti a tassa di soggiorno'], ['bookingCleaningField', 'Costi di pulizia']]
        .forEach(([containerId, name]) => {
            const field = form.elements.namedItem(name);
            document.getElementById(containerId).hidden = !showField;
            field.disabled = !showField;
            field.required = showField;
            if (!showField) {
                field.setCustomValidity('');
                field.removeAttribute('aria-invalid');
                delete field.dataset.touched;
            }
        });
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
        'Costi di pulizia': 'Inserisci un importo maggiore o uguale a 0 con massimo due decimali, ad esempio 250,75.',
        data: 'Inserisci una data valida in formato gg/mm/aaaa.',
        categoria: 'Inserisci la categoria.',
        descrizione: 'Inserisci la descrizione.',
        importo: 'Inserisci un importo maggiore o uguale a 0 con massimo due decimali, ad esempio 250,75.'
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
        !data.get('Cellulare').trim() || /^[0-9 +]+$/.test(data.get('Cellulare'))
            ? '' : 'Inserisci il cellulare usando solo numeri, spazi e il carattere +.'
    );

    const taxField = fields.namedItem('Ospiti soggetti a tassa di soggiorno');
    if (!taxField.disabled) {
        taxField.setCustomValidity(
            /^[0-9]+$/.test(taxField.value) && Number.isSafeInteger(Number(taxField.value))
                ? '' : 'Inserisci solo cifre da 0 a 9, senza segni o decimali.'
        );
    }

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
    const updating = form.dataset.action === 'modificaPrenotazione';

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
            body: JSON.stringify({ action: form.dataset.action, token: localStorage.getItem('token') ?? '', ...bookingData })
        });
        if (!response.ok) throw new Error('Risposta HTTP ' + response.status);
        const result = await response.json();
        if (result.code === 401) {
            handleUnauthorized();
            return;
        }
        if (result.status !== 'success') {
            return;
        }

        status.textContent = 'Prenotazione salvata.';
        if (!updating) form.reset();
        await changeBookingYear(selectedBookingYear === ALL_YEARS ? ALL_YEARS : String(result.year));
    } catch (error) {
        status.textContent = 'Salvataggio non confermato. Verifica il foglio prima di riprovare. ' + error.message;
    } finally {
        controls.forEach((control, index) => { control.disabled = disabledStates[index]; });
        button.textContent = updating ? 'Salva modifiche' : 'Inserisci prenotazione';
        form.dataset.saving = 'false';
        updateBookingChannelFields(form);
    }
}

function collectBookingData(event) {
    event.preventDefault();
    const form = event.currentTarget;
    validateBookingForm(form);
    if (!form.reportValidity()) return null;

    const data = new FormData(form);
    const guests = data.get('Numero Ospiti').trim();
    const euroFormatter = new Intl.NumberFormat('it-IT', {
        useGrouping: false,
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
    const bookingData = {
        channel: form.elements.namedItem('channel').value,
        nominativo: data.get('Nominativo').trim(),
        cellulare: data.get('Cellulare').trim(),
        checkin: data.get('Check-in'),
        checkout: data.get('Check-out'),
        numeroOspiti: guests,
        note: data.get('Note').trim(),
        guadagnoLordo: euroFormatter.format(Number(data.get('Guadagno Lordo').replace(',', '.'))),
        costiPulizia: form.elements.namedItem('channel').value === 'Black'
            ? euroFormatter.format(0)
            : euroFormatter.format(Number(data.get('Costi di pulizia').replace(',', '.')))
    };
    if (form.dataset.action === 'modificaPrenotazione') {
        bookingData.id = Number(form.dataset.id);
    }

    if (bookingData.channel !== 'Black') {
        bookingData.ospitiSoggettiATassa = Number(data.get('Ospiti soggetti a tassa di soggiorno'));
    }

    return bookingData;
}

function showInsertExpense() {
    sessionStorage.setItem('lastActivePanel', 'inserisciSpesa');
    document.querySelectorAll('.view, .tabButton').forEach(element => element.classList.remove('active'));
    updateViewHeader(titleWithYear('Nuova Spesa'), 'inserisciSpesa');

    const container = document.getElementById('insertExpenseView');
    container.innerHTML = `
        <form id="expenseForm" class="booking-form">
            <div class="booking-field">
                <label for="expenseDate">Data (gg/mm/aaaa)</label>
                <input id="expenseDate" name="data" type="text" inputmode="numeric" placeholder="gg/mm/aaaa" pattern="[0-9]{2}/[0-9]{2}/[0-9]{4}" maxlength="10" required>
            </div>
            <div class="booking-field">
                <label for="expenseCategory">Categoria</label>
                <select id="expenseCategory" name="categoria" required>
                    <option value="">Seleziona una categoria</option>
                    <option>Pulizia e lavanderia</option>
                    <option>Utenze e consumi</option>
                    <option>Manutenzione e riparazioni</option>
                    <option>Materiale di consumo e cortesia</option>
                    <option>Arredamento e dotazioni</option>
                    <option>Assicurazione e sicurezza</option>
                    <option>Lavori e investimenti</option>
                    <option>Spese straordinarie e varie</option>
                </select>
            </div>
            <div class="booking-field booking-field-wide">
                <label for="expenseDescription">Descrizione</label>
                <input id="expenseDescription" name="descrizione" type="text" required>
            </div>
            <div class="booking-field booking-field-wide">
                <label for="expenseAmount">Importo (euro)</label>
                <input id="expenseAmount" name="importo" type="text" inputmode="decimal" placeholder="0,00" pattern="[0-9]+([.,][0-9]{1,2})?" required>
            </div>
            <div class="booking-field booking-field-wide">
                <label for="expenseNotes">Note</label>
                <textarea id="expenseNotes" name="note" rows="4"></textarea>
            </div>
            <button type="submit" class="booking-submit booking-field-wide">Inserisci spesa</button>
            <p id="expenseSaveStatus" class="booking-field-wide" role="status" hidden></p>
        </form>
    `;
    const form = document.getElementById('expenseForm');
    form.dataset.year = String(selectedBookingYear);
    const today = new Date();
    form.elements.namedItem('data').value =
        `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
    form.addEventListener('submit', submitExpense);
    form.addEventListener('focusout', event => {
        if (!event.target.matches('input, select, textarea')) return;
        event.target.dataset.touched = 'true';
        validateExpenseForm(form);
        if (!event.target.validity.valid) event.target.reportValidity();
    });
    form.addEventListener('input', event => {
        if (!event.target.matches('input, select, textarea')) return;
        if (event.target.name === 'data') formatBookingDateInput(event.target);
        event.target.setCustomValidity('');
        event.target.removeAttribute('aria-invalid');
    });
    form.addEventListener('invalid', event => {
        event.target.dataset.touched = 'true';
        updateBookingFieldError(event.target);
    }, true);
}

function validateExpenseForm(form) {
    form.querySelectorAll('input, select, textarea').forEach(field => field.setCustomValidity(''));
    const dateField = form.elements.namedItem('data');
    const date = parseItalianDate(dateField.value);
    const formattedDate = `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
    dateField.setCustomValidity(
        /^[0-9]{2}\/[0-9]{2}\/[0-9]{4}$/.test(dateField.value) && formattedDate === dateField.value
            ? '' : 'Inserisci una data valida in formato gg/mm/aaaa.'
    );
    const amountField = form.elements.namedItem('importo');
    const amount = amountField.value.trim();
    amountField.setCustomValidity(
        /^\d+([.,]\d{1,2})?$/.test(amount) && Number.isFinite(Number(amount.replace(',', '.')))
            ? '' : 'Inserisci un importo maggiore o uguale a 0 con massimo due decimali.'
    );
    form.querySelectorAll('input, select, textarea').forEach(field => {
        if (field.required && !field.value.trim()) field.setCustomValidity('Campo obbligatorio: inserisci un valore.');
        if (field.dataset.touched === 'true') updateBookingFieldError(field);
    });
}

async function submitExpense(event) {
    event.preventDefault();
    const form = event.currentTarget;
    if (form.dataset.saving === 'true') return;
    validateExpenseForm(form);
    if (!form.reportValidity()) return;

    const data = new FormData(form);
    const year = normalizeBookingYear(form.dataset.year);
    const expenseData = {
        action: 'saveSpesa',
        token: localStorage.getItem('token') ?? '',
        anno: year,
        data: data.get('data'),
        categoria: data.get('categoria').trim(),
        descrizione: data.get('descrizione').trim(),
        importo: data.get('importo').trim(),
        note: data.get('note').trim()
    };
    const status = form.querySelector('#expenseSaveStatus');
    const controls = Array.from(form.querySelectorAll('input, select, textarea, button'));
    const button = form.querySelector('button[type="submit"]');
    form.dataset.saving = 'true';
    controls.forEach(control => { control.disabled = true; });
    button.textContent = 'Salvataggio...';
    status.hidden = false;
    status.textContent = 'Salvataggio in corso...';

    try {
        if (!bookingScriptURL) throw new Error('URL della Web App Google Apps Script non configurato.');
        const response = await fetch(bookingScriptURL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify(expenseData)
        });
        if (!response.ok) throw new Error('Risposta HTTP ' + response.status);
        const result = await response.json();
        if (result.code === 401) {
            localStorage.removeItem('token');
            status.textContent = 'Sessione scaduta. Accedi nuovamente e salva la spesa.';
            showLoginView(() => {
                document.querySelectorAll('.view, .tabButton').forEach(element => element.classList.remove('active'));
                updateViewHeader(titleWithYear('Nuova Spesa'), 'inserisciSpesa');
            });
            return;
        }
        if (result.status !== 'success') throw new Error(result.message || 'Salvataggio spesa non riuscito.');
        if (!Number.isInteger(result.year)) throw new Error('Anno nella risposta di salvataggio non valido.');

        expensesRequestId++;
        graphicRequestId++;
        expensesCacheVersion++;
        sessionStorage.removeItem('arrayExpenses_' + year);
        sessionStorage.removeItem('arrayExpenses_' + result.year);
        sessionStorage.removeItem('arrayExpenses_' + ALL_YEARS);
        status.textContent = 'Spesa salvata.';
        form.reset();
        if (document.getElementById('expensesView').classList.contains('active') ||
            (document.getElementById('insertExpenseView').classList.contains('active') &&
                document.getElementById('expenseForm') === form)) {
            await showListExpenses();
        } else if (document.getElementById('graphicView').classList.contains('active')) {
            await renderGraphicView();
        }
    } catch (error) {
        status.textContent = 'Salvataggio non confermato. Verifica il foglio prima di riprovare. ' + error.message;
    } finally {
        controls.forEach(control => { control.disabled = false; });
        button.textContent = 'Inserisci spesa';
        form.dataset.saving = 'false';
    }
}

async function loadExpenses(year) {
    const cacheKey = 'arrayExpenses_' + year;
    const cacheVersion = expensesCacheVersion;
    const cachedExpenses = sessionStorage.getItem(cacheKey);
    let expenses;
    if (cachedExpenses !== null) {
        expenses = JSON.parse(cachedExpenses);
    } else {
        if (!bookingScriptURL) throw new Error('URL della Web App Google Apps Script non configurato.');
        const parameters = new URLSearchParams({
            anno: String(year),
            action: 'leggiSpese',
            token: localStorage.getItem('token') ?? ''
        });
        const response = await fetch(bookingScriptURL + (bookingScriptURL.includes('?') ? '&' : '?') + parameters.toString());
        if (!response.ok) throw new Error('Caricamento non riuscito (HTTP ' + response.status + ')');
        const result = await response.json();
        if (result.code !== 200) {
            const error = new Error(result.message || 'Caricamento spese non riuscito.');
            error.code = result.code;
            throw error;
        }
        expenses = typeof result.data === 'string' ? JSON.parse(result.data) : result.data;
    }
    if (!Array.isArray(expenses) || expenses.some(expense =>
        !expense || typeof expense !== 'object' || Array.isArray(expense) ||
        (expense.Importo !== '' && expense.Importo != null &&
            (typeof expense.Importo !== 'number' || !Number.isFinite(expense.Importo))))) {
        throw new Error('Formato dati spese non valido.');
    }
    expenses.sort((a, b) => parseItalianDate(b.Data) - parseItalianDate(a.Data));
    if (cachedExpenses === null && cacheVersion === expensesCacheVersion) {
        sessionStorage.setItem(cacheKey, JSON.stringify(expenses));
    }
    return expenses;
}

async function showListExpenses() {
    const year = selectedBookingYear;
    const requestId = ++expensesRequestId;
    sessionStorage.setItem('lastActivePanel', 'spese');
    document.querySelectorAll('.view, .tabButton').forEach(element => element.classList.remove('active'));
    updateViewHeader(titleWithYear('Spese'), 'spese');

    const view = document.getElementById('expensesView');
    const items = document.getElementById('expenseItems');
    const loading = document.getElementById('expensesLoadingIndicator');
    const status = document.getElementById('expensesStatus');
    items.replaceChildren();
    status.hidden = true;
    status.textContent = '';
    status.classList.remove('expenses-error');
    loading.hidden = true;
    view.setAttribute('aria-busy', 'false');

    try {
        loading.hidden = sessionStorage.getItem('arrayExpenses_' + year) !== null;
        view.setAttribute('aria-busy', String(!loading.hidden));
        const expenses = await loadExpenses(year);
        if (requestId !== expensesRequestId) return;
        expenses.forEach(expense => {
            const card = document.createElement('article');
            card.className = 'expense-item';
            card.innerHTML = `
                <div class="expense-heading">
                    <strong class="expense-category"></strong>
                    <span class="expense-amount"></span>
                </div>
                <p class="expense-date"><i class="fa fa-calendar-o" aria-hidden="true"></i> <span></span></p>
                <p class="expense-description"><i class="fa fa-file-text-o" aria-hidden="true"></i> <span></span></p>
                <p class="expense-note"><i class="fa fa-sticky-note-o" aria-hidden="true"></i> <span></span></p>
            `;
            card.querySelector('.expense-category').textContent = expense.Categoria || 'Senza categoria';
            card.querySelector('.expense-amount').textContent = expense.Importo === '' || expense.Importo == null
                ? '-' : formatEuro(expense.Importo) + ' \u20ac';
            card.querySelector('.expense-date span').textContent = expense.Data || '-';
            card.querySelector('.expense-description span').textContent = expense.Descrizione || '-';
            card.querySelector('.expense-note span').textContent = expense.Note || '-';
            items.appendChild(card);
        });
        if (!expenses.length) {
            status.textContent = year === ALL_YEARS ? 'Nessuna spesa.' : 'Nessuna spesa per il ' + year + '.';
            status.hidden = false;
        }
    } catch (error) {
        if (requestId !== expensesRequestId) return;
        if (error.code === 401) {
            localStorage.removeItem('token');
            showLoginView(() => showListExpenses());
            return;
        }
        items.replaceChildren();
        status.textContent = 'Caricamento spese non riuscito. ' + error.message;
        status.classList.add('expenses-error');
        status.hidden = false;
        console.error('Errore durante il caricamento delle spese:', error);
    } finally {
        if (requestId === expensesRequestId) {
            loading.hidden = true;
            view.setAttribute('aria-busy', 'false');
        }
    }
}

function showListView() {

    sessionStorage.setItem('lastActivePanel', 'lista');

    // Rimuove la classe active da tutti i pulsanti
    document.getElementById('calendarViewBtn').classList.remove('active');
    document.getElementById('graphicViewBtn').classList.remove('active');
    // Aggiunge la classe active al pulsante Lista
    document.getElementById('listViewBtn').classList.add('active');

    // Aggiorna il titolo
    updateViewHeader(titleWithYear('Prenotazioni'), 'lista');
 
    // svuota la visualizzazione precedente
    const bookingItems = document.getElementById('bookingItems');
    bookingItems.innerHTML = '';
    
    // Filtra solo le prenotazioni con Check-in o Check-out >= data odierna
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Reset ore per confronto solo data

    const filteredBookings = arrayBooking.filter(booking => {
        const checkinDate = parseItalianDate(booking['Check-in']);
        const checkoutDate = parseItalianDate(booking['Check-out']);
        return checkinDate >= today || checkoutDate >= today;
    });

    // visualizza una div per ogni prenotazione con i campi Nominativo, Check-in, Check-out, Notti, Numero Ospiti
    filteredBookings.forEach(booking => {
        const bookingDiv = document.createElement('div');
        bookingDiv.classList.add('booking-item');
           
        printCard(bookingDiv, booking);
 
        bookingItems.appendChild(bookingDiv);
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
    updateViewHeader(titleWithYear('Disponibilità'), 'calendario');
    
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
    const guests = booking['Numero Ospiti'] == null ? '' : String(booking['Numero Ospiti']).trim();
        const phone = typeof booking.Cellulare === 'string'
            ? booking.Cellulare.trim()
            : typeof booking.Cellulare === 'number' && Number.isFinite(booking.Cellulare)
                ? String(booking.Cellulare) : '';
        component.innerHTML = `
            <p><strong>${booking.Nominativo}</strong></p>
            <p><i class="fa fa-calendar-o" aria-hidden="true"></i> ${booking['Check-in']} - ${booking['Check-out']}</p>
            <p><i class="fa fa-moon-o" aria-hidden="true"></i> ${bookingNights(booking)} Notti</p>
            <p><i class="fa fa-user-o" aria-hidden="true"></i> ${guests || '-'} ${guests.length === 1 ? 'Ospiti' : ''}</p>
            <p><i class="fa fa-percent" aria-hidden="true"></i> ${booking['Tassa di soggiorno'] && booking.hasOwnProperty('Tassa di soggiorno') ? formatEuro(booking['Tassa di soggiorno']) + ' € Tassa di soggiorno' : 'No Tax'}</p>
            <p><i class="fa fa-phone" aria-hidden="true"></i> ${phone ? phone.replace(/\\/g, '+') : ' - '}</p>
            <p><i class="fa fa-sticky-note-o" aria-hidden="true"></i> ${booking.Note && booking.hasOwnProperty("Note") ? booking.Note : ' - '}</p>
            <div class="logo"><img src="./img/${booking.channel}.png" alt="Logo" width="40"></div>
            <details class="booking-actions">
                <summary class="booking-actions-toggle" aria-label="Altre azioni" title="Altre azioni">
                    <i class="fa fa-ellipsis-v" aria-hidden="true"></i>
                </summary>
                <div class="booking-actions-menu">
                    <button type="button" class="edit-action"><i class="fa fa-pencil" aria-hidden="true"></i> Modifica</button>
                    <button type="button" class="delete-action"><i class="fa fa-trash" aria-hidden="true"></i> Elimina</button>
                </div>
            </details>
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
        const deleteButton = component.querySelector('.delete-action');
        deleteButton.addEventListener('click', () => deleteBooking(booking, deleteButton));
        const editButton = component.querySelector('.edit-action');
        editButton.addEventListener('click', () => showInsertBooking(booking));
 
}

function confirmBookingDeletion(booking) {
    return new Promise(resolve => {
        const dialog = document.createElement('dialog');
        dialog.className = 'booking-delete-dialog';
        dialog.setAttribute('aria-labelledby', 'bookingDeleteTitle');
        dialog.setAttribute('aria-describedby', 'bookingDeleteDescription');
        dialog.innerHTML = `
            <form method="dialog">
                <div class="booking-delete-heading">
                    <i class="fa fa-trash" aria-hidden="true"></i>
                    <h2 id="bookingDeleteTitle">Elimina prenotazione</h2>
                </div>
                <p id="bookingDeleteDescription">Confermi la cancellazione della prenotazione di <strong></strong>?</p>
                <div class="booking-delete-buttons">
                    <button type="submit" value="cancel" autofocus>Annulla</button>
                    <button type="submit" value="delete" class="booking-delete-confirm">
                        <i class="fa fa-trash" aria-hidden="true"></i> Elimina
                    </button>
                </div>
            </form>
        `;
        dialog.querySelector('strong').textContent = booking.Nominativo;
        dialog.addEventListener('cancel', event => {
            event.preventDefault();
            dialog.close('cancel');
        });
        dialog.addEventListener('close', () => {
            const confirmed = dialog.returnValue === 'delete';
            dialog.remove();
            resolve(confirmed);
        }, { once: true });
        document.body.appendChild(dialog);
        dialog.showModal();
    });
}

async function deleteBooking(booking, button) {
    if (pendingBookingDeletions.has(booking)) return;

    const year = parseItalianDate(booking['Check-in']).getFullYear();
    pendingBookingDeletions.add(booking);
    button.disabled = true;
    try {
        if (!await confirmBookingDeletion(booking)) return;
        const response = await fetch(bookingScriptURL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                action: 'cancellaPrenotazione',
                token: localStorage.getItem('token') ?? '',
                id: Number(booking.ID)
            })
        });
        if (!response.ok) throw new Error('Risposta HTTP ' + response.status);
        const result = await response.json();
        if (result.code === 401) {
            handleUnauthorized();
            return;
        }
        if (result.status !== 'success') throw new Error(result.message || 'Cancellazione non riuscita.');

        arrayBooking = arrayBooking.filter(item => item !== booking);
        try {
            sessionStorage.setItem('arrayBooking', JSON.stringify(arrayBooking));
            sessionStorage.setItem('arrayBookingYear', String(selectedBookingYear));
            if (localStorage.getItem('arrayBookingYear') === String(year)) {
                localStorage.removeItem('arrayBooking');
                localStorage.removeItem('arrayBookingYear');
            }
        } catch (error) {
            console.error('Aggiornamento cache dopo la cancellazione non riuscito:', error);
        }

        const panel = sessionStorage.getItem('lastActivePanel');
        if (panel === 'calendario') showCalendarView();
        else if (panel === 'grafico') showGraphicView();
        else if (panel === 'lista') showListView();
    } catch (error) {
        window.alert('Cancellazione non confermata. Verifica il foglio prima di riprovare. ' + error.message);
    } finally {
        pendingBookingDeletions.delete(booking);
        button.disabled = false;
    }
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
function bookingNights(booking) {
    const checkin = parseItalianDate(booking['Check-in']);
    const checkout = parseItalianDate(booking['Check-out']);
    // Math.round compensa il cambio dell'ora legale
    const nights = Math.round((checkout - checkin) / 86400000);
    return checkin.getTime() && checkout.getTime() && nights > 0 ? nights : '-';
}

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