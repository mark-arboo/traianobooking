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

function bookingsCacheKey(year) {
    return 'arrayBooking_' + year;
}

function clearBookingsCache() {
    Object.keys(sessionStorage)
        .filter(key => key.startsWith('arrayBooking_'))
        .forEach(key => sessionStorage.removeItem(key));
}

function changeBookingYear(year) {
    const bookingYear = normalizeBookingYear(year);
    const cachedBookings = readCachedBookings(bookingYear);
    if (cachedBookings) {
        arrayBooking = cachedBookings;
        selectedBookingYear = bookingYear;
        sessionStorage.setItem('arrayBookingYear', String(bookingYear));
        updateBookingYearSelection();
        showListView();
        return Promise.resolve();
    }

    pendingBookingYear = bookingYear;
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
            sessionStorage.setItem(bookingsCacheKey(bookingYear), JSON.stringify(arrayBooking));
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
            const message = document.createElement('p');
            message.className = 'expenses-error';
            message.setAttribute('role', 'alert');
            message.textContent = 'Caricamento prenotazioni non riuscito. ' + error.message;
            document.getElementById('bookingItems').appendChild(message);
            console.error('Errore durante il caricamento delle prenotazioni:', error);
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

function readCachedBookings(year = selectedBookingYear) {
    try {
        const bookings = JSON.parse(sessionStorage.getItem(bookingsCacheKey(year)));
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
    updateViewHeader('Profitto', 'grafico');

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
    if (graphicContainer.classList.contains('active')) updateViewHeader('Profitto', 'grafico');
    
    // Crea l'header con navigazione
    const header = document.createElement('div');
    header.className = 'profit-month-nav';
    const monthNames = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                        'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    header.innerHTML = `
        <button type="button" class="nav-button" onclick="previousMonthGraphic()" aria-label="Mese precedente">
            <i class="fa fa-angle-left" aria-hidden="true"></i>
        </button>
        <h2>${monthNames[month]} ${year}</h2>
        <button type="button" class="nav-button" onclick="nextMonthGraphic()" aria-label="Mese successivo">
            <i class="fa fa-angle-right" aria-hidden="true"></i>
        </button>
    `;
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
    const showChannelShares = Object.values(revenueByChannel).every(amount => amount >= 0);

    // Calcola la revenue annuale
    const yearlyRevenue = calculateYearlyRevenue(year);

    statsContainer.innerHTML = `
        <section class="stats-card profit-month-card" aria-labelledby="profitMonthTitle">
            <div class="profit-month-summary">
                <div class="profit-month-heading">
                    <h3 id="profitMonthTitle" class="profit-label">Netto del mese</h3>
                    <p class="booking-count">${bookingsInMonth.length} ${bookingsInMonth.length === 1 ? 'prenotazione' : 'prenotazioni'}</p>
                </div>
                <p class="total-revenue">Caricamento...</p>
                <p class="profit-month-caption">Guadagni delle prenotazioni meno spese</p>
            </div>
            <div class="profit-month-totals">
                <div class="profit-month-income">
                    <p class="profit-label">Guadagni del mese</p>
                    <p class="profit-month-amount">${formatEuro(totalRevenue)} €</p>
                </div>
                <div class="profit-month-expenses">
                    <p class="profit-label">Spese del mese</p>
                    <p class="profit-month-amount monthly-expenses-amount">Caricamento...</p>
                </div>
            </div>
        </section>
        <section aria-labelledby="profitChannelsTitle">
            <h3 id="profitChannelsTitle" class="profit-section-title">Guadagni per canale</h3>
            <div class="stats-card profit-channels">
                ${[['booking', 'Booking.com'], ['airbnb', 'Airbnb'], ['black', 'Diretto']].map(([channel, label]) => `
                    <div class="revenue-item profit-channel-${channel}">
                        <div class="channel-info">
                            <p class="channel-name">${label}</p>
                            <p class="channel-revenue">${formatEuro(revenueByChannel[channel])} €</p>
                        </div>
                        ${showChannelShares ? `
                            <div class="profit-channel-track" aria-hidden="true">
                                <span style="width: ${totalRevenue > 0 ? revenueByChannel[channel] / totalRevenue * 100 : 0}%"></span>
                            </div>
                        ` : ''}
                    </div>
                `).join('')}
            </div>
        </section>
        <section aria-labelledby="profitYearTitle">
            <h3 id="profitYearTitle" class="profit-section-title">Il tuo anno <span>${year}</span></h3>
            <div class="stats-card yearly-card">
                <div class="yearly-revenue yearly-net">
                    <i class="fa fa-bar-chart" aria-hidden="true"></i>
                    <p class="yearly-label">Netto annuale</p>
                    <p class="yearly-amount annual-net-amount">Caricamento...</p>
                </div>
                <div class="profit-year-totals">
                    <div class="yearly-revenue">
                        <p class="yearly-label">Guadagni annuali</p>
                        <p class="yearly-amount">${formatEuro(yearlyRevenue)} €</p>
                    </div>
                    <div class="yearly-revenue yearly-expenses">
                        <p class="yearly-label">Spese annuali</p>
                        <p class="yearly-amount annual-expenses-amount">Caricamento...</p>
                    </div>
                </div>
            </div>
        </section>
        <p class="profit-footnote">Guadagni attribuiti al mese di check-in.</p>
    `;

    statsContainer.querySelector('.profit-month-card').prepend(header);
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
        statsContainer.querySelector('.annual-net-amount').textContent = formatEuro(yearlyRevenue - annualTotal) + ' \u20ac';
        statsContainer.querySelector('.total-revenue').textContent = formatEuro(totalRevenue - monthlyTotal) + ' \u20ac';
    } catch (error) {
        if (requestId !== graphicRequestId) return;
        if (error.code === 401) {
            localStorage.removeItem('token');
            showLoginView(() => showGraphicView());
            return;
        }
        statsContainer.querySelectorAll('.monthly-expenses-amount, .annual-expenses-amount, .annual-net-amount, .total-revenue')
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
    document.querySelector('.page-header h1').textContent = title;
    const yearLabel = selectedBookingYear === ALL_YEARS ? 'Tutti gli anni' : `Anno ${selectedBookingYear}`;
    const subtitle = panel === 'calendario' ? `Prenotazioni · ${yearLabel}`
        : panel === 'grafico' ? `Riepilogo economico · ${yearLabel}`
        : panel === 'inserisci' ? 'Dati del soggiorno' : yearLabel;
    document.getElementById('pageSubtitle').textContent = subtitle;
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
            <p class="booking-form-caption">I campi con * sono obbligatori</p>
            <div class="booking-form-card booking-field">
                <label for="bookingChannel">Canale</label>
                <div class="booking-channel-control">
                    <img class="booking-channel-logo" src="./img/booking.png" alt="">
                    <select id="bookingChannel" name="channel" required>
                        <option value="Booking.com">Booking.com</option>
                        <option value="Airbnb">Airbnb</option>
                        <option value="Black">Black</option>
                    </select>
                </div>
            </div>
            <section class="booking-form-section" aria-labelledby="bookingGuestTitle">
                <h2 id="bookingGuestTitle">Ospite</h2>
                <div class="booking-form-card">
                    <div class="booking-field">
                        <label for="bookingName">Nominativo</label>
                        <input id="bookingName" name="Nominativo" type="text" autocomplete="name" required>
                    </div>
                    <div class="booking-field">
                        <label for="bookingPhone">Cellulare <span class="booking-optional">Facoltativo</span></label>
                        <input id="bookingPhone" name="Cellulare" type="tel" autocomplete="tel" placeholder="Numero di telefono" title="Sono ammessi solo numeri, spazi e il carattere +">
                    </div>
                    <div class="booking-field">
                        <label for="bookingGuests">Numero ospiti</label>
                        <input id="bookingGuests" name="Numero Ospiti" type="text" aria-describedby="bookingGuestsHint" required>
                        <p id="bookingGuestsHint" class="booking-field-hint">Puoi indicare anche adulti e bambini.</p>
                    </div>
                </div>
            </section>
            <section class="booking-form-section" aria-labelledby="bookingStayTitle">
                <h2 id="bookingStayTitle">Soggiorno</h2>
                <div class="booking-form-card">
                    <div class="booking-form-pair">
                        <div class="booking-field">
                            <label for="bookingCheckin">Check-in</label>
                            <input id="bookingCheckin" name="Check-in" type="text" inputmode="numeric" placeholder="gg/mm/aaaa" pattern="[0-9]{2}/[0-9]{2}/[0-9]{4}" maxlength="10" required>
                        </div>
                        <div class="booking-field">
                            <label for="bookingCheckout">Check-out</label>
                            <input id="bookingCheckout" name="Check-out" type="text" inputmode="numeric" placeholder="gg/mm/aaaa" pattern="[0-9]{2}/[0-9]{2}/[0-9]{4}" maxlength="10" required>
                        </div>
                    </div>
                    <div id="bookingTaxGuestsField" class="booking-field">
                        <label for="bookingTaxGuests">Ospiti soggetti a tassa di soggiorno</label>
                        <input id="bookingTaxGuests" name="Ospiti soggetti a tassa di soggiorno" type="number" inputmode="numeric" min="0" step="1" required>
                    </div>
                </div>
            </section>
            <section class="booking-form-section" aria-labelledby="bookingAmountsTitle">
                <h2 id="bookingAmountsTitle">Importi</h2>
                <div class="booking-form-card booking-form-pair">
                    <div class="booking-field">
                        <label for="bookingGross">Guadagno lordo</label>
                        <div class="booking-money-control">
                            <input id="bookingGross" name="Guadagno Lordo" type="number" inputmode="decimal" min="0" step="0.01" aria-label="Guadagno lordo in euro" required>
                            <span aria-hidden="true">€</span>
                        </div>
                    </div>
                    <div id="bookingCleaningField" class="booking-field">
                        <label for="bookingCleaning">Costi di pulizia</label>
                        <div class="booking-money-control">
                            <input id="bookingCleaning" name="Costi di pulizia" type="number" inputmode="decimal" min="0" step="0.01" value="0" aria-label="Costi di pulizia in euro" required>
                            <span aria-hidden="true">€</span>
                        </div>
                    </div>
                </div>
            </section>
            <section class="booking-form-section" aria-labelledby="bookingNotesTitle">
                <h2 id="bookingNotesTitle"><label for="bookingNotes">Note</label> <span class="booking-optional">Facoltative</span></h2>
                <div class="booking-form-card booking-field">
                    <textarea id="bookingNotes" name="Note" rows="3"></textarea>
                </div>
            </section>
            <button type="submit" class="booking-submit">
                Inserisci prenotazione
            </button>
            <p id="bookingSaveStatus" role="status" hidden></p>
        </form>
    `;

    const form = document.getElementById('bookingForm');
    initializeFormFieldErrors(form);
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
        clearBookingFieldError(event.target);
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
    const channel = form.elements.namedItem('channel').value;
    const logo = form.querySelector('.booking-channel-logo');
    logo.hidden = !channel;
    if (channel) logo.src = './img/' + (channel === 'Booking.com' ? 'booking' : channel.toLowerCase()) + '.png';
    const showField = channel !== 'Black';
    [['bookingTaxGuestsField', 'Ospiti soggetti a tassa di soggiorno'], ['bookingCleaningField', 'Costi di pulizia']]
        .forEach(([containerId, name]) => {
            const field = form.elements.namedItem(name);
            document.getElementById(containerId).hidden = !showField;
            field.disabled = !showField;
            field.required = showField;
            if (!showField) {
                clearBookingFieldError(field);
                delete field.dataset.touched;
            }
        });
}

function initializeFormFieldErrors(form) {
    form.querySelectorAll('input, select, textarea').forEach(field => {
        const error = document.createElement('p');
        error.id = field.id + 'Error';
        error.className = 'booking-field-error';
        error.hidden = true;
        field.closest('.booking-field').appendChild(error);
        field.setAttribute('aria-describedby', [field.getAttribute('aria-describedby'), error.id].filter(Boolean).join(' '));
    });
}

function clearBookingFieldError(field) {
    field.setCustomValidity('');
    field.removeAttribute('aria-invalid');
    const error = document.getElementById(field.id + 'Error');
    if (error) {
        error.textContent = '';
        error.hidden = true;
    }
}

function updateBookingFieldError(field) {
    const invalid = !field.validity.valid;
    field.setAttribute('aria-invalid', String(invalid));
    const error = document.getElementById(field.id + 'Error');
    if (!invalid) {
        if (error) {
            error.textContent = '';
            error.hidden = true;
        }
        return;
    }

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
    if (error) {
        error.textContent = message;
        error.hidden = false;
    }
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
        clearBookingsCache();
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
    updateViewHeader('Nuova Spesa', 'inserisciSpesa');

    const container = document.getElementById('insertExpenseView');
    container.innerHTML = `
        <form id="expenseForm" class="booking-form">
            <p class="booking-form-caption">I campi con * sono obbligatori</p>
            <section class="booking-form-section" aria-labelledby="expenseDetailsTitle">
                <h2 id="expenseDetailsTitle">Dettagli della spesa</h2>
                <div class="booking-form-card">
                    <div class="booking-field">
                        <label for="expenseDate">Data</label>
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
                    <div class="booking-field">
                        <label for="expenseDescription">Descrizione</label>
                        <input id="expenseDescription" name="descrizione" type="text" required>
                    </div>
                </div>
            </section>
            <section class="booking-form-section" aria-labelledby="expenseAmountTitle">
                <h2 id="expenseAmountTitle">Importo</h2>
                <div class="booking-form-card booking-field">
                    <label for="expenseAmount">Totale spesa</label>
                    <div class="booking-money-control">
                        <input id="expenseAmount" name="importo" type="text" inputmode="decimal" placeholder="0,00" pattern="[0-9]+([.,][0-9]{1,2})?" aria-label="Totale spesa in euro" required>
                        <span aria-hidden="true">€</span>
                    </div>
                </div>
            </section>
            <section class="booking-form-section" aria-labelledby="expenseNotesTitle">
                <h2 id="expenseNotesTitle"><label for="expenseNotes">Note</label> <span class="booking-optional">Facoltative</span></h2>
                <div class="booking-form-card booking-field">
                    <textarea id="expenseNotes" name="note" rows="3"></textarea>
                </div>
            </section>
            <button type="submit" class="booking-submit">Inserisci spesa</button>
            <p id="expenseSaveStatus" role="status" hidden></p>
        </form>
    `;
    const form = document.getElementById('expenseForm');
    initializeFormFieldErrors(form);
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
        clearBookingFieldError(event.target);
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
                updateViewHeader('Nuova Spesa', 'inserisciSpesa');
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
    updateViewHeader('Spese', 'spese');

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
        const monthFormatter = new Intl.DateTimeFormat('it-IT', { month: 'long', year: 'numeric' });
        const dateFormatter = new Intl.DateTimeFormat('it-IT', { day: 'numeric', month: 'long', year: 'numeric' });
        const groups = new Map();
        let total = 0;
        expenses.forEach(expense => {
            const date = parseItalianDate(expense.Data);
            const parts = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(expense.Data ?? '');
            if (!parts || date.getDate() !== Number(parts[1]) ||
                date.getMonth() + 1 !== Number(parts[2]) || date.getFullYear() !== Number(parts[3])) {
                throw new Error('Data spesa non valida: ' + String(expense.Data ?? ''));
            }
            const key = `${date.getFullYear()}-${date.getMonth()}`;
            if (!groups.has(key)) groups.set(key, { date, total: 0, entries: [] });
            const group = groups.get(key);
            group.total += expense.Importo || 0;
            group.entries.push({ expense, date });
            total += expense.Importo || 0;
        });

        const summary = document.createElement('section');
        summary.className = 'expenses-summary';
        summary.setAttribute('aria-label', 'Riepilogo spese');
        summary.innerHTML = `
            <p class="expenses-summary-label">Totale spese</p>
            <p class="expenses-summary-amount">${formatEuro(total)} \u20ac</p>
            <p class="expenses-summary-caption"></p>
            <i class="fa fa-credit-card" aria-hidden="true"></i>
        `;
        summary.querySelector('.expenses-summary-caption').textContent =
            `${expenses.length} ${expenses.length === 1 ? 'spesa registrata' : 'spese registrate'} · ` +
            (year === ALL_YEARS ? 'Tutti gli anni' : `Anno ${year}`);
        items.appendChild(summary);

        groups.forEach(group => {
            const section = document.createElement('section');
            section.className = 'expenses-month';
            const heading = document.createElement('div');
            heading.className = 'expenses-month-heading';
            heading.innerHTML = '<h2></h2><p></p>';
            heading.querySelector('h2').textContent = monthFormatter.format(group.date);
            heading.querySelector('p').textContent = formatEuro(group.total) + ' \u20ac';
            section.appendChild(heading);
            group.entries.forEach(({ expense, date }) => {
                const note = String(expense.Note ?? '').trim();
                const card = document.createElement('article');
                card.className = 'expense-item';
                card.innerHTML = `
                    <div class="expense-heading">
                        <strong class="expense-category"></strong>
                        <span class="expense-amount"></span>
                    </div>
                    <p class="expense-description"></p>
                    <p class="expense-date"><i class="fa fa-calendar-o" aria-hidden="true"></i> <span></span></p>
                    ${note ? '<p class="expense-note"><i class="fa fa-sticky-note-o" aria-hidden="true"></i> <span></span></p>' : ''}
                `;
                card.querySelector('.expense-category').textContent = expense.Categoria || 'Senza categoria';
                card.querySelector('.expense-amount').textContent = expense.Importo === '' || expense.Importo == null
                    ? '-' : formatEuro(expense.Importo) + ' \u20ac';
                card.querySelector('.expense-description').textContent = expense.Descrizione || '-';
                card.querySelector('.expense-date span').textContent = dateFormatter.format(date);
                if (note) card.querySelector('.expense-note span').textContent = note;
                section.appendChild(card);
            });
            items.appendChild(section);
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
    updateViewHeader('Prenotazioni', 'lista');
 
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
    filteredBookings.forEach((booking, index) => {
        const bookingDiv = document.createElement('div');
        bookingDiv.classList.add('booking-item');
           
        printCard(bookingDiv, booking, index === 0);
 
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
        const nominativo = item.querySelector('.booking-name').textContent.toLowerCase();
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
    calendarContainer.replaceChildren();
    const hasYearData = selectedBookingYear === ALL_YEARS || currentYear === selectedBookingYear;
    const surface = document.createElement('div');
    surface.className = 'calendar-surface';
    calendarContainer.appendChild(surface);

    const header = document.createElement('div');
    header.className = 'calendar-header';
    const monthNames = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 
                        'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
    header.innerHTML = `
        <button type="button" class="nav-button" onclick="previousMonth()" aria-label="Mese precedente"><i class="fa fa-angle-left" aria-hidden="true"></i></button>
        <h2>${monthNames[currentMonth]} ${currentYear}</h2>
        <button type="button" class="nav-button" onclick="nextMonth()" aria-label="Mese successivo"><i class="fa fa-angle-right" aria-hidden="true"></i></button>
    `;
    surface.appendChild(header);

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
        const dayCell = document.createElement('button');
        dayCell.type = 'button';
        dayCell.className = 'calendar-day';
        
        const currentDate = new Date(currentYear, currentMonth, day);
        const dateString = `${String(day).padStart(2, '0')}/${String(currentMonth + 1).padStart(2, '0')}/${currentYear}`;
        
        // Verifica se è il giorno corrente
        const today = new Date();
        if (currentDate.getDate() === today.getDate() && 
            currentDate.getMonth() === today.getMonth() && 
            currentDate.getFullYear() === today.getFullYear()) {
            dayCell.classList.add('today');
            dayCell.setAttribute('aria-current', 'date');
        }

        const bookingsForDay = hasYearData ? getBookingsForDate(currentDate) : [];
        const departures = hasYearData ? getCheckOutBookingsForDate(currentDate) : [];
        const hasCheckIn = bookingsForDay.some(booking =>
            parseItalianDate(booking['Check-in']).getTime() === currentDate.getTime());
        let state;
        if (!hasYearData) {
            dayCell.classList.add('unavailable');
            dayCell.disabled = true;
            state = 'Dati non caricati';
        } else if (departures.length && hasCheckIn) {
            dayCell.classList.add('checkout-occupied');
            state = 'Check-out e check-in';
        } else if (departures.length) {
            dayCell.classList.add('checkout-free');
            state = 'Check-out';
        } else if (bookingsForDay.length > 0) {
            dayCell.classList.add('occupied');
            state = 'Occupato';
        } else {
            dayCell.classList.add('free');
            state = 'Libero';
        }

        dayCell.dataset.date = dateString;
        dayCell.setAttribute('aria-label', `${dateString}: ${state}`);
        dayCell.setAttribute('aria-pressed', 'false');
        dayCell.innerHTML = `<span class="day-number">${day}</span>`;
        dayCell.addEventListener('click', () => showDayDetails(dateString, bookingsForDay));
        
        calendarGrid.appendChild(dayCell);
    }

    surface.appendChild(calendarGrid);

    if (!hasYearData) {
        const message = document.createElement('p');
        message.className = 'calendar-data-status';
        message.setAttribute('role', 'status');
        message.textContent = 'Dati non caricati per questo anno.';
        surface.appendChild(message);
    }

    // Aggiungi la sezione per i dettagli del giorno
    const detailsSection = document.createElement('div');
    detailsSection.id = 'day-details';
    detailsSection.className = 'day-details';
    calendarContainer.appendChild(detailsSection);
}

function getBookingsForDate(date) {
    return arrayBooking.filter(booking => {
        const checkin = parseItalianDate(booking['Check-in']);
        const checkout = parseItalianDate(booking['Check-out']);
        
        // La prenotazione è attiva se la data è >= check-in e < check-out
        return date >= checkin && date < checkout;
    });
}

function getCheckOutBookingsForDate(date) {
    return arrayBooking.filter(booking => {
        const checkout = parseItalianDate(booking['Check-out']);
        return date.getTime() === checkout.getTime();
    });
}

function showDayDetails(dateString, bookings) {
    const detailsContainer = document.getElementById('day-details');
    const date = parseItalianDate(dateString);
    const departures = getCheckOutBookingsForDate(date);
    const arrivals = bookings.filter(booking =>
        parseItalianDate(booking['Check-in']).getTime() === date.getTime());
    document.querySelectorAll('#calendarView .calendar-day[data-date]').forEach(day => {
        const selected = day.dataset.date === dateString;
        day.classList.toggle('selected', selected);
        day.setAttribute('aria-pressed', String(selected));
    });

    detailsContainer.replaceChildren();
    const title = document.createElement('h3');
    title.textContent = new Intl.DateTimeFormat('it-IT', {
        weekday: 'long', day: 'numeric', month: 'long'
    }).format(date);
    detailsContainer.appendChild(title);
    const counts = [];
    if (arrivals.length) counts.push(`${arrivals.length} ${arrivals.length === 1 ? 'arrivo' : 'arrivi'}`);
    const staying = bookings.length - arrivals.length;
    if (staying) counts.push(`${staying} in soggiorno`);
    if (departures.length) counts.push(`${departures.length} ${departures.length === 1 ? 'partenza' : 'partenze'}`);
    const description = document.createElement('p');
    description.className = 'calendar-day-summary';
    description.textContent = counts.length ? counts.join(' · ') : 'Nessuna prenotazione per questo giorno.';
    detailsContainer.appendChild(description);

    [...bookings, ...departures].forEach(booking => {
        const bookingCard = document.createElement('div');
        bookingCard.className = 'booking-card';
        printCard(bookingCard, booking, false);
        const departure = departures.includes(booking);
        const arrival = arrivals.includes(booking);
        const badge = document.createElement('span');
        badge.className = 'booking-stay-status ' + (departure ? 'departure' : arrival ? 'arrival' : 'staying');
        badge.textContent = departure ? 'In partenza' : arrival ? 'In arrivo' : 'In soggiorno';
        const channel = bookingCard.querySelector('.booking-channel');
        const badges = document.createElement('span');
        badges.className = 'booking-channel-row';
        channel.replaceWith(badges);
        badges.append(channel, badge);
        detailsContainer.appendChild(bookingCard);
    });
    
    // Scroll al div dei dettagli
    detailsContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function printCard(component, booking, expanded = true) {
    const guests = booking['Numero Ospiti'] == null ? '' : String(booking['Numero Ospiti']).trim();
    const phone = typeof booking.Cellulare === 'string'
        ? booking.Cellulare.trim().replace(/\\/g, '+')
        : typeof booking.Cellulare === 'number' && Number.isFinite(booking.Cellulare)
            ? String(booking.Cellulare) : '';
    const checkin = parseItalianDate(booking['Check-in']);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const isPastBooking = checkin.getTime() > 0 && checkin < today;
    const channels = { booking: 'Booking.com', airbnb: 'Airbnb', black: 'Diretto · Black' };
    const channel = Object.prototype.hasOwnProperty.call(channels, booking.channel) ? booking.channel : 'unknown';
    const note = String(booking.Note ?? '').trim();
    component.classList.add('booking-summary-card');
    component.dataset.channel = channel;
    component.classList.toggle('booking-past', isPastBooking);
    component.innerHTML = `
        <details class="booking-disclosure" ${expanded ? 'open' : ''}>
            <summary class="booking-overview">
                <span class="booking-name"></span>
                <i class="fa fa-angle-down booking-chevron" aria-hidden="true"></i>
                <span class="booking-channel"></span>
                <span class="booking-dates">
                    <span><span class="booking-date-label">Check-in</span><span class="booking-checkin"></span></span>
                    <span class="booking-date-arrow" aria-hidden="true">→</span>
                    <span><span class="booking-date-label">Check-out</span><span class="booking-checkout"></span></span>
                    <span class="booking-nights"><i class="fa fa-moon-o" aria-hidden="true"></i> ${bookingNights(booking)} notti</span>
                </span>
            </summary>
            <div class="booking-expanded">
                <p class="booking-guests-label">Ospiti</p>
                <p class="booking-guests"></p>
                ${phone ? '<p class="booking-contact"><i class="fa fa-phone" aria-hidden="true"></i> <a></a></p>' : ''}
                ${note ? `<details class="booking-notes">
                    <summary><i class="fa fa-sticky-note-o" aria-hidden="true"></i> <span class="booking-note-preview"></span><span class="booking-note-heading">Note</span><i class="fa fa-angle-down" aria-hidden="true"></i></summary>
                    <p></p>
                </details>` : ''}
                <p class="booking-tax"><span>Tassa di soggiorno</span><span class="booking-tax-amount"></span></p>
                <div class="booking-footer">
                    <button type="button" class="pay" aria-label="Mostra importi" aria-pressed="false">
                        <span class="booking-amount-column"><span class="booking-amount-label">Lordo</span>
                            <span class="lordo" style="display:none;"></span><span class="noLordo">•••• €</span>
                        </span>
                        <span class="booking-amount-column"><span class="booking-amount-label">Netto</span>
                            <span class="netto" style="display:none;"></span><span class="noNetto">•••• €</span>
                        </span>
                    </button>
                    ${isPastBooking ? '' : `<details class="booking-actions">
                <summary class="booking-actions-toggle" aria-label="Altre azioni" title="Altre azioni">
                    <i class="fa fa-ellipsis-v" aria-hidden="true"></i>
                </summary>
                <div class="booking-actions-menu">
                    <button type="button" class="edit-action"><i class="fa fa-pencil" aria-hidden="true"></i> Modifica</button>
                    <button type="button" class="delete-action"><i class="fa fa-trash" aria-hidden="true"></i> Elimina</button>
                </div>
                    </details>`}
                </div>
            </div>
        </details>
    `;
    component.querySelector('.booking-name').textContent = booking.Nominativo || '';
    component.querySelector('.booking-channel').textContent = channels[channel] || booking.channel || 'Altro';
    const dateFormatter = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: 'short', year: 'numeric' });
    ['Check-in', 'Check-out'].forEach((field, index) => {
        const date = parseItalianDate(booking[field]);
        component.querySelector(index === 0 ? '.booking-checkin' : '.booking-checkout').textContent =
            booking[field] && Number.isFinite(date.getTime()) ? dateFormatter.format(date) : '-';
    });
    component.querySelector('.booking-guests').textContent = guests || '-';
    if (phone) {
        const link = component.querySelector('.booking-contact a');
        link.textContent = phone;
        link.href = 'tel:' + phone.replace(/[^\d+]/g, '');
    }
    if (note) {
        component.querySelector('.booking-note-preview').textContent = note;
        component.querySelector('.booking-notes p').textContent = note;
    }
    component.querySelector('.booking-tax-amount').textContent = booking['Tassa di soggiorno']
        ? formatEuro(booking['Tassa di soggiorno']) + ' €' : 'Non applicata';
    component.querySelector('.lordo').textContent = booking['Guadagno Lordo'] == null || booking['Guadagno Lordo'] === ''
        ? '-' : formatEuro(booking['Guadagno Lordo']) + ' €';
    component.querySelector('.netto').textContent = booking['Guadagno Netto'] == null || booking['Guadagno Netto'] === ''
        ? '-' : formatEuro(booking['Guadagno Netto']) + ' €';
    component.querySelector('.pay').addEventListener('click', showHidePayDetails);
    const deleteButton = component.querySelector('.delete-action');
    if (deleteButton) deleteButton.addEventListener('click', () => deleteBooking(booking, deleteButton));
    const editButton = component.querySelector('.edit-action');
    if (editButton) editButton.addEventListener('click', () => showInsertBooking(booking));
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
            // Gli altri anni (e "Tutti") possono contenere la prenotazione cancellata o spostata
            clearBookingsCache();
            sessionStorage.setItem(bookingsCacheKey(selectedBookingYear), JSON.stringify(arrayBooking));
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
    document.querySelectorAll('.pay').forEach(button => {
        button.setAttribute('aria-pressed', String(isCurrentlyHidden));
        button.setAttribute('aria-label', isCurrentlyHidden ? 'Nascondi importi' : 'Mostra importi');
    });

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