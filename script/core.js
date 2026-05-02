let arrayBooking = [];
let currentMonth = new Date().getMonth();
let currentYear = new Date().getFullYear();

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
 
    // legge i primi 3 sheet di un file excel letto da una URL 
    const url = 'https://docs.google.com/spreadsheets/d/1eZ2t1dVZqAiZTflLigA9y8sHzbXSVdISgmlCt8MeOyk/export?format=xlsx';
    fetch(url)
        .then(response => response.arrayBuffer())
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

            sessionStorage.setItem('arrayBooking', JSON.stringify(arrayBooking));
    
            showListView(); // Mostra la vista lista all'avvio dell'app
        })
        .catch(error => console.error('Errore durante il caricamento del file Excel:', error));

}

/**
 * Logica eseguita durante il refresh
 */
function onPageRefresh() {
    // Qui puoi aggiungere logica per il refresh
    // Ad esempio: ripristinare stato, saltare intro, etc.
    
    // Controlla se c'era un pannello salvato
    const lastPanel = sessionStorage.getItem('lastActivePanel');

    if (sessionStorage.getItem('arrayBooking')) {
        arrayBooking = JSON.parse(sessionStorage.getItem('arrayBooking'));
    }
    
    if (lastPanel) {
        switch(lastPanel) {
            case 'calendario':
                showCalendarView();
                break;
            case 'lista':
                showListView();
                break;
            case 'grafico':
                showGraphicView();
                break;    
            default:
                showListView();
        }
    } else {
        showListView();
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
    // Rimuove la classe active da tutti i pulsanti
    document.getElementById('listViewBtn').classList.remove('active');
    document.getElementById('calendarViewBtn').classList.remove('active');
    // Aggiunge la classe active al pulsante Grafico
    document.getElementById('graphicViewBtn').classList.add('active');

    // Mostra la vista grafico e nasconde le altre viste
    document.getElementById('graphicView').classList.add('active');
    document.getElementById('listView').classList.remove('active');
    document.getElementById('calendarView').classList.remove('active');

    sessionStorage.setItem('lastActivePanel', 'grafico');
}

function showListView() {

    // Rimuove la classe active da tutti i pulsanti
    document.getElementById('calendarViewBtn').classList.remove('active');
    document.getElementById('graphicViewBtn').classList.remove('active');
    // Aggiunge la classe active al pulsante Lista
    document.getElementById('listViewBtn').classList.add('active');
 
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

    // inserisce una text per la ricerca rapida del nominativo, posizionata in alto prima delle card.
    // La ricerca deve partire ad ogni lettera digitata e deve filtrare le card in base al nominativo
    listView.innerHTML = `
        <input type="text" class="search-input" placeholder="Cerca prenotazione" onkeyup="filterBookings(this)">
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

    sessionStorage.setItem('lastActivePanel', 'lista');
    
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
    // Rimuove la classe active da tutti i pulsanti
    document.getElementById('listViewBtn').classList.remove('active');
    document.getElementById('graphicViewBtn').classList.remove('active');
    // Aggiunge la classe active al pulsante Calendario
    document.getElementById('calendarViewBtn').classList.add('active');
    
    renderCalendar();

    // Mostra la vista calendario e nasconde le altre viste
    document.getElementById('calendarView').classList.add('active');
    document.getElementById('listView').classList.remove('active');
    document.getElementById('graphicView').classList.remove('active');

    sessionStorage.setItem('lastActivePanel', 'calendario');
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
            <p><i class="fa fa-sticky-note-o" aria-hidden="true"></i> ${booking.Note && booking.hasOwnProperty("Note") ? booking.Note : ' - '}</p>
            <div class="logo"><img src="./img/${booking.channel}.png" alt="Logo" width="40"></div>
            <div class="pay">
                <span class="lordo">${booking['Guadagno Lordo'] && booking.hasOwnProperty('Guadagno Lordo') ? formatEuro(booking['Guadagno Lordo']) + ' € /' : ''}</span>
                <span class="netto">${booking['Guadagno Netto'] && booking.hasOwnProperty('Guadagno Netto') ? formatEuro(booking['Guadagno Netto']) + ' €' : ''}</span>
                <span class="noLordo" style="display:none;"> -- /</span>
                <span class="noNetto" style="display:none;"> -- €</span>
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