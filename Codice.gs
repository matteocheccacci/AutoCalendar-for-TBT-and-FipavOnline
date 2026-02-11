/**
 * SISTEMA AUTOMATICO DI SINCRONIZZAZIONE GARE
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏐 AutoCalendar')
      .addItem('Scarica Nuove Gare da Gmail', 'syncGmailToSheetAndCalendar')
      .addItem('Sincronizza solo Calendario', 'createCalendarEvents')
      .addSeparator()
      .addItem('Configura Automazione (Trigger)', 'setupTrigger')
      .addToUi();
}

// --- LOGICA CALENDARIO DEDICATO ---

/**
 * Cerca il calendario "Partite - AC". Se non esiste, lo crea.
 */
function getOrCreateCalendar() {
  var calName = "Partite - AC";
  var calendars = CalendarApp.getCalendarsByName(calName);
  
  if (calendars.length > 0) {
    return calendars[0];
  } else {
    Logger.log("Calendario non trovato. Creazione in corso...");
    return CalendarApp.createCalendar(calName, {
      summary: "Calendario sincronizzato automaticamente per le designazioni arbitrali.",
      timeZone: "Europe/Rome"
    });
  }
}

// --- LOGICA DI CONFIGURAZIONE AUTOMATICA ---

function setupTrigger() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Configurazione Automazione',
    'Ogni quante ore vuoi che lo script controlli la posta? (Inserisci un numero da 1 a 24)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    var ore = parseInt(response.getResponseText());
    if (isNaN(ore) || ore < 1 || ore > 24) {
      ui.alert('Errore: inserisci un numero valido tra 1 e 24.');
      return;
    }

    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() == 'syncGmailToSheetAndCalendar') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    ScriptApp.newTrigger('syncGmailToSheetAndCalendar')
        .timeBased()
        .everyHours(ore)
        .create();

    // Inizializza il calendario al primo setup per sicurezza
    getOrCreateCalendar();

    ui.alert('✅ Automazione configurata! Il calendario "Partite - AC" è pronto e verrà controllato ogni ' + ore + ' ore.');
  }
}

// --- FUNZIONE DI SINCRONIZZAZIONE GMAIL ---

function syncGmailToSheetAndCalendar() {
  var threads = [];
  threads = threads.concat(GmailApp.search('from:info@tiebreaktech.com is:unread'));
  threads = threads.concat(GmailApp.search('from:designazioni.varese@federvolley.it is:unread'));
  
  if (threads.length === 0) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arbitro") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var count = 0;

  threads.forEach(thread => {
    var msg = thread.getMessages()[0];
    var body = msg.getPlainBody();
    var from = msg.getFrom();
    var dataGara = from.includes("tiebreaktech") ? parseTieBreak(body) : parseFipavVarese(body);

    if (dataGara && dataGara.numeroGara) {
      sheet.appendRow([
        dataGara.data, dataGara.ora, dataGara.luogo, 
        dataGara.squadraCasa, dataGara.squadraOspite, 
        dataGara.categoria, dataGara.numeroGara, 
        dataGara.codAttivazione || "", dataGara.codFirma || "",
        dataGara.primoArbitro, dataGara.secondoArbitro, "", "", dataGara.fase || ""
      ]);
      thread.markRead();
      count++;
    }
  });

  if(count > 0) createCalendarEvents();
}

// --- LOGICA CALENDARIO ---

function createCalendarEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arbitro") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  
  // Utilizza la nuova funzione per ottenere il calendario corretto
  var cal = getOrCreateCalendar();
  
  var startRow = findStartRow(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 14).getValues();
  data.forEach(row => {
    var dateInput = new Date(row[0]);
    var timeStr = row[1].toString();
    if (isNaN(dateInput.getTime()) || !timeStr.includes(':')) return;
    
    var t = timeStr.split(":");
    var start = new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate(), t[0], t[1]);
    var end = new Date(start.getTime() + (3 * 60 * 60 * 1000));
    var title = "🏐 Partita: " + row[3] + " / " + row[4];
    
    if (cal.getEvents(start, end, {search: title}).length === 0) {
      var desc = "Gara n: " + row[6] + "\nCategoria: " + row[5] + "\nCod. Attivazione: " + row[7] + "\nCod. Firma: " + row[8];
      cal.createEvent(title, start, end, {location: row[2], description: desc}).addPopupReminder(1440);
    }
  });
}

// --- UTILITY E PARSER ---

function findStartRow(sheet) {
  var dates = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  for (var r = dates.length - 1; r >= 0; r--) {
    var d = new Date(dates[r][0]);
    if (!isNaN(d.getTime()) && d < yesterday) return r + 2;
  }
  return 1;
}

function parseTieBreak(text) {
  try {
    return {
      numeroGara: text.match(/Gara numero (\d+)/)[1],
      categoria: text.match(/Gara numero \d+ (.*?) del/)[1],
      data: text.match(/del (\d{2}-\d{2}-\d{4})/)[1].replace(/-/g, "/"),
      ora: text.match(/(\d{2}:\d{2})/)[1],
      squadraCasa: text.match(/([^-]+) - /)[1].trim(),
      squadraOspite: text.match(/ - ([^\n]+)/)[1].split("\n")[0].trim(),
      luogo: text.match(/([A-Z\s]+ - [A-Z\s\d,]+)\n/)[1].trim(),
      primoArbitro: "Io", secondoArbitro: ""
    };
  } catch(e) { return null; }
}

function parseFipavVarese(text) {
  try {
    return {
      numeroGara: text.match(/gara (\d+)/)[1],
      categoria: text.match(/\d+ (.*?) -/)[1].trim(),
      data: text.match(/del (\d{2}\/\d{2}\/\d{4})/)[1],
      ora: text.match(/ore (\d{2}:\d{2})/)[1],
      squadraCasa: text.match(/B\n(.*?)-/s)[1].trim(),
      squadraOspite: text.match(/-(.*?)del/s)[1].trim(),
      luogo: text.match(/Campo: (.*?)\n/)[1].trim(),
      codAttivazione: text.match(/REFERTO: (\d+)/)[1],
      codFirma: text.match(/FIRMA REFERTO: (\d+)/)[1],
      primoArbitro: "Io", secondoArbitro: ""
    };
  } catch(e) { return null; }
}
