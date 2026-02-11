/**
 * SISTEMA AUTOMATICO DI SINCRONIZZAZIONE GARE 
 * Versione: 0.3.1
 */

const CALENDAR_NAME = "Partite - AC";

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏐 AutoCalendar')
      .addItem('Sincronizzazione MANUALE (Tutto)', 'manualSync')
      .addItem('Sincronizzazione AUTOMATICA (Solo nuove)', 'syncGmailToSheetAndCalendar')
      .addSeparator()
      .addItem('Configura API Gemini e Trigger', 'setupTrigger')
      .addToUi();
}

/**
 * LOGICA DI CONFIGURAZIONE INIZIALE
 */
function setupTrigger() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Richiesta API Key per Gemini
  var keyResponse = ui.prompt(
    'Configurazione API Gemini',
    'Inserisci la tua API Key di Gemini (lascia vuoto per non usare l\'IA nel selezionare le gare):',
    ui.ButtonSet.OK_CANCEL
  );

  if (keyResponse.getSelectedButton() == ui.Button.OK) {
    var newKey = keyResponse.getResponseText().trim();
    if (newKey !== "") {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', newKey);
      ui.alert('✅ API Key salvata. Il controllo AI è attivo.');
    } else {
      PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
      ui.alert('⚠️ API Key non inserita. Verrà utilizzato il parser standard.');
    }
  }

  // 2. Setup Trigger Temporale (Ogni 2 ore)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'syncGmailToSheetAndCalendar') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('syncGmailToSheetAndCalendar').timeBased().everyHours(2).create();
  
  getOrCreateCalendar();
  ui.alert('✅ Automazione configurata (ogni 2 ore) sul calendario "' + CALENDAR_NAME + '".');
}

/**
 * SINCRONIZZAZIONE
 */
function manualSync() {
  syncGmailToSheetAndCalendar('(from:info@tiebreaktech.com OR subject:"Designazione gara") newer_than:30d');
}

function syncGmailToSheetAndCalendar(customQuery) {
  var query = customQuery || 'is:unread (from:info@tiebreaktech.com OR subject:"Designazione gara") newer_than:30d';
  var threads = GmailApp.search(query);
  if (threads.length === 0) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arbitro") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var existingIds = sheet.getRange(1, 7, sheet.getLastRow() || 1, 1).getValues().flat().map(String);
  
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  threads.forEach(thread => {
    var msg = thread.getMessages()[thread.getMessageCount() - 1];
    var htmlBody = msg.getBody();
    var plainBody = msg.getPlainBody();
    var from = msg.getFrom().toLowerCase();
    var dataGara = null;

    if (apiKey) {
      dataGara = callGeminiAI(plainBody, apiKey);
    } 
    
    if (!dataGara) {
      dataGara = (from.includes("tiebreaktech") || htmlBody.includes("TieBreak")) ? 
                 parseRegionaleStandard(htmlBody) : parseTerritorialeStandard(htmlBody);
    }

    if (dataGara && dataGara.numeroGara && existingIds.indexOf(String(dataGara.numeroGara)) === -1) {
      sheet.appendRow([
        dataGara.data, dataGara.ora, dataGara.luogo, 
        dataGara.squadraCasa, dataGara.squadraOspite, 
        dataGara.categoria, dataGara.numeroGara, 
        dataGara.codA || "", dataGara.codF || "", dataGara.arb1 || "", dataGara.arb2 || "", ""
      ]);
      existingIds.push(String(dataGara.numeroGara));
    }
  });

  createCalendarEvents();
}

/**
 * MOTORE AI GEMINI
 */
function callGeminiAI(text, key) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`;
  const prompt = `Analizza questa email di designazione pallavolo. Estrai i dati e verifica che i nomi delle squadre NON siano nomi di dirigenti o indirizzi. Rispondi SOLO con un JSON valido: {"data":"DD/MM/YYYY","ora":"HH:MM","luogo":"...","squadraCasa":"...","squadraOspite":"...","categoria":"...","numeroGara":"...","codA":"...","codF":"...","arb1":"...","arb2":"..."}. Testo: ${text}`;

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": { "response_mime_type": "application/json", "temperature": 0.1 }
    }),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    return JSON.parse(json.candidates[0].content.parts[0].text);
  } catch (e) { return null; }
}

// --- PARSER STANDARD ---
function parseRegionaleStandard(html) {
  var text = cleanEmail(html);
  var res = { codA: "", codF: "" };
  try {
    res.numeroGara = text.match(/Gara numero (\d+)/)[1];
    res.data = text.match(/(\d{2}[-\/]\d{2}[-\/]\d{4})/)[1].replace(/-/g, "/");
    res.ora = text.match(/(\d{2}:\d{2})/)[1];
    var lines = text.split('\n').map(s => s.trim()).filter(s => s.length > 0);
    var capIndex = lines.findIndex(l => l.match(/\d{5}$/));
    if (capIndex > 1) {
      var teamLine = lines[capIndex - 1];
      var lastDash = teamLine.lastIndexOf(" - ");
      res.squadraCasa = teamLine.substring(0, lastDash).trim();
      res.squadraOspite = teamLine.substring(lastDash + 3).trim();
      res.luogo = lines[capIndex];
    }
    res.categoria = text.match(/Gara numero \d+\s+(.*?)\s+del/i)[1].trim();
    res.arb1 = extractField(text, "1° arbitro:");
    res.arb2 = extractField(text, "2° arbitro:");
    return res;
  } catch(e) { return null; }
}

function parseTerritorialeStandard(html) {
  var text = cleanEmail(html);
  var res = {};
  try {
    res.numeroGara = text.match(/gara (\d+)/i)[1];
    res.data = text.match(/del (\d{2}\/\d{2}\/\d{4})/)[1];
    res.ora = text.match(/ore (\d{2}:\d{2})/i)[1];
    var matchTeams = text.match(/Girone.*?\n(.*?)\s+del\s+\d{2}/i) || text.match(/Girone.*?\s+(.*?)\s+del\s+\d{2}/i);
    if (matchTeams) {
      var fullString = matchTeams[1].replace(/\n/g, " ");
      var lastDash = fullString.lastIndexOf(" - ");
      res.squadraCasa = fullString.substring(0, lastDash).trim();
      res.squadraOspite = fullString.substring(lastDash + 3).trim();
    }
    res.luogo = text.match(/Campo:(.*?)\n/i)[1].trim();
    res.codA = text.match(/REFERTO:\s*(\d+)/)[1];
    res.codF = text.match(/FIRMA REFERTO:\s*(\d+)/)[1];
    res.arb1 = extractField(text, "I Arbitro:");
    res.arb2 = extractField(text, "Secondo arbitro:");
    res.categoria = text.match(/gara \d+\s+(.*?)\s+-/i)[1].trim();
    return res;
  } catch(e) { return null; }
}

function cleanEmail(html) {
  return html.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/=3D/g, '=').replace(/=20/g, ' ').replace(/=\r?\n/g, '').replace(/&nbsp;/g, ' ').replace(/\n\s*\n/g, '\n'); 
}

function extractField(text, label) {
  var regex = new RegExp(label + "\\s*([A-Z\\s]+)", "i");
  var match = text.match(regex);
  return match ? match[1].split('\n')[0].trim() : "";
}

function getOrCreateCalendar() {
  var calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (calendars.length > 0) return calendars[0];
  return CalendarApp.createCalendar(CALENDAR_NAME, {timeZone: "Europe/Rome"});
}

function createCalendarEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arbitro") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var cal = getOrCreateCalendar();
  var data = sheet.getDataRange().getValues();
  var yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1);
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || isNaN(new Date(row[0]).getTime())) continue;
    var d = new Date(row[0]);
    if (d < yesterday) continue;
    
    var t = row[1].toString().split(":");
    var start = new Date(d.getFullYear(), d.getMonth(), d.getDate(), t[0], t[1]);
    var title = "🏐 " + row[3] + " vs " + row[4];
    
    if (cal.getEvents(start, new Date(start.getTime() + 10800000), {search: title}).length === 0) {
      var desc = "Gara: " + row[6] + "\nCod. Att: " + row[7] + " | Cod. Fir: " + row[8];
      cal.createEvent(title, start, new Date(start.getTime() + 10800000), {location: row[2], description: desc});
    }
  }
}
