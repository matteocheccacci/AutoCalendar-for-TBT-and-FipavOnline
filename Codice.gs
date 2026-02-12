const CALENDAR_NAME = "Partite - AC";
const MAJOR_VERSION = 0;
const MINOR_VERSION = 4;
const PATCH_VERSION = 5;
const githubUrl = "https://github.com/matteocheccacci/AutoCalendar-for-TBT-and-FipavOnline";

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏐 AutoCalendar')
      .addItem('Sincronizzazione Manuale', 'manualSync')
      .addSeparator()
      .addItem('Configurazione Iniziale Completa', 'setupTrigger')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Impostazioni Singole')
        .addItem('Imposta il Tuo Nome', 'setUserName')
        .addItem('Imposta API Gemini', 'setGeminiKey')
        .addItem('Cambia Frequenza Aggiornamento', 'setFrequency')
        .addItem('Gestisci Invitati Calendario', 'setGuests'))
      .addSeparator()
      .addItem('Info', 'showInfo')
      .addToUi();
}

function showInfo() {
  const annoCorrente = new Date().getFullYear();
  const versione = `${MAJOR_VERSION}.${MINOR_VERSION}.${PATCH_VERSION}`;
  const htmlContent = `
    <div style="font-family: sans-serif; line-height: 1.5; color: #333; text-align: center;">
      <h2 style="color: #1a73e8;">🏐 AutoCalendar</h2>
      <p><b>Versione:</b> ${versione}<br>
      <b>Creatore:</b> Matteo Checcacci</p>
      <p style="font-size: 0.9em;">Il prodotto è concesso in <b>Licenza MIT</b>.</p>
      <div style="margin: 20px 0;">
        <a href="${githubUrl}" target="_blank" 
           style="background-color: #24292e; color: white; padding: 10px 20px; 
                  text-decoration: none; border-radius: 5px; font-weight: bold;">
           Vai su GitHub / Segnala Bug
        </a>
      </div>
      <hr style="border: 0; border-top: 1px solid #eee;">
      <p style="font-size: 0.8em; color: #666;">
        © ${annoCorrente} KekkoTech Softwares - RefPublic Team
      </p>
    </div>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Informazioni su AutoCalendar');
}

function setUserName() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Configurazione Nome', 'Inserisci COGNOME e NOME (es: ROSSI MARIO):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var name = res.getResponseText().trim();
    if (name !== "") {
      PropertiesService.getScriptProperties().setProperty('USER_FULL_NAME', name);
      ui.alert('✅ Nome salvato: ' + name);
    }
  }
}

function setGeminiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Configurazione API Gemini', 'Inserisci API Key (vuoto per disattivare IA):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var key = res.getResponseText().trim();
    if (key !== "") PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
    else PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
  }
}

function setFrequency() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Frequenza', 'Ogni quante ore? (1, 2, 4, 6, 8, 12):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var freq = parseInt(res.getResponseText());
    if ([1, 2, 4, 6, 8, 12].includes(freq)) {
      var triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(t => { if(t.getHandlerFunction() == 'syncGmailToSheetAndCalendar') ScriptApp.deleteTrigger(t); });
      ScriptApp.newTrigger('syncGmailToSheetAndCalendar').timeBased().everyHours(freq).create();
    }
  }
}

function setGuests() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Invitati', 'Email separate da virgola:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('CALENDAR_GUESTS', res.getResponseText().trim());
  }
}

function setupTrigger() {
  setUserName(); setGeminiKey(); setFrequency(); setGuests();
  getOrCreateCalendar();
  SpreadsheetApp.getUi().alert('✅ Configurazione Iniziale completata.');
}

function manualSync() {
  var query = '(from:info@tiebreaktech.com OR subject:"Designazione gara") newer_than:30d';
  syncGmailToSheetAndCalendar(query);
  SpreadsheetApp.getUi().alert('✅ Sincronizzazione manuale terminata.');
}

function syncGmailToSheetAndCalendar(customQuery) {
  var query = customQuery || 'is:unread (from:info@tiebreaktech.com OR subject:"Designazione gara") newer_than:30d';
  var threads = GmailApp.search(query);
  if (threads.length === 0) return;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arbitro") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var dataValues = sheet.getDataRange().getValues();
  var existingKeys = dataValues.map(row => {
    var d = (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(row[0]);
    return d + "|" + String(row[6]);
  });
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  threads.forEach(thread => {
    var messages = thread.getMessages();
    messages.forEach(msg => {
      var htmlBody = msg.getBody();
      var from = msg.getFrom().toLowerCase();
      var dataGara = null;
      var isRegionale = from.includes("tiebreaktech") || htmlBody.includes("TieBreak");
      if (apiKey) dataGara = callGeminiAI(htmlBody, apiKey);
      if (!dataGara) dataGara = isRegionale ? parseRegionaleStandard(htmlBody) : parseTerritorialeStandard(htmlBody);
      if (dataGara && dataGara.numeroGara) {
        var currentKey = dataGara.data + "|" + dataGara.numeroGara;
        if (existingKeys.indexOf(currentKey) === -1) {
          var finalCodA = isRegionale ? "-" : (dataGara.codA || "-");
          var finalCodF = isRegionale ? "-" : (dataGara.codF || "-");
          sheet.appendRow([
            dataGara.data, dataGara.ora, dataGara.luogo, 
            dataGara.squadraCasa, dataGara.squadraOspite, 
            dataGara.categoria, dataGara.numeroGara, 
            finalCodA, finalCodF, 
            dataGara.arb1 || "", dataGara.arb2 || "", ""
          ]);
          existingKeys.push(currentKey);
        }
      }
    });
  });
  createCalendarEvents();
}

function createCalendarEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arbitro") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var cal = getOrCreateCalendar();
  var data = sheet.getDataRange().getValues();
  var yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1);
  var guests = PropertiesService.getScriptProperties().getProperty('CALENDAR_GUESTS');
  var myName = (PropertiesService.getScriptProperties().getProperty('USER_FULL_NAME') || "").toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || isNaN(new Date(row[0]).getTime())) continue;
    var d = new Date(row[0]);
    if (d < yesterday) continue;
    var t = row[1].toString().split(":");
    var start = new Date(d.getFullYear(), d.getMonth(), d.getDate(), t[0], t[1]);
    var title = "🏐 " + row[3] + " vs " + row[4];
    if (cal.getEvents(start, new Date(start.getTime() + 10800000), {search: title}).length === 0) {
      var descLines = [];
      descLines.push("N. Gara: " + row[6]);
      descLines.push("Categoria: " + row[5]);
      var arb1 = String(row[9]).trim();
      var arb2 = String(row[10]).trim();
      var arb1Lower = arb1.toLowerCase();
      var arb2Lower = arb2.toLowerCase();
      var amIArb1 = myName !== "" && (arb1Lower.includes(myName) || myName.includes(arb1Lower));
      var amIArb2 = myName !== "" && (arb2Lower.includes(myName) || myName.includes(arb2Lower));
      if (amIArb1) {
        if (arb2 && arb2 !== "" && arb2 !== "-") descLines.push("Secondo arbitro: " + arb2);
      } else if (amIArb2) {
        if (arb1 && arb1 !== "" && arb1 !== "-") descLines.push("Primo arbitro: " + arb1);
      } else {
        if (arb1 && arb1 !== "" && arb1 !== "-") descLines.push("Primo arbitro: " + arb1);
        if (arb2 && arb2 !== "" && arb2 !== "-") descLines.push("Secondo arbitro: " + arb2);
      }
      if (row[7] && row[7] !== "-" && row[7] !== "") descLines.push("Codice attivazione: " + row[7]);
      if (row[8] && row[8] !== "-" && row[8] !== "") descLines.push("Codice firma: " + row[8]);
      var event = cal.createEvent(title, start, new Date(start.getTime() + 10800000), { 
        location: row[2], 
        description: descLines.join("\n"), 
        guests: guests ? guests : null 
      });
      event.removeAllReminders();
      event.addPopupReminder(1440);
      var morningOf = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 8, 0);
      var minutesToMorning = (start.getTime() - morningOf.getTime()) / 60000;
      if (minutesToMorning > 0) event.addPopupReminder(minutesToMorning);
    }
  }
}

function callGeminiAI(text, key) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`;
  const prompt = `Analizza l'HTML di questa email di designazione arbitri. Ignora date e orari di inoltro o ricezione nell'intestazione. Estrai i dati della gara effettiva. Rispondi SOLO con un JSON: {"data":"DD/MM/YYYY","ora":"HH:MM","luogo":"...","squadraCasa":"...","squadraOspite":"...","categoria":"...","numeroGara":"...","codA":"...","codF":"...","arb1":"...","arb2":"..."}. Testo HTML: ${text}`;
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }], "generationConfig": { "response_mime_type": "application/json", "temperature": 0.1 } }), "muteHttpExceptions": true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const resText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
    return JSON.parse(resText);
  } catch (e) { return null; }
}

function parseRegionaleStandard(html) {
  var text = cleanEmail(html);
  var res = { codA: "-", codF: "-" };
  try {
    res.numeroGara = text.match(/Gara numero (\d+)/)[1];
    res.data = text.match(/(\d{2}[-\/]\d{2}[-\/]\d{4})/)[1].replace(/-/g, "/");
    res.ora = text.match(/(\d{2}:\d{2})/)[1];
    var lines = text.split('\n').map(s => s.trim()).filter(s => s.length > 0);
    var capIndex = lines.findIndex(l => l.match(/\d{5}$/));
    if (capIndex > 1) {
      var teamLine = lines[capIndex - 1];
      var parts = teamLine.split(" - ");
      if (parts.length === 2) {
        res.squadraCasa = parts[0].trim();
        res.squadraOspite = parts[1].trim();
      } else {
        var mid = Math.ceil(parts.length / 2);
        res.squadraCasa = parts.slice(0, mid).join(" - ").trim();
        res.squadraOspite = parts.slice(mid).join(" - ").trim();
      }
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
    var lines = text.split('\n').map(s => s.trim()).filter(s => s.length > 0);
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].includes(" - ") && !lines[i].toLowerCase().includes("girone") && !lines[i].toLowerCase().includes("gara")) {
        var parts = lines[i].split(" - ");
        var mid = Math.ceil(parts.length / 2);
        res.squadraCasa = parts.slice(0, mid).join(" - ").trim();
        var rawOspite = parts.slice(mid).join(" - ");
        res.squadraOspite = rawOspite.split(/del \d|alle|ore/i)[0].trim();
        break;
      }
    }
    res.luogo = text.match(/Campo:(.*?)\n/i) ? text.match(/Campo:(.*?)\n/i)[1].trim() : "Vedi mail";
    res.codA = text.match(/REFERTO:\s*(\d+)/) ? text.match(/REFERTO:\s*(\d+)/)[1] : "";
    res.codF = text.match(/FIRMA REFERTO:\s*(\d+)/) ? text.match(/FIRMA REFERTO:\s*(\d+)/)[1] : "";
    res.arb1 = extractField(text, "I Arbitro:");
    res.arb2 = extractField(text, "Secondo arbitro:");
    res.categoria = text.match(/gara \d+\s+(.*?)\s+-/i) ? text.match(/gara \d+\s+(.*?)\s+-/i)[1].trim() : "Territoriale";
    return res;
  } catch(e) { return null; }
}

function cleanEmail(html) { return html.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/=3D/g, '=').replace(/=20/g, ' ').replace(/=\r?\n/g, '').replace(/&nbsp;/g, ' ').replace(/\n\s*\n/g, '\n'); }
function extractField(text, label) { var regex = new RegExp(label + "\\s*([A-Z\\s]+)", "i"); var match = text.match(regex); return match ? match[1].split('\n')[0].trim() : ""; }
function getOrCreateCalendar() { var calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME); if (calendars.length > 0) return calendars[0]; return CalendarApp.createCalendar(CALENDAR_NAME, {timeZone: "Europe/Rome"}); }
