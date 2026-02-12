const CALENDAR_NAME = "Partite - AC";
const MAJOR_VERSION = 0;
const MINOR_VERSION = 5;
const PATCH_VERSION = 2;
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
      .addItem('Verifica Aggiornamenti', 'checkUpdates')
      .addItem('Info', 'showInfo')
      .addToUi();

  checkUpdates();
}

function getUpdateAvailable() {
  const rawUrl = githubUrl.replace("github.com", "raw.githubusercontent.com") + "/main/Code.gs?t=" + new Date().getTime();
  try {
    const response = UrlFetchApp.fetch(rawUrl);
    const content = response.getContentText();
    const mMaj = content.match(/const MAJOR_VERSION = (\d+);/);
    const mMin = content.match(/const MINOR_VERSION = (\d+);/);
    const mPat = content.match(/const PATCH_VERSION = (\d+);/);
    if (mMaj && mMin && mPat) {
      const rMaj = parseInt(mMaj[1], 10);
      const rMin = parseInt(mMin[1], 10);
      const rPat = parseInt(mPat[1], 10);
      const isNewer = (rMaj > MAJOR_VERSION) || 
                      (rMaj === MAJOR_VERSION && rMin > MINOR_VERSION) ||
                      (rMaj === MAJOR_VERSION && rMin === MINOR_VERSION && rPat > PATCH_VERSION);
      return { isNewer: isNewer, version: `${rMaj}.${rMin}.${rPat}` };
    }
  } catch (e) { console.log("Errore check: " + e); }
  return { isNewer: false };
}

function checkUpdates() {
  const updateInfo = getUpdateAvailable();
  if (updateInfo.isNewer) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('🚀 Aggiornamento Disponibile', `Nuova versione: ${updateInfo.version}\nVersione attuale: ${MAJOR_VERSION}.${MINOR_VERSION}.${PATCH_VERSION}\n\nScarica l'ultima versione da GitHub.`, ui.ButtonSet.OK);
  }
}

function checkUpdatesAutomated() {
  const props = PropertiesService.getScriptProperties();
  let count = parseInt(props.getProperty('UPDATE_CHECK_COUNT') || "0");
  count++;
  if (count >= 4) {
    const updateInfo = getUpdateAvailable();
    if (updateInfo.isNewer) {
      const myEmail = Session.getEffectiveUser().getEmail();
      const subject = "🚀 Aggiornamento disponibile per AutoCalendar (" + updateInfo.version + ")";
      const body = "Ciao!\n\nÈ disponibile una nuova versione di AutoCalendar su GitHub.\n\n" +
                   "Versione attuale: " + MAJOR_VERSION + "." + MINOR_VERSION + "." + PATCH_VERSION + "\n" +
                   "Nuova versione: " + updateInfo.version + "\n\n" +
                   "Puoi scaricare il nuovo codice da qui: " + githubUrl;
      GmailApp.sendEmail(myEmail, subject, body);
    }
    count = 0;
  }
  props.setProperty('UPDATE_CHECK_COUNT', count.toString());
}

function showInfo() {
  const anno = new Date().getFullYear();
  const versione = `${MAJOR_VERSION}.${MINOR_VERSION}.${PATCH_VERSION}`;
  const html = `<div style="font-family: sans-serif; line-height: 1.4; color: #333; text-align: center;"><h2 style="color: #1a73e8;">🏐 AutoCalendar</h2><p><b>Versione:</b> ${versione}<br><b>Creatore:</b> Matteo Checcacci</p><div style="font-size: 0.8em; background: #f9f9f9; padding: 10px; border-radius: 5px; text-align: left; margin: 10px 0;"><strong>© Info Copyright:</strong><br>• TBT, FWM, Fipav Web Manager sono copyright di TieBreakTech.<br>• FipavOnline è copyright di Manufacturing Point Software.</div><p style="font-size: 0.85em;">Prodotto concesso in <b>Licenza MIT</b>.</p><div style="margin: 15px 0;"><a href="${githubUrl}" target="_blank" style="background-color: #24292e; color: white; padding: 8px 16px; text-decoration: none; border-radius: 5px; font-weight: bold;">GitHub / Segnala Bug</a></div><hr style="border: 0; border-top: 1px solid #eee;"><p style="font-size: 0.75em; color: #666;">© ${anno} KekkoTech Softwares - RefPublic Team</p></div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(350), 'Informazioni su AutoCalendar');
}

function setUserName() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Configurazione Nome', 'Inserisci COGNOME e NOME (es: ROSSI MARIO):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) { PropertiesService.getScriptProperties().setProperty('USER_FULL_NAME', res.getResponseText().trim()); SpreadsheetApp.getUi().alert('✅ Nome salvato.'); }
}

function setGeminiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Configurazione API Gemini', 'Inserisci API Key:', ui.ButtonSet.OK_CANCEL);
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
  if (res.getSelectedButton() == ui.Button.OK) { PropertiesService.getScriptProperties().setProperty('CALENDAR_GUESTS', res.getResponseText().trim()); }
}

function setupTrigger() { setUserName(); setGeminiKey(); setFrequency(); setGuests(); getOrCreateCalendar(); SpreadsheetApp.getUi().alert('✅ Configurazione completata.'); }

function manualSync() { syncGmailToSheetAndCalendar('(from:info@tiebreaktech.com OR subject:"Designazione gara" OR subject:"VARIAZIONE GARA" OR subject:"Spostamento Gara") newer_than:30d'); SpreadsheetApp.getUi().alert('✅ Sincronizzazione terminata.'); }

function syncGmailToSheetAndCalendar(customQuery) {
  var query = customQuery || 'is:unread (from:info@tiebreaktech.com OR subject:"Designazione gara" OR subject:"VARIAZIONE GARA" OR subject:"Spostamento Gara") newer_than:30d';
  var threads = GmailApp.search(query);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arbitro") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      var htmlBody = msg.getBody();
      var from = msg.getFrom().toLowerCase();
      var subject = msg.getSubject().toUpperCase();
      var isRegionale = from.includes("tiebreaktech") || htmlBody.includes("TieBreak");
      
      var dataGara = null;
      if (subject.includes("VARIAZIONE") || subject.includes("SPOSTAMENTO")) {
        dataGara = parseSpostamento(htmlBody);
      } else {
        dataGara = isRegionale ? parseRegionaleStandard(htmlBody) : parseTerritorialeStandard(htmlBody);
      }

      if (apiKey) {
        var aiData = callGeminiAI(htmlBody, apiKey);
        if (aiData) {
          if (!dataGara) dataGara = aiData;
          else {
            if (!dataGara.ora || dataGara.ora == "") dataGara.ora = aiData.ora;
            if (!dataGara.codA || dataGara.codA == "") dataGara.codA = aiData.codA;
            if (!dataGara.codF || dataGara.codF == "") dataGara.codF = aiData.codF;
            if (!dataGara.luogo || dataGara.luogo == "") dataGara.luogo = aiData.luogo;
          }
        }
      }

      if (dataGara && dataGara.numeroGara) {
        var dataValues = sheet.getDataRange().getValues();
        var rowIndex = -1;
        for (var i = 1; i < dataValues.length; i++) {
          if (String(dataValues[i][6]).includes(String(dataGara.numeroGara))) { rowIndex = i + 1; break; }
        }
        var finalCodA = isRegionale ? "-" : (dataGara.codA || "-");
        var finalCodF = isRegionale ? "-" : (dataGara.codF || "-");
        
        if (rowIndex !== -1) {
          var oldRow = dataValues[rowIndex-1];
          var rowData = [
            dataGara.data || oldRow[0], 
            dataGara.ora || oldRow[1], 
            dataGara.luogo || oldRow[2], 
            dataGara.squadraCasa || oldRow[3], 
            dataGara.squadraOspite || oldRow[4], 
            dataGara.categoria || oldRow[5], 
            oldRow[6], 
            finalCodA, finalCodF, 
            dataGara.arb1 || oldRow[9], 
            dataGara.arb2 || oldRow[10], ""
          ];
          sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
        } else if (!subject.includes("VARIAZIONE") && !subject.includes("SPOSTAMENTO")) {
          var rowData = [dataGara.data, dataGara.ora, dataGara.luogo, dataGara.squadraCasa, dataGara.squadraOspite, dataGara.categoria, dataGara.numeroGara, finalCodA, finalCodF, dataGara.arb1 || "", dataGara.arb2 || "", ""];
          sheet.appendRow(rowData);
        }
        msg.markRead();
      }
    });
  });
  createCalendarEvents();
  checkUpdatesAutomated();
}

function parseSpostamento(html) {
  var text = cleanEmail(html);
  var res = {};
  try {
    var matchGara = text.match(/gara:\s*([A-Z0-9\s\-]+)\s*e['\s]/i);
    res.numeroGara = matchGara ? matchGara[1].trim() : null;
    var matchData = text.match(/al giorno:\s*\w+\s*(\d{2}\/\d{2}\/\d{4})/i);
    res.data = matchData ? matchData[1] : null;
    var matchOra = text.match(/ore:\s*(\d{2}\.\d{2})/i);
    res.ora = matchOra ? matchOra[1].replace('.', ':') : null;
    var lines = text.split('\n');
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].includes("ore:")) {
        if (lines[i+1] && lines[i+1].trim() !== "" && !lines[i+1].includes("Mail generata")) {
          res.luogo = lines[i+1].trim();
          break;
        }
      }
    }
    return res.numeroGara ? res : null;
  } catch (e) { return null; }
}

function parseDateCell(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/);
    if (m) return new Date(parseInt(m[3],10), parseInt(m[2],10)-1, parseInt(m[1],10));
  }
  return null;
}

function parseTimeCell(v) {
  if (v == null) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return { hh: v.getHours(), mm: v.getMinutes() };
  const m = String(v).trim().match(/^(\d{1,2})[:\.](\d{2})$/);
  return m ? { hh: parseInt(m[1],10), mm: parseInt(m[2],10) } : null;
}

function createCalendarEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Arbitro") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var cal = getOrCreateCalendar();
  var data = sheet.getDataRange().getValues();
  var yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1);
  var guests = PropertiesService.getScriptProperties().getProperty('CALENDAR_GUESTS');
  var myName = (PropertiesService.getScriptProperties().getProperty('USER_FULL_NAME') || "").toLowerCase();
  
  var excelGaraIds = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][6]) excelGaraIds.push("[GARA:" + data[i][6] + "]");
  }

  var futureLimit = new Date();
  futureLimit.setMonth(futureLimit.getMonth() + 6);
  var allEvents = cal.getEvents(yesterday, futureLimit);
  
  allEvents.forEach(ev => {
    var desc = ev.getDescription();
    var match = desc.match(/\[GARA:([^\]]+)\]/);
    if (match) {
      var tag = match[0];
      if (excelGaraIds.indexOf(tag) === -1) {
        ev.deleteEvent();
      }
    }
  });

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dOnly = parseDateCell(row[0]);
    var time = parseTimeCell(row[1]);
    if (!dOnly || !time || dOnly < yesterday) continue;

    var start = new Date(dOnly.getFullYear(), dOnly.getMonth(), dOnly.getDate(), time.hh, time.mm);
    var end = new Date(start.getTime() + 10800000);
    var title = "🏐 " + row[3] + " vs " + row[4];
    var searchTag = "[GARA:" + row[6] + "]";
    var descLines = ["N. Gara: " + row[6], "Categoria: " + row[5]];
    var arb1 = String(row[9]).trim(), arb2 = String(row[10]).trim();
    var amI1 = myName !== "" && (arb1.toLowerCase().includes(myName) || myName.includes(arb1.toLowerCase()));
    var amI2 = myName !== "" && (arb2.toLowerCase().includes(myName) || myName.includes(arb2.toLowerCase()));
    
    if (amI1) { if (arb2 && arb2 !== "-") descLines.push("Secondo arbitro: " + arb2); }
    else if (amI2) { if (arb1 && arb1 !== "-") descLines.push("Primo arbitro: " + arb1); }
    else { if (arb1 && arb1 !== "-") descLines.push("Primo arbitro: " + arb1); if (arb2 && arb2 !== "-") descLines.push("Secondo arbitro: " + arb2); }
    
    if (row[7] && row[7] !== "-" && row[7] !== "") descLines.push("Codice attivazione: " + row[7]);
    if (row[8] && row[8] !== "-" && row[8] !== "") descLines.push("Codice firma: " + row[8]);
    descLines.push("\n" + searchTag);

    var existing = cal.getEvents(new Date(dOnly.getTime() - 172800000 * 3), new Date(dOnly.getTime() + 172800000 * 30), {search: searchTag});
    if (existing.length > 0) {
      var ev = existing[0];
      if (ev.getStartTime().getTime() !== start.getTime() || ev.getLocation() !== row[2] || ev.getTitle() !== title) {
        ev.setTime(start, end); ev.setLocation(row[2]); ev.setDescription(descLines.join("\n")); ev.setTitle(title);
      }
    } else {
      var params = { location: row[2], description: descLines.join("\n") };
      if (guests && guests.trim() !== "") params.guests = guests;
      var event = cal.createEvent(title, start, end, params);
      event.removeAllReminders(); event.addPopupReminder(1440);
      var morning = new Date(dOnly.getFullYear(), dOnly.getMonth(), dOnly.getDate(), 8, 0);
      var rem = (start.getTime() - morning.getTime()) / 60000;
      if (rem > 0) event.addPopupReminder(rem);
    }
  }
}

function callGeminiAI(html, key) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`;
  const prompt = `Analizza HTML designazione o variazione volley. Ignora inoltri. Estrai dati gara. Rispondi SOLO JSON: {"data":"DD/MM/YYYY","ora":"HH:MM","luogo":"...","squadraCasa":"...","squadraOspite":"...","categoria":"...","numeroGara":"...","codA":"...","codF":"..."}. Se non trovi codA o codF lascia "". Testo: ${html}`;
  try {
    const res = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }], "generationConfig": { "response_mime_type": "application/json", "temperature": 0.1 } }), "muteHttpExceptions": true });
    return JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts[0].text);
  } catch (e) { return null; }
}

function parseRegionaleStandard(html) {
  var text = cleanEmail(html);
  var res = { codA: "-", codF: "-" };
  try {
    var matchGara = text.match(/Gara numero (\d+)/);
    if (!matchGara) return null;
    res.numeroGara = matchGara[1];
    res.data = text.match(/(\d{2}[-\/]\d{2}[-\/]\d{4})/)[1].replace(/-/g, "/");
    res.ora = text.match(/(\d{2}:\d{2})/)[1];
    var lines = text.split('\n').map(s => s.trim()).filter(s => s.length > 0);
    var capIndex = lines.findIndex(l => l.match(/\d{5}$/));
    if (capIndex > 1) {
      var p = lines[capIndex - 1].split(" - ");
      var mid = Math.ceil(p.length / 2);
      res.squadraCasa = p.slice(0, mid).join(" - ").trim();
      res.squadraOspite = p.slice(mid).join(" - ").trim();
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
    var matchGara = text.match(/gara (\d+)/i);
    if (!matchGara) return null;
    res.numeroGara = matchGara[1];
    res.data = text.match(/del (\d{2}\/\d{2}\/\d{4})/)[1];
    res.ora = text.match(/ore (\d{2}:\d{2})/i)[1];
    var lines = text.split('\n').map(s => s.trim()).filter(s => s.length > 0);
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].includes(" - ") && !lines[i].toLowerCase().includes("girone") && !lines[i].toLowerCase().includes("gara")) {
        var p = lines[i].split(" - ");
        var mid = Math.ceil(p.length / 2);
        res.squadraCasa = p.slice(0, mid).join(" - ").trim();
        res.squadraOspite = p.slice(mid).join(" - ").split(/del \d|alle|ore/i)[0].trim();
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

function extractField(text, label) {
  var safeLabel = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var m = text.match(new RegExp(safeLabel + "\\s*([^\\n\\r]+)", "i"));
  return m ? m[1].trim() : "";
}

function getOrCreateCalendar() {
  var c = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  return c.length > 0 ? c[0] : CalendarApp.createCalendar(CALENDAR_NAME, {timeZone: "Europe/Rome"});
}
