const CALENDAR_NAME = "Partite - AC";
const MAJOR_VERSION = 0;
const MINOR_VERSION = 6;
const PATCH_VERSION = 1;
const CURRENT_VERSION = `${MAJOR_VERSION}.${MINOR_VERSION}.${PATCH_VERSION}`;
const githubUrl = "https://github.com/matteocheccacci/AutoCalendar-for-TBT-and-FipavOnline";

const WHATS_NEW_TEXT = `
  <ul>
    <li><b>Correzione BUG:</b> E' stato corretto un BUG che non consentiva di rilevare correttamente le variazioni di FWM.</li>
  </ul>`;

const updtMailBody = `
    <p>È disponibile la nuova versione <b>${CURRENT_VERSION}</b> su GitHub.</p>
    <p>Installala copiando il codice da <a href="${githubUrl}">GitHub</a>.</p>
  `;

function getRawFileUrl_(fileName) {
  const base = githubUrl.replace("https://github.com/", "https://raw.githubusercontent.com/");
  return base + "/main/" + fileName + "?t=" + new Date().getTime();
}

function fetchRemoteCodeText_() {
  const urls = [getRawFileUrl_("Codice.gs"), getRawFileUrl_("Code.gs"), getRawFileUrl_("Code.js"), getRawFileUrl_("Codice.js")];
  let lastErr = null;
  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = resp.getResponseCode();
      const text = resp.getContentText() || "";
      if (code >= 200 && code < 300 && /MAJOR_VERSION|MINOR_VERSION|PATCH_VERSION/.test(text)) {
        return { url, text, code };
      }
      lastErr = new Error("HTTP " + code + " for " + url);
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr || new Error("Impossibile scaricare le file remoto.");
}

function normalizeGaraId_(v) {
  if (v == null) return "";
  return String(v).toUpperCase().replace(/\s+/g, " ").replace(/\s*-\s*/g, " - ").trim();
}

function normalizeText_(s) {
  return String(s || "")
    .replace(/[\u2012\u2013\u2014\u2212]/g, "-")
    .replace(/\s*-\s*/g, " - ")
    .replace(/\s+/g, " ")
    .trim();
}

function capitalizeName_(name) {
  if (!name || typeof name !== 'string') return name;
  return name.toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
}

function sortSheetByDateTime_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Arbitro") || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
}

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
      .addItem('Gestisci Invitati Calendario', 'setGuests')
      .addItem('Imposta Notifiche Extra', 'setExtraReminder'))
    .addSeparator()
    .addItem('Verifica Aggiornamenti', 'checkUpdatesManual')
    .addItem('Novità', 'checkWhatsNew')
    .addItem('Info', 'showInfo')
    .addToUi();
  
  checkUpdates(false);
  Utilities.sleep(1000);
  checkWhatsNew(false);
}

function checkWhatsNew(isManual = true) {
  const props = PropertiesService.getScriptProperties();
  const lastVersion = props.getProperty('LAST_INSTALLED_VERSION');
  
  if (!isManual && lastVersion === CURRENT_VERSION) {
    return;
  }

  const html = `
    <div style="font-family: sans-serif; padding: 10px;">
      <h2 style="color: #1a73e8;">✨ Novità Versione ${CURRENT_VERSION}</h2>
      ${WHATS_NEW_TEXT}
      <p style="font-size: 0.9em; color: #666;">Se non hai tempo ora, clicca il tasto sotto per ricevere le novità via email.</p>
      <div style="margin-top: 20px; text-align: center;">
        <input type="button" value="Inviamele via Email" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).sendWhatsNewEmail();" style="padding: 10px; background-color: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer;">
        <input type="button" value="Chiudi" onclick="google.script.host.close();" style="padding: 10px; margin-left: 10px; border: 1px solid #ccc; border-radius: 4px; cursor: pointer;">
      </div>
    </div>`;
  const ui = HtmlService.createHtmlOutput(html).setWidth(420).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(ui, "AutoCalendar - Novità");
  props.setProperty('LAST_INSTALLED_VERSION', CURRENT_VERSION);
}

function sendWhatsNewEmail() {
  const email = Session.getEffectiveUser().getEmail();
  const body = `<h2>AutoCalendar - Novità Versione ${CURRENT_VERSION}</h2>${WHATS_NEW_TEXT}<br><p>Segnala un Bug su GitHub: ${githubUrl}</p>`;
  GmailApp.sendEmail(email, `AutoCalendar: Novità versione ${CURRENT_VERSION}`, "", {htmlBody: body});
}

function checkUpdatesManual() { checkUpdates(true); }

function getUpdateAvailable() {
  try {
    const remote = fetchRemoteCodeText_();
    const content = remote.text;
    const mMaj = content.match(/const MAJOR_VERSION = (\d+);/);
    const mMin = content.match(/const MINOR_VERSION = (\d+);/);
    const mPat = content.match(/const PATCH_VERSION = (\d+);/);
    if (mMaj && mMin && mPat) {
      const rMaj = parseInt(mMaj[1], 10);
      const rMin = parseInt(mMin[1], 10);
      const rPat = parseInt(mPat[1], 10);
      const isNewer = rMaj > MAJOR_VERSION || (rMaj === MAJOR_VERSION && rMin > MINOR_VERSION) || (rMaj === MAJOR_VERSION && rMin === MINOR_VERSION && rPat > PATCH_VERSION);
      return { isNewer: isNewer, version: `${rMaj}.${rMin}.${rPat}` };
    }
  } catch (e) {}
  return { isNewer: false };
}

function checkUpdates(showIfUpdated) {
  const updateInfo = getUpdateAvailable();
  const ui = SpreadsheetApp.getUi();
  if (updateInfo.isNewer) {
    ui.alert('🚀 Aggiornamento Disponibile', `Nuova versione: ${updateInfo.version}\nVersione attuale: ${CURRENT_VERSION}\n\nScarica l'ultima versione da GitHub.`, ui.ButtonSet.OK);
  } else if (showIfUpdated) {
    ui.alert('🔝 Sei aggiornato!', `Versione attuale: ${CURRENT_VERSION}`, ui.ButtonSet.OK);
  }
}

function showInfo() {
  const anno = new Date().getFullYear();
  const versione = CURRENT_VERSION;
  const html = `<div style="font-family: sans-serif; line-height: 1.4; color: #333; text-align: center;">
    <h2 style="color: #1a73e8;">🏐 AutoCalendar</h2>
    <p><b>Versione:</b> ${versione}<br><b>Creatore:</b> Matteo Checcacci</p>
    <div style="font-size: 0.8em; background: #f9f9f9; padding: 10px; border-radius: 5px; text-align: left; margin: 10px 0;">
      <strong>© Info Copyright:</strong><br>
      • TBT, FWM, Fipav Web Manager sono copyright di TieBreakTech.<br>
      • FipavOnline è copyright di Manufacturing Point Software.
    </div>
    <p style="font-size: 0.85em;">Prodotto concesso in <b>Licenza MIT</b>.</p>
    <div style="margin: 15px 0;">
      <a href="${githubUrl}" target="_blank" style="background-color: #24292e; color: white; padding: 8px 16px; text-decoration: none; border-radius: 5px; font-weight: bold;">GitHub / Segnala Bug</a>
    </div>
    <hr style="border: 0; border-top: 1px solid #eee;">
    <p style="font-size: 0.75em; color: #666;">© ${anno} KekkoTech Softwares - <a href="https://refpublic.it/" target="_blank">RefPublic</a></p>
  </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(350), 'Informazioni su AutoCalendar');
}

function setUserName() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('USER_FULL_NAME') || "Non impostato";
  var res = ui.prompt('Impostazione Nome', 'Valore attuale: ' + current + '\n\nInserisci COGNOME NOME:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var v = (res.getResponseText() || "").trim();
    if (v) PropertiesService.getScriptProperties().setProperty('USER_FULL_NAME', v);
  }
}

function setGeminiKey() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ? "Chiave presente (nascosta)" : "Non impostata";
  var res = ui.prompt('Impostazione API Gemini', 'Stato attuale: ' + current + '\n\nInserisci la nuova API Key:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var v = (res.getResponseText() || "").trim();
    if (v) PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', v);
  }
}

function setFrequency() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var currentFreq = "Non impostata";
  triggers.forEach(t => { if (t.getHandlerFunction() == 'syncGmailToSheetAndCalendar') {
    currentFreq = "Trigger attivo"; 
  }});
  var res = ui.prompt('Frequenza Aggiornamento', 'Stato: ' + currentFreq + '\n\nInserisci ore (1, 2, 4, 6, 8, 12):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var freq = parseInt((res.getResponseText() || "").trim(), 10);
    var allowed = { 1: true, 2: true, 4: true, 6: true, 8: true, 12: true };
    if (!allowed[freq]) {
      ui.alert('Valore non valido', 'Inserisci solo: 1, 2, 4, 6, 8, 12', ui.ButtonSet.OK);
      return;
    }
    triggers.forEach(t => { if (t.getHandlerFunction() == 'syncGmailToSheetAndCalendar') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('syncGmailToSheetAndCalendar').timeBased().everyHours(freq).create();
  }
}

function setGuests() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('CALENDAR_GUESTS') || "Nessuno";
  var res = ui.prompt('Gestione Invitati', 'Attuali: ' + current + '\n\nInserisci Email (separate da virgola):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('CALENDAR_GUESTS', (res.getResponseText() || "").trim());
  }
}

function setExtraReminder() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('EXTRA_REMINDER_MINUTES') || "Nessuno";
  var res = ui.prompt('Notifiche Extra Personalizzate', 'Valori attuali (minuti prima): ' + current + '\n\nInserisci i minuti di anticipo separati dalla virgola (es: 60, 120, 1440):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var val = (res.getResponseText() || "").trim();
    if (val) {
      PropertiesService.getScriptProperties().setProperty('EXTRA_REMINDER_MINUTES', val);
    }
  }
}

function setupTrigger() {
  setUserName();
  setGeminiKey();
  setFrequency();
  setGuests();
  setExtraReminder();
  getOrCreateCalendar();
  SpreadsheetApp.getUi().alert('✅ Configurazione completata.');
}

function manualSync() {
  const summary = syncGmailToSheetAndCalendar('(from:info@tiebreaktech.com OR subject:"Designazione gara" OR subject:"VARIAZIONE GARA" OR subject:"SPOSTAMENTO GARA") newer_than:30d');
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '✅ Sincronizzazione terminata',
    `Gare aggiunte a Excel: ${summary.added}\nGare aggiornate/spostate sul calendario: ${summary.moved}`,
    ui.ButtonSet.OK
  );
}

function syncGmailToSheetAndCalendar(customQuery) {
  var query = customQuery || 'is:unread (from:info@tiebreaktech.com OR subject:"Designazione gara" OR subject:"VARIAZIONE GARA" OR subject:"SPOSTAMENTO GARA") newer_than:30d';
  var threads = GmailApp.search(query);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Arbitro") || ss.getSheets()[0];
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var cal = getOrCreateCalendar();

  var addedCount = 0;
  var allMsgs = [];
  threads.forEach(thread => { thread.getMessages().forEach(m => allMsgs.push(m)); });
  allMsgs.sort((a, b) => a.getDate().getTime() - b.getDate().getTime());

  allMsgs.forEach(msg => {
    var htmlBody = msg.getBody();
    var from = (msg.getFrom() || "").toLowerCase();
    var subject = (msg.getSubject() || "").toUpperCase();
    var isRegionale = from.includes("tiebreaktech") || htmlBody.includes("TieBreak");

    var dataGara = null;
    if (subject.includes("VARIAZIONE") || subject.includes("SPOSTAMENTO") || subject.includes("AUTORIZZAZIONE")) {
      dataGara = isRegionale ? parseSpostamentoTBT(htmlBody) : parseSpostamentoFipavOnline(htmlBody);
    } else {
      dataGara = isRegionale ? parseRegionaleStandard(htmlBody) : parseTerritorialeStandard(htmlBody);
    }

    if (apiKey) {
      var aiData = callGeminiAI(htmlBody, apiKey);
      if (aiData) {
        if (!dataGara) dataGara = aiData;
        else {
          if (aiData.ora) dataGara.ora = aiData.ora;
          if (aiData.codA) dataGara.codA = aiData.codA;
          if (aiData.codF) dataGara.codF = aiData.codF;
          if (aiData.luogo) dataGara.luogo = aiData.luogo;
          if (aiData.squadraCasa) dataGara.squadraCasa = aiData.squadraCasa;
          if (aiData.squadraOspite) dataGara.squadraOspite = aiData.squadraOspite;
          if (aiData.data) dataGara.data = aiData.data;
          if (aiData.numeroGara) dataGara.numeroGara = aiData.numeroGara;
          if (aiData.categoria) dataGara.categoria = aiData.categoria;
          if (aiData.arb1) dataGara.arb1 = aiData.arb1;
          if (aiData.arb2) dataGara.arb2 = aiData.arb2;
        }
      }
    }

    if (dataGara && dataGara.numeroGara) {
      var dataValues = sheet.getDataRange().getValues();
      var rowIndex = -1;
      var targetId = normalizeGaraId_(dataGara.numeroGara);

      for (var i = 1; i < dataValues.length; i++) {
        var sheetId = normalizeGaraId_(dataValues[i][6]);
        if (sheetId && (sheetId === targetId || sheetId.includes(targetId) || targetId.includes(sheetId))) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        var oldRow = dataValues[rowIndex - 1];

        if (subject.includes("VARIAZIONE") || subject.includes("SPOSTAMENTO") || subject.includes("AUTORIZZAZIONE")) {
          deleteCalendarEventById_(cal, targetId);
        }

        var finalCodA = (dataGara.codA && dataGara.codA !== "" && dataGara.codA !== "-") ? dataGara.codA : oldRow[7];
        var finalCodF = (dataGara.codF && dataGara.codF !== "" && dataGara.codF !== "-") ? dataGara.codF : oldRow[8];

        var rowData = [
          dataGara.data || oldRow[0],
          dataGara.ora || oldRow[1],
          dataGara.luogo || oldRow[2],
          dataGara.squadraCasa || oldRow[3],
          dataGara.squadraOspite || oldRow[4],
          dataGara.categoria || oldRow[5],
          oldRow[6],
          finalCodA,
          finalCodF,
          capitalizeName_(dataGara.arb1 || oldRow[9]),
          capitalizeName_(dataGara.arb2 || oldRow[10]),
          ""
        ];
        
        var targetRange = sheet.getRange(rowIndex, 1, 1, rowData.length);
        sheet.getRange(rowIndex, 7, 1, 3).setNumberFormat("@");
        targetRange.setValues([rowData]);
        
      } else if (!subject.includes("VARIAZIONE") && !subject.includes("SPOSTAMENTO") && !subject.includes("AUTORIZZAZIONE")) {
        var rowDataNew = [
          dataGara.data, dataGara.ora, dataGara.luogo,
          dataGara.squadraCasa, dataGara.squadraOspite,
          dataGara.categoria, dataGara.numeroGara,
          isRegionale ? "-" : (dataGara.codA || ""),
          isRegionale ? "-" : (dataGara.codF || ""),
          capitalizeName_(dataGara.arb1 || ""), 
          capitalizeName_(dataGara.arb2 || ""), 
          ""
        ];
        sheet.appendRow(rowDataNew);
        sheet.getRange(sheet.getLastRow(), 7, 1, 3).setNumberFormat("@");
        addedCount++;
      }
      msg.markRead();
    }
  });

  SpreadsheetApp.flush();
  sortSheetByDateTime_();
  const movedCount = createCalendarEvents();
  return { added: addedCount, moved: movedCount };
}

function deleteCalendarEventById_(cal, garaId) {
  if (!garaId) return;
  var targetId = normalizeGaraId_(garaId);

  var now = new Date();
  var startSearch = new Date(); startSearch.setDate(now.getDate() - 60);
  var endSearch = new Date(); endSearch.setMonth(now.getMonth() + 6);

  var events = cal.getEvents(startSearch, endSearch);
  events.forEach(ev => {
    var desc = ev.getDescription() || "";
    var match = desc.match(/N\. Gara:\s*([^\n\r]+)/i) || desc.match(/\[GARA:([^\]]+)\]/);
    if (match) {
      var eventId = normalizeGaraId_(match[1]);
      if (eventId === targetId) {
        ev.deleteEvent();
      }
    }
  });
}

function createCalendarEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Arbitro") || ss.getSheets()[0];
  var cal = getOrCreateCalendar();
  var data = sheet.getDataRange().getValues();
  var extraReminder = PropertiesService.getScriptProperties().getProperty('EXTRA_REMINDER_MINUTES');

  var now = new Date();
  var yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);

  var lookbackLimit = new Date();
  lookbackLimit.setDate(lookbackLimit.getDate() - 30);

  var futureLimit = new Date();
  futureLimit.setMonth(futureLimit.getMonth() + 6);

  var allEvents = cal.getEvents(lookbackLimit, futureLimit);

  var movedCount = 0;
  var eventsMap = {};

  allEvents.forEach(ev => {
    var desc = ev.getDescription() || "";
    var match = desc.match(/N\. Gara:\s*([^\n\r]+)/i) || desc.match(/\[GARA:([^\]]+)\]/);
    if (match) {
      var id = normalizeGaraId_(match[1]);
      if (!eventsMap[id]) {
        eventsMap[id] = ev;
      } else {
        ev.deleteEvent();
      }
    }
  });

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var garaIdRaw = row[6];
    if (!garaIdRaw) continue;

    var garaIdNorm = normalizeGaraId_(garaIdRaw);
    var dOnly = parseDateCell(row[0]);
    var time = parseTimeCell(row[1]);

    if (!dOnly || !time) continue;

    if (dOnly < yesterday) {
      delete eventsMap[garaIdNorm];
      continue;
    }

    var start = new Date(dOnly.getFullYear(), dOnly.getMonth(), dOnly.getDate(), time.hh, time.mm);
    var end = new Date(start.getTime() + 9000000);
    var title = "🏐 " + row[3] + " vs " + row[4];
    var searchTag = "N. Gara: " + garaIdRaw;

    var descLines = [searchTag, "Categoria: " + row[5]];
    var arb1 = String(row[9]).trim(), arb2 = String(row[10]).trim();
    var myName = (PropertiesService.getScriptProperties().getProperty('USER_FULL_NAME') || "").toLowerCase();
    var amI1 = myName !== "" && (arb1.toLowerCase().includes(myName) || myName.includes(arb1.toLowerCase()));
    var amI2 = myName !== "" && (arb2.toLowerCase().includes(myName) || myName.includes(arb2.toLowerCase()));

    if (amI1) { if (arb2 && arb2 !== "-" && arb2 !== "") descLines.push("Secondo arbitro: " + arb2); }
    else if (amI2) { if (arb1 && arb1 !== "-" && arb1 !== "") descLines.push("Primo arbitro: " + arb1); }
    else {
      if (arb1 && arb1 !== "-" && arb1 !== "") descLines.push("Primo arbitro: " + arb1);
      if (arb2 && arb2 !== "-" && arb2 !== "") descLines.push("Secondo arbitro: " + arb2);
    }

    if (row[7] && row[7] !== "-" && row[7] !== "") descLines.push("Codice attivazione: " + row[7]);
    if (row[8] && row[8] !== "-" && row[8] !== "") descLines.push("Codice firma: " + row[8]);

    var finalDesc = descLines.join("\n");

    if (eventsMap[garaIdNorm]) {
      var ev = eventsMap[garaIdNorm];
      if (ev.getStartTime().getTime() !== start.getTime() || ev.getLocation() !== row[2] || ev.getTitle() !== title || ev.getDescription() !== finalDesc) {
        movedCount++;
        ev.setTime(start, end);
        ev.setLocation(row[2]);
        ev.setDescription(finalDesc);
        ev.setTitle(title);
        
        ev.removeAllReminders();
        ev.addPopupReminder(1440);
        var morning = new Date(dOnly.getFullYear(), dOnly.getMonth(), dOnly.getDate(), 8, 0);
        var rem = (start.getTime() - morning.getTime()) / 60000;
        if (rem > 0) ev.addPopupReminder(rem);
        
        if (extraReminder) {
          extraReminder.split(',').forEach(function(m) {
            var val = parseInt(m.trim(), 10);
            if (!isNaN(val)) ev.addPopupReminder(val);
          });
        }
      }
      delete eventsMap[garaIdNorm];
    } else {
      var guests = PropertiesService.getScriptProperties().getProperty('CALENDAR_GUESTS');
      var params = { location: row[2], description: finalDesc };
      if (guests && guests.trim() !== "") params.guests = guests;
      var event = cal.createEvent(title, start, end, params);
      event.removeAllReminders();
      event.addPopupReminder(1440);
      var morning = new Date(dOnly.getFullYear(), dOnly.getMonth(), dOnly.getDate(), 8, 0);
      var rem = (start.getTime() - morning.getTime()) / 60000;
      if (rem > 0) event.addPopupReminder(rem);
      
      if (extraReminder) {
        extraReminder.split(',').forEach(function(m) {
          var val = parseInt(m.trim(), 10);
          if (!isNaN(val)) event.addPopupReminder(val);
        });
      }
    }
  }

  for (var id in eventsMap) {
    eventsMap[id].deleteEvent();
  }

  return movedCount;
}

function splitTeamsSmart_(teamLineRaw) {
  var teamLine = normalizeText_(teamLineRaw || "");
  if (!teamLine || teamLine.indexOf(" - ") === -1) return null;

  var parts = teamLine.split(" - ").map(p => normalizeText_(p)).filter(p => p);
  if (parts.length < 2) return null;
  if (parts.length === 2) return { casa: parts[0], ospite: parts[1] };

  var best = null;
  for (var k = 1; k < parts.length; k++) {
    var left = parts.slice(0, k).join(" - ").trim();
    var right = parts.slice(k).join(" - ").trim();
    if (!left || !right) continue;

    var score = 0;
    var leftTokens = left.split(" ").filter(Boolean).length;
    var rightTokens = right.split(" ").filter(Boolean).length;
    score += Math.abs(leftTokens - rightTokens) * 5;
    score += Math.abs(left.length - right.length) * 0.5;

    if (/^[A-Z\s\d]+$/.test(parts[k-1]) && /[a-z]/.test(parts[k])) {
      score += 150; 
    }

    var mid = Math.floor(parts.length / 2);
    score += Math.abs(k - mid) * 10;

    if (best == null || score < best.score) best = { score: score, casa: left, ospite: right };
  }

  return best ? { casa: best.casa, ospite: best.ospite } : null;
}

function parseTerritorialeStandard(html) {
  var text = cleanEmail(html);
  var res = {};
  try {
    var matchGara = text.match(/gara (\d+)/i);
    if (!matchGara) return null;
    res.numeroGara = matchGara[1];

    var mData = text.match(/del (\d{2}\/\d{2}\/\d{4})/);
    if (mData) res.data = mData[1];

    var mOra = text.match(/ore (\d{1,2}[:\.]\d{2})/i);
    if (mOra) res.ora = mOra[1].replace(".", ":");

    var lines = text.split('\n').map(s => normalizeText_(s)).filter(s => s.length > 0);
    for (var i = 0; i < lines.length; i++) {
      var l = lines[i];
      if (l.includes(" - ") && !l.toLowerCase().includes("girone") && !l.toLowerCase().includes("gara") && !l.toLowerCase().includes("campo:")) {
        var teamLine = l.split(/del \d|alle|ore/i)[0].trim();
        var pair = splitTeamsSmart_(teamLine);
        if (pair) {
          res.squadraCasa = pair.casa;
          res.squadraOspite = pair.ospite;
          break;
        }
      }
    }

    var mCampo = text.match(/Campo:(.*?)\n/i);
    res.luogo = mCampo ? normalizeText_(mCampo[1]) : "Vedi mail";

    res.codA = text.match(/CODICE ATTIVAZIONE REFERTO:\s*(\d+)/i) ? text.match(/CODICE ATTIVAZIONE REFERTO:\s*(\d+)/i)[1] : (text.match(/REFERTO:\s*(\d+)/) ? text.match(/REFERTO:\s*(\d+)/)[1] : "");
    res.codF = text.match(/CODICE FIRMA REFERTO:\s*(\d+)/i) ? text.match(/CODICE FIRMA REFERTO:\s*(\d+)/i)[1] : (text.match(/FIRMA REFERTO:\s*(\d+)/) ? text.match(/FIRMA REFERTO:\s*(\d+)/)[1] : "");

    res.arb1 = extractField(text, "Primo arbitro:") || extractField(text, "I Arbitro:");
    res.arb2 = extractField(text, "II Arbitro:") || extractField(text, "Secondo arbitro:");

    res.categoria = text.match(/gara \d+\s+(.*?)\s+-/i) ? normalizeText_(text.match(/gara \d+\s+(.*?)\s+-/i)[1]) : "Territoriale";
    return res;
  } catch (e) { return null; }
}

function parseRegionaleStandard(html) {
  var text = cleanEmail(html);
  var res = { codA: "-", codF: "-" };
  try {
    var matchGara = text.match(/Gara numero (\d+)/i);
    if (!matchGara) return null;
    res.numeroGara = matchGara[1];

    var mData = text.match(/(\d{2}\s*[-\/]\s*\d{2}\s*[-\/]\s*\d{4})/);
    if (mData) {
      res.data = mData[1].replace(/\s/g, "").replace(/-/g, "/");
    }

    var mOra = text.match(/(\d{2}[:\.]\d{2})/);
    if (mOra) res.ora = mOra[1].replace(".", ":");

    var lines = text.split('\n').map(s => normalizeText_(s)).filter(s => s.length > 0);
    var capIndex = lines.findIndex(l => l.match(/\d{5}$/));
    if (capIndex > 1) {
      var teamLine = lines[capIndex - 1];
      var pair = splitTeamsSmart_(teamLine);
      if (pair) {
        res.squadraCasa = pair.casa;
        res.squadraOspite = pair.ospite;
      }
      res.luogo = lines[capIndex];
    }

    var mCat = text.match(/Gara numero \d+\s+(.*?)\s+del/i);
    if (mCat) res.categoria = normalizeText_(mCat[1]);

    res.arb1 = extractField(text, "1° arbitro:");
    res.arb2 = extractField(text, "2° arbitro:");
    return res;
  } catch (e) { return null; }
}

function parseSpostamentoFipavOnline(html) {
  var text = cleanEmail(html);
  var res = {};
  try {
    var matchGara = text.match(/gara:\s*([A-Z0-9\s\-]+)\s*e['\s]/i);
    res.numeroGara = matchGara ? normalizeText_(matchGara[1]) : null;

    var matchData = text.match(/al giorno:\s*\w+\s*(\d{2}\/\d{2}\/\d{4})/i);
    res.data = matchData ? matchData[1] : null;

    var matchOra = text.match(/ore:\s*(\d{1,2}[.:]\d{2})/i);
    res.ora = matchOra ? matchOra[1].replace('.', ':') : null;

    var lines = text.split('\n').map(s => normalizeText_(s));
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].toLowerCase().includes("ore:")) {
        if (lines[i + 1] && lines[i + 1].trim() !== "" && !lines[i + 1].toLowerCase().includes("mail generata")) {
          res.luogo = lines[i + 1].trim();
          break;
        }
      }
    }
    return res.numeroGara ? res : null;
  } catch (e) { return null; }
}

function parseSpostamentoTBT(html) {
  var text = cleanEmail(html);
  var res = {};
  try {
    // 1. Estrazione Numero Gara
    var matchGara = text.match(/modifiche alla gara n\.\s*(\d+)/i) || 
                    text.match(/Numero:\s*(\d+)/i) ||
                    text.match(/gara n\.\s*(\d+)/i);
    if (!matchGara) return null;
    res.numeroGara = matchGara[1];

    // 2. Estrazione Data e Ora Attuale (formato YYYY-MM-DD o DD/MM/YYYY)
    // Cerchiamo la riga "Data attuale:"
    var matchDataOra = text.match(/Data attuale:\s*(\d{4})-(\d{2})-(\d{2})\s*(\d{2}:\d{2})/i) ||
                       text.match(/Data attuale:\s*(\d{2}\/\d{2}\/\d{4})\s*(\d{2}:\d{2})/i);
    
    if (matchDataOra) {
      // Se il primo gruppo catturato ha 4 cifre è anno (YYYY-MM-DD), altrimenti è già DD/MM/YYYY
      if (matchDataOra[1].length === 4) {
        res.data = matchDataOra[3] + "/" + matchDataOra[2] + "/" + matchDataOra[1];
        res.ora = matchDataOra[4];
      } else {
        res.data = matchDataOra[1];
        res.ora = matchDataOra[2];
      }
    }

    // 3. Estrazione Impianto Attuale
    var matchImpianto = text.match(/Impianto attuale:\s*([^\n\r]+)/i);
    if (matchImpianto) {
      res.luogo = normalizeText_(matchImpianto[1]);
    }

    return res.numeroGara ? res : null;
  } catch (e) {
    return null;
  }
}

function cleanEmail(html) {
  var s = String(html || "")
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/=C3=A8/g, 'è') // Fix codifica Quoted-Printable comune
    .replace(/=3D/g, '=')
    .replace(/=20/g, ' ')
    .replace(/=\r?\n/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/\r/g, '\n')
    .replace(/\n\s*\n/g, '\n');
  s = s.split('\n').map(line => normalizeText_(line)).join('\n');
  return s;
}

function extractField(text, label) {
  var safeLabel = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var m = text.match(new RegExp(safeLabel + "\\s*([^\\n\\r]+)", "i"));
  return m ? normalizeText_(m[1]) : "";
}

function parseDateCell(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/);
    if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  }
  return null;
}

function parseTimeCell(v) {
  if (v == null) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return { hh: v.getHours(), mm: v.getMinutes() };
  const m = String(v).trim().match(/^(\d{1,2})[:\.](\d{2})$/);
  if (!m) return null;
  return { hh: parseInt(m[1], 10), mm: parseInt(m[2], 10) };
}

function getOrCreateCalendar() {
  var c = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  return c.length > 0 ? c[0] : CalendarApp.createCalendar(CALENDAR_NAME, { timeZone: "Europe/Rome" });
}

function callGeminiAI(html, key) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`;
  const prompt = `Analizza HTML designazione volley. Estrai dati gara. Rispondi SOLO JSON: {"data":"DD/MM/YYYY","ora":"HH:MM","luogo":"...","squadraCasa":"...","squadraOspite":"...","categoria":"...","numeroGara":"...","codA":"...","codF":"...","arb1":"...","arb2":"..."}. Testo: ${html}`;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { response_mime_type: "application/json", temperature: 0.1 } }),
      muteHttpExceptions: true
    });
    return JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts[0].text);
  } catch (e) { return null; }
}

function checkUpdatesAutomated() {
  const props = PropertiesService.getScriptProperties();
  let count = parseInt(props.getProperty('UPDATE_CHECK_COUNT') || "0", 10);
  count++;
  if (count >= 4) {
    const updateInfo = getUpdateAvailable();
    if (updateInfo.isNewer) {
      const myEmail = Session.getEffectiveUser().getEmail();
      if (myEmail) {
      GmailApp.sendEmail(myEmail, "🚀 Aggiornamento AutoCalendar", "", {htmlBody: updtMailBody});
    }
    count = 0;
  }
  props.setProperty('UPDATE_CHECK_COUNT', count.toString());
}
}
