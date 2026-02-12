const CALENDAR_NAME = "Partite - AC";
const MAJOR_VERSION = 0;
const MINOR_VERSION = 5;
const PATCH_VERSION = 5;
const githubUrl = "https://github.com/matteocheccacci/AutoCalendar-for-TBT-and-FipavOnline";

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
  throw lastErr || new Error("Impossibile scaricare il file remoto.");
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
    .addItem('Verifica Aggiornamenti', 'checkUpdatesManual')
    .addItem('Info', 'showInfo')
    .addToUi();
  checkUpdates(false);
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
    ui.alert('🚀 Aggiornamento Disponibile', `Nuova versione: ${updateInfo.version}\nVersione attuale: ${MAJOR_VERSION}.${MINOR_VERSION}.${PATCH_VERSION}\n\nScarica l'ultima versione da GitHub.`, ui.ButtonSet.OK);
  } else if (showIfUpdated) {
    ui.alert('🔝 Sei aggiornato!', `Versione attuale: ${MAJOR_VERSION}.${MINOR_VERSION}.${PATCH_VERSION}`, ui.ButtonSet.OK);
  }
}

function showInfo() {
  const anno = new Date().getFullYear();
  const html = `<div style="font-family:sans-serif;text-align:center;"><h2 style="color:#1a73e8;">🏐 AutoCalendar</h2><p>Versione: ${MAJOR_VERSION}.${MINOR_VERSION}.${PATCH_VERSION}<br>Autore: Matteo Checcacci</p><a href="${githubUrl}" target="_blank" style="background:#24292e;color:white;padding:8px;text-decoration:none;border-radius:5px;">GitHub</a><p style="font-size:0.7em;">© ${anno} KekkoTech</p></div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(300), 'Info');
}

function setUserName() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Nome', 'Inserisci COGNOME NOME:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var v = (res.getResponseText() || "").trim();
    if (v) PropertiesService.getScriptProperties().setProperty('USER_FULL_NAME', v);
  }
}

function setGeminiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('API Gemini', 'Inserisci API Key:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var v = (res.getResponseText() || "").trim();
    if (v) PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', v);
  }
}

function setFrequency() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Frequenza', 'Ore (1, 2, 4, 6, 8, 12):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var freq = parseInt((res.getResponseText() || "").trim(), 10);
    var allowed = { 1: true, 2: true, 4: true, 6: true, 8: true, 12: true };
    if (!allowed[freq]) {
      ui.alert('Valore non valido', 'Inserisci solo: 1, 2, 4, 6, 8, 12', ui.ButtonSet.OK);
      return;
    }
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => { if (t.getHandlerFunction() == 'syncGmailToSheetAndCalendar') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('syncGmailToSheetAndCalendar').timeBased().everyHours(freq).create();
  }
}

function setGuests() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Invitati', 'Email (separate da virgola):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('CALENDAR_GUESTS', (res.getResponseText() || "").trim());
  }
}

function setupTrigger() {
  setUserName();
  setGeminiKey();
  setFrequency();
  setGuests();
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
          if (aiData.ora) dataGara.ora = aiData.ora;
          if (aiData.codA) dataGara.codA = aiData.codA;
          if (aiData.codF) dataGara.codF = aiData.codF;
          if (aiData.luogo) dataGara.luogo = aiData.luogo;
          if (aiData.squadraCasa) dataGara.squadraCasa = aiData.squadraCasa;
          if (aiData.squadraOspite) dataGara.squadraOspite = aiData.squadraOspite;
          if (aiData.data) dataGara.data = aiData.data;
          if (aiData.numeroGara) dataGara.numeroGara = aiData.numeroGara;
          if (aiData.categoria) dataGara.categoria = aiData.categoria;
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

        if (subject.includes("VARIAZIONE") || subject.includes("SPOSTAMENTO")) {
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
          dataGara.arb1 || oldRow[9],
          dataGara.arb2 || oldRow[10],
          ""
        ];
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else if (!subject.includes("VARIAZIONE") && !subject.includes("SPOSTAMENTO")) {
        var rowDataNew = [
          dataGara.data, dataGara.ora, dataGara.luogo,
          dataGara.squadraCasa, dataGara.squadraOspite,
          dataGara.categoria, dataGara.numeroGara,
          isRegionale ? "-" : (dataGara.codA || ""),
          isRegionale ? "-" : (dataGara.codF || ""),
          dataGara.arb1 || "", dataGara.arb2 || "", ""
        ];
        sheet.appendRow(rowDataNew);
        addedCount++;
      }
      msg.markRead();
    }
  });

  SpreadsheetApp.flush();
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
    var end = new Date(start.getTime() + 10800000);
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

    // Estrazione data corretta per formati TieBreak normalizzati (es: "01 - 11 - 2025" o "01/11/2025")
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

function parseSpostamento(html) {
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

function cleanEmail(html) {
  var s = String(html || "")
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
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
  const prompt = `Analizza HTML designazione volley. Estrai dati gara. Rispondi SOLO JSON: {"data":"DD/MM/YYYY","ora":"HH:MM","luogo":"...","squadraCasa":"...","squadraOspite":"...","categoria":"...","numeroGara":"...","codA":"...","codF":"..."}. Testo: ${html}`;
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
      if (myEmail) GmailApp.sendEmail(myEmail, "🚀 Aggiornamento AutoCalendar", "Nuova versione " + updateInfo.version + " disponibile su GitHub.");
    }
    count = 0;
  }
  props.setProperty('UPDATE_CHECK_COUNT', count.toString());
}
