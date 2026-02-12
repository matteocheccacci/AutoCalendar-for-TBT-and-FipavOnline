# 🏐 AutoCalendar: Sincronizzazione Automatica Gare

AutoCalendar è uno script avanzato per **Google Apps Script** progettato per gli Ufficiali di Gara. Automatizza la gestione delle designazioni leggendo le email di **TieBreakTech** e **FipavOnline**, popolando un database su Google Sheets e creando eventi intelligenti sul tuo Calendario.

---

## 🛠️ Requisiti Iniziali del Foglio

1. **Crea un nuovo Foglio Google**.
2. **Rinomina il foglio di lavoro**: La linguetta in basso deve chiamarsi esattamente **`Arbitro`**.
3. **Prepara le intestazioni**: Inserisci questi titoli nella **Riga 1**:

| Colonna | Intestazione |
| :--- | :--- |
| **A** | Data |
| **B** | Ora |
| **C** | Luogo |
| **D** | Casa |
| **E** | Ospite |
| **F** | Categoria |
| **G** | Nr. Gara |
| **H** | Cod. Attivazione |
| **I** | Cod. Firma |
| **J** | 1° Arb |
| **K** | 2° Arb |

---

## 🚀 Installazione Rapida

### 1. Caricamento Script
* Dal tuo Foglio Google, vai su **Estensioni** > **Apps Script**.
* Incolla il contenuto del file `Codice.gs`.
* Salva il progetto con il nome "AutoCalendar".

### 2. Ottenere la Gemini API Key (non obbligatorio)
Per un'estrazione perfetta che distingua squadre da indirizzi o dirigenti, lo script utilizza l'IA di Google.
1. Accedi a [Google AI Studio](https://aistudio.google.com/).
2. Clicca su **"Get API key"** e poi su **"Create API key in new project"**.
3. Copia la chiave generata (trattala come una password personale).

### 3. Configurazione Finale
* Torna al Foglio Google e ricarica la pagina.
* Nel nuovo menù **🏐 AutoCalendar**, seleziona **Configurazione Iniziale Completa**.
* Inserisci il tuo cognome e il tuo nome, inserisci l'API Key se la stai utilizzando, scegli la frequenza (es. ogni 2 ore) e aggiungi eventuali invitati (es. i tuoi familiari).


---

## ⚠️ Se usi un altro provider per la posta elettronica (es. Outlook)

1. Andare nelle impostazioni di Gmail (da PC):

2.  Vai in Impostazioni > Visualizza tutte le impostazioni.

3. Scheda Account e importazione.

4. Sezione Controlla la posta da altri account > Aggiungi un account email.

5. Inserisce i dati di Outlook. In questo modo Gmail "pesca" le mail da Outlook e lo script le troverà come se fossero native.

---

## 📖 Funzionalità Principali

* **Lettura Intelligente**: Distingue le email inviate da Fipav Web Manager e quelle di FipavOnline.
* **Filtro Duplicati Stagionale**: Verifica la combinazione `Data + Numero Gara`, per evitare la creazione di duplicati.
* **Notifiche Calendario**: 
  * Un promemoria **24 ore prima** del match.
  * Un promemoria la **mattina stessa alle ore 08:00**.
* **Sincronizzazione Manuale**: Scansiona tutte le email degli ultimi 30 giorni (anche se già lette) per recuperare designazioni perse.

---

## ⚠️ Note di Sicurezza e Mobile

* **Privacy**: La tua API Key viene salvata nelle `ScriptProperties` protette di Google. Non è visibile nel codice se condividi il foglio.
* **Google Calendar**: Se non vedi gli eventi sul tuo smartphone, apri l'app **Google Calendar** > **Impostazioni** > **Partite - AC** e attiva la **Sincronizzazione**.

---

## 📄 Licenza e Segnalazione Bug

Il prodotto è concesso in **Licenza MIT**. Il codice è aperto e disponibile per modifiche.
Puoi segnalare bug o suggerire migliorie tramite la sezione **Issues** su GitHub.

---
## ©️ Info Copyright

* **TBT, FWM, Fipav Web Manager sono copyright di TieBreakTech. 
* **FipavOnline è copyright di Manufacturing Point Software.
---

**© 2026 KekkoTech Softwares - RefPublic Team** *Sviluppato da Matteo Checcacci*
