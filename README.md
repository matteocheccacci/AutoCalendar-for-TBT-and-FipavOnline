# 🏐 Guida Setup: Sincronizzazione Automatica Gare

Questo script trasforma il tuo Foglio Google in un assistente personale per l'arbitraggio. Legge le email di designazione, aggiorna l'archivio gare e sincronizza tutto sul tuo calendario.

---

## 🛠️ Requisiti Iniziali

1.  **Foglio Google**: Crea un nuovo Foglio Google.
2.  **Nome Foglio**: Rinomina la linguetta in basso (il foglio di lavoro) esattamente come **`Arbitro`**.
3.  **Intestazioni**: Inserisci i seguenti titoli nella prima riga (Riga 1) per mantenere l'ordine:
    * **A:** Data | **B:** Ora | **C:** Luogo | **D:** Casa | **E:** Ospite | **F:** Categoria | **G:** Nr. Gara | **H:** Cod. Attivazione | **I:** Cod. Firma | **J:** 1° Arb | **K:** 2° Arb

---

## 🚀 Configurazione in 5 Minuti

### 1. Installazione del Codice
* Nel tuo Foglio Google, vai su **Estensioni** > **Apps Script**.
* Cancella tutto il codice presente nell'editor.
* Incolla il codice fornito nel file `Codice.gs`.
* Clicca sull'icona del floppy (**Salva**) e assegna un nome al progetto (es. *Gestione Arbitri*).

### 2. Primo Avvio e Autorizzazione
* Torna al Foglio Google e **ricarica la pagina** del browser.
* Vedrai apparire il menù **🏐 Sync Arbitri**.
* Seleziona **Scarica Nuove Gare da Gmail**.
* Google ti chiederà di autorizzare lo script. 
    * *Nota:* Se appare l'avviso "App non verificata", clicca su **Avanzate** e poi su **Vai a [Nome Progetto] (non sicura)** per procedere.

### 3. Attivazione del Pilota Automatico
Per fare in modo che lo script lavori da solo senza che tu debba aprire il foglio:
* Vai nel menù **🏐 Sync Arbitri** > **Configura Automazione (Trigger)**.
* Inserisci ogni quante ore vuoi che lo script controlli la posta (es. `6`).
* Clicca **OK**. Lo script ora è attivo sui server Google e girerà anche a PC spento.

---

## 📖 Come Funziona

1.  **Lettura Gmail**: Lo script cerca email non lette provenienti da *TieBreak* o *Fipav Varese*. 
2.  **Archiviazione**: Estrae Numero Gara, Squadre, Data e Codici, inserendo una nuova riga nel foglio.
3.  **Calendario**: Crea automaticamente l'evento sul tuo calendario principale.
4.  **Notifiche**: L'evento include promemoria popup:
    * **24 ore prima** della gara.
    * **2 ore prima** della gara.
5.  **Sicurezza**: Se ricevi un'email per una gara già presente, lo script la ignorerà per evitare duplicati.

---

## ⚠️ Note Importanti

* **Email "Non Lette"**: Lo script elabora solo le email contrassegnate come non lette. Se vuoi ricaricare una gara vecchia, segnala come "da leggere" su Gmail e avvia la funzione manualmente.
* **Modifiche**: Se cambi l'ID del calendario o il nome del foglio, ricordati di aggiornare le variabili corrispondenti nel codice.

---
*Dal creatore di RefPublic*
