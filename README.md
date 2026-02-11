# 🏐 Guida Setup: Sincronizzazione Automatica Gare

Questo script trasforma il tuo Foglio Google in un assistente personale per l'arbitraggio. Legge le email di designazione (Regionali e Territoriali), compila il tuo archivio gare e sincronizza gli eventi su un calendario Google dedicato.

---

## 🛠️ Requisiti Iniziali

1.  **Foglio Google**: Crea un nuovo foglio.
2.  **Nome Foglio**: Rinomina la linguetta in basso esattamente come **`Arbitro`**.
3.  **Intestazioni**: Inserisci questi titoli nella **Riga 1**:
    * **A:** Data | **B:** Ora | **C:** Luogo | **D:** Casa | **E:** Ospite | **F:** Categoria | **G:** Nr. Gara | **H:** Cod. Attivazione | **I:** Cod. Firma | **J:** 1° Arb | **K:** 2° Arb

---

## 🚀 Configurazione in 5 Minuti

### 1. Installazione del Codice
* Nel tuo Foglio Google: **Estensioni** > **Apps Script**.
* Incolla il codice fornito nel file `Codice.gs` (cancellando eventuali testi preesistenti).
* Clicca sull'icona del floppy (**Salva**) e dai un nome al progetto.

### 2. Attivazione e API Gemini (Opzionale ma Consigliato)
L'integrazione con Gemini permette un parsing intelligente che evita errori con i nomi dei dirigenti o indirizzi complessi.
* Ottieni una chiave gratuita su [Google AI Studio](https://aistudio.google.com/).
* Torna al foglio, ricarica la pagina e vai nel menù **🏐 AutoCalendar** > **Configura API Gemini e Trigger**.
* Incolla la tua chiave quando richiesto. Se non vuoi usarla, lascia il campo vuoto: lo script userà il sistema di ricerca standard.

### 3. Primo Avvio
* Dal menù **🏐 AutoCalendar**, seleziona **Sincronizzazione MANUALE (Tutto)**.
* Google ti chiederà di autorizzare lo script. Clicca su **Avanzate** > **Vai a [Nome Progetto] (non sicura)** per procedere.
* Lo script creerà automaticamente un calendario chiamato **"Partite - AC"**.

---

## 📖 Funzionalità

1.  **Sincronizzazione Automatica**: Una volta configurato il Trigger, lo script controllerà Gmail ogni 2 ore alla ricerca di nuove gare.
2.  **Calendario Dedicato**: Gli eventi vengono inseriti nel calendario "Partite - AC" per non intasare il tuo calendario personale.
3.  **Controllo Duplicati**: Lo script verifica il numero gara: non avrai mai doppioni nel foglio o nel calendario.
4.  **Descrizioni Pulite**: Gli eventi in calendario mostrano solo i dati essenziali (Gara e Codici Referto) per una consultazione rapida da smartphone.

---

## ⚠️ Note Importanti

* **Email Lette/Non Lette**: L'automazione scarica solo le email **non lette**. La funzione manuale scansiona invece tutte le email degli ultimi 30 giorni.
* **Mobile**: Per vedere le partite sul telefono, ricordati di attivare la sincronizzazione del nuovo calendario "Partite - AC" nelle impostazioni dell'app Google Calendar.

---
*Dal creatore di RefPublic*
