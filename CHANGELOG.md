# Changelog Tecnico Dettagliato: BDAP Automation


## 1. Evoluzione della Struttura di `Template_Analisi.xlsx`

| Versione | Fogli Presenti nel Template | Dimensioni / Righe Chiave | Modifiche Strutturali |
| :--- | :--- | :--- | :--- |
| **v0 / v0.1.0** | `CONCLUSIONI`<br>`VERIFICA QUESTIONARI`<br>`FCDE`<br>`ANALISI CONTABILE`<br>`CONTROLLI-POST`<br>`MAPPATURE_BDAP` | **FCDE:** 7 righe, 9 colonne.<br>**ANALISI CONTABILE:** 37 righe contabili compilabili. | Struttura iniziale ereditata dal progetto originario di stage. |
| **v0.2.0 / v0.3.1** | *Stessi fogli di v0* | **ANALISI CONTABILE:** Aggiunte righe (10 nuove voci contabili). | Inserimento delle voci per spese legali, avvocatura e fondo perequativo TARI. |
| **v0.4.0** | Aggiunto **`Piano Indicatori`** | **Piano Indicatori:** 32 righe, 11 colonne.<br>**FCDE:** Esteso a 10 colonne (supporto anni fino al 2026). | Introduzione del foglio dedicato al monitoraggio degli indicatori sintetici BDAP. |
| **v0.4.1 -> v0.4.3** | **`ANAGRAFICA`**<br>`FCDE`<br>`ANALISI CONTABILE`<br>`Piano Indicatori`<br>`CONTROLLI-POST`<br>`MAPPATURE_BDAP` | **ANAGRAFICA:** 28 righe, 8 colonne.<br>**FCDE:** Esteso a **9 righe**.<br>*(Eliminati definitivamente `CONCLUSIONI` e `VERIFICA QUESTIONARI`)* | Introduzione del foglio anagrafica istituzionale; revisione formule e percentuali FCDE. |

<br><br><br>

## 2. Changelog Analitico Versione per Versione

### BDAP_Automation (Stage Locatelli)
Versione iniziale del repository, derivata dal lavoro di stage (*Stage_Locatelli*).

* **Architettura dei Moduli:**
  * `core/automation.py`: Compilazione del foglio `ANALISI CONTABILE`. <br>
  Lettura riga per riga dei riferimenti contabili da colonna A/B, normalizzazione delle etichette e scrittura dei valori numerici estratti dai rendiconti BDAP nelle colonne corrispondenti agli anni analizzati.
  * `core/automation_fcde.py`: Popolamento delle righe del foglio `FCDE` (Importo FCDE a riga 2, Residui Titolo I a riga 3, Residui Titolo III a riga 4, Totale residui a riga 5 con formula `=C3+C4`, percentuale a riga 6 con formula `=C2*100/C5`, metodo a riga 7 impostato su `"ordinario"`).
  * `core/automation_questionario.py`: Individuazione dei questionari del bilancio consuntivo tramite pattern nel nome ed estrazione con tolleranza di riga (`row_tolerance`).
  * `core/automation_search.py`: Utility per l'individuazione dinamica delle colonne relative agli anni (`find_year_column`) e per la risoluzione dei fogli tramite indice o sottostringa (`find_sheet`).
  * `orchestration/cli_discovery.py`: Rilevamento automatico della cartella del comune e dei file di rendiconto BDAP tramite regex `Rend.\s*(\d{4})\.xlsx$`.
* **Tracciabilità dei Dati:** Inserimento di note native `openpyxl.comments.Comment` con schema a riga singola:
  `Fonte BDAP: <file> | foglio <nome_foglio>, riga <r>, colonna <c> (<coord>)`.
* **Vincoli Iniziali:** Percorso del template Excel vincolato rigidamente alla cartella di lavoro locale; assenza di supporto ad ambienti congelati (PyInstaller / py2app)

<br>
<br>

### BDAP_Automation v0.1.0
*Obiettivo: Predisposizione per la creazione di pacchetti eseguibili ed installer per Windows e macOS; correzione geometria interfaccia grafica.*

* **Risoluzione Percorsi Risorse (`cli_discovery.py`):**
  * Introdotta la funzione `get_resource_root() -> Path`:
    * Rileva la presenza dell'attributo `sys._MEIPASS` (cartella temporanea creata da PyInstaller).
    * Rileva il flag `sys.frozen` per bundle macOS creati con py2app (risolvendo il percorso interno `Contents/Resources`).
    * Fallback su percorso standard relativo al file di codice sorgente (`parents[2]`).
  * Aggiornata la funzione `resolve_default_template_path()`: antepone `get_resource_root()` alla lista dei percorsi di ricerca, consentendo il recupero immediato di `Template_Analisi.xlsx` all'interno dell'eseguibile.
* **Layout Interfaccia Grafica (`texts.py`):**
  * Costante `WINDOW_GEOMETRY` modificata da `"1050x680"` a `"1050x810"` per evitare il troncamento dei selettori di percorso e della finestra di avanzamento.

<br>
<br>

### BDAP_Automation v0.2.0
*Obiettivo: Estensione delle rilevazioni contabili a spese legali, avvocatura e fondo TARI; supporto a coordinate multiple dai questionari.*

* **Struttura Template Excel (`Template_Analisi.xlsx`):**
  * Aggiunte 10 nuove righe contabili nel foglio `ANALISI CONTABILE`:
    * Riga 38: `UFFICIO AVVOCATURA / AVVOCATURA`
    * Riga 39: `AVVOCATI DIRIGENTI`
    * Riga 40: `AVVOCATI FUNZIONARI`
    * Riga 41: `FONDI AVVOCATURA DIRIGENTI`
    * Riga 42: `FONDI AVVOCATURA FUNZIONARI`
    * Riga 43: `INCARICHI LEGALI ESTERNI`
    * Riga 44: `PATROCINIO LEGALE`
    * Riga 45: `CONSULENZA LEGALE`
    * Riga 46: `SERVIZI LEGALI`
    * Riga 47: `FONDO PEREQUATIVO TARI`

* **Estrazione Questionari (`automation_questionario.py`):**
  * **Filtro candidati:** In `_is_questionario_candidate()`, aggiunta l'esclusione esplicita per i file contenenti `"novità"` nel nome (`novità.xlsx`), evitando falsi positivi causati da documenti di aggiornamento salvati nella stessa cartella.
  * **Supporto a liste di coordinate (`cell_refs`):** Gestione di mappature contenenti più celle da aggregare:
    * Scansione sequenziale delle celle indicate con tolleranza di riga (`row_tolerance`).
    * Fallback su ricerca dell'etichetta tramite `find_nearby_title_row` nel caso in cui le celle abbiano subito uno slittamento rispetto al modello standard.
* **Mappature Predefinite (`default_mappings.py`):**
  * Inserite le definizioni di estrazione per le nuove voci dal foglio `SPESA_PERSONALE` (`ufficioavvocatura`, `avvocatidirigenti`, `avvocatifunzionari`, `fondiavvocaturadirigenti`, `fondiavvocaturafunzionari`, `incarichilegaliesterni`, `patrociniolegale`, `consulenzalegale`, `servizilegali`).
  * Inserita la mappatura per `fondoperequativotari` (foglio `GEST. FIN. RESIDUI`, cella `I42`).
  * Ristrutturata la voce `contrastoallevasionetributaria` con `table_refs` multipli su fogli entrate alternativi (`SEZ. II - DATI ENTRATE`) per recuperare sia le entrate IMU/TASI sia le entrate TARSU/TIA/TARI.

<br>
<br>

### BDAP_Automation v0.3.0
*Obiettivo: Documentazione tecnica e standardizzazione operativa delle mappature.*

* **Manuale Mappature (`MAPPINGS.md`):**
  * Riorganizzazione della documentazione con spiegazione dettagliata del funzionamento delle mappature deterministiche (`sheet_idx`, `cell_ref`, `expected_label`, `label_check_cells`).
  * Indicazioni sulle modalità di fallback in caso di layout BDAP differenti.

<br>
<br>

### BDAP_Automation v0.3.1
*Obiettivo: Integrazione dei requisiti per packaging PyInstaller e canali di distribuzione.*

* **Dipendenze (`requirements.txt`):**
  * Aggiunto `PyInstaller>=6.0`.
* **Distribuzione (`README.md`):**
  * Inseriti i riferimenti al repository GitHub di build (`Fil0807/BDAP-Automation-Deployment`) e alla cartella cloud Google Drive contenente gli installer precompilati.

<br>
<br>

### BDAP_Automation v0.4.0
*Obiettivo: Introduzione del modulo Piano Indicatori, centralizzazione e formattazione dei commenti Excel, estrazione anagrafica comune e supporto alla lettura PDF.*

* **Nuovo Modulo `core/automation_piano_indicatori.py`:**
  * Gestione e compilazione del nuovo foglio **`Piano Indicatori`** inserito in `Template_Analisi.xlsx`.
  * Rilevamento automatico dei file ausiliari degli indicatori sintetici tramite `_discover_aux_workbook(bdap_path, "indicatori", year)` (file con pattern `*_ind_*` o `*indicatori*`).
  * Risoluzione per ciascun anno dei seguenti indicatori di bilancio:
    * **1.1:** Incidenza spese rigide su entrate correnti
    * **2.8:** Incidenza incassi entrate proprie su previsioni definitive
    * **3.2:** Anticipazioni chiuse solo contabilmente
    * **4.1:** Incidenza spesa di personale su spesa corrente
    * **10.3:** Sostenibilità debiti finanziari
    * **12.4:** Sostenibilità disavanzo a carico dell'esercizio
    * **13.1:** Debiti riconosciuti e finanziati
    * **13.2 + 13.3:** Somma debiti in corso di riconoscimento e in corso di finanziamento
    * **9.2:** Indicatore tempi medi di pagamento fornitori
* **Nuovo Modulo `support/comment_builder.py`:**
  * Calcolo delle dimensioni dei commenti per prevenire il troncamento su Excel:
    * `width = max(450, min(750, max_len * 9 + 60))`
    * `height = max(170, len(lines) * 26 + 45)`
  * Formattazione delle coordinate di cella: `E45 (Riga 45, Colonna 5)`.
  * Funzione `fix_all_comments_in_workbook()`: applicata sistematicamente a tutti i fogli prima del salvataggio finale per reimpostare dimensioni minime garantite.
  * Funzione `extract_a1_breakdown()`: analizza l'Allegato A1 del rendiconto BDAP ed inserisce nel commento della quota accantonata il dettaglio completo dei singoli fondi (contenzioso, perdite partecipate, rinnovi contrattuali, altri accantonamenti).
* **Nuovo Modulo `support/conclusioni.py`:**
  * Compilazione del foglio `CONCLUSIONI`:
    * Cella `C2`: Nome dell'ente rilevato dai rendiconti o dalla struttura delle directory.
    * Cella `C3`: Popolazione residente dell'ultimo anno analizzato, rilevata tramite regex su relazioni revisori in formato PDF o dai prospetti BDAP.
* **Miglioramenti al Modulo FCDE (`automation_fcde.py`):**
  * Introduzione della ricerca estesa a tutto il foglio in caso di mancata corrispondenza della cella base, controllando le etichette nelle colonne A e B ed estraendo il dato contabile da colonna D.
  * Aggiunta funzione `read_fcde_totale_from_workbook()` per il recupero della somma residui attivi già calcolata dinamicamente.
* **Correzioni di Stabilità (`cli_discovery.py`):**
  * Inizializzazione protetta dei limiti massimi di ricerca (`max_row = 40`, `max_col = 12`) prima del controllo dell'etichetta *Denominazione ente* su cella A7, prevenendo cicli vuoti su workbook con `sheet.max_column = None`.
* **Formattazione Dati (`value_formatter.py`):**
  * Aggiunto supporto per la conversione di flag booleani/numerici ($1/0$) nelle stringhe `"Sì"` e `"No"` (`boolean_yes_no=True`).
* **Dipendenze PDF (`requirements.txt`):**
  * Aggiunti i pacchetti `pymupdf>=1.24.0` e `pypdf>=4.0.0` per consentire l'ispezione documentale degli allegati PDF.

<br>
<br>

### BDAP_Automation v0.4.1
*Obiettivo: Aggiunta del foglio ANAGRAFICA, calcolo tempestività deliberazione consiliare, ristrutturazione foglio FCDE e robustezza multipiattaforma.*

* **Nuovo Modulo `core/automation_anagrafica.py`:**
  * Estrazione delle informazioni anagrafiche dai questionari di bilancio tramite ispezione dei fogli `ANAGRAFICA`/`SCHEDA_ANAGRAFICA` e del foglio `DATI`.
  * Risoluzione delle formule di puntamento ministeriali basate sulla funzione `=SCARTO(...)` / `=OFFSET(...)`.
  * Popolamento delle tre sezioni del nuovo foglio **`ANAGRAFICA`**:
    * **Sezione 1 - Dati Generali dell'Ente (riga 5..10, colonna C):** Denominazione, Codice Fiscale, Provincia, Regione, Tipologia Ente, Popolazione residente.
    * **Sezione 2 - Estremi Deliberazione Rendiconto (righe 14..19 per ciascun esercizio):**
      * Colonna C: Numero deliberazione consiglio (`n. <delib_n>`)
      * Colonna D: Data deliberazione
      * Colonna E: Termine di legge stabilito al 30 aprile dell'anno successivo (`30/04/<anno+1>`)
      * Colonna F: **Valutazione automatica della tempestività:**
        Imposta `"Tempestivo"` se la data di delibera è $\le$ 30 aprile dell'anno successivo, altrimenti imposta `"In ritardo"`.
    * **Sezione 3 - Relazione dell'Organo di Revisione (righe 23..27):**
      Numero e data del verbale, ruolo, nominativo del revisore, codice fiscale ed indirizzo e-mail.
* **Ristrutturazione `Template_Analisi.xlsx`:**
  * Eliminazione dei fogli `CONCLUSIONI` e `VERIFICA QUESTIONARI`.
  * Inserimento del foglio `ANAGRAFICA` come prima scheda del documento.
* **Revisione e Formule Foglio FCDE (`automation_fcde.py`):**
  * Griglia di calcolo estesa da 7 a **9 righe**:
    * Riga 2: Importo FCDE
    * Riga 3: Residui attivi Titolo I
    * Riga 4: Residui attivi Titolo III
    * Riga 5: Totale residui attivi (`=C3+C4`)
    * Riga 6: Quota avanzo disponibile da `ANALISI CONTABILE` riga 19.
    * Riga 7: `% rispetto a residui attivi conservati` (formula: `=C2*100/C5`).
    * Riga 8: `% FCDE + parte disponibile su residui attivi`:
      Formula: `=({Col}2+{Col}6)*100/{Col}5`.
    * Riga 9: Metodo di calcolo (`"ordinario"`).
  * Aggiunta funzione `read_fcde_percentage_plus_disponibile_from_workbook()` per esportare la percentuale combinata verso la riga 26 del foglio `ANALISI CONTABILE`.
* **Disambiguazione Righe Omonime (`automation.py`):**
  * La funzione `_resolve_special_source_entry()` integra logiche contestuali di disambiguazione basate sul numero di riga e sul testo di origine (`source_hint`):
    * **Anticipazione di Tesoreria:**
      * Riga 32 (o sorgente rendiconto/BDAP): associata a `anticipazionetesoreria` (Titolo 7 entrate rendiconto).
      * Riga 33 (o sorgente questionario): associata a `anticipazionetesoreriaquestionario` (risposta Sì/No dal questionario).
    * **Fondo Anticipazioni di Liquidità (FAL):**
      * Riga 36 (rendiconto): associata a `fal`.
      * Riga 37 (questionario): associata a `falquestionario`.
    * **Altre voci disambiguate:** `importononrestituitoanticipazione`, `incidenzaprime3titoli`, `debitocomplessivocontratto`.
* **Filtro File di Sistema macOS:**
  * In `automation_indicatori.py` e `controlli_post.py`, esclusi i file ombra generati su macOS con prefisso `._` o residenti nella directory `__MACOSX`, prevenendo errori di lettura su archivi compressi o chiavette USB.

<br>
<br>

### BDAP_Automation v0.4.2
* Modifica di allineamento documentale: Corretta la sintassi di alcuni file HTML e i collegamenti nel file `README.md`.

<br>
<br>

### BDAP_Automation v0.4.3
*Obiettivo: Allineamento visivo centrato dei dati contabili compilati, bonifica delle note esplicative per valori predefiniti, correzione coordinate contabili e completamento della documentazione.*

* **Centratura Automatica Dati (`automation.py` e `cli_pipeline.py`):**
  * Introdotta la funzione `format_inserted_data_cells_center(workbook_or_path, log=None)`, eseguita come fase conclusiva dell'intera pipeline:
    * Esclude i fogli di puro testo (`MAPPATURE_BDAP` e `CONTROLLI-POST`).
    * **Fogli contabili (`ANALISI CONTABILE`, `FCDE`, `Piano Indicatori`):**
      Rileva le colonne degli anni nelle intestazioni (righe 1..6) e applica l'allineamento centrato orizzontale e verticale (`horizontal="center", vertical="center"`) a tutte le celle con dati effettivi, preservando l'eventuale impostazione di testo a capo (`wrap_text`).
    * **Foglio `ANAGRAFICA`:**
      Centra le celle della colonna C nelle righe 5..10 (Dati generali) e le colonne B..H nelle tabelle delle deliberazioni e dei verbali.
* **Bonifica Commenti per Valori Predefiniti (`comment_builder.py`):**
  * Aggiunta la tracciabilità specifica per `fondivincolaticovid`: annota l'importo lordo derivato da *All. a2 col. O (Fondo solidarietà alimentare)*.
* **Correzione Coordinate Contabili (`default_mappings.py`):**
  * **`fondocontenzioso`:**
    * Coordinata primaria aggiornata da `I16` a `I17`.
    * Lista riferimenti di fallback impostata a `['I17', 'I16', 'I14']`.
    * Verifica etichetta estesa a `B17` e `B16`.
  * **`fondogaranziadebiticommerciali`:**
    * Riordinata la lista delle etichette con priorità per `FONDO GARANZIA DEBITI COMMERCIALI`.
    * Inserite le celle di fallback `I19` e `I23`.

<br>
<br>

### BDAP_Automation v1.0.0
* Modifica di allineamento documentale e testuale: Corretti file e stringhe riportanti la versione del programma e aggiornati alle 1.0.0

<br>
<br>
