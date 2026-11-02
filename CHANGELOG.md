# Changelog Tecnico Dettagliato: BDAP Automation (v0 -> v0.4.3)

Il presente documento descrive in dettaglio la cronologia e l'evoluzione tecnica del progetto **BDAP Automation**, analizzando le modifiche apportate alla struttura del codice, alle regole di estrazione dai rendiconti e dai questionari, al file modello Excel e all'orchestrazione della pipeline a partire dalla versione iniziale (**BDAP_Automation-0**) fino all'ultima versione (**BDAP_Automation-0.4.3**).

---

## 1. Evoluzione della Struttura di `Template_Analisi.xlsx`

| Versione | Fogli Presenti nel Template | Dimensioni / Righe Chiave | Modifiche Strutturali |
| :--- | :--- | :--- | :--- |
| **v0 / v0.1.0** | `CONCLUSIONI`<br>`VERIFICA QUESTIONARI`<br>`FCDE`<br>`ANALISI CONTABILE`<br>`CONTROLLI-POST`<br>`MAPPATURE_BDAP` | **FCDE:** 7 righe, 9 colonne.<br>**ANALISI CONTABILE:** 37 righe contabili compilabili. | Struttura iniziale ereditata dal progetto originario di stage. |
| **v0.2.0 / v0.3.1** | *Stessi fogli di v0* | **ANALISI CONTABILE:** Aggiunte righe da 38 a 47 (10 nuove voci contabili). | Inserimento delle voci per spese legali, avvocatura e fondo perequativo TARI. |
| **v0.4.0** | Aggiunto **`Piano Indicatori`** | **Piano Indicatori:** 32 righe, 11 colonne.<br>**FCDE:** Esteso a 10 colonne (supporto anni fino al 2026). | Introduzione del foglio dedicato al monitoraggio degli indicatori sintetici BDAP. |
| **v0.4.1 -> v0.4.3** | **`ANAGRAFICA`**<br>`FCDE`<br>`ANALISI CONTABILE`<br>`Piano Indicatori`<br>`CONTROLLI-POST`<br>`MAPPATURE_BDAP` | **ANAGRAFICA:** 28 righe, 8 colonne.<br>**FCDE:** Esteso a **9 righe**.<br>*(Eliminati definitivamente `CONCLUSIONI` e `VERIFICA QUESTIONARI`)* | Introduzione del foglio anagrafica istituzionale; revisione formule e percentuali FCDE. |

---

## 2. Changelog Analitico Versione per Versione

---

### BDAP_Automation-0 (Baseline)
Versione iniziale del repository, derivata dal lavoro di stage (*Stage_Locatelli*).

* **Architettura dei Moduli:**
  * `core/automation.py`: Compilazione del foglio `ANALISI CONTABILE`. Lettura riga per riga dei riferimenti contabili da colonna A/B, normalizzazione delle etichette e scrittura dei valori numerici estratti dai rendiconti BDAP nelle colonne corrispondenti agli anni analizzati.
  * `core/automation_fcde.py`: Popolamento delle righe 1..7 del foglio `FCDE` (Importo FCDE a riga 2, Residui Titolo I a riga 3, Residui Titolo III a riga 4, Totale residui a riga 5 con formula `=C3+C4`, percentuale a riga 6 con formula `=C2*100/C5`, metodo a riga 7 impostato su `"ordinario"`).
  * `core/automation_questionario.py`: Individuazione dei questionari del bilancio consuntivo tramite pattern nel nome ed estrazione con tolleranza di riga (`row_tolerance`).
  * `core/automation_search.py`: Utility per l'individuazione dinamica delle colonne relative agli anni (`find_year_column`) e per la risoluzione dei fogli tramite indice o sottostringa (`find_sheet`).
  * `orchestration/cli_discovery.py`: Rilevamento automatico della cartella del comune e dei file di rendiconto BDAP tramite regex `Rend.\s*(\d{4})\.xlsx$`.
* **Tracciabilità dei Dati:** Inserimento di note native `openpyxl.comments.Comment` con schema a riga singola:
  `Fonte BDAP: <file> | foglio <nome_foglio>, riga <r>, colonna <c> (<coord>)`.
* **Vincoli Iniziali:** Percorso del template Excel vincolato rigidamente alla cartella di lavoro locale; assenza di supporto ad ambienti congelati (PyInstaller / py2app); finestra GUI con risoluzione fissa `1050x680`.

---

### BDAP_Automation-0.1.0 (Transizione da v0)
*Scopo: Predisposizione per la creazione di pacchetti eseguibili ed installer per Windows e macOS; correzione geometria interfaccia grafica.*

* **Risoluzione Percorsi Risorse (`cli_discovery.py`):**
  * Introdotta la funzione `get_resource_root() -> Path`:
    * Rileva la presenza dell'attributo `sys._MEIPASS` (cartella temporanea creata da PyInstaller).
    * Rileva il flag `sys.frozen` per bundle macOS creati con py2app (risolvendo il percorso interno `Contents/Resources`).
    * Fallback su percorso standard relativo al file di codice sorgente (`parents[2]`).
  * Aggiornata la funzione `resolve_default_template_path()`: antepone `get_resource_root()` alla lista dei percorsi di ricerca, consentendo il recupero immediato di `Template_Analisi.xlsx` all'interno dell'eseguibile.
* **Layout Interfaccia Grafica (`texts.py`):**
  * Costante `WINDOW_GEOMETRY` modificata da `"1050x680"` a `"1050x810"` per evitare il troncamento dei selettori di percorso e della finestra di avanzamento.
* **Documentazione (`README.md`):**
  * Aggiornato per specificare lo status di fork orientato alla conversione in installer standalone.

---

### BDAP_Automation-0.2.0 (Transizione da v0.1.0)
*Scopo: Estensione delle rilevazioni contabili a spese legali, avvocatura e fondo perequativo TARI; supporto a coordinate multiple dai questionari.*

* **Struttura Template Excel (`Template_Analisi.xlsx`):**
  * Aggiunte 10 nuove righe contabili nel foglio `ANALISI CONTABILE` (righe 38..47):
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
  * **Filtro candidati:** In `_is_questionario_candidate()`, aggiunta l'esclusione esplicita per i file contenenti `"novit"` nel nome (`novità.xlsx`), evitando falsi positivi causati da documenti di aggiornamento normativo salvati nella stessa cartella.
  * **Supporto a liste di coordinate (`cell_refs`):** Gestione di mappature contenenti più celle da aggregare:
    * Scansione sequenziale delle celle indicate con tolleranza di riga (`row_tolerance`).
    * Concatenazione dei valori individuati tramite separatore di riga `"\n"`.
    * Fallback su ricerca dell'etichetta tramite `find_nearby_title_row` nel caso in cui le celle abbiano subito uno slittamento rispetto al modello standard.
* **Mappature Predefinite (`default_mappings.py`):**
  * Inserite le definizioni di estrazione per le nuove voci dal foglio `SPESA_PERSONALE` (`ufficioavvocatura`, `avvocatidirigenti`, `avvocatifunzionari`, `fondiavvocaturadirigenti`, `fondiavvocaturafunzionari`, `incarichilegaliesterni`, `patrociniolegale`, `consulenzalegale`, `servizilegali`).
  * Inserita la mappatura per `fondoperequativotari` (foglio `GEST. FIN. RESIDUI`, cella `I42`).
  * Ristrutturata la voce `contrastoallevasionetributaria` con `table_refs` multipli su fogli entrate alternativi (`SEZ. II - DATI ENTRATE`) per recuperare sia le entrate IMU/TASI sia le entrate TARSU/TIA/TARI.

---

### BDAP_Automation-0.3.0 (Transizione da v0.2.0)
*Scopo: Documentazione tecnica e standardizzazione operativa delle mappature.*

* **Manuale Mappature (`MAPPINGS.md`):**
  * Riorganizzazione della documentazione con spiegazione dettagliata del funzionamento delle mappature deterministiche (`sheet_idx`, `cell_ref`, `expected_label`, `label_check_cells`).
  * Indicazioni sulle modalità di fallback fuzzy in caso di layout BDAP difformi.
  * Checklist operativa di manutenzione per l'aggiunta di nuove voci contabili senza alterare i test pregressi.

---

### BDAP_Automation-0.3.1 (Transizione da v0.3.0)
*Scopo: Integrazione dei requisiti per packaging PyInstaller e canali di distribuzione.*

* **Dipendenze (`requirements.txt`):**
  * Aggiunto `PyInstaller>=6.0`.
* **Distribuzione (`README.md`):**
  * Inseriti i riferimenti al repository GitHub di build (`Fil0807/BDAP-Automation-Deployment`) e alla cartella cloud Google Drive contenente gli installer precompilati.

---

### BDAP_Automation-0.4.0 (Transizione da v0.3.1 - Major Release)
*Scopo: Introduzione del modulo Piano Indicatori, centralizzazione e formattazione dei commenti Excel, estrazione anagrafica comune e supporto alla lettura PDF.*

* **Nuovo Modulo `core/automation_piano_indicatori.py`:**
  * Gestione e compilazione del nuovo foglio **`Piano Indicatori`** inserito in `Template_Analisi.xlsx`.
  * Rilevamento automatico dei file ausiliari degli indicatori sintetici tramite `_discover_aux_workbook(bdap_path, "indicatori", year)` (file con pattern `*_ind_*` o `*indicatori*`).
  * Risoluzione per ciascun anno dei seguenti indicatori di bilancio:
    * **1.1:** Incidenza spese rigide su entrate correnti ($> 48\%$)
    * **2.8:** Incidenza incassi entrate proprie su previsioni definitive ($< 22\%$)
    * **3.2:** Anticipazioni chiuse solo contabilmente ($> 0$)
    * **4.1:** Incidenza spesa di personale su spesa corrente ($< 40\%$)
    * **10.3:** Sostenibilità debiti finanziari ($> 16\%$)
    * **12.4:** Sostenibilità disavanzo a carico dell'esercizio ($> 1,20\%$)
    * **13.1:** Debiti riconosciuti e finanziati ($> 1\%$)
    * **13.2 + 13.3:** Somma debiti in corso di riconoscimento e in corso di finanziamento ($> 0,60\%$)
    * **9.2:** Indicatore tempi medi di pagamento fornitori ($\le 0$)
* **Nuovo Modulo `support/comment_builder.py`:**
  * Centralizzazione e calcolo geometrico delle dimensioni dei commenti per prevenire il troncamento visivo su Microsoft Excel:
    * `width = max(450, min(750, max_len * 9 + 60))`
    * `height = max(170, len(lines) * 26 + 45)`
  * Formattazione esplicita delle coordinate di cella: `E45 (Riga 45, Colonna 5)`.
  * Funzione `fix_all_comments_in_workbook()`: applicata sistematicamente a tutti i fogli prima del salvataggio finale per reimpostare dimensioni minime garantite.
  * Funzione `extract_a1_breakdown()`: analizza l'Allegato A1 del rendiconto BDAP ed inserisce nel commento della quota accantonata il dettaglio completo dei singoli fondi (contenzioso, perdite partecipate, rinnovi contrattuali, altri accantonamenti).
* **Nuovo Modulo `support/conclusioni.py`:**
  * Compilazione del foglio `CONCLUSIONI`:
    * Cella `C2`: Nome dell'ente rilevato dai rendiconti o dalla struttura delle directory.
    * Cella `C3`: Popolazione residente dell'ultimo anno analizzato, rilevata tramite regex su relazioni revisori in formato PDF o dai prospetti BDAP.
* **Miglioramenti al Modulo FCDE (`automation_fcde.py`):**
  * Introduzione della ricerca estesa a tutto il foglio in caso di mancata corrispondenza della cella base, controllando le etichette nelle colonne A e B ed estraendo il dato contabile da colonna D.
  * Aggiunta funzione `read_fcde_totale_from_workbook()` per il recupero della somma residui attivi EP già valorizzata o calcolata dinamicamente.
* **Correzioni di Stabilità (`cli_discovery.py`):**
  * Inizializzazione protetta dei limiti massimi di ricerca (`max_row = 40`, `max_col = 12`) prima del controllo dell'etichetta *Denominazione ente* su cella A7, prevenendo cicli vuoti su workbook con `sheet.max_column = None`.
* **Formattazione Dati (`value_formatter.py`):**
  * Aggiunto supporto per la conversione di flag booleani/numerici ($1/0$) nelle stringhe italiane `"Sì"` e `"No"` (`boolean_yes_no=True`).
* **Dipendenze PDF (`requirements.txt`):**
  * Aggiunti i pacchetti `pymupdf>=1.24.0` e `pypdf>=4.0.0` per consentire l'ispezione documentale degli allegati PDF.

---

### BDAP_Automation-0.4.1 (Transizione da v0.4.0)
*Scopo: Creazione del foglio ANAGRAFICA, calcolo tempestività deliberazione consiliare, ristrutturazione foglio FCDE e robustezza multipiattaforma.*

* **Nuovo Modulo `core/automation_anagrafica.py`:**
  * Estrazione strutturata delle informazioni anagrafiche dai questionari di bilancio tramite ispezione congiunta dei fogli `ANAGRAFICA`/`SCHEDA_ANAGRAFICA` e del foglio `DATI`.
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
  * Eliminazione definitiva dei fogli `CONCLUSIONI` e `VERIFICA QUESTIONARI`.
  * Inserimento del foglio `ANAGRAFICA` come prima scheda del documento.
* **Revisione e Formule Foglio FCDE (`automation_fcde.py`):**
  * Griglia di calcolo estesa da 7 a **9 righe**:
    * Riga 2: Importo FCDE
    * Riga 3: Residui attivi Titolo I
    * Riga 4: Residui attivi Titolo III
    * Riga 5: Totale residui attivi (`=C3+C4`)
    * **Riga 6 (Nuova):** Quota avanzo disponibile da `ANALISI CONTABILE` riga 19.
      Scrittura dinamica della formula di collegamento: `='ANALISI CONTABILE'!<ColonnaAnno>19`.
    * Riga 7: `% rispetto a residui attivi conservati` (formula: `=C2*100/C5`).
    * **Riga 8 (Nuova):** `% FCDE + parte disponibile su residui attivi`:
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

---

### BDAP_Automation-0.4.2 (Transizione da v0.4.1)
* Release di allineamento documentale: rettificata la sintassi di formattazione HTML e la resa dei collegamenti nel file `README.md`.

---

### BDAP_Automation-0.4.3 (Transizione da v0.4.2 - Versione Corrente)
*Scopo: Allineamento visivo centrato dei dati contabili compilati, bonifica delle note esplicative per valori predefiniti, correzione coordinate contabili e completamento della documentazione in italiano.*

* **Centratura Automatica Dati (`automation.py` e `cli_pipeline.py`):**
  * Introdotta la funzione `format_inserted_data_cells_center(workbook_or_path, log=None)`, eseguita come fase conclusiva dell'intera pipeline:
    * Esclude i fogli di puro testo (`MAPPATURE_BDAP` e `CONTROLLI-POST`).
    * **Fogli contabili (`ANALISI CONTABILE`, `FCDE`, `Piano Indicatori`):**
      Rileva le colonne degli anni nelle intestazioni (righe 1..6) e applica l'allineamento centrato orizzontale e verticale (`horizontal="center", vertical="center"`) a tutte le celle con dati effettivi, preservando l'eventuale impostazione di testo a capo (`wrap_text`).
    * **Foglio `ANAGRAFICA`:**
      Centra le celle della colonna C nelle righe 5..10 (Dati generali) e le colonne B..H nelle tabelle delle deliberazioni e dei verbali.
* **Bonifica Commenti per Valori Predefiniti (`comment_builder.py`):**
  * La funzione `create_source_comment()` intercetta i token convenzionali (`DEFAULT_IF_MISSING`, `DEFAULT`, `CALCULATED`, `NONE`), evitando di stampare coordinate inesistenti nel commento della cella e inserendo la nota descrittiva standard:
    `voce non rilevata nella sorgente (valore predefinito)`.
  * Aggiunta la tracciabilità specifica per `fondivincolaticovid`: annota l'importo lordo derivato da *All. a2 col. O (Fondo solidarietà alimentare)*.
* **Correzione Coordinate Contabili (`default_mappings.py`):**
  * **`fondocontenzioso`:**
    * Coordinata primaria aggiornata da `I16` a `I17`.
    * Lista riferimenti di fallback impostata a `['I17', 'I16', 'I14']`.
    * Verifica etichetta estesa a `B17` e `B16`.
  * **`fondogaranziadebiticommerciali`:**
    * Riordinata la lista delle etichette con priorità per la forma maiuscola `FONDO GARANZIA DEBITI COMMERCIALI`.
    * Inserite le celle di fallback `I19` e `I23`.
* **Documentazione Tecnica e Tipizzazione del Codice:**
  * Riconversione sistematica in lingua italiana di tutti i docstring e dei commenti interni in `cli_discovery.py`, `cli_pipeline.py`, `value_resolver.py`, `automation_anagrafica.py`, `automation_fcde.py`, `automation_piano_indicatori.py`.
  * Aggiunti type hints espliciti sulle funzioni chiave (ad es. `tuple[Optional[object], Optional[str]]` per `find_bdap_value_by_label`).
