# Mappature deterministiche

Questo documento descrive le voci configurate in `bdap_app/support/default_mappings.py`.
Le mappature servono a ridurre ricerche ambigue nei workbook BDAP: per ogni voce del template si indica da quale workbook leggere, quale foglio usare, quale cella o strategia applicare e quale etichetta validare.

## Come leggere una mappatura

Ogni voce del dizionario `TEMPLATE_SOURCES` usa come chiave una versione normalizzata della voce nel template. Il valore è un dizionario di regole. Esempio:

```python
'parteaccantonatab': {
    'cell_ref': 'B33',
    'expected_label': 'Totale parte accantonata',
    'label_check_cells': ['A33'],
    'optional': True,
    'row_tolerance': 2,
    'sheet_idx': '11',
    'source_workbook': 'rendiconto',
    'strict_expected_label': False,
}
```

In questo caso il programma:

1. usa il rendiconto BDAP dell'anno elaborato;
2. apre il foglio 11;
3. controlla l'etichetta attesa vicino ad `A33`;
4. legge il valore da `B33`;
5. scrive nel template il valore e il commento di tracciabilità.

## Sorgenti supportate

| `source_workbook` | Significato | Note operative |
| --- | --- | --- |
| `rendiconto` | Workbook BDAP annuale principale, normalmente `Rend. YYYY.xlsx`. | È la sorgente più frequente. Usa `sheet_idx`, `sheet_name_contains`, `cell_ref`, `cell_refs` o ricerche per etichetta. |
| `questionario` | Questionario Bilancio/Consuntivo dell'anno. | Il file deve contenere `questionario` e anno nel nome ed essere vicino al rendiconto; alcune regole controllano anche le intestazioni del workbook. |
| `questionario_debiti` | Questionario Debiti Fuori Bilancio. | Il nome deve richiamare questionario e debiti fuori bilancio. |
| `indicatori` | Workbook degli indicatori o file sintetici. | Sono accettati nomi con `indicatori`, `ind`, `sintetico` o `sintetici`; l'anno deve essere nel nome o nelle prime righe. |
| `analysis` | Workbook di analisi in compilazione. | Usato per valori già calcolati in altri fogli del template, ad esempio `FCDE`. |
| `relazione` | Voce da relazione o PDF. | Al momento il flusso lascia la cella vuota: è una voce da gestire manualmente o con una futura estrazione da PDF. |
| assente | Nessuna sorgente automatica definita. | La voce resta non risolta o viene gestita da logiche specifiche/fallback se presenti. |

## Come aggiungere una nuova voce

1. Individuare la voce nel template e ricavare la chiave normalizzata usata dal codice.
2. Stabilire la sorgente (`rendiconto`, `questionario`, `indicatori`, ecc.).
3. Preferire un riferimento deterministico (`sheet_idx`/`sheet_name_contains` + `cell_ref`) quando la struttura BDAP è stabile.
4. Aggiungere `expected_label` e `label_check_cells` quando possibile: servono a evitare di leggere una cella corretta solo per posizione ma sbagliata per contenuto.
5. Usare `row_tolerance` solo se la riga può spostarsi tra versioni del file.
6. Impostare `optional=True` per voci che possono mancare senza bloccare l'elaborazione.
7. Se il valore è una percentuale, impostare `is_percentage=True`.
8. Aggiornare questa tabella e aggiungere o adattare un test mirato.

## Checklist di manutenzione

- Verificare che il foglio sia identificabile in modo stabile: usare `sheet_idx` solo se l'ordine dei fogli non cambia; in caso contrario preferire `sheet_name_contains`.
- Evitare ricerche globali quando è disponibile una cella precisa.
- Non usare `strict_expected_label=True` se la dicitura cambia spesso tra anni o versioni BDAP.
- Documentare qui le voci che restano manuali, così l'utente sa che non sono errori di automazione.
- Dopo ogni nuova mappatura, eseguire i test e provare almeno un workbook reale o demo.

## Elenco mappature

| Key | Cell / Cell refs | Sheet (idx or contains) | Source workbook | Expected label / Notes |
|---|---:|:---|:---|---|
| `prospettoevoluzionerisultatoammne` | D24 | 11 | rendiconto | Risultato di amministrazione al 31 dicembre |
| `parteaccantonatab` | B33 | 11 | rendiconto | Totale parte accantonata |
| `dicuifcde` | B27 | 11 | rendiconto | Fondo crediti di dubbia esigibilità al 31/12 |
| `partevincolatac` | B40 | 11 | rendiconto | Totale parte vincolata |
| `partedestinataagliinvestimentid` | B42 | 11 | rendiconto | Totale parte destinata agli investimenti |
| `partedisponibileeabcd` | B43 | 11 | rendiconto | Totale parte disponibile — calcolata se mancante (E = A - B - C - D) |
| `prospettocassa` | D17 | 11 | rendiconto | Fondo di cassa al 31 dicembre |
| `importocassavincolata` | D8 (year cell C6) | sheet name contains `SEZ. II-GEST. FIN. CASSA` | questionario | di cui cassa vincolata |
| `debitifuoribilancio` | P19 | Sezione preliminare | questionario_debiti | 4. Sono stati riconosciuti debiti fuori bilancio nel |
| `fcdeinrapportoaresiduiattivi` | I6 | sheet name contains `FCDE` | analysis | FCDE in rapporto a residui attivi (is_percentage=False) |
| `fpvpartecorrenteal3112` | D21 | 11 | rendiconto | Fondo pluriennale vincolato per spese correnti |
| `fpvpartecapitaleal3112` | D22 | 11 | rendiconto | Fondo pluriennale vincolato per spese in conto capitale |
| `fondocontenzioso` | I16 | 12 | rendiconto | Totale Fondo contenzioso |
| `potenzialipassivita` | I26 | 12 | rendiconto | FONDO ACCANTONAMENTO PASSIVITA POTENZIALI |
| `fondogaranziadebiticommerciali` | I23 | 12 | rendiconto | Totale Fondo di garanzia debiti commerciali (only_fill_if_found) |
| `fondivincolaticovid` | (sum cells in col O) | 13 | rendiconto | Trasferimenti correnti … COVID-19 — somma colonna `O` |
| `rispettoequilibri` | B69 | 7 | rendiconto | W1) Risultato di competenza |
| `equilibriodibilanciow2` | B72 | 7 | rendiconto | W2) Equilibrio di bilancio |
| `equilibriocomplessivow3` | B74 | 7 | rendiconto | W3) Equilibrio complessivo |
| `fondocassadaidatisiope` | - | - | - | Fondo cassa dai dati SIOPE (no explicit cell) |
| `anticipazionetesoreria` | B31 | 6 | rendiconto | Titolo 7: Anticipazioni da istituto tesoriere/cassiere |
| `fal` | B29 | 11 | rendiconto | Fondo anticipazioni liquidità |
| `contrazionenuovimutui` | E25 (fallback F25) | sheet name contains `SEZ. IV -INDEBITAMENTO_DATI _` | questionario | 3) Debito complessivo contratto (static_column, check_year_in_label) |
| `accantonamentocifrarelativaairinnovicontrattualidelccnl` | I26 | 12 | rendiconto | FONDO RINNOVI CONTRATTUALI |
| `spesadelpersonalesostenuta` | - | - | relazione | Spesa del personale sostenuta (no explicit cell) |
| `incidenzaspesapersonalesuspesacorrente` | D27 | 1 | indicatori | Incidenza della spesa di personale sulla spesa corrente |
| `%riscossionisanzionidelcodicedellastrada` | C20 / C24 / D25 / B32 | sheet name contains `SEZ. II - DATI ENTRATE` or `SEZ. II- Gest. ENTRATE` | questionario | Percentuale di riscossione (is_percentage=True) |
| `%riscossioniproventidapermessoacostruire` | - | - | relazione | is_percentage=True |
| `contrastoallevasionetributaria` | - | - | questionario | - |
| `%diriscossionecomplessiva` | - | - | - | is_percentage=True |
| `residuiattividatatiprecedenti2021_titolo_uno` | C71 | sheet name contains `SEZ. II -DATI RISULT. FINAN` | questionario | display_refs [('TITOLO I','C71')], year cells C69..G69 |
| `residuiattividatatiprecedenti2021_titolo_tre` | C73 | sheet name contains `SEZ. II -DATI RISULT. FINAN` | questionario | display_refs [('TITOLO III','C73')], year cells C69..G69 |
| `residuiattividatatiprecedenti2021_titolo_quattro` | C74 | sheet name contains `SEZ. II -DATI RISULT. FINAN` | questionario | display_refs [('TITOLO IV','C74')], year cells C69..G69 |
| `residuipassividatati` | C83 | sheet name contains `SEZ. II -DATI RISULT. FINAN` | questionario | display_refs [('TITOLO I','C83')], year cells C81..G81 |


## Chiavi supportate in `bdap_app/support/default_mappings.py`

Di seguito le chiavi che possono comparire in ogni voce del dizionario di mappatura e il loro significato.

| Chiave | Tipo | Descrizione | Esempio |
|---|---|---|---|
| `cell_ref` | string | Riferimento singolo alla cella contenente il valore. | `B33` |
| `cell_refs` | list[string] | Lista di riferimenti cella quando il valore è distribuito su più celle. | `['B31','C31']` |
| `fallback_cell_refs` | list[string] | Riferimenti alternativi da provare se `cell_ref` è vuota/assente. | `['F25']` |
| `year_cell_ref` | string or list[string] | Cella (o celle) che contengono gli anni di intestazione per tabelle orizzontali. | `C6` o `['C69','D69',...]` |
| `expected_label` | string or list[string] | Testo atteso vicino al valore per validazione; può essere una stringa o una lista di possibili etichette. | `'Totale parte accantonata'` |
| `strict_expected_label` | bool | Se True richiede match esatto con `expected_label`; altrimenti usa confronto più permissivo. | `True` |
| `label_check_cells` | list[string] | Celle (o intervalli) da verificare per individuare la label prima di leggere la riga di valore. Utile quando la label può spostarsi di poche righe o occupare più colonne. | `['A25','A25:C25']` |
| `optional` | bool | Se True la voce è opzionale e l'assenza non genera errore. | `True` |
| `row_tolerance` | int or float | Numero di righe da cercare (sopra/sotto) quando si cerca la label o la cella di riferimento. | `2` o `2.0` |
| `sheet_idx` | string or int | Indice del foglio (1-based) da cui leggere il valore. | `'11'` |
| `sheet_name_contains` | string | Sottostringa da cercare nel nome del foglio (utile quando il foglio non ha indice fisso). | `'SEZ. II -DATI RISULT. FINAN'` |
| `source_workbook` | string | Nome logico del workbook BDAP dove cercare (`rendiconto`, `questionario`, `analysis`, `relazione`, `indicatori`, ...). | `rendiconto` |
| `search_key_prefix` | string | Prefisso usato nella ricerca per etichetta quando la label contiene un identificatore (es. `W1)`). | `'W1)'` |
| `is_percentage` | bool | Indica che il valore è una percentuale e potrebbe necessitare di formattazione/normalizzazione diversa. | `True` |
| `sum_cells` | list[string] | Colonne o celle da sommare per ottenere il valore, può essere una colonna (`['O']`) o celle specifiche (`['O17','O35']`). | `['O']` |
| `compute_if_missing` | dict | Istruzioni per calcolare il valore se mancante. Tipicamente contiene `base` (chiave del valore base) e `subtract` (lista di chiavi da sottrarre). | `{'base':'risultatoamministrazionea','subtract':['parteaccantonatab',...]}` |
| `display_refs` | list[tuple] | Riferimenti da usare per la visualizzazione (label, cell). | `[('TITOLO I','C71')]` |
| `fallback_display_refs` | list[dict] | Riferimenti alternativi da mostrare o usare quando la tabella principale non contiene il valore. | `{'title':'TITOLO I','cell_ref':'F6'}` |
| `only_fill_if_found` | bool | Se True scrive il valore solo se la label viene trovata. | `True` |
| `static_column` | bool | Indica che la voce si trova sempre in una colonna fissa (utile per tabelle con colonna statica). | `True` |
| `check_year_in_label` | bool | Se True verifica che l'anno sia presente nella label (utile per valori per anno). | `True` |
| `table_refs` | list[dict] | Elenco di coppie `cell_ref`/`year_cell_ref` per layout alternativi dello stesso questionario. | `{'cell_ref':'D8','year_cell_ref':'D6'}` |
| `prefer_exact_table_workbook` | bool | Se True preferisce il questionario che contiene esattamente l'anno richiesto nella tabella. | `True` |
| `prefer_latest_table_workbook` | bool | Se True preferisce il questionario con tabella più recente quando più file sono candidati. | `True` |
| `default_if_missing` | number/string | Valore da usare quando la voce non viene trovata ma deve essere comunque compilata. | `0` |

Note:
- `expected_label` può essere una lista di alternative quando la stessa voce è chiamata in modi diversi nei rendiconti.
- Le celle e gli intervalli possono essere specificati anche come range Excel (`A25:C25`) quando necessario.
- `sheet_idx` è 1-based (foglio 1 = primo foglio nel workbook).




## Voci Non Auto-Compilate

Queste voci sono lasciate intenzionalmente vuote o non hanno una sorgente automatica affidabile, perché richiedono relazione, PDF, dati SIOPE o verifica manuale:

- FONDO CASSA DAI DATI SIOPE 1450 - FONDO DI CASSA DELL'ENTE ALLA FINE DEL PERIODO DI RIFERIMENTO - QUOTA VINCOLATA
- SPESA DEL PERSONALE art. 1 comma 557
- % RISCOSSIONI PROVENTI DA PERMESSO A COSTRUIRE
- CONTRASTO ALL'EVASIONE TRIBUTARIA

Quando una di queste voci viene automatizzata in futuro, spostarla nell'elenco mappature, indicare la sorgente usata e aggiungere un test dedicato.
