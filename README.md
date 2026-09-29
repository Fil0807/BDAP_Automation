# Automazione della compilazione di file Excel di analisi contabile utilizzando i dati BDAP dei comuni

Fork (pubblico) del progetto https://github.com/Elisabetta-43/Stage_Locatelli con modifiche per adattare il progetto alla conversione a installer Windows e MacOS


Sono state apportate modifiche a file di ricerca come ```./bdap_app/orchestration/cli_discovery.py``` per migliorare la gestione dei Path di alcuni componenti e sono stati rimossi alcuni file non necessari come le immagini di spiegazione e diagrammi


<br><br>

# Comandi fase testing
In tutti i passaggi è consigliato la creazione di un ambiente di sviluppo tramite i seguenti comandi:

Creazione ambiente virtuale:
```bash
python -m venv .venv
```

<br>

Attivazione ambiente virtuale

MacOS/Linux:  
```bash
source .venv/bin/activate
```

Windows:  
```bash
.\.venv\Scripts\Activate.ps1
```

<br>

Installazione/Aggiornamento dei requisiti:
```bash
pip install -r requirements.txt
```

<br>

Avvio dell'applicazione:
```bash
python -m bdap_app
```
