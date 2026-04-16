# Personal Python Scripts

Questa cartella contiene script Python personali e un launcher unico (`main.py`) per eseguirli rapidamente.

## Setup veloce

Setup in un comando:

```bash
./setup.sh
```

Oppure manuale:

1. Crea/attiva ambiente virtuale:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

2. Installa dipendenze Python:

```bash
pip install -r requirements.txt
```

3. Installa dipendenze di sistema (richieste da alcuni script):

```bash
brew install pandoc libreoffice
```

Per conversione PDF con `xelatex`, installa anche una distribuzione TeX (es. MacTeX).

## Uso di main.py

Elenca gli script disponibili:

```bash
python main.py list
```

Esegui uno script tramite chiave:

```bash
python main.py run md_to_pdf_batch -- --help
```

Nuovo convertitore unificato verso PDF:

```bash
python main.py run to_pdf_converter -- /percorso/input /percorso/output
```

GUI batch del convertitore:

```bash
python main.py run to_pdf_converter_gui
```

Misura distanza Terra-Luna giorno per giorno:

```bash
python main.py run earth_moon_distance_daily
python main.py run earth_moon_distance_daily -- --start-date 2026-04-16 --days 90 --time-utc 12:00 --output ~/Desktop/earth_moon_distance.csv
python main.py run earth_moon_distance_daily -- --plot --plot-output ~/Desktop/earth_moon_distance.png
python main.py run earth_moon_distance_daily -- --append-daily --start-date 2026-04-16
python main.py run earth_moon_distance_daily -- --gui
```

Esempi:

```bash
# Converte tutti i file supportati in una cartella, incluse sottocartelle
python main.py run to_pdf_converter -- ~/Documenti/input ~/Documenti/output_pdf

# Usa Pandoc per Markdown ed EPUB
python main.py run to_pdf_converter -- ~/Documenti/input ~/Documenti/output_pdf --prefer-pandoc

# Converte un singolo file DOCX
python main.py run to_pdf_converter -- ~/Documenti/file.docx ~/Documenti/output_pdf

# Apre la GUI batch per selezionare input, output e opzioni
python main.py run to_pdf_converter_gui
```

Note:
- Le chiavi sono generate automaticamente dai nomi file in `scripts/`.
- In alternativa puoi passare il nome file esatto (incluso `.py`).
- Tutti gli argomenti dopo `--` vengono inoltrati allo script scelto.
- `to_pdf_converter` supporta: `.docx`, `.wps`, `.epub`, `.txt`, `.rtf`, `.md`.
- `to_pdf_converter_gui` fornisce una GUI Tkinter per scansione batch, opzioni di conversione, barra di avanzamento e log.
- `earth_moon_distance_daily` calcola la distanza geocentrica Terra-Luna usando effemeridi JPL (`de421`) e salva un CSV giornaliero.
- `earth_moon_distance_daily --plot` genera anche un grafico PNG dell'andamento della distanza.
- `earth_moon_distance_daily --append-daily` aggiunge una sola riga al CSV per la data scelta (salta i duplicati).
- `earth_moon_distance_daily --gui` apre una GUI desktop (Tkinter) con campi input/output e opzioni plot/append.
- La GUI include il pulsante `Apri output` per aprire la cartella di destinazione nel Finder.
- La GUI include il pulsante `Apri report` per aprire il JSON finale generato dall'ultima conversione.
- La GUI mostra anche una tabella `File falliti` con il motivo sintetico degli errori di conversione.
- La GUI include i preset `Bilanciato`, `Qualita stampa` e `Conversione veloce`.
- Per `.md` ed `.epub` serve `pandoc` con un PDF engine come `xelatex`.
- Per `.docx`, `.wps`, `.txt` e `.rtf` serve `libreoffice` (`soffice`).
- Su macOS il convertitore prova anche il percorso standard `/Applications/LibreOffice.app/Contents/MacOS/soffice` se `soffice` non e presente nel `PATH`.

## Task VS Code

Sono disponibili task in `.vscode/tasks.json`:

- Python: List Personal Scripts
- Python: Run Personal Script

Uso:

1. Apri Command Palette e avvia Run Task.
2. Seleziona Python: List Personal Scripts per vedere le chiavi disponibili.
3. Seleziona Python: Run Personal Script e inserisci:
	- scriptKey: la chiave script mostrata dalla task list
	- scriptArgs: argomenti opzionali (esempio: `--help`)

## Changelog

### 2026-04-15

- Commit: `c0efef1`
- Aggiunto launcher centrale `main.py` con comandi `list` e `run`.
- Aggiornato `requirements.txt` con dipendenza Python (`pypdf`) e note su dipendenze di sistema.
- Aggiunte task VS Code in `.vscode/tasks.json` per usare rapidamente gli script.
- Importati e versionati gli script personali nella cartella `scripts/`.
