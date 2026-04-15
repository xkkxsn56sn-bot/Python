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

Note:
- Le chiavi sono generate automaticamente dai nomi file in `scripts/`.
- In alternativa puoi passare il nome file esatto (incluso `.py`).
- Tutti gli argomenti dopo `--` vengono inoltrati allo script scelto.

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
