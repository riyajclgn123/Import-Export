# Firestore Import-Export

A Python command-line tool that moves Cloud Firestore data to and from Excel, and makes timestamped JSON backups.

![Python](https://img.shields.io/badge/Python-3-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-2.2-150458?logo=pandas&logoColor=white)
![Firebase](https://img.shields.io/badge/Cloud%20Firestore-FFCA28?logo=firebase&logoColor=black)
![Node.js](https://img.shields.io/badge/Node.js-npx-339933?logo=nodedotjs&logoColor=white)

## Overview

Editing Firestore data by hand in the Firebase console is slow. This tool lets you export a whole Firestore database to an Excel workbook, edit it in a spreadsheet, and import it back. Each top-level collection becomes one worksheet, and each document becomes one row. Before you make changes, you can save a timestamped JSON backup and restore it later.

## Features

- **Export (Firestore to Excel).** Dumps the database to JSON, then writes one sheet per collection to `Exported/data.xlsx`. Document IDs go in the first column.
- **Import (Excel to Firestore).** Reads every sheet in the workbook, turns it back into Firestore's collection and document JSON, and uploads it.
- **Keeps nested data.** Cells holding JSON-like strings (`{...}` or `[...]`) are parsed back into objects or arrays, and empty cells become empty strings instead of `NaN`.
- **Backups.** Saves the database to `backup/backup_YYYYMMDD_HHMMSS.json`, and can restore any backup file from the menu.
- **Menu-driven.** `Main.py` installs the dependencies and shows a numbered menu.

## How it works

```
                 firestore-export                pandas + openpyxl
Firestore  ───────────────────────►  JSON  ───────────────────────►  Excel (.xlsx)
Firestore  ◄───────────────────────  JSON  ◄───────────────────────  Excel (.xlsx)
                 firestore-import                pandas
```

The Python scripts call the [`node-firestore-import-export`](https://www.npmjs.com/package/node-firestore-import-export) CLI through `npx` to read and write Firestore. The scripts handle the JSON-to-Excel conversion themselves. All file paths are read from `config.properties`.

| File | Role |
| --- | --- |
| `Main.py` | Installs requirements and shows the menu: backup, import, export, restore backup |
| `Export.py` | Exports Firestore to JSON, then to Excel. `--backup` saves a timestamped JSON only |
| `Import.py` | Converts Excel to JSON and imports it. If you pass a JSON file as an argument, it imports that file directly |
| `config.properties` | Paths for the Excel, JSON and backup files and the Firebase credentials |

## Tech stack

Python 3, pandas, openpyxl, NumPy, configparser and argparse. Node.js (`npx`) with `node-firestore-import-export`. Google Cloud Firestore.

## Getting started

### Prerequisites

- Python 3 and pip (the `python` command must be on your PATH)
- Node.js, so that `npx` is available
- A Firebase **service account key** (JSON) for your project

### Setup

```bash
git clone https://github.com/riyajclgn123/Import-Export.git
cd Import-Export
pip install -r requirements.txt     # Main.py also does this automatically
```

Save your service account key at the path set in `config.properties` (default: `AppConfig/appConfig.json`). The `.gitignore` excludes `*.json` and credential files, so the key will not be committed.

### Run

```bash
python Main.py
```

```
1. Create Backup For Safety
2. Import Data (Excel to Firestore)
3. Export Data (Firestore to Excel)
4. Import Backup File
```

You can also run the scripts directly:

```bash
python Export.py              # Firestore -> Exported/data.xlsx
python Export.py --backup     # Firestore -> backup/backup_<timestamp>.json
python Import.py              # Exported/data.xlsx -> Firestore
python Import.py backup/<file>  # restore a backup file
```

> Importing writes to your live Firestore database. Make a backup first.

## Project structure

```
Import-Export/
├── Main.py             menu entry point
├── Export.py           Firestore -> JSON -> Excel, and backups
├── Import.py           Excel -> JSON -> Firestore
├── config.properties   file paths and credentials location
├── requirements.txt
├── Exported/           generated Excel workbook
└── backup/             JSON backups
```

## Author

**Riyaj Chaulagain**, Computer Science student at Southeastern Louisiana University

- GitHub: [riyajclgn123](https://github.com/riyajclgn123)
- Portfolio: [riyajchaulagain.com.np](https://riyajchaulagain.com.np)
