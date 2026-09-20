# PITC Suite

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Configure credentials:
   - Copy `config.example.ini` to `config.ini`
   - Fill in your credentials, or configure them directly via the in-app **Settings** panel.

## Run

Run `python Pitcsuite.py`, or open `dist/pitcsuite.exe` after building.

The collapsible left tools panel contains all 13 tools. PDF Merger replaces the former PDF Merge and Multi-Merge tools with the integrated PDFMerger project: single-folder or up to five folders, numeric ordering, duplex/N-up blank-page padding, filename stamping, progress, and cancellation.

The suite uses PDFMerger's Fusion dark palette, larger system text, green primary actions, and orange secondary actions. Tool forms scroll when needed on smaller screens.

## Validation & Build

Validation:
```bash
python -m unittest discover -s tests -v
```

Build executable:
```bash
python -m PyInstaller --noconfirm pitcsuite.spec
```
