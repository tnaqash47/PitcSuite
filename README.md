# PITC Suite

Run `python Pitcsuite.py`, or open `dist/pitcsuite.exe` after building.

The collapsible left tools panel contains all 13 tools. PDF Merger replaces the former PDF Merge and Multi-Merge tools with the integrated PDFMerger project: single-folder or up to five folders, numeric ordering, duplex/N-up blank-page padding, filename stamping, progress, and cancellation.

The suite uses PDFMerger's Fusion dark palette, larger system text, green primary actions, and orange secondary actions. Tool forms scroll when needed on smaller screens.

PDF Merger requires PySide6>=6.10,<7, pypdf>=6.6,<7, and reportlab>=4.4,<5, in addition to the existing suite dependencies.

Validation: `python -m unittest discover -s tests -v`

Build: `python -m PyInstaller --noconfirm pitcsuite.spec`
