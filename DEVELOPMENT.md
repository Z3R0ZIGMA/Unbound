# Developing Unbound

How to run Unbound from source, rebuild the EXE, and find your way around the code. For using the app, see [README.md](README.md); for the design and roadmap, see [PLAN.md](PLAN.md).

## Run from source

Needs Python 3.10 or newer (developed on 3.15) and Outlook Classic. In PowerShell or Command Prompt, from this folder:

```powershell
py -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python unbound_app.py
```

libpff-python ships prebuilt Windows wheels, so no C compiler is needed.

## Build the EXE

```powershell
.venv\Scripts\python -m pip install pyinstaller
.venv\Scripts\pyinstaller --noconfirm --clean --onefile --windowed --name Unbound --icon assets/icon.ico --add-data "assets/icon.ico;assets" --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32com.client --exclude-module tkinter unbound_app.py
```

This replaces `dist\Unbound.exe`. The EXE is committed to the repo, so after changing the code, rebuild it and commit it together with the code change so the two stay in step. Each committed build adds about 43 MB to the repo's history.

To check a build contains all its libraries:

```powershell
Start-Process dist\Unbound.exe -ArgumentList '--check', 'check.txt' -Wait
Get-Content check.txt
```

Every line should start with `OK`.

## Project layout

| Path | What it is |
|---|---|
| `dist/Unbound.exe` | The built app, ready to run |
| `unbound_app.py` | Entry point (also the `--check` self-test) |
| `unbound/app.py` | The window (PySide6, dark theme; colours are in `COLORS` at the top) |
| `unbound/converter.py` | Folder scan and the OST → PST conversion loop, log writing |
| `unbound/ost_reader.py` | Reads messages out of the OST with pypff |
| `unbound/pst_writer.py` | Writes messages into the PST through Outlook |
| `unbound/mapi.py` | MAPI property tags shared by reader and writer |
| `scripts/list_folders.py` | Command-line check: prints an OST's folder tree and test-reads items |
| `scripts/copy_one_message.py` | Command-line check: copies one message to a PST and compares every field |
| `assets/make_icon.py` | Draws the app icon; re-run it after changing the colours, then rebuild |
