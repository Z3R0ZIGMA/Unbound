# Unbound

Windows desktop app that converts an Outlook OST file into a PST. See [PLAN.md](PLAN.md).

## Setup (Windows)

Requires Python 3.10 or newer (developed on 3.15). libpff-python ships prebuilt Windows wheels, so no C compiler is needed.

```bash
py -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

Check it works:

```bash
.venv\Scripts\python -c "import pypff; print(pypff.get_version())"
```

## Run the app

```bash
.venv\Scripts\python unbound_app.py
```

Needs Outlook Classic installed (it writes the PST). Version 1 converts mail only;
calendar, contacts and tasks folders are shown but greyed out. A log of anything
skipped or failed is saved next to the PST as `<name> - Unbound log.txt`.

## Build the EXE

```bash
.venv\Scripts\python -m pip install pyinstaller
.venv\Scripts\pyinstaller --noconfirm --clean --onefile --windowed --name Unbound --icon assets/icon.ico --add-data "assets/icon.ico;assets" --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32com.client --exclude-module tkinter unbound_app.py
```

The result is `dist\Unbound.exe`, a single file that runs without Python
(the PC still needs Outlook Classic). To check a build has all its libraries:

```bash
dist\Unbound.exe --check check.txt
```

Exit code 0 means everything loaded; `check.txt` lists each library.
