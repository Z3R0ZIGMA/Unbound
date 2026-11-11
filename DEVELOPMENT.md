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
.venv\Scripts\pyinstaller --noconfirm --clean --onefile --windowed --name Unbound --icon assets/icon.ico --add-data "assets/icon.ico;assets" --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32com.client --hidden-import win32timezone --exclude-module tkinter unbound_app.py
```

The result is `dist\Unbound.exe` (about 43 MB). `dist/` is git-ignored so builds don't bloat the repo's history; the EXE is published as a file on a [GitHub Release](https://github.com/Z3R0ZIGMA/Unbound/releases) instead. To publish a new version after rebuilding (with the [GitHub CLI](https://cli.github.com/), signed in):

```powershell
gh release create v1.2.1 dist\Unbound.exe --title "Unbound 1.2.1" --notes "What changed"
```

Rebuild right before publishing, so the EXE matches the pushed code. To check a build contains all its libraries and can pass dates to Outlook:

```powershell
Start-Process dist\Unbound.exe -ArgumentList '--check', 'check.txt' -Wait
Get-Content check.txt
```

Every line should start with `OK`.

**Hidden imports matter.** PyInstaller only bundles modules it can see being imported. pywin32 loads `win32timezone` on demand, the first time a date marked with a time zone goes through COM, so the build command lists it explicitly. Without it, 1.1.0 and 1.2.0 shipped an EXE where every calendar item failed and emails got the wrong received date, although the code worked fine from source. That's why the self-check passes a real date through COM, and why the EXE itself should be tested before a release (see [Testing](#testing)).

## Testing

The scripts in `scripts/` check conversions end to end by writing a PST through Outlook and reading it back with pypff. They need Outlook Classic. They print pass/fail per field and never print message content.

The OST tests need an OST to read. Outlook locks the OST of the profile on this PC while it runs, so close Outlook and copy the OST from `%LOCALAPPDATA%\Microsoft\Outlook` somewhere else first, then point the scripts at the copy. Delete the copy and the test PSTs afterwards; they contain your mail.

```powershell
# Apple Mail: builds a test export, converts it and checks every message (no OST needed)
.venv\Scripts\python scripts\mbox_test.py C:\Temp\unbound_test

# OST: folder tree and a read of every item
.venv\Scripts\python scripts\list_folders.py C:\Temp\copy.ost --sample 100000

# OST: one email with attachments, every field compared
.venv\Scripts\python scripts\copy_one_message.py C:\Temp\copy.ost C:\Temp\one.pst

# OST: repeating and all-day appointments, a meeting and a contact; every property and the recurrence compared
.venv\Scripts\python scripts\copy_items_test.py C:\Temp\copy.ost C:\Temp\items.pst
```

Known results: `copy_one_message.py` fails "internet headers" (Outlook doesn't allow writing them), and `copy_items_test.py` fails "time zone" for appointments with an IANA zone name such as `America/New_York`. Everything else should pass.

### Testing the EXE itself

Code that works from source can still fail in the EXE if PyInstaller left something out, so test the built EXE before each release. `Unbound.exe` has a hidden mode that converts without the window (the folders ticked by default), writes the usual log next to the PST, and exits with 0 if nothing failed:

```powershell
# Apple Mail: convert the test export with the EXE, then check every message
.venv\Scripts\python scripts\make_mbox_fixture.py C:\Temp\exe_test\export
Start-Process dist\Unbound.exe -ArgumentList '--convert', 'C:\Temp\exe_test\export', 'C:\Temp\exe_test\mbox_test.pst' -Wait
.venv\Scripts\python scripts\mbox_test.py C:\Temp\exe_test --verify-only

# OST: convert with the EXE, then compare every email, appointment and contact with the source
Start-Process dist\Unbound.exe -ArgumentList '--convert', 'C:\Temp\copy.ost', 'C:\Temp\exe_test\ost.pst' -Wait
.venv\Scripts\python scripts\compare_pst.py C:\Temp\copy.ost C:\Temp\exe_test\ost.pst
```

`compare_pst.py` reports how many items in the PST are identical to the source (emails: subject, sender, times, read state, body type, attachments, recipients; appointments: times and full recurrence; contacts: name and email), and which field differs when one isn't. Times are compared to the second.

A quick static check catches undefined names and leftover imports:

```powershell
.venv\Scripts\python -m pip install pyflakes
.venv\Scripts\python -m pyflakes unbound unbound_app.py scripts
```

## Project layout

| Path | What it is |
|---|---|
| `dist/Unbound.exe` | The built app (not in git; created by the build above) |
| `unbound_app.py` | Entry point (also the `--check` self-test) |
| `unbound/app.py` | The window (PySide6, dark theme; colours are in `COLORS` at the top) |
| `unbound/converter.py` | Folder scan and the conversion loop (OST, PST or MBOX → PST), log writing |
| `unbound/ost_reader.py` | Reads messages out of an OST or PST with pypff |
| `unbound/mbox_reader.py` | Finds the mailboxes in an Apple Mail export and reads MBOX messages |
| `unbound/pst_writer.py` | Writes mail, appointments and contacts into the PST through Outlook |
| `unbound/props.py` | Reads every property of an item, resolving named properties via the file's name-to-ID map |
| `unbound/appointment.py` | Decodes recurrence patterns and time zones, which Outlook only accepts through its object model |
| `unbound/mapi.py` | MAPI property tags shared by reader and writer |
| `scripts/list_folders.py` | Command-line check: prints an OST's folder tree and test-reads items |
| `scripts/copy_one_message.py` | Command-line check: copies one message to a PST and compares every field |
| `scripts/copy_items_test.py` | Command-line check: copies sample appointments and a contact to a PST and compares every property and the recurrence |
| `scripts/make_mbox_fixture.py` | Builds a test Apple Mail export covering the tricky cases |
| `scripts/mbox_test.py` | End-to-end MBOX check: builds the test export, converts it and verifies every message (`--verify-only` checks a PST made by the EXE) |
| `scripts/compare_pst.py` | Compares a converted PST with its source, item by item |
| `assets/make_icon.py` | Draws the app icon; re-run it after changing the colours, then rebuild |
