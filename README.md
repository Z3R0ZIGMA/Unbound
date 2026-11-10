# Unbound

Windows desktop app that converts an Outlook OST file (for example one copied from another PC) into a PST. It reads the OST directly with [libpff](https://github.com/libyal/libpff), so it needs neither the original PC nor its Outlook profile, and writes the PST through Outlook Classic on this PC. See [PLAN.md](PLAN.md) for the design and roadmap.

## Download

A ready-built **[dist/Unbound.exe](dist/Unbound.exe)** is included in this repo. It's a single file that runs without installing anything else (apart from Outlook Classic, below). On GitHub, open the file and click **Download raw file**. Windows SmartScreen may warn the first time because the EXE isn't signed; choose **More info → Run anyway**.

## Requirements

- Windows
- **Outlook Classic** installed on the PC doing the conversion. It writes the PST. The New Outlook can't, because it has no automation (COM) support.
- Python 3.10 or newer, only if running from source (developed on 3.15)

## Using it

1. Start `dist\Unbound.exe` (or run from source, below).
2. Click **Browse…** and choose the OST file. The folder tree appears with item counts.
3. Tick the folders to convert. Your mail folders start ticked; hidden system folders and **Shared mailboxes** start unticked. Ticking a folder ticks its subfolders too. **Select all mail** puts the ticks back to that starting selection, and **Clear** unticks everything.
4. Click **Convert to PST** and choose where to save the PST. Save it somewhere other than Outlook's own data folder, such as Documents.
5. Wait for the progress bar. You can cancel at any time; what's already copied stays in the PST.

When it finishes, a log is saved next to the PST as `<name> - Unbound log.txt`. It lists what was copied, skipped and failed, plus the folder and item types found in the file.

Outlook opens in the background during conversion (if it isn't already open) and stays open afterwards. The new PST is not left attached to your Outlook profile; to look at it, use **File → Open & Export → Open Outlook Data File** in Outlook.

## Limitations (version 1)

- **Mail only.** Calendar, contacts, tasks, notes and journal folders are shown greyed out. Meeting invites and other non-mail items inside mail folders are skipped and counted in the log.
- **No internet headers.** Outlook doesn't let other programs write them, so the raw `Received:`/routing headers aren't copied. Subject, sender, recipients, dates, read state, bodies and attachments are.
- **Speed.** About 0.15 seconds per message through Outlook, so a 28,000-message mailbox takes roughly 70 minutes.
- **Outlook's reminders window** may appear when Outlook starts. Leave Outlook running until the conversion finishes.
- **IMAP OSTs** are supported but not yet tested on a real file. Exchange/Microsoft 365 OSTs are tested.

## Run from source

In PowerShell or Command Prompt, from this folder:

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
| `dist/Unbound.exe` | The built app, ready to run (see [Download](#download)) |
| `unbound_app.py` | Entry point (also the `--check` self-test) |
| `unbound/app.py` | The window (PySide6, dark theme; colours are in `COLORS` at the top) |
| `unbound/converter.py` | Folder scan and the OST → PST conversion loop, log writing |
| `unbound/ost_reader.py` | Reads messages out of the OST with pypff |
| `unbound/pst_writer.py` | Writes messages into the PST through Outlook |
| `unbound/mapi.py` | MAPI property tags shared by reader and writer |
| `scripts/list_folders.py` | Command-line check: prints an OST's folder tree and test-reads items |
| `scripts/copy_one_message.py` | Command-line check: copies one message to a PST and compares every field |
| `assets/make_icon.py` | Draws the app icon; re-run it after changing the colours, then rebuild |
