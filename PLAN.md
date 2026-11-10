# OST to PST Converter – Project Plan

## Goal

A Windows desktop app that converts an Outlook OST file (copied from another machine) into a PST file, using Outlook Classic on the local machine to write the PST.

## Version 1 features

- Select an OST file
- Preview folder tree with item counts
- Choose which folders to export
- Convert to PST with a progress bar
- Save a log of items that failed to convert

## Architecture

### 1. Reader (OST parsing)

- Uses libpff via its Python bindings (pypff)
- Walks folders, extracts messages (headers, body, HTML/RTF), attachments
- Does not need Outlook or the original profile

### 2. Writer (PST creation)

- Uses Outlook Classic via COM automation (pywin32)
- Creates a new PST store, recreates folder structure, adds items
- Note: New Outlook does not support COM automation

### 3. Interface

- PySide6 desktop GUI
- Packaged as a single Windows executable with PyInstaller

## Interface layout

- Dark mode throughout (dark window, light text)
- Top: file picker to select the OST file
- Middle: folder tree with item counts and checkboxes to choose folders
- Bottom: large Convert button with a progress bar
- Fully customisable later (colours, logo, layout) by asking Claude Code in plain words

## Later additions

- Contacts and calendar items
- Date range filters
- Splitting very large PSTs
- Apple Mail (Mac) MBOX to PST conversion: read MBOX with Python's built-in `mailbox` module, reuse the Outlook Classic writer

## Risks and first steps

1. ~~Build/install libpff (pypff) on Windows — highest risk~~ Done: `libpff-python` 20261030 installs from a prebuilt Windows wheel (no compiler needed)
2. ~~Test reading a real, recent OST (Outlook 2013+ format support is partial)~~ Done with `scripts/list_folders.py` on two current Microsoft 365 OSTs (96 MB and 5.6 GB): every message, calendar item, contact and task read with no failures (28,292 in the large file, about 8 minutes). Findings for the reader:
   - Open files with an absolute, backslash path (`os.path.abspath`); libpff's `\\?\` long-path handling rejects forward slashes
   - pypff raises `OSError` when a message has no HTML or RTF body; treat that as "absent", not a failure
   - Skip non-message items (`pypff.item`, e.g. the hidden `PersonMetadata` and `Help` folders)
3. Milestone: small app that opens a sample OST and lists its folders
4. Then prove writing one message into a PST via Outlook COM
