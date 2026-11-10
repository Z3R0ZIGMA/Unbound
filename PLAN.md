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

1. Build/install libpff (pypff) on Windows — highest risk
2. Test reading a real, recent OST (Outlook 2013+ format support is partial)
3. Milestone: small app that opens a sample OST and lists its folders
4. Then prove writing one message into a PST via Outlook COM
