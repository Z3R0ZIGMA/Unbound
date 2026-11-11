# Unbound

Windows desktop app that converts an Outlook OST file (for example one copied from another PC) into a PST. It reads the OST directly with [libpff](https://github.com/libyal/libpff), so it needs neither the original PC nor its Outlook profile, and writes the PST through Outlook Classic on this PC. See [PLAN.md](PLAN.md) for the design and roadmap.

## Download

Unbound runs as a single file, `Unbound.exe`, that needs nothing else installed apart from Outlook Classic (below). The EXE isn't stored in this repo, to keep it small; build it with the steps in [DEVELOPMENT.md](DEVELOPMENT.md#build-the-exe), which put it at `dist\Unbound.exe`. Windows SmartScreen may warn the first time because the EXE isn't signed; choose **More info → Run anyway**.

## Requirements

- Windows
- **Outlook Classic** installed on the PC doing the conversion. It writes the PST. The New Outlook can't, because it has no automation (COM) support.

## Using it

1. Start `Unbound.exe`.
2. Click **Browse…** and choose the OST file. The folder tree appears with item counts.
3. Tick the folders to convert. Your mail folders start ticked; hidden system folders and **Shared mailboxes** start unticked. Ticking a folder ticks its subfolders too. **Select all mail** puts the ticks back to that starting selection, and **Clear** unticks everything.
4. Click **Convert to PST** and choose where to save the PST. Save it somewhere other than Outlook's own data folder, such as Documents.
5. Wait for the progress bar. You can cancel at any time; what's already copied stays in the PST.

When it finishes, a log is saved next to the PST as `<name> - Unbound log.txt`. It lists what was copied, skipped and failed, plus the folder and item types found in the file.

**Converting this PC's own OST.** An OST in Outlook's data folder (`%LOCALAPPDATA%\Microsoft\Outlook`) may belong to the Outlook profile on this PC, and Outlook locks it while running. Unbound first makes a temporary working copy of such a file and converts from that, so it needs free disk space equal to the OST's size; the copy is deleted afterwards. If you see **"Outlook is using this file"**, close Outlook, wait a few seconds, and try again. OSTs copied from another PC aren't affected.

Outlook opens in the background during conversion (if it isn't already open) and stays open afterwards. The new PST is not left attached to your Outlook profile; to look at it, use **File → Open & Export → Open Outlook Data File** in Outlook.

## Limitations (version 1)

- **Mail only.** Calendar, contacts, tasks, notes and journal folders are shown greyed out. Meeting invites and other non-mail items inside mail folders are skipped and counted in the log.
- **No internet headers.** Outlook doesn't let other programs write them, so the raw `Received:`/routing headers aren't copied. Subject, sender, recipients, dates, read state, bodies and attachments are.
- **Speed.** About 0.15 seconds per message through Outlook, so a 28,000-message mailbox takes roughly 70 minutes.
- **Outlook's reminders window** may appear when Outlook starts. Leave Outlook running until the conversion finishes.
- **IMAP OSTs** are supported but not yet tested on a real file. Exchange/Microsoft 365 OSTs are tested.

## Development

To run from source, rebuild the EXE, or find your way around the code, see [DEVELOPMENT.md](DEVELOPMENT.md).
