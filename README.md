# Unbound

Windows desktop app that converts an Outlook OST file (for example one copied from another PC) into a PST: mail, calendar and contacts. It also converts **Apple Mail** mailboxes exported on a Mac (MBOX), for moving from a Mac to Windows, and can re-save an existing **PST** as a new one. It reads the OST directly with [libpff](https://github.com/libyal/libpff), so it needs neither the original PC nor its Outlook profile, and writes the PST through Outlook Classic on this PC. See [PLAN.md](PLAN.md) for the design and roadmap.

## Download

Download **`Unbound.exe`** from the [latest release](https://github.com/Z3R0ZIGMA/Unbound/releases/latest). It's a single file that needs nothing else installed apart from Outlook Classic (below).

**Windows may warn the first time you run it.** Windows shows "Windows protected your PC" for programs it hasn't seen often yet. Choose **More info → Run anyway**. To be sure the file is the one that was published, check it first, in PowerShell:

```powershell
Get-FileHash Unbound.exe
```

The result should match the line in `SHA256SUMS.txt` on the release page (releases from 1.3.0 on have one).

To build it yourself instead, see [DEVELOPMENT.md](DEVELOPMENT.md#build-the-exe).

## Requirements

- Windows
- **Outlook Classic** installed on the PC doing the conversion. It writes the PST. The New Outlook can't, because it has no automation (COM) support.

Tested with OSTs from **Exchange / Microsoft 365** and **IMAP** accounts.

## Using it

1. Start `Unbound.exe`.
2. Click **Browse…** and choose **Outlook OST file** to pick an OST, or **Apple Mail export (folder)** for mail exported from a Mac (see below). You can also drag the file or folder onto the window. The folder tree appears with item counts.
3. Tick the folders to convert. Your mail, calendar and contacts folders start ticked; hidden system folders and **Shared mailboxes** start unticked. Ticking a folder ticks its subfolders too. **Select all mail** puts the ticks back to that starting selection, and **Clear** unticks everything.
4. Click **Convert to PST** and choose where to save the PST, such as in Documents. Unbound won't save into Outlook's own data folder, where a file could belong to your Outlook profile. If you choose a PST that already exists, the old one goes to the Recycle Bin. You're warned first if the PST probably won't fit on the drive, or would pass the 50 GB at which Outlook stops adding to a PST.
5. Wait for the progress bar. You can cancel at any time; what's already copied stays in the PST. Closing the window does the same: the conversion stops, and the window closes once Outlook has let go of the PST.

When it finishes, a log is saved next to the PST as `<name> - Unbound log.txt`. It lists what was copied, skipped and failed, plus the folder and item types found in the file. The finished message has buttons to show the PST in its folder and to open the log.

**If a conversion is interrupted** by a crash or a power cut, start Unbound again: it deletes that conversion's temporary files and detaches its PST from Outlook. Outlook can keep hold of that PST file for several minutes afterwards, so to convert again straight away, save under another name.

**Moving from a Mac (Apple Mail).** On the Mac, open Mail, select the mailboxes to move (Cmd-click to pick several), then choose **Mailbox → Export Mailbox…** and pick a folder. Mail writes one `Name.mbox` folder per mailbox, with sub-mailboxes inside. Copy that whole folder to the Windows PC (USB stick, cloud drive or network share), then in Unbound choose **Browse… → Apple Mail export (folder)** and select it. It's a folder rather than a file because Apple Mail exports each mailbox as its own folder. Loose `.mbox` files from other mail programs work too: put them in a folder and select that.

**Converting a PST to a new PST.** Unbound also reads PST files, which is useful to:

- **Upgrade an old-format PST.** PSTs from Outlook 97–2002 (and ones made by Unbound 1.0.0) use the old ANSI format, which can't hold characters outside the Western alphabet (such as emoji or non-Latin text) and is limited to 2 GB. Converting one gives a modern Unicode PST.
- **Copy only some folders** of a PST into a new, smaller PST.

Click **Browse… → Outlook OST file**, change the file type at the bottom right of the file window from "Outlook OST files" to **All files**, and pick the PST. Everything else works as for an OST. (Characters already lost in an old-format PST can't be recovered.)

**Converting this PC's own OST.** An OST in Outlook's data folder (`%LOCALAPPDATA%\Microsoft\Outlook`) may belong to the Outlook profile on this PC, and Outlook locks it while running. Unbound first makes a temporary working copy of such a file and converts from that, so it needs free disk space equal to the OST's size; the copy shows its progress, can be cancelled, and is deleted afterwards. If you see **"Outlook is using this file"**, close Outlook and try again; if Outlook is already closed, Windows Search may be indexing the file, so wait a minute or two and try again. OSTs copied from another PC aren't affected.

Outlook opens in the background during conversion (if it isn't already open) and stays open afterwards. The new PST is not left attached to your Outlook profile; to look at it, use **File → Open & Export → Open Outlook Data File** in Outlook.

## What's converted

- **Mail:** subject, sender, recipients, dates, read state, bodies and attachments. In Apple Mail exports, an email attached to another email is kept as an `.eml` attachment.
- **Calendar:** appointments and meetings with their times, time zone (including zones named the way phones and Google Calendar name them, such as `America/New_York`), all-day flag, attendees, reminders, location, body and attachments. Repeating appointments keep their pattern, end date, cancelled occurrences and changed occurrences (new time, subject and location).
- **Contacts:** every field Outlook stores on the contact (names, emails, phones, addresses, company, notes and so on), plus the contact photo.

## Limitations

- **Not converted yet:** tasks, notes, journal entries and contact groups (shown greyed out or counted as skipped in the log). Meeting requests and delivery reports inside mail folders are skipped too.
- **Emails attached to other emails** (forwarded as an attachment) in an OST or PST aren't copied yet; the log counts them.
- **Damaged files.** If part of an OST or PST can't be read, the items affected are listed as failed in the log and the rest is converted.
- **No internet headers.** Outlook doesn't let other programs write them, so the raw `Received:`/routing headers aren't copied.
- **Changed occurrences** of a repeating appointment keep their new time, subject and location, but not a changed body or attendee list.
- **Apple Mail exports** don't record which messages were read or flagged, so every message arrives as read and unflagged. Mailbox names are kept as they are (for example "Sent Messages" stays "Sent Messages").
- **Speed.** About 0.15 seconds per email and 0.5 seconds per calendar item or contact through Outlook, so a 28,000-message mailbox takes roughly 70 minutes.
- **Outlook's reminders window** may appear when Outlook starts. Leave Outlook running until the conversion finishes.

## Privacy

Unbound works entirely on your PC and makes no internet connections of its own. It never sends mail: every item is only saved into the PST. Temporary files (a working copy of an OST, attachments on their way into Outlook) are deleted when a conversion ends.

## Development

To run from source, rebuild the EXE, or find your way around the code, see [DEVELOPMENT.md](DEVELOPMENT.md).
