# Unbound

Windows desktop app that converts an Outlook OST file (for example one copied from another PC) into a PST: mail, calendar and contacts. It also converts **Apple Mail** mailboxes exported on a Mac (MBOX), for moving from a Mac to Windows. It reads the OST directly with [libpff](https://github.com/libyal/libpff), so it needs neither the original PC nor its Outlook profile, and writes the PST through Outlook Classic on this PC. See [PLAN.md](PLAN.md) for the design and roadmap.

## Download

Download **`Unbound.exe`** from the [latest release](https://github.com/Z3R0ZIGMA/Unbound/releases/latest). It's a single file that needs nothing else installed apart from Outlook Classic (below). Windows SmartScreen may warn the first time because the EXE isn't signed; choose **More info → Run anyway**.

To build it yourself instead, see [DEVELOPMENT.md](DEVELOPMENT.md#build-the-exe).

## Requirements

- Windows
- **Outlook Classic** installed on the PC doing the conversion. It writes the PST. The New Outlook can't, because it has no automation (COM) support.

Tested with OSTs from **Exchange / Microsoft 365** and **IMAP** accounts.

## Using it

1. Start `Unbound.exe`.
2. Click **Browse…** and choose **Outlook OST file** to pick an OST, or **Apple Mail export (folder)** for mail exported from a Mac (see below). The folder tree appears with item counts.
3. Tick the folders to convert. Your mail, calendar and contacts folders start ticked; hidden system folders and **Shared mailboxes** start unticked. Ticking a folder ticks its subfolders too. **Select all mail** puts the ticks back to that starting selection, and **Clear** unticks everything.
4. Click **Convert to PST** and choose where to save the PST. Save it somewhere other than Outlook's own data folder, such as Documents.
5. Wait for the progress bar. You can cancel at any time; what's already copied stays in the PST.

When it finishes, a log is saved next to the PST as `<name> - Unbound log.txt`. It lists what was copied, skipped and failed, plus the folder and item types found in the file.

**Moving from a Mac (Apple Mail).** On the Mac, open Mail, select the mailboxes to move (Cmd-click to pick several), then choose **Mailbox → Export Mailbox…** and pick a folder. Mail writes one `Name.mbox` folder per mailbox, with sub-mailboxes inside. Copy that whole folder to the Windows PC (USB stick, cloud drive or network share), then in Unbound choose **Browse… → Apple Mail export (folder)** and select it. It's a folder rather than a file because Apple Mail exports each mailbox as its own folder. Loose `.mbox` files from other mail programs work too: put them in a folder and select that.

**Converting this PC's own OST.** An OST in Outlook's data folder (`%LOCALAPPDATA%\Microsoft\Outlook`) may belong to the Outlook profile on this PC, and Outlook locks it while running. Unbound first makes a temporary working copy of such a file and converts from that, so it needs free disk space equal to the OST's size; the copy is deleted afterwards. If you see **"Outlook is using this file"**, close Outlook, wait a few seconds, and try again. OSTs copied from another PC aren't affected.

Outlook opens in the background during conversion (if it isn't already open) and stays open afterwards. The new PST is not left attached to your Outlook profile; to look at it, use **File → Open & Export → Open Outlook Data File** in Outlook.

## What's converted

- **Mail:** subject, sender, recipients, dates, read state, bodies and attachments.
- **Calendar:** appointments and meetings with their times, time zone, all-day flag, attendees, reminders, location, body and attachments. Repeating appointments keep their pattern, end date, cancelled occurrences and changed occurrences (new time, subject and location).
- **Contacts:** every field Outlook stores on the contact (names, emails, phones, addresses, company, notes and so on), plus the contact photo.

## Limitations

- **Not converted yet:** tasks, notes, journal entries and contact groups (shown greyed out or counted as skipped in the log). Meeting requests and delivery reports inside mail folders are skipped too.
- **Emails attached to other emails** (forwarded as an attachment) aren't copied yet; the log counts them.
- **No internet headers.** Outlook doesn't let other programs write them, so the raw `Received:`/routing headers aren't copied.
- **Changed occurrences** of a repeating appointment keep their new time, subject and location, but not a changed body or attendee list.
- **Apple Mail exports** don't record which messages were read or flagged, so every message arrives as read and unflagged. Mailbox names are kept as they are (for example "Sent Messages" stays "Sent Messages").
- **Speed.** About 0.15 seconds per email and 0.5 seconds per calendar item or contact through Outlook, so a 28,000-message mailbox takes roughly 70 minutes.
- **Outlook's reminders window** may appear when Outlook starts. Leave Outlook running until the conversion finishes.

## Development

To run from source, rebuild the EXE, or find your way around the code, see [DEVELOPMENT.md](DEVELOPMENT.md).
