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

- ~~Contacts and calendar items~~ Done in version 1.1 (see below)
- Date range filters
- Splitting very large PSTs
- ~~Apple Mail (Mac) MBOX to PST conversion~~ Done in version 1.2 (see below)

## Risks and first steps

1. ~~Build/install libpff (pypff) on Windows — highest risk~~ Done: `libpff-python` 20261030 installs from a prebuilt Windows wheel (no compiler needed)
2. ~~Test reading a real, recent OST (Outlook 2013+ format support is partial)~~ Done with `scripts/list_folders.py` on two current Microsoft 365 OSTs (96 MB and 5.6 GB): every message, calendar item, contact and task read with no failures (28,292 in the large file, about 8 minutes). Findings for the reader:
   - Open files with an absolute, backslash path (`os.path.abspath`); libpff's `\\?\` long-path handling rejects forward slashes
   - pypff raises `OSError` when a message has no HTML or RTF body; treat that as "absent", not a failure
   - Skip non-message items (`pypff.item`, e.g. the hidden `PersonMetadata` and `Help` folders)
3. ~~Milestone: small app that opens a sample OST and lists its folders~~ Done, and v1 is built: `unbound_app.py` (window in `unbound/app.py`, engine in `unbound/converter.py`). Tested on a real OST: Inbox, 84 messages in 12s (~0.15s each, so a 28k-message mailbox is roughly 70 minutes). Notes:
   - The tree shows the mailbox folders (`Root - Mailbox/IPM_SUBTREE`) as the top level, plus cached shared mailboxes (`Shared Data`) as an unticked "Shared mailboxes" group
   - Folders are classed by container class (`IPF.Note` = mail); non-mail folders are greyed out, and system folders start unticked
   - Meeting invites and other non-`IPM.Note` items in mail folders are skipped and counted in the log
4. ~~Then prove writing one message into a PST via Outlook COM~~ Done with `scripts/copy_one_message.py` (code in `unbound/`): a real message with 8 recipients and 4 attachments copied into a new PST and read back with pypff. Subject, sender, delivery/sent times, read state, HTML and plain body, recipients and attachments (byte-identical) all match, and the item is received mail, not a draft. Findings for the writer:
   - A new `MailItem` always saves as an unsent draft. Create a `PostItem`, set `MessageClass = "IPM.Note"`, save, then reopen it with `GetItemFromID`; it's now a `MailItem` that accepts real recipients
   - The first save overwrites delivery/sent times, so set them after reopening
   - `PropertyAccessor` treats dates as local time; convert pypff's UTC times to local first
   - Call `Recipient.Resolve()` or the address is lost (the whole text becomes the display name)
   - Known limitation: Outlook won't let scripts write internet headers (`PR_TRANSPORT_MESSAGE_HEADERS`) or the To/Cc display lines directly (the latter are filled in from the recipients)
   - Starting Outlook from the script can show its reminders window, which keeps Outlook running and the PST locked; release all COM references and wait for the file to unlock

## Real-world testing

- Version 1.0.0 converts OSTs from both **Exchange / Microsoft 365** and **IMAP** accounts. For IMAP, the reader treats `IPF.Imap` folders as mail and handles the PST-style "Top of Outlook data file" layout as well as Exchange's `Root - Mailbox/IPM_SUBTREE`.
- The OST of the profile on the converting PC is locked by Outlook once it starts, so Unbound converts such files from a temporary copy taken before Outlook starts.

## Version 1.1: calendar and contacts

- Appointments and contacts are copied property by property: every MAPI property of the source item is read (named properties resolved by decoding the file's name-to-ID map, since pypff only exposes it as raw data) and written with `PropertyAccessor`.
- Outlook refuses raw writes of an appointment's recurrence and time-zone properties and recomputes start/end on save, so those go through the object model: `StartTimeZone`, `StartUTC`/`EndUTC`, `GetRecurrencePattern()`, and `GetOccurrence()` for cancelled and changed occurrences. The recurrence blob is decoded per [MS-OXOCAL] 2.2.1.44; checked against real files, `ARO_APPTCOLOR` carries no data in `ExceptionInfo`.
- pywin32 converts naive datetimes from local time to UTC on their way into COM but passes UTC-aware ones unchanged, so all dates are handed over marked UTC.
- Tested on a real OST: 616/616 appointments (including 64 repeating series with cancelled and changed occurrences) and all contacts identical after conversion.
- Fixed in passing: the PST was created with `olStoreANSI` (3) instead of `olStoreUnicode` (2), which lost characters outside the Western code page and limited the PST to 2 GB; hidden attachments (a series' changed occurrences) and attached emails were copied as broken files.

## Version 1.2: Apple Mail (MBOX)

- Opens Apple Mail's "Export Mailbox" output: a folder of `Name.mbox` folders, each holding an `mbox` file, nested for sub-mailboxes. A single mbox file works too.
- Messages are split on `From ` lines (with mboxrd `>From` quoting undone), parsed with Python's `email` package into the same `Message` record as the OST reader, and written by the same PST writer. Attached emails (`message/rfc822`) become `.eml` attachments.
- Inline images keep their content ID so the HTML shows them. Outlook ignores attempts to set `PR_ATTACHMENT_HIDDEN` (reset on save) and refuses `PR_ATTACH_FLAGS`, so it decides itself whether to list them.
- Plain-text mail is now saved as plain text (`BodyFormat = olFormatPlain`); before, Outlook turned it into HTML and the PST had no plain body.
- Fixed a regression from 1.1: hidden attachments were all skipped, which dropped the inline images of OST emails; now only hidden embedded items (a repeating appointment's changed occurrences) are skipped.
- Tested end to end with a generated export (`scripts/mbox_test.py`): emoji and encoded subjects, HTML with inline image, attachments, Cc, read/unread, nested mailboxes; all checks pass. Not yet tested with a real export from a Mac.
- Review fixes before release: a repeating series in another time zone now takes its zone from the recurrence (`PidLidAppointmentTimeZoneDefinitionRecur`) rather than the start "display" zone, which Outlook resets to the local zone when the pattern is set (it shifted such a series by the zone difference when a converted PST was converted again); rich-text (RTF-only) mail keeps its formatting; an Apple mailbox that only holds other mailboxes is no longer reported as failed; attachments without a long file name use their short name.
- Full test run: live OST through the window (working copy, 184 emails, 587 calendar items, 0 failed), OST/PST scripts, PST round trip (184/184 emails, 587/587 appointments identical, 281 inline images kept), MBOX end-to-end and edge cases, EXE self-check.
