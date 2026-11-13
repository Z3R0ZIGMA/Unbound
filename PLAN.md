# Unbound – Project Plan

## Goal

A Windows desktop app that converts mail data into an Outlook PST file, using Outlook Classic on the local machine to write the PST:

- an Outlook **OST** (for example copied from another machine), without needing the original PC or profile
- an **Apple Mail** export (MBOX), for moving from a Mac to Windows
- an existing **PST**, to upgrade an old ANSI-format PST to Unicode or copy only some of its folders

## Features (current: 1.3.3)

- Open an OST, a PST, or an Apple Mail export folder (Browse, or drop it on the window)
- Preview the folder tree with item counts; mail, calendar and contacts folders are ticked by default, system folders and shared mailboxes are not
- Convert mail, calendar (including repeating meetings) and contacts, with a progress bar, time remaining and cancel
- Save a log next to the PST: what was copied, skipped, copied with warnings or failed, plus the folder and item types found
- Converts this PC's own (locked) OST from a temporary working copy, with progress and cancel, also while Word, Excel, Teams or Windows Search has the file open
- Looks after the user's files: an existing PST is replaced by way of the Recycle Bin, Outlook's own data folder is refused as a destination, and a PST that probably won't fit (or would pass Outlook's 50 GB limit) is warned about first
- Cleans up after itself: closing the window stops the conversion properly, and a conversion cut short by a crash is tidied away at the next start
- Carries on past damaged parts of an OST or PST, listing the affected items in the log

## Architecture

### 1. Readers

- **OST/PST:** libpff through its Python bindings (pypff), in `unbound/ost_reader.py`. Named properties (most calendar and contact fields) are resolved by decoding the file's name-to-ID map in `unbound/props.py`, because pypff only exposes it as raw data. Recurrence patterns are decoded in `unbound/appointment.py`.
- **MBOX:** Python's `email` package, in `unbound/mbox_reader.py`. Finds the `Name.mbox` folders of an Apple Mail export and splits each mbox file into messages.
- Both produce the same `Message` record, so the rest of the app doesn't care where an item came from.
- OST/PST reading runs in a child process (`unbound/pff_process.py`): libpff is a C library decoding files from anywhere, and a crash in it would otherwise close the window mid-conversion. If the reader dies or hangs, that item is logged as failed and a new reader takes over.
- Neither needs Outlook or the original profile.

### 2. Writer (PST creation)

- Outlook Classic via COM automation (pywin32), in `unbound/pst_writer.py`
- Creates a new Unicode PST, recreates the folder structure (with calendar and contacts folder types), adds items, then detaches the PST from the profile
- Note: the New Outlook has no COM automation, so it can't be used

### 3. Interface

- PySide6 desktop window, `unbound/app.py`; conversion engine in `unbound/converter.py`
- Packaged as a single Windows EXE with PyInstaller (`Unbound.spec`), built and published to GitHub Releases by a GitHub Actions workflow

## Interface layout

- Dark mode throughout (dark window, light text); colours in one place (`COLORS`)
- Header: grey geometric triangle logo, "Unbound", one-line description; window title "Unbound – Convert to PST"
- Top: path box and a **Browse…** menu (Ctrl+O) with "Outlook data file (OST or PST)…" and "Apple Mail export (folder)…"; both remember the last folder used
- Middle: folder tree with item counts and checkboxes; unsupported folders greyed out
- Bottom: large Convert button, progress bar with Cancel, status line

## Roadmap

Done:

- ~~Contacts and calendar items~~ 1.1
- ~~Apple Mail (Mac) MBOX to PST conversion~~ 1.2
- ~~Map time zone names in IANA format (e.g. `America/New_York`, written by phones and Google) to Windows zones~~ 1.3
- ~~Emails attached to other emails in Apple Mail exports~~ 1.3 (they were being dropped)

Next candidates:

- Tasks, notes and journal entries; contact groups (distribution lists)
- Emails attached to other emails (embedded messages) in OST/PST sources
- Changed occurrences of a repeating meeting: copy a changed body and attendee list too
- Apple Mail: map mailbox names to Outlook's ("Sent Messages" to "Sent Items"); read Mail's own storage (`.emlx` in `~/Library/Mail`) so read/flagged status survives
- Date range filters
- Splitting very large PSTs
- Resuming an interrupted conversion (skipping what's already in the PST)
- Signing the EXE so Windows SmartScreen names the publisher and trust carries over between releases: the release workflow is ready for it, and what's left is choosing a signing route and adding its credentials (see "Signing and Windows SmartScreen" in DEVELOPMENT.md)
- Publishing through the Microsoft Store (MSIX), the only route with no SmartScreen warning from the first day
- ~~A licence for Unbound's own code~~ MIT License, chosen 12 November 2026 ([LICENSE](LICENSE))

## Releases

- **1.0.0:** mail from OST to PST (Exchange/Microsoft 365 and IMAP OSTs)
- **1.1.0:** calendar and contacts; PSTs created as Unicode (1.0.0 made ANSI PSTs)
- **1.2.0:** Apple Mail (MBOX); inline images in OST emails kept again; plain-text and rich-text mail keep their format; repeating meetings in other time zones fixed; new title, logo and Browse menu
- **1.3.3 (not released yet):** fixes and small improvements from a review of the code; see "Version 1.3.3" below. Also the MIT License
- **1.3.2:** a mailbox that other programs have open can be converted. Word and Excel hold the Outlook profile's OST open while they run (so do Teams and Windows Search) and lock part of it, and Unbound reported "Outlook is using this file" although Outlook was closed. It now asks Windows which programs have the file open (Restart Manager) and names them. If Outlook isn't among them it reads past the lock: Windows applies such locks to ordinary reads but not to a memory-mapped view, so the folder scan reads through one and the working copy is taken through one. If Outlook has the file, it's still refused, because Outlook changes the file as it syncs; the message then offers to close Outlook, for when it's running without a window. Tested with a real lock held by a second program (scan, working copy, a full conversion compared item by item) and by scanning a real OST while Word and Excel had it open; the "Close Outlook" button is tested with a stand-in, not against a live Outlook
- **1.3.1:** the first published build of 1.3. The 1.3.0 tag has no release: all its tests passed, but the test process crashed as it exited (windows left over from the tests), which stopped the build. Only the tests changed
- **1.3.0:** safety and reliability; see "Version 1.3" below (tagged, not published)
- **1.2.1:** fixes the EXE (1.1.0 and 1.2.0): calendar items failed with "No module named 'win32timezone'" and emails got the conversion date as their received date, because PyInstaller left out a module pywin32 loads on demand. The self-check now passes a date through COM, the log lists refused properties, and `Unbound.exe --convert` lets the EXE itself be tested end to end

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
   - ~~`PropertyAccessor` treats dates as local time; convert pypff's UTC times to local first~~ Corrected in 1.1: it's pywin32 that treats naive datetimes as local time; dates are now passed marked as UTC (see below)
   - Call `Recipient.Resolve()` or the address is lost (the whole text becomes the display name)
   - Known limitation: Outlook won't let scripts write internet headers (`PR_TRANSPORT_MESSAGE_HEADERS`) or the To/Cc display lines directly (the latter are filled in from the recipients)
   - Starting Outlook from the script can show its reminders window, which keeps Outlook running and the PST locked; release all COM references and wait for the file to unlock

## Real-world testing (1.0)

- Version 1.0.0 converts OSTs from both **Exchange / Microsoft 365** and **IMAP** accounts. For IMAP, the reader treats `IPF.Imap` folders as mail and handles the PST-style "Top of Outlook data file" layout as well as Exchange's `Root - Mailbox/IPM_SUBTREE`.
- The OST of the profile on the converting PC is locked by Outlook once it starts, so Unbound converts such files from a temporary copy taken before Outlook starts.

## Version 1.1: calendar and contacts

- Appointments and contacts are copied property by property: every MAPI property of the source item is read (named properties resolved by decoding the file's name-to-ID map, since pypff only exposes it as raw data) and written with `PropertyAccessor`.
- Outlook refuses raw writes of an appointment's recurrence and time-zone properties and recomputes start/end on save, so those go through the object model: `StartTimeZone`, `StartUTC`/`EndUTC`, `GetRecurrencePattern()`, and `GetOccurrence()` for cancelled and changed occurrences. The recurrence blob is decoded per [MS-OXOCAL] 2.2.1.44; checked against real files, `ARO_APPTCOLOR` carries no data in `ExceptionInfo`.
- pywin32 converts naive datetimes from local time to UTC on their way into COM but passes UTC-aware ones unchanged, so all dates are handed over marked UTC.
- Tested on a real OST: 616/616 appointments (including 64 repeating series with cancelled and changed occurrences) and all contacts identical after conversion.
- Fixed in passing: the PST was created with `olStoreANSI` (3) instead of `olStoreUnicode` (2), which lost characters outside the Western code page and limited the PST to 2 GB; hidden attachments (a series' changed occurrences) and attached emails were copied as broken files.

## Version 1.2: Apple Mail (MBOX)

- Opens Apple Mail's "Export Mailbox" output: a folder of `Name.mbox` folders, each holding an `mbox` file, nested for sub-mailboxes. Loose `.mbox` files in the chosen folder work too.
- Messages are split on `From ` lines (with mboxrd `>From` quoting undone), parsed with Python's `email` package into the same `Message` record as the OST reader, and written by the same PST writer. Attached emails (`message/rfc822`) become `.eml` attachments.
- Inline images keep their content ID so the HTML shows them. Outlook ignores attempts to set `PR_ATTACHMENT_HIDDEN` (reset on save) and refuses `PR_ATTACH_FLAGS`, so it decides itself whether to list them.
- Plain-text mail is now saved as plain text (`BodyFormat = olFormatPlain`); before, Outlook turned it into HTML and the PST had no plain body.
- Fixed a regression from 1.1: hidden attachments were all skipped, which dropped the inline images of OST emails; now only hidden embedded items (a repeating appointment's changed occurrences) are skipped.
- Tested end to end with a generated export (`scripts/mbox_test.py`): emoji and encoded subjects, HTML with inline image, attachments, Cc, read/unread, nested mailboxes; all checks pass. Not yet tested with a real export from a Mac.
- Review fixes before release: a repeating series in another time zone now takes its zone from the recurrence (`PidLidAppointmentTimeZoneDefinitionRecur`) rather than the start "display" zone, which Outlook resets to the local zone when the pattern is set (it shifted such a series by the zone difference when a converted PST was converted again); rich-text (RTF-only) mail keeps its formatting; an Apple mailbox that only holds other mailboxes is no longer reported as failed; attachments without a long file name use their short name.
- Full test run: live OST through the window (working copy, 184 emails, 587 calendar items, 0 failed), OST/PST scripts, PST round trip (184/184 emails, 587/587 appointments identical, 281 inline images kept), MBOX end-to-end and edge cases, EXE self-check.

## Version 1.3: safety and reliability

From a review of the code for security, bugs and dependencies.

- **Nothing left behind.** Closing the window during a conversion used to give up after two minutes and exit anyway, which left the PST attached to the Outlook profile and, for this PC's own OST, a full copy of the mailbox in `%TEMP%`. A crash did the same. Now the window waits for the conversion to stop; every conversion has its own temporary folder with a marker file that stays open while it runs; and the next start removes the folder of a conversion that died and detaches its PST. Tested by killing a conversion partway. Outlook keeps the file of such a PST locked for about six minutes after it's detached, until Windows clears up the dead program's connections.
- **Reader in its own process**, as described under Architecture. Tested with a stand-in reader that crashes and hangs, and with damaged files.
- **Recipients and senders.** A quote in a display name broke the `"Name" <address>` text handed to Outlook. A recipient with only a name was looked up in *this* PC's address book and could become a different person. An Exchange sender with only a directory address (`/o=...`) was stored as an SMTP address. All three fixed.
- **Bugs found.** A folder without a name was passed to Outlook as an empty name, so every item in it failed. A body line starting with "From " split an MBOX message in two when the file didn't quote such lines; a separator must now carry a date, as real ones do. An email attached to another email in an MBOX was dropped, and its own attachments were shown on the outer message. A long attachment name lost its extension. A failed conversion left no log, and a log that couldn't be saved left the window stuck.
- **Dependencies.** All pinned with checksums, including PyInstaller, which wasn't pinned before. No known vulnerabilities in any of them at the time (PyPI and OSV databases). Dependabot proposes updates; CI audits them weekly.
- **Tests and CI.** Unit tests that need neither Outlook nor a mailbox, run with pyflakes, a build of the EXE and its self-check on every push. A generated PST fixture (`scripts/make_pst_fixture.py`) lets calendar and contacts be tested end to end without a real OST.
- **Releases** are built by GitHub Actions from a version tag, with a checksum file, the licence notices of the bundled libraries, an optional signing step, and (for a public repository) a build attestation.
- Full test run: unit tests; MBOX end to end (7 messages, all checks pass); PST round trips through the reader process (7/7 emails, 5/5 appointments including a series with a cancelled and a moved occurrence, 2/2 contacts identical); the same with the built EXE; the window driven through a conversion, a close during a conversion and a start after a killed conversion. Not yet run against a real OST: IANA time zones, an Exchange sender without an internet address, and recipients without an address.

## Version 1.3.3: review fixes

From a second review of the code for security, bugs and quick improvements to the window.

- **Data that was lost.** Appointment notes stored as HTML (as Outlook stores them) were dropped, because appointments and contacts can't take HTML from another program; their text is now kept. Yearly series ("second Tuesday of March", "every two years") failed, because Outlook counts a yearly interval in months. When Outlook refuses rich text (RTF), which it did for every form tried on the test PC, the text is kept as plain text and the log says so; before, an email without a separate plain body was saved with no body. RTF read from an OST also had padding after its end that Outlook rejects.
- **Conversions that couldn't start or finish.** One store in the Outlook profile that Outlook couldn't reach (an unavailable shared mailbox, say) stopped every conversion, after Outlook had hung for minutes on it; the new PST is now found by index, newest first, stepping over such stores. A reader that died or hung while opening a file froze the window; opening and the folder scan now have time limits. An error with no text aborted the conversion. Two folders whose names differ only in capitals or accent encoding: every item of the second failed; they now share a folder, and a name Outlook still refuses gets " (2)".
- **Smaller bugs.** An item that failed after its first save stayed in the PST half-written; it's now removed. A recipient type with an unusual flag skipped resolving the address. Damaged values (an empty true/false, a date in the year 9999) aborted reading an item. A line break in a subject could add fake lines to the log. A file that merely starts with "From " was taken for a mailbox. The EXE's `--check` and `--convert` could hang on an error instead of returning an exit code. CI and release jobs have time limits.
- **Security.** Unbound's own two Windows DLLs (Restart Manager, ICU) are loaded by full path from System32. The README says to keep the EXE in a folder of its own and how to check a release's build attestation.
- **Window.** The Browse menu offers "Outlook data file (OST or PST)", so a PST no longer needs the "All files" trick; Ctrl+O opens it. "Select all mail" is now **Reset selection**. Both file windows remember the last folder (in `HKEY_CURRENT_USER\Software\Unbound`); the save window starts in the real Documents folder. Plain-English reasons when a file can't be opened, with the technical text under "Show details". If Outlook Classic isn't installed, the window says so at once and won't start a conversion. A moving bar while Outlook starts, time left measured from the first item, and the end message counts emails attached to other emails that weren't copied. Keyboard shortcuts on the buttons and names for screen readers.
- **Found by a real conversion of the owner's OST.** During a Microsoft 365 outage, Outlook stopped answering while the PST was being detached at the end, and the window stayed at 100% for ever. When that Outlook was ended, the run was reported as "Conversion failed" with no counts. Detaching now has a 3-minute limit, and a PST that couldn't be detached is reported with the counts and how to close it in Outlook. A yearly series on 29 February (a birthday) failed, because Outlook works out the day and month from the start and then refuses to be told them; each part of a pattern is now set only if Outlook hasn't already got it. An item removed after failing went to the PST's Deleted Items rather than away; it's now deleted from there too.
- Tests: unit tests for each fix, with stand-ins that behave as Outlook was seen to (`tests/fake_outlook.py`). End to end with generated data: the PST fixture, now with both yearly series, a 29 February birthday and HTML notes (8/8 appointments and 2/2 contacts identical, notes included), and the MBOX test (all checks pass). A real OST (Microsoft 365): 229/229 emails identical, and 619 of 620 appointments; the odd one out is the failed birthday that build left in Deleted Items, which both fixes now prevent. Two appointments' notes differed only in invisible characters (soft hyphens, a text-direction mark), which `compare_pst.py` now ignores. Not done: signed or encrypted (S/MIME) mail, which may fail to convert and needs a real signed email to test.
