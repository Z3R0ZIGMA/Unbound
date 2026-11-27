# Developing Unbound

How to run Unbound from source, test it, build the EXE, publish a release, and find your way around the code. For using the app, see [README.md](README.md); for the design and roadmap, see [PLAN.md](PLAN.md).

## Run from source

Needs Python 3.10 or newer (developed on 3.15) and Outlook Classic. In PowerShell or Command Prompt, from this folder:

```powershell
py -m venv .venv
.venv\Scripts\python -m pip install --require-hashes -r requirements.txt -r requirements-dev.txt
.venv\Scripts\python unbound_app.py
```

`requirements.txt` is what the app needs; `requirements-dev.txt` adds PyInstaller (to build the EXE) and pyflakes. Both pin exact versions with checksums, so pip refuses a package that differs from the one reviewed. libpff-python ships prebuilt Windows wheels, so no C compiler is needed.

## Testing

### Unit tests

Fast, and they need neither Outlook nor a mailbox:

```powershell
.venv\Scripts\python -m unittest discover -s tests -t .
```

They cover the folder scan, the MBOX reader, the recurrence and property decoders (including with damaged data), the reader process (including a reader that crashes or hangs), crash recovery, the conversion loop with a stand-in for Outlook, the PST writer against stand-ins that behave as Outlook was seen to (`tests/fake_outlook.py`: stores it can't reach, folder names it treats as equal, yearly intervals in months, a 29 February series, refused rich text, deleting that only moves to Deleted Items, an Outlook that stops answering while the PST is detached), and the window's rules (what may be saved where, closing during a conversion, remembered folders, a PC without Outlook). The window tests keep their settings in a temporary file, never in the registry. They use only Python's own `unittest`, so testing adds no dependencies.

A static check catches undefined names and leftover imports:

```powershell
.venv\Scripts\python -m pyflakes unbound unbound_app.py scripts assets tests
```

### End to end, through Outlook

The scripts in `scripts/` write a PST through Outlook and read it back with pypff. They need Outlook Classic. They print pass/fail per field and never print message content.

```powershell
# Apple Mail: builds a test export, converts it and checks every message
.venv\Scripts\python scripts\mbox_test.py C:\Temp\unbound_test

# Calendar and contacts: builds a test PST, converts it, compares every item
.venv\Scripts\python scripts\make_pst_fixture.py C:\Temp\unbound_test\fixture.pst
.venv\Scripts\python unbound_app.py --convert C:\Temp\unbound_test\fixture.pst C:\Temp\unbound_test\out.pst
.venv\Scripts\python scripts\compare_pst.py C:\Temp\unbound_test\fixture.pst C:\Temp\unbound_test\out.pst
```

Those two need no real mailbox. The OST scripts do: Outlook locks the OST of the profile on this PC while it runs, so close Outlook and copy the OST from `%LOCALAPPDATA%\Microsoft\Outlook` somewhere else first, then point the scripts at the copy. Delete the copy and the test PSTs afterwards; they contain your mail.

```powershell
# OST: folder tree and a read of every item
.venv\Scripts\python scripts\list_folders.py C:\Temp\copy.ost --sample 100000

# OST: one email with attachments, every field compared
.venv\Scripts\python scripts\copy_one_message.py C:\Temp\copy.ost C:\Temp\one.pst

# OST: repeating and all-day appointments, a meeting and a contact; every property and the recurrence compared
.venv\Scripts\python scripts\copy_items_test.py C:\Temp\copy.ost C:\Temp\items.pst
```

Known results: `copy_one_message.py` fails "internet headers" (Outlook doesn't allow writing them). `copy_items_test.py` used to fail "time zone" for appointments with an IANA zone name such as `America/New_York`; since 1.3.4 those are given the matching Windows zone. (1.3.0 to 1.3.3 tried to, but missed that Outlook answers an unknown zone name with nothing rather than an error, so the translation never ran.) Everything else should pass.

`compare_pst.py` reports how many items in the PST are identical to the source (emails: subject, sender, times, read state, body type, attachments, recipients; appointments: times, full recurrence and notes; contacts: name, email and notes), and which field differs when one isn't. Times are compared to the second, and notes as text, since Outlook stores appointment notes as HTML that Unbound can only write back as text; invisible characters such as soft hyphens are ignored, because Outlook drops some of them. It reads every folder of both files, so against a real OST "source items not in the PST" also counts folders that weren't ticked.

### Testing the EXE itself

Code that works from source can still fail in the EXE if PyInstaller left something out, so test the built EXE before a release. `Unbound.exe` has two hidden modes for that. The first loads every library, starts a reader process and passes a date through COM:

```powershell
Start-Process dist\Unbound.exe -ArgumentList '--check', 'check.txt' -Wait
Get-Content check.txt
```

Every line should start with `OK`. One line checks that secure (HTTPS) connections work, for Check for updates, without connecting to anything. The second converts without the window (the folders ticked by default), writes the usual log next to the PST, and exits with 0 if nothing failed:

```powershell
# Apple Mail: convert the test export with the EXE, then check every message
.venv\Scripts\python scripts\make_mbox_fixture.py C:\Temp\exe_test\export
Start-Process dist\Unbound.exe -ArgumentList '--convert', 'C:\Temp\exe_test\export', 'C:\Temp\exe_test\mbox_test.pst' -Wait
.venv\Scripts\python scripts\mbox_test.py C:\Temp\exe_test --verify-only

# Calendar and contacts: convert the test PST with the EXE, then compare
Start-Process dist\Unbound.exe -ArgumentList '--convert', 'C:\Temp\unbound_test\fixture.pst', 'C:\Temp\exe_test\out.pst' -Wait
.venv\Scripts\python scripts\compare_pst.py C:\Temp\unbound_test\fixture.pst C:\Temp\exe_test\out.pst
```

## Build the EXE

```powershell
.venv\Scripts\pyinstaller --noconfirm --clean Unbound.spec
```

The result is `dist\Unbound.exe` (about 43 MB; the release build on GitHub is about 47 MB, because it bundles some Windows runtime files a local build leaves out). [Unbound.spec](Unbound.spec) holds everything about the build: the icon, the version shown in the file's properties (read from `unbound/__init__.py`), and the modules PyInstaller can't find by itself.

**Hidden imports matter.** PyInstaller only bundles modules it can see being imported. pywin32 loads `win32timezone` on demand, the first time a date marked with a time zone goes through COM, so the spec lists it. Without it, 1.1.0 and 1.2.0 shipped an EXE where every calendar item failed and emails got the wrong received date, although the code worked fine from source. That's why the self-check passes a real date through COM.

`dist/` is git-ignored so builds don't bloat the repo's history.

## Continuous integration

[.github/workflows/ci.yml](.github/workflows/ci.yml) runs on every push to `main` and every pull request:

- **Test and build** (Windows): installs the pinned dependencies with checksums enforced, runs pyflakes and the unit tests, builds the EXE and runs its `--check`. The EXE is kept for a week as a download on the run's page, for trying a change out.
- **Dependency audit:** looks every pinned package up in the public vulnerability databases with `pip-audit`.

Both jobs also run every Monday, so a vulnerability announced after the last push still shows up, and so does a build broken by a change on GitHub's side.

Outlook isn't available there, so the end-to-end scripts above stay a manual step before a release.

## Publishing a release

1. Set the new version in [unbound/__init__.py](unbound/__init__.py). It's the only place it's written. Add the release to "Releases" in [PLAN.md](PLAN.md).
2. Run the unit tests, pyflakes and the end-to-end scripts above. Rebuild `dist\Unbound.exe` from the code you're about to tag and run its `--check`.
3. Commit, push, then tag that commit and push the tag (the tag must match the version, with a `v` in front):

   ```powershell
   git tag -a v1.4.0 -m "Unbound 1.4.0"
   git push origin v1.4.0
   ```

   Never move or re-push a tag that's already on GitHub; a ruleset on the repository refuses it for any `v*` tag. If a release build fails, fix it and release the next patch version.

[.github/workflows/release.yml](.github/workflows/release.yml) then builds the EXE on GitHub from the tagged code (so it always matches what's in the repo), signs it if signing is set up, and publishes a release with `Unbound.exe`, its checksum in `SHA256SUMS.txt`, and the licence notices.

Releases are immutable (a repository setting since November 2026, so from 1.3.6 on): once published, a release's files and its tag can't be changed or deleted, and a deleted release's tag name can't be used again. Its title and notes can still be edited on its page. So a mistake in a published release is fixed by releasing the next patch version. Nothing is published if a test or the EXE's self-check fails, or if the tag doesn't match the version in the code. Watch the run under the repository's **Actions** tab.

While the repository is public, the workflow also records a build attestation: GitHub's signed statement that this exact file was built by this workflow from this commit. Anyone can check a download with `gh attestation verify Unbound.exe --repo Z3R0ZIGMA/Unbound`. GitHub doesn't offer that for private repositories on standard plans, so the step is skipped there.

## Signing and Windows SmartScreen

SmartScreen warns about a download when neither the file nor its publisher has a track record. An unsigned EXE has no publisher, so every release starts from nothing and shows "Windows protected your PC" with "Unknown publisher". Signing changes two things: the warning names the verified publisher, and trust builds up on the publisher, so it carries over to later releases. It isn't instant. Microsoft's own guidance is that a new publisher's downloads can still be warned about for some weeks, until enough people have run them, and that the pricier "EV" certificates no longer skip that wait. The one route with no warning from the first day is the Microsoft Store.

| Route | Cost | Who can use it | What the release needs |
|---|---|---|---|
| **Azure Artifact Signing** (was Trusted Signing) | about $10 a month | Individuals in the USA and Canada; organisations there, in the EU, the UK and some other countries | Nothing more: the release workflow's signing step is written for it |
| **Certificate from a certificate authority** (DigiCert, Sectigo, Certum and others) | roughly $150–300 a year | Anyone, anywhere | A different signing step: the key sits on a hardware token or in the authority's cloud service |
| **SignPath Foundation** | free | Open-source projects that meet its conditions (public code under an open-source licence), approved at its discretion. Unbound qualifies on paper, but this route was set aside: every release needs SignPath's manual approval | Its own signing step; the publisher shown is SignPath Foundation, not you |
| **Microsoft Store** | free | Anyone, anywhere | Packaging the app as MSIX and submitting it; Microsoft signs it |

Check each provider's current terms before choosing; the first row's limits are from Microsoft's documentation at the time of writing.

**Setting up Azure Artifact Signing.** Follow Microsoft's [quickstart](https://learn.microsoft.com/azure/artifact-signing/quickstart) to create a signing account, pass identity validation and create a *Public Trust* certificate profile. Then create an app registration in Microsoft Entra, give it the *Artifact Signing Certificate Profile Signer* role on the certificate profile, and create a client secret for it. In the GitHub repository, under **Settings → Secrets and variables → Actions**, add:

| Kind | Name | Value |
|---|---|---|
| Variable | `SIGNING_ENDPOINT` | The account's URI, such as `https://eus.codesigning.azure.net/` |
| Variable | `SIGNING_ACCOUNT` | The signing account's name |
| Variable | `SIGNING_CERTIFICATE_PROFILE` | The certificate profile's name |
| Secret | `AZURE_TENANT_ID` | The directory (tenant) ID |
| Secret | `AZURE_CLIENT_ID` | The app registration's application (client) ID |
| Secret | `AZURE_CLIENT_SECRET` | The client secret |

From the next tag on, the release workflow signs the EXE, checks the signature and runs the EXE's self-check again. Until `SIGNING_ACCOUNT` exists it builds unsigned, and says so in the run's summary.

## Updating dependencies

[Dependabot](.github/dependabot.yml) opens a pull request each week for any dependency or GitHub Action with a new version, with the new checksums filled in; CI then builds and tests it. PySide6 and shiboken6 are updated together because they must match. A new version is only proposed once it has been out for 7 days (a "cooldown"), so a hijacked or broken release is usually pulled before it gets here; security updates for a known vulnerability come straight away. A merged pull request's branch is deleted automatically.

To change a version by hand, the checksums have to be regenerated too. In a separate, throwaway environment:

```powershell
py -m venv C:\Temp\lock
C:\Temp\lock\Scripts\python -m pip install pip-tools
# in.txt: the packages marked "used directly" in the file being updated, one "name==version" per line
C:\Temp\lock\Scripts\python -m piptools compile --generate-hashes --allow-unsafe --strip-extras --no-header -o C:\Temp\lock\out.txt C:\Temp\lock\in.txt
```

Then replace everything below the comment at the top of `requirements.txt` (or `requirements-dev.txt`) with the contents of `out.txt`. Rebuilding with a newer Python updates the Python runtime inside the EXE, which is a dependency too. Update the versions in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) to match.

## How it's put together

- **Reading happens in a second process.** OST and PST files are decoded by libpff, a C library, and may come from anywhere. If a damaged file made it crash inside the window's process, the window would vanish mid-conversion. So `unbound/pff_process.py` opens the file in a child process: if that dies or hangs on an item, the item is logged as failed, a new reader is started, and the conversion carries on. MBOX files are read by Python itself, in the main process.
- **Each conversion has its own temporary folder** (`%TEMP%\unbound_…`) for the working copy of an OST and for attachments on their way into Outlook. A marker file in it stays open while the conversion runs. If Unbound dies, the next start finds the folder, sees that nothing holds the marker any more, deletes the folder and detaches that conversion's PST from Outlook. A conversion still running in another window keeps its marker locked and is left alone.
- **Closing the window never abandons a conversion.** It asks the conversion to stop and closes once the PST is detached and the temporary folder is gone.
- **Outlook's objects share what's known about their type.** pywin32 asks a COM server to describe every new object from scratch, which was about half the time an email took. `_Dispatch` in `unbound/pst_writer.py` shares a type's description between its objects for one conversion. It leans on pywin32's insides (`win32com.client.dynamic`), so a unit test and the EXE's self-check ("shared COM types") both exercise it; if a pywin32 update breaks it, they fail, and `_outlook()` falls back to pywin32's ordinary objects.
- **Properties are copied one at a time, on purpose.** `PropertyAccessor.SetProperties` is far quicker, but with an item's copied properties it crashed Outlook on a contact and behaved differently on appointments. Only an email's sender and date fields, a fixed set of text and dates, go in one call.
- **Detaching the PST has a time limit.** Outlook was seen to stop answering at that point during a Microsoft 365 outage. Detaching is asked of Outlook from a thread of its own (`DETACH_TIMEOUT` in `unbound/pst_writer.py`, 3 minutes); if it doesn't answer, the conversion finishes anyway and reports the PST as possibly still attached, with the counts.
- **An update is swapped in when Unbound closes, not when it's downloaded.** Windows lets a running program be renamed but not overwritten, so Check for updates downloads the new EXE as `Unbound.exe.new` next to the running one and checks it against the release's `SHA256SUMS.txt`. When Unbound closes (or restarts, if the user chooses Restart now), the running `Unbound.exe` is renamed to `Unbound.exe.old` and the new one takes its name; the next start deletes the `.old` file. Swapping only then matters because the reader process is `Unbound.exe` started again: a session swapped mid-way would start its readers from the new version. A restart is started with `PYINSTALLER_RESET_ENVIRONMENT=1`, so the new version unpacks its own files instead of reusing the old one's. Only files under this repository's own release downloads are fetched, and only after the user agrees.

## Project layout

| Path | What it is |
|---|---|
| `dist/Unbound.exe` | The built app (not in git; created by the build above) |
| `Unbound.spec` | How the EXE is built |
| `unbound_app.py` | Entry point (also the `--check` and `--convert` test modes) |
| `unbound/__init__.py` | The version number |
| `unbound/app.py` | The window (PySide6, dark theme; colours are in `COLORS` at the top) |
| `unbound/converter.py` | Folder scan and the conversion loop (OST, PST or MBOX → PST), temporary files and crash recovery, log writing |
| `unbound/pff_process.py` | Runs the OST/PST reader in a child process and restarts it if it dies |
| `unbound/ost_reader.py` | Reads messages out of an OST or PST with pypff |
| `unbound/mbox_reader.py` | Finds the mailboxes in an Apple Mail export and reads MBOX messages |
| `unbound/pst_writer.py` | Writes mail, appointments and contacts into the PST through Outlook |
| `unbound/props.py` | Reads every property of an item, resolving named properties via the file's name-to-ID map |
| `unbound/appointment.py` | Decodes recurrence patterns and time zones, which Outlook only accepts through its object model |
| `unbound/timezones.py` | Windows names for IANA time zones (`America/New_York`), asked of Windows itself |
| `unbound/updates.py` | Check for updates: asks GitHub for the latest release, downloads `Unbound.exe` and checks it against `SHA256SUMS.txt`, and swaps it in for the running EXE |
| `unbound/mapi.py` | MAPI property tags shared by reader and writer |
| `unbound/errors.py` | Error text as shown in the window and the log |
| `tests/` | Unit tests (no Outlook needed); `tests/fakes.py` and `tests/fake_outlook.py` are the stand-ins for the reader and for Outlook |
| `scripts/list_folders.py` | Command-line check: prints an OST's folder tree and test-reads items |
| `scripts/copy_one_message.py` | Command-line check: copies one message to a PST and compares every field |
| `scripts/copy_items_test.py` | Command-line check: copies sample appointments and a contact to a PST and compares every property and the recurrence |
| `scripts/make_mbox_fixture.py` | Builds a test Apple Mail export covering the tricky cases |
| `scripts/mbox_test.py` | End-to-end MBOX check: builds the test export, converts it and verifies every message (`--verify-only` checks a PST made by the EXE) |
| `scripts/make_pst_fixture.py` | Builds a test PST with calendar items and contacts through Outlook |
| `scripts/compare_pst.py` | Compares a converted PST with its source, item by item |
| `assets/make_icon.py` | Draws the app icon (also the window's logo) and the repository's social preview; re-run it after changing the colours or shape, then rebuild. Upload `assets/social-preview.png` under the repository's Settings → Social preview |
| `requirements.txt`, `requirements-dev.txt` | Pinned dependencies with checksums: the app's, and the build tools' |
| `.github/` | CI, the release workflow and Dependabot's settings |
| `.gitattributes` | Line endings: text is stored with LF, PowerShell scripts always use CRLF, images are binary |
| `SECURITY.md` | How to report a security problem privately, and how releases are built and can be verified |
| `LICENSE` | Unbound's own licence (MIT); the copyright line is also in the EXE's file details (`Unbound.spec`) |
| `THIRD-PARTY-NOTICES.md`, `third-party-licenses/` | The software inside the EXE and the licence texts it requires (LGPL and GPL), shipped with each release |
