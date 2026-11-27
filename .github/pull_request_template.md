## What this changes


## Checks
- [ ] `python -m pyflakes unbound unbound_app.py scripts assets tests` and `python -m unittest discover -s tests -t .` pass
- [ ] No real mailbox data added: no names, email addresses, subjects or message text, and no real OST, PST or MBOX files (generated test data only)
- [ ] Still no network use except Check for updates (only when clicked), and Unbound still never sends mail
- [ ] New or updated dependencies are pinned with checksums in `requirements*.txt`, and THIRD-PARTY-NOTICES.md is updated
- [ ] README.md (Privacy section included) and DEVELOPMENT.md updated if behaviour changed
- [ ] If packaging changed: `Unbound.exe --check check.txt` passes on a fresh build
