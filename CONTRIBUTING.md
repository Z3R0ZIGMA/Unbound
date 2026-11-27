# Contributing to Unbound

Thanks for helping. Unbound is a small tool kept by one person, so please open an issue before starting on a big
change.

- **Bugs:** use the bug report form under Issues. Don't attach mailbox files, converted PSTs or full conversion logs:
  they hold email, names and addresses. Copy the counts and error lines from the log instead.
- **Security problems:** report them privately (see [SECURITY.md](SECURITY.md)), never as an issue.
- **Code:** [DEVELOPMENT.md](DEVELOPMENT.md) explains running from source, testing, building the EXE and the layout.
  The ground rules that matter most:
  - no network use except Check for updates (only when clicked), and Unbound never sends mail
  - no personal or real mailbox data in the repository: tests use generated OST, PST and MBOX data
  - a damaged item is logged as failed and the conversion carries on; it never stops the whole conversion
  - dependencies are pinned with checksums, and the bundled software is listed in
    [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)
  - pyflakes and the unit tests must pass

By contributing you agree that your work is released under the [MIT licence](LICENSE).
