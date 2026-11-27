# Security policy

## Supported versions
Only the latest release gets security fixes.

## Reporting a vulnerability
Please report it privately: on GitHub, **Security > Report a vulnerability** on this repository. Don't open a public
issue. Say what is affected, how to reproduce it (a sample OST, PST or MBOX that triggers it, if you can share one
safely), and the version.

You should get a reply within a week. Fixes are released as a new version, with the issue described in the release
notes once the release is out.

## How Unbound is built and released
- Unbound works entirely on the PC. Its only internet connection is **Check for updates**, and only when the user
  clicks it: it asks GitHub for the latest release and, if the user agrees, downloads `Unbound.exe` from this
  repository's releases and checks it against the release's `SHA256SUMS.txt` before using it (see the README's
  Privacy section).
- Every dependency is pinned with checksums (`requirements*.txt`, installed with `--require-hashes`), and CI checks
  them against the public advisory databases on every push and weekly.
- Releases are built by GitHub Actions from the tagged code (`.github/workflows/release.yml`), never on someone's own
  PC. Each release has `SHA256SUMS.txt` and a build provenance attestation:
  `gh attestation verify Unbound.exe --repo Z3R0ZIGMA/Unbound`.
- From 1.3.6 on, releases are immutable: once published, a release's files and its tag can't be changed.
