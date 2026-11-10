# Unbound

Windows desktop app that converts an Outlook OST file into a PST. See [PLAN.md](PLAN.md).

## Setup (Windows)

Requires Python 3.10 or newer (developed on 3.15). libpff-python ships prebuilt Windows wheels, so no C compiler is needed.

```bash
py -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

Check it works:

```bash
.venv\Scripts\python -c "import pypff; print(pypff.get_version())"
```

## Run the app

```bash
.venv\Scripts\python unbound_app.py
```

Needs Outlook Classic installed (it writes the PST). Version 1 converts mail only;
calendar, contacts and tasks folders are shown but greyed out. A log of anything
skipped or failed is saved next to the PST as `<name> - Unbound log.txt`.
