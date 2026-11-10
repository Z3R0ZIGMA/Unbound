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
