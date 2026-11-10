"""Start the Unbound app: .venv\\Scripts\\python unbound_app.py [file.ost]

`Unbound.exe --check REPORT.txt` loads every library the app needs, writes
the result to REPORT.txt and exits (0 = all good). Used to test the build.
"""

import sys


def self_check(report_path):
    lines, ok = [], True
    for name in ("pypff", "pythoncom", "pywintypes", "win32com.client",
                 "PySide6.QtWidgets", "unbound.app", "unbound.converter",
                 "unbound.pst_writer"):
        try:
            mod = __import__(name, fromlist=["_"])
            extra = f" {mod.get_version()}" if name == "pypff" else ""
            lines.append(f"OK   {name}{extra}")
        except Exception as e:
            ok = False
            lines.append(f"FAIL {name}: {e}")
    with open(report_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    return 0 if ok else 1


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--check":
        sys.exit(self_check(sys.argv[2]))
    from unbound.app import main
    main()
