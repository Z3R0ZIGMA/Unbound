"""Start the Unbound app: .venv\\Scripts\\python unbound_app.py [file.ost]

Test modes (used to check a build; Unbound.exe has no console, so results go
to files and the exit code):

  Unbound.exe --check REPORT.txt
      Loads every library the app needs, starts a reader process and passes
      a date through COM the way the converter does, then writes the result
      to REPORT.txt. Exit 0 = OK.

  Unbound.exe --convert SOURCE OUTPUT.pst
      Converts the folders that are ticked by default, without the window,
      and writes the usual log next to OUTPUT.pst. Exit 0 = nothing failed.
"""

import multiprocessing
import os
import sys


def self_check(report_path):
    from unbound import __version__
    lines, ok = [f"Unbound {__version__}"], True
    for name in ("pypff", "pythoncom", "pywintypes", "win32com.client", "win32timezone",
                 "PySide6.QtWidgets", "unbound.app", "unbound.converter",
                 "unbound.pst_writer", "unbound.props", "unbound.appointment",
                 "unbound.mbox_reader", "unbound.pff_process", "unbound.timezones"):
        try:
            mod = __import__(name, fromlist=["_"])
            extra = f" {mod.get_version()}" if name == "pypff" else ""
            lines.append(f"OK   {name}{extra}")
        except Exception as e:
            ok = False
            lines.append(f"FAIL {name}: {e}")

    # OST and PST files are read in a child process, which a packaged EXE
    # starts differently from Python run from source
    try:
        from unbound.pff_process import child_process_works
        if child_process_works():
            lines.append("OK   reader process")
        else:
            ok = False
            lines.append("FAIL reader process: it didn't answer")
    except Exception as e:
        ok = False
        lines.append(f"FAIL reader process: {e}")

    # The converter hands Outlook dates marked as UTC; pywin32 loads extra
    # modules for that only when it happens, so do it once for real
    try:
        from datetime import datetime
        import win32com.client
        from unbound.pst_writer import com_date, from_com_date
        sent = datetime(2026, 3, 15, 9, 30)
        d = win32com.client.Dispatch("Scripting.Dictionary")
        d.Add("date", com_date(sent))
        back = from_com_date(d.Item("date"))
        if back == sent:
            lines.append("OK   dates through COM")
        else:
            ok = False
            lines.append(f"FAIL dates through COM: sent {sent}, got {back}")
    except Exception as e:
        ok = False
        lines.append(f"FAIL dates through COM: {e}")

    # Not a failure when missing: Windows before 10 (1903) has no ICU, and
    # such appointments then keep their times but not their time zone name
    try:
        from unbound.timezones import windows_zone
        zone = windows_zone("America/New_York")
        lines.append(f"OK   time zone names (America/New_York = {zone})" if zone
                     else "OK   time zone names: not available on this Windows")
    except Exception as e:
        ok = False
        lines.append(f"FAIL time zone names: {e}")

    from unbound.app import _icon_path
    if os.path.exists(_icon_path()):
        lines.append("OK   icon")
    else:
        ok = False
        lines.append(f"FAIL icon: not found at {_icon_path()}")
    with open(report_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    return 0 if ok else 1


def convert_without_window(source, pst):
    from unbound.converter import (
        convert, log_path_for, recover_interrupted, scan_folders, write_failure_log, write_log,
    )
    log = log_path_for(pst)
    try:
        nodes = [n for top in scan_folders(source) for n in top.walk() if n.default_checked]
        result = convert(source, pst, nodes, detach=recover_interrupted())
    except Exception as e:
        write_failure_log(e, source, pst, log)
        return 2
    write_log(result, source, pst, log)
    return 1 if result.failures else 0


if __name__ == "__main__":
    # Must come first: the reader process is this same program started again
    multiprocessing.freeze_support()
    if len(sys.argv) == 3 and sys.argv[1] == "--check":
        sys.exit(self_check(sys.argv[2]))
    if len(sys.argv) == 4 and sys.argv[1] == "--convert":
        sys.exit(convert_without_window(os.path.abspath(sys.argv[2]), os.path.abspath(sys.argv[3])))
    from unbound.app import main
    main()
