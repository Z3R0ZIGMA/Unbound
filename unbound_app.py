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
                 "unbound.mbox_reader", "unbound.pff_process", "unbound.timezones",
                 "unbound.updates"):
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

    # Outlook's objects are wrapped so that what's learnt about a type is
    # shared (pst_writer._Dispatch). That leans on pywin32's insides, so try
    # it with a COM object every Windows has
    try:
        import win32com.client
        from unbound.pst_writer import _Dispatch, _wrap
        types = {}
        first, second = (_wrap(types, win32com.client.Dispatch("Scripting.Dictionary")._oleobj_)
                         for _ in range(2))
        first.Add("a", 1)
        second.Add("b", 2)
        shared = isinstance(second, _Dispatch) and second._olerepr_ is first._olerepr_
        if shared and (first.Count, second.Count, second.Item("b")) == (1, 1, 2):
            lines.append("OK   shared COM types")
        else:
            ok = False
            lines.append("FAIL shared COM types: objects of one type don't share")
    except Exception as e:
        ok = False
        lines.append(f"FAIL shared COM types: {e}")

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

    # Check for updates talks to GitHub over HTTPS, which needs Python's ssl
    # module and Windows' trusted certificates. Nothing is sent here
    try:
        import ssl
        context = ssl.create_default_context()
        roots = len(context.get_ca_certs())
        if roots:
            lines.append(f"OK   secure connections ({ssl.OPENSSL_VERSION}, {roots} trusted roots)")
        else:
            ok = False
            lines.append("FAIL secure connections: no trusted certificates found")
    except Exception as e:
        ok = False
        lines.append(f"FAIL secure connections: {e}")

    try:
        from unbound.app import _icon_path
        if os.path.exists(_icon_path()):
            lines.append("OK   icon")
        else:
            ok = False
            lines.append(f"FAIL icon: not found at {_icon_path()}")
    except Exception as e:
        ok = False
        lines.append(f"FAIL icon: {e}")
    with open(report_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    return 0 if ok else 1


def _guarded(fn, *args, fallback_file=None, code=1):
    """Run a test mode so that any error ends in an exit code, never in the
    windowed EXE's "Unhandled exception" box, which would wait for a click
    (and in CI, for hours)."""
    try:
        return fn(*args)
    except BaseException as e:  # noqa: BLE001 - an exit code is the whole point
        if fallback_file:
            try:
                with open(fallback_file, "w", encoding="utf-8") as fh:
                    fh.write(f"FAIL {type(e).__name__}: {e}\n")
            except OSError:
                pass
        return code


def convert_without_window(source, pst):
    from unbound.converter import (
        convert, log_path_for, recover_interrupted, scan_folders, write_failure_log, write_log,
    )
    log = log_path_for(pst)
    try:
        nodes = [n for top in scan_folders(source) for n in top.walk() if n.default_checked]
        result = convert(source, pst, nodes, detach=recover_interrupted())
    except Exception as e:
        try:
            write_failure_log(e, source, pst, log)
        except OSError:
            pass  # e.g. the output folder doesn't exist; the exit code still says it failed
        return 2
    write_log(result, source, pst, log)
    return 1 if result.failures or result.still_attached else 0


if __name__ == "__main__":
    # Must come first: the reader process is this same program started again
    multiprocessing.freeze_support()
    if len(sys.argv) == 3 and sys.argv[1] == "--check":
        sys.exit(_guarded(self_check, sys.argv[2], fallback_file=sys.argv[2], code=1))
    if len(sys.argv) == 4 and sys.argv[1] == "--convert":
        sys.exit(_guarded(convert_without_window, os.path.abspath(sys.argv[2]),
                          os.path.abspath(sys.argv[3]), code=2))
    from unbound.app import main
    main()
