"""End-to-end MBOX test: build a test Apple Mail export, convert it to a PST
through Outlook, read the PST back with pypff and check every message.

Usage:
    .venv\\Scripts\\python scripts\\mbox_test.py WORK_FOLDER

Needs Outlook Classic. WORK_FOLDER is created and holds the export and PST.
"""

import json
import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pypff  # noqa: E402

from unbound import mapi  # noqa: E402
from unbound.converter import convert, scan_folders, write_log  # noqa: E402
from unbound.ost_reader import _entries, _integer, _string, open_file, read_message  # noqa: E402


def main():
    work = os.path.abspath(sys.argv[1])
    export = os.path.join(work, "export")
    pst = os.path.join(work, "mbox_test.pst")
    if os.path.exists(pst):
        os.remove(pst)
    subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), "make_mbox_fixture.py"),
                    export], check=True)
    expected = json.load(open(os.path.join(export, "expected.json"), encoding="utf-8"))

    nodes = [n for top in scan_folders(export) for n in top.walk()]
    r = convert(export, pst, nodes)
    write_log(r, export, pst, os.path.join(work, "mbox_test - Unbound log.txt"))
    print(f"Converted: {dict(r.copied)}, skipped {r.skipped}, failed {len(r.failures)}")
    for fl in r.failures:
        print("  failure:", fl.folder, fl.error)

    for _ in range(120):
        try:
            f = open_file(pst)
            break
        except OSError:
            time.sleep(1)
    else:
        print("PST still locked")
        return 1
    with open(pst, "rb") as fh:
        version = int.from_bytes(fh.read(12)[10:12], "little")

    got = []

    def visit(fo, trail):
        for i in range(fo.number_of_sub_messages):
            it = fo.get_sub_message(i)
            if not isinstance(it, pypff.message):
                continue
            m = read_message(it)
            props = _entries(it.record_sets[0])
            atts = []
            for k in range(it.number_of_attachments):
                a = it.get_attachment(k)
                ap = _entries(a.record_sets[0])
                hidden = ap.get(0x7FFE)
                atts.append({"filename": a.long_filename, "data": a.read_buffer(a.size or 0),
                             "content_id": _string(ap, 0x3712),
                             "hidden": bool(hidden and hidden.data and hidden.data[0])})
            got.append(("/".join(trail), m, _integer(props, mapi.MESSAGE_FLAGS) or 0, atts))
        for i in range(fo.number_of_sub_folders):
            s = fo.get_sub_folder(i)
            visit(s, trail + (s.name,))
    visit(f.get_root_folder(), ())
    f.close()

    print(f"PST format version {version} (23 = Unicode)")
    failures = 0
    for e in expected:
        hits = [g for g in got if g[0].endswith(e["folder"]) and g[1].subject == e["subject"]]
        print(f"\n== {e['folder']} / {e['subject']}")
        if not hits:
            print("  FAIL  not found in the PST (folder or subject wrong)")
            failures += 1
            continue
        _, m, flags, atts = hits[0]
        src_atts = {a["filename"]: a for a in e["attachments"]}
        checks = {
            "sender": (m.sender_name, m.sender_email.lower()) == (e["sender_name"], e["sender_email"]),
            "To": sorted(x.email.lower() for x in m.recipients if x.kind == mapi.MAPI_TO) == sorted(e["to"]),
            "Cc": sorted(x.email.lower() for x in m.recipients if x.kind == mapi.MAPI_CC) == sorted(e["cc"]),
            "sent time": m.submit_time is not None and m.submit_time.isoformat() == e["date"],
            "received time": m.delivery_time is not None and m.delivery_time.isoformat() == e["date"],
            "read state": bool(flags & mapi.MSGFLAG_READ) == e["read"],
            "not a draft": not flags & mapi.MSGFLAG_UNSENT,
            "HTML body": bool(m.html_body) == e["has_html"],
            "attachments": sorted((a["filename"], len(a["data"])) for a in atts)
                           == sorted((a["filename"], a["size"]) for a in e["attachments"]),
            "inline image id": all((a["content_id"] or None) == src_atts[a["filename"]]["content_id"]
                                   for a in atts if a["filename"] in src_atts),
            "inline image linked to HTML": all(f"cid:{a['content_id']}" in (m.html_body or "")
                                               for a in atts if a["content_id"]),
        }
        if e["subject"] == "Plain text only":
            body = m.plain_body or ""
            checks["'From ' line unquoted"] = "\nFrom the start of a line" in body and ">From the start" not in body
        for k, ok in checks.items():
            print(f"  {'PASS' if ok else 'FAIL'}  {k}")
            failures += not ok
    print(f"\n{'ALL PASSED' if not failures else f'{failures} checks failed'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
