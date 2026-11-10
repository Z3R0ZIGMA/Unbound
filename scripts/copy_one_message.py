"""Step 4 proof: copy one message from an OST into a new PST via Outlook,
then read the PST back with pypff and check the fields survived.

Usage:
    .venv\\Scripts\\python scripts\\copy_one_message.py SOURCE.ost OUTPUT.pst

Picks the first message in an Inbox that has an HTML body and an
attachment. Prints only pass/fail per field, never message content.
"""

import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pypff  # noqa: E402

from unbound import mapi  # noqa: E402
from unbound.ost_reader import open_file, read_message, _entries, _integer, _string  # noqa: E402
from unbound.pst_writer import PstWriter  # noqa: E402


def find_sample(folder, in_inbox=False):
    in_inbox = in_inbox or (folder.name or "").lower() == "inbox"
    if in_inbox:
        for i in range(folder.number_of_sub_messages):
            item = folder.get_sub_message(i)
            if not isinstance(item, pypff.message) or item.number_of_attachments == 0:
                continue
            m = read_message(item)
            if m.html_body and m.message_class == "IPM.Note" and m.sender_email:
                return m
    for i in range(folder.number_of_sub_folders):
        found = find_sample(folder.get_sub_folder(i), in_inbox)
        if found:
            return found
    return None


def first_message(folder):
    for i in range(folder.number_of_sub_messages):
        item = folder.get_sub_message(i)
        if isinstance(item, pypff.message):
            return item
    for i in range(folder.number_of_sub_folders):
        found = first_message(folder.get_sub_folder(i))
        if found:
            return found
    return None


def outlook_running():
    import subprocess
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/NH"],
                         capture_output=True, text=True, encoding="mbcs",
                         errors="replace").stdout
    return "OUTLOOK.EXE" in out.upper()


def main():
    src, out = sys.argv[1], os.path.abspath(sys.argv[2])
    if os.path.exists(out):
        print(f"{out} already exists; delete it first")
        return 1
    os.makedirs(os.path.dirname(out), exist_ok=True)

    f = open_file(src)
    m = find_sample(f.get_root_folder())
    f.close()
    if not m:
        print("No suitable sample message found")
        return 1
    print(f"Sample: {len(m.attachments)} attachment(s), {len(m.recipients)} recipient(s), "
          f"read={m.is_read}")

    was_running = outlook_running()
    w = PstWriter(out, display_name="Unbound test")
    try:
        w.add_message(w.get_folder(["Inbox"]), m)
    finally:
        w.close(quit_outlook=not was_running)
    errors = w.property_errors
    del w
    if errors:
        print("Outlook refused these properties:")
        for k, v in errors.items():
            print(f"  {k}: {v.splitlines()[0][:150]}")

    # Read the PST back independently of Outlook. Outlook keeps the file
    # locked for a few seconds after detaching/quitting, so retry.
    for _ in range(120):
        try:
            pst = open_file(out)
            break
        except OSError:
            time.sleep(1)
    else:
        print("FAIL: PST still locked by Outlook after 120s")
        return 1
    raw = first_message(pst.get_root_folder())
    if raw is None:
        print("FAIL: no message found in the PST")
        return 1
    back = read_message(raw)
    props = _entries(raw.record_sets[0])
    flags = _integer(props, mapi.MESSAGE_FLAGS) or 0
    headers_ok = bool(back.transport_headers) == bool(m.transport_headers)
    pst.close()

    def norm(s):
        return " ".join((s or "").split())

    def same_second(a, b):
        return bool(a and b) and abs((a - b).total_seconds()) < 1

    def recips(msg):
        return sorted((r.kind, r.email.lower()) for r in msg.recipients)

    checks = {
        "subject": back.subject == m.subject,
        "message class": back.message_class == "IPM.Note",
        "not a draft (unsent flag clear)": not (flags & mapi.MSGFLAG_UNSENT),
        "read/unread state": back.is_read == m.is_read,
        "sender name": back.sender_name == m.sender_name,
        "sender email": back.sender_email.lower() == m.sender_email.lower(),
        "delivery time": same_second(back.delivery_time, m.delivery_time),
        "submit time": same_second(back.submit_time, m.submit_time),
        "HTML body present": bool(back.html_body),
        "plain body text matches": norm(back.plain_body)[:500] == norm(m.plain_body)[:500],
        "attachment names": sorted(a.filename for a in back.attachments)
                            == sorted(a.filename for a in m.attachments),
        "attachment bytes": sorted(a.data for a in back.attachments)
                            == sorted(a.data for a in m.attachments),
        "recipients (To/Cc/Bcc + addresses)": recips(back) == recips(m),
        "internet headers": headers_ok,
    }
    for name, ok in checks.items():
        print(f"  {'PASS' if ok else 'FAIL'}  {name}")
    if recips(back) != recips(m):
        print(f"  recipients: {len(back.recipients)} in PST vs {len(m.recipients)} in source")
    if not same_second(back.delivery_time, m.delivery_time):
        print(f"  delivery time off by {back.delivery_time - m.delivery_time}"
              if back.delivery_time and m.delivery_time else "  delivery time missing")
    print(f"\n{sum(checks.values())}/{len(checks)} checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
