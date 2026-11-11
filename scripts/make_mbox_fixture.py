"""Build a test Apple Mail export: nested Name.mbox folders, each holding an
"mbox" file and an Info.plist, the way Mail's Mailbox > Export Mailbox does.

Usage:
    .venv\\Scripts\\python scripts\\make_mbox_fixture.py OUTPUT_FOLDER

The messages cover the cases a converter gets wrong: encoded and emoji
subjects, HTML with inline (cid:) images, attachments, many recipients,
lines starting with "From " (quoted as ">From " in mbox), plain text only,
and a missing Status header. Writes expected.json describing each message so
a test can check the converted PST against it.
"""

import json
import os
import sys
from email.headerregistry import Address
from email.message import EmailMessage
from email.utils import format_datetime
from datetime import datetime, timedelta, timezone

PNG = bytes.fromhex(  # 1x1 red pixel
    "89504e470d0a1a0a0000000d4948445200000001000000010802000000907753de"
    "0000000c4944415408d763f8cfc000000301010018dd8db00000000049454e44ae426082")
PDF = b"%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n" + bytes(range(256)) * 4

INFO_PLIST = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>LocalFileVersion</key>
	<integer>1</integer>
</dict>
</plist>
"""


def base(subject, sender, to, when, cc=(), status="RO"):
    m = EmailMessage()
    m["Subject"] = subject
    m["From"] = Address(*sender)
    m["To"] = [Address(*t) for t in to]
    if cc:
        m["Cc"] = [Address(*c) for c in cc]
    m["Date"] = format_datetime(when)
    m["Message-Id"] = f"<{abs(hash((subject, when))):x}@example.test>"
    if status:
        m["Status"] = status
    return m


def build():
    t0 = datetime(2026, 3, 2, 9, 15, tzinfo=timezone(timedelta(hours=2)))
    alice = ("Alice Müller", "alice", "example.test")
    bob = ("Bob Smith", "bob", "example.test")
    carol = ("Carol 王", "carol", "example.test")
    me = ("Me", "me", "example.test")
    msgs = {}

    m = base("Plain text only", alice, [me], t0)
    m.set_content("Hello,\nFrom the start of a line this must survive mbox quoting.\n\nAlice\n")
    msgs["Inbox"] = [m]

    m = base("Café ☕ meeting 🎉 – Zürich", bob, [me, alice], t0 + timedelta(hours=3), cc=[carol])
    m.set_content("Plain part with ünïcödé.\n")
    m.add_alternative("<html><body><p>HTML part with <b>ünïcödé</b> 🎉</p>"
                      "<p>Logo: <img src=\"cid:logo@example.test\"></p></body></html>",
                      subtype="html")
    m.get_payload()[1].add_related(PNG, "image", "png", cid="<logo@example.test>",
                                   filename="logo.png")
    msgs["Inbox"].append(m)

    m = base("Report attached", carol, [me], t0 + timedelta(days=1), status="O")  # seen, not read
    m.set_content("See the attached report and photo.\n")
    m.add_attachment(PDF, "application", "pdf", filename="Quarterly report.pdf")
    m.add_attachment(PNG, "image", "png", filename="photo.png")
    msgs["Inbox"].append(m)

    m = base("Sent from my Mac", me, [alice, bob, carol], t0 + timedelta(days=2))
    m.set_content("A reply.\n\n> From: Alice\n> quoted text\n")
    msgs["Sent Messages"] = [m]

    m = base("Project kickoff", alice, [me], t0 + timedelta(days=3), status=None)  # no Status, as Apple writes
    m.set_content("Nested mailbox message.\n")
    msgs["Projects/Alpha"] = [m]

    m = base("=?UTF-8?B?w4RwZmVsIHVuZCBCaXJuZW4=?=", bob, [me], t0 + timedelta(days=4))
    m.set_content("Subject was sent as an encoded word.\n")
    msgs["Projects"] = [m]
    return msgs


def write_mbox(path, messages):
    with open(path, "wb") as fh:
        for m in messages:
            raw = m.as_bytes(policy=m.policy.clone(linesep="\n"))
            # mboxrd quoting, as Apple Mail writes it
            lines = []
            for line in raw.split(b"\n"):
                stripped = line.lstrip(b">")
                lines.append(b">" + line if stripped.startswith(b"From ") else line)
            fh.write(b"From " + str(m["From"].addresses[0].addr_spec).encode() +
                     b" Mon Mar  2 09:15:00 2026\n")
            fh.write(b"\n".join(lines).rstrip(b"\n") + b"\n\n")


def main():
    out = os.path.abspath(sys.argv[1])
    os.makedirs(out, exist_ok=True)
    expected = []
    for mailbox, messages in build().items():
        folder = out
        for part in mailbox.split("/"):
            folder = os.path.join(folder, part + ".mbox")
        os.makedirs(folder, exist_ok=True)
        write_mbox(os.path.join(folder, "mbox"), messages)
        with open(os.path.join(folder, "Info.plist"), "w", encoding="utf-8") as fh:
            fh.write(INFO_PLIST)
        for m in messages:
            atts = [{"filename": p.get_filename(), "size": len(p.get_content()),
                     "content_id": (p["Content-ID"] or "").strip("<>") or None}
                    for p in m.walk() if p.get_filename()]
            expected.append({
                "folder": mailbox, "subject": str(m["Subject"]),
                "sender_email": m["From"].addresses[0].addr_spec,
                "sender_name": m["From"].addresses[0].display_name,
                "to": [a.addr_spec for a in m["To"].addresses],
                "cc": [a.addr_spec for a in m["Cc"].addresses] if m["Cc"] else [],
                "date": m["Date"].datetime.astimezone(timezone.utc).replace(tzinfo=None).isoformat(),
                "read": m["Status"] is None or "R" in str(m["Status"]),  # no Status header = read
                "has_html": any(p.get_content_type() == "text/html" for p in m.walk()),
                "attachments": atts,
            })
    with open(os.path.join(out, "expected.json"), "w", encoding="utf-8") as fh:
        json.dump(expected, fh, ensure_ascii=False, indent=1)
    print(f"Wrote {len(expected)} messages in {out}")


if __name__ == "__main__":
    main()
