"""Read Apple Mail exports and other MBOX files.

Apple Mail's Mailbox > Export Mailbox writes a folder named "Name.mbox"
holding a file called "mbox" (plus Info.plist and sometimes a
table_of_contents); exporting a mailbox with sub-mailboxes nests further
"Child.mbox" folders inside it. A plain "Name.mbox" file (as other mail
programs write) is accepted too.

Messages are parsed with Python's email package into the same Message
record the OST reader produces, so the PST writer handles both alike.
"""

import os
import re
from dataclasses import dataclass, field
from datetime import timezone
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime

from . import mapi
from .ost_reader import Attachment, Message, Recipient

MBOX_FILE = "mbox"  # the data file inside an Apple Mail .mbox folder
_FROM_LINE = re.compile(rb"^From \S")
_QUOTED_FROM = re.compile(rb"^>(>*From )", re.M)


@dataclass
class Mailbox:
    name: str
    path: tuple  # mailbox names from the top of the export down to this one
    file: str | None  # the mbox data file, or None for a folder with only children
    children: list = field(default_factory=list)


def is_mbox_source(path):
    """True for an Apple Mail export folder, a .mbox folder or an mbox file."""
    path = os.path.abspath(path)
    if os.path.isdir(path):
        return path.lower().endswith(".mbox") or any(
            e.lower().endswith(".mbox") for e in os.listdir(path))
    if path.lower().endswith(".mbox") or os.path.basename(path) == MBOX_FILE:
        return True
    try:
        with open(path, "rb") as fh:
            return bool(_FROM_LINE.match(fh.read(64)))
    except OSError:
        return False


def _mailbox_name(entry):
    return entry[:-5] if entry.lower().endswith(".mbox") else entry


def find_mailboxes(path):
    """The mailbox tree under an export folder, a .mbox folder or a file."""
    path = os.path.abspath(path)
    if os.path.isfile(path):
        name = os.path.basename(path)
        if name == MBOX_FILE:  # the file inside an Apple .mbox folder
            name = os.path.basename(os.path.dirname(path))
        return [Mailbox(_mailbox_name(name), (_mailbox_name(name),), path)]
    if path.lower().endswith(".mbox"):  # one exported mailbox folder
        return [_bundle(path, ())]
    return _children(path, ())


def _bundle(folder, parent_path):
    name = _mailbox_name(os.path.basename(folder))
    data = os.path.join(folder, MBOX_FILE)
    box = Mailbox(name, parent_path + (name,), data if os.path.isfile(data) else None)
    box.children = _children(folder, box.path)
    return box


def _children(folder, parent_path):
    boxes = []
    for entry in sorted(os.listdir(folder), key=str.lower):
        full = os.path.join(folder, entry)
        if not entry.lower().endswith(".mbox"):
            continue
        if os.path.isdir(full):
            boxes.append(_bundle(full, parent_path))
        elif os.path.isfile(full):
            name = _mailbox_name(entry)
            boxes.append(Mailbox(name, parent_path + (name,), full))
    return boxes


# ---------- splitting an mbox into messages ----------

_offsets_cache = {}


def message_offsets(file):
    """Byte offsets where each message starts (its "From " line)."""
    st = os.stat(file)
    key = (os.path.abspath(file), st.st_size, st.st_mtime_ns)
    if key in _offsets_cache:
        return _offsets_cache[key]
    offsets = []
    pos = 0
    previous_blank = True
    with open(file, "rb") as fh:
        for line in fh:
            if previous_blank and _FROM_LINE.match(line):
                offsets.append(pos)
            previous_blank = line in (b"\n", b"\r\n")
            pos += len(line)
    _offsets_cache[key] = offsets
    return offsets


def read_raw(file, index):
    """The message bytes (without its "From " line), quoting undone."""
    offsets = message_offsets(file)
    start = offsets[index]
    end = offsets[index + 1] if index + 1 < len(offsets) else os.path.getsize(file)
    with open(file, "rb") as fh:
        fh.seek(start)
        data = fh.read(end - start)
    data = data.split(b"\n", 1)[1] if b"\n" in data else b""
    # mboxrd quoting: ">From " in the file was "From " in the message
    return _QUOTED_FROM.sub(rb"\1", data)


# ---------- turning a message into a Message ----------

def _utc(dt):
    if dt is None:
        return None
    if dt.tzinfo is None:
        return dt
    return dt.astimezone(timezone.utc).replace(tzinfo=None)


def _date(value):
    try:
        return _utc(parsedate_to_datetime(str(value))) if value else None
    except (TypeError, ValueError, IndexError):
        return None


def _addresses(msg, header):
    values = msg.get_all(header, [])
    return [(n, a) for n, a in getaddresses([str(v) for v in values]) if a or n]


def _text(part):
    try:
        return part.get_content()
    except (LookupError, UnicodeError, KeyError, AssertionError):
        payload = part.get_payload(decode=True) or b""
        return payload.decode("utf-8", errors="replace")


def parse_message(raw):
    msg = BytesParser(policy=policy.default).parsebytes(raw)

    sender_name, sender_email = (_addresses(msg, "From") or [("", "")])[0]
    recipients = []
    for header, kind in (("To", mapi.MAPI_TO), ("Cc", mapi.MAPI_CC), ("Bcc", mapi.MAPI_BCC)):
        recipients += [Recipient(n, a, kind) for n, a in _addresses(msg, header)]

    sent = _date(msg["Date"])
    received = None
    first_received = msg.get_all("Received", [])[:1]  # the last hop is listed first
    if first_received and ";" in str(first_received[0]):
        received = _date(str(first_received[0]).rsplit(";", 1)[1].strip())

    # Apple Mail exports usually have no Status header; treat those as read
    status = str(msg.get("Status", "")) + str(msg.get("X-Status", ""))
    is_read = "R" in status if (msg["Status"] is not None or msg["X-Status"] is not None) else True

    plain_part = msg.get_body(preferencelist=("plain",))
    html_part = msg.get_body(preferencelist=("html",))
    body_parts = {id(p) for p in (plain_part, html_part) if p is not None}

    attachments, embedded = [], 0
    for part in msg.walk():
        if part.is_multipart() or id(part) in body_parts:
            continue
        ctype = part.get_content_type()
        filename = part.get_filename()
        cid = (part["Content-ID"] or "").strip().strip("<>") or None
        if ctype == "message/rfc822":
            inner = part.get_payload()
            inner = inner[0] if isinstance(inner, list) and inner else inner
            data = inner.as_bytes() if hasattr(inner, "as_bytes") else bytes(part.get_payload(decode=True) or b"")
            subject = (inner.get("Subject", "") if hasattr(inner, "get") else "") or "Attached message"
            attachments.append(Attachment(_safe_name(str(subject)) + ".eml", data))
            continue
        if part.get_content_maintype() == "text" and not filename and not cid:
            continue  # stray extra text parts of the body
        data = part.get_payload(decode=True) or b""
        if not filename:
            ext = {"image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif",
                   "application/pdf": ".pdf"}.get(ctype, ".bin")
            filename = f"attachment_{len(attachments) + 1}{ext}"
        attachments.append(Attachment(filename, data, content_id=cid))

    header_block = raw.split(b"\n\n", 1)[0].decode("utf-8", errors="replace")
    return Message(
        message_class="IPM.Note",
        subject=str(msg["Subject"] or ""),
        sender_name=sender_name,
        sender_email=sender_email,
        delivery_time=received or sent,
        submit_time=sent or received,
        is_read=is_read,
        plain_body=_text(plain_part) if plain_part is not None else None,
        html_body=_text(html_part) if html_part is not None else None,
        rtf_body=None,
        transport_headers=header_block,
        recipients=recipients,
        attachments=attachments,
        skipped_embedded=embedded,
    )


def _safe_name(text):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text).strip(" .")[:100] or "Attached message"
