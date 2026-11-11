"""Read messages out of an OST/PST with pypff into plain Python objects."""

import os
from dataclasses import dataclass, field
from datetime import datetime

import pypff

from . import mapi


@dataclass
class Recipient:
    name: str
    email: str
    kind: int  # mapi.MAPI_TO / MAPI_CC / MAPI_BCC


@dataclass
class Attachment:
    filename: str
    data: bytes


@dataclass
class Message:
    message_class: str
    subject: str
    sender_name: str
    sender_email: str
    delivery_time: datetime | None  # UTC
    submit_time: datetime | None  # UTC
    is_read: bool
    plain_body: str | None
    html_body: str | None
    transport_headers: str | None
    recipients: list[Recipient] = field(default_factory=list)
    attachments: list[Attachment] = field(default_factory=list)


class OstInUseError(Exception):
    """The file is locked, almost always by Outlook using it for its profile."""

    def __init__(self, path):
        super().__init__(
            "Outlook is using this file, so it can't be read right now.\n\n"
            "It's the data file of the Outlook profile on this PC. Close Outlook, "
            "wait a few seconds for it to finish closing, then try again.\n\n"
            f"File: {path}")


def is_lock_error(e):
    text = str(e).lower()
    return ("locked a portion" in text or "being used by another process" in text
            or getattr(e, "winerror", None) in (32, 33))


def open_file(path):
    f = pypff.file()
    # libpff uses \\?\ long paths on Windows, which reject forward slashes
    path = os.path.abspath(path)
    try:
        f.open(path)
    except OSError as e:
        if is_lock_error(e):
            raise OstInUseError(path) from e
        raise
    return f


def _entries(record_set):
    return {e.entry_type: e for e in record_set.entries}


def _string(entries, prop_id):
    e = entries.get(prop_id)
    if e is None:
        return None
    try:
        return e.data_as_string
    except (OSError, TypeError):
        return None


def _integer(entries, prop_id):
    e = entries.get(prop_id)
    if e is None:
        return None
    try:
        return e.data_as_integer
    except (OSError, TypeError):
        return None


def _optional(fn):
    """pypff raises OSError when a value (e.g. an HTML body) is absent."""
    try:
        return fn()
    except OSError:
        return None


def _decode_html(raw, codepage):
    if raw is None:
        return None
    for enc in (f"cp{codepage}" if codepage else None, "utf-8", "cp1252"):
        if not enc:
            continue
        try:
            return raw.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("utf-8", errors="replace")


def read_message(msg):
    props = _entries(msg.record_sets[0])
    flags = _integer(props, mapi.MESSAGE_FLAGS) or 0
    codepage = _integer(props, mapi.INTERNET_CPID)

    sender_email = (
        _string(props, mapi.SENDER_SMTP_ADDRESS)
        or _string(props, mapi.SENT_REPRESENTING_SMTP_ADDRESS)
        or _string(props, mapi.SENDER_EMAIL_ADDRESS)
        or ""
    )

    recipients = []
    rec_obj = _optional(lambda: msg.recipients)
    if rec_obj is not None:
        for rs in rec_obj.record_sets:
            r = _entries(rs)
            recipients.append(Recipient(
                name=_string(r, mapi.DISPLAY_NAME) or "",
                email=_string(r, mapi.SMTP_ADDRESS) or _string(r, mapi.EMAIL_ADDRESS) or "",
                kind=_integer(r, mapi.RECIPIENT_TYPE) or mapi.MAPI_TO,
            ))

    attachments = []
    for i in range(msg.number_of_attachments):
        a = msg.get_attachment(i)
        size = a.size or 0
        data = a.read_buffer(size) if size else b""
        name = _optional(lambda: a.long_filename) or f"attachment_{i + 1}"
        attachments.append(Attachment(filename=name, data=data))

    plain = _optional(lambda: msg.plain_text_body)
    return Message(
        message_class=_string(props, mapi.MESSAGE_CLASS) or "IPM.Note",
        subject=msg.subject or "",
        sender_name=msg.sender_name or "",
        sender_email=sender_email,
        delivery_time=_optional(lambda: msg.delivery_time),
        submit_time=_optional(lambda: msg.client_submit_time),
        is_read=bool(flags & mapi.MSGFLAG_READ),
        plain_body=plain.decode("utf-8", errors="replace") if plain else None,
        html_body=_decode_html(_optional(lambda: msg.html_body), codepage),
        transport_headers=_optional(lambda: msg.transport_headers),
        recipients=recipients,
        attachments=attachments,
    )
