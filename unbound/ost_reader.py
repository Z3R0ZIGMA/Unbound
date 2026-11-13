"""Read messages out of an OST/PST with pypff into plain Python objects."""

import ctypes
import mmap
import os
from ctypes import wintypes
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
    content_id: str | None = None  # set for images shown inside HTML (cid: links)


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
    rtf_body: bytes | None
    transport_headers: str | None
    recipients: list[Recipient] = field(default_factory=list)
    attachments: list[Attachment] = field(default_factory=list)
    # Attached emails (embedded messages) can't be recreated yet; counted
    # so the log can report them
    skipped_embedded: int = 0


ATTACH_METHOD = 0x3705
ATTACH_EMBEDDED_MSG = 5
ATTACHMENT_HIDDEN = 0x7FFE
ATTACH_CONTENT_ID = 0x3712
ATTACH_SHORT_FILENAME = 0x3704  # 8.3 name, used when there is no long one


def _has_outlook(programs):
    return any("outlook" in p.lower() for p in programs)


def _names(programs):
    """'Word', 'Word and Excel', 'Word, Excel and Teams'."""
    programs = list(programs)
    if len(programs) <= 1:
        return "".join(programs)
    return ", ".join(programs[:-1]) + " and " + programs[-1]


class OstInUseError(OSError):
    """The file can't be read because another program is using it.

    `programs` names the ones Windows reports as having it open; `outlook`
    says whether Outlook is among them.
    """

    def __init__(self, path, programs=()):
        self.programs = list(programs)
        self.outlook = _has_outlook(self.programs)
        if self.outlook:
            text = ("Outlook is using this file, so it can't be read right now.\n\n"
                    "Close Outlook and try again. If you can't see an Outlook window, it's "
                    "running in the background: Unbound starts Outlook to write a PST, and "
                    "Outlook keeps running afterwards.")
        elif self.programs:
            one = len(self.programs) == 1
            text = (f"{_names(self.programs)} {'has' if one else 'have'} this file open, "
                    "so it can't be read right now.\n\n"
                    f"Close {'it' if one else 'them'} and try again.")
        else:
            text = ("Another program is using this file, so it can't be read right now.\n\n"
                    "If it's the data file of the Outlook profile on this PC, close Outlook "
                    "and try again. If Outlook is already closed, Windows Search may be "
                    "indexing the file in the background: wait a minute or two and try again.")
        super().__init__(f"{text}\n\nFile: {path}")


def is_lock_error(e):
    text = str(e).lower()
    return ("locked a portion" in text or "being used by another process" in text
            or getattr(e, "winerror", None) in (32, 33))


def is_locked(path):
    """Whether another program has locked the start of the file, so that an
    ordinary read of it is refused. Outlook's data files are locked there by
    every program that has them open: Outlook, and also Word, Excel, Teams
    and Windows Search, which all open the mailbox while they run."""
    try:
        with open(path, "rb") as fh:
            fh.read(65536)
        return False
    except PermissionError:
        return True
    except OSError as e:
        return is_lock_error(e)


class _RM_UNIQUE_PROCESS(ctypes.Structure):
    _fields_ = [("dwProcessId", wintypes.DWORD), ("ProcessStartTime", wintypes.FILETIME)]


class _RM_PROCESS_INFO(ctypes.Structure):
    _fields_ = [("Process", _RM_UNIQUE_PROCESS), ("strAppName", wintypes.WCHAR * 256),
                ("strServiceShortName", wintypes.WCHAR * 64), ("ApplicationType", ctypes.c_int),
                ("AppStatus", wintypes.ULONG), ("TSSessionId", wintypes.DWORD),
                ("bRestartable", wintypes.BOOL)]


def programs_using(path):
    """Names of the programs that have this file open, as Windows knows them
    ("Microsoft Word"), asked of the Restart Manager. Empty if there are
    none, or if Windows can't say."""
    try:
        rm = ctypes.WinDLL(mapi.system_dll("rstrtmgr.dll"))
    except OSError:
        return []
    session = wintypes.DWORD()
    if rm.RmStartSession(ctypes.byref(session), 0, ctypes.create_unicode_buffer(33)):
        return []
    try:
        files = (wintypes.LPCWSTR * 1)(os.path.abspath(path))
        if rm.RmRegisterResources(session, 1, files, 0, None, 0, None):
            return []
        needed, reasons = wintypes.UINT(0), wintypes.DWORD(0)
        count = wintypes.UINT(64)
        found = (_RM_PROCESS_INFO * 64)()
        if rm.RmGetList(session, ctypes.byref(needed), ctypes.byref(count), found,
                        ctypes.byref(reasons)):
            return []
        own = os.getpid()
        return sorted({i.strAppName for i in found[:count.value]
                       if i.Process.dwProcessId != own and i.strAppName})
    finally:
        rm.RmEndSession(session)


class MappedFile:
    """A file read through a memory-mapped view, as a file object for libpff.

    Windows applies another program's lock on part of a file to ordinary
    reads, but not to a mapped view. That lets a mailbox be read while Word
    or Excel has it open.
    """

    def __init__(self, path):
        self._fh = open(path, "rb")
        try:
            self._view = mmap.mmap(self._fh.fileno(), 0, access=mmap.ACCESS_READ)
        except (OSError, ValueError):
            self._fh.close()
            raise

    def read(self, size=-1):
        return self._view.read(size)

    def seek(self, offset, whence=os.SEEK_SET):
        return self._view.seek(offset, whence)

    def tell(self):
        return self._view.tell()

    def get_size(self):
        return self._view.size()

    def close(self):
        self._view.close()
        self._fh.close()


def open_file(path):
    f = pypff.file()
    # libpff uses \\?\ long paths on Windows, which reject forward slashes
    path = os.path.abspath(path)
    try:
        f.open(path)
        return f
    except OSError as e:
        if not is_lock_error(e):
            raise
        cause = e
    programs = programs_using(path)
    if _has_outlook(programs):
        # Outlook changes the file as it syncs, so reading under it could
        # give a mixture of old and new: it has to be closed
        raise OstInUseError(path, programs) from cause
    # Anything else (Word, Excel, Teams, Windows Search) only holds the
    # mailbox open, so it's safe to read past its lock
    f = pypff.file()
    try:
        f.open_file_object(MappedFile(path))
    except (OSError, ValueError) as e:
        raise OstInUseError(path, programs) from e
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
                kind=mapi.recipient_kind(_integer(r, mapi.RECIPIENT_TYPE)),
            ))

    attachments = []
    skipped_embedded = 0
    for i in range(msg.number_of_attachments):
        a = msg.get_attachment(i)
        try:
            att_props = _entries(a.record_sets[0])
        except (OSError, IndexError):
            att_props = {}
        hidden = att_props.get(ATTACHMENT_HIDDEN)
        is_hidden = hidden is not None and hidden.data and hidden.data[0]
        if _integer(att_props, ATTACH_METHOD) == ATTACH_EMBEDDED_MSG:
            # Hidden embedded items are a repeating appointment's changed
            # occurrences (recreated from the recurrence data); visible ones
            # are attached emails, which can't be recreated yet
            if not is_hidden:
                skipped_embedded += 1
            continue
        # Other hidden attachments (usually images shown inside an HTML
        # email) are kept, with their content id so the HTML can show them
        size = a.size or 0
        data = a.read_buffer(size) if size else b""
        name = (_optional(lambda: a.long_filename) or _string(att_props, ATTACH_SHORT_FILENAME)
                or f"attachment_{i + 1}")
        cid = _string(att_props, ATTACH_CONTENT_ID)
        attachments.append(Attachment(filename=name, data=data,
                                      content_id=cid.strip().strip("<>") if cid else None))

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
        rtf_body=_optional(lambda: msg.rtf_body),
        transport_headers=_optional(lambda: msg.transport_headers),
        recipients=recipients,
        attachments=attachments,
        skipped_embedded=skipped_embedded,
    )
