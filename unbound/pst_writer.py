"""Write messages into a new PST using Outlook Classic over COM.

Outlook saves a new MailItem as an unsent draft and won't let us clear that
flag afterwards. Instead each message is created as a PostItem (which is
never "unsent"), then given the original sender, times and recipients via
PropertyAccessor and re-classed as IPM.Note so it shows as normal mail.
"""

import contextlib
import functools
import gc
import html
import os
import re
import tempfile
import threading
import time
import unicodedata
import uuid
from datetime import timedelta, timezone
from html.parser import HTMLParser

import pythoncom
import win32com.client
from win32com.client import dynamic

from . import appointment as cal
from . import mapi
from .errors import first_line
from .props import PT_SYSTIME
from .timezones import windows_zone

# OlStoreType: olStoreDefault = 1, olStoreUnicode = 2, olStoreANSI = 3.
# Must be Unicode: ANSI (Outlook 97-2002) PSTs lose characters outside the
# Western code page (emoji, non-Latin text) and stop at 2 GB.
OL_STORE_UNICODE = 2
OL_POST_ITEM = 6  # olPostItem
OL_APPOINTMENT_ITEM = 1  # olAppointmentItem
OL_CONTACT_ITEM = 2  # olContactItem
OL_BY_VALUE = 1  # olByValue
OL_FORMAT_PLAIN = 1  # olFormatPlain
OL_FOLDER_CALENDAR = 9  # olFolderCalendar
OL_FOLDER_CONTACTS = 10  # olFolderContacts
OL_FOLDER_DELETED_ITEMS = 3  # olFolderDeletedItems
ATTACH_CONTENT_ID = 0x3712  # PR_ATTACH_CONTENT_ID

# Properties that belong to the source file or that Outlook computes itself;
# copying them is pointless or harmful. Ids only (any type).
SKIP_PROP_IDS = {
    0x001A,  # message class: set through the object model
    0x0037, 0x003D, 0x0E1D, 0x0070,  # subject and its parts: set via Subject
    0x1000, 0x1009, 0x1013, 0x1016, 0x1035,  # bodies, set via Body/RTFBody
    0x0E04, 0x0E03, 0x0E02,  # To/Cc/Bcc display lines: computed from recipients
    0x0E06, 0x0E07, 0x0E08, 0x0E17, 0x0E1B, 0x0E1F, 0x0E21,  # flags, size, status
    0x0FF4, 0x0FF6, 0x0FF8, 0x0FF9, 0x0FFE, 0x0FFF,  # access, instance/record keys, entry id
    0x300B,  # search key
    0x3007, 0x3008,  # created / modified times (set by the store)
    0x65E0, 0x65E1, 0x65E2, 0x65E3,  # sync source/parent keys, change key, change list
    0x0E09, 0x0E0A, 0x0E0B,  # parent entry id, sent-mail entry id
    0x3FDE, 0x3FF1, 0x3FFD,  # code pages / locale: Outlook sets its own
    0x7FFA, 0x7FFB, 0x7FFC, 0x7FFD,  # attachment hiding flags
}
# Sync and per-file bookkeeping kept by Exchange in the store, not the item
SKIP_PROP_RANGES = [(0x6600, 0x67FF), (0x0E30, 0x0E3F)]


def should_copy(p):
    if p.named is not None:
        return True
    if p.tag_id in SKIP_PROP_IDS:
        return False
    if any(lo <= p.tag_id <= hi for lo, hi in SKIP_PROP_RANGES):
        return False
    return p.tag_id < 0x8000


def _prop_label(p):
    if p.named:
        n = p.named
        return f"{n.guid}/{n.lid:04X}" if n.lid is not None else f"{n.guid}/{n.name}"
    return f"0x{p.tag_id:04X}{p.prop_type:04X}"


def com_date(dt):
    """Hand a date to Outlook exactly as given.

    pywin32 treats a naive datetime as local time and converts it to UTC on
    the way into COM, but passes one marked as UTC through unchanged. So mark
    it UTC: then UTC values (StartUTC, raw properties) and local wall-clock
    values (PatternStartDate, Start) both arrive as written.
    """
    if dt is None:
        return None
    return dt.replace(tzinfo=timezone.utc)


def from_com_date(value):
    """The wall-clock value of a date read from Outlook."""
    return value.replace(tzinfo=None) if value is not None else None


# Names Windows 10 and earlier treat as devices, whatever the extension:
# a file called "nul.txt" would be written to nowhere
_DEVICE_NAMES = ({"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"}
                 | {f"{port}{n}" for port in ("COM", "LPT") for n in range(1, 10)})


def _safe_filename(name):
    """A name safe to write in a temporary folder. Outlook takes the
    attachment's file name from it, so the extension is kept."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    stem, ext = os.path.splitext(name)
    if len(name) > 150 and stem and len(ext) <= 20:
        name = stem[:150 - len(ext)].rstrip(" .") + ext
    name = name[:150] or "attachment"
    if name.split(".")[0].strip().upper() in _DEVICE_NAMES:
        name = "_" + name
    return name


class _TextFromHtml(HTMLParser):
    """Collects the readable text of an HTML document."""

    BREAKS = {"br", "p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "table"}
    HIDDEN = {"head", "style", "script", "title"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.hidden = 0

    def handle_starttag(self, tag, attrs):
        if tag in self.HIDDEN:
            self.hidden += 1
        elif tag in self.BREAKS:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in self.HIDDEN:
            self.hidden = max(0, self.hidden - 1)
        elif tag in self.BREAKS:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.hidden:
            self.parts.append(data)


def html_to_text(text):
    """The readable text of an HTML body, for an item that can only take a
    plain one. Line breaks are kept; runs of blank lines are squeezed."""
    parser = _TextFromHtml()
    try:
        parser.feed(text or "")
        parser.close()
        raw = "".join(parser.parts)
    except Exception:
        raw = html.unescape(re.sub(r"<[^>]*>", " ", text or ""))
    lines = [" ".join(line.split()) for line in raw.replace("\xa0", " ").splitlines()]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


def clean_rtf(rtf):
    """RTF as Outlook will accept it. RTF decoded from an OST's compressed
    form ends with padding (NUL bytes) after the closing brace, and Outlook
    rejects that as "The parameter is incorrect"."""
    if isinstance(rtf, str):
        rtf = rtf.encode("cp1252", errors="replace")
    return (rtf or b"").rstrip(b"\x00\r\n ")


# RTF groups that hold no text: fonts, colours, styles, document info and
# "\*" destinations such as \*\generator
_RTF_SKIPPED_GROUPS = re.compile(r"\{\\(?:\*|(?:fonttbl|colortbl|stylesheet|info|pict)\b)")


def _drop_rtf_groups(text):
    out, i = [], 0
    while True:
        m = _RTF_SKIPPED_GROUPS.search(text, i)
        if not m:
            out.append(text[i:])
            return "".join(out)
        out.append(text[i:m.start()])
        depth, j = 0, m.start()
        while j < len(text):
            ch = text[j]
            if ch == "\\":
                j += 2  # an escaped brace doesn't count
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    break
            j += 1
        i = j + 1


def rtf_to_text(rtf):
    """A rough plain-text version of an RTF body, for when Outlook won't take
    the RTF itself and there's no other body: enough to keep the words."""
    if isinstance(rtf, bytes):
        rtf = rtf.decode("cp1252", errors="replace")
    text = _drop_rtf_groups((rtf or "").rstrip("\x00"))
    text = re.sub(r"\\(par|line)\b ?", "\n", text)
    text = re.sub(r"\\'([0-9a-fA-F]{2})", lambda m: bytes([int(m.group(1), 16)]).decode("cp1252", "replace"), text)
    text = re.sub(r"\\u(-?\d+)\??", lambda m: chr(int(m.group(1)) % 0x10000), text)
    text = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", text)  # other control words
    text = text.replace("\\{", "{").replace("\\}", "}").replace("\\\\", "\\")
    text = text.replace("{", "").replace("}", "")
    lines = [" ".join(line.split()) for line in text.splitlines()]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


def _folder_key(name):
    """How Outlook compares folder names: ignoring case and how accents are
    encoded. "Archive" and "archive" are the same folder to it."""
    return unicodedata.normalize("NFC", name or "").casefold()


def _is_smtp(address):
    return bool(address) and "@" in address and not address.startswith("/")


def recipient_text(name, email):
    """(text for Recipients.Add, whether Outlook should resolve it).

    With an internet address Outlook is given '"Name" <address>' and told to
    resolve it, which is what makes it store the address. Without one (only
    a name, or an Exchange directory address) just the name is stored:
    resolving would look it up in *this* PC's address book and could quietly
    put in a different person who happens to match.
    """
    # A quote inside the name would end the quoted part early
    name = re.sub(r"[\x00-\x1f\x7f]+", " ", name or "").replace('"', "'").strip()
    email = (email or "").strip()
    if _is_smtp(email):
        return (f'"{name}" <{email}>' if name else email), True
    return (name or email or None), False


def _stores_newest_first(ns):
    """Each store in the profile with its file path, newest first, stepping
    over any that Outlook can't open right now.

    A store that can't be reached (a shared mailbox, an account that can't
    sign in, an archive on a drive that isn't connected) can block for
    minutes and then make Outlook's own `for store in Stores` enumeration
    fail. Going by index, one bad store costs only itself, and the PST just
    attached, which is the newest, is found first.
    """
    stores = ns.Stores
    for i in range(stores.Count, 0, -1):
        try:
            store = stores.Item(i)
            path = store.FilePath
        except Exception:
            continue
        yield store, path


class _Dispatch(dynamic.CDispatch):
    """pywin32's late-bound COM object, remembering what it learns about a
    type of object (its properties, and how each method is called) for every
    object of that type.

    pywin32 keeps that per object, so each new email, recipient and
    attachment asked Outlook all over again: some 80 extra round trips an
    email, about half the time a conversion took. What's sent to Outlook to
    do the work is unchanged.
    """

    # Objects reached through this one (an item's recipients, say) share too.
    # pywin32 wraps them in two places: here, for a property's value...
    def _wrap_dispatch_(self, ob, userName=None, returnCLSID=None):
        return _wrap(self.__dict__["_types_"], ob, userName)

    # ...and, for what a method returns, with the "Dispatch" that the code
    # it generates for the method finds in its globals
    def _make_method_(self, name):
        method = super()._make_method_(name)
        if method is not None:
            method.__func__.__globals__["Dispatch"] = functools.partial(
                _wrap, self.__dict__["_types_"])
        return method


def _wrap(types, ob, userName=None, *_):
    """A COM object from Outlook as a _Dispatch. `types` is what's known so
    far: {interface id: (type description, methods built, type information)}.
    It belongs to one PstWriter, so to one conversion and one thread."""
    if not isinstance(ob, dynamic.PyIDispatchType):
        return win32com.client.Dispatch(ob, userName)
    try:
        typeinfo = ob.GetTypeInfo()
        key = str(typeinfo.GetTypeAttr()[0])
    except (pythoncom.com_error, AttributeError):
        return dynamic.Dispatch(ob, userName)  # it won't say what it is
    known = types.get(key)
    if known is None:
        obj = dynamic.Dispatch(ob, userName, createClass=_Dispatch, typeinfo=typeinfo)
        types[key] = (obj._olerepr_, obj._builtMethods_, obj._lazydata_)
    else:
        obj = _Dispatch(ob, known[0], userName, lazydata=known[2])
        obj.__dict__["_builtMethods_"] = known[1]
    obj.__dict__["_types_"] = types
    return obj


def _outlook():
    """Outlook's Application object. Starts Outlook if it isn't running."""
    app = win32com.client.Dispatch("Outlook.Application")
    if isinstance(app, dynamic.CDispatch):
        # (Not when this PC has pywin32's generated classes for Outlook:
        # those already know their types.)
        try:
            return _wrap({}, app._oleobj_, app._username_)
        except Exception:
            pass  # slower, but it works: pywin32's own object
    return app


def _set_if_changed(obj, name, value):
    try:
        if getattr(obj, name) == value:
            return
    except Exception:
        pass  # can't be read: set it
    setattr(obj, name, value)


def _detach(ns, paths, raise_errors=False):
    """Detach these PST files from the Outlook profile where they're
    attached (the files stay on disk). Returns the paths that were.
    Errors are stepped over unless raise_errors is set."""
    want = {os.path.normcase(os.path.abspath(p)) for p in paths}
    removed = []
    if not want:
        return removed
    try:
        for store, path in list(_stores_newest_first(ns)):
            if not (path and os.path.normcase(os.path.abspath(path)) in want):
                continue
            try:
                ns.RemoveStore(store.GetRootFolder())
                removed.append(path)
            except Exception:
                if raise_errors:
                    raise
                # a store Outlook won't let go of
    except Exception:
        if raise_errors:
            raise
        # Outlook couldn't even say how many stores there are
    return removed


def detach_stores(paths, timeout=None):
    """Detach PST files that a conversion which was interrupted left attached
    to Outlook. Starts Outlook if it isn't running. Returns the paths that
    were attached. With a timeout (seconds), DetachTimeout is raised if
    Outlook hasn't answered by then."""
    if timeout is not None:
        return _detach_within(paths, timeout, raise_errors=False)
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    try:
        return _detach(ns, paths)
    finally:
        del ns
        gc.collect()


class PstWriter:
    def __init__(self, pst_path, display_name=None, temp_dir=None, detach=()):
        self.pst_path = os.path.abspath(pst_path)
        self.temp_dir = temp_dir  # where attachments are staged; None = the system's
        self.app = _outlook()
        self.ns = self.app.GetNamespace("MAPI")
        if detach:
            _detach(self.ns, detach)
        self.property_errors = {}  # property name -> error text, for the log
        self.warnings = []  # things copied only partly, for the log
        self.ns.AddStoreEx(self.pst_path, OL_STORE_UNICODE)
        try:
            self.store = self._find_store()
            self.root = self.store.GetRootFolder()
            if display_name:
                self.root.Name = display_name
        except BaseException:
            # The PST is attached but can't be used: don't leave it in the
            # user's profile
            _detach(self.ns, [self.pst_path])
            raise

    def _find_store(self):
        want = os.path.normcase(self.pst_path)
        for store, path in _stores_newest_first(self.ns):
            if path and os.path.normcase(os.path.abspath(path)) == want:
                return store
        raise RuntimeError(f"Outlook did not attach the PST: {self.pst_path}")

    def get_folder(self, parts, folder_types=None):
        """Return the folder at root/parts[0]/parts[1]/..., creating as needed.

        folder_types gives each level's type when it has to be created
        (OL_FOLDER_CALENDAR, OL_FOLDER_CONTACTS, or None for mail): a list
        with one entry per level, or a single value for the last level only.
        """
        if not isinstance(folder_types, (list, tuple)):
            folder_types = [None] * (len(parts) - 1) + [folder_types]
        # Folders already found or made, by path: listing a folder's
        # subfolders through Outlook again for each new one adds up in a
        # mailbox with hundreds of folders
        known = self.__dict__.setdefault("_folders", {})
        folder, path = self.root, ()
        for name, ftype in zip(parts, folder_types):
            # Matched the way Outlook matches: it won't create "archive" next
            # to "Archive" (IMAP accounts can have both), so they share one
            key = _folder_key(name)
            path += (key,)
            existing = known.get(path)
            if existing is None:
                for sub in folder.Folders:
                    if _folder_key(sub.Name) == key:
                        existing = sub
                        break
            if existing is None:
                try:
                    existing = self._add_folder(folder, name, ftype)
                except Exception:
                    # A name Outlook still treats as taken: keep the items apart
                    existing = self._add_folder(folder, f"{name} (2)", ftype)
            folder = known[path] = existing
        return folder

    def _items(self, folder):
        """A folder's item collection, fetched once while items keep going
        into the same folder."""
        last = self.__dict__.get("_last_items")
        if last is None or last[0] is not folder:
            last = self.__dict__["_last_items"] = (folder, folder.Items)
        return last[1]

    @staticmethod
    def _add_folder(parent, name, ftype):
        return parent.Folders.Add(name, ftype) if ftype else parent.Folders.Add(name)

    def _staging(self, attachments=True):
        """A temporary folder for one item's attachments, inside the run's
        own folder when there is one (short name there: paths have a limit).
        None for an item without attachments, which needs no folder."""
        if not attachments:
            return contextlib.nullcontext()
        return tempfile.TemporaryDirectory(
            prefix="unbound_" if self.temp_dir is None else "a", dir=self.temp_dir)

    def _item(self, entry_id):
        """A saved item, opened afresh. The store's id is asked for once."""
        store_id = self.__dict__.get("_store_id")
        if store_id is None:
            store_id = self.__dict__["_store_id"] = self.store.StoreID
        return self.ns.GetItemFromID(entry_id, store_id)

    def _add_recipients(self, item, recipients, label):
        if not recipients:
            return
        added = item.Recipients
        for r in recipients:
            text, resolve = recipient_text(r.name, r.email)
            if not text:
                continue
            try:
                rec = added.Add(text)
            except Exception as e:
                self.property_errors[label] = first_line(e, 200)
                continue
            try:
                rec.Type = r.kind
            except Exception as e:
                # Kept as To; the address matters more than the type
                self.property_errors[f"{label} type"] = first_line(e, 200)
            if resolve:
                # Without this Outlook stores the raw text as the display
                # name and leaves the email address empty
                try:
                    rec.Resolve()
                except Exception as e:
                    self.property_errors[label] = first_line(e, 200)

    def _add_attachments(self, item, attachments, tmp):
        for i, att in enumerate(attachments):
            # unique subfolder per attachment so duplicate names don't clash
            d = os.path.join(tmp, str(i))
            os.mkdir(d)
            path = os.path.join(d, _safe_filename(att.filename))
            with open(path, "wb") as fh:
                fh.write(att.data)
            if att.filename == "ContactPicture.jpg" and hasattr(item, "AddPicture"):
                item.AddPicture(path)  # shows as the contact's photo
                continue
            added = item.Attachments.Add(path, OL_BY_VALUE, 1, att.filename)
            if att.content_id:
                # The HTML body shows this image through "cid:<content id>";
                # without the id the image appears broken. (Outlook decides
                # itself whether to list it as an attachment: it resets
                # PR_ATTACHMENT_HIDDEN and refuses PR_ATTACH_FLAGS on save.)
                try:
                    added.PropertyAccessor.SetProperty(
                        mapi.proptag_url(ATTACH_CONTENT_ID, mapi.PT_UNICODE), att.content_id)
                except Exception as e:
                    self.property_errors["inline image id"] = first_line(e, 200)

    def _copy_props(self, item, props, skip=None):
        """Copy raw properties onto the item; returns (set, failed) counts."""
        entries = []
        for p in props:
            if not should_copy(p) or (skip and skip(p)):
                continue
            value = com_date(p.value) if p.prop_type == PT_SYSTIME else p.value
            entries.append((p.url(), value, _prop_label(p)))
        # One by one: see _set_properties for why not in one call
        refused = self._set_properties(item.PropertyAccessor, entries, together=False)
        return len(entries) - len(refused), len(refused)

    def _set_properties(self, accessor, entries, together=False):
        """Set raw properties, given as (schema name, value, label for the
        log). Returns (label, Outlook's error) for each one it refused, and
        notes them for the log.

        `together` sends them in one call (SetProperties), which is much
        quicker: 40 properties took 25 ms one by one and 1 ms together.
        Outlook answers with a result per property and sets the rest when it
        refuses one. Only for a known, fixed set of text and date properties,
        though. With an item's copied properties it isn't safe: one such
        call on a contact (text, numbers, yes/no values and binary together)
        crashed Outlook, and on appointments it let through properties that
        Outlook refuses when asked singly.
        """
        refused = []

        def one_by_one(some):
            for name, value, label in some:
                try:
                    accessor.SetProperty(name, value)
                except Exception as e:
                    self.property_errors[label] = first_line(e, 200)
                    refused.append((label, self.property_errors[label]))

        if not together or len(entries) < 2:
            one_by_one(entries)
            return refused
        try:
            results = accessor.SetProperties([e[0] for e in entries], [e[1] for e in entries])
        except Exception:
            one_by_one(entries)  # the call as a whole wasn't taken
            return refused
        if not isinstance(results, (list, tuple)) or len(results) != len(entries):
            return refused  # an empty answer: nothing was refused
        again = []
        for entry, result in zip(entries, results):
            if result is None:
                continue
            if entry[2] in self.property_errors:
                # The batch only gives an error number. The text for this
                # property is already known from an earlier item
                refused.append((entry[2], self.property_errors[entry[2]]))
            else:
                again.append(entry)  # ask again by itself, for the text
        one_by_one(again)
        return refused

    # ---------- appointments ----------

    def _tz(self, blob):
        name = cal.tz_key_name(blob)
        if not name:
            return None
        # A calendar uses a handful of zones, so each is looked up once
        zones = self.__dict__.setdefault("_zones", {})
        if name not in zones:
            zones[name] = self._find_zone(name)
        if zones[name] is None:
            self.warnings.append(f"time zone '{name}' not known to this PC")
        return zones[name]

    def _find_zone(self, name):
        # Outlook knows Windows zone names; phones and Google Calendar write
        # IANA ones ("America/New_York"), which Windows can translate
        for candidate in (lambda: name, lambda: windows_zone(name)):
            candidate = candidate()
            if not candidate:
                continue
            try:
                zone = self.app.TimeZones.Item(candidate)
            except Exception:
                continue
            # Outlook answers with nothing, not an error, for a name it
            # doesn't know. Taking that as an answer meant the Windows name
            # was never tried, and such appointments lost their time zone
            if zone is not None:
                return zone
        return None

    def _to_pc_local(self, dt, tz):
        """A time in the appointment's own zone -> this PC's local time."""
        if tz is None:
            return dt
        try:
            pc = self.__dict__.get("_pc_zone")
            if pc is None:
                current = self.app.TimeZones.CurrentTimeZone
                pc = self.__dict__["_pc_zone"] = (current, current.ID)
            if tz.ID == pc[1]:
                return dt
            return from_com_date(self.app.TimeZones.ConvertTime(com_date(dt), tz, pc[0]))
        except Exception:
            return dt

    def _set_times(self, item, a):
        """Set an appointment's time zones and times; returns its start zone.

        A meeting received from someone else (state flags, named property
        8217, with the "received" bit) still ends up in this PC's zone:
        Outlook resets its zone at every save, whatever is set here or
        afterwards, and refuses the zone properties written raw. Its times
        are right; only the zone it's shown in differs from the source."""
        # A repeating series' times are defined in its recurrence time zone;
        # the start "display" zone can differ (or be wrong), so prefer it
        noted = len(self.warnings)
        if a.get(cal.LID_RECUR):
            tz_start = self._tz(a.get(cal.LID_TZ_RECUR) or a.get(cal.LID_TZ_START_DISPLAY))
        else:
            tz_start = self._tz(a.get(cal.LID_TZ_START_DISPLAY) or a.get(cal.LID_TZ_RECUR))
        tz_end = self._tz(a.get(cal.LID_TZ_END_DISPLAY)) or tz_start
        # Start and end are usually in the same zone: say it once
        self.warnings[noted:] = dict.fromkeys(self.warnings[noted:])
        if tz_start is not None:
            item.StartTimeZone = tz_start
            item.EndTimeZone = tz_end
        if a.get(cal.LID_ALL_DAY):
            item.AllDayEvent = True
        if a.get(cal.LID_START) is not None:
            item.StartUTC = com_date(a[cal.LID_START])
        if a.get(cal.LID_END) is not None:
            item.EndUTC = com_date(a[cal.LID_END])
        return tz_start

    def _set_recurrence(self, item, r):
        ol_type = r.outlook_type()
        if ol_type is None:
            raise ValueError("this kind of repeat (e.g. Hijri calendar) isn't supported")
        rp = item.GetRecurrencePattern()
        rp.RecurrenceType = ol_type
        # Each part is set only if Outlook hasn't already worked it out from
        # the start: for a yearly series on 29 February it refuses even to
        # be told the day, month or interval it already has
        if ol_type in (cal.OL_WEEKLY, cal.OL_MONTH_NTH, cal.OL_YEAR_NTH):
            _set_if_changed(rp, "DayOfWeekMask", r.day_of_week_mask)
        if ol_type in (cal.OL_MONTH_NTH, cal.OL_YEAR_NTH):
            _set_if_changed(rp, "Instance", r.instance)
        if ol_type in (cal.OL_MONTHLY, cal.OL_YEARLY):
            _set_if_changed(rp, "DayOfMonth",
                            31 if r.pattern_type == cal.PT_MONTH_END else r.day_of_month)
        if ol_type in (cal.OL_YEARLY, cal.OL_YEAR_NTH):
            _set_if_changed(rp, "MonthOfYear", r.start_date.month)
        if ol_type in (cal.OL_YEARLY, cal.OL_YEAR_NTH):
            # Outlook counts a yearly interval in months (12 = every year,
            # 24 = every two) and refuses 1 or 2 as "not valid"
            _set_if_changed(rp, "Interval", 12 * r.interval())
        else:
            _set_if_changed(rp, "Interval", r.interval())
        rp.PatternStartDate = com_date(r.start_date)
        if not item.AllDayEvent:
            rp.StartTime = com_date(r.start_date + timedelta(minutes=r.start_offset))
            rp.Duration = r.end_offset - r.start_offset
        if r.end_type == cal.END_AFTER_DATE:
            rp.PatternEndDate = com_date(r.end_date)
        elif r.end_type == cal.END_AFTER_N:
            rp.Occurrences = r.occurrences
        else:
            rp.NoEndDate = True

    def _apply_exceptions(self, item, r, tz):
        rp = item.GetRecurrencePattern()
        changed_days = {x.original_start.date() for x in r.exceptions}
        for day in r.deleted:
            if day.date() in changed_days:
                continue
            original = self._to_pc_local(day + timedelta(minutes=r.start_offset), tz)
            try:
                rp.GetOccurrence(com_date(original)).Delete()
            except Exception:
                self.warnings.append(f"couldn't remove a cancelled occurrence of '{item.Subject}'")
        for x in r.exceptions:
            try:
                occ = rp.GetOccurrence(com_date(self._to_pc_local(x.original_start, tz)))
                occ.Start = com_date(self._to_pc_local(x.start, tz))
                occ.End = com_date(self._to_pc_local(x.end, tz))
                if x.subject is not None:
                    occ.Subject = x.subject
                if x.location is not None:
                    occ.Location = x.location
                if x.busy_status is not None:
                    occ.BusyStatus = x.busy_status
                if x.reminder_set is not None:
                    occ.ReminderSet = x.reminder_set
                if x.reminder_minutes is not None:
                    occ.ReminderMinutesBeforeStart = x.reminder_minutes
                occ.Save()
            except Exception:
                self.warnings.append(f"couldn't apply a changed occurrence of '{item.Subject}'")

    def add_item(self, folder, kind, m, props):
        """Recreate an appointment or contact from its raw properties.

        kind is OL_APPOINTMENT_ITEM or OL_CONTACT_ITEM. m is the Message from
        ost_reader (subject, bodies, recipients, attachments); props are all
        its properties from props.read_props.
        """
        is_appt = kind == OL_APPOINTMENT_ITEM
        skip = cal.is_managed if is_appt else None
        a = cal.appointment_named(props) if is_appt else {}
        recurrence = None
        if is_appt and a.get(cal.LID_RECUR):
            recurrence = cal.parse_recurrence(a[cal.LID_RECUR])  # ValueError -> item fails

        item = self._items(folder).Add(kind)
        item.MessageClass = m.message_class
        item.Subject = m.subject
        self._set_item_body(item, m)
        if kind == OL_APPOINTMENT_ITEM:
            # attendees; saving never sends anything
            self._add_recipients(item, m.recipients, "attendees")
        with self._staging(m.attachments) as tmp:
            self._add_attachments(item, m.attachments, tmp)
            self._copy_props(item, props, skip)
            tz = self._set_times(item, a) if is_appt else None
            item.Save()
        entry_id = item.EntryID
        try:
            if recurrence is not None:
                # (Outlook resets the series' start "display" zone to this PC's
                # here whatever we set; the recurrence keeps the original zone,
                # which is what defines the series' times.)
                self._set_recurrence(item, recurrence)
                item.Save()
                self._apply_exceptions(item, recurrence, tz)
            # Outlook recomputes some fields on save, so apply the raw properties
            # once more on a fresh copy of the item and save again
            del item
            item = self._item(entry_id)
            self._copy_props(item, props, skip)
            item.Save()
        except BaseException:
            # Reported as failed, so don't leave a half-made copy behind
            # (e.g. a series that couldn't repeat saved as a single meeting)
            self._delete_saved(entry_id)
            raise
        return item

    def _delete_saved(self, entry_id):
        """Remove an item for good. Delete() only moves it to the PST's
        Deleted Items (where a failed birthday series was found, saved
        without its pattern), and Outlook doesn't say where it went. So it's
        marked first, found there by the mark and deleted again, which
        removes it. (Moving it to Deleted Items and deleting the moved copy
        didn't: Outlook put that in Drafts.)"""
        mark = f"unbound-remove-{uuid.uuid4().hex}"
        try:
            item = self._item(entry_id)
            item.Subject = mark
            item.Save()
            item.Delete()
            del item
        except Exception:
            return  # best effort; the failure itself is what gets reported
        try:
            bin_ = self.store.GetDefaultFolder(OL_FOLDER_DELETED_ITEMS)
            found = bin_.Items.Find(f"[Subject] = '{mark}'")
            if found is not None:
                found.Delete()
        except Exception:
            pass  # left in Deleted Items, under the mark

    def _set_item_body(self, item, m):
        """An appointment's or contact's notes. Outlook writes appointment
        notes as HTML only, which these items can't take directly, so the
        HTML's text is kept rather than losing the notes."""
        if m.rtf_body:
            try:
                item.RTFBody = clean_rtf(m.rtf_body)
                return
            except Exception:
                self.warnings.append("the notes' formatting couldn't be kept; saved as plain text")
        text = m.plain_body or (html_to_text(m.html_body) if m.html_body else "")
        if not text and m.rtf_body:
            text = rtf_to_text(m.rtf_body)
        if text:
            item.Body = text

    def _set_sender_and_times(self, item, m):
        S, T = mapi.PT_UNICODE, mapi.PT_SYSTIME
        address = m.sender_email
        # An Exchange sender whose internet address isn't in the source has
        # only a directory address ("/o=.../cn=..."). That's type EX, not an
        # SMTP address, so it mustn't be stored as one.
        addr_type = ("EX" if address.startswith("/") else "SMTP") if address else None
        smtp = address if _is_smtp(address) else None
        wanted = (
            ("sender name", mapi.SENDER_NAME, S, m.sender_name),
            ("sent-representing name", mapi.SENT_REPRESENTING_NAME, S, m.sender_name),
            ("sender address type", mapi.SENDER_ADDRTYPE, S, addr_type),
            ("sent-representing address type", mapi.SENT_REPRESENTING_ADDRTYPE, S, addr_type),
            ("sender email", mapi.SENDER_EMAIL_ADDRESS, S, address),
            ("sent-representing email", mapi.SENT_REPRESENTING_EMAIL_ADDRESS, S, address),
            ("sender SMTP", mapi.SENDER_SMTP_ADDRESS, S, smtp),
            ("sent-representing SMTP", mapi.SENT_REPRESENTING_SMTP_ADDRESS, S, smtp),
            ("delivery time", mapi.MESSAGE_DELIVERY_TIME, T, com_date(m.delivery_time)),
            ("submit time", mapi.CLIENT_SUBMIT_TIME, T, com_date(m.submit_time)),
        )
        entries = [(mapi.proptag_url(prop_id, prop_type), value, label)
                   for label, prop_id, prop_type, value in wanted
                   if value is not None and value != ""]
        for label, error in self._set_properties(item.PropertyAccessor, entries, together=True):
            # Sender and dates matter to the reader, so flag the item itself
            self.warnings.append(f"couldn't set the {label} ({error})")

    def add_message(self, folder, m):
        # 1. Create as a PostItem so it is saved as delivered mail, not a draft
        post = self._items(folder).Add(OL_POST_ITEM)
        post.Subject = m.subject
        if m.html_body:
            post.HTMLBody = m.html_body
        elif m.rtf_body:
            # Rich-text mail (common between Outlook users) keeps its formatting
            try:
                post.RTFBody = clean_rtf(m.rtf_body)
            except Exception:
                # Outlook can refuse it; keep the words rather than save an
                # empty email, and say so in the log
                self.warnings.append("the email's rich-text formatting couldn't be kept; "
                                     "saved as plain text")
                post.BodyFormat = OL_FORMAT_PLAIN
                post.Body = m.plain_body or rtf_to_text(m.rtf_body)
        elif m.plain_body:
            # Keep plain-text mail plain; otherwise Outlook stores it as HTML
            post.BodyFormat = OL_FORMAT_PLAIN
            post.Body = m.plain_body

        with self._staging(m.attachments) as tmp:
            self._add_attachments(post, m.attachments, tmp)
            post.MessageClass = m.message_class
            post.Save()
        entry_id = post.EntryID
        del post

        # 2. Reopen it: now it's a MailItem, so it takes real recipients.
        #    Times are set here because the first save overwrites them.
        try:
            item = self._item(entry_id)
            self._add_recipients(item, m.recipients, "recipients")
            self._set_sender_and_times(item, m)
            item.UnRead = not m.is_read
            item.Save()
        except BaseException:
            # Reported as failed, so don't leave a copy without sender,
            # recipients or dates behind
            self._delete_saved(entry_id)
            raise
        return item

    def close(self, quit_outlook=False, timeout=None):
        """Detach the PST from the Outlook profile (the file stays on disk).

        With a timeout (seconds), Outlook is asked from a thread of its own,
        and DetachError is raised if it hasn't answered by then or refuses.
        An Outlook started in the background was seen to stop answering
        altogether at this point, while Microsoft 365 was down for its
        Exchange account, which left the window at 100% for ever. The request stays queued, so Outlook may still detach the PST
        once it recovers.
        """
        if timeout is None:
            try:
                self.ns.RemoveStore(self.root)
            finally:
                if quit_outlook:
                    self.app.Quit()
                self._drop_references()
            return
        try:
            _detach_within([self.pst_path], timeout)
        except DetachTimeout:
            # Releasing these would wait on the same unanswering Outlook
            _stuck_references.append((self.root, self.store, self.ns, self.app))
            self.root = self.store = self.ns = self.app = None
            raise
        except BaseException:
            self._drop_references()
            raise
        if quit_outlook:
            self.app.Quit()
        self._drop_references()

    def _drop_references(self):
        # Drop every COM reference; Outlook keeps the PST open while any are
        # alive.
        self.root = self.store = self.ns = self.app = None
        gc.collect()


DETACH_TIMEOUT = 180  # seconds Outlook gets to detach the PST at the end
_stuck_references = []  # COM references to an Outlook that stopped answering


class DetachError(Exception):
    """The PST couldn't be detached from the Outlook profile."""


class DetachTimeout(DetachError):
    pass


def _detach_within(paths, timeout, raise_errors=True):
    """Detach these PSTs from the profile through a connection to Outlook
    made on a thread of its own, waiting at most timeout seconds. Returns
    the paths that were attached."""
    outcome = {}

    def work():
        pythoncom.CoInitialize()
        try:
            ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            # A PST that isn't attached any more (Outlook let go of it) is fine
            outcome["removed"] = _detach(ns, paths, raise_errors=raise_errors)
            del ns
        except Exception as e:
            outcome["error"] = e
        finally:
            gc.collect()
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=work, daemon=True, name="unbound-detach")
    thread.start()
    # Keep handling this thread's COM messages while waiting, as a thread
    # holding COM objects should, in case Outlook calls back into it
    deadline = time.monotonic() + timeout
    while thread.is_alive() and time.monotonic() < deadline:
        pythoncom.PumpWaitingMessages()
        thread.join(0.05)
    if thread.is_alive():
        raise DetachTimeout(f"Outlook stopped responding while the PST was being detached "
                            f"from it (waited {_minutes(timeout)})")
    if "error" in outcome:
        raise DetachError(f"Outlook couldn't detach the PST: {first_line(outcome['error'], 200)}")
    return outcome.get("removed", [])


def _minutes(seconds):
    if seconds < 120:
        return f"{seconds:g} seconds"
    return f"{seconds / 60:g} minutes"
