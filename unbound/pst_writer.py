"""Write messages into a new PST using Outlook Classic over COM.

Outlook saves a new MailItem as an unsent draft and won't let us clear that
flag afterwards. Instead each message is created as a PostItem (which is
never "unsent"), then given the original sender, times and recipients via
PropertyAccessor and re-classed as IPM.Note so it shows as normal mail.
"""

import os
import re
import tempfile
from datetime import timedelta, timezone

import win32com.client

from . import appointment as cal
from . import mapi
from .props import PT_SYSTIME

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


def _safe_filename(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name[:150] or "attachment"


class PstWriter:
    def __init__(self, pst_path, display_name=None):
        self.pst_path = os.path.abspath(pst_path)
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self.ns.AddStoreEx(self.pst_path, OL_STORE_UNICODE)
        self.store = self._find_store()
        self.root = self.store.GetRootFolder()
        if display_name:
            self.root.Name = display_name
        self.property_errors = {}  # property name -> error text, for the log
        self.warnings = []  # things copied only partly, for the log

    def _find_store(self):
        want = os.path.normcase(self.pst_path)
        for store in self.ns.Stores:
            try:
                path = store.FilePath
            except Exception:
                continue
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
        folder = self.root
        for name, ftype in zip(parts, folder_types):
            existing = None
            for sub in folder.Folders:
                if sub.Name == name:
                    existing = sub
                    break
            if existing is None:
                existing = folder.Folders.Add(name, ftype) if ftype else folder.Folders.Add(name)
            folder = existing
        return folder

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
                    self.property_errors["inline image id"] = str(e).splitlines()[0][:200]

    def _copy_props(self, item, props, skip=None):
        """Copy raw properties onto the item; returns (set, failed) counts."""
        pa = item.PropertyAccessor
        ok = failed = 0
        for p in props:
            if not should_copy(p) or (skip and skip(p)):
                continue
            value = p.value
            if p.prop_type == PT_SYSTIME:
                value = com_date(value)
            try:
                pa.SetProperty(p.url(), value)
                ok += 1
            except Exception as e:
                failed += 1
                self.property_errors[_prop_label(p)] = str(e).splitlines()[0][:200]
        return ok, failed

    # ---------- appointments ----------

    def _tz(self, blob):
        name = cal.tz_key_name(blob)
        if not name:
            return None
        try:
            return self.app.TimeZones.Item(name)
        except Exception:
            self.warnings.append(f"time zone '{name}' not known to this PC")
            return None

    def _to_pc_local(self, dt, tz):
        """A time in the appointment's own zone -> this PC's local time."""
        if tz is None:
            return dt
        try:
            current = self.app.TimeZones.CurrentTimeZone
            if tz.ID == current.ID:
                return dt
            return from_com_date(self.app.TimeZones.ConvertTime(com_date(dt), tz, current))
        except Exception:
            return dt

    def _set_times(self, item, a):
        # A repeating series' times are defined in its recurrence time zone;
        # the start "display" zone can differ (or be wrong), so prefer it
        if a.get(cal.LID_RECUR):
            tz_start = self._tz(a.get(cal.LID_TZ_RECUR) or a.get(cal.LID_TZ_START_DISPLAY))
        else:
            tz_start = self._tz(a.get(cal.LID_TZ_START_DISPLAY) or a.get(cal.LID_TZ_RECUR))
        tz_end = self._tz(a.get(cal.LID_TZ_END_DISPLAY)) or tz_start
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
        if ol_type in (cal.OL_WEEKLY, cal.OL_MONTH_NTH, cal.OL_YEAR_NTH):
            rp.DayOfWeekMask = r.day_of_week_mask
        if ol_type in (cal.OL_MONTH_NTH, cal.OL_YEAR_NTH):
            rp.Instance = r.instance
        if ol_type in (cal.OL_MONTHLY, cal.OL_YEARLY):
            rp.DayOfMonth = 31 if r.pattern_type == cal.PT_MONTH_END else r.day_of_month
        if ol_type in (cal.OL_YEARLY, cal.OL_YEAR_NTH):
            rp.MonthOfYear = r.start_date.month
        if ol_type != cal.OL_YEARLY or r.interval() > 1:
            rp.Interval = r.interval()
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

        item = folder.Items.Add(kind)
        item.MessageClass = m.message_class
        item.Subject = m.subject
        if m.rtf_body:
            try:
                item.RTFBody = m.rtf_body
            except Exception:
                item.Body = m.plain_body or ""
        elif m.plain_body:
            item.Body = m.plain_body
        if kind == OL_APPOINTMENT_ITEM:
            for r in m.recipients:  # attendees; saving never sends anything
                addr = f'"{r.name}" <{r.email}>' if r.name and r.email else (r.email or r.name)
                if not addr:
                    continue
                try:
                    rec = item.Recipients.Add(addr)
                    rec.Type = r.kind
                    rec.Resolve()
                except Exception as e:
                    self.property_errors["attendees"] = str(e)
        with tempfile.TemporaryDirectory(prefix="unbound_") as tmp:
            self._add_attachments(item, m.attachments, tmp)
            self._copy_props(item, props, skip)
            tz = self._set_times(item, a) if is_appt else None
            item.Save()
        if recurrence is not None:
            # (Outlook resets the series' start "display" zone to this PC's
            # here whatever we set; the recurrence keeps the original zone,
            # which is what defines the series' times.)
            self._set_recurrence(item, recurrence)
            item.Save()
            self._apply_exceptions(item, recurrence, tz)
        # Outlook recomputes some fields on save, so apply the raw properties
        # once more on a fresh copy of the item and save again
        entry_id = item.EntryID
        del item
        item = self.ns.GetItemFromID(entry_id, self.store.StoreID)
        self._copy_props(item, props, skip)
        item.Save()
        return item

    def _set(self, accessor, label, prop_id, prop_type, value):
        if value is None or value == "":
            return
        try:
            accessor.SetProperty(mapi.proptag_url(prop_id, prop_type), value)
        except Exception as e:
            error = str(e).splitlines()[0][:200]
            self.property_errors[label] = error
            # Sender and dates matter to the reader, so flag the item itself
            self.warnings.append(f"couldn't set the {label} ({error})")

    def _set_sender_and_times(self, item, m):
        pa = item.PropertyAccessor
        S, T = mapi.PT_UNICODE, mapi.PT_SYSTIME
        for label, pid in (("sender name", mapi.SENDER_NAME),
                           ("sent-representing name", mapi.SENT_REPRESENTING_NAME)):
            self._set(pa, label, pid, S, m.sender_name)
        for label, pid in (("sender address type", mapi.SENDER_ADDRTYPE),
                           ("sent-representing address type", mapi.SENT_REPRESENTING_ADDRTYPE)):
            self._set(pa, label, pid, S, "SMTP" if m.sender_email else None)
        for label, pid in (("sender email", mapi.SENDER_EMAIL_ADDRESS),
                           ("sent-representing email", mapi.SENT_REPRESENTING_EMAIL_ADDRESS),
                           ("sender SMTP", mapi.SENDER_SMTP_ADDRESS),
                           ("sent-representing SMTP", mapi.SENT_REPRESENTING_SMTP_ADDRESS)):
            self._set(pa, label, pid, S, m.sender_email)
        self._set(pa, "delivery time", mapi.MESSAGE_DELIVERY_TIME, T, com_date(m.delivery_time))
        self._set(pa, "submit time", mapi.CLIENT_SUBMIT_TIME, T, com_date(m.submit_time))

    def add_message(self, folder, m):
        # 1. Create as a PostItem so it is saved as delivered mail, not a draft
        post = folder.Items.Add(OL_POST_ITEM)
        post.Subject = m.subject
        if m.html_body:
            post.HTMLBody = m.html_body
        elif m.rtf_body:
            # Rich-text mail (common between Outlook users) keeps its formatting
            try:
                post.RTFBody = m.rtf_body
            except Exception:
                post.BodyFormat = OL_FORMAT_PLAIN
                post.Body = m.plain_body or ""
        elif m.plain_body:
            # Keep plain-text mail plain; otherwise Outlook stores it as HTML
            post.BodyFormat = OL_FORMAT_PLAIN
            post.Body = m.plain_body

        with tempfile.TemporaryDirectory(prefix="unbound_") as tmp:
            self._add_attachments(post, m.attachments, tmp)
            post.MessageClass = m.message_class
            post.Save()
        entry_id = post.EntryID
        del post

        # 2. Reopen it: now it's a MailItem, so it takes real recipients.
        #    Times are set here because the first save overwrites them.
        item = self.ns.GetItemFromID(entry_id, self.store.StoreID)
        for r in m.recipients:
            addr = f'"{r.name}" <{r.email}>' if r.name and r.email else (r.email or r.name)
            if not addr:
                continue
            try:
                rec = item.Recipients.Add(addr)
                rec.Type = r.kind
                # Without this Outlook stores the raw text as the display
                # name and leaves the email address empty
                rec.Resolve()
            except Exception as e:
                self.property_errors["recipients"] = str(e)
        self._set_sender_and_times(item, m)
        item.UnRead = not m.is_read
        item.Save()
        return item

    def close(self, quit_outlook=False):
        """Detach the PST from the Outlook profile (the file stays on disk)."""
        try:
            self.ns.RemoveStore(self.root)
        finally:
            if quit_outlook:
                self.app.Quit()
            # Drop every COM reference; Outlook keeps the PST open while
            # any are alive.
            self.root = self.store = self.ns = self.app = None
            import gc
            gc.collect()
