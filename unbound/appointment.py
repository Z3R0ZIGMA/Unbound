"""Calendar details Outlook won't accept as raw properties.

Outlook protects an appointment's recurrence pattern and time-zone
properties, and recomputes its start/end on save, so these are decoded here
and applied through the object model instead ([MS-OXOCAL] 2.2.1.44).
"""

import struct
from dataclasses import dataclass, field
from datetime import datetime, timedelta

from .props import PSETID_APPOINTMENT

# Appointment named properties (PSETID_Appointment LIDs)
LID_START = 0x820D  # PidLidAppointmentStartWhole
LID_END = 0x820E  # PidLidAppointmentEndWhole
LID_DURATION = 0x8213
LID_ALL_DAY = 0x8215  # PidLidAppointmentSubType
LID_RECUR = 0x8216  # PidLidAppointmentRecur
LID_RECURRING = 0x8223
LID_RECURRENCE_TYPE = 0x8231
LID_RECURRENCE_PATTERN = 0x8232
LID_TZ_STRUCT = 0x8233
LID_TZ_DESCRIPTION = 0x8234
LID_CLIP_START = 0x8235
LID_CLIP_END = 0x8236
LID_TZ_START_DISPLAY = 0x825E
LID_TZ_END_DISPLAY = 0x825F
LID_TZ_RECUR = 0x8260

# Set by our code through the object model (or recomputed by Outlook), so
# not copied as raw properties
MANAGED_LIDS = {
    LID_START, LID_END, LID_DURATION, LID_ALL_DAY, LID_RECUR, LID_RECURRING,
    LID_RECURRENCE_TYPE, LID_RECURRENCE_PATTERN, LID_TZ_STRUCT, LID_TZ_DESCRIPTION,
    LID_CLIP_START, LID_CLIP_END, LID_TZ_START_DISPLAY, LID_TZ_END_DISPLAY, LID_TZ_RECUR,
}
MANAGED_TAGS = {0x0060, 0x0061}  # PR_START_DATE / PR_END_DATE

EPOCH_1601 = datetime(1601, 1, 1)

# RecurrencePattern.PatternType
PT_DAY, PT_WEEK, PT_MONTH, PT_MONTH_NTH, PT_MONTH_END = 0x0, 0x1, 0x2, 0x3, 0x4
# RecurFrequency
FREQ_DAILY, FREQ_WEEKLY, FREQ_MONTHLY, FREQ_YEARLY = 0x200A, 0x200B, 0x200C, 0x200D
END_AFTER_DATE, END_AFTER_N, END_NEVER, END_NEVER2 = 0x2021, 0x2022, 0x2023, 0xFFFFFFFF

# Outlook object model OlRecurrenceType
OL_DAILY, OL_WEEKLY, OL_MONTHLY, OL_MONTH_NTH, OL_YEARLY, OL_YEAR_NTH = 0, 1, 2, 3, 5, 6

# ExceptionInfo.OverrideFlags
ARO_SUBJECT, ARO_MEETINGTYPE, ARO_REMINDERDELTA, ARO_REMINDER = 0x1, 0x2, 0x4, 0x8
ARO_LOCATION, ARO_BUSYSTATUS, ARO_ATTACHMENT, ARO_SUBTYPE = 0x10, 0x20, 0x40, 0x80


def minutes_to_dt(minutes):
    return EPOCH_1601 + timedelta(minutes=minutes)


@dataclass
class ExceptionInfo:
    start: datetime  # new start, in the appointment's own time zone
    end: datetime
    original_start: datetime
    subject: str | None = None
    location: str | None = None
    busy_status: int | None = None
    reminder_set: bool | None = None
    reminder_minutes: int | None = None
    all_day: bool | None = None


@dataclass
class Recurrence:
    frequency: int
    pattern_type: int
    period: int
    day_of_week_mask: int = 0
    day_of_month: int = 0
    instance: int = 0  # 1-4, 5 = last
    end_type: int = END_NEVER
    occurrences: int = 0
    first_dow: int = 0
    deleted: list = field(default_factory=list)  # original dates (midnight)
    modified: list = field(default_factory=list)
    start_date: datetime | None = None  # pattern start (midnight)
    end_date: datetime | None = None
    start_offset: int = 0  # minutes after midnight, appointment's time zone
    end_offset: int = 0
    exceptions: list = field(default_factory=list)

    def outlook_type(self):
        yearly = self.frequency == FREQ_YEARLY
        if self.pattern_type == PT_DAY:
            return OL_DAILY
        if self.pattern_type == PT_WEEK:
            return OL_WEEKLY
        if self.pattern_type in (PT_MONTH, PT_MONTH_END):
            return OL_YEARLY if yearly else OL_MONTHLY
        if self.pattern_type == PT_MONTH_NTH:
            return OL_YEAR_NTH if yearly else OL_MONTH_NTH
        return None  # Hijri calendars etc.: not supported

    def interval(self):
        if self.pattern_type == PT_DAY:
            return max(1, self.period // 1440)
        if self.frequency == FREQ_YEARLY:
            return max(1, self.period // 12)
        return max(1, self.period)


def _read_string(buf, off, wide):
    (count,) = struct.unpack_from("<H", buf, off)
    off += 2
    size = count * 2 if wide else count
    raw = buf[off:off + size]
    return raw.decode("utf-16-le" if wide else "cp1252", errors="replace"), off + size


def parse_recurrence(blob):
    """Decode PidLidAppointmentRecur. Raises ValueError if it can't."""
    try:
        return _parse_recurrence(blob)
    except (struct.error, IndexError) as e:
        raise ValueError(f"unreadable recurrence data: {e}") from e


def _parse_recurrence(b):
    off = 0
    reader, writer, freq, ptype, _cal, _first, period, _sliding = struct.unpack_from("<HHHHHIII", b, off)
    off += 22
    if reader != 0x3004:
        raise ValueError(f"unknown recurrence version 0x{reader:04X}")
    r = Recurrence(frequency=freq, pattern_type=ptype, period=period)
    if ptype == PT_WEEK:
        (r.day_of_week_mask,) = struct.unpack_from("<I", b, off)
        off += 4
    elif ptype in (PT_MONTH, PT_MONTH_END, 0xA, 0xC):
        (r.day_of_month,) = struct.unpack_from("<I", b, off)
        off += 4
    elif ptype in (PT_MONTH_NTH, 0xB):
        r.day_of_week_mask, r.instance = struct.unpack_from("<II", b, off)
        off += 8
    r.end_type, r.occurrences, r.first_dow, n_deleted = struct.unpack_from("<IIII", b, off)
    off += 16
    r.deleted = [minutes_to_dt(m) for m in struct.unpack_from(f"<{n_deleted}I", b, off)]
    off += 4 * n_deleted
    (n_modified,) = struct.unpack_from("<I", b, off)
    off += 4
    r.modified = [minutes_to_dt(m) for m in struct.unpack_from(f"<{n_modified}I", b, off)]
    off += 4 * n_modified
    start, end = struct.unpack_from("<II", b, off)
    off += 8
    r.start_date = minutes_to_dt(start)
    r.end_date = minutes_to_dt(end)

    # AppointmentRecurrencePattern part
    _r2, _w2, r.start_offset, r.end_offset, n_exc = struct.unpack_from("<IIIIH", b, off)
    off += 18
    for _ in range(n_exc):
        s, e, orig, flags = struct.unpack_from("<IIIH", b, off)
        off += 14
        x = ExceptionInfo(minutes_to_dt(s), minutes_to_dt(e), minutes_to_dt(orig))
        if flags & ARO_SUBJECT:
            off += 2  # SubjectLength (includes the length field)
            x.subject, off = _read_string(b, off, wide=False)
        if flags & ARO_MEETINGTYPE:
            off += 4
        if flags & ARO_REMINDERDELTA:
            (x.reminder_minutes,) = struct.unpack_from("<I", b, off)
            off += 4
        if flags & ARO_REMINDER:
            (rs,) = struct.unpack_from("<I", b, off)
            x.reminder_set = bool(rs)
            off += 4
        if flags & ARO_LOCATION:
            off += 2
            x.location, off = _read_string(b, off, wide=False)
        if flags & ARO_BUSYSTATUS:
            (x.busy_status,) = struct.unpack_from("<I", b, off)
            off += 4
        if flags & ARO_ATTACHMENT:
            off += 4
        if flags & ARO_SUBTYPE:
            (sub,) = struct.unpack_from("<I", b, off)
            x.all_day = bool(sub)
            off += 4
        # ARO_EXCEPTIONAL_BODY (0x100) and ARO_APPTCOLOR (0x200) carry no
        # data here; checked against real files, where a 4-byte colour field
        # misaligned every following exception
        r.exceptions.append(x)
    # Extended exceptions carry Unicode versions of subject/location
    try:
        (reserved1,) = struct.unpack_from("<I", b, off)
        off += 4 + reserved1
        for i, x in enumerate(r.exceptions):
            flags = None
            # ExtendedException: ChangeHighlight (if writer2 >= 0x3009), reserved, times, strings
            if _w2 >= 0x3009:
                (ch_size,) = struct.unpack_from("<I", b, off)
                off += 4 + ch_size
            (res_size,) = struct.unpack_from("<I", b, off)
            off += 4 + res_size
            if x.subject is not None or x.location is not None:
                off += 12  # start, end, original start
                if x.subject is not None:
                    x.subject, off = _read_string(b, off, wide=True)
                if x.location is not None:
                    x.location, off = _read_string(b, off, wide=True)
                (res2,) = struct.unpack_from("<I", b, off)
                off += 4 + res2
    except struct.error:
        pass  # keep the cp1252 strings from ExceptionInfo
    return r


def tz_key_name(blob):
    """Windows time zone key name from a TZDEFINITION blob, e.g.
    'South Africa Standard Time'."""
    if not blob or len(blob) < 8:
        return None
    (n,) = struct.unpack_from("<H", blob, 6)
    try:
        return blob[8:8 + n * 2].decode("utf-16-le")
    except UnicodeDecodeError:
        return None


def appointment_named(props):
    """{LID: value} for this item's PSETID_Appointment properties."""
    return {p.named.lid: p.value for p in props
            if p.named and p.named.guid == PSETID_APPOINTMENT and p.named.lid is not None}


def is_managed(p):
    if p.named is not None:
        return p.named.guid == PSETID_APPOINTMENT and p.named.lid in MANAGED_LIDS
    return p.tag_id in MANAGED_TAGS
