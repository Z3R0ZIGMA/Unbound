"""Proof for calendar and contacts: copy a few appointments and a contact
into a new PST, then compare every copied property by reading the PST back
with pypff, and check key fields as Outlook itself reports them.

Usage:
    .venv\\Scripts\\python scripts\\copy_items_test.py SOURCE.ost OUTPUT.pst

Prints pass/fail and property names only, never values.
"""

import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pypff  # noqa: E402

from unbound.ost_reader import open_file, read_message  # noqa: E402
from unbound.props import (  # noqa: E402
    PSETID_ADDRESS, PSETID_APPOINTMENT, load_name_map, read_props,
)
from unbound import appointment as appt  # noqa: E402
from unbound.pst_writer import (  # noqa: E402
    OL_APPOINTMENT_ITEM, OL_CONTACT_ITEM, OL_FOLDER_CALENDAR, OL_FOLDER_CONTACTS,
    PstWriter, from_com_date, should_copy,
)

RECUR = (PSETID_APPOINTMENT, 0x8216)
ALL_DAY = (PSETID_APPOINTMENT, 0x8215)
START = (PSETID_APPOINTMENT, 0x820D)
END = (PSETID_APPOINTMENT, 0x820E)
EMAIL1 = (PSETID_ADDRESS, 0x8083)
TZ = (PSETID_APPOINTMENT, 0x825E)


def named(props):
    return {(p.named.guid, p.named.lid): p for p in props if p.named and p.named.lid is not None}


def container_class(fo):
    for e in fo.record_sets[0].entries:
        if e.entry_type == 0x3613:
            return (e.data_as_string or "").upper()
    return ""


def pick_samples(f, name_map):
    wanted = {"recurring appointment": None, "recurring with changes": None,
              "other time zone": None, "meeting with attendees": None,
              "all-day event": None, "contact with email": None}

    def visit(fo):
        cls = container_class(fo)
        if cls.startswith("IPF.APPOINTMENT") or cls == "IPF.CONTACT":
            for i in range(fo.number_of_sub_messages):
                if all(wanted.values()):
                    return
                item = fo.get_sub_message(i)
                if not isinstance(item, pypff.message):
                    continue
                props = read_props(item, name_map)
                n = named(props)
                m = read_message(item)
                if cls == "IPF.CONTACT":
                    if not wanted["contact with email"] and EMAIL1 in n and m.message_class == "IPM.Contact":
                        wanted["contact with email"] = (m, props)
                    continue
                if m.message_class != "IPM.Appointment":
                    continue
                tz = appt.tz_key_name(n[TZ].value) if TZ in n else None
                if RECUR in n:
                    r = appt.parse_recurrence(n[RECUR].value)
                    if not wanted["recurring with changes"] and r.exceptions and r.deleted:
                        wanted["recurring with changes"] = (m, props)
                    elif not wanted["recurring appointment"] and not r.exceptions:
                        wanted["recurring appointment"] = (m, props)
                if not wanted["other time zone"] and tz and tz != "South Africa Standard Time":
                    wanted["other time zone"] = (m, props)
                elif not wanted["meeting with attendees"] and m.recipients and RECUR not in n:
                    wanted["meeting with attendees"] = (m, props)
                elif not wanted["all-day event"] and n.get(ALL_DAY) and n[ALL_DAY].value:
                    wanted["all-day event"] = (m, props)
        for i in range(fo.number_of_sub_folders):
            visit(fo.get_sub_folder(i))

    visit(f.get_root_folder())
    return {k: v for k, v in wanted.items() if v}


LOCATION = (PSETID_APPOINTMENT, 0x8208)


def effective(r, series_subject, n):
    series_location = n[LOCATION].value if LOCATION in n else ""
    return sorted((x.original_start,
                   x.subject if x.subject is not None else series_subject,
                   x.location if x.location is not None else series_location)
                  for x in r.exceptions)


def same(a, b):
    if hasattr(a, "timestamp") and hasattr(b, "timestamp"):
        return abs((a - b).total_seconds()) < 1
    return a == b


def oom_dt(v):
    return v.replace(tzinfo=None) if v is not None else None


def write_items(out, samples):
    """Write the samples and check them through Outlook; all COM objects
    are local here so they are released when this returns."""
    w = PstWriter(out, display_name="Unbound items test")
    oom = {}
    try:
        cal = w.get_folder(["Calendar"], OL_FOLDER_CALENDAR)
        con = w.get_folder(["Contacts"], OL_FOLDER_CONTACTS)
        for label, (m, props) in samples.items():
            is_contact = label.startswith("contact")
            item = w.add_item(con if is_contact else cal,
                              OL_CONTACT_ITEM if is_contact else OL_APPOINTMENT_ITEM, m, props)
            n = named(props)
            checks = {"subject": item.Subject == m.subject,
                      "attachments": item.Attachments.Count >= len(
                          [a for a in m.attachments if a.filename != "ContactPicture.jpg"])}
            if is_contact:
                checks["Email1 (Outlook view)"] = (item.Email1Address or "").lower() == (n[EMAIL1].value or "").lower()
                checks["has a name"] = bool(item.FullName or item.FileAs)
            else:
                checks["start (Outlook view)"] = same(from_com_date(item.StartUTC), n[START].value)
                if not checks["start (Outlook view)"]:
                    print(f"   {label}: StartUTC off by {from_com_date(item.StartUTC) - n[START].value}")
                checks["end (Outlook view)"] = same(from_com_date(item.EndUTC), n[END].value)
                checks["recurring flag"] = bool(item.IsRecurring) == (RECUR in n)
                if TZ in n:
                    checks["time zone"] = item.StartTimeZone.ID == appt.tz_key_name(n[TZ].value)
                    if not checks["time zone"]:
                        print(f"   {label}: zone {item.StartTimeZone.ID!r} vs {appt.tz_key_name(n[TZ].value)!r}")
                checks["attendee count"] = item.Recipients.Count == len(m.recipients)
                if ALL_DAY in n:
                    checks["all-day flag"] = bool(item.AllDayEvent) == bool(n[ALL_DAY].value)
            oom[label] = checks
            del item
    finally:
        errors = dict(w.property_errors)
        errors.update({f"warning {i}": t for i, t in enumerate(w.warnings)})
        w.close()
        del w
    return oom, errors


def main():
    src, out = sys.argv[1], os.path.abspath(sys.argv[2])
    if os.path.exists(out):
        print(f"{out} already exists; delete it first")
        return 1
    os.makedirs(os.path.dirname(out), exist_ok=True)

    f = open_file(src)
    name_map = load_name_map(f)
    samples = pick_samples(f, name_map)
    f.close()
    print("Samples found:", ", ".join(samples) or "none")

    oom, errors = write_items(out, samples)
    import gc
    gc.collect()

    for _ in range(120):
        try:
            pst = open_file(out)
            break
        except OSError:
            time.sleep(1)
    else:
        print("PST still locked")
        return 1
    out_map = load_name_map(pst)
    written = {}

    def collect(fo):
        for i in range(fo.number_of_sub_messages):
            it = fo.get_sub_message(i)
            written[(it.subject or "")] = read_props(it, out_map)
        for i in range(fo.number_of_sub_folders):
            collect(fo.get_sub_folder(i))
    collect(pst.get_root_folder())
    pst.close()

    for label, (m, props) in samples.items():
        print(f"\n== {label}")
        for k, ok in oom[label].items():
            print(f"  {'PASS' if ok else 'FAIL'}  {k}")
        got = written.get(m.subject)
        if got is None:
            print("  FAIL  item not found in PST")
            continue
        def key(p):
            return (p.named.guid, p.named.lid, p.named.name) if p.named else p.tag_id
        got_by = {key(p): p for p in got}
        copied = [p for p in props if should_copy(p) and not appt.is_managed(p)]
        missing = [p for p in copied if key(p) not in got_by]
        differ = [p for p in copied if key(p) in got_by and not same(got_by[key(p)].value, p.value)]
        print(f"  raw properties: {len(copied) - len(missing) - len(differ)}/{len(copied)} identical, "
              f"{len(missing)} missing, {len(differ)} changed")
        for p in (missing + differ)[:12]:
            tag = (f"{p.named.guid[:9]}../{p.named.lid:04X}" if p.named and p.named.lid is not None
                   else (f"{p.named.guid[:9]}../{p.named.name}" if p.named else f"0x{p.tag_id:04X}"))
            print(f"     {'missing' if p in missing else 'changed'}: {tag} type 0x{p.prop_type:04X}")
        src_n, out_n = named(props), named(got)
        if RECUR in src_n:
            a = appt.parse_recurrence(src_n[RECUR].value)
            if RECUR not in out_n:
                print("  FAIL  recurrence missing in PST")
                continue
            b = appt.parse_recurrence(out_n[RECUR].value)
            never = (0x2023, 0xFFFFFFFF)

            def end_of(r):
                if r.end_type in never:
                    return "never"
                return (r.end_type, r.end_date if r.end_type == 0x2021 else r.occurrences)

            rc = {
                "repeat type and interval": (a.outlook_type(), a.interval()) == (b.outlook_type(), b.interval()),
                "days / instance": (a.day_of_week_mask, a.day_of_month, a.instance)
                                   == (b.day_of_week_mask, b.day_of_month, b.instance),
                "series start date": a.start_date == b.start_date,
                "series end": end_of(a) == end_of(b),
                "time of day and length": (a.start_offset, a.end_offset) == (b.start_offset, b.end_offset),
                "cancelled occurrences": set(a.deleted) == set(b.deleted),
                "changed occurrences": sorted((x.original_start, x.start, x.end) for x in a.exceptions)
                                       == sorted((x.original_start, x.start, x.end) for x in b.exceptions),
                # Compare what each occurrence shows: an override equal to the
                # series value is the same thing as no override
                "changed subjects/locations": effective(a, m.subject, src_n) == effective(b, m.subject, out_n),
            }
            for k, ok in rc.items():
                print(f"  {'PASS' if ok else 'FAIL'}  recurrence: {k}")
            if not all(rc.values()):
                print(f"     source: type {a.outlook_type()} every {a.interval()} start {a.start_date:%Y-%m-%d} "
                      f"end {end_of(a)} offsets {a.start_offset}-{a.end_offset} "
                      f"cancelled {len(a.deleted)} changed {len(a.exceptions)}")
                print(f"     PST:    type {b.outlook_type()} every {b.interval()} start {b.start_date:%Y-%m-%d} "
                      f"end {end_of(b)} offsets {b.start_offset}-{b.end_offset} "
                      f"cancelled {len(b.deleted)} changed {len(b.exceptions)}")
    if errors:
        print(f"\nOutlook refused {len(errors)} properties, e.g.:")
        for k, v in list(errors.items())[:8]:
            print(f"  {k}: {v[:110]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
