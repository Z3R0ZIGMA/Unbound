"""Compare a converted PST with its source OST/PST, item by item.

Usage:
    .venv\\Scripts\\python scripts\\compare_pst.py SOURCE.ost OUTPUT.pst

Checks, for every email, appointment and contact in OUTPUT, that the source
has one with the same details:
  emails:        subject, sender, sent and received time, read state, body
                 type, attachments (name, size, inline id), recipients
  appointments:  subject, start, end, all-day, full recurrence (pattern,
                 cancelled and changed occurrences), notes
  contacts:      name, email address, notes
Notes are compared as text (formatting ignored): Outlook stores appointment
notes as HTML, which Unbound can only write back as text.
Also reports how many source items of each kind are missing from OUTPUT
(only meaningful if every folder was converted). Prints counts only, never
message content. Exit code 0 if everything in OUTPUT matches its source.
"""

import collections
from datetime import timedelta
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pypff  # noqa: E402

from unbound import appointment as A  # noqa: E402
from unbound.ost_reader import open_file, read_message  # noqa: E402
from unbound.pst_writer import html_to_text, rtf_to_text  # noqa: E402
from unbound.props import PSETID_ADDRESS, load_name_map, read_props  # noqa: E402

NEVER = (0x2023, 0xFFFFFFFF)
EMAIL1 = (PSETID_ADDRESS, 0x8083)


def _recurrence(a):
    if A.LID_RECUR not in a:
        return None
    r = A.parse_recurrence(a[A.LID_RECUR])
    end = "never" if r.end_type in NEVER else (r.end_type, r.end_date if r.end_type == 0x2021 else r.occurrences)
    return (r.outlook_type(), r.interval(), r.day_of_week_mask, r.day_of_month, r.instance, r.start_date, end,
            r.start_offset, r.end_offset, frozenset(r.deleted),
            tuple(sorted((x.original_start, x.start, x.end) for x in r.exceptions)))


def _second(dt):
    """Times to the nearest second: Outlook stores dates a little less
    precisely than an OST, so sub-second differences aren't real changes."""
    if dt is None:
        return None
    return (dt + timedelta(microseconds=500_000)).replace(microsecond=0)


def _notes(m):
    """An item's notes as plain words, whatever form they're stored in."""
    text = m.plain_body or (html_to_text(m.html_body) if m.html_body else "")         or (rtf_to_text(m.rtf_body) if m.rtf_body else "")
    return " ".join(text.split())


def load(path):
    f = open_file(path)
    name_map = load_name_map(f)
    items = {"email": collections.Counter(), "appointment": collections.Counter(),
             "contact": collections.Counter()}

    def visit(folder):
        for i in range(folder.number_of_sub_messages):
            it = folder.get_sub_message(i)
            if not isinstance(it, pypff.message):
                continue
            m = read_message(it)
            cls = m.message_class.upper()
            if cls.startswith("IPM.NOTE"):
                items["email"][(
                    m.subject, m.sender_email.lower(), _second(m.submit_time), _second(m.delivery_time), m.is_read,
                    "html" if m.html_body else ("rtf" if m.rtf_body else "plain"),
                    tuple(sorted((a.filename, len(a.data), a.content_id) for a in m.attachments)),
                    tuple(sorted((r.kind, r.email.lower()) for r in m.recipients)))] += 1
            elif cls.startswith("IPM.APPOINTMENT"):
                a = A.appointment_named(read_props(it, name_map))
                items["appointment"][(m.subject, _second(a.get(A.LID_START)), _second(a.get(A.LID_END)),
                                      bool(a.get(A.LID_ALL_DAY)), _recurrence(a), _notes(m))] += 1
            elif cls == "IPM.CONTACT":
                n = {(p.named.guid, p.named.lid): p.value for p in read_props(it, name_map)
                     if p.named and p.named.lid is not None}
                items["contact"][(m.subject, (n.get(EMAIL1) or "").lower(), _notes(m))] += 1
        for i in range(folder.number_of_sub_folders):
            visit(folder.get_sub_folder(i))

    visit(f.get_root_folder())
    f.close()
    return items


def main():
    source, output = load(sys.argv[1]), load(sys.argv[2])
    ok = True
    for kind in ("email", "appointment", "contact"):
        out, src = output[kind], source[kind]
        n_out = sum(out.values())
        matched = sum((out & src).values())
        print(f"{kind + 's':13} {matched}/{n_out} in the PST identical to the source"
              f"   ({sum((src - out).values())} source items not in the PST)")
        if matched != n_out:
            ok = False
            # Say which field differs for the first few mismatches
            fields = {"email": ["subject", "sender", "sent time", "received time", "read state",
                                "body type", "attachments", "recipients"],
                      "appointment": ["subject", "start", "end", "all-day", "recurrence", "notes"],
                      "contact": ["name", "email", "notes"]}[kind]
            by_subject = collections.defaultdict(list)
            for k in src:
                by_subject[k[0]].append(k)
            for k in list(out - src)[:5]:
                candidates = by_subject.get(k[0])
                if not candidates:
                    print("    one item's subject isn't in the source")
                    continue
                diff = [fields[i] for i, (x, y) in enumerate(zip(k, candidates[0])) if x != y]
                print(f"    differs in: {', '.join(diff)}")
    print("ALL MATCH" if ok else "MISMATCHES FOUND")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
