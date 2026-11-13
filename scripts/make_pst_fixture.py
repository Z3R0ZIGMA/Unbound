"""Build a small test PST through Outlook, with calendar items and contacts.

Usage:
    .venv\\Scripts\\python scripts\\make_pst_fixture.py OUTPUT.pst

Lets the calendar and contacts conversion be checked end to end without a
real mailbox: convert OUTPUT.pst with Unbound, then compare the two with
compare_pst.py. It holds a one-off appointment, an all-day event, a meeting
with attendees, a weekly series with a cancelled and a moved occurrence, a
monthly series with an end date, a yearly "second Tuesday of March" series
with notes, a series every two years, and two contacts. Outlook stores the
appointments' notes as HTML only, which is what the notes check relies on.

Needs Outlook Classic. Nothing is ever sent: the items are only saved, the
addresses end in .test (which can't exist), and reminders are off.
"""

import gc
import os
import sys
import time
from datetime import datetime

import win32com.client

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from unbound.pst_writer import (  # noqa: E402
    OL_APPOINTMENT_ITEM, OL_CONTACT_ITEM, OL_FOLDER_CALENDAR, OL_FOLDER_CONTACTS,
    OL_STORE_UNICODE, _stores_newest_first, com_date,
)

OL_WEEKLY, OL_MONTH_NTH, OL_YEARLY, OL_YEAR_NTH = 1, 3, 5, 6
MONDAY, TUESDAY, WEDNESDAY = 2, 4, 8
OL_MEETING = 1


def appointment(folder, subject, start, minutes, **fields):
    a = folder.Items.Add(OL_APPOINTMENT_ITEM)
    a.Subject = subject
    a.Start = com_date(start)
    a.Duration = minutes
    a.ReminderSet = False
    for name, value in fields.items():
        setattr(a, name, value)
    return a


def build(root):
    calendar = root.Folders.Add("Calendar", OL_FOLDER_CALENDAR)
    contacts = root.Folders.Add("Contacts", OL_FOLDER_CONTACTS)

    appointment(calendar, "Dentist", datetime(2027, 3, 3, 14, 0), 45,
                Location="Main Road 12", Body="Bring the form.").Save()

    a = appointment(calendar, "Public holiday – Zürich ☕", datetime(2027, 3, 22), 0)
    a.AllDayEvent = True
    a.Save()

    a = appointment(calendar, "Project review", datetime(2027, 3, 4, 10, 0), 60, Location="Room 1")
    a.MeetingStatus = OL_MEETING
    for address in ("Alice Müller <alice@example.test>", "bob@example.test"):
        a.Recipients.Add(address).Resolve()
    a.Save()  # saved only, never sent

    a = appointment(calendar, "Team sync", datetime(2027, 3, 1, 9, 0), 30)
    pattern = a.GetRecurrencePattern()
    pattern.RecurrenceType = OL_WEEKLY
    pattern.DayOfWeekMask = MONDAY | WEDNESDAY
    pattern.PatternStartDate = com_date(datetime(2027, 3, 1))
    pattern.Occurrences = 8
    a.Save()
    pattern = a.GetRecurrencePattern()
    pattern.GetOccurrence(com_date(datetime(2027, 3, 3, 9, 0))).Delete()
    moved = pattern.GetOccurrence(com_date(datetime(2027, 3, 8, 9, 0)))
    moved.Start = com_date(datetime(2027, 3, 8, 14, 0))
    moved.Subject = "Team sync (moved)"
    moved.Location = "Room 2"
    moved.Save()

    a = appointment(calendar, "Monthly report", datetime(2027, 3, 9, 16, 0), 90)
    pattern = a.GetRecurrencePattern()
    pattern.RecurrenceType = OL_MONTH_NTH
    pattern.DayOfWeekMask = TUESDAY
    pattern.Instance = 2
    pattern.PatternStartDate = com_date(datetime(2027, 3, 9))
    pattern.PatternEndDate = com_date(datetime(2027, 9, 14))
    a.Save()

    # Yearly series. Outlook counts a yearly interval in months, which once
    # made both of these fail to convert
    a = appointment(calendar, "Board meeting", datetime(2027, 3, 9, 9, 0), 120,
                    Body="Agenda and dial-in details in the notes.")
    pattern = a.GetRecurrencePattern()
    pattern.RecurrenceType = OL_YEAR_NTH  # the second Tuesday of March
    pattern.DayOfWeekMask = TUESDAY
    pattern.Instance = 2
    pattern.MonthOfYear = 3
    pattern.PatternStartDate = com_date(datetime(2027, 3, 9))
    pattern.NoEndDate = True
    a.Save()

    a = appointment(calendar, "Licence renewal", datetime(2027, 4, 1, 8, 0), 15)
    pattern = a.GetRecurrencePattern()
    pattern.RecurrenceType = OL_YEARLY
    pattern.Interval = 24  # every two years, in months
    pattern.DayOfMonth = 1
    pattern.MonthOfYear = 4
    pattern.PatternStartDate = com_date(datetime(2027, 4, 1))
    pattern.Occurrences = 5
    a.Save()

    for first, last, email, company in (("Alice", "Müller", "alice@example.test", "Example GmbH"),
                                        ("王", "Carol", "carol@example.test", "")):
        c = contacts.Items.Add(OL_CONTACT_ITEM)
        c.FirstName, c.LastName = first, last
        c.Email1Address = email
        c.CompanyName = company
        c.MobileTelephoneNumber = "+27 82 555 0100"
        c.Body = "Met at the conference."
        c.Save()


def main():
    path = os.path.abspath(sys.argv[1])
    if os.path.exists(path):
        os.remove(path)
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    ns.AddStoreEx(path, OL_STORE_UNICODE)
    # Newest store first, stepping over any Outlook can't open right now
    # (see PstWriter._find_store)
    root = next(s for s, p in _stores_newest_first(ns)
                if p and os.path.normcase(p) == os.path.normcase(path)).GetRootFolder()
    try:
        root.Name = "Unbound fixture"
        build(root)
    finally:
        ns.RemoveStore(root)  # detach it again; the file stays
        del root, ns
        gc.collect()
    for _ in range(120):  # Outlook lets go of the file a moment later
        try:
            with open(path, "ab"):
                break
        except OSError:
            time.sleep(1)
    print(f"Wrote {path}")


if __name__ == "__main__":
    main()
