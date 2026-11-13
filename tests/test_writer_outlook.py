"""PstWriter's dealings with Outlook, against stand-ins (tests/fake_outlook.py)
that behave as Outlook was seen to in the 1.3.3 review."""

import os
import threading
import time
import unittest
import uuid
from datetime import datetime
from types import SimpleNamespace
from unittest import mock

from tests import fake_outlook as fo
from unbound import appointment as cal
from unbound import pst_writer
from unbound.props import Prop
from unbound.pst_writer import PstWriter, _detach, html_to_text, rtf_to_text


def writer_with(existing=(), pst=r"C:\out\new.pst"):
    ns = fo.Namespace(existing)
    with mock.patch.object(pst_writer.win32com.client, "Dispatch", return_value=fo.App(ns)):
        w = PstWriter(pst, display_name="Mail")
    return w, ns


def bare_writer():
    """A PstWriter without Outlook, for testing one method at a time."""
    w = PstWriter.__new__(PstWriter)
    w.warnings, w.property_errors = [], {}
    return w


def message(**kw):
    base = dict(subject="s", message_class="IPM.Appointment", plain_body=None, html_body=None,
                rtf_body=None, recipients=[], attachments=[], is_read=True, sender_name="",
                sender_email="", delivery_time=None, submit_time=None)
    base.update(kw)
    return SimpleNamespace(**base)


class FindingTheNewPst(unittest.TestCase):
    """F3: one store Outlook can't open used to stop every conversion."""

    def test_new_pst_is_found_despite_a_store_outlook_cannot_open(self):
        w, ns = writer_with([fo.Store(r"C:\a.ost"), fo.Store(r"C:\shared.ost", broken=True)])
        self.assertEqual(w.store.FilePath, r"C:\out\new.pst")
        self.assertEqual(w.root.Name, "Mail")

    def test_detach_steps_over_a_store_that_errors(self):
        ns = fo.Namespace([fo.Store(r"C:\old.pst"), fo.Store(r"C:\x.ost", broken=True),
                           fo.Store(r"C:\other.pst")])
        self.assertEqual(_detach(ns, [r"C:\old.pst", r"C:\other.pst"]),
                         [r"C:\other.pst", r"C:\old.pst"])
        self.assertEqual([s._path for s in ns.Stores.list], [r"C:\x.ost"])

    def test_a_pst_that_cannot_be_used_is_not_left_attached(self):
        ns = fo.Namespace()
        with mock.patch.object(pst_writer.win32com.client, "Dispatch", return_value=fo.App(ns)), \
                mock.patch.object(PstWriter, "_find_store", side_effect=RuntimeError("lost")):
            with self.assertRaises(RuntimeError):
                PstWriter(r"C:\out\new.pst")
        self.assertEqual(ns.Stores.list, [])
        self.assertEqual(ns.removed, [r"C:\out\new.pst"])


class Hangs(fo.Namespace):
    """An Outlook that stops answering, as one started in the background was
    seen to when asked to detach the PST at the end of a real conversion."""

    def __init__(self, existing=()):
        super().__init__(existing)
        self.answer = threading.Event()

    def RemoveStore(self, root):
        self.answer.wait(30)
        super().RemoveStore(root)


class Refuses(fo.Namespace):
    def RemoveStore(self, root):
        raise fo.OutlookError("The remote procedure call failed.")


class DetachingAtTheEnd(unittest.TestCase):
    """Detaching the PST has a time limit, so a silent Outlook can't keep the
    window at 100% for ever."""

    def close(self, ns, timeout=5):
        with mock.patch.object(pst_writer.win32com.client, "Dispatch", return_value=fo.App(ns)):
            w = PstWriter(r"C:\out\new.pst")
            try:
                w.close(timeout=timeout)
            finally:
                self.writer = w

    def test_detaches_the_pst(self):
        ns = fo.Namespace([fo.Store(r"C:\a.ost", broken=True)])
        self.close(ns)
        self.assertEqual(ns.removed, [r"C:\out\new.pst"])
        self.assertIsNone(self.writer.ns)

    def test_gives_up_on_an_outlook_that_stops_answering(self):
        ns = Hangs()
        self.addCleanup(ns.answer.set)  # let the waiting thread finish
        started = time.time()
        with self.assertRaises(pst_writer.DetachTimeout) as caught:
            self.close(ns, timeout=0.3)
        self.assertLess(time.time() - started, 5)
        self.assertIn("stopped responding", str(caught.exception))
        self.assertIsNone(self.writer.ns)
        # Still queued: Outlook detaches it once it answers again
        ns.answer.set()
        for _ in range(50):
            if ns.removed:
                break
            time.sleep(0.1)
        self.assertEqual(ns.removed, [r"C:\out\new.pst"])

    def test_a_refusal_is_reported(self):
        with self.assertRaises(pst_writer.DetachError) as caught:
            self.close(Refuses())
        self.assertIn("remote procedure call failed", str(caught.exception))

    def test_a_pst_outlook_already_let_go_of_is_fine(self):
        ns = fo.Namespace()
        with mock.patch.object(pst_writer.win32com.client, "Dispatch", return_value=fo.App(ns)):
            w = PstWriter(r"C:\out\new.pst")
            ns.Stores.list.clear()
            w.close(timeout=5)
        self.assertEqual(ns.removed, [])


class Folders(unittest.TestCase):
    """F7: Outlook treats names differing only in case (or accent encoding)
    as the same folder, so every item of the second one used to fail."""

    def test_names_that_differ_only_in_case_share_a_folder(self):
        w, _ = writer_with()
        first = w.get_folder(["Archive"])
        self.assertIs(w.get_folder(["archive"]), first)
        self.assertIs(w.get_folder(["Café"]), w.get_folder(["Cafe\u0301"]))

    def test_a_name_outlook_still_refuses_gets_a_second_folder(self):
        w, _ = writer_with()
        with mock.patch.object(pst_writer, "_folder_key", side_effect=lambda n: n):
            w.get_folder(["Inbox"])
            second = w.get_folder(["INBOX"])
        self.assertEqual(second.Name, "INBOX (2)")


class YearlyRepeats(unittest.TestCase):
    """F2: Outlook counts a yearly interval in months."""

    def recurrence(self, pattern_type, period, freq=cal.FREQ_YEARLY):
        return cal.Recurrence(frequency=freq, pattern_type=pattern_type, period=period,
                              day_of_week_mask=4, day_of_month=9, instance=2,
                              start_date=datetime(2026, 3, 10), end_type=cal.END_NEVER)

    def test_leap_day_birthday(self):
        """A yearly series on 29 February: Outlook works out the day, month
        and interval itself and refuses to be told them (a real birthday
        failed to convert)."""
        r = cal.Recurrence(frequency=cal.FREQ_YEARLY, pattern_type=cal.PT_MONTH, period=12,
                           day_of_month=29, start_date=datetime(2024, 2, 29),
                           end_type=cal.END_NEVER)
        s = self.set_pattern(r, derived={"DayOfMonth": 29, "MonthOfYear": 2, "Interval": 12})
        self.assertEqual(s["RecurrenceType"], cal.OL_YEARLY)
        self.assertNotIn("DayOfMonth", s)
        self.assertTrue(s["NoEndDate"])

    def set_pattern(self, r, derived=None):
        pattern = fo.Pattern(derived)
        item = SimpleNamespace(GetRecurrencePattern=lambda: pattern, AllDayEvent=False)
        bare_writer()._set_recurrence(item, r)
        return pattern.set

    def test_second_tuesday_of_march_every_year(self):
        s = self.set_pattern(self.recurrence(cal.PT_MONTH_NTH, 12))
        self.assertEqual((s["RecurrenceType"], s["Interval"], s["Instance"], s["MonthOfYear"]),
                         (cal.OL_YEAR_NTH, 12, 2, 3))

    def test_every_two_years(self):
        self.assertEqual(self.set_pattern(self.recurrence(cal.PT_MONTH, 24))["Interval"], 24)

    def test_monthly_and_weekly_count_in_their_own_units(self):
        monthly = self.set_pattern(self.recurrence(cal.PT_MONTH, 3, cal.FREQ_MONTHLY))
        weekly = self.set_pattern(self.recurrence(cal.PT_WEEK, 2, cal.FREQ_WEEKLY))
        self.assertEqual((monthly["Interval"], weekly["Interval"]), (3, 2))


class Notes(unittest.TestCase):
    """F1 and F4: notes and bodies are kept as text when the form Outlook
    stores them in can't be written."""

    def test_html_only_notes_are_kept_as_text(self):
        item = fo.Item()
        w = bare_writer()
        w._set_item_body(item, message(html_body="<html><head><style>p{}</style></head><body>"
                                                  "<p>Bring the form.</p><p>Join: <b>https://x.test</b>"
                                                  "</p></body></html>"))
        self.assertEqual(item.Body, "Bring the form.\n\nJoin: https://x.test")
        self.assertEqual(w.warnings, [])

    def test_rtf_kept_when_outlook_takes_it(self):
        item = fo.Item()
        bare_writer()._set_item_body(item, message(rtf_body=b"{\\rtf1 hi}"))
        self.assertEqual(item.RTFBody, b"{\\rtf1 hi}")

    def test_refused_rtf_falls_back_to_text_with_a_warning(self):
        item = fo.Item(refuse_rtf=True)
        w = bare_writer()
        w._set_item_body(item, message(rtf_body=b"{\\rtf1\\ansi Hello\\par world}"))
        self.assertEqual(item.Body, "Hello\nworld")
        self.assertEqual(len(w.warnings), 1)

    def test_text_from_html(self):
        self.assertEqual(html_to_text("a&amp;b<br>c&nbsp;d<script>x()</script>"), "a&b\nc d")
        self.assertEqual(html_to_text(""), "")

    def test_text_from_rtf(self):
        self.assertEqual(rtf_to_text(b"{\\rtf1{\\*\\generator x;}Caf\\'e9 \\u9749?\\par ok}"),
                         "Café ☕\nok")
        # Font and colour tables hold no text
        self.assertEqual(rtf_to_text(b"{\\rtf1{\\fonttbl{\\f0\\fnil Calibri;}}{\\colortbl;\\red0;}"
                                     b"\\pard Met at the conference.\\par\r\n}\r\n\x00\x00"),
                         "Met at the conference.")

    def test_rtf_padding_is_trimmed_before_outlook_sees_it(self):
        """RTF decoded from an OST ends in NUL padding, which Outlook refuses."""
        item = fo.Item()
        bare_writer()._set_item_body(item, message(rtf_body=b"{\\rtf1 hi}\r\n\x00\x00"))
        self.assertEqual(item.RTFBody, b"{\\rtf1 hi}")


class SharedTypeDescriptions(unittest.TestCase):
    """pywin32 asks a COM server to describe every new object afresh. With
    Outlook that was about half the time an email took, so objects of one
    type now share what's been learnt. Tried here with a COM object every
    Windows PC has."""

    def setUp(self):
        raw = pst_writer.win32com.client.Dispatch("Scripting.FileSystemObject")
        self.types = {}
        self.fso = pst_writer._wrap(self.types, raw._oleobj_, "FileSystemObject")

    def test_objects_of_one_type_share_what_is_known(self):
        a = self.fso.GetSpecialFolder(2)  # what a method returns...
        b = self.fso.GetSpecialFolder(2)
        self.assertIsInstance(a, pst_writer._Dispatch)
        self.assertIs(a._olerepr_, b._olerepr_)
        self.assertIs(a._builtMethods_, b._builtMethods_)
        self.assertEqual(a.Path, b.Path)
        self.assertTrue(os.path.isdir(a.Path))
        drives = self.fso.Drives          # ...and a property's value
        self.assertIsInstance(drives, pst_writer._Dispatch)
        self.assertGreaterEqual(drives.Count, 1)
        self.assertIs(a._types_, self.types)

    def test_still_works_like_an_ordinary_object(self):
        folder = self.fso.GetSpecialFolder(2)
        name = "unbound-test-" + uuid.uuid4().hex
        made = folder.SubFolders.Add(name)        # a method with an argument
        try:
            self.assertTrue(self.fso.FolderExists(made.Path))
            self.assertEqual(made.Name, name)
        finally:
            made.Delete()
        self.assertFalse(self.fso.FolderExists(os.path.join(folder.Path, name)))

    def test_outlooks_stand_in_is_left_alone(self):
        app = fo.App(fo.Namespace())
        with mock.patch.object(pst_writer.win32com.client, "Dispatch", return_value=app):
            self.assertIs(pst_writer._outlook(), app)


URL = "http://schemas.microsoft.com/mapi/proptag/0x"


class SettingProperties(unittest.TestCase):
    """An email's sender and dates go to Outlook in one call. Copied
    properties never do: such a call crashed Outlook on a contact."""

    def entries(self):
        return [(URL + "0C1A001F", "Ann", "sender name"), (URL + "0E060040", "date", "delivery time"),
                (URL + "0C1F001F", "ann@example.test", "sender email")]

    def test_one_call_when_asked(self):
        pa, w = fo.PropertyAccessor(), bare_writer()
        self.assertEqual(w._set_properties(pa, self.entries(), together=True), [])
        self.assertEqual([kind for kind, _ in pa.calls], ["batch"])
        self.assertEqual(pa.values[URL + "0C1A001F"], "Ann")

    def test_a_refusal_is_asked_about_again_for_its_text_once(self):
        pa, w = fo.PropertyAccessor(refuse={URL + "0E060040"}), bare_writer()
        refused = w._set_properties(pa, self.entries(), together=True)
        self.assertEqual([label for label, _ in refused], ["delivery time"])
        self.assertIn("does not support", w.property_errors["delivery time"])
        self.assertEqual(pa.calls[1], ("one", URL + "0E060040"))
        self.assertEqual(pa.values[URL + "0C1F001F"], "ann@example.test")  # the rest were set
        pa.calls.clear()
        again = w._set_properties(pa, self.entries(), together=True)   # the next email
        self.assertEqual(again, refused)
        self.assertEqual([kind for kind, _ in pa.calls], ["batch"])    # the text is known by now

    def test_falls_back_to_one_by_one(self):
        pa, w = fo.PropertyAccessor(batches_fail=True), bare_writer()
        self.assertEqual(w._set_properties(pa, self.entries(), together=True), [])
        self.assertEqual([kind for kind, _ in pa.calls], ["batch", "one", "one", "one"])
        self.assertEqual(len(pa.values), 3)

    def test_copied_properties_go_one_by_one(self):
        props = [Prop(0x3A17, 0x1F, None, "Room 5"), Prop(0x3A42, 0x40, None, datetime(2001, 2, 3)),
                 Prop(0x003F, 0x102, None, b"\x01")]
        item = fo.Item()
        item.PropertyAccessor = fo.PropertyAccessor(refuse={URL + "003F0102"})
        w = bare_writer()
        self.assertEqual(w._copy_props(item, props), (2, 1))
        self.assertEqual([kind for kind, _ in item.PropertyAccessor.calls], ["one", "one", "one"])
        self.assertIn("0x003F0102", w.property_errors)

    def test_a_sender_field_outlook_refuses_is_flagged_on_the_email(self):
        item = fo.Item()
        item.PropertyAccessor = fo.PropertyAccessor(refuse={URL + "0C1A001F"})
        w = bare_writer()
        w._set_sender_and_times(item, message(sender_name="Ann", sender_email="ann@example.test",
                                              delivery_time=datetime(2026, 3, 2, 10, 0)))
        self.assertEqual(len(w.warnings), 1)
        self.assertIn("sender name", w.warnings[0])
        self.assertEqual(item.PropertyAccessor.values[URL + "0C1F001F"], "ann@example.test")
        self.assertNotIn(URL + "00390040", item.PropertyAccessor.values)  # no submit time given


class SmallSavings(unittest.TestCase):
    def test_no_temporary_folder_without_attachments(self):
        w = bare_writer()
        w.temp_dir = None
        with mock.patch.object(pst_writer.tempfile, "TemporaryDirectory") as made:
            with w._staging([]) as tmp:
                self.assertIsNone(tmp)
        made.assert_not_called()
        with w._staging([object()]) as tmp:
            self.assertTrue(os.path.isdir(tmp))
        self.assertFalse(os.path.exists(tmp))

    def test_store_id_is_asked_for_once(self):
        w = bare_writer()
        asked = []

        class Store:
            @property
            def StoreID(self):
                asked.append(1)
                return "S"
        w.store = Store()
        w.ns = SimpleNamespace(GetItemFromID=lambda entry, store: (entry, store))
        self.assertEqual([w._item("E1"), w._item("E2")], [("E1", "S"), ("E2", "S")])
        self.assertEqual(len(asked), 1)

    def test_each_time_zone_is_looked_up_once(self):
        looked_up = []

        def zone(name):
            looked_up.append(name)
            if name != "South Africa Standard Time":
                raise fo.OutlookError("not found")
            return "zone object"
        w = bare_writer()
        w.app = SimpleNamespace(TimeZones=SimpleNamespace(Item=zone))
        with mock.patch.object(pst_writer.cal, "tz_key_name", side_effect=lambda blob: blob), \
                mock.patch.object(pst_writer, "windows_zone", return_value=None):
            self.assertEqual([w._tz("South Africa Standard Time") for _ in range(3)],
                             ["zone object"] * 3)
            self.assertEqual([w._tz("Nowhere/Zone") for _ in range(2)], [None, None])
        self.assertEqual(looked_up, ["South Africa Standard Time", "Nowhere/Zone"])
        self.assertEqual(len(w.warnings), 2)  # an unknown zone is still noted on each item

    def zones_writer(self):
        """Outlook's TimeZones, which answers with nothing (not an error)
        for a name it doesn't know."""
        w = bare_writer()
        w.app = SimpleNamespace(TimeZones=SimpleNamespace(
            Item=lambda name: "SAST zone" if name == "South Africa Standard Time" else None))
        return w

    def test_a_phone_style_zone_name_is_translated(self):
        """'Africa/Johannesburg' used to be given up on, silently: Outlook's
        empty answer was taken for a zone, so the Windows name was never
        tried. Found in a real mailbox, where some 40 appointments had it."""
        w = self.zones_writer()
        names = {"Africa/Johannesburg": "South Africa Standard Time"}
        with mock.patch.object(pst_writer.cal, "tz_key_name", side_effect=lambda blob: blob), \
                mock.patch.object(pst_writer, "windows_zone", side_effect=names.get):
            self.assertEqual(w._tz("Africa/Johannesburg"), "SAST zone")
            self.assertIsNone(w._tz("(no TZ description)"))
        self.assertEqual(w.warnings, ["time zone '(no TZ description)' not known to this PC"])

    def test_an_unknown_zone_is_noted_once_per_appointment(self):
        w = self.zones_writer()
        item = SimpleNamespace()
        a = {cal.LID_TZ_START_DISPLAY: "Nowhere/Zone", cal.LID_TZ_END_DISPLAY: "Nowhere/Zone"}
        with mock.patch.object(pst_writer.cal, "tz_key_name", side_effect=lambda blob: blob), \
                mock.patch.object(pst_writer, "windows_zone", return_value=None):
            w.warnings.append("an earlier item's warning")
            self.assertIsNone(w._set_times(item, a))
            self.assertIsNone(w._set_times(item, a))  # the next appointment
        self.assertEqual(w.warnings, ["an earlier item's warning"]
                         + ["time zone 'Nowhere/Zone' not known to this PC"] * 2)


class Recipients(unittest.TestCase):
    """F12: a type Outlook refuses used to skip Resolve, losing the address."""

    def test_address_is_resolved_even_when_the_type_is_refused(self):
        item = SimpleNamespace(Recipients=fo.Recipients(refuse_type=True))
        w = bare_writer()
        w._add_recipients(item, [SimpleNamespace(name="Bob", email="bob@example.test", kind=1)],
                          "recipients")
        rec = item.Recipients.added[0]
        self.assertTrue(rec.resolved)
        self.assertIn("recipients type", w.property_errors)


class NothingHalfWritten(unittest.TestCase):
    """F9: an item reported as failed used to stay in the PST, half made."""

    def test_email_that_fails_after_its_first_save_is_removed(self):
        w = bare_writer()
        post = fo.Item("E1")
        reopened = fo.Item("E1", fail_save=True)
        leftover = fo.Item("E1")
        lookups = iter([reopened, leftover])
        w.ns = SimpleNamespace(GetItemFromID=lambda *a: next(lookups))
        w.store = SimpleNamespace(StoreID="S")
        w.temp_dir = None
        folder = SimpleNamespace(Items=SimpleNamespace(Add=lambda kind: post))
        with self.assertRaises(fo.OutlookError):
            w.add_message(folder, message(message_class="IPM.Note", plain_body="hi"))
        self.assertTrue(leftover.deleted)

    def test_removed_for_good_not_left_in_deleted_items(self):
        """Delete() in a PST only moves an item to Deleted Items, where a
        failed birthday series was found after a real conversion."""
        w = bare_writer()
        bin_ = fo.DeletedItems()
        saved = fo.Item("E1")
        saved.deleted_items = bin_
        w.ns = SimpleNamespace(GetItemFromID=lambda *a: saved)
        w.store = SimpleNamespace(StoreID="S", GetDefaultFolder=lambda kind: bin_)
        w._delete_saved("E1")
        self.assertTrue(saved.deleted)
        self.assertEqual(len(bin_.list), 1)
        self.assertTrue(bin_.list[0].deleted)  # found in Deleted Items and deleted from there


if __name__ == "__main__":
    unittest.main()
