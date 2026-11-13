"""PstWriter's dealings with Outlook, against stand-ins (tests/fake_outlook.py)
that behave as Outlook was seen to in the 1.3.3 review."""

import unittest
from datetime import datetime
from types import SimpleNamespace
from unittest import mock

from tests import fake_outlook as fo
from unbound import appointment as cal
from unbound import pst_writer
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

    def set_pattern(self, r):
        pattern = fo.Pattern()
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


if __name__ == "__main__":
    unittest.main()
