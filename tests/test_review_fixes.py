"""Fixes from the 1.3.3 code review that aren't about the window or the
PST writer (those are in test_app.py and test_writer_outlook.py)."""

import os
import tempfile
import time
import unittest
from unittest import mock

from tests.fakes import DiesWhileOpening, FailsSilently, HangsWhileOpening
from tests.test_mbox import mbox
from unbound import mapi, mbox_reader, pff_process
from unbound.converter import MAIL, Failure, Result, convert, scan_folders, write_log
from unbound.errors import first_line
from unbound.mbox_reader import is_mbox_source, message_offsets, parse_message
from unbound.pff_process import PffProcess, ReaderLost


class TempDir(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = tmp.name

    def write(self, name, data):
        path = os.path.join(self.dir, name)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(data)
        return path


class ErrorsWithoutAMessage(TempDir):
    """F5: `str(e).splitlines()[0]` raised IndexError for an exception with no
    message, inside the very handler meant to report it."""

    def test_first_line(self):
        self.assertEqual(first_line(ValueError("one\ntwo")), "one")
        self.assertEqual(first_line(MemoryError()), "MemoryError")
        self.assertEqual(first_line(ValueError("x" * 500), 10), "x" * 10)

    def test_such_an_error_costs_one_item_not_the_run(self):
        export = os.path.join(self.dir, "export")
        self.write("export/Inbox.mbox/mbox", mbox(b"Subject: one\n\nx\n", b"Subject: two\n\nx\n",
                                                   b"Subject: three\n\nx\n"))
        nodes = [n for top in scan_folders(export) for n in top.walk()]
        result = convert(export, os.path.join(self.dir, "out.pst"), nodes,
                         outlook_writer=FailsSilently(fail_on={"two"}))
        self.assertEqual(result.copied[MAIL], 2)
        self.assertEqual(result.failures, [Failure("Inbox", 1, "two", "SilentError")])


class LogLines(TempDir):
    """F15: a line break in a subject could write fake lines into the log."""

    def test_subjects_stay_on_one_line(self):
        result = Result()
        result.failures.append(Failure("In\nbox", 0, "x\nFailed: 0\nNo items failed.", "bad\nmore"))
        log = os.path.join(self.dir, "log.txt")
        write_log(result, "in.ost", "out.pst", log)
        with open(log, encoding="utf-8") as fh:
            lines = fh.read().splitlines()
        self.assertIn("In box | 1 | x Failed: 0 No items failed. | bad more", lines)
        self.assertEqual(lines.count("Failed: 0"), 0)


class ParserEdges(unittest.TestCase):
    """F13: a damaged value or an absurd date failed a whole item."""

    def test_empty_true_false_value_is_skipped(self):
        from tests.test_props import Entry, record_sets
        from unbound.props import PT_BOOLEAN, read_props
        item = type("Item", (), {"record_sets": record_sets([Entry(0x0E1F, PT_BOOLEAN, b""),
                                                              Entry(0x0037, 0x1F, b"")])})
        props = read_props(item, {})
        self.assertNotIn(0x0E1F, [p.tag_id for p in props])

    def test_email_dated_in_the_year_9999_keeps_no_date(self):
        m = parse_message(b"Subject: far\nDate: Fri, 31 Dec 9999 23:59:59 -0500\n\nbody\n")
        self.assertEqual(m.subject, "far")
        self.assertIsNone(m.submit_time)


class WhatCountsAsAMailbox(TempDir):
    """F18: any file starting "From " was taken for a mailbox."""

    def test_text_file_starting_with_from_is_not_a_mailbox(self):
        notes = self.write("notes.txt", b"From here on this is just a text file.\n")
        self.assertFalse(is_mbox_source(notes))

    def test_unnamed_mailbox_with_a_dated_separator_still_counts(self):
        box = self.write("export", b"From a@b.test Mon Jan  1 00:00:00 2024\nSubject: x\n\nhi\n")
        self.assertTrue(is_mbox_source(box))

    def test_a_file_with_no_line_breaks_is_read_in_pieces(self):
        data = b"From a@b.test Mon Jan  1 00:00:00 2024\nSubject: x\n\n" + b"y" * 5000 + b"\n"
        path = self.write("big.mbox", data)
        with mock.patch.object(mbox_reader, "_MAX_PIECE", 64):
            self.assertEqual(message_offsets(path), [0])


class ReaderStartUp(unittest.TestCase):
    """F6: a reader that died or hung while opening the file froze the window."""

    def test_reader_that_dies_while_opening_is_noticed(self):
        started = time.time()
        with self.assertRaises(ReaderLost):
            PffProcess("mailbox", DiesWhileOpening)
        self.assertLess(time.time() - started, 30)

    def test_reader_that_hangs_while_opening_can_be_stopped(self):
        import threading
        self.addCleanup(setattr, pff_process, "_shutting_down", False)
        threading.Timer(1.0, pff_process.kill_readers).start()
        started = time.time()
        with self.assertRaises(ReaderLost):
            PffProcess("mailbox", HangsWhileOpening)
        self.assertLess(time.time() - started, 30)


class SmallThings(unittest.TestCase):
    def test_recipient_types(self):
        """F12: flags in the high bits made Outlook refuse the type."""
        self.assertEqual([mapi.recipient_kind(v) for v in (1, 2, 3, 0x80000001, 0x10000002, 4, None)],
                         [1, 2, 3, 1, 2, 1, 1])

    def test_system_dlls_are_loaded_by_full_path(self):
        """F8: a DLL with a system DLL's name next to Unbound.exe was loaded."""
        with mock.patch.dict(os.environ, {"SystemRoot": r"D:\Win"}):
            self.assertEqual(mapi.system_dll("icu.dll"), r"D:\Win\System32\icu.dll")

    def test_exe_test_modes_end_in_an_exit_code(self):
        """F11: an error in the windowed EXE's test modes showed a box that
        waited for a click, instead of returning an exit code."""
        import unbound_app
        with tempfile.TemporaryDirectory() as folder:
            report = os.path.join(folder, "check.txt")

            def explode(*args):
                raise ImportError("no module")
            self.assertEqual(unbound_app._guarded(explode, report, fallback_file=report, code=1), 1)
            with open(report, encoding="utf-8") as fh:
                self.assertIn("FAIL ImportError: no module", fh.read())
            missing = os.path.join(folder, "no such folder", "out.pst")
            with mock.patch("unbound.converter.scan_folders", side_effect=OSError("unreadable")):
                self.assertEqual(unbound_app.convert_without_window("in.ost", missing), 2)


if __name__ == "__main__":
    unittest.main()
