"""The separate process that reads OST/PST files."""

import os
import tempfile
import unittest
from unittest import mock

from tests.fakes import AlwaysCrashes, TroubledReader
from unbound import pff_process
from unbound.converter import FolderNode, scan_folders
from unbound.pff_process import PffProcess, ReaderCrashed, ReaderLost, child_process_works


class ReaderProcess(unittest.TestCase):
    def open(self, factory=TroubledReader, path="mailbox"):
        reader = PffProcess(path, factory)
        self.addCleanup(reader.close)
        return reader

    def test_a_child_process_can_be_started(self):
        self.assertTrue(child_process_works())

    def test_answers_requests(self):
        reader = self.open()
        self.assertEqual(reader.scan(), ["tree"])
        self.assertEqual(reader.folder(FolderNode(5, "Inbox", ("Inbox",), 1, None)), 5)
        self.assertIsNone(reader.folder(FolderNode(99, "Gone", ("Gone",), 1, None)))
        message, load_props = reader.item(5, 0)
        self.assertEqual((message, load_props()), ("message 0", None))

    def test_error_on_one_item_is_reported_for_that_item(self):
        reader = self.open()
        with self.assertRaises(RuntimeError) as caught:
            reader.item(5, 4)
        self.assertEqual(str(caught.exception), "bad item")  # first line only
        self.assertEqual(reader.restarts, 0)
        self.assertEqual(reader.item(5, 1)[0], "message 1")

    def test_crash_costs_one_item_and_reading_goes_on(self):
        reader = self.open()
        self.assertEqual(reader.item(5, 1)[0], "message 1")
        with self.assertRaises(ReaderCrashed):
            reader.item(5, 2)
        self.assertEqual(reader.restarts, 1)
        self.assertEqual(reader.item(5, 5)[0], "message 5")  # a new reader took over

    def test_hung_reader_is_given_up_on(self):
        reader = self.open()
        with mock.patch.object(pff_process, "ITEM_TIMEOUT", 0.5):
            with self.assertRaises(ReaderCrashed):
                reader.item(5, 3)
        self.assertEqual(reader.item(5, 0)[0], "message 0")

    def test_reader_that_keeps_crashing_ends_the_run(self):
        reader = self.open(AlwaysCrashes)
        with mock.patch.object(pff_process, "MAX_RESTARTS", 2):
            for _ in range(2):
                with self.assertRaises(ReaderCrashed):
                    reader.item(5, 0)
            with self.assertRaises(ReaderLost):
                reader.item(5, 0)
            with self.assertRaises(ReaderLost):  # and stays ended
                reader.item(5, 1)

    def test_file_that_cannot_be_opened(self):
        with self.assertRaises(OSError) as caught:
            PffProcess("the unopenable", TroubledReader)
        self.assertEqual(str(caught.exception), "no such mailbox")

    def test_killing_readers_unblocks_a_waiting_thread(self):
        reader = self.open()
        self.addCleanup(setattr, pff_process, "_shutting_down", False)
        with mock.patch.object(pff_process, "ITEM_TIMEOUT", 60):
            import threading
            threading.Timer(0.5, pff_process.kill_readers).start()
            with self.assertRaises(ReaderLost):
                reader.item(5, 3)  # would hang for two minutes


class RealFiles(unittest.TestCase):
    """Through libpff itself, with files that aren't mailboxes."""

    def test_file_that_is_not_a_mailbox(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "junk.ost")
            with open(path, "wb") as fh:
                fh.write(b"this is not an OST file " * 50)
            with self.assertRaises(OSError) as caught:
                scan_folders(path)
            self.assertIn("unable to open file", str(caught.exception))

    def test_damaged_files_never_take_the_caller_down(self):
        """Random bytes behind a valid-looking header: libpff may reject the
        file or its process may die, but the caller always gets an exception."""
        import random
        rng = random.Random(42)
        with tempfile.TemporaryDirectory() as folder:
            for n in range(8):
                path = os.path.join(folder, f"damaged{n}.pst")
                with open(path, "wb") as fh:
                    fh.write(b"!BDN" + bytes(rng.randrange(256) for _ in range(4)) + b"SM"
                             + bytes([23, 0, 19, 0]) + bytes(rng.randrange(256) for _ in range(4096)))
                with self.assertRaises((OSError, RuntimeError)):
                    scan_folders(path)


if __name__ == "__main__":
    unittest.main()
