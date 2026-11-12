"""Files that another program has open and has locked part of, as Word and
Excel do to the Outlook mailbox while they run."""

import os
import subprocess
import sys
import tempfile
import unittest
from unittest import mock

from unbound import converter
from unbound.ost_reader import (
    MappedFile, OstInUseError, _names, is_locked, programs_using,
)

# Holds a lock on the first 64 KB of a file, as a separate program, until
# its input is closed
LOCKER = """
import msvcrt, sys
fh = open(sys.argv[1], "r+b")
msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 65536)
print("locked", flush=True)
sys.stdin.read()
"""


class LockedFile(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = tmp.name
        self.path = os.path.join(self.dir, "mail.ost")
        self.data = os.urandom(300_000)
        with open(self.path, "wb") as fh:
            fh.write(self.data)

    def lock(self):
        proc = subprocess.Popen([sys.executable, "-c", LOCKER, self.path],
                                stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
        self.addCleanup(self.unlock, proc)
        self.assertEqual(proc.stdout.readline().strip(), "locked")
        return proc

    def unlock(self, proc):
        if proc.poll() is None:
            proc.stdin.close()
            proc.wait(10)
        proc.stdout.close()

    def test_lock_is_noticed_and_goes_away(self):
        self.assertFalse(is_locked(self.path))
        proc = self.lock()
        self.assertTrue(is_locked(self.path))
        with open(self.path, "rb") as fh:
            with self.assertRaises(PermissionError):
                fh.read(100)
        self.unlock(proc)
        self.assertFalse(is_locked(self.path))

    def test_windows_names_the_program_that_has_the_file_open(self):
        self.assertEqual(programs_using(self.path), [])
        self.lock()
        programs = programs_using(self.path)
        self.assertEqual(len(programs), 1, programs)
        self.assertIn("python", programs[0].lower())

    def test_mapped_view_reads_past_the_lock(self):
        self.lock()
        mapped = MappedFile(self.path)
        self.addCleanup(mapped.close)
        self.assertEqual(mapped.get_size(), len(self.data))
        self.assertEqual(mapped.read(4096), self.data[:4096])
        self.assertEqual(mapped.tell(), 4096)
        mapped.seek(-10, os.SEEK_END)
        self.assertEqual(mapped.read(), self.data[-10:])
        mapped.seek(70_000)
        self.assertEqual(mapped.read(5), self.data[70_000:70_005])

    def test_working_copy_is_complete_despite_the_lock(self):
        self.lock()
        seen = []
        with mock.patch.object(converter, "COPY_CHUNK", 50_000):
            copy = converter._working_copy(self.path, tempfile.mkdtemp(dir=self.dir),
                                           lambda done, size: seen.append(done))
        with open(copy, "rb") as fh:
            self.assertEqual(fh.read(), self.data)
        self.assertEqual(seen[-1], len(self.data))

    def test_copy_of_an_unlocked_file_is_unchanged(self):
        target = os.path.join(self.dir, "copy.ost")
        converter._copy_file(self.path, target)
        with open(target, "rb") as fh:
            self.assertEqual(fh.read(), self.data)


class Messages(unittest.TestCase):
    def test_names(self):
        self.assertEqual(_names(["Microsoft Word"]), "Microsoft Word")
        self.assertEqual(_names(["Microsoft Excel", "Microsoft Word"]),
                         "Microsoft Excel and Microsoft Word")
        self.assertEqual(_names(["A", "B", "C"]), "A, B and C")

    def test_outlook_has_the_file(self):
        e = OstInUseError("C:\\mail.ost", ["Microsoft Excel", "Microsoft Outlook"])
        self.assertTrue(e.outlook)
        self.assertIn("Close Outlook and try again", str(e))
        self.assertIn("running in the background", str(e))
        self.assertIn("File: C:\\mail.ost", str(e))

    def test_other_programs_have_the_file(self):
        e = OstInUseError("C:\\mail.ost", ["Microsoft Excel", "Microsoft Word"])
        self.assertFalse(e.outlook)
        self.assertIn("Microsoft Excel and Microsoft Word have this file open", str(e))
        self.assertIn("Close them and try again", str(e))
        one = OstInUseError("C:\\mail.ost", ["Microsoft Word"])
        self.assertIn("Microsoft Word has this file open", str(one))
        self.assertIn("Close it and try again", str(one))

    def test_windows_cannot_say_who_has_the_file(self):
        e = OstInUseError("C:\\mail.ost")
        self.assertFalse(e.outlook)
        self.assertEqual(e.programs, [])
        self.assertIn("Another program is using this file", str(e))


if __name__ == "__main__":
    unittest.main()
