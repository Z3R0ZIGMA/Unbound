import os
import tempfile
import unittest
from unittest import mock

from tests.fakes import FakeFolder, FakeWriter
from tests.test_mbox import mbox
from unbound import __version__, converter
from unbound.converter import (
    CALENDAR, CONTACTS, MAIL, Failure, FolderNode, Result, convert, estimated_pst_size,
    is_outlook_data_file, item_kind, recover_interrupted, scan_folders, write_failure_log,
    write_log,
)


def node(name, items=0, container_class="IPF.Note", children=(), path=None):
    return FolderNode(1, name, path or (name,), items, container_class, list(children))


class Kinds(unittest.TestCase):
    def test_folder_kinds(self):
        for container_class, kind in (("IPF.Note", MAIL), (None, MAIL), ("IPF.Imap", MAIL),
                                      ("IPF.Appointment", CALENDAR), ("IPF.Contact", CONTACTS),
                                      ("IPF.Contact.RecipientCache", None), ("IPF.Task", None)):
            self.assertEqual(node("x", container_class=container_class).kind, kind, container_class)

    def test_item_kinds(self):
        for message_class, kind in (("IPM.Note", MAIL), ("ipm.note.smime", MAIL),
                                    ("IPM.Appointment", CALENDAR), ("IPM.Contact", CONTACTS),
                                    ("IPM.Contact.Custom", CONTACTS), ("IPM.ContactGroup", None),
                                    ("IPM.Schedule.Meeting.Request", None), ("IPM.Task", None)):
            self.assertEqual(item_kind(message_class), kind, message_class)

    def test_what_starts_ticked(self):
        self.assertTrue(node("Inbox").default_checked)
        self.assertFalse(node("Sync Issues").default_checked)        # system folder
        self.assertFalse(node("Tasks", container_class="IPF.Task").default_checked)
        shared = FolderNode(1, "Inbox", ("Shared mailboxes", "Inbox"), 3, "IPF.Note", [], True)
        self.assertFalse(shared.default_checked)


class Scanning(unittest.TestCase):
    def tree(self, root):
        return converter._visible_tree(converter._scan(root, ()))

    def test_exchange_layout(self):
        root = FakeFolder(1, "", children=[
            FakeFolder(2, "Root - Mailbox", children=[
                FakeFolder(3, "IPM_SUBTREE", children=[
                    FakeFolder(4, "Inbox", 5, [FakeFolder(5, "Sub", 2)], "IPF.Note"),
                    FakeFolder(6, "Calendar", 3, container_class="IPF.Appointment")]),
                FakeFolder(7, "Shared Data", children=[FakeFolder(8, "Team", 9)]),
                FakeFolder(9, "Finder", 1)])])
        top = self.tree(root)
        self.assertEqual([n.name for n in top], ["Inbox", "Calendar", "Shared mailboxes"])
        self.assertEqual(top[0].children[0].path, ("Inbox", "Sub"))
        self.assertEqual(top[1].kind, CALENDAR)
        team = top[2].children[0]
        self.assertEqual(team.path, ("Shared mailboxes", "Team"))
        self.assertTrue(team.shared)
        self.assertFalse(team.default_checked)

    def test_pst_and_imap_layout(self):
        root = FakeFolder(1, "", children=[
            FakeFolder(2, "Search Root"),
            FakeFolder(3, "Top of Outlook data file", children=[FakeFolder(4, "Inbox", 5)])])
        self.assertEqual([n.path for n in self.tree(root)], [("Inbox",)])

    def test_unknown_layout_shows_everything(self):
        root = FakeFolder(1, "", children=[FakeFolder(2, "A", 1), FakeFolder(3, "B", 1)])
        self.assertEqual([n.name for n in self.tree(root)], ["A", "B"])

    def test_folder_without_a_name_gets_one_outlook_can_create(self):
        root = FakeFolder(1, "", children=[
            FakeFolder(2, "Top of Outlook data file", children=[
                FakeFolder(3, "Inbox", 1, [FakeFolder(4, None, 7), FakeFolder(5, "  ", 2)])])])
        inbox = self.tree(root)[0]
        self.assertEqual([c.name for c in inbox.children], ["Folder 1", "Folder 2"])
        self.assertEqual([c.path for c in inbox.children],
                         [("Inbox", "Folder 1"), ("Inbox", "Folder 2")])
        for n in inbox.walk():
            self.assertTrue(all(part.strip() for part in n.path), n.path)


class TempDirTest(unittest.TestCase):
    """Runs with the system temporary folder pointed at an empty one."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = tmp.name
        self.temp = os.path.join(self.dir, "temp")
        os.mkdir(self.temp)
        patch = mock.patch.object(tempfile, "tempdir", self.temp)
        patch.start()
        self.addCleanup(patch.stop)

    def write(self, name, data):
        path = os.path.join(self.dir, name)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(data)
        return path


class Recovery(TempDirTest):
    def test_running_conversion_is_left_alone(self):
        run = converter._Run(os.path.join(self.dir, "out.pst"))
        self.addCleanup(run.close)
        self.assertEqual(recover_interrupted(), [])
        self.assertTrue(os.path.isfile(os.path.join(run.dir, converter.RUN_MARKER)))

    def test_crashed_conversion_is_cleaned_up_once(self):
        pst = os.path.join(self.dir, "out.pst")
        run = converter._Run(pst)
        with open(os.path.join(run.dir, "mail.ost"), "wb") as fh:
            fh.write(b"a copy of someone's mailbox")
        run._marker.close()  # the process died: nothing holds the marker any more
        self.assertEqual(recover_interrupted(), [pst])
        self.assertFalse(os.path.exists(run.dir))
        self.assertEqual(recover_interrupted(), [])

    def test_finished_conversion_leaves_nothing(self):
        run = converter._Run(os.path.join(self.dir, "out.pst"))
        run.close()
        self.assertEqual(os.listdir(self.temp), [])

    def test_folders_that_are_not_ours_are_left_alone(self):
        other = os.path.join(self.temp, "unbound_dns_cache")
        os.mkdir(other)
        with open(os.path.join(other, "data"), "w") as fh:
            fh.write("x")
        self.assertEqual(recover_interrupted(), [])
        self.assertTrue(os.path.isfile(os.path.join(other, "data")))

    def test_unreadable_marker_still_gets_cleaned_up(self):
        folder = os.path.join(self.temp, "unbound_abc")
        os.mkdir(folder)
        with open(os.path.join(folder, converter.RUN_MARKER), "w") as fh:
            fh.write("{not json")
        self.assertEqual(recover_interrupted(), [])
        self.assertFalse(os.path.exists(folder))


class WorkingCopy(TempDirTest):
    def test_copies_and_reports_progress(self):
        source = self.write("mail.ost", os.urandom(50_000))
        seen = []
        with mock.patch.object(converter, "COPY_CHUNK", 16_384):
            copy = converter._working_copy(source, self.temp, lambda done, size: seen.append((done, size)))
        self.assertEqual(os.path.dirname(copy), self.temp)
        with open(source, "rb") as a, open(copy, "rb") as b:
            self.assertEqual(a.read(), b.read())
        self.assertEqual(seen[-1], (50_000, 50_000))
        self.assertEqual(len(seen), 4)

    def test_can_be_cancelled_partway(self):
        source = self.write("mail.ost", os.urandom(50_000))
        calls = []

        def cancel_on_third_look():
            calls.append(1)
            return len(calls) >= 3

        with mock.patch.object(converter, "COPY_CHUNK", 16_384):
            with self.assertRaises(converter._Cancelled):
                converter._working_copy(source, self.temp, None, cancel_on_third_look)

    def test_status_text_changes_once_per_percent(self):
        seen = []
        report = converter._copy_reporter(lambda done, total, text: seen.append((done, total, text)), 10)
        for copied in range(0, 1001):
            report(copied, 1000)
        self.assertEqual(len(seen), 100)
        self.assertEqual(seen[0], (0, 10, "Making a working copy of the OST… 0%"))
        self.assertTrue(seen[-1][2].endswith("99%"))

    def test_outlook_data_folder_is_recognised(self):
        with mock.patch.dict(os.environ, {"LOCALAPPDATA": self.dir}):
            inside = os.path.join(self.dir, "Microsoft", "Outlook", "me.ost")
            self.assertTrue(is_outlook_data_file(inside))
            self.assertTrue(is_outlook_data_file(inside.upper()))
            self.assertFalse(is_outlook_data_file(os.path.join(self.dir, "Documents", "me.ost")))
            self.assertFalse(is_outlook_data_file(os.path.join(self.dir, "Microsoft", "Outlook2", "x")))


class Converting(TempDirTest):
    """The conversion loop, with an Apple Mail export as the source and a
    stand-in for Outlook."""

    def setUp(self):
        super().setUp()
        self.export = os.path.join(self.dir, "export")
        self.write("export/Inbox.mbox/mbox", mbox(
            b"Subject: one\n\nx\n", b"Subject: two\n\nx\n", b"Subject: three\n\nx\n"))
        self.write("export/Work.mbox/Alpha.mbox/mbox", mbox(b"Subject: four\n\nx\n"))
        self.pst = os.path.join(self.dir, "out.pst")
        self.nodes = [n for top in scan_folders(self.export) for n in top.walk()]

    def test_copies_everything_and_tidies_up(self):
        writer = FakeWriter()
        progress = []
        result = convert(self.export, self.pst, self.nodes, outlook_writer=writer,
                         progress=lambda *a: progress.append(a))
        self.assertEqual(result.copied[MAIL], 4)
        self.assertEqual(result.failures, [])
        self.assertFalse(result.cancelled)
        self.assertEqual(writer.messages, [(("Inbox",), "one"), (("Inbox",), "two"),
                                           (("Inbox",), "three"), (("Work", "Alpha"), "four")])
        self.assertEqual(progress[-1], (4, 4, "Work/Alpha"))
        self.assertEqual(result.item_types["IPM.Note"], 4)
        self.assertEqual(os.listdir(self.temp), [])  # no temporary files left behind

    def test_one_bad_item_does_not_stop_the_run(self):
        result = convert(self.export, self.pst, self.nodes, outlook_writer=FakeWriter(fail_on={"two"}))
        self.assertEqual(result.copied[MAIL], 3)
        self.assertEqual(result.failures, [Failure("Inbox", 1, "two", "Outlook said no")])

    def test_cancel_stops_between_items_and_tidies_up(self):
        writer = FakeWriter()
        result = convert(self.export, self.pst, self.nodes, outlook_writer=writer,
                         should_cancel=lambda: len(writer.messages) >= 2)
        self.assertTrue(result.cancelled)
        self.assertEqual(result.copied[MAIL], 2)
        self.assertEqual(os.listdir(self.temp), [])

    def test_temporary_files_are_removed_when_outlook_fails_to_start(self):
        with mock.patch("unbound.pst_writer.PstWriter", side_effect=RuntimeError("no Outlook")):
            with self.assertRaises(RuntimeError):
                convert(self.export, self.pst, self.nodes)
        self.assertEqual(os.listdir(self.temp), [])

    def test_size_estimate(self):
        inbox = next(n for n in self.nodes if n.name == "Inbox")
        self.assertEqual(estimated_pst_size(self.export, [inbox], self.nodes),
                         os.path.getsize(inbox.identifier))
        ost = self.write("mail.ost", b"x" * 1000)
        nodes = [node("a", 25), node("b", 75)]
        self.assertEqual(estimated_pst_size(ost, nodes[:1], nodes), 250)
        self.assertEqual(estimated_pst_size(ost, [], []), 0)
        self.assertEqual(estimated_pst_size(os.path.join(self.dir, "gone.ost"), nodes, nodes), 0)


class Logs(TempDirTest):
    def read(self, path):
        with open(path, encoding="utf-8") as fh:
            return fh.read()

    def test_log_names_the_version_and_what_happened(self):
        result = Result()
        result.copied[MAIL] = 2
        result.failures.append(Failure("Inbox", 4, "Subject ☕", "broke"))
        result.stopped = "The file's reader stopped too many times."
        log = os.path.join(self.dir, "x - Unbound log.txt")
        write_log(result, "in.ost", "out.pst", log)
        text = self.read(log)
        self.assertTrue(text.startswith(f"Unbound {__version__} conversion log\n"))
        self.assertIn("Stopped early: The file's reader stopped too many times.", text)
        self.assertIn("Copied: 2 (mail 2, calendar 0, contacts 0)", text)
        self.assertIn("Inbox | 5 | Subject ☕ | broke", text)

    def test_failure_log(self):
        log = os.path.join(self.dir, "x - Unbound log.txt")
        write_failure_log("Couldn't start Outlook.", "in.ost", "out.pst", log)
        text = self.read(log)
        self.assertIn(f"Unbound {__version__}", text)
        self.assertIn("Conversion failed: Couldn't start Outlook.", text)


if __name__ == "__main__":
    unittest.main()
