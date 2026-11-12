"""The window, driven without a screen."""

import os
import tempfile
import unittest
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QEvent, Qt  # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402

from unbound import app as ui  # noqa: E402
from unbound.converter import PST_SIZE_LIMIT, FolderNode  # noqa: E402


def sample_tree():
    sub = FolderNode(2, "Projects", ("Inbox", "Projects"), 4, "IPF.Note")
    return [FolderNode(1, "Inbox", ("Inbox",), 10, "IPF.Note", [sub]),
            FolderNode(3, "Calendar", ("Calendar",), 6, "IPF.Appointment"),
            FolderNode(4, "Tasks", ("Tasks",), 2, "IPF.Task"),
            FolderNode(5, "Sync Issues", ("Sync Issues",), 1, "IPF.Note")]


class Helpers(unittest.TestCase):
    def test_durations(self):
        self.assertEqual([ui._duration(s) for s in (5, 65, 3725)], ["5s", "1m 05s", "1h 02m"])

    def test_sizes(self):
        self.assertEqual(ui._size(5 * 2**20), "5 MB")
        self.assertEqual(ui._size(100), "1 MB")
        self.assertEqual(ui._size(int(2.5 * 2**30)), "2.5 GB")

    def test_same_file(self):
        with tempfile.TemporaryDirectory() as folder:
            a = os.path.join(folder, "Mail.pst")
            with open(a, "w"):
                pass
            self.assertTrue(ui._same_file(a, os.path.join(folder, ".", "MAIL.PST")))
            self.assertFalse(ui._same_file(a, os.path.join(folder, "other.pst")))
            missing = os.path.join(folder, "new.pst")
            self.assertTrue(ui._same_file(missing, missing.upper()))


class Window(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qapp = QApplication.instance() or QApplication([])

    def setUp(self):
        self.win = ui.MainWindow()
        self.addCleanup(self.dispose)
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = tmp.name

    def dispose(self):
        # Destroy the window now. Left to the end of the run, Qt objects are
        # torn down in no particular order as Python exits, which can crash
        # the test process after every test has passed.
        self.win.thread = self.win.worker = None
        self.win.close()
        self.win.deleteLater()
        del self.win
        QApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        QApplication.processEvents()

    def load(self):
        self.win.ost_path = os.path.join(self.dir, "mail.ost")
        with open(self.win.ost_path, "wb") as fh:
            fh.write(b"x" * 2200)
        self.win.on_scanned(sample_tree())

    def names(self):
        return [n.name for n in self.win.selected_nodes()]

    def top(self, name):
        return next(self.win.tree.topLevelItem(i) for i in range(self.win.tree.topLevelItemCount())
                    if self.win.tree.topLevelItem(i).text(0) == name)

    def test_starts_with_nothing_to_convert(self):
        self.assertFalse(self.win.convert_btn.isEnabled())
        self.assertIn(ui.__version__, self.win.findChild(ui.QLabel, "muted").text())

    def test_ticks(self):
        self.load()
        self.assertEqual(self.names(), ["Inbox", "Projects", "Calendar"])  # not Tasks, not Sync Issues
        self.assertTrue(self.win.convert_btn.isEnabled())
        self.assertIn("3 folders selected · 20 items", self.win.summary.text())
        self.assertFalse(self.top("Tasks").flags() & Qt.ItemIsUserCheckable)

        self.top("Inbox").setCheckState(0, Qt.Unchecked)  # takes its subfolder with it
        self.assertEqual(self.names(), ["Calendar"])
        self.win.set_all(False)
        self.assertEqual(self.names(), [])
        self.assertFalse(self.win.convert_btn.isEnabled())
        self.win.set_all(True)
        self.assertEqual(self.names(), ["Inbox", "Projects", "Calendar"])

    def start(self, pst):
        """Press Convert with `pst` chosen in the save dialog; returns the
        worker that was started, or None."""
        with mock.patch.object(ui.QFileDialog, "getSaveFileName", return_value=(pst, "")), \
                mock.patch.object(self.win, "run_in_thread") as run:
            self.win.start_convert()
        return run.call_args[0][0] if run.called else None

    def test_convert_starts_for_a_new_file(self):
        self.load()
        pst = os.path.join(self.dir, "out.pst")
        worker = self.start(pst)
        self.assertEqual((worker.ost, worker.pst), (self.win.ost_path, pst))
        self.assertEqual([n.name for n in worker.nodes], ["Inbox", "Projects", "Calendar"])

    def test_source_cannot_be_the_output(self):
        self.load()
        with mock.patch.object(QMessageBox, "warning") as warning:
            self.assertIsNone(self.start(self.win.ost_path.upper()))
        self.assertIn("different file", warning.call_args[0][2])
        self.assertTrue(os.path.exists(self.win.ost_path))

    def test_output_cannot_go_in_outlooks_own_folder(self):
        self.load()
        with mock.patch.dict(os.environ, {"LOCALAPPDATA": self.dir}), \
                mock.patch.object(QMessageBox, "warning") as warning:
            target = os.path.join(self.dir, "Microsoft", "Outlook", "archive.pst")
            os.makedirs(os.path.dirname(target))
            with open(target, "w") as fh:
                fh.write("someone's real mail")
            self.assertIsNone(self.start(target))
        self.assertIn("Outlook keeps its own data files", warning.call_args[0][2])
        self.assertTrue(os.path.exists(target))

    def test_existing_file_goes_to_the_recycle_bin(self):
        self.load()
        pst = os.path.join(self.dir, "out.pst")
        with open(pst, "w") as fh:
            fh.write("old")
        with mock.patch.object(ui.QFile, "moveToTrash", return_value=True) as trash:
            self.assertIsNotNone(self.start(pst))
        trash.assert_called_once_with(pst)

    def test_permanent_delete_is_asked_about_when_there_is_no_recycle_bin(self):
        self.load()
        pst = os.path.join(self.dir, "out.pst")
        with open(pst, "w") as fh:
            fh.write("old")
        with mock.patch.object(ui.QFile, "moveToTrash", return_value=False):
            with mock.patch.object(QMessageBox, "question", return_value=QMessageBox.No):
                self.assertIsNone(self.start(pst))
            self.assertTrue(os.path.exists(pst))
            with mock.patch.object(QMessageBox, "question", return_value=QMessageBox.Yes):
                self.assertIsNotNone(self.start(pst))
            self.assertFalse(os.path.exists(pst))

    def test_warns_when_the_pst_probably_will_not_fit(self):
        self.load()
        pst = os.path.join(self.dir, "out.pst")
        nodes = self.win.selected_nodes()
        self.assertIsNone(self.win._room_problem(pst, nodes))
        with mock.patch.object(ui, "estimated_pst_size", return_value=PST_SIZE_LIMIT + 1):
            self.assertIn("50 GB", self.win._room_problem(pst, nodes))
            with mock.patch.object(QMessageBox, "question", return_value=QMessageBox.No):
                self.assertIsNone(self.start(pst))
        usage = mock.Mock(free=1000)
        with mock.patch.object(ui.shutil, "disk_usage", return_value=usage):
            self.assertIn("may not fit", self.win._room_problem(pst, nodes))

    def test_closing_while_idle(self):
        self.assertTrue(self.win.close())

    def test_closing_during_a_conversion_waits_for_it_to_tidy_up(self):
        worker = ui.ConvertWorker("in.ost", "out.pst", [])
        self.win.thread, self.win.worker = mock.Mock(), worker
        with mock.patch.object(QMessageBox, "question", return_value=QMessageBox.No):
            self.assertFalse(self.win.close())
        self.assertFalse(worker._cancel)  # "No": carries on converting

        with mock.patch.object(QMessageBox, "question", return_value=QMessageBox.Yes), \
                mock.patch.object(ui.QTimer, "singleShot") as later, \
                mock.patch.object(ui, "kill_readers") as kill:
            self.assertFalse(self.win.close())  # not yet: the PST is still attached
            self.assertFalse(self.win.close())  # asking again changes nothing
        self.assertTrue(worker._cancel)
        kill.assert_not_called()                # the reader gets time to finish its item
        self.assertEqual(later.call_args[0][1], kill)

        with mock.patch.object(QMessageBox, "information") as box, \
                mock.patch.object(QMessageBox, "critical") as error:
            self.win.on_convert_failed("ignored while closing")
        box.assert_not_called()
        error.assert_not_called()
        self.win._thread_finished()             # the conversion has stopped and tidied up
        self.assertFalse(self.win.isVisible())

    def test_file_given_while_busy_is_loaded_afterwards(self):
        self.win.thread, self.win.worker = mock.Mock(), mock.Mock()
        self.win.load_file("later.ost")
        self.assertIsNone(self.win.ost_path)
        with mock.patch.object(self.win, "run_in_thread") as run:
            self.win._thread_finished()
        self.assertEqual(run.call_args[0][0].path, os.path.abspath("later.ost"))

    def test_says_which_programs_have_the_file_open(self):
        self.win.ost_path = os.path.join(self.dir, "mail.ost")
        self.win.worker = mock.Mock(in_use_by=["Microsoft Excel", "Microsoft Word"])
        self.win.on_scanned(sample_tree())
        self.assertIn("Microsoft Excel and Microsoft Word have this file open",
                      self.win.status.text())
        self.assertTrue(self.win.convert_btn.isEnabled())  # and it can still be converted

    def scan_fails(self, outlook_has_it, click):
        """A scan fails; the user clicks the button with this label. Returns
        the labels of the buttons offered and the worker started, if any."""
        self.win.ost_path = os.path.join(self.dir, "mail.ost")
        self.win.worker = mock.Mock(outlook_has_it=outlook_has_it)
        offered = []

        def exec_(box):
            offered.extend(b.text() for b in box.buttons())
            box.clickedButton = lambda: next((b for b in box.buttons() if b.text() == click), None)
            return 0

        with mock.patch.object(QMessageBox, "exec", exec_), \
                mock.patch.object(self.win, "run_in_thread") as run:
            self.win.on_scan_failed("Outlook is using this file")
        return offered, (run.call_args[0][0] if run.called else None)

    def test_offers_to_close_outlook_when_outlook_has_the_file(self):
        offered, worker = self.scan_fails(True, "Close Outlook and try again")
        self.assertIn("Close Outlook and try again", offered)
        self.assertIsInstance(worker, ui.CloseOutlookWorker)
        self.assertEqual(worker.path, os.path.join(self.dir, "mail.ost"))
        self.win.on_outlook_closed(worker.path)   # Outlook has gone:
        with mock.patch.object(self.win, "run_in_thread") as run:
            self.win.thread = mock.Mock()
            self.win._thread_finished()           # the file is read again
        self.assertIsInstance(run.call_args[0][0], ui.ScanWorker)

    def test_outlook_is_left_alone_unless_asked(self):
        offered, worker = self.scan_fails(True, "OK")
        self.assertIsNone(worker)
        offered, worker = self.scan_fails(False, "OK")
        self.assertNotIn("Close Outlook and try again", offered)
        self.assertIsNone(worker)


class ClosingOutlook(unittest.TestCase):
    def run_worker(self, still_has_it):
        """Run the worker with a stand-in for Outlook; `still_has_it` is what
        each successive check of the file reports."""
        outlook, done, failed = mock.Mock(), [], []
        worker = ui.CloseOutlookWorker("C:\\mail.ost")
        worker.done.connect(done.append)
        worker.failed.connect(failed.append)
        checks = iter(still_has_it)
        clock = iter(range(0, 10_000, 30))
        with mock.patch("win32com.client.GetActiveObject", return_value=outlook), \
                mock.patch.object(worker, "_outlook_has_it", lambda: next(checks)), \
                mock.patch.object(ui.time, "sleep"), \
                mock.patch.object(ui.time, "time", lambda: next(clock)):
            worker.run()
        return outlook, done, failed

    def test_closes_outlook_and_waits_for_the_file(self):
        outlook, done, failed = self.run_worker([True, True, False])
        outlook.Quit.assert_called_once_with()
        self.assertEqual((done, failed), (["C:\\mail.ost"], []))

    def test_gives_up_if_outlook_stays(self):
        outlook, done, failed = self.run_worker([True] * 50)
        self.assertEqual(done, [])
        self.assertIn("Task Manager", failed[0])


if __name__ == "__main__":
    unittest.main()
