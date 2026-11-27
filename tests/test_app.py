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


class WindowCase(unittest.TestCase):
    """Set-up shared by the window tests: a fresh window per test."""

    @classmethod
    def setUpClass(cls):
        cls.qapp = QApplication.instance() or QApplication([])

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = tmp.name
        # The same on every PC (CI has no Outlook), and nothing written to the
        # registry: remembered folders go to a throwaway file
        settings_file = os.path.join(self.dir, "settings.ini")
        for target, value in (("outlook_installed", lambda: True),
                              ("_settings", lambda: ui.QSettings(settings_file, ui.QSettings.IniFormat))):
            patcher = mock.patch.object(ui, target, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        self.win = ui.MainWindow()
        self.addCleanup(self.dispose)

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

    def start(self, pst):
        """Press Convert with `pst` chosen in the save dialog; returns the
        worker that was started, or None."""
        with mock.patch.object(ui.QFileDialog, "getSaveFileName", return_value=(pst, "")), \
                mock.patch.object(self.win, "run_in_thread") as run:
            self.win.start_convert()
        return run.call_args[0][0] if run.called else None


class Window(WindowCase):
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


class Improvements(WindowCase):
    """The quick improvements from the 1.3.3 review (QW1-QW12)."""

    def test_pst_can_be_chosen_without_the_all_files_trick(self):  # QW1
        self.assertEqual([a.text() for a in self.win.browse_btn.menu().actions()][0],
                         "Outlook data file (OST or PST)…")
        with mock.patch.object(ui.QFileDialog, "getOpenFileName", return_value=("", "")) as dialog:
            self.win.choose_file()
        self.assertIn("*.pst", dialog.call_args[0][3])

    def finished(self, **result_fields):
        from unbound.converter import Result
        result = Result(**result_fields)
        self.win.pst_path = os.path.join(self.dir, "out.pst")
        shown = {}

        def capture(box):
            shown["text"], shown["icon"] = box.informativeText(), box.icon()
            shown["headline"] = box.text()
            return 0
        with mock.patch.object(QMessageBox, "exec", capture):
            self.win.on_converted(result, "")
        return shown

    def test_attached_emails_left_out_are_mentioned(self):  # QW2
        shown = self.finished(skipped_attached_emails=3)
        self.assertIn("Emails attached to other emails, not copied yet: 3", shown["text"])
        self.assertEqual(shown["icon"], QMessageBox.Warning)
        self.assertEqual(self.finished()["icon"], QMessageBox.Information)

    def test_a_pst_left_attached_to_outlook_is_explained(self):
        shown = self.finished(still_attached="Outlook stopped responding")
        self.assertEqual(shown["icon"], QMessageBox.Warning)
        self.assertIn("Outlook stopped responding.", shown["headline"])
        self.assertIn("choose Close", shown["headline"])

    def test_plain_english_when_a_file_cannot_be_opened(self):  # QW3
        empty = os.path.join(self.dir, "empty.ost")
        open(empty, "wb").close()
        self.assertIn("empty", ui.explain_open_error(empty, OSError("unable to read")))
        junk = os.path.join(self.dir, "junk.pst")
        with open(junk, "wb") as fh:
            fh.write(b"junk" * 500)
        self.assertIn("damaged or incomplete", ui.explain_open_error(
            junk, OSError("libpff_file_header_read: unable to read file header data")))
        self.assertIn("isn't an Outlook data file", ui.explain_open_error(
            junk, OSError("libpff: invalid file signature")))
        worker = ui.ScanWorker(junk)
        failed = []
        worker.failed.connect(failed.append)
        with mock.patch.object(ui, "scan_folders", side_effect=OSError("pypff: unable to read x")):
            worker.run()
        self.assertNotIn("pypff", failed[0])
        self.assertIn("pypff", worker.details)

    def test_progress_while_outlook_starts_and_during_the_working_copy(self):  # QW5
        self.load()
        with mock.patch.object(ui.QFileDialog, "getSaveFileName",
                               return_value=(os.path.join(self.dir, "o.pst"), "")), \
                mock.patch.object(self.win, "run_in_thread"):
            self.win.start_convert()
        self.assertEqual(self.win.progress.maximum(), 0)  # Qt's moving "busy" bar
        self.win.on_progress(0, 20, "Making a working copy of the OST… 40%")
        self.assertEqual((self.win.progress.maximum(), self.win.progress.value()), (100, 40))
        self.win.on_progress(0, 20, "Starting Outlook…")
        self.assertEqual(self.win.progress.maximum(), 0)
        with mock.patch.object(QMessageBox, "critical"):
            self.win.on_convert_failed("no")
        self.assertGreater(self.win.progress.maximum(), 0)  # no longer animating

    def test_bar_shows_how_far_a_cancelled_run_got(self):
        """The "Detaching…" step shows the busy bar; afterwards the bar went
        back to nothing, as if nothing had been copied."""
        from unbound.converter import Result
        self.load()
        with mock.patch.object(ui.QFileDialog, "getSaveFileName",
                               return_value=(os.path.join(self.dir, "o.pst"), "")), \
                mock.patch.object(self.win, "run_in_thread"):
            self.win.start_convert()
        total = self.win.total_items
        self.win.on_progress(7, total, "Inbox")
        self.win.on_progress(0, total, "Detaching the PST from Outlook…")
        with mock.patch.object(QMessageBox, "exec", lambda box: 0):
            self.win.on_converted(Result(cancelled=True), "")
        self.assertEqual((self.win.progress.maximum(), self.win.progress.value()), (total, 7))

    def test_time_left_ignores_set_up_time(self):  # QW8
        self.win.items_started, self.win.total_items = None, 100
        clock = iter([1000.0, 1010.0])
        with mock.patch.object(ui.time, "time", lambda: next(clock)):
            self.win.on_progress(1, 100, "Inbox")
            self.win.on_progress(11, 100, "Inbox")  # 10 items in 10 s
        self.assertIn("about 1m 29s left", self.win.status.text())  # 89 left at 1 a second

    def test_save_starts_in_the_documents_folder_windows_reports(self):  # QW6
        self.load()
        with mock.patch.object(ui.QStandardPaths, "writableLocation", return_value=self.dir), \
                mock.patch.object(ui.QFileDialog, "getSaveFileName", return_value=("", "")) as dialog:
            self.win.start_convert()
        self.assertEqual(os.path.dirname(dialog.call_args[0][2]), self.dir)

    def test_button_names_what_it_does(self):  # QW7
        self.assertEqual(self.win.all_btn.text().replace("&", ""), "Reset selection")

    def test_keyboard_shortcuts(self):  # QW9
        self.assertEqual(self.win.browse_btn.shortcut().toString(), "Alt+B")
        self.assertEqual(self.win.convert_btn.shortcut().toString(), "Alt+C")
        self.assertEqual(self.win.tree.accessibleName(), "Folders to convert")

    def test_last_folders_are_remembered(self):  # QW10
        source = os.path.join(self.dir, "mail.ost")
        open(source, "wb").close()
        with mock.patch.object(ui.QFileDialog, "getOpenFileName", return_value=(source, "")), \
                mock.patch.object(self.win, "load_file"):
            self.win.choose_file()
        with mock.patch.object(ui.QFileDialog, "getOpenFileName", return_value=("", "")) as dialog:
            self.win.choose_file()
        self.assertEqual(os.path.normcase(dialog.call_args[0][2]), os.path.normcase(self.dir))

    def test_missing_outlook_is_said_up_front(self):  # QW11
        with mock.patch.object(ui, "outlook_installed", lambda: False):
            win = ui.MainWindow()
        try:
            self.assertIn("Outlook Classic isn't installed", win.status.text())
            win.ost_path = os.path.join(self.dir, "mail.ost")
            open(win.ost_path, "wb").close()
            win.on_scanned(sample_tree())
            with mock.patch.object(QMessageBox, "warning") as warning, \
                    mock.patch.object(win, "run_in_thread") as run:
                win.start_convert()
            self.assertFalse(run.called)
            self.assertIn("Outlook Classic", warning.call_args[0][2])
        finally:
            win.thread = win.worker = None
            win.close()
            win.deleteLater()
            QApplication.sendPostedEvents(None, QEvent.DeferredDelete)


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

    def test_reaches_an_outlook_running_in_the_background(self):
        """An Outlook started without a window (as Unbound starts it) isn't
        in Windows' table of running programs: GetActiveObject fails with
        "Operation unavailable". It was the very case the button is for."""
        outlook, done = mock.Mock(), []
        worker = ui.CloseOutlookWorker("C:\\mail.ost")
        worker.done.connect(done.append)
        with mock.patch("win32com.client.GetActiveObject",
                        side_effect=OSError("Operation unavailable")), \
                mock.patch("win32com.client.Dispatch", return_value=outlook) as dispatch, \
                mock.patch.object(worker, "_outlook_has_it", lambda: False):
            worker.run()
        dispatch.assert_called_once_with("Outlook.Application")
        outlook.Quit.assert_called_once_with()
        self.assertEqual(done, ["C:\\mail.ost"])


class CheckingForUpdates(WindowCase):
    """The Check for updates button, with GitHub and the download stood in."""

    def setUp(self):
        super().setUp()
        from unbound import updates
        self.updates = updates
        self.exe = os.path.join(self.dir, "Unbound.exe")
        with open(self.exe, "wb") as fh:
            fh.write(b"this version")
        self.choices = []  # the button to click in each message box, in turn
        for patcher in (mock.patch.object(self.win, "run_in_thread", self.run_now),
                        mock.patch.object(ui, "_show_in_folder"),
                        mock.patch.object(ui.QDesktopServices, "openUrl"),
                        mock.patch.object(QMessageBox, "exec", lambda box: 0),
                        mock.patch.object(QMessageBox, "clickedButton", lambda box: self.click(box))):
            patcher.start()
            self.addCleanup(patcher.stop)

    def click(self, box):
        """The next choice for a new box; the same answer when asked again."""
        if box.property("clicked") is None:
            box.setProperty("clicked", self.choices.pop(0))
        return next(b for b in box.buttons() if b.text() == box.property("clicked"))

    def run_now(self, worker, on_done, on_failed, progress=None):
        """Run the job straight away, then finish as the real thread does."""
        win = self.win
        worker.done.connect(on_done)
        worker.failed.connect(on_failed)
        if progress:
            worker.progress.connect(progress)
        win.thread = win.worker = worker
        worker.run()
        win.thread = win.worker = None
        win.update_controls()
        if win._after_job:
            job, win._after_job = win._after_job, None
            job()

    def release(self, version):
        base = "https://github.com/Z3R0ZIGMA/Unbound/releases/download/v{}/".format(version)
        return self.updates.Release(self.updates.parse_version(version), f"v{version}",
                                    f"https://github.com/Z3R0ZIGMA/Unbound/releases/tag/v{version}",
                                    base + "Unbound.exe", base + "SHA256SUMS.txt")

    def check(self, latest, exe=True, download=None):
        def fake_download(release, target, progress=None, cancelled=None):
            if download:
                raise download
            progress(100)
            with open(target, "wb") as fh:
                fh.write(b"new version")
        with mock.patch.object(self.updates, "get_latest", return_value=latest), \
                mock.patch.object(self.updates, "running_exe", return_value=self.exe if exe else None), \
                mock.patch.object(self.updates, "download", fake_download):
            self.win.update_btn.click()

    def test_up_to_date(self):
        self.check(self.release(ui.__version__))
        self.assertIn(f"up to date ({ui.__version__})", self.win.status.text())
        self.assertTrue(self.win.update_btn.isEnabled())

    def test_nothing_to_update_to(self):
        self.check(None)
        self.assertIn("Couldn't find a release", self.win.status.text())

    def test_github_unreachable(self):
        with mock.patch.object(self.updates, "get_latest",
                               side_effect=self.updates.UpdateError("Couldn't reach GitHub.")):
            self.win.update_btn.click()
        self.assertEqual(self.win.status.text(), "Couldn't update: Couldn't reach GitHub.")
        self.assertTrue(self.win.update_btn.isEnabled())

    def test_run_from_source_opens_the_release_page(self):
        with mock.patch.object(ui.QDesktopServices, "openUrl") as opened:
            self.check(self.release("99.0.0"), exe=False)
        self.assertIn("/releases/tag/v99.0.0", opened.call_args[0][0].toString())
        self.assertIn("99.0.0 is available", self.win.status.text())

    def test_not_now(self):
        self.choices = ["Not now"]
        self.check(self.release("99.0.0"))
        self.assertIn("click Check for updates when you're ready", self.win.status.text())
        self.assertFalse(os.path.exists(self.exe + ".new"))

    def test_later_swaps_it_in_when_unbound_closes(self):
        """The running copy stays as it is until Unbound closes, so this
        session never starts its file reader from the new version."""
        self.choices = ["Download", "Later"]
        self.check(self.release("99.0.0"))
        with open(self.exe, "rb") as fh:
            self.assertEqual(fh.read(), b"this version")
        self.assertFalse(self.win.update_btn.isEnabled())  # one update at a time
        with mock.patch.object(self.updates, "running_exe", return_value=self.exe), \
                mock.patch.object(self.updates, "start_new_instance") as started:
            self.win.close()
        with open(self.exe, "rb") as fh:
            self.assertEqual(fh.read(), b"new version")
        self.assertTrue(os.path.exists(self.exe + ".old"))
        started.assert_not_called()

    def test_restart_now(self):
        self.choices = ["Download", "Restart now"]
        with mock.patch.object(self.updates, "start_new_instance") as started:
            self.check(self.release("99.0.0"))
        with open(self.exe, "rb") as fh:
            self.assertEqual(fh.read(), b"new version")
        started.assert_called_once_with(self.exe)

    def test_a_bad_download_changes_nothing(self):
        self.choices = ["Download"]
        self.check(self.release("99.0.0"), download=self.updates.UpdateError(
            "The download doesn't match the release's checksum, so it wasn't used."))
        self.assertIn("checksum", self.win.status.text())
        self.assertTrue(self.win.update_btn.isEnabled())
        with open(self.exe, "rb") as fh:
            self.assertEqual(fh.read(), b"this version")

    def test_read_only_folder_saves_to_downloads(self):
        self.choices = ["Download"]
        targets = []

        def download(release, target, progress=None, cancelled=None):
            targets.append(target)
            if target.endswith(".new"):
                raise PermissionError("Access is denied")
        with mock.patch.object(self.updates, "get_latest", return_value=self.release("99.0.0")), \
                mock.patch.object(self.updates, "running_exe", return_value=self.exe), \
                mock.patch.object(self.updates, "download", download):
            self.win.update_btn.click()
        self.assertTrue(targets[1].endswith(os.path.join("Downloads", "Unbound-99.0.0.exe")))
        self.assertIn("saved to", self.win.status.text())
        ui._show_in_folder.assert_called_once_with(targets[1])
        self.assertIsNone(self.win._staged_update)

    def test_not_while_converting(self):
        self.load()
        self.win.thread = object()
        self.win.update_controls()
        self.assertFalse(self.win.update_btn.isEnabled())


class AfterACrash(unittest.TestCase):
    def test_detaching_the_old_pst_has_a_time_limit(self):
        """Without one, an Outlook that doesn't answer kept the window from
        ever being closed."""
        from unbound import pst_writer
        done, failed = [], []
        worker = ui.RecoverWorker(["C:\\old.pst"])
        worker.done.connect(done.append)
        worker.failed.connect(failed.append)
        with mock.patch.object(pst_writer, "_detach_within", return_value=["C:\\old.pst"]) as detach:
            worker.run()
        self.assertEqual(detach.call_args[0][:2], (["C:\\old.pst"], pst_writer.DETACH_TIMEOUT))
        self.assertEqual((done, failed), ([["C:\\old.pst"]], []))
        with mock.patch.object(pst_writer, "_detach_within",
                               side_effect=pst_writer.DetachTimeout("Outlook stopped responding")):
            worker.run()
        self.assertEqual(failed, ["Outlook stopped responding"])


if __name__ == "__main__":
    unittest.main()
