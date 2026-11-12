"""Unbound desktop window (PySide6)."""

import os
import shutil
import subprocess
import sys
import time

from PySide6.QtCore import QFile, QObject, QSize, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QColor, QIcon, QPalette
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QMainWindow, QMenu, QMessageBox, QProgressBar, QPushButton, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QWidget,
)

from . import __version__
from .converter import (
    CALENDAR, CONTACTS, MAIL, PST_SIZE_LIMIT, convert, estimated_pst_size,
    is_outlook_data_file, log_path_for, recover_interrupted, scan_folders,
    write_failure_log, write_log,
)
from .mbox_reader import is_mbox_source
from .ost_reader import OstInUseError
from .pff_process import kill_readers

APP_NAME = "Unbound"

# Colours in one place so they're easy to change later
COLORS = {
    "window": "#1e1f22",
    "panel": "#2b2d31",
    "text": "#e6e6e6",
    "muted": "#8b8f97",
    "accent": "#3b82f6",
    "accent_hover": "#2563eb",
    "danger": "#ef4444",
    "border": "#3a3d43",
}

FOLDER_KIND = {  # labels for folders that aren't converted yet
    "IPF.CONTACT.RECIPIENTCACHE": "address cache",
    "IPF.TASK": "tasks",
    "IPF.STICKYNOTE": "notes",
    "IPF.JOURNAL": "journal",
}

NODE_ROLE = Qt.UserRole


def apply_dark_theme(app):
    app.setStyle("Fusion")
    c = {k: QColor(v) for k, v in COLORS.items()}
    pal = QPalette()
    pal.setColor(QPalette.Window, c["window"])
    pal.setColor(QPalette.WindowText, c["text"])
    pal.setColor(QPalette.Base, c["panel"])
    pal.setColor(QPalette.AlternateBase, c["window"])
    pal.setColor(QPalette.Text, c["text"])
    pal.setColor(QPalette.Button, c["panel"])
    pal.setColor(QPalette.ButtonText, c["text"])
    pal.setColor(QPalette.Highlight, c["accent"])
    pal.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    pal.setColor(QPalette.ToolTipBase, c["panel"])
    pal.setColor(QPalette.ToolTipText, c["text"])
    pal.setColor(QPalette.PlaceholderText, c["muted"])
    for role in (QPalette.Text, QPalette.WindowText, QPalette.ButtonText):
        pal.setColor(QPalette.Disabled, role, c["muted"])
    app.setPalette(pal)
    app.setStyleSheet(f"""
        QWidget {{ font-size: 10pt; }}
        QLineEdit, QTreeWidget {{
            background: {COLORS['panel']}; border: 1px solid {COLORS['border']};
            border-radius: 6px; padding: 6px;
        }}
        QHeaderView::section {{
            background: {COLORS['window']}; color: {COLORS['muted']};
            border: none; padding: 4px 6px;
        }}
        QPushButton {{
            background: {COLORS['panel']}; border: 1px solid {COLORS['border']};
            border-radius: 6px; padding: 6px 14px;
        }}
        QPushButton:hover {{ border-color: {COLORS['accent']}; }}
        QPushButton:disabled {{ color: {COLORS['muted']}; }}
        QPushButton#convert {{
            background: {COLORS['accent']}; border: none; color: white;
            font-size: 14pt; font-weight: 600; padding: 14px; border-radius: 8px;
        }}
        QPushButton#convert:hover {{ background: {COLORS['accent_hover']}; }}
        QPushButton#convert:disabled {{ background: {COLORS['border']}; color: {COLORS['muted']}; }}
        QPushButton#cancel {{ border-color: {COLORS['danger']}; }}
        QProgressBar {{
            background: {COLORS['panel']}; border: 1px solid {COLORS['border']};
            border-radius: 6px; height: 22px; text-align: center;
        }}
        QProgressBar::chunk {{ background: {COLORS['accent']}; border-radius: 5px; }}
        QLabel#muted {{ color: {COLORS['muted']}; }}
        QLabel#title {{ font-size: 20pt; font-weight: 700; }}
    """)


class ScanWorker(QObject):
    done = Signal(list)
    failed = Signal(str)

    def __init__(self, path):
        super().__init__()
        self.path = path

    def run(self):
        try:
            if os.path.isdir(self.path) and not is_mbox_source(self.path):
                self.failed.emit("No Apple Mail mailboxes were found in this folder.\n\n"
                                 "Choose the folder you exported to from Apple Mail "
                                 "(it contains folders ending in .mbox), or one .mbox folder.")
                return
            nodes = scan_folders(self.path)
            if not nodes:
                self.failed.emit("No folders or mailboxes were found here.")
                return
            self.done.emit(nodes)
        except OstInUseError as e:
            self.failed.emit(str(e))
        except Exception as e:
            self.failed.emit("This couldn't be opened as an Outlook data file or an MBOX export.\n\n"
                             + str(e).splitlines()[0][:400])


class ConvertWorker(QObject):
    progress = Signal(int, int, str)
    done = Signal(object, str)
    failed = Signal(str)

    def __init__(self, ost, pst, nodes):
        super().__init__()
        self.ost, self.pst, self.nodes = ost, pst, nodes
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def _fail(self, msg):
        # Keep a record next to where the PST was going, as a finished run does
        log = log_path_for(self.pst)
        try:
            write_failure_log(msg, self.ost, self.pst, log)
            msg += f"\n\nThis was also saved to:\n{log}"
        except OSError:
            pass
        self.failed.emit(msg)

    def run(self):
        import pythoncom  # COM must be initialised on the thread that uses it
        pythoncom.CoInitialize()
        try:
            try:
                result = convert(self.ost, self.pst, self.nodes,
                                 progress=lambda d, t, f: self.progress.emit(d, t, f),
                                 should_cancel=lambda: self._cancel)
            except OstInUseError as e:
                self._fail(str(e))
                return
            except Exception as e:
                msg = str(e).splitlines()[0][:400]
                if "Outlook.Application" in msg or "Invalid class string" in msg:
                    msg = ("Couldn't start Outlook. Unbound needs Outlook Classic "
                           "installed on this PC (the New Outlook can't write PST files).")
                self._fail(msg)
                return
            log = log_path_for(self.pst)
            try:
                write_log(result, self.ost, self.pst, log)
            except OSError:
                log = ""  # the PST is still fine; say the log couldn't be saved
            self.done.emit(result, log)
        finally:
            pythoncom.CoUninitialize()


class RecoverWorker(QObject):
    """Detaches the PSTs that a conversion which was interrupted (a crash, a
    power cut) left attached to the Outlook profile."""
    done = Signal(list)
    failed = Signal(str)

    def __init__(self, psts):
        super().__init__()
        self.psts = psts

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            from .pst_writer import detach_stores
            self.done.emit(detach_stores(self.psts))
        except Exception as e:
            self.failed.emit(str(e).splitlines()[0][:300])
        finally:
            pythoncom.CoUninitialize()


class MainWindow(QMainWindow):
    def __init__(self, interrupted=()):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} – Convert to PST")
        self.resize(760, 680)
        self.setAcceptDrops(True)  # a file or folder can be dropped on the window
        self.ost_path = None
        self.pst_path = None
        self.started = 0.0
        self.thread = None
        self.worker = None
        self._syncing = False
        self._closing = False
        self._pending_path = None  # a file to load once the current job is done
        self._interrupted = list(interrupted)

        root = QWidget()
        self.setCentralWidget(root)
        lay = QVBoxLayout(root)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(12)

        # Header: triangle logo, name, one-line description
        title_row = QHBoxLayout()
        title_row.setSpacing(12)
        if os.path.exists(_icon_path()):
            logo = QLabel()
            logo.setObjectName("logo")
            size = 40
            # pixmap at the screen's scale so it stays sharp on high-DPI displays
            logo.setPixmap(QIcon(_icon_path()).pixmap(QSize(size, size), self.devicePixelRatioF()))
            logo.setFixedSize(size, size)
            title_row.addWidget(logo)
        title = QLabel(APP_NAME)
        title.setObjectName("title")
        title_row.addWidget(title)
        title_row.addStretch(1)
        version = QLabel(f"Version {__version__}")
        version.setObjectName("muted")
        title_row.addWidget(version, 0, Qt.AlignBottom)
        subtitle = QLabel("Convert an Outlook OST or an Apple Mail export (MBOX) into a PST")
        subtitle.setObjectName("muted")
        lay.addLayout(title_row)
        lay.addWidget(subtitle)

        # Top: file picker
        row = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        self.path_edit.setPlaceholderText("Choose an OST file or an Apple Mail export, or drop one here…")
        self.browse_btn = QPushButton("Browse…")
        menu = QMenu(self.browse_btn)
        menu.addAction("Outlook OST file…", self.choose_file)
        menu.addAction("Apple Mail export (folder)…", self.choose_mbox_folder)
        self.browse_btn.setMenu(menu)
        row.addWidget(self.path_edit, 1)
        row.addWidget(self.browse_btn)
        lay.addLayout(row)

        # Middle: folder tree
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Folder", "Items"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tree.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tree.itemChanged.connect(self.on_item_changed)
        lay.addWidget(self.tree, 1)

        sel_row = QHBoxLayout()
        self.summary = QLabel("No file loaded")
        self.summary.setObjectName("muted")
        self.all_btn = QPushButton("Select all mail")
        self.none_btn = QPushButton("Clear")
        self.all_btn.clicked.connect(lambda: self.set_all(True))
        self.none_btn.clicked.connect(lambda: self.set_all(False))
        sel_row.addWidget(self.summary, 1)
        sel_row.addWidget(self.all_btn)
        sel_row.addWidget(self.none_btn)
        lay.addLayout(sel_row)

        # Bottom: convert + progress
        self.convert_btn = QPushButton("Convert to PST")
        self.convert_btn.setObjectName("convert")
        self.convert_btn.clicked.connect(self.start_convert)
        lay.addWidget(self.convert_btn)

        prog_row = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setObjectName("cancel")
        self.cancel_btn.clicked.connect(self.cancel_convert)
        self.cancel_btn.hide()
        prog_row.addWidget(self.progress, 1)
        prog_row.addWidget(self.cancel_btn)
        lay.addLayout(prog_row)

        self.status = QLabel("")
        self.status.setObjectName("muted")
        self.status.setWordWrap(True)
        lay.addWidget(self.status)

        self.update_controls()
        if self._interrupted:
            self.status.setText("Cleaning up after a conversion that was interrupted…")
            self.run_in_thread(RecoverWorker(self._interrupted),
                               self.on_recovered, self.on_recover_failed)
            self.update_controls()

    # ---------- after an interrupted conversion ----------

    def on_recovered(self, detached):
        self.status.setText(
            "A conversion was interrupted last time. Its temporary files were removed"
            + (" and its PST was detached from Outlook." if detached else "."))

    def on_recover_failed(self, _msg):
        names = ", ".join(os.path.basename(p) for p in self._interrupted)
        self.status.setText(
            "A conversion was interrupted last time. Its temporary files were removed, but "
            f"{names} may still be open in Outlook: right-click it there and choose Close.")

    # ---------- file and tree ----------

    def choose_mbox_folder(self):
        # Apple Mail's Export Mailbox writes one "Name.mbox" folder per
        # mailbox, so this picks a folder: the one they were exported into
        # (or a single .mbox folder). Loose .mbox files in it work too.
        path = QFileDialog.getExistingDirectory(
            self, "Choose the folder you exported from Apple Mail")
        if path:
            self.load_file(path)

    def choose_file(self):
        start = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Outlook")
        # A PST can still be opened through "All files" (e.g. to re-save an
        # old ANSI PST as Unicode), but OST is what this is for
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose an OST file", start if os.path.isdir(start) else "",
            "Outlook OST files (*.ost);;All files (*)")
        if path:
            self.load_file(path)

    def _dropped_path(self, event):
        """The one local file or folder being dragged, or None."""
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        if len(urls) == 1 and urls[0].isLocalFile() and not self.busy():
            return urls[0].toLocalFile()
        return None

    def dragEnterEvent(self, event):
        if self._dropped_path(event):
            event.acceptProposedAction()

    def dropEvent(self, event):
        path = self._dropped_path(event)
        if path:
            event.acceptProposedAction()
            self.load_file(path)

    def load_file(self, path):
        if self.busy():
            self._pending_path = path  # e.g. still cleaning up at startup
            return
        self.ost_path = os.path.abspath(path)
        self.path_edit.setText(self.ost_path)
        self.tree.clear()
        self.summary.setText("Reading folders…")
        self.status.setText("")
        self.progress.setValue(0)
        self.run_in_thread(ScanWorker(self.ost_path), self.on_scanned, self.on_scan_failed)
        self.update_controls()

    def on_scan_failed(self, msg):
        if self._closing:
            return
        self.summary.setText("Couldn't read this file")
        QMessageBox.warning(self, APP_NAME, msg)
        self.ost_path = None
        self.update_controls()

    def on_scanned(self, top_nodes):
        if self._closing:
            return
        self.tree.blockSignals(True)
        for node in top_nodes:
            self.tree.addTopLevelItem(self._make_item(node))
        self.tree.blockSignals(False)
        for i in range(self.tree.topLevelItemCount()):
            self.tree.topLevelItem(i).setExpanded(True)
        self.update_summary()
        self.update_controls()

    def _make_item(self, node):
        item = QTreeWidgetItem([node.name or "(unnamed)", ""])
        item.setData(0, NODE_ROLE, node)
        count = f"{node.item_count:,}" if node.item_count else ""
        if node.is_supported:
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(0, Qt.Checked if node.default_checked else Qt.Unchecked)
            if node.kind != MAIL:
                count = f"{count} · {node.kind}" if count else node.kind
        else:
            # Tree items are tickable by default in Qt; unsupported folders mustn't be
            item.setFlags(item.flags() & ~Qt.ItemIsUserCheckable)
            kind = FOLDER_KIND.get((node.container_class or "").upper(), "not converted")
            count = f"{count} ({kind} – later version)" if count else f"({kind})"
            item.setForeground(0, QColor(COLORS["muted"]))
            item.setForeground(1, QColor(COLORS["muted"]))
            item.setToolTip(0, "This kind of folder isn't converted yet")
        item.setText(1, count)
        item.setTextAlignment(1, Qt.AlignRight | Qt.AlignVCenter)
        for c in node.children:
            item.addChild(self._make_item(c))
        return item

    def _items(self):
        def walk(item):
            yield item
            for i in range(item.childCount()):
                yield from walk(item.child(i))
        for i in range(self.tree.topLevelItemCount()):
            yield from walk(self.tree.topLevelItem(i))

    def _checkable(self, item):
        return bool(item.flags() & Qt.ItemIsUserCheckable)

    def on_item_changed(self, item, column):
        if self._syncing or column != 0 or not self._checkable(item):
            return
        # Ticking or unticking a folder applies to all its subfolders too
        state = item.checkState(0)
        self._syncing = True

        def apply(it):
            for i in range(it.childCount()):
                ch = it.child(i)
                if self._checkable(ch):
                    ch.setCheckState(0, state)
                apply(ch)
        apply(item)
        self._syncing = False
        self.update_summary()

    def set_all(self, on):
        self._syncing = True
        for it in self._items():
            if self._checkable(it):
                node = it.data(0, NODE_ROLE)
                it.setCheckState(0, Qt.Checked if on and node.default_checked else Qt.Unchecked)
        self._syncing = False
        self.update_summary()

    def selected_nodes(self):
        return [it.data(0, NODE_ROLE) for it in self._items()
                if self._checkable(it) and it.checkState(0) == Qt.Checked]

    def update_summary(self):
        nodes = self.selected_nodes()
        items = sum(n.item_count for n in nodes)
        self.summary.setText(f"{len(nodes)} folders selected · {items:,} items")
        self.update_controls()

    # ---------- conversion ----------

    def busy(self):
        return self.thread is not None

    def update_controls(self):
        busy = self.busy()
        has_tree = self.tree.topLevelItemCount() > 0
        self.browse_btn.setEnabled(not busy)
        self.tree.setEnabled(not busy)
        self.all_btn.setEnabled(has_tree and not busy)
        self.none_btn.setEnabled(has_tree and not busy)
        self.convert_btn.setEnabled(has_tree and not busy and bool(self.selected_nodes()))

    def _replace_existing(self, pst):
        """Make way for the new PST. The save dialog already asked to replace
        the file; Outlook would add to an existing PST instead, so it has to
        go first. It goes to the Recycle Bin, so a slip can be undone."""
        try:
            with open(pst, "ab"):
                pass  # only checks nothing else has it open
        except OSError as e:
            QMessageBox.warning(
                self, APP_NAME,
                "The existing file can't be replaced because it's in use.\n\n"
                "If it's open in Outlook, right-click it there and choose Close, then try "
                "again. After a conversion that was interrupted, Outlook can keep hold of "
                "the file for several minutes: wait a little, or save under another name."
                f"\n\n{e}")
            return False
        if QFile.moveToTrash(pst):
            return True
        if QMessageBox.question(
                self, APP_NAME,
                "The existing file can't be moved to the Recycle Bin (this drive may not "
                "have one).\n\nDelete it permanently instead?") != QMessageBox.Yes:
            return False
        try:
            os.remove(pst)
        except OSError as e:
            QMessageBox.warning(self, APP_NAME, f"Couldn't delete the existing file.\n\n{e}")
            return False
        return True

    def _room_problem(self, pst, nodes):
        """A warning if the PST probably won't fit, else None."""
        size = estimated_pst_size(self.ost_path, nodes,
                                  [it.data(0, NODE_ROLE) for it in self._items()])
        if size > PST_SIZE_LIMIT:
            return (f"This selection is roughly {_size(size)}. Outlook stops adding to a PST "
                    "at 50 GB, so the conversion would probably stop partway.\n\n"
                    "Converting fewer folders at a time, into separate PSTs, avoids that.")
        try:
            free = shutil.disk_usage(os.path.dirname(pst)).free
        except OSError:
            return None
        if free < size * 1.1:
            return (f"The PST will be roughly {_size(size)}, and the drive you chose has "
                    f"{_size(free)} free, so it may not fit.")
        return None

    def start_convert(self):
        nodes = self.selected_nodes()
        if not nodes:
            return
        base = os.path.basename(os.path.normpath(self.ost_path))
        if base == "mbox":  # the data file inside an Apple .mbox folder
            base = os.path.basename(os.path.dirname(self.ost_path))
        default = os.path.join(os.path.expanduser("~"), "Documents",
                               os.path.splitext(base)[0] + ".pst")
        pst, _ = QFileDialog.getSaveFileName(self, "Save PST as", default,
                                             "Outlook data file (*.pst)")
        if not pst:
            return
        pst = os.path.abspath(pst)
        if _same_file(pst, self.ost_path):
            QMessageBox.warning(self, APP_NAME, "Choose a different file from the source.")
            return
        if is_outlook_data_file(pst):
            # Files there can belong to the Outlook profile on this PC, and
            # replacing one of those would lose mail
            QMessageBox.warning(
                self, APP_NAME,
                "That folder is where Outlook keeps its own data files.\n\n"
                "Save the PST somewhere else, such as Documents.")
            return
        problem = self._room_problem(pst, nodes)
        if problem and QMessageBox.question(
                self, APP_NAME, problem + "\n\nConvert anyway?") != QMessageBox.Yes:
            return
        if os.path.exists(pst) and not self._replace_existing(pst):
            return

        self.pst_path = pst
        self.started = time.time()
        self.progress.setValue(0)
        self.progress.setMaximum(max(1, sum(n.item_count for n in nodes)))
        self.status.setText("Starting Outlook…")
        self.cancel_btn.show()
        self.cancel_btn.setEnabled(True)
        self.run_in_thread(ConvertWorker(self.ost_path, pst, nodes),
                           self.on_converted, self.on_convert_failed,
                           progress=self.on_progress)
        self.update_controls()

    def on_progress(self, done, total, folder):
        if self._closing:
            return
        if done == 0:  # a setup step, e.g. making a working copy of the OST
            self.status.setText(folder)
            return
        self.progress.setMaximum(max(1, total))
        self.progress.setValue(done)
        elapsed = time.time() - self.started
        eta = ""
        if done >= 5 and elapsed > 3:
            left = elapsed / done * (total - done)
            eta = f" · about {_duration(left)} left"
        self.status.setText(f"{done:,} of {total:,} · {folder}{eta}")

    def cancel_convert(self):
        if self.worker and hasattr(self.worker, "cancel"):
            self.worker.cancel()
            self.cancel_btn.setEnabled(False)
            self.status.setText("Cancelling after the current item…")

    def on_converted(self, result, log_path):
        if self._closing:
            return
        self.cancel_btn.hide()
        word = "Cancelled" if result.cancelled else "Stopped early" if result.stopped else "Done"
        log_note = f"Log: {log_path}" if log_path else "The log couldn't be saved."
        self.status.setText(f"{word} in {_duration(result.seconds)}. {log_note}")
        if not result.cancelled and not result.stopped:
            self.progress.setValue(self.progress.maximum())
        c = result.copied
        lines = [f"Copied: {c[MAIL]:,} emails, {c[CALENDAR]:,} calendar items, "
                 f"{c[CONTACTS]:,} contacts"]
        if result.skipped:
            lines.append(f"Skipped (tasks, notes, contact groups and other items "
                         f"not converted yet): {result.skipped:,}")
        if result.warnings:
            lines.append(f"Copied with warnings: {len(result.warnings):,} (see the log)")
        lines.append(f"Failed: {len(result.failures):,}")
        lines += ["", f"PST: {self.pst_path}", log_note]
        box = QMessageBox(self)
        box.setWindowTitle(APP_NAME)
        box.setIcon(QMessageBox.Warning if result.failures or result.cancelled or result.stopped
                    else QMessageBox.Information)
        if result.stopped:
            detail = result.stopped + " What was copied before that is in the PST."
        elif result.failures:
            detail = "Some items couldn't be converted; they're listed in the log."
        else:
            detail = ""
        box.setText(f"{word}. {detail}")
        box.setInformativeText("\n".join(lines))
        show_btn = box.addButton("Show PST in folder", QMessageBox.ActionRole)
        log_btn = box.addButton("Open log", QMessageBox.ActionRole) if log_path else None
        box.addButton(QMessageBox.Ok)
        box.exec()
        if box.clickedButton() == show_btn:
            _show_in_folder(self.pst_path)
        elif log_btn is not None and box.clickedButton() == log_btn:
            os.startfile(log_path)

    def on_convert_failed(self, msg):
        if self._closing:
            return
        self.cancel_btn.hide()
        self.status.setText("Conversion failed")
        QMessageBox.critical(self, APP_NAME, msg)

    # ---------- threading ----------

    def run_in_thread(self, worker, on_done, on_failed, progress=None):
        thread = QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.done.connect(on_done)
        worker.failed.connect(on_failed)
        if progress:
            worker.progress.connect(progress)
        for sig in (worker.done, worker.failed):
            sig.connect(thread.quit)
        thread.finished.connect(self._thread_finished)
        self.thread, self.worker = thread, worker
        thread.start()

    def _thread_finished(self):
        self.thread.deleteLater()
        self.thread = self.worker = None
        if self._closing:
            self.close()  # the close that was waiting for this job to stop
            return
        self.update_controls()
        if self._pending_path:
            path, self._pending_path = self._pending_path, None
            self.load_file(path)

    def closeEvent(self, event):
        if not self.busy():
            event.accept()
            return
        # Never close with a job still running: the conversion has to detach
        # the PST from Outlook and delete its temporary files first. So ask
        # it to stop, and close when it has (_thread_finished).
        event.ignore()
        if self._closing:
            return
        converting = isinstance(self.worker, ConvertWorker)
        if converting and QMessageBox.question(
                self, APP_NAME, "A conversion is running. Stop it and close?") != QMessageBox.Yes:
            return
        self._closing = True
        self.setEnabled(False)
        if converting:
            self.worker.cancel()
            self.status.setText("Stopping… Unbound closes as soon as Outlook has let go of the PST.")
            # A reader stuck on a damaged file would never notice the cancel
            QTimer.singleShot(15_000, kill_readers)
        else:
            kill_readers()  # only reading folders: nothing to tidy up


def _duration(seconds):
    seconds = int(seconds)
    if seconds < 60:
        return f"{seconds}s"
    m, s = divmod(seconds, 60)
    if m < 60:
        return f"{m}m {s:02d}s"
    h, m = divmod(m, 60)
    return f"{h}h {m:02d}m"


def _size(n_bytes):
    gb = n_bytes / 2**30
    return f"{gb:.1f} GB" if gb >= 1 else f"{max(1, round(n_bytes / 2**20))} MB"


def _same_file(a, b):
    try:
        return os.path.samefile(a, b)  # also sees through short names and links
    except OSError:  # one of them doesn't exist
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _show_in_folder(path):
    """Open the file's folder in Explorer with the file selected."""
    try:
        subprocess.Popen(f'explorer /select,"{os.path.normpath(path)}"')
    except OSError:
        os.startfile(os.path.dirname(path))


def _icon_path():
    # PyInstaller unpacks bundled files to sys._MEIPASS
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, "assets", "icon.ico")


def main():
    if sys.platform == "win32":
        # Own taskbar identity, so Windows shows our icon rather than Python's
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Unbound.OSTtoPST")
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(__version__)
    if os.path.exists(_icon_path()):
        app.setWindowIcon(QIcon(_icon_path()))
    apply_dark_theme(app)
    # Temporary files of a conversion that crashed are removed here; the
    # window then detaches that conversion's PST from Outlook
    win = MainWindow(recover_interrupted())
    win.show()
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):  # file or export folder
        win.load_file(sys.argv[1])
    sys.exit(app.exec())
