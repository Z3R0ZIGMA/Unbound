"""Unbound desktop window (PySide6)."""

import os
import sys
import time

from PySide6.QtCore import QObject, Qt, QThread, Signal
from PySide6.QtGui import QColor, QIcon, QPalette
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QProgressBar, QPushButton, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QWidget,
)

from .converter import CALENDAR, CONTACTS, MAIL, convert, scan_folders, write_log
from .ost_reader import OstInUseError

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
            self.done.emit(scan_folders(self.path))
        except OstInUseError as e:
            self.failed.emit(str(e))
        except Exception as e:
            self.failed.emit("This file couldn't be opened as an Outlook data file.\n\n"
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

    def run(self):
        import pythoncom  # COM must be initialised on the thread that uses it
        pythoncom.CoInitialize()
        try:
            try:
                result = convert(self.ost, self.pst, self.nodes,
                                 progress=lambda d, t, f: self.progress.emit(d, t, f),
                                 should_cancel=lambda: self._cancel)
            except OstInUseError as e:
                self.failed.emit(str(e))
                return
            except Exception as e:
                msg = str(e).splitlines()[0][:400]
                if "Outlook.Application" in msg or "Invalid class string" in msg:
                    msg = ("Couldn't start Outlook. Unbound needs Outlook Classic "
                           "installed on this PC (the New Outlook can't write PST files).")
                self.failed.emit(msg)
                return
            log = os.path.splitext(self.pst)[0] + " - Unbound log.txt"
            write_log(result, self.ost, self.pst, log)
            self.done.emit(result, log)
        finally:
            pythoncom.CoUninitialize()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} – OST to PST")
        self.resize(760, 680)
        self.ost_path = None
        self.thread = None
        self.worker = None
        self._syncing = False

        root = QWidget()
        self.setCentralWidget(root)
        lay = QVBoxLayout(root)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(12)

        title = QLabel(APP_NAME)
        title.setObjectName("title")
        subtitle = QLabel("Convert an Outlook OST file into a PST")
        subtitle.setObjectName("muted")
        lay.addWidget(title)
        lay.addWidget(subtitle)

        # Top: file picker
        row = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        self.path_edit.setPlaceholderText("Choose an OST file…")
        self.browse_btn = QPushButton("Browse…")
        self.browse_btn.clicked.connect(self.choose_file)
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
        lay.addWidget(self.status)

        self.update_controls()

    # ---------- file and tree ----------

    def choose_file(self):
        start = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Outlook")
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose an OST file", start if os.path.isdir(start) else "",
            "Outlook data files (*.ost *.pst);;All files (*)")
        if path:
            self.load_file(path)

    def load_file(self, path):
        self.ost_path = os.path.abspath(path)
        self.path_edit.setText(self.ost_path)
        self.tree.clear()
        self.summary.setText("Reading folders…")
        self.status.setText("")
        self.progress.setValue(0)
        self.run_in_thread(ScanWorker(self.ost_path), self.on_scanned, self.on_scan_failed)
        self.update_controls()

    def on_scan_failed(self, msg):
        self.summary.setText("Couldn't read this file")
        QMessageBox.warning(self, APP_NAME, msg)
        self.ost_path = None
        self.update_controls()

    def on_scanned(self, top_nodes):
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

    def start_convert(self):
        nodes = self.selected_nodes()
        if not nodes:
            return
        default = os.path.join(os.path.expanduser("~"), "Documents",
                               os.path.splitext(os.path.basename(self.ost_path))[0] + ".pst")
        pst, _ = QFileDialog.getSaveFileName(self, "Save PST as", default,
                                             "Outlook data file (*.pst)")
        if not pst:
            return
        pst = os.path.abspath(pst)
        if os.path.normcase(pst) == os.path.normcase(self.ost_path):
            QMessageBox.warning(self, APP_NAME, "Choose a different file from the source.")
            return
        if os.path.exists(pst):
            # The save dialog already asked to replace it. Outlook would add to an
            # existing PST rather than replace it, so remove it first.
            try:
                os.remove(pst)
            except OSError as e:
                QMessageBox.warning(self, APP_NAME,
                                    f"Couldn't replace the existing file (is it open in Outlook?)\n\n{e}")
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
        self.cancel_btn.hide()
        word = "Cancelled" if result.cancelled else "Done"
        self.status.setText(f"{word} in {_duration(result.seconds)}. Log: {log_path}")
        if not result.cancelled:
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
        lines += ["", f"PST: {self.pst_path}", f"Log: {log_path}"]
        box = QMessageBox(self)
        box.setWindowTitle(APP_NAME)
        box.setIcon(QMessageBox.Warning if result.failures or result.cancelled
                    else QMessageBox.Information)
        box.setText(f"{word}. " + ("Some items couldn't be converted; they're listed in the log."
                                   if result.failures else ""))
        box.setInformativeText("\n".join(lines))
        open_btn = box.addButton("Open folder", QMessageBox.ActionRole)
        box.addButton(QMessageBox.Ok)
        box.exec()
        if box.clickedButton() == open_btn:
            os.startfile(os.path.dirname(self.pst_path))

    def on_convert_failed(self, msg):
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
        self.update_controls()

    def closeEvent(self, event):
        if self.busy():
            if QMessageBox.question(self, APP_NAME,
                                    "A conversion is running. Stop it and close?") != QMessageBox.Yes:
                event.ignore()
                return
            if hasattr(self.worker, "cancel"):
                self.worker.cancel()
            self.thread.quit()
            self.thread.wait(120_000)
        event.accept()


def _duration(seconds):
    seconds = int(seconds)
    if seconds < 60:
        return f"{seconds}s"
    m, s = divmod(seconds, 60)
    if m < 60:
        return f"{m}m {s:02d}s"
    h, m = divmod(m, 60)
    return f"{h}h {m:02d}m"


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
    if os.path.exists(_icon_path()):
        app.setWindowIcon(QIcon(_icon_path()))
    apply_dark_theme(app)
    win = MainWindow()
    win.show()
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        win.load_file(sys.argv[1])
    sys.exit(app.exec())
