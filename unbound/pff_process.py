"""Read an OST/PST in a separate process.

libpff is a C library that decodes a complicated binary format, and the file
may come from anywhere. A damaged file can make it crash or hang, and a crash
takes the whole process with it. In the window's own process that would close
the app mid-conversion, leaving the PST attached to Outlook. So the file is
opened and read in a child process instead: if that dies, the item being read
is reported as failed, the reader is started again and the conversion goes on.
"""

import multiprocessing
import multiprocessing.connection
import weakref

from .converter import _PffReader
from .errors import first_line as _first_line
from .ost_reader import OstInUseError

MAX_RESTARTS = 20
ITEM_TIMEOUT = 600  # seconds for one item; a hung reader is treated like a crashed one
OPEN_TIMEOUT = 600  # seconds for the reader to open the file and say it's ready
SCAN_TIMEOUT = 3600  # seconds to scan the folder tree (very large files take minutes)


class ReaderCrashed(RuntimeError):
    """The reader died on one item; it has been restarted."""


class ReaderLost(RuntimeError):
    """The reader can't be started (again), so nothing more can be read."""


def _serve(conn, path, reader_factory):
    """The child process: open the file, then answer requests until told to stop."""
    try:
        reader = reader_factory(path)
    except OstInUseError as e:
        conn.send(("in_use", e.programs))
        return
    except Exception as e:
        conn.send(("failed", _first_line(e)))
        return
    conn.send(("ready", None))
    try:
        while True:
            try:
                request = conn.recv()
            except EOFError:
                break  # the main process has gone
            if request is None:
                break
            op, *args = request
            try:
                conn.send(("ok", getattr(reader, op)(*args)))
            except Exception as e:
                conn.send(("error", _first_line(e)))
    finally:
        reader.close()


_live = weakref.WeakSet()
_shutting_down = False


def kill_readers():
    """Stop every reader process now and start no more. For closing the app:
    it unblocks a thread that's waiting on a reader that has hung."""
    global _shutting_down
    _shutting_down = True
    for reader in list(_live):
        # Only end the process. The thread waiting on it then sees the pipe
        # close and tidies up its own end; closing that from here could
        # pull it out from under a read in progress.
        proc = reader.proc
        if proc is not None:
            try:
                proc.terminate()
            except (OSError, ValueError, AttributeError):
                pass


class PffProcess:
    """An OST/PST opened in a child process.

    Has the same folder()/item()/close() shape as the converter's other
    readers, plus scan() for the folder tree.
    """

    def __init__(self, path, reader_factory=_PffReader):
        self.path = path
        self.reader_factory = reader_factory
        self.restarts = 0
        self.proc = self.conn = None
        # Listed before it starts, so closing the app can end a reader that
        # hangs while still opening the file
        _live.add(self)
        try:
            self._start()
        except BaseException:
            _live.discard(self)
            raise

    def _start(self):
        if _shutting_down:
            raise ReaderLost("Unbound is closing.")
        ctx = multiprocessing.get_context("spawn")
        self.conn, child_end = ctx.Pipe()
        self.proc = ctx.Process(target=_serve, args=(child_end, self.path, self.reader_factory),
                                daemon=True, name="unbound-reader")
        self.proc.start()
        child_end.close()
        # Wait for the reader's first word, or for it to end. A process that
        # dies before it has taken over its end of the pipe never closes it,
        # so waiting on the pipe alone could wait for ever.
        lost = ReaderLost("The file's reader stopped while opening the file. "
                          "The file is probably damaged.")
        try:
            ready = multiprocessing.connection.wait([self.conn, self.proc.sentinel],
                                                    timeout=OPEN_TIMEOUT)
            if self.conn not in ready:
                raise EOFError
            status, detail = self.conn.recv()
        except (EOFError, OSError, ValueError, TypeError):
            self._stop(polite=False)
            raise lost from None
        if status == "ready":
            return
        self._stop(polite=False)
        if status == "in_use":
            raise OstInUseError(self.path, detail or ())
        raise OSError(detail)

    def _stop(self, polite=True):
        conn, proc, self.conn, self.proc = self.conn, self.proc, None, None
        if conn is not None:
            if polite:
                try:
                    conn.send(None)
                except OSError:
                    pass
            conn.close()
        if proc is not None:
            proc.join(5 if polite else 0)
            if proc.is_alive():
                proc.terminate()
                proc.join(5)

    def _crashed(self, why):
        self._stop(polite=False)
        self.restarts += 1
        if self.restarts > MAX_RESTARTS:
            raise ReaderLost("The file's reader stopped too many times. "
                             "The file is probably badly damaged.")
        try:
            self._start()
        except ReaderLost:
            raise
        except Exception as e:
            raise ReaderLost(f"The file couldn't be reopened after its reader stopped: "
                             f"{_first_line(e)}") from e
        raise ReaderCrashed(why)

    def _ask(self, op, *args, timeout=None):
        if self.conn is None:
            raise ReaderLost("The file's reader isn't running.")
        try:
            self.conn.send((op, *args))
            if not self.conn.poll(timeout):
                self._crashed("the reader took too long on this item, so it was skipped "
                              "(the data is probably damaged here)")
            status, reply = self.conn.recv()
        except (EOFError, OSError):
            self._crashed("the reader stopped on this item "
                          "(the data is probably damaged here)")
        if status == "error":
            raise RuntimeError(reply)
        return reply

    def scan(self):
        """The folders to show, as converter.FolderNode trees."""
        return self._ask("scan", timeout=SCAN_TIMEOUT)

    def folder(self, node):
        return (node.identifier if self._ask("has_folder", node.identifier, timeout=ITEM_TIMEOUT)
                else None)

    def item(self, folder_id, i):
        """(Message, props loader), or (None, None) for a hidden system item."""
        m, props = self._ask("read", folder_id, i, timeout=ITEM_TIMEOUT)
        return m, (lambda: props)

    def close(self):
        self._stop()
        _live.discard(self)


def _pong(conn):
    conn.send("pong")


def child_process_works():
    """Whether this build can start a reader process at all (for the self-check)."""
    ctx = multiprocessing.get_context("spawn")
    parent_end, child_end = ctx.Pipe()
    proc = ctx.Process(target=_pong, args=(child_end,), daemon=True)
    proc.start()
    child_end.close()
    try:
        return bool(parent_end.poll(60)) and parent_end.recv() == "pong"
    except (EOFError, OSError):
        return False
    finally:
        parent_end.close()
        proc.join(5)
        if proc.is_alive():
            proc.terminate()
