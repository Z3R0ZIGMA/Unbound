"""Conversion engine: folder tree scan and OST -> PST copy. No GUI code here.

Copies mail (IPM.Note*), appointments (IPM.Appointment*) and contacts
(IPM.Contact). Tasks, notes, contact groups, meeting requests and other item
types are skipped and counted in the log.
"""

import json
import mmap
import os
import shutil
import tempfile
import time
from collections import Counter
from dataclasses import dataclass, field

import pypff

from . import __version__
from .errors import first_line
from .mbox_reader import find_mailboxes, is_mbox_source, message_offsets, parse_message, read_raw
from .ost_reader import (
    OstInUseError, is_lock_error, is_locked, open_file, programs_using, read_message,
)
from .props import load_name_map, read_props

MAIL, CALENDAR, CONTACTS = "mail", "calendar", "contacts"

# Outlook stops writing to a PST at 50 GB unless its limit has been raised
PST_SIZE_LIMIT = 50 * 2**30

# Hidden system folders that hold no user mail; unticked by default
SYSTEM_FOLDERS = {
    "itemprocsearch", "personmetadata", "recipient cache", "yammer root",
    "conversation action settings", "files", "gal contacts", "companies",
    "organizational contacts", "peoplecentricconversation buddies",
    "externalcontacts", "quick step settings", "sync issues", "conflicts",
    "local failures", "server failures", "recoverable items", "deletions",
    "purges", "versions", "calendar logging", "audits", "spooler queue",
    "search root", "finder", "views", "common views", "freebusy data",
    "shortcuts", "reminders", "tracked mail processing", "deferred action",
    "schedule", "to-do search", "mrm", "social activity notifications",
    "inbound", "outbound", "feeds", "team chat", "teamschatsession",
}


CONTAINER_CLASS = 0x3613


@dataclass
class FolderNode:
    identifier: int
    name: str
    path: tuple  # folder names in the output PST, top level first
    item_count: int
    container_class: str | None  # IPF.Note = mail, IPF.Appointment = calendar, ...
    children: list = field(default_factory=list)
    shared: bool = False  # inside a cached shared mailbox

    @property
    def is_system(self):
        n = self.name.strip().lower()
        return n in SYSTEM_FOLDERS or n.startswith("{") or n.startswith("ms-olk-")

    @property
    def kind(self):
        """MAIL, CALENDAR, CONTACTS, or None for folders not converted yet."""
        c = (self.container_class or "").upper()
        # IPF.Note = Exchange/PST mail folders; IPF.Imap = IMAP account folders
        if c == "" or c.startswith("IPF.NOTE") or c.startswith("IPF.IMAP"):
            return MAIL
        if c.startswith("IPF.APPOINTMENT"):
            return CALENDAR
        # Plain contact folders only; IPF.Contact.RecipientCache etc. are
        # Outlook's hidden address caches
        if c == "IPF.CONTACT":
            return CONTACTS
        return None

    @property
    def is_supported(self):
        return self.kind is not None

    @property
    def default_checked(self):
        return self.is_supported and not self.is_system and not self.shared

    def total_items(self):
        return self.item_count + sum(c.total_items() for c in self.children)

    def walk(self):
        yield self
        for c in self.children:
            yield from c.walk()


def _container_class(folder):
    try:
        for e in folder.record_sets[0].entries:
            if e.entry_type == CONTAINER_CLASS:
                return e.data_as_string
    except (OSError, IndexError, TypeError):
        pass
    return None


def _scan(folder, path, name=None):
    try:
        count = folder.number_of_sub_messages
    except OSError:
        count = 0
    if name is None:
        name = folder.name or ""
    node = FolderNode(folder.identifier, name, path, count, _container_class(folder))
    try:
        n_sub = folder.number_of_sub_folders
    except OSError:
        n_sub = 0
    for i in range(n_sub):
        try:
            sub = folder.get_sub_folder(i)
        except OSError:
            continue
        # Outlook can't create a folder without a name, so a nameless one
        # gets a name here that the tree and the PST then both use
        sub_name = sub.name if (sub.name or "").strip() else f"Folder {i + 1}"
        node.children.append(_scan(sub, path + (sub_name,), sub_name))
    return node


def _child(node, name):
    for c in node.children:
        if c.name == name:
            return c
    return None


def _rebase(node, path, shared=False):
    """Copy of node with output paths starting at `path`."""
    return FolderNode(node.identifier, node.name, path, node.item_count,
                      node.container_class,
                      [_rebase(c, path + (c.name,), shared) for c in node.children],
                      shared)


# Names Outlook gives the folder that holds the user's mail ("IPM subtree")
# in PST-style files, which is the layout IMAP and POP OSTs use
MAIL_ROOT_NAMES = {
    "top of outlook data file", "top of personal folders",
    "top of information store", "ipm_subtree",
}
# Internal folders that sit beside the mail root and are never it
NOT_MAIL_ROOT = {"search root", "spam search folder 2", "finder", "itemprocsearch",
                 "root - public", "freebusy data", "common views", "views"}


def _find_mail_root(raw):
    """Pick the folder holding the user's mail in a non-Exchange layout."""
    candidates = [c for c in raw.children if c.name.strip().lower() not in NOT_MAIL_ROOT]
    for c in candidates:
        if c.name.strip().lower() in MAIL_ROOT_NAMES:
            return c
    # Unknown name (e.g. translated Outlook): if one folder holds the mail, use it
    with_items = [c for c in candidates if c.total_items()]
    if len(with_items) == 1 and with_items[0].children:
        return with_items[0]
    return None


def scan_folders(path):
    """Top-level folders of an OST/PST or an Apple Mail / MBOX export."""
    if is_mbox_source(path):
        return _scan_mbox(path)
    from .pff_process import PffProcess
    reader = PffProcess(path)
    try:
        return reader.scan()
    finally:
        reader.close()


def _scan_mbox(path):
    def node(box):
        count = len(message_offsets(box.file)) if box.file else 0
        # identifier: the mbox file to read; folder-only mailboxes get a key
        # that never matches a file
        ident = box.file or ("folder", box.path)
        return FolderNode(ident, box.name, box.path, count, "IPF.Note",
                          [node(c) for c in box.children])
    return [node(b) for b in find_mailboxes(path)]


def _visible_tree(raw):
    """The top-level folders to show, as Outlook would, from the raw tree
    of the whole file.

    An Exchange OST keeps the user's folders under Root - Mailbox/IPM_SUBTREE
    and cached shared mailboxes under Root - Mailbox/Shared Data. IMAP OSTs
    and PSTs use a "Top of Outlook data file" folder instead. Either way that
    folder's contents become the top level of the output PST; the rest is
    internal Outlook data.
    """
    mailbox = _child(raw, "Root - Mailbox")
    ipm = mailbox and _child(mailbox, "IPM_SUBTREE")
    if not ipm:
        root = _find_mail_root(raw)
        if root is None:  # unrecognised layout: show everything as it is
            return [_rebase(c, (c.name,)) for c in raw.children]
        return [_rebase(c, (c.name,)) for c in root.children]

    top = [_rebase(c, (c.name,)) for c in ipm.children]
    shared = _child(mailbox, "Shared Data")
    if shared and shared.total_items():
        group = _rebase(shared, ("Shared mailboxes",), shared=True)
        group.name = "Shared mailboxes"
        top.append(group)
    return top


@dataclass
class Failure:
    folder: str
    index: int
    subject: str
    error: str


@dataclass
class Result:
    copied: Counter = field(default_factory=Counter)  # MAIL/CALENDAR/CONTACTS -> count
    skipped: int = 0
    failures: list = field(default_factory=list)
    warnings: list = field(default_factory=list)  # items copied only partly
    cancelled: bool = False
    stopped: str = ""  # why the run ended early by itself (the source became unreadable)
    still_attached: str = ""  # why the PST couldn't be detached from Outlook at the end
    source_in_use_by: list = field(default_factory=list)  # programs that had the source open
    seconds: float = 0.0
    skipped_attached_emails: int = 0
    # Diagnostics for the log: what kinds of folders and items were met
    folder_types: Counter = field(default_factory=Counter)
    item_types: Counter = field(default_factory=Counter)  # message class -> count
    skipped_types: Counter = field(default_factory=Counter)
    refused_properties: dict = field(default_factory=dict)  # property -> Outlook's error

    @property
    def total_copied(self):
        return sum(self.copied.values())


def item_kind(message_class):
    c = message_class.upper()
    if c.startswith("IPM.NOTE"):
        return MAIL
    if c.startswith("IPM.APPOINTMENT"):
        return CALENDAR
    if c == "IPM.CONTACT" or c.startswith("IPM.CONTACT."):
        return CONTACTS
    return None  # tasks, notes, contact groups, meeting requests, ...


class _PffReader:
    """An OST/PST through pypff. Lives in the reader process (pff_process),
    not in the window's own process, so a crash in libpff can't take the
    window down."""

    def __init__(self, path):
        self.f = open_file(path)
        self.name_map = None
        self.folders = None

    def scan(self):
        return _visible_tree(_scan(self.f.get_root_folder(), ()))

    def _folder(self, identifier):
        if self.folders is None:
            self.folders = {}

            def index(folder):
                self.folders[folder.identifier] = folder
                for i in range(folder.number_of_sub_folders):
                    try:
                        index(folder.get_sub_folder(i))
                    except OSError:
                        pass
            index(self.f.get_root_folder())
        return self.folders.get(identifier)

    def has_folder(self, identifier):
        return self._folder(identifier) is not None

    def read(self, identifier, i):
        """(Message, properties) of item i in a folder. (None, None) for a
        hidden system item; properties only for calendar items and contacts,
        which are the ones recreated from them."""
        folder = self._folder(identifier)
        if folder is None:
            raise LookupError("folder could not be opened")
        item = folder.get_sub_message(i)
        if not isinstance(item, pypff.message):
            return None, None
        m = read_message(item)
        props = None
        if item_kind(m.message_class) in (CALENDAR, CONTACTS):
            if self.name_map is None:
                self.name_map = load_name_map(self.f)
            props = read_props(item, self.name_map)
        return m, props

    def close(self):
        self.f.close()


class _MboxReader:
    """Messages of mbox files; a folder node's identifier is its file."""

    def folder(self, node):
        return node.identifier if isinstance(node.identifier, str) else None

    def item(self, file, i):
        return parse_message(read_raw(file, i)), None

    def close(self):
        pass


# ---------- a run's temporary files, and cleaning up after a crash ----------

RUN_MARKER = "unbound-run.json"


class _Run:
    """A conversion's own temporary folder, for the working copy of the OST
    and attachments on their way into Outlook.

    The marker file in it stays open for as long as the conversion runs.
    Windows won't let an open file be renamed, so the next start can tell a
    run that crashed (the marker can be renamed) from one still going in
    another window (it can't), and clean up only the first.
    """

    def __init__(self, pst_path):
        self.dir = tempfile.mkdtemp(prefix="unbound_")
        self._marker = open(os.path.join(self.dir, RUN_MARKER), "w", encoding="utf-8")
        json.dump({"pst": os.path.abspath(pst_path), "version": __version__,
                   "started": time.strftime("%Y-%m-%d %H:%M:%S")}, self._marker)
        self._marker.flush()

    def close(self):
        self._marker.close()
        shutil.rmtree(self.dir, ignore_errors=True)


def recover_interrupted():
    """Clean up after conversions that didn't end normally (a crash, a power
    cut, a forced close).

    Deletes their temporary folders, which can hold a full copy of the
    mailbox, and returns the PST files they were writing. Those may still be
    attached to the Outlook profile: detach them with
    pst_writer.detach_stores, or pass them to convert() as `detach`.
    """
    stale = []
    root = tempfile.gettempdir()
    try:
        names = os.listdir(root)
    except OSError:
        return stale
    for name in names:
        folder = os.path.join(root, name)
        marker = os.path.join(folder, RUN_MARKER)
        dead = marker + ".dead"
        if not name.startswith("unbound_"):
            continue
        newly_found = not os.path.exists(dead)
        if newly_found:
            if not os.path.isfile(marker):
                continue  # not a run folder of ours
            try:
                os.replace(marker, dead)
            except OSError:
                continue  # still open, so that conversion is still running
        if newly_found:
            try:
                with open(dead, encoding="utf-8") as fh:
                    pst = json.load(fh).get("pst")
                if isinstance(pst, str) and pst:
                    stale.append(pst)
            except (OSError, ValueError, AttributeError):
                pass
        # A file in it may be briefly held (a virus scanner on the copy);
        # the renamed marker gets the folder another try next time
        shutil.rmtree(folder, ignore_errors=True)
    return stale


class _Cancelled(Exception):
    """The user cancelled during a setup step."""


COPY_CHUNK = 8 * 2**20


def _copy_file(source, target, on_progress=None, should_cancel=None):
    """Copy a file in pieces, so a big one shows progress and can be cancelled.

    If another program has locked part of the file (Word, Excel and others
    do that to a mailbox they have open), ordinary reads of that part are
    refused. Windows doesn't apply such locks to a memory-mapped view, so
    the copy carries on through one.
    """
    size = os.path.getsize(source)
    copied = 0
    view = None
    with open(source, "rb") as fin, open(target, "wb") as fout:
        try:
            while True:
                if should_cancel and should_cancel():
                    raise _Cancelled()
                if view is None:
                    try:
                        chunk = fin.read(COPY_CHUNK)
                    except OSError as e:
                        if not (isinstance(e, PermissionError) or is_lock_error(e)):
                            raise
                        view = mmap.mmap(fin.fileno(), 0, access=mmap.ACCESS_READ)
                        continue
                else:
                    chunk = view[copied:copied + COPY_CHUNK]
                if not chunk:
                    break
                fout.write(chunk)
                copied += len(chunk)
                if on_progress:
                    on_progress(copied, size)
        finally:
            if view is not None:
                view.close()


def is_outlook_data_file(path):
    """True for files in Outlook's own data folder, i.e. ones a profile on
    this PC may be using. OSTs copied from elsewhere are read in place."""
    folder = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Outlook")
    p = os.path.normcase(os.path.abspath(path))
    return p.startswith(os.path.normcase(os.path.abspath(folder)) + os.sep)


def _working_copy(path, folder, on_progress=None, should_cancel=None):
    """Copy the OST into `folder` (the run's temporary folder); returns the copy."""
    size = os.path.getsize(path)
    free = shutil.disk_usage(folder).free
    if free < size * 1.05:
        raise RuntimeError(
            f"Not enough free space for a working copy of the OST: it needs "
            f"{size / 2**30:.1f} GB in {folder}, and {free / 2**30:.1f} GB is free.")
    copy = os.path.join(folder, os.path.basename(path))
    # A file can be held briefly in a way that refuses even opening it, so
    # retry a few times before reporting who has it
    for attempt in range(4):
        try:
            _copy_file(path, copy, on_progress, should_cancel)
            return copy
        except OSError as e:
            locked = isinstance(e, PermissionError) or is_lock_error(e)
            if locked and attempt < 3:
                time.sleep(2)
                continue
            if locked:
                raise OstInUseError(os.path.abspath(path), programs_using(path)) from e
            raise


def _copy_reporter(progress, total):
    """Progress of the working copy, as a status line that changes once per percent."""
    last = [-1]

    def report(copied, size):
        percent = min(99, copied * 100 // size) if size else 99
        if percent != last[0]:
            last[0] = percent
            progress(0, total, f"Making a working copy of the OST… {percent}%")
    return report


def estimated_pst_size(source, selected, all_nodes):
    """Roughly how big the PST will be, in bytes: the source's size scaled
    by the share of its items that are selected."""
    try:
        if is_mbox_source(source):
            return sum(os.path.getsize(n.identifier) for n in selected
                       if isinstance(n.identifier, str) and n.item_count)
        everything = sum(n.item_count for n in all_nodes)
        if not everything:
            return 0
        return int(os.path.getsize(source) * sum(n.item_count for n in selected) / everything)
    except OSError:
        return 0


def convert(ost_path, pst_path, selected, progress=None, should_cancel=None,
            outlook_writer=None, detach=()):
    """Copy mail, appointments and contacts from the selected FolderNodes
    of ost_path into a new PST.

    progress(done, total, folder_path) is called after each item, and with
    done = 0 and a status text during setup steps.
    should_cancel() is polled between items.
    outlook_writer lets tests inject a writer; normally a PstWriter is made.
    detach lists PST files an interrupted conversion may have left attached
    to Outlook (see recover_interrupted); they're detached first.
    """
    from .pff_process import PffProcess, ReaderLost
    from .pst_writer import (
        DETACH_TIMEOUT, OL_APPOINTMENT_ITEM, OL_CONTACT_ITEM, OL_FOLDER_CALENDAR,
        OL_FOLDER_CONTACTS, DetachError, PstWriter,
    )
    folder_type = {CALENDAR: OL_FOLDER_CALENDAR, CONTACTS: OL_FOLDER_CONTACTS}
    item_type = {CALENDAR: OL_APPOINTMENT_ITEM, CONTACTS: OL_CONTACT_ITEM}

    start = time.time()
    result = Result()
    total = sum(n.item_count for n in selected)
    for n in selected:
        result.folder_types[n.container_class or "(none)"] += 1

    mbox = is_mbox_source(ost_path)
    reader = writer = None
    done = 0
    run = _Run(pst_path)
    try:
        # Outlook locks the OST of its own profile as soon as it starts, and we
        # start Outlook to write the PST. So read such files from a copy taken
        # before Outlook starts. The same goes for a file another program has
        # open right now (Word and Excel hold the mailbox open while they run).
        read_path = ost_path
        in_use = not mbox and is_locked(ost_path)
        if in_use:
            result.source_in_use_by = programs_using(ost_path)
            error = OstInUseError(os.path.abspath(ost_path), result.source_in_use_by)
            if error.outlook:
                raise error  # Outlook changes the file as it syncs: it has to be closed
        if not mbox and (in_use or is_outlook_data_file(ost_path)):
            if progress:
                progress(0, total, "Making a working copy of the OST…")
            read_path = _working_copy(ost_path, run.dir,
                                      progress and _copy_reporter(progress, total), should_cancel)

        if mbox:
            reader, tree = _MboxReader(), _scan_mbox(read_path)
        else:
            reader = PffProcess(read_path)
            tree = reader.scan()
        # Folder kind for every path, so parents of a calendar or contacts
        # folder are created with the right type too
        kinds_by_path = {n.path: n.kind for top in tree for n in top.walk()}
        name = os.path.splitext(os.path.basename(os.path.normpath(ost_path)))[0]
        writer = outlook_writer or PstWriter(pst_path, display_name=name,
                                             temp_dir=run.dir, detach=detach)

        for node in selected:
            if not node.item_count:
                continue  # nothing to copy (e.g. an Apple mailbox that only holds others)
            label = "/".join(node.path)
            try:
                source = reader.folder(node)
            except ReaderLost as e:
                result.failures.append(Failure(label, -1, "", str(e)))
                result.stopped = str(e)
                return result
            except Exception:
                source = None
            if source is None:
                result.failures.append(Failure(label, -1, "", "folder could not be opened"))
                done += node.item_count
                continue
            target = None  # created on the first item copied, so empty folders aren't made
            for i in range(node.item_count):
                if should_cancel and should_cancel():
                    result.cancelled = True
                    return result
                subject = ""
                try:
                    m, load_props = reader.item(source, i)
                    if m is None:
                        result.skipped += 1
                        result.item_types["(system item)"] += 1
                        result.skipped_types["(system item)"] += 1
                    else:
                        subject = m.subject
                        result.item_types[m.message_class] += 1
                        kind = item_kind(m.message_class)
                        if kind is None:
                            result.skipped += 1
                            result.skipped_types[m.message_class] += 1
                        else:
                            if target is None:
                                parts = list(node.path)
                                types = [folder_type.get(kinds_by_path.get(tuple(parts[:k + 1])))
                                         for k in range(len(parts))]
                                target = writer.get_folder(parts, types)
                            warnings_before = len(writer.warnings)
                            if kind == MAIL:
                                writer.add_message(target, m)
                            else:
                                writer.add_item(target, item_type[kind], m, load_props())
                            result.copied[kind] += 1
                            result.skipped_attached_emails += m.skipped_embedded
                            for w in writer.warnings[warnings_before:]:
                                result.warnings.append(Failure(label, i, subject, w))
                except ReaderLost as e:  # the source can't be read any more: stop here
                    result.failures.append(Failure(label, i, subject, str(e)))
                    result.stopped = str(e)
                    return result
                except Exception as e:  # one bad item must not stop the run
                    result.failures.append(Failure(label, i, subject, first_line(e)))
                done += 1
                if progress:
                    progress(done, total, label)
    except _Cancelled:
        result.cancelled = True
    finally:
        result.seconds = time.time() - start
        try:
            if reader is not None:
                reader.close()
        finally:
            try:
                if writer is not None:
                    result.refused_properties = dict(writer.property_errors)
                if writer is not None and outlook_writer is None:
                    # Always detach the PST from the profile. If Outlook won't,
                    # what was copied is still in the PST: say so rather than
                    # report the whole run as failed
                    if progress:
                        progress(0, total, "Detaching the PST from Outlook…")
                    try:
                        writer.close(timeout=DETACH_TIMEOUT)
                    except DetachError as e:
                        result.still_attached = str(e)
            finally:
                run.close()
    return result


STILL_ATTACHED_ADVICE = (
    "Everything copied is in the PST, but it may still be open in Outlook. If Outlook "
    "lists it as an extra data file, right-click its name in the folder list and "
    "choose Close. If Outlook can't open the PST, run Outlook's Inbox Repair Tool "
    "(SCANPST.EXE) on it.")


def _one_line(text):
    """Text from the mailbox (a subject, a folder name) on one line: a line
    break in it could otherwise write lines that look like the log's own,
    such as a fake "Failed: 0"."""
    return " ".join(str(text or "").split())


def log_path_for(pst_path):
    return os.path.splitext(pst_path)[0] + " - Unbound log.txt"


def write_failure_log(error, ost_path, pst_path, log_path):
    """The log for a conversion that couldn't start or broke off with an error."""
    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write(f"Unbound {__version__} conversion log\n")
        fh.write(f"Source: {ost_path}\nOutput: {pst_path}\n")
        fh.write(f"Finished: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        fh.write(f"Conversion failed: {error}\n")


def write_log(result, ost_path, pst_path, log_path):
    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write(f"Unbound {__version__} conversion log\n")
        fh.write(f"Source: {ost_path}\nOutput: {pst_path}\n")
        fh.write(f"Finished: {time.strftime('%Y-%m-%d %H:%M:%S')}"
                 f"{'  (cancelled)' if result.cancelled else ''}\n")
        if result.stopped:
            fh.write(f"Stopped early: {result.stopped}\n")
        if result.still_attached:
            fh.write(f"Not detached from Outlook: {_one_line(result.still_attached)}. "
                     f"{STILL_ATTACHED_ADVICE}\n")
        if result.source_in_use_by:
            fh.write(f"The source was open in {', '.join(result.source_in_use_by)}; "
                     f"it was converted from a copy taken at the start.\n")
        fh.write(f"Copied: {result.total_copied} (mail {result.copied[MAIL]}, "
                 f"calendar {result.copied[CALENDAR]}, contacts {result.copied[CONTACTS]})\n")
        fh.write(f"Skipped (tasks, notes, contact groups, meeting requests and other "
                 f"items not converted yet): {result.skipped}\n")
        fh.write(f"Failed: {len(result.failures)}\n")
        if result.skipped_attached_emails:
            fh.write(f"Emails attached to other items, not copied yet: "
                     f"{result.skipped_attached_emails}\n")
        fh.write(f"Time: {result.seconds:.0f}s\n\n")

        def counts(title, counter):
            fh.write(f"{title}\n")
            for name, n in counter.most_common():
                fh.write(f"  {name}: {n}\n")
            if not counter:
                fh.write("  (none)\n")
            fh.write("\n")
        counts("Folder types converted (container class: folders)", result.folder_types)
        counts("Item types seen (message class: items)", result.item_types)
        counts("Item types skipped (message class: items)", result.skipped_types)
        if result.refused_properties:
            fh.write("Properties Outlook didn't accept (a few are normal, e.g. ones only "
                     "Outlook itself may set; listed to help diagnose problems)\n")
            for prop, error in sorted(result.refused_properties.items()):
                fh.write(f"  {prop}: {error}\n")
            fh.write("\n")

        if result.warnings:
            fh.write("Copied with warnings (folder | item number | subject | warning)\n")
            for w in result.warnings:
                fh.write(f"{_one_line(w.folder)} | {w.index + 1} | {_one_line(w.subject)} | "
                         f"{_one_line(w.error)}\n")
            fh.write("\n")
        if result.failures:
            fh.write("Failed items (folder | item number | subject | error)\n")
            for fl in result.failures:
                fh.write(f"{_one_line(fl.folder)} | {fl.index + 1} | {_one_line(fl.subject)} | "
                         f"{_one_line(fl.error)}\n")
        else:
            fh.write("No items failed.\n")
