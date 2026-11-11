"""Conversion engine: folder tree scan and OST -> PST copy. No GUI code here.

Copies mail (IPM.Note*), appointments (IPM.Appointment*) and contacts
(IPM.Contact). Tasks, notes, contact groups, meeting requests and other item
types are skipped and counted in the log.
"""

import os
import shutil
import tempfile
import time
from collections import Counter
from dataclasses import dataclass, field

import pypff

from .mbox_reader import find_mailboxes, is_mbox_source, message_offsets, parse_message, read_raw
from .ost_reader import OstInUseError, is_lock_error, open_file, read_message
from .props import load_name_map, read_props

MAIL, CALENDAR, CONTACTS = "mail", "calendar", "contacts"

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


def _scan(folder, path):
    try:
        count = folder.number_of_sub_messages
    except OSError:
        count = 0
    node = FolderNode(folder.identifier, folder.name or "", path, count,
                      _container_class(folder))
    try:
        n_sub = folder.number_of_sub_folders
    except OSError:
        n_sub = 0
    for i in range(n_sub):
        try:
            sub = folder.get_sub_folder(i)
        except OSError:
            continue
        node.children.append(_scan(sub, path + ((sub.name or f"Folder {i + 1}"),)))
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
    return _scan_pff(path)


def _scan_mbox(path):
    def node(box):
        count = len(message_offsets(box.file)) if box.file else 0
        # identifier: the mbox file to read; folder-only mailboxes get a key
        # that never matches a file
        ident = box.file or ("folder", box.path)
        return FolderNode(ident, box.name, box.path, count, "IPF.Note",
                          [node(c) for c in box.children])
    return [node(b) for b in find_mailboxes(path)]


def _scan_pff(ost_path):
    """Return the top-level folders to show, as Outlook would.

    An Exchange OST keeps the user's folders under Root - Mailbox/IPM_SUBTREE
    and cached shared mailboxes under Root - Mailbox/Shared Data. IMAP OSTs
    and PSTs use a "Top of Outlook data file" folder instead. Either way that
    folder's contents become the top level of the output PST; the rest is
    internal Outlook data.
    """
    f = open_file(ost_path)
    try:
        raw = _scan(f.get_root_folder(), ())
    finally:
        f.close()

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
    seconds: float = 0.0
    skipped_attached_emails: int = 0
    # Diagnostics for the log: what kinds of folders and items were met
    folder_types: Counter = field(default_factory=Counter)
    item_types: Counter = field(default_factory=Counter)  # message class -> count
    skipped_types: Counter = field(default_factory=Counter)

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
    """Items of an OST/PST through pypff."""

    def __init__(self, path):
        self.f = open_file(path)
        self.name_map = load_name_map(self.f)
        self.folders = {}

        def index(folder):
            self.folders[folder.identifier] = folder
            for i in range(folder.number_of_sub_folders):
                try:
                    index(folder.get_sub_folder(i))
                except OSError:
                    pass
        index(self.f.get_root_folder())

    def folder(self, node):
        return self.folders.get(node.identifier)

    def item(self, folder, i):
        """(Message, props loader), or (None, None) for a hidden system item."""
        item = folder.get_sub_message(i)
        if not isinstance(item, pypff.message):
            return None, None
        return read_message(item), lambda: read_props(item, self.name_map)

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


def convert(ost_path, pst_path, selected, progress=None, should_cancel=None,
            outlook_writer=None):
    """Copy mail, appointments and contacts from the selected FolderNodes
    of ost_path into a new PST.

    progress(done, total, folder_path) is called after each item.
    should_cancel() is polled between items.
    outlook_writer lets tests inject a writer; normally a PstWriter is made.
    """
    from .pst_writer import (
        OL_APPOINTMENT_ITEM, OL_CONTACT_ITEM, OL_FOLDER_CALENDAR, OL_FOLDER_CONTACTS, PstWriter,
    )
    folder_type = {CALENDAR: OL_FOLDER_CALENDAR, CONTACTS: OL_FOLDER_CONTACTS}
    item_type = {CALENDAR: OL_APPOINTMENT_ITEM, CONTACTS: OL_CONTACT_ITEM}

    start = time.time()
    result = Result()
    total = sum(n.item_count for n in selected)
    for n in selected:
        result.folder_types[n.container_class or "(none)"] += 1

    # Outlook locks the OST of its own profile as soon as it starts, and we
    # start Outlook to write the PST. So read such files from a copy taken
    # before Outlook starts.
    mbox = is_mbox_source(ost_path)
    read_path, temp_dir = ost_path, None
    if not mbox and is_outlook_data_file(ost_path):
        if progress:
            progress(0, total, "Making a working copy of the OST…")
        read_path, temp_dir = _working_copy(ost_path)

    reader = writer = None
    done = 0
    try:
        # Folder kind for every path, so parents of a calendar or contacts
        # folder are created with the right type too
        kinds_by_path = {n.path: n.kind for top in scan_folders(read_path) for n in top.walk()}
        reader = _MboxReader() if mbox else _PffReader(read_path)
        name = os.path.splitext(os.path.basename(os.path.normpath(ost_path)))[0]
        writer = outlook_writer or PstWriter(pst_path, display_name=name)

        for node in selected:
            if not node.item_count:
                continue  # nothing to copy (e.g. an Apple mailbox that only holds others)
            source = reader.folder(node)
            label = "/".join(node.path)
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
                except Exception as e:  # one bad item must not stop the run
                    result.failures.append(Failure(label, i, subject, str(e).splitlines()[0][:300]))
                done += 1
                if progress:
                    progress(done, total, label)
    finally:
        result.seconds = time.time() - start
        try:
            if reader is not None:
                reader.close()
        finally:
            try:
                if writer is not None and outlook_writer is None:
                    writer.close()  # always detach the PST from the profile
            finally:
                if temp_dir:
                    shutil.rmtree(temp_dir, ignore_errors=True)
    return result


def is_outlook_data_file(path):
    """True for files in Outlook's own data folder, i.e. ones a profile on
    this PC may be using. OSTs copied from elsewhere are read in place."""
    folder = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Outlook")
    p = os.path.normcase(os.path.abspath(path))
    return p.startswith(os.path.normcase(os.path.abspath(folder)) + os.sep)


def _working_copy(path):
    size = os.path.getsize(path)
    temp_dir = tempfile.mkdtemp(prefix="unbound_")
    free = shutil.disk_usage(temp_dir).free
    if free < size * 1.05:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise RuntimeError(
            f"Not enough free space for a working copy of the OST: it needs "
            f"{size / 2**30:.1f} GB in {temp_dir}, and {free / 2**30:.1f} GB is free.")
    copy = os.path.join(temp_dir, os.path.basename(path))
    # A running Outlook locks parts of its OST only briefly (e.g. while
    # syncing), so retry a few times before asking the user to close it
    for attempt in range(4):
        try:
            shutil.copyfile(path, copy)
            return copy, temp_dir
        except OSError as e:
            locked = isinstance(e, PermissionError) or is_lock_error(e)
            if locked and attempt < 3:
                time.sleep(2)
                continue
            shutil.rmtree(temp_dir, ignore_errors=True)
            if locked:
                raise OstInUseError(os.path.abspath(path)) from e
            raise


def write_log(result, ost_path, pst_path, log_path):
    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write("Unbound conversion log\n")
        fh.write(f"Source: {ost_path}\nOutput: {pst_path}\n")
        fh.write(f"Finished: {time.strftime('%Y-%m-%d %H:%M:%S')}"
                 f"{'  (cancelled)' if result.cancelled else ''}\n")
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

        if result.warnings:
            fh.write("Copied with warnings (folder | item number | subject | warning)\n")
            for w in result.warnings:
                fh.write(f"{w.folder} | {w.index + 1} | {w.subject} | {w.error}\n")
            fh.write("\n")
        if result.failures:
            fh.write("Failed items (folder | item number | subject | error)\n")
            for fl in result.failures:
                fh.write(f"{fl.folder} | {fl.index + 1} | {fl.subject} | {fl.error}\n")
        else:
            fh.write("No items failed.\n")
