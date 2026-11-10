"""Conversion engine: folder tree scan and OST -> PST copy. No GUI code here.

Version 1 copies mail only (message classes starting with IPM.Note).
Calendar, contacts and other item types are skipped and counted.
"""

import os
import time
from dataclasses import dataclass, field

import pypff

from .ost_reader import open_file, read_message

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
    def is_mail(self):
        c = self.container_class
        return c is None or c.upper().startswith("IPF.NOTE")

    @property
    def default_checked(self):
        return self.is_mail and not self.is_system and not self.shared

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


def scan_folders(ost_path):
    """Return the top-level folders to show, as Outlook would.

    An OST keeps the user's folders under Root - Mailbox/IPM_SUBTREE, and
    cached shared mailboxes under Root - Mailbox/Shared Data. Those become
    the top level of the output PST; the rest is internal Outlook data.
    """
    f = open_file(ost_path)
    try:
        raw = _scan(f.get_root_folder(), ())
    finally:
        f.close()

    mailbox = _child(raw, "Root - Mailbox")
    ipm = mailbox and _child(mailbox, "IPM_SUBTREE")
    if not ipm:  # not the usual OST layout (e.g. a PST): show everything
        return [_rebase(c, (c.name,)) for c in raw.children]

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
    copied: int = 0
    skipped_not_mail: int = 0
    failures: list = field(default_factory=list)
    cancelled: bool = False
    seconds: float = 0.0


def convert(ost_path, pst_path, selected, progress=None, should_cancel=None,
            outlook_writer=None):
    """Copy mail from the selected FolderNodes of ost_path into a new PST.

    progress(done, total, folder_path) is called after each item.
    should_cancel() is polled between items.
    outlook_writer lets tests inject a writer; normally a PstWriter is made.
    """
    from .pst_writer import PstWriter

    start = time.time()
    result = Result()
    total = sum(n.item_count for n in selected)

    writer = outlook_writer or PstWriter(
        pst_path, display_name=os.path.splitext(os.path.basename(ost_path))[0])
    f = open_file(ost_path)
    done = 0
    try:
        folders_by_id = {}

        def index(folder):
            folders_by_id[folder.identifier] = folder
            for i in range(folder.number_of_sub_folders):
                try:
                    index(folder.get_sub_folder(i))
                except OSError:
                    pass
        index(f.get_root_folder())

        for node in selected:
            source = folders_by_id.get(node.identifier)
            label = "/".join(node.path)
            if source is None:
                result.failures.append(Failure(label, -1, "", "folder could not be opened"))
                done += node.item_count
                continue
            target = None  # created on first mail item, so empty folders aren't made
            for i in range(node.item_count):
                if should_cancel and should_cancel():
                    result.cancelled = True
                    return result
                subject = ""
                try:
                    item = source.get_sub_message(i)
                    if not isinstance(item, pypff.message):
                        result.skipped_not_mail += 1
                    else:
                        m = read_message(item)
                        subject = m.subject
                        if not m.message_class.upper().startswith("IPM.NOTE"):
                            result.skipped_not_mail += 1
                        else:
                            if target is None:
                                target = writer.get_folder(list(node.path))
                            writer.add_message(target, m)
                            result.copied += 1
                except Exception as e:  # one bad item must not stop the run
                    result.failures.append(Failure(label, i, subject, str(e).splitlines()[0][:300]))
                done += 1
                if progress:
                    progress(done, total, label)
    finally:
        f.close()
        result.seconds = time.time() - start
        if outlook_writer is None:
            writer.close()
    return result


def write_log(result, ost_path, pst_path, log_path):
    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write("Unbound conversion log\n")
        fh.write(f"Source: {ost_path}\nOutput: {pst_path}\n")
        fh.write(f"Finished: {time.strftime('%Y-%m-%d %H:%M:%S')}"
                 f"{'  (cancelled)' if result.cancelled else ''}\n")
        fh.write(f"Copied: {result.copied}\n")
        fh.write(f"Skipped (not mail; calendar/contacts come in a later version): "
                 f"{result.skipped_not_mail}\n")
        fh.write(f"Failed: {len(result.failures)}\n")
        fh.write(f"Time: {result.seconds:.0f}s\n\n")
        if result.failures:
            fh.write("Failed items (folder | item number | subject | error)\n")
            for fl in result.failures:
                fh.write(f"{fl.folder} | {fl.index + 1} | {fl.subject} | {fl.error}\n")
        else:
            fh.write("No items failed.\n")
