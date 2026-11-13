"""Stand-ins used by the tests."""

import os
import time


class FakeFolder:
    """Looks enough like a pypff folder for the folder scan."""

    def __init__(self, identifier, name, items=0, children=(), container_class=None):
        self.identifier = identifier
        self.name = name
        self.number_of_sub_messages = items
        self._children = list(children)
        self._container_class = container_class

    @property
    def number_of_sub_folders(self):
        return len(self._children)

    def get_sub_folder(self, i):
        return self._children[i]

    @property
    def record_sets(self):
        if self._container_class is None:
            return []
        entry = type("Entry", (), {"entry_type": 0x3613, "data_as_string": self._container_class})
        return [type("RecordSet", (), {"entries": [entry]})]


class FakeWriter:
    """Records what the converter asks Outlook to do."""

    def __init__(self, fail_on=()):
        self.warnings = []
        self.property_errors = {}
        self.folders = []
        self.messages = []  # (folder path, subject)
        self.items = []
        self.fail_on = set(fail_on)
        self.closed = False

    def get_folder(self, parts, types=None):
        self.folders.append((tuple(parts), tuple(types or ())))
        return tuple(parts)

    def add_message(self, folder, m):
        if m.subject in self.fail_on:
            raise RuntimeError("Outlook said no\nsecond line")
        self.messages.append((folder, m.subject))

    def add_item(self, folder, kind, m, props):
        self.items.append((folder, kind, m.subject))

    def close(self, quit_outlook=False, timeout=None):
        self.closed = True


class TroubledReader:
    """A reader for the reader-process tests: item 2 of any folder kills the
    process, item 3 hangs, folder 99 doesn't exist."""

    def __init__(self, path):
        if path.endswith("unopenable"):
            raise OSError("no such mailbox")

    def scan(self):
        return ["tree"]

    def has_folder(self, identifier):
        return identifier != 99

    def read(self, identifier, i):
        if i == 2:
            os._exit(3)  # as a crash inside the C library would
        if i == 3:
            time.sleep(120)
        if i == 4:
            raise ValueError("bad item\nmore detail")
        return f"message {i}", None

    def close(self):
        pass


class AlwaysCrashes(TroubledReader):
    def read(self, identifier, i):
        os._exit(3)


class DiesWhileOpening:
    """A reader whose process ends while opening the file, as one killed by
    a virus scanner or by running out of memory would."""

    def __init__(self, path):
        os._exit(3)


class HangsWhileOpening:
    """A reader that never finishes opening the file."""

    def __init__(self, path):
        time.sleep(120)


class SilentError(Exception):
    """An exception with no message, like MemoryError()."""

    def __str__(self):
        return ""


class FailsSilently(FakeWriter):
    def add_message(self, folder, m):
        if m.subject in self.fail_on:
            raise SilentError()
        self.messages.append((folder, m.subject))
