"""Stand-ins for the Outlook objects PstWriter talks to, behaving the way
Outlook was seen to behave where it matters (see the test docstrings)."""

import os
import unicodedata


class OutlookError(Exception):
    pass


def _key(name):
    return unicodedata.normalize("NFC", name).casefold()


class Folders:
    def __init__(self):
        self._items = []

    def __iter__(self):
        return iter(list(self._items))

    def Add(self, name, kind=None):
        # Outlook won't create a folder whose name matches an existing one
        # apart from case or how accents are encoded
        if any(_key(f.Name) == _key(name) for f in self._items):
            raise OutlookError("Cannot create the folder.")
        folder = Folder(name, kind)
        self._items.append(folder)
        return folder


class Folder:
    def __init__(self, name, kind=None):
        self.Name = name
        self.kind = kind
        self.Folders = Folders()


class Store:
    def __init__(self, path, broken=False):
        self._path = path
        self.broken = broken
        self.root = Folder(os.path.basename(path or "store"))

    @property
    def FilePath(self):
        if self.broken:
            raise OutlookError("OLE error 0x8624011d")  # a store Outlook can't open
        return self._path

    def GetRootFolder(self):
        return self.root


class Stores:
    """Item(i) is 1-based. Enumerating with `for` fails, as Outlook's own
    enumeration does once it reaches a store it can't open."""

    def __init__(self, stores):
        self.list = stores

    @property
    def Count(self):
        return len(self.list)

    def Item(self, i):
        return self.list[i - 1]

    def __iter__(self):
        for s in self.list:
            if s.broken:
                raise OutlookError("OLE error 0x8624011d")
            yield s


class Namespace:
    def __init__(self, existing=()):
        self.Stores = Stores(list(existing))
        self.removed = []
        self.items = {}

    def AddStoreEx(self, path, kind):
        self.Stores.list.append(Store(path))

    def RemoveStore(self, root):
        store = next(s for s in self.Stores.list if s.root is root)
        self.Stores.list.remove(store)
        self.removed.append(store._path)

    def GetItemFromID(self, entry_id, store_id=None):
        return self.items[entry_id]


class App:
    def __init__(self, ns):
        self.ns = ns

    def GetNamespace(self, name):
        return self.ns


class Pattern:
    """A RecurrencePattern that records what's set. Like Outlook, a yearly
    pattern refuses an Interval below 12: it counts it in months."""

    def __init__(self, derived=None):
        # derived: what Outlook works out from the start once RecurrenceType
        # is set, and then refuses to be told again (seen for a yearly series
        # on 29 February: day 29, month 2, interval 12)
        object.__setattr__(self, "set", {})
        object.__setattr__(self, "derived", derived or {})

    def __getattr__(self, name):
        if self.set.get("RecurrenceType") is not None and name in self.derived:
            return self.derived[name]
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name in self.derived and self.set.get("RecurrenceType") is not None:
            raise OutlookError("The recurrence pattern is not valid.")
        if name == "Interval" and self.set.get("RecurrenceType") in (5, 6) and value < 12:
            raise OutlookError("The recurrence pattern is not valid.")
        self.set[name] = value


class Recipient:
    def __init__(self, text, refuse_type=False):
        self.text = text
        self.refuse_type = refuse_type
        self.resolved = False
        self.type = None

    @property
    def Type(self):
        return self.type

    @Type.setter
    def Type(self, value):
        if self.refuse_type:
            raise OutlookError("The parameter is incorrect.")
        self.type = value

    def Resolve(self):
        self.resolved = True


class Recipients:
    def __init__(self, refuse_type=False):
        self.added = []
        self.refuse_type = refuse_type

    def Add(self, text):
        rec = Recipient(text, self.refuse_type)
        self.added.append(rec)
        return rec


class PropertyAccessor:
    def __init__(self):
        self.values = {}

    def SetProperty(self, url, value):
        self.values[url] = value


class Item:
    """An item whose Save can be made to fail, and which records Delete."""

    def __init__(self, entry_id="E1", fail_save=False, refuse_rtf=False):
        object.__setattr__(self, "values", {})
        self.EntryID = entry_id
        self.fail_save = fail_save
        self.refuse_rtf = refuse_rtf
        self.saved = 0
        self.deleted = False
        self.Recipients = Recipients()
        self.PropertyAccessor = PropertyAccessor()

    def __setattr__(self, name, value):
        if name == "RTFBody" and self.__dict__.get("refuse_rtf"):
            raise OutlookError("The parameter is incorrect.")
        object.__setattr__(self, name, value)

    def Save(self):
        if self.fail_save:
            raise OutlookError("Outlook refused to save")
        self.saved += 1

    def Delete(self):
        """Like Outlook in a PST: the item goes to Deleted Items (a new
        object there), unless it's already in it."""
        self.deleted = True
        bin_ = self.__dict__.get("deleted_items")
        if bin_ is not None:
            copy = Item(self.EntryID + "-in-bin")
            copy.Subject = getattr(self, "Subject", "")
            bin_.list.append(copy)


class DeletedItems:
    """A Deleted Items folder whose Items.Find matches on Subject."""

    def __init__(self):
        self.list = []
        self.Items = self

    def Find(self, restriction):
        for item in self.list:
            if restriction == f"[Subject] = '{item.Subject}'":
                return item
        return None
