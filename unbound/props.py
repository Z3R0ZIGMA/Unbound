"""Read every MAPI property of a pypff item, with named properties resolved.

Outlook stores most calendar and contact fields as *named* properties: the
item only holds a number (0x8000 and up) that's specific to this file, and
the file's name-to-ID map says which (property set GUID, id or name) it
stands for. pypff exposes that map only as raw data, so it's decoded here
following [MS-PST] 2.4.7.
"""

import struct
import uuid
from dataclasses import dataclass

PS_MAPI = "{00020328-0000-0000-C000-000000000046}"
PS_PUBLIC_STRINGS = "{00020329-0000-0000-C000-000000000046}"

# Well-known property sets
PSETID_APPOINTMENT = "{00062002-0000-0000-C000-000000000046}"
PSETID_ADDRESS = "{00062004-0000-0000-C000-000000000046}"
PSETID_COMMON = "{00062008-0000-0000-C000-000000000046}"
PSETID_TASK = "{00062003-0000-0000-C000-000000000046}"
PSETID_MEETING = "{6ED8DA90-450B-101B-98DA-00AA003F1305}"
PS_INTERNET_HEADERS = "{00020386-0000-0000-C000-000000000046}"

# Property types
PT_SHORT, PT_LONG, PT_FLOAT, PT_DOUBLE, PT_CURRENCY = 0x2, 0x3, 0x4, 0x5, 0x6
PT_APPTIME, PT_ERROR, PT_BOOLEAN, PT_OBJECT, PT_I8 = 0x7, 0xA, 0xB, 0xD, 0x14
PT_STRING8, PT_UNICODE, PT_SYSTIME, PT_CLSID, PT_BINARY = 0x1E, 0x1F, 0x40, 0x48, 0x102
MV_FLAG = 0x1000


@dataclass(frozen=True)
class NamedKey:
    guid: str  # "{...}" upper case
    lid: int | None = None  # numeric name
    name: str | None = None  # string name

    def url(self, prop_type):
        """PropertyAccessor schema name for this property."""
        base = "http://schemas.microsoft.com/mapi/"
        if self.lid is not None:
            return f"{base}id/{self.guid}/{self.lid:04X}{prop_type:04X}"
        return f"{base}string/{self.guid}/{self.name}"


def load_name_map(pff_file):
    """{0x8000+ prop id: NamedKey} for this file."""
    streams = {}
    for rs in pff_file.name_to_id_map.record_sets:
        for e in rs.entries:
            if e.entry_type in (0x2, 0x3, 0x4):
                streams[e.entry_type] = e.data or b""
    guids, entries, strings = streams.get(0x2, b""), streams.get(0x3, b""), streams.get(0x4, b"")

    def guid_at(index):
        if index == 1:
            return PS_MAPI
        if index == 2:
            return PS_PUBLIC_STRINGS
        off = (index - 3) * 16
        if index < 3 or off + 16 > len(guids):
            return None
        return "{" + str(uuid.UUID(bytes_le=guids[off:off + 16])).upper() + "}"

    result = {}
    for off in range(0, len(entries) - 7, 8):
        prop_id_or_offset, guid_and_kind, prop_index = struct.unpack_from("<IHH", entries, off)
        is_string = guid_and_kind & 1
        guid = guid_at(guid_and_kind >> 1)
        if guid is None:
            continue
        if is_string:
            if prop_id_or_offset + 4 > len(strings):
                continue
            (length,) = struct.unpack_from("<I", strings, prop_id_or_offset)
            raw = strings[prop_id_or_offset + 4:prop_id_or_offset + 4 + length]
            key = NamedKey(guid, name=raw.decode("utf-16-le", errors="replace"))
        else:
            key = NamedKey(guid, lid=prop_id_or_offset)
        result[0x8000 + prop_index] = key
    return result


def _value(entry):
    """Python value for a record entry, or raise ValueError if unsupported."""
    t = entry.value_type
    data = entry.data
    if data is None:
        raise ValueError("no data")
    if t in (PT_SHORT,):
        return struct.unpack("<h", data[:2])[0]
    if t == PT_LONG:
        return struct.unpack("<i", data[:4])[0]
    if t == PT_BOOLEAN:
        return bool(data[0])
    if t == PT_DOUBLE or t == PT_APPTIME:
        return struct.unpack("<d", data[:8])[0]
    if t == PT_FLOAT:
        return struct.unpack("<f", data[:4])[0]
    if t == PT_I8 or t == PT_CURRENCY:
        return struct.unpack("<q", data[:8])[0]
    if t == PT_SYSTIME:
        return entry.data_as_datetime  # naive UTC
    if t in (PT_UNICODE, PT_STRING8):
        return entry.data_as_string
    if t in (PT_BINARY, PT_CLSID):
        return bytes(data)
    if t == PT_UNICODE | MV_FLAG:
        return _mv_strings(data)
    if t == PT_LONG | MV_FLAG:
        n = len(data) // 4
        return list(struct.unpack(f"<{n}i", data[:n * 4]))
    if t == PT_BINARY | MV_FLAG:
        return _mv_binary(data)
    raise ValueError(f"type 0x{t:04X} not supported")


def _mv_strings(data):
    # [MS-PST] 2.3.3.4.2: count, offsets, then UTF-16 strings back to back
    (count,) = struct.unpack_from("<I", data, 0)
    offsets = list(struct.unpack_from(f"<{count}I", data, 4)) + [len(data)]
    return [data[offsets[i]:offsets[i + 1]].decode("utf-16-le").rstrip("\x00")
            for i in range(count)]


def _mv_binary(data):
    (count,) = struct.unpack_from("<I", data, 0)
    offsets = list(struct.unpack_from(f"<{count}I", data, 4)) + [len(data)]
    return [bytes(data[offsets[i]:offsets[i + 1]]) for i in range(count)]


@dataclass
class Prop:
    tag_id: int  # property id as stored in this file
    prop_type: int
    named: NamedKey | None
    value: object

    def url(self):
        if self.named:
            return self.named.url(self.prop_type)
        return f"http://schemas.microsoft.com/mapi/proptag/0x{self.tag_id:04X}{self.prop_type:04X}"


def read_props(item, name_map):
    """All readable properties of a pypff item (first record set)."""
    props = []
    try:
        entries = item.record_sets[0].entries
    except (OSError, IndexError):
        return props
    for e in entries:
        try:
            value = _value(e)
        except (ValueError, OSError, struct.error, TypeError, UnicodeDecodeError,
                IndexError, OverflowError):
            continue  # one damaged value is skipped, not the whole item
        named = name_map.get(e.entry_type) if e.entry_type >= 0x8000 else None
        if e.entry_type >= 0x8000 and named is None:
            continue  # can't say what it is, so can't recreate it
        props.append(Prop(e.entry_type, e.value_type, named, value))
    return props
