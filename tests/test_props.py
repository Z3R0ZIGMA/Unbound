import random
import struct
import unittest
import uuid

from unbound import props
from unbound.props import NamedKey, load_name_map, read_props


class Entry:
    def __init__(self, entry_type, value_type=0, data=b"", text=None):
        self.entry_type = entry_type
        self.value_type = value_type
        self.data = data
        self.data_as_string = text


def record_sets(entries):
    return [type("RecordSet", (), {"entries": entries})]


def pff_file(guids=b"", entries=b"", strings=b""):
    """Looks enough like a pypff file for the name-to-ID map."""
    name_map = type("Map", (), {"record_sets": record_sets(
        [Entry(0x2, data=guids), Entry(0x3, data=entries), Entry(0x4, data=strings)])})
    return type("File", (), {"name_to_id_map": name_map})


class NameMap(unittest.TestCase):
    def test_numbered_and_named_properties(self):
        guid = uuid.UUID("00062002-0000-0000-C000-000000000046")
        name = "Keywords".encode("utf-16-le")
        strings = struct.pack("<I", len(name)) + name
        entries = (struct.pack("<IHH", 0x820D, (3 << 1) | 0, 0)       # id in the first listed set
                   + struct.pack("<IHH", 0, (2 << 1) | 1, 1)           # a name, in PS_PUBLIC_STRINGS
                   + struct.pack("<IHH", 0x8501, (1 << 1) | 0, 2)      # id in PS_MAPI
                   + struct.pack("<IHH", 0x1234, (9 << 1) | 0, 3))     # a set that isn't listed
        result = load_name_map(pff_file(guid.bytes_le, entries, strings))
        self.assertEqual(result, {
            0x8000: NamedKey("{00062002-0000-0000-C000-000000000046}", lid=0x820D),
            0x8001: NamedKey(props.PS_PUBLIC_STRINGS, name="Keywords"),
            0x8002: NamedKey(props.PS_MAPI, lid=0x8501)})

    def test_addresses_outlook_understands(self):
        key = NamedKey("{00062002-0000-0000-C000-000000000046}", lid=0x820D)
        self.assertEqual(key.url(0x40), "http://schemas.microsoft.com/mapi/id/"
                                        "{00062002-0000-0000-C000-000000000046}/820D0040")
        named = NamedKey(props.PS_PUBLIC_STRINGS, name="Keywords")
        self.assertEqual(named.url(0x101F), "http://schemas.microsoft.com/mapi/string/"
                                            "{00020329-0000-0000-C000-000000000046}/Keywords")

    def test_damaged_map_never_raises(self):
        rng = random.Random(7)
        for _ in range(500):
            blobs = [bytes(rng.randrange(256) for _ in range(rng.randint(0, 80))) for _ in range(3)]
            load_name_map(pff_file(*blobs))


class Values(unittest.TestCase):
    def read(self, *entries, name_map=None):
        item = type("Item", (), {"record_sets": record_sets(list(entries))})
        return read_props(item, name_map or {})

    def test_simple_types(self):
        got = self.read(Entry(0x3A06, props.PT_UNICODE, b"x", "Given"),
                        Entry(0x0E07, props.PT_LONG, struct.pack("<i", -5)),
                        Entry(0x0E1F, props.PT_BOOLEAN, b"\x01"),
                        Entry(0x3A40, props.PT_BINARY, b"\x00\x01\x02"),
                        Entry(0x1000, props.PT_DOUBLE, struct.pack("<d", 1.5)))
        self.assertEqual([(p.tag_id, p.value) for p in got],
                         [(0x3A06, "Given"), (0x0E07, -5), (0x0E1F, True),
                          (0x3A40, b"\x00\x01\x02"), (0x1000, 1.5)])
        self.assertEqual(got[0].url(), "http://schemas.microsoft.com/mapi/proptag/0x3A06001F")

    def test_lists_of_strings_and_binaries(self):
        a, b = "one".encode("utf-16-le"), "två\x00".encode("utf-16-le")
        strings = struct.pack("<III", 2, 12, 12 + len(a)) + a + b
        binaries = struct.pack("<III", 2, 12, 14) + b"\x01\x02" + b"\x03"
        got = self.read(Entry(0x3A58, props.PT_UNICODE | props.MV_FLAG, strings),
                        Entry(0x3A59, props.PT_BINARY | props.MV_FLAG, binaries),
                        Entry(0x3A5A, props.PT_LONG | props.MV_FLAG, struct.pack("<ii", 7, -1)))
        self.assertEqual([p.value for p in got], [["one", "två"], [b"\x01\x02", b"\x03"], [7, -1]])

    def test_named_property_needs_an_entry_in_the_map(self):
        key = NamedKey("{00062004-0000-0000-C000-000000000046}", lid=0x8083)
        got = self.read(Entry(0x8010, props.PT_UNICODE, b"x", "a@b.c"),
                        Entry(0x8011, props.PT_UNICODE, b"x", "unknown"),
                        name_map={0x8010: key})
        self.assertEqual([(p.named, p.value) for p in got], [(key, "a@b.c")])

    def test_unreadable_values_are_skipped_not_fatal(self):
        got = self.read(Entry(0x3A06, props.PT_LONG, b"\x01"),                 # too short
                        Entry(0x3A07, 0x00FB, b"1234"),                        # unknown type
                        Entry(0x3A08, props.PT_UNICODE, None),                 # no data
                        Entry(0x3A58, props.PT_UNICODE | props.MV_FLAG, struct.pack("<I", 2**32 - 1)),
                        Entry(0x3A59, props.PT_BINARY | props.MV_FLAG, b"\x05"),
                        Entry(0x3A09, props.PT_LONG, struct.pack("<i", 9)))
        self.assertEqual([(p.tag_id, p.value) for p in got], [(0x3A09, 9)])

    def test_item_without_properties(self):
        self.assertEqual(read_props(type("Item", (), {"record_sets": []}), {}), [])


if __name__ == "__main__":
    unittest.main()
