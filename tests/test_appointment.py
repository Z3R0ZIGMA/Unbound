import random
import struct
import unittest
from datetime import datetime

from unbound import appointment as A
from unbound.props import PSETID_APPOINTMENT, NamedKey, Prop


def minutes(dt):
    return int((dt - A.EPOCH_1601).total_seconds() // 60)


def wide(text):
    return struct.pack("<H", len(text)) + text.encode("utf-16-le")


def narrow(text):
    data = text.encode("cp1252")
    return struct.pack("<HH", len(data) + 1, len(data)) + data


def weekly_blob(writer2=0x3009, exceptions=True):
    """A weekly series (Mon + Wed, every 2 weeks, 10 times) with one cancelled
    occurrence and one moved and renamed one, laid out per [MS-OXOCAL]."""
    start_day = datetime(2026, 3, 2)
    deleted = [datetime(2026, 3, 4), datetime(2026, 3, 16)] if exceptions else []
    modified = [datetime(2026, 3, 16)] if exceptions else []
    b = struct.pack("<HHHHHIII", 0x3004, 0x3004, A.FREQ_WEEKLY, A.PT_WEEK, 0, 0, 2, 0)
    b += struct.pack("<I", 0b0001010)  # Monday + Wednesday
    b += struct.pack("<IIII", A.END_AFTER_N, 10, 1, len(deleted))
    b += b"".join(struct.pack("<I", minutes(d)) for d in deleted)
    b += struct.pack("<I", len(modified))
    b += b"".join(struct.pack("<I", minutes(d)) for d in modified)
    b += struct.pack("<II", minutes(start_day), minutes(datetime(2026, 5, 6)))
    b += struct.pack("<IIIIH", 0x3006, writer2, 9 * 60, 10 * 60, len(modified))
    if exceptions:
        new_start, new_end = datetime(2026, 3, 16, 14, 0), datetime(2026, 3, 16, 15, 30)
        original = datetime(2026, 3, 16, 9, 0)
        flags = A.ARO_SUBJECT | A.ARO_REMINDERDELTA | A.ARO_LOCATION | A.ARO_BUSYSTATUS
        b += struct.pack("<IIIH", minutes(new_start), minutes(new_end), minutes(original), flags)
        b += narrow("Moved?") + struct.pack("<I", 30) + narrow("Room 2") + struct.pack("<I", 2)
    b += struct.pack("<I", 0)  # reserved block 1
    if exceptions:
        if writer2 >= 0x3009:
            b += struct.pack("<I", 4) + b"\x00" * 4  # change highlight
        b += struct.pack("<I", 0)  # reserved block
        b += struct.pack("<III", minutes(new_start), minutes(new_end), minutes(original))
        b += wide("Moved – Zürich ☕") + wide("Room 2 – 王")
        b += struct.pack("<I", 0)
    b += struct.pack("<I", 0)
    return b


class Recurrence(unittest.TestCase):
    def test_weekly_series_with_changed_occurrences(self):
        for writer2 in (0x3008, 0x3009):
            with self.subTest(writer2=hex(writer2)):
                r = A.parse_recurrence(weekly_blob(writer2))
                self.assertEqual(r.outlook_type(), A.OL_WEEKLY)
                self.assertEqual(r.interval(), 2)
                self.assertEqual(r.day_of_week_mask, 0b0001010)
                self.assertEqual((r.end_type, r.occurrences), (A.END_AFTER_N, 10))
                self.assertEqual(r.start_date, datetime(2026, 3, 2))
                self.assertEqual((r.start_offset, r.end_offset), (540, 600))
                self.assertEqual(r.deleted, [datetime(2026, 3, 4), datetime(2026, 3, 16)])
                (x,) = r.exceptions
                self.assertEqual(x.original_start, datetime(2026, 3, 16, 9, 0))
                self.assertEqual((x.start, x.end),
                                 (datetime(2026, 3, 16, 14, 0), datetime(2026, 3, 16, 15, 30)))
                # the Unicode strings of the extended part replace the cp1252 ones
                self.assertEqual((x.subject, x.location), ("Moved – Zürich ☕", "Room 2 – 王"))
                self.assertEqual((x.reminder_minutes, x.busy_status), (30, 2))

    def test_series_without_exceptions(self):
        r = A.parse_recurrence(weekly_blob(exceptions=False))
        self.assertEqual((r.deleted, r.exceptions), ([], []))

    def test_intervals(self):
        self.assertEqual(A.Recurrence(A.FREQ_DAILY, A.PT_DAY, 1440 * 3).interval(), 3)
        self.assertEqual(A.Recurrence(A.FREQ_YEARLY, A.PT_MONTH, 24).interval(), 2)
        self.assertEqual(A.Recurrence(A.FREQ_MONTHLY, A.PT_MONTH, 0).interval(), 1)

    def test_outlook_types(self):
        for freq, pattern, expected in ((A.FREQ_DAILY, A.PT_DAY, A.OL_DAILY),
                                        (A.FREQ_MONTHLY, A.PT_MONTH_END, A.OL_MONTHLY),
                                        (A.FREQ_YEARLY, A.PT_MONTH, A.OL_YEARLY),
                                        (A.FREQ_MONTHLY, A.PT_MONTH_NTH, A.OL_MONTH_NTH),
                                        (A.FREQ_YEARLY, A.PT_MONTH_NTH, A.OL_YEAR_NTH),
                                        (A.FREQ_MONTHLY, 0xA, None)):  # Hijri
            self.assertEqual(A.Recurrence(freq, pattern, 1).outlook_type(), expected)

    def test_unknown_version_is_refused(self):
        with self.assertRaises(ValueError):
            A.parse_recurrence(struct.pack("<H", 0x1234) + weekly_blob()[2:])

    def test_damaged_data_only_ever_raises_value_error(self):
        """Whatever a damaged file holds, the caller gets ValueError (the item
        is then reported as failed), never a crash or another exception."""
        good = weekly_blob()
        rng = random.Random(1601)
        samples = [good[:n] for n in range(len(good))]
        for _ in range(3000):
            b = bytearray(good)
            for _ in range(rng.randint(1, 6)):
                b[rng.randrange(len(b))] = rng.randrange(256)
            samples.append(bytes(b[:rng.randint(0, len(b))] if rng.random() < 0.3 else b))
        samples += [bytes(rng.randrange(256) for _ in range(rng.randint(0, 200))) for _ in range(500)]
        for blob in samples:
            try:
                A.parse_recurrence(blob)
            except ValueError:
                pass


class TimeZones(unittest.TestCase):
    def test_key_name_is_read_from_a_definition(self):
        name = "South Africa Standard Time"
        blob = struct.pack("<BBHHH", 2, 1, 0, 2, len(name)) + name.encode("utf-16-le") + b"\x00" * 8
        self.assertEqual(A.tz_key_name(blob), name)

    def test_missing_or_damaged_definitions(self):
        self.assertIsNone(A.tz_key_name(None))
        self.assertIsNone(A.tz_key_name(b"\x02\x01"))
        self.assertIsNone(A.tz_key_name(struct.pack("<BBHHH", 2, 1, 0, 2, 1) + b"\x00\xd8"))


class ManagedProperties(unittest.TestCase):
    def test_times_and_recurrence_go_through_outlooks_own_fields(self):
        start = Prop(0x8005, 0x40, NamedKey(PSETID_APPOINTMENT, lid=A.LID_START), datetime(2026, 1, 1))
        location = Prop(0x8006, 0x1F, NamedKey(PSETID_APPOINTMENT, lid=0x8208), "Room")
        self.assertTrue(A.is_managed(start))
        self.assertFalse(A.is_managed(location))
        self.assertTrue(A.is_managed(Prop(0x0060, 0x40, None, datetime(2026, 1, 1))))
        self.assertEqual(A.appointment_named([start, location]),
                         {A.LID_START: datetime(2026, 1, 1), 0x8208: "Room"})


if __name__ == "__main__":
    unittest.main()
