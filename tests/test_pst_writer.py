"""The parts of the Outlook writer that don't need Outlook."""

import unittest
from datetime import datetime, timezone

from unbound.props import NamedKey, Prop
from unbound.pst_writer import (
    _is_smtp, _safe_filename, com_date, from_com_date, recipient_text, should_copy,
)


class SafeFilenames(unittest.TestCase):
    def test_ordinary_names_are_kept(self):
        for name in ("report.pdf", "Quarterly report (final).xlsx", "résumé 王.docx"):
            self.assertEqual(_safe_filename(name), name)

    def test_cannot_leave_the_temporary_folder(self):
        for name in ("..\\..\\evil.exe", "../../evil.exe", "C:\\Windows\\x.dll", "a:b", "x\x00y", ".."):
            safe = _safe_filename(name)
            self.assertNotRegex(safe, r'[\\/:\x00]')
            self.assertFalse(safe.startswith("."))
            self.assertTrue(safe)

    def test_empty_names_get_a_name(self):
        for name in ("", "   ", "...", "."):
            self.assertEqual(_safe_filename(name), "attachment")

    def test_device_names_are_not_used_as_file_names(self):
        for name in ("CON", "nul.txt", "Aux.pdf", "COM1", "lpt9.doc", "PRN.tar.gz"):
            self.assertEqual(_safe_filename(name), "_" + name)
        for name in ("console.txt", "COM10.txt", "nullable.pdf"):
            self.assertEqual(_safe_filename(name), name)

    def test_long_names_keep_their_extension(self):
        safe = _safe_filename("a" * 300 + ".pdf")
        self.assertEqual(len(safe), 150)
        self.assertTrue(safe.endswith(".pdf"))
        self.assertEqual(len(_safe_filename("b" * 300)), 150)


class Recipients(unittest.TestCase):
    def test_name_and_address(self):
        self.assertEqual(recipient_text("Bob Smith", "bob@example.test"),
                         ('"Bob Smith" <bob@example.test>', True))
        self.assertEqual(recipient_text("", "bob@example.test"), ("bob@example.test", True))

    def test_quote_in_a_name_cannot_break_the_address(self):
        text, resolve = recipient_text('Smith, "JJ" John', "jj@example.test")
        self.assertEqual(text, "\"Smith, 'JJ' John\" <jj@example.test>")
        self.assertEqual(text.count('"'), 2)
        self.assertTrue(resolve)

    def test_control_characters_are_removed_from_names(self):
        self.assertEqual(recipient_text("Bob\r\n\tSmith", "b@example.test")[0],
                         '"Bob Smith" <b@example.test>')

    def test_name_alone_is_not_looked_up_in_this_pcs_address_book(self):
        self.assertEqual(recipient_text("John", ""), ("John", False))
        self.assertEqual(recipient_text("John", None), ("John", False))

    def test_exchange_directory_address_is_not_treated_as_an_email_address(self):
        dn = "/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=abc123-john"
        self.assertEqual(recipient_text("John Doe", dn), ("John Doe", False))
        self.assertEqual(recipient_text("", dn), (dn, False))

    def test_nothing_to_add(self):
        self.assertEqual(recipient_text("", ""), (None, False))
        self.assertEqual(recipient_text(None, None), (None, False))

    def test_what_counts_as_an_internet_address(self):
        self.assertTrue(_is_smtp("a@b.c"))
        for address in ("", None, "/o=Org/cn=a@b", "MAILER-DAEMON"):
            self.assertFalse(_is_smtp(address))


class Dates(unittest.TestCase):
    def test_dates_are_marked_utc_and_come_back_unchanged(self):
        dt = datetime(2026, 3, 15, 9, 30)
        self.assertEqual(com_date(dt), dt.replace(tzinfo=timezone.utc))
        self.assertEqual(from_com_date(com_date(dt)), dt)
        self.assertIsNone(com_date(None))
        self.assertIsNone(from_com_date(None))


class WhichPropertiesAreCopied(unittest.TestCase):
    def test_skips_what_outlook_computes_itself(self):
        self.assertFalse(should_copy(Prop(0x0037, 0x1F, None, "subject")))
        self.assertFalse(should_copy(Prop(0x6610, 0x03, None, 1)))     # sync bookkeeping
        self.assertFalse(should_copy(Prop(0x8123, 0x03, None, 1)))     # named but unresolved

    def test_copies_ordinary_and_named_properties(self):
        self.assertTrue(should_copy(Prop(0x3A06, 0x1F, None, "Given name")))
        named = NamedKey("{00062004-0000-0000-C000-000000000046}", lid=0x8083)
        self.assertTrue(should_copy(Prop(0x8123, 0x1F, named, "a@b.c")))


if __name__ == "__main__":
    unittest.main()
