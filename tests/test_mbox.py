import os
import tempfile
import unittest

from unbound import mapi, mbox_reader
from unbound.mbox_reader import (
    find_mailboxes, is_mbox_source, message_offsets, parse_message, read_raw,
)

SEPARATOR = b"From sender@example.test Mon Mar  2 09:15:00 2026\n"


def mbox(*messages, separator=SEPARATOR):
    return b"".join(separator + m + b"\n" for m in messages)


class TempDirTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = tmp.name

    def write(self, name, data):
        path = os.path.join(self.dir, name)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(data)
        return path


class Splitting(TempDirTest):
    def subjects(self, path):
        return [parse_message(read_raw(path, i)).subject
                for i in range(len(message_offsets(path)))]

    def test_splits_on_separator_lines(self):
        path = self.write("a.mbox", mbox(b"Subject: one\n\nhello\n", b"Subject: two\n\nbye\n"))
        self.assertEqual(self.subjects(path), ["one", "two"])

    def test_body_line_starting_with_from_does_not_split(self):
        # Written by a program that doesn't quote such lines as ">From"
        path = self.write("a.mbox", mbox(
            b"Subject: one\n\nhello\n\nFrom here on, we agree.\nbye\n",
            b"Subject: two\n\nsecond\n"))
        self.assertEqual(self.subjects(path), ["one", "two"])
        body = parse_message(read_raw(path, 0)).plain_body
        self.assertIn("From here on, we agree.", body)
        self.assertIn("bye", body)

    def test_separator_forms_of_other_mail_programs(self):
        for separator in (b"From - Tue Oct 12 08:01:02 2021\n",                  # Thunderbird
                          b"From 17000000@xxx Mon Jan 01 00:00:00 +0000 2024\n",  # Google Takeout
                          b"From MAILER-DAEMON Fri Jul  8 12:08:34 2011\n",
                          b"From a@b.c Mon, 1 Jan 2024 10:00:00 +0200\r\n"):
            with self.subTest(separator=separator):
                path = self.write("a.mbox", mbox(b"Subject: one\n\nx\n", b"Subject: two\n\ny\n",
                                                 separator=separator))
                self.assertEqual(self.subjects(path), ["one", "two"])
                mbox_reader._offsets_cache.clear()

    def test_file_without_dated_separators_still_splits(self):
        path = self.write("a.mbox", mbox(b"Subject: one\n\nx\n", b"Subject: two\n\ny\n",
                                         separator=b"From someone\n"))
        self.assertEqual(self.subjects(path), ["one", "two"])

    def test_from_line_needs_a_blank_line_before_it(self):
        path = self.write("a.mbox", mbox(b"Subject: one\n\nline\n" + SEPARATOR.rstrip() + b"\nend\n"))
        self.assertEqual(self.subjects(path), ["one"])

    def test_quoted_from_lines_are_restored(self):
        path = self.write("a.mbox", mbox(b"Subject: one\n\n>From the start\n>>From deeper\n"))
        body = parse_message(read_raw(path, 0)).plain_body
        self.assertIn("\nFrom the start", "\n" + body)
        self.assertIn(">From deeper", body)

    def test_empty_file(self):
        path = self.write("a.mbox", b"")
        self.assertEqual(message_offsets(path), [])


class Parsing(unittest.TestCase):
    def test_headers_recipients_and_dates(self):
        m = parse_message(
            b"From: =?UTF-8?Q?Alice_M=C3=BCller?= <alice@example.test>\n"
            b"To: Bob <bob@example.test>, carol@example.test\n"
            b"Cc: \"Smith, \\\"JJ\\\" John\" <jj@example.test>\n"
            b"Subject: =?UTF-8?B?w4RwZmVsIHVuZCBCaXJuZW4=?=\n"
            b"Date: Mon, 02 Mar 2026 09:15:00 +0200\n"
            b"Status: RO\n\nBody text\n")
        self.assertEqual(m.subject, "Äpfel und Birnen")
        self.assertEqual((m.sender_name, m.sender_email), ("Alice Müller", "alice@example.test"))
        self.assertEqual([(r.name, r.email, r.kind) for r in m.recipients], [
            ("Bob", "bob@example.test", mapi.MAPI_TO),
            ("", "carol@example.test", mapi.MAPI_TO),
            ('Smith, "JJ" John', "jj@example.test", mapi.MAPI_CC)])
        self.assertEqual(m.submit_time.isoformat(), "2026-03-02T07:15:00")  # UTC
        self.assertTrue(m.is_read)
        self.assertEqual(m.plain_body.strip(), "Body text")
        self.assertIsNone(m.html_body)
        self.assertEqual(m.message_class, "IPM.Note")

    def test_read_state(self):
        unread = parse_message(b"Subject: x\nStatus: O\n\nbody\n")
        no_header = parse_message(b"Subject: x\n\nbody\n")  # as Apple Mail exports
        self.assertFalse(unread.is_read)
        self.assertTrue(no_header.is_read)

    def test_attachments_inline_images_and_attached_email(self):
        raw = (
            b"Subject: parts\nMIME-Version: 1.0\n"
            b"Content-Type: multipart/mixed; boundary=XX\n\n"
            b"--XX\nContent-Type: text/html; charset=utf-8\n\n<p>hi <img src=\"cid:logo@x\"></p>\n"
            b"--XX\nContent-Type: image/png\nContent-ID: <logo@x>\n"
            b"Content-Transfer-Encoding: base64\n\niVBORw0KGgo=\n"
            b"--XX\nContent-Type: application/pdf\n"
            b"Content-Disposition: attachment; filename=\"..\\\\..\\\\report.pdf\"\n"
            b"Content-Transfer-Encoding: base64\n\nJVBERi0=\n"
            b"--XX\nContent-Type: message/rfc822\n\nSubject: inner one\n\ninner body\n"
            b"--XX--\n")
        m = parse_message(raw)
        self.assertIn("cid:logo@x", m.html_body)
        names = [(a.filename, a.content_id) for a in m.attachments]
        self.assertEqual(names, [("attachment_1.png", "logo@x"),
                                 ("..\\..\\report.pdf", None),
                                 ("inner one.eml", None)])
        self.assertEqual(m.attachments[0].data, b"\x89PNG\r\n\x1a\n")
        self.assertEqual(m.attachments[1].data, b"%PDF-")
        self.assertIn(b"inner body", m.attachments[2].data)

    def test_headers_end_at_the_first_blank_line_with_windows_line_ends(self):
        m = parse_message(b"Subject: x\r\nTo: a@example.test\r\n\r\nthe body\r\n")
        self.assertNotIn("the body", m.transport_headers)
        self.assertIn("Subject: x", m.transport_headers)

    def test_broken_input_does_not_raise(self):
        for raw in (b"", b"\n\n", b"no headers at all", b"Subject: \xff\xfe\n\n\xff",
                    b"Content-Type: multipart/mixed; boundary=Q\n\n--Q\n",
                    b"Date: not a date\nFrom: <<>>\nTo: ;;;\n\nx"):
            with self.subTest(raw=raw):
                parse_message(raw)


class FindingMailboxes(TempDirTest):
    def test_apple_mail_export_folder(self):
        self.write("Inbox.mbox/mbox", mbox(b"Subject: a\n\nx\n"))
        self.write("Inbox.mbox/Info.plist", b"<plist/>")
        self.write("Projects.mbox/Alpha.mbox/mbox", mbox(b"Subject: b\n\nx\n"))  # parent has no mail
        self.write("loose.mbox", mbox(b"Subject: c\n\nx\n"))
        self.write("notes.txt", b"ignored")
        boxes = find_mailboxes(self.dir)
        self.assertEqual([b.name for b in boxes], ["Inbox", "loose", "Projects"])
        projects = boxes[2]
        self.assertIsNone(projects.file)
        self.assertEqual([(c.name, c.path) for c in projects.children],
                         [("Alpha", ("Projects", "Alpha"))])

    def test_what_counts_as_an_mbox_source(self):
        export = self.write("Inbox.mbox/mbox", mbox(b"Subject: a\n\nx\n"))
        ost = self.write("mail.ost", b"!BDN" + b"\x00" * 60)
        self.assertTrue(is_mbox_source(self.dir))
        self.assertTrue(is_mbox_source(os.path.dirname(export)))
        self.assertTrue(is_mbox_source(export))
        self.assertFalse(is_mbox_source(ost))
        self.assertFalse(is_mbox_source(os.path.join(self.dir, "missing.ost")))


if __name__ == "__main__":
    unittest.main()
