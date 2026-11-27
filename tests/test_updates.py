"""Check for updates, without the internet: GitHub's answers are stand-ins."""

import hashlib
import io
import json
import os
import tempfile
import unittest
import urllib.error
from unittest import mock

from unbound import updates

EXE = b"MZ new Unbound" * 1000
SUMS = f"{hashlib.sha256(EXE).hexdigest()}  Unbound.exe\r\n"
BASE = "https://github.com/Z3R0ZIGMA/Unbound/releases/download/v1.4.0/"


def release_json(tag="v1.4.0", assets=("Unbound.exe", "SHA256SUMS.txt", "Third-party-licenses.txt"),
                 base=BASE, **extra):
    return json.dumps({
        "tag_name": tag, "draft": False, "prerelease": False,
        "html_url": f"https://github.com/Z3R0ZIGMA/Unbound/releases/tag/{tag}",
        "assets": [{"name": n, "browser_download_url": base + n} for n in assets], **extra})


class Response(io.BytesIO):
    def __init__(self, body, length=True):
        super().__init__(body)
        self.headers = {"Content-Length": str(len(body))} if length else {}

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()


def serve(files):
    """A stand-in for urlopen that answers from {url: bytes}."""
    def urlopen(request, timeout):
        url = request.full_url
        if url not in files:
            raise urllib.error.HTTPError(url, 404, "Not Found", {}, None)
        return Response(files[url])
    return urlopen


class ReadingTheRelease(unittest.TestCase):
    def test_finds_the_exe_and_its_checksums(self):
        r = updates.parse_latest(release_json())
        self.assertEqual((r.version, r.name, r.tag), ((1, 4, 0), "1.4.0", "v1.4.0"))
        self.assertEqual(r.exe_url, BASE + "Unbound.exe")
        self.assertEqual(r.sums_url, BASE + "SHA256SUMS.txt")
        self.assertTrue(r.page_url.endswith("/releases/tag/v1.4.0"))

    def test_nothing_to_install_without_the_exe_or_its_checksums(self):
        self.assertIsNone(updates.parse_latest(release_json(assets=("Unbound.exe",))))
        self.assertIsNone(updates.parse_latest(release_json(assets=("SHA256SUMS.txt",))))

    def test_drafts_pre_releases_and_odd_tags_are_ignored(self):
        self.assertIsNone(updates.parse_latest(release_json(draft=True)))
        self.assertIsNone(updates.parse_latest(release_json(prerelease=True)))
        self.assertIsNone(updates.parse_latest(release_json(tag="nightly")))

    def test_only_this_projects_own_downloads(self):
        """A release whose files point anywhere else is never used."""
        other = "https://example.com/Z3R0ZIGMA/Unbound/releases/download/v1.4.0/"
        self.assertIsNone(updates.parse_latest(release_json(base=other)))

    def test_versions(self):
        self.assertEqual(updates.parse_version("v1.3.5"), (1, 3, 5))
        self.assertEqual(updates.parse_version("1.4"), (1, 4, 0))
        self.assertIsNone(updates.parse_version("1.4.0-beta"))
        self.assertTrue(updates.is_newer((1, 3, 10), (1, 3, 9)))  # not compared as text
        self.assertFalse(updates.is_newer((1, 3, 5), (1, 3, 5)))
        self.assertFalse(updates.is_newer((1, 3, 4), (1, 3, 5)))

    def test_checksum_lines(self):
        h = "ab" * 32
        self.assertEqual(updates.expected_hash(f"{h}  Unbound.exe\n"), h)
        self.assertEqual(updates.expected_hash(f"{h.upper()} *unbound.exe\r\n"), h)
        self.assertIsNone(updates.expected_hash(f"{h}  Other.exe\n"))
        self.assertIsNone(updates.expected_hash("not-a-hash  Unbound.exe\n"))


class AskingGitHub(unittest.TestCase):
    def test_latest_release(self):
        api = {updates.LATEST_RELEASE_API: release_json().encode()}
        with mock.patch("urllib.request.urlopen", serve(api)):
            self.assertEqual(updates.get_latest().name, "1.4.0")

    def test_says_who_is_asking_and_nothing_else(self):
        seen = []

        def urlopen(request, timeout):
            seen.append(request)
            return Response(release_json().encode())
        with mock.patch("urllib.request.urlopen", urlopen):
            updates.get_latest()
        request = seen[0]
        self.assertEqual(request.get_method(), "GET")
        self.assertIsNone(request.data)
        self.assertEqual(request.get_header("User-agent"), f"Unbound/{updates.__version__}")

    def test_errors_are_explained(self):
        cases = [
            (urllib.error.URLError("getaddrinfo failed"), "internet connection"),
            (urllib.error.HTTPError("u", 403, "rate limit exceeded", {}, None), "Try again in an hour"),
            (urllib.error.HTTPError("u", 500, "Server Error", {}, None), "500"),
        ]
        for error, words in cases:
            with mock.patch("urllib.request.urlopen", side_effect=error), \
                    self.assertRaises(updates.UpdateError) as caught:
                updates.get_latest()
            self.assertIn(words, str(caught.exception))
        with mock.patch("urllib.request.urlopen", serve({updates.LATEST_RELEASE_API: b"<html>"})), \
                self.assertRaises(updates.UpdateError):
            updates.get_latest()


class Downloading(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.target = os.path.join(tmp.name, "Unbound.exe.new")
        self.release = updates.parse_latest(release_json())

    def download(self, exe=EXE, sums=SUMS, **kwargs):
        files = {BASE + "Unbound.exe": exe, BASE + "SHA256SUMS.txt": sums.encode()}
        with mock.patch("urllib.request.urlopen", serve(files)):
            updates.download(self.release, self.target, **kwargs)

    def test_checked_download(self):
        seen = []
        self.download(progress=seen.append)
        with open(self.target, "rb") as fh:
            self.assertEqual(fh.read(), EXE)
        self.assertEqual(seen[-1], 100)

    def test_a_file_that_doesnt_match_is_deleted(self):
        with self.assertRaises(updates.UpdateError) as caught:
            self.download(exe=EXE + b"tampered")
        self.assertIn("checksum", str(caught.exception))
        self.assertFalse(os.path.exists(self.target))

    def test_not_listed_in_the_checksums(self):
        with self.assertRaises(updates.UpdateError):
            self.download(sums="ab" * 32 + "  Other.exe\n")
        self.assertFalse(os.path.exists(self.target))

    def test_cancelled_part_way(self):
        with self.assertRaises(updates.UpdateError):
            self.download(cancelled=lambda: True)
        self.assertFalse(os.path.exists(self.target))


class SwappingIn(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.exe = os.path.join(tmp.name, "Unbound.exe")
        self.write(self.exe, b"old")
        self.staged = updates.staged_path(self.exe)
        self.write(self.staged, b"new")

    @staticmethod
    def write(path, data):
        with open(path, "wb") as fh:
            fh.write(data)

    @staticmethod
    def read(path):
        with open(path, "rb") as fh:
            return fh.read()

    def test_new_in_place_old_kept_aside(self):
        updates.install(self.staged, self.exe)
        self.assertEqual(self.read(self.exe), b"new")
        self.assertEqual(self.read(self.exe + ".old"), b"old")
        self.assertFalse(os.path.exists(self.staged))

    def test_put_back_when_the_new_one_cant_be_moved(self):
        real_replace = os.replace

        def replace(src, dst):
            if src == self.staged:
                raise PermissionError("in use")
            real_replace(src, dst)
        with mock.patch.object(updates.os, "replace", replace), self.assertRaises(OSError):
            updates.install(self.staged, self.exe)
        self.assertEqual(self.read(self.exe), b"old")

    def test_leftovers_removed_at_the_next_start(self):
        self.write(self.exe + ".old", b"older")
        updates.clean_up(self.exe)
        self.assertFalse(os.path.exists(self.exe + ".old"))
        self.assertFalse(os.path.exists(self.staged))
        self.assertTrue(os.path.exists(self.exe))
        updates.clean_up(None)  # run from source: nothing to do

    def test_running_from_source_has_no_exe_to_replace(self):
        with mock.patch.object(updates.sys, "frozen", False, create=True):
            self.assertIsNone(updates.running_exe())

    def test_restart_starts_a_separate_copy(self):
        with mock.patch("subprocess.Popen") as popen:
            updates.start_new_instance(self.exe)
        args, kwargs = popen.call_args
        self.assertEqual(args[0], [self.exe])
        self.assertEqual(kwargs["env"]["PYINSTALLER_RESET_ENVIRONMENT"], "1")


if __name__ == "__main__":
    unittest.main()
