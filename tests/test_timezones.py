import unittest

from unbound import timezones
from unbound.timezones import windows_zone


@unittest.skipIf(timezones._mapper() is None, "this Windows has no ICU library")
class IanaNames(unittest.TestCase):
    def test_iana_names_map_to_windows_names(self):
        for iana, windows in (("America/New_York", "Eastern Standard Time"),
                              ("Africa/Johannesburg", "South Africa Standard Time"),
                              ("Europe/London", "GMT Standard Time"),
                              ("Asia/Kolkata", "India Standard Time")):
            self.assertEqual(windows_zone(iana), windows)

    def test_names_that_are_not_iana_zones(self):
        for name in ("Not/AZone", "South Africa Standard Time", "", None, "a\x00b"):
            self.assertIsNone(windows_zone(name))


if __name__ == "__main__":
    unittest.main()
