"""Windows names for IANA time zones.

Outlook only knows Windows time zone names ("Eastern Standard Time").
Appointments made on phones or in Google Calendar carry IANA names
("America/New_York") instead. Windows 10 (1903) and later include the ICU
library, which holds the official mapping between the two, so it's asked
rather than shipping a table here that would go out of date.
"""

import ctypes
from functools import lru_cache


@lru_cache(maxsize=None)
def _mapper():
    try:
        fn = ctypes.WinDLL("icu.dll").ucal_getWindowsTimeZoneID
    except (OSError, AttributeError):
        return None  # older Windows, or not Windows at all
    fn.argtypes = [ctypes.c_wchar_p, ctypes.c_int32, ctypes.c_wchar_p, ctypes.c_int32,
                   ctypes.POINTER(ctypes.c_int)]
    fn.restype = ctypes.c_int32
    return fn


@lru_cache(maxsize=None)
def windows_zone(name):
    """'America/New_York' -> 'Eastern Standard Time'; None if there's no
    match or this Windows can't say."""
    fn = _mapper()
    if fn is None or not name or "\x00" in name:
        return None
    out = ctypes.create_unicode_buffer(128)
    status = ctypes.c_int(0)
    length = fn(name, -1, out, len(out), ctypes.byref(status))
    if status.value > 0 or length <= 0:  # ICU: above zero is an error
        return None
    return out.value or None
