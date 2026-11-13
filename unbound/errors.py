"""Turning exceptions into short text for the window and the log."""


def first_line(e, limit=300):
    """The first line of an exception's message, at most `limit` characters.

    Some exceptions have no message at all (MemoryError(), a bare
    `raise ValueError`), so `str(e).splitlines()[0]` would itself raise
    IndexError inside the very handler meant to report the problem. Those
    get the exception's name instead.
    """
    return (str(e).splitlines() or [type(e).__name__])[0][:limit]
