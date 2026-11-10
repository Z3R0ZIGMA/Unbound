"""Read-only check: open an OST/PST with pypff and print its folder tree.

Usage:
    .venv\\Scripts\\python scripts\\list_folders.py PATH_TO_FILE.ost [--sample N]

For each folder it prints the item count. With --sample N it also tries to
read up to N messages per folder (subject, sender, body, attachment count) and
reports how many read cleanly. It never prints message content.
"""

import argparse
import os
import sys
import time

import pypff


def try_read_message(msg):
    """Touch the fields the converter will need; raise if any fail."""
    _ = msg.subject
    _ = msg.sender_name
    _ = msg.transport_headers
    _ = msg.plain_text_body
    # pypff raises OSError when a body format is absent; that's normal
    for attr in ("html_body", "rtf_body"):
        try:
            getattr(msg, attr)
        except OSError:
            pass
    for i in range(msg.number_of_attachments):
        msg.get_attachment(i).get_size()


def walk(folder, depth, sample, totals):
    name = folder.name or "(root)"
    try:
        count = folder.number_of_sub_messages
    except OSError as e:
        print(f"{'  ' * depth}{name}  [ERROR reading count: {e}]")
        totals["folder_errors"] += 1
        count = 0
    else:
        line = f"{'  ' * depth}{name}  ({count})"
        if sample and count:
            ok = bad = skipped = 0
            for i in range(min(sample, count)):
                try:
                    item = folder.get_sub_message(i)
                    if not isinstance(item, pypff.message):
                        skipped += 1  # hidden system item, not a message
                        continue
                    try_read_message(item)
                    ok += 1
                except Exception:
                    bad += 1
            totals["sampled_ok"] += ok
            totals["sampled_bad"] += bad
            totals["sampled_skipped"] += skipped
            line += f"  sample: {ok} ok, {bad} failed"
            if skipped:
                line += f", {skipped} system items skipped"
        print(line)
    totals["folders"] += 1
    totals["items"] += count

    try:
        n_sub = folder.number_of_sub_folders
    except OSError as e:
        print(f"{'  ' * (depth + 1)}[ERROR reading subfolders: {e}]")
        totals["folder_errors"] += 1
        return
    for i in range(n_sub):
        try:
            sub = folder.get_sub_folder(i)
        except OSError as e:
            print(f"{'  ' * (depth + 1)}[ERROR opening subfolder {i}: {e}]")
            totals["folder_errors"] += 1
            continue
        walk(sub, depth + 1, sample, totals)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    parser.add_argument("--sample", type=int, default=0)
    args = parser.parse_args()

    start = time.time()
    f = pypff.file()
    try:
        # libpff uses \\?\ long paths on Windows, which reject forward slashes
        f.open(os.path.abspath(args.path))  # pypff opens read-only
    except OSError as e:
        print(f"Could not open file: {e}")
        return 1

    totals = dict(folders=0, items=0, folder_errors=0,
                  sampled_ok=0, sampled_bad=0, sampled_skipped=0)
    try:
        walk(f.get_root_folder(), 0, args.sample, totals)
    finally:
        f.close()

    print()
    print(f"Folders: {totals['folders']}  Items: {totals['items']}  "
          f"Folder errors: {totals['folder_errors']}")
    if args.sample:
        print(f"Sampled messages: {totals['sampled_ok']} ok, {totals['sampled_bad']} failed, "
              f"{totals['sampled_skipped']} system items skipped")
    print(f"Time: {time.time() - start:.1f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
