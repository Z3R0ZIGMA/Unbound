"""Write messages into a new PST using Outlook Classic over COM.

Outlook saves a new MailItem as an unsent draft and won't let us clear that
flag afterwards. Instead each message is created as a PostItem (which is
never "unsent"), then given the original sender, times and recipients via
PropertyAccessor and re-classed as IPM.Note so it shows as normal mail.
"""

import os
import re
import tempfile
from datetime import timezone

import win32com.client

from . import mapi

OL_STORE_UNICODE = 3  # olStoreUnicode: Outlook 2003+ PST format
OL_POST_ITEM = 6  # olPostItem
OL_BY_VALUE = 1  # olByValue


def _as_local(utc_dt):
    """PropertyAccessor treats dates as local time and converts them to UTC
    itself, so hand it local time (pypff gives naive UTC)."""
    if utc_dt is None:
        return None
    return utc_dt.replace(tzinfo=timezone.utc).astimezone().replace(tzinfo=None)


def _safe_filename(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name[:150] or "attachment"


class PstWriter:
    def __init__(self, pst_path, display_name=None):
        self.pst_path = os.path.abspath(pst_path)
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self.ns.AddStoreEx(self.pst_path, OL_STORE_UNICODE)
        self.store = self._find_store()
        self.root = self.store.GetRootFolder()
        if display_name:
            self.root.Name = display_name
        self.property_errors = {}  # property name -> error text, for the log

    def _find_store(self):
        want = os.path.normcase(self.pst_path)
        for store in self.ns.Stores:
            try:
                path = store.FilePath
            except Exception:
                continue
            if path and os.path.normcase(os.path.abspath(path)) == want:
                return store
        raise RuntimeError(f"Outlook did not attach the PST: {self.pst_path}")

    def get_folder(self, parts):
        """Return the folder at root/parts[0]/parts[1]/..., creating as needed."""
        folder = self.root
        for name in parts:
            existing = None
            for sub in folder.Folders:
                if sub.Name == name:
                    existing = sub
                    break
            folder = existing or folder.Folders.Add(name)
        return folder

    def _set(self, accessor, label, prop_id, prop_type, value):
        if value is None or value == "":
            return
        try:
            accessor.SetProperty(mapi.proptag_url(prop_id, prop_type), value)
        except Exception as e:
            self.property_errors[label] = str(e)

    def _set_sender_and_times(self, item, m):
        pa = item.PropertyAccessor
        S, T = mapi.PT_UNICODE, mapi.PT_SYSTIME
        for label, pid in (("sender name", mapi.SENDER_NAME),
                           ("sent-representing name", mapi.SENT_REPRESENTING_NAME)):
            self._set(pa, label, pid, S, m.sender_name)
        for label, pid in (("sender address type", mapi.SENDER_ADDRTYPE),
                           ("sent-representing address type", mapi.SENT_REPRESENTING_ADDRTYPE)):
            self._set(pa, label, pid, S, "SMTP" if m.sender_email else None)
        for label, pid in (("sender email", mapi.SENDER_EMAIL_ADDRESS),
                           ("sent-representing email", mapi.SENT_REPRESENTING_EMAIL_ADDRESS),
                           ("sender SMTP", mapi.SENDER_SMTP_ADDRESS),
                           ("sent-representing SMTP", mapi.SENT_REPRESENTING_SMTP_ADDRESS)):
            self._set(pa, label, pid, S, m.sender_email)
        self._set(pa, "delivery time", mapi.MESSAGE_DELIVERY_TIME, T, _as_local(m.delivery_time))
        self._set(pa, "submit time", mapi.CLIENT_SUBMIT_TIME, T, _as_local(m.submit_time))

    def add_message(self, folder, m):
        # 1. Create as a PostItem so it is saved as delivered mail, not a draft
        post = folder.Items.Add(OL_POST_ITEM)
        post.Subject = m.subject
        if m.html_body:
            post.HTMLBody = m.html_body
        elif m.plain_body:
            post.Body = m.plain_body

        with tempfile.TemporaryDirectory(prefix="unbound_") as tmp:
            for i, att in enumerate(m.attachments):
                # unique subfolder per attachment so duplicate names don't clash
                d = os.path.join(tmp, str(i))
                os.mkdir(d)
                path = os.path.join(d, _safe_filename(att.filename))
                with open(path, "wb") as fh:
                    fh.write(att.data)
                post.Attachments.Add(path, OL_BY_VALUE, 1, att.filename)
            post.MessageClass = m.message_class
            post.Save()
        entry_id = post.EntryID
        del post

        # 2. Reopen it: now it's a MailItem, so it takes real recipients.
        #    Times are set here because the first save overwrites them.
        item = self.ns.GetItemFromID(entry_id, self.store.StoreID)
        for r in m.recipients:
            addr = f'"{r.name}" <{r.email}>' if r.name and r.email else (r.email or r.name)
            if not addr:
                continue
            try:
                rec = item.Recipients.Add(addr)
                rec.Type = r.kind
                # Without this Outlook stores the raw text as the display
                # name and leaves the email address empty
                rec.Resolve()
            except Exception as e:
                self.property_errors["recipients"] = str(e)
        self._set_sender_and_times(item, m)
        item.UnRead = not m.is_read
        item.Save()
        return item

    def close(self, quit_outlook=False):
        """Detach the PST from the Outlook profile (the file stays on disk)."""
        try:
            self.ns.RemoveStore(self.root)
        finally:
            if quit_outlook:
                self.app.Quit()
            # Drop every COM reference; Outlook keeps the PST open while
            # any are alive.
            self.root = self.store = self.ns = self.app = None
            import gc
            gc.collect()
