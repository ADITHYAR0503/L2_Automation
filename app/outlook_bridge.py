import os
import re
import subprocess
import sys
from typing import Any, Dict, List, Optional
from urllib.parse import quote


def derive_soti_search_url(device_name: str, env: str = "Production") -> str:
    dev = (device_name or "").strip().upper()
    if not dev:
        return "https://soti-mobicontrol.prestoprod.net/MobiControl/WebConsole/home/devices"

    search_terms = [dev]
    m_bftp = re.match(r"^([A-Z0-9]+)_BFTP_([A-Z0-9]+)_(\d+)$", dev)
    m_dcu = re.match(r"^([A-Z0-9]+)_DCU_([A-Z0-9]+)_(\d+)$", dev)

    if m_bftp:
        prefix, mid, _ = m_bftp.groups()
        search_terms.append(f"{prefix}_DCU_{mid}_1")
    elif m_dcu:
        prefix, mid, _ = m_dcu.groups()
        search_terms.append(f"{prefix}_BFTP_{mid}_2")

    joined = " OR ".join(f'Text AlmostEqual "{t}"' for t in search_terms)
    encoded_q = quote(f" {joined}", safe="")

    env_key = (env or "Production").strip().lower()
    base = "https://soti-mobicontrol.prestoprod.net"
    if env_key == "ctt":
        base = "https://soti-mobicontrol.prestoctt.net"
    elif env_key == "preprod":
        base = "https://soti-mobicontrol.prestopreprod.net"

    return f"{base}/MobiControl/WebConsole/home/devices?search={encoded_q}&subGroups=true&count=250"


def _strip_html_tags(raw_html: str) -> str:
    text = re.sub(r"<style.*?>.*?</style>", " ", raw_html or "", flags=re.S | re.I)
    text = re.sub(r"<script.*?>.*?</script>", " ", text, flags=re.S | re.I)
    text = re.sub(r"</(tr|p|div|br|li|h[1-6])>", "\n", text, flags=re.I)
    text = re.sub(r"</(td|th)>", " | ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    return (
        text.replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
    )


def _extract_rows_from_text_or_html(body_text: str, html_body: str = "") -> List[Dict[str, str]]:
    combined = (body_text or "") + "\n" + _strip_html_tags(html_body or "")
    rows: List[Dict[str, str]] = []
    seen = set()

    for line in combined.splitlines():
        clean = " ".join(line.split()).strip()
        if not clean:
            continue

        dev_match = re.search(
            r"\b([A-Z]{2,8}_(?:BFTP|DCU|MFTP|SFTP|MPOS|SPOS|TPT)_[A-Za-z0-9]+_\d+)\b",
            clean,
            re.I,
        )
        if not dev_match:
            continue

        dev_name = dev_match.group(1).upper()
        rest = clean[dev_match.end():].strip(" |-\t")

        status_match = re.search(r"\b([0-9])\b", rest)
        status_val = status_match.group(1) if status_match else "4"

        dur_match = re.search(
            r"(\d+\s*Hour\(s\)\s*\d+\s*Minute\(s\)|\d+\s*(?:hrs?|hours?|mins?|minutes?))",
            rest,
            re.I,
        )
        duration_val = dur_match.group(1).strip() if dur_match else "0 Hour(s) 31 Minute(s)"

        key = (dev_name, status_val, duration_val)
        if key not in seen:
            seen.add(key)
            rows.append(
                {
                    "device_name": dev_name,
                    "android_device_status": status_val,
                    "online_duration": duration_val,
                }
            )

    return rows


def _collect_all_folders(root_folder, collected: list, depth: int = 0) -> None:
    if depth > 4:
        return
    collected.append(root_folder)
    try:
        for sub in root_folder.Folders:
            _collect_all_folders(sub, collected, depth + 1)
    except Exception:
        pass


def fetch_mail_by_subject_from_outlook(subject_query: str) -> Dict[str, Any]:
    """
    Search Outlook Inbox and all subfolders (e.g., DR2 REPORTS > Devices Stuck)
    for the most recent email matching the pasted subject/title.
    Returns EntryID and StoreID so Generate Mail can reply directly to that exact email.
    """
    if sys.platform != "win32":
        return {"connected": False, "found": False, "message": "Windows Outlook required.", "rows": []}

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        return {"connected": False, "found": False, "message": "pywin32 not installed.", "rows": []}

    needle = (subject_query or "").strip().lower()
    if not needle:
        return {"connected": True, "found": False, "message": "Paste a mail subject/title first.", "rows": []}

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # Inbox

        folders: list = []
        _collect_all_folders(inbox, folders)

        # Prioritize folders named 'Devices Stuck' or 'DR2'
        folders.sort(
            key=lambda f: 0
            if any(k in str(getattr(f, "Name", "")).lower() for k in ("stuck", "dr2"))
            else 1
        )

        best_mail = None
        best_time = None

        for folder in folders:
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                checked = 0
                for item in items:
                    checked += 1
                    if checked > 120:
                        break
                    subj = str(getattr(item, "Subject", "") or "")
                    if needle in subj.lower():
                        rec_time = getattr(item, "ReceivedTime", None)
                        if best_time is None or (rec_time is not None and rec_time > best_time):
                            best_time = rec_time
                            best_mail = item
                        break
            except Exception:
                continue

        if best_mail is None:
            return {
                "connected": True,
                "found": False,
                "message": f"No mail found matching: {subject_query}",
                "rows": [],
            }

        subj = str(getattr(best_mail, "Subject", "") or "")
        entry_id = str(getattr(best_mail, "EntryID", "") or "")
        store_id = str(getattr(getattr(best_mail, "Parent", None), "StoreID", "") or "")
        body = str(getattr(best_mail, "Body", "") or "")
        html_body = str(getattr(best_mail, "HTMLBody", "") or "")
        rows = _extract_rows_from_text_or_html(body, html_body)

        return {
            "connected": True,
            "found": True,
            "subject": subj,
            "entry_id": entry_id,
            "store_id": store_id,
            "message": f"Fetched: {subj}",
            "rows": rows,
        }
    except Exception as exc:
        return {"connected": False, "found": False, "message": f"Outlook error: {exc}", "rows": []}
    finally:
        pythoncom.CoUninitialize()


def create_outlook_reply_draft(
    entry_id: str,
    store_id: str,
    to_addr: str,
    cc_addr: str,
    subject: str,
    html_fragment: str,
    attachments: Optional[List[Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """
    Reply directly to the exact fetched Outlook email item (preserving thread + original mail below)
    and display the draft window in Outlook for human review before sending.
    """
    if sys.platform != "win32" or not entry_id:
        return {"replied_in_thread": False, "detail": "No Outlook EntryID; falling back to .eml draft."}

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        return {"replied_in_thread": False, "detail": "pywin32 unavailable."}

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        original_mail = (
            namespace.GetItemFromID(entry_id, store_id)
            if store_id
            else namespace.GetItemFromID(entry_id)
        )

        reply_item = original_mail.ReplyAll()
        if to_addr.strip():
            reply_item.To = to_addr.strip()
        if cc_addr.strip():
            reply_item.CC = cc_addr.strip()
        if subject.strip():
            reply_item.Subject = subject.strip()

        # Attach inline screenshots via Content-ID so they render cleanly inside Outlook reply
        for att in attachments or []:
            abs_file = os.path.abspath(att["path"])
            if os.path.exists(abs_file):
                mail_att = reply_item.Attachments.Add(abs_file)
                # PR_ATTACH_CONTENT_ID MAPI property tag
                mail_att.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    att["cid"],
                )

        existing_html = str(getattr(reply_item, "HTMLBody", "") or "")
        body_open_match = re.search(r"<body[^>]*>", existing_html, flags=re.I)
        if body_open_match:
            insert_pos = body_open_match.end()
            reply_item.HTMLBody = (
                existing_html[:insert_pos] + html_fragment + existing_html[insert_pos:]
            )
        else:
            reply_item.HTMLBody = html_fragment + existing_html

        reply_item.Save()
        reply_item.Display()
        return {
            "replied_in_thread": True,
            "detail": "Reply draft opened directly on the original Outlook email thread.",
        }
    except Exception as exc:
        return {"replied_in_thread": False, "detail": f"Thread reply fallback ({exc})"}
    finally:
        pythoncom.CoUninitialize()


def open_eml_fallback(eml_path: str) -> Dict[str, Any]:
    abs_path = os.path.abspath(eml_path)
    output_root = os.path.abspath("output")
    if os.path.commonpath([output_root, abs_path]) != output_root or not os.path.exists(abs_path):
        return {"opened": False, "detail": "Invalid draft file path."}

    if sys.platform == "win32":
        try:
            os.startfile(abs_path)  # type: ignore[attr-defined]
            return {"opened": True, "detail": "Opened unsent .eml draft in Outlook."}
        except Exception:
            try:
                subprocess.Popen(["outlook.exe", "/eml", abs_path], shell=False)
                return {"opened": True, "detail": "Launched Outlook with .eml draft."}
            except Exception as exc:
                return {"opened": False, "detail": f"Draft saved in output/: {exc}"}

    return {"opened": False, "detail": "Draft saved in output/."}