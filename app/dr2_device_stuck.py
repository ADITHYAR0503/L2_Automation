import csv
import html
import io
import os
import re
from datetime import datetime
from email.message import EmailMessage
from typing import Any, Dict, List, Optional

from app.outlook_bridge import (
    create_outlook_reply_draft,
    derive_soti_search_url,
    open_eml_fallback,
)

ALLOWED_IMAGE_MIMES = {
    ".png": ("image", "png"),
    ".jpg": ("image", "jpeg"),
    ".jpeg": ("image", "jpeg"),
    ".webp": ("image", "webp"),
}

MAX_ROWS = 200
MAX_SCREENSHOT_BYTES = 5 * 1024 * 1024  # 5 MB

QUICK_COMMENTS: List[str] = [
    "DCU is in Print Screen, Hence BFTP Android Status 4",
    "DCU is not On Route, Hence BFTP Android Status 4",
    "Both DCU and BFTP Offline in SOTI",
    "DCU Offline in SOTI",
    "DCU is in Login Screen",
    "BFTP Device Recovered, Android Status 1 and in Tap Screen",
    "DCU is in Choose Mode Screen",
]


def _clean_text(value: str) -> str:
    cleaned = re.sub(r"[\r\n\x00-\x1f\x7f]", " ", value or "")
    return " ".join(cleaned.split()).strip()


def _sanitize_label(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]", "", (value or "").strip())
    return cleaned[:80]


def parse_stuck_rows(raw_text: str, soti_env: str = "Production") -> List[Dict[str, Any]]:
    lines = [line.strip() for line in (raw_text or "").splitlines() if line.strip()]
    if not lines:
        return []

    parsed: List[Dict[str, Any]] = []
    for idx, line in enumerate(lines):
        if len(parsed) >= MAX_ROWS:
            break

        lower_line = line.lower()
        if idx == 0 and "devicename" in lower_line and "status" in lower_line:
            continue

        if "\t" in line:
            parts = [p.strip() for p in line.split("\t")]
        elif "|" in line:
            parts = [p.strip() for p in line.split("|")]
        elif "," in line:
            reader = csv.reader(io.StringIO(line))
            parts = [p.strip() for p in next(reader, [])]
        else:
            m = re.match(r"^(\S+)\s+([0-9A-Za-z]+)\s+(.+)$", line)
            parts = [m.group(1), m.group(2), m.group(3)] if m else [line]

        if len(parts) < 2:
            continue

        device_name = _clean_text(parts[0]).upper()
        android_status = _clean_text(parts[1])
        online_duration = _clean_text(parts[2]) if len(parts) >= 3 else "0 Hour(s) 31 Minute(s)"
        comment_val = _clean_text(parts[3]) if len(parts) >= 4 else ""

        if not device_name:
            continue

        parsed.append(
            {
                "device_name": device_name,
                "android_device_status": android_status,
                "online_duration": online_duration,
                "soti_url": derive_soti_search_url(device_name, soti_env),
                "comments": comment_val,
            }
        )

    return parsed


def build_reply_fragment_html(
    rows: List[Dict[str, Any]],
    regards_name: str,
    screenshots: Optional[List[Dict[str, Any]]] = None,
) -> str:
    safe_regards = html.escape(_clean_text(regards_name) or "Adithya")

    row_html_list: List[str] = []
    for r in rows:
        dev_esc = html.escape(_clean_text(r.get("device_name", "")))
        st_esc = html.escape(_clean_text(r.get("android_device_status", "")))
        dur_esc = html.escape(_clean_text(r.get("online_duration", "")))
        cmt_esc = html.escape(_clean_text(r.get("comments", "")))

        row_html_list.append(
            "<tr>"
            f"<td style='border:1px solid #999999;padding:8px 10px;font-family:Calibri,Arial,sans-serif;font-size:11pt;'>{dev_esc}</td>"
            f"<td style='border:1px solid #999999;padding:8px 10px;font-family:Calibri,Arial,sans-serif;font-size:11pt;text-align:center;'>{st_esc}</td>"
            f"<td style='border:1px solid #999999;padding:8px 10px;font-family:Calibri,Arial,sans-serif;font-size:11pt;'>{dur_esc}</td>"
            f"<td style='border:1px solid #999999;padding:8px 10px;font-family:Calibri,Arial,sans-serif;font-size:11pt;'>{cmt_esc}</td>"
            "</tr>"
        )

    table_rows_html = "\n".join(row_html_list)

    shots_html_list: List[str] = []
    for shot in screenshots or []:
        label = html.escape(_sanitize_label(shot.get("device_label", "Device")))
        cid = html.escape(_clean_text(shot.get("cid", "")))
        if cid:
            shots_html_list.append(
                f"<div style='margin-top:14px;'>"
                f"<div style='font-family:Calibri,Arial,sans-serif;font-size:11pt;font-weight:bold;margin-bottom:4px;'>{label}</div>"
                f"<img src='cid:{cid}' alt='{label}' style='max-width:560px;border:1px solid #999999;' />"
                f"</div>"
            )

    shots_block = "\n".join(shots_html_list)

    return f"""<div style="font-family:Calibri,Arial,sans-serif;color:#111111;font-size:11pt;">
  <p>Hi All,</p>
  <p>Please find the list of devices with Status 4 or 5 and Online &gt; 30 minutes.</p>
  <table style="border-collapse:collapse;width:100%;max-width:860px;margin-top:8px;">
    <thead>
      <tr style="background-color:#d9d9d9;font-weight:bold;">
        <th style="border:1px solid #999999;padding:8px 10px;text-align:left;">DeviceName</th>
        <th style="border:1px solid #999999;padding:8px 10px;text-align:center;">AndroidDeviceStatus</th>
        <th style="border:1px solid #999999;padding:8px 10px;text-align:left;">OnlineDuration</th>
        <th style="border:1px solid #999999;padding:8px 10px;text-align:left;">Comments</th>
      </tr>
    </thead>
    <tbody>
      {table_rows_html}
    </tbody>
  </table>
  {shots_block}
  <p style="margin-top:16px;">Thanks &amp; Regards,<br/>{safe_regards}</p>
</div><br/>"""


def generate_device_stuck_mail(
    sender: str,
    to_addr: str,
    cc_addr: str,
    subject: str,
    regards_name: str,
    rows: List[Dict[str, Any]],
    entry_id: str = "",
    store_id: str = "",
    raw_screenshots: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    os.makedirs("output", exist_ok=True)
    shots_dir = os.path.abspath(os.path.join("output", "tmp_shots"))
    os.makedirs(shots_dir, exist_ok=True)

    clean_from = _clean_text(sender) or "adithya.rajeev@accenture.com"
    clean_to = _clean_text(to_addr) or "Presto.PDS.OpsL2Team@accenture.com"
    clean_cc = _clean_text(cc_addr)
    clean_subject = _clean_text(subject) or "[External] Action Required - Devices Stuck in Status 4 or 5"
    clean_regards = _clean_text(regards_name) or "Adithya"

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    prepared_shots: List[Dict[str, Any]] = []

    for idx, item in enumerate(raw_screenshots or []):
        raw_bytes: bytes = item.get("content", b"")
        filename = item.get("filename", f"proof_{idx+1}.png")
        ext = os.path.splitext(filename.lower())[1] or ".png"
        if not raw_bytes or len(raw_bytes) > MAX_SCREENSHOT_BYTES or ext not in ALLOWED_IMAGE_MIMES:
            continue

        dev_label = _sanitize_label(item.get("device_label") or f"Device_{idx+1}")
        safe_slug = re.sub(r"[^A-Za-z0-9_-]", "_", dev_label) or f"shot_{idx+1}"
        cid = f"dr2_shot_{timestamp}_{idx+1}_{safe_slug}@l2automation.local"
        maintype, subtype = ALLOWED_IMAGE_MIMES[ext]

        temp_img_path = os.path.join(shots_dir, f"{timestamp}_{idx+1}_{safe_slug}{ext}")
        with open(temp_img_path, "wb") as img_f:
            img_f.write(raw_bytes)

        prepared_shots.append(
            {
                "device_label": dev_label,
                "cid": cid,
                "maintype": maintype,
                "subtype": subtype,
                "filename": f"{safe_slug}{ext}",
                "path": temp_img_path,
                "content": raw_bytes,
            }
        )

    fragment_html = build_reply_fragment_html(rows, clean_regards, prepared_shots)

    # 1. Try replying directly to the exact fetched Outlook email thread
    thread_res = create_outlook_reply_draft(
        entry_id=entry_id,
        store_id=store_id,
        to_addr=clean_to,
        cc_addr=clean_cc,
        subject=clean_subject,
        html_fragment=fragment_html,
        attachments=prepared_shots,
    )

    if thread_res.get("replied_in_thread"):
        return {
            "mode": "outlook_thread_reply",
            "detail": thread_res["detail"],
        }

    # 2. Fallback: generate unsent .eml draft with exact subject and open in Outlook
    msg = EmailMessage()
    msg["From"] = clean_from
    msg["To"] = clean_to
    if clean_cc:
        msg["Cc"] = clean_cc
    msg["Subject"] = clean_subject
    msg["X-Unsent"] = "1"

    plain_lines = [
        "Hi All,",
        "Please find the list of devices with Status 4 or 5 and Online > 30 minutes.",
        "",
    ]
    for r in rows:
        plain_lines.append(
            f"- {r.get('device_name', '')} | Status: {r.get('android_device_status', '')} | "
            f"Duration: {r.get('online_duration', '')} | Comments: {r.get('comments', '')}"
        )
    plain_lines.extend(["", "Thanks & Regards,", clean_regards])
    msg.set_content("\n".join(plain_lines))

    full_html = f"<html><body>{fragment_html}</body></html>"
    msg.add_alternative(full_html, subtype="html")
    html_part = msg.get_body(preferencelist=("html",))
    if html_part is not None:
        for shot in prepared_shots:
            html_part.add_related(
                shot["content"],
                maintype=shot["maintype"],
                subtype=shot["subtype"],
                cid=f"<{shot['cid']}>",
                filename=shot["filename"],
            )

    eml_filename = f"dr2_device_stuck_{timestamp}.eml"
    eml_path = os.path.abspath(os.path.join("output", eml_filename))
    output_root = os.path.abspath("output")
    if os.path.commonpath([output_root, eml_path]) != output_root:
        raise ValueError("Invalid output path.")

    with open(eml_path, "wb") as f:
        f.write(msg.as_bytes())

    open_res = open_eml_fallback(eml_path)
    return {
        "mode": "eml_draft",
        "filename": eml_filename,
        "detail": open_res.get("detail", f"Saved {eml_filename} in output/."),
    }