import html
import io
import re
import requests
import smtplib
import time
import shutil
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
import os
import csv
import json
import subprocess
import uuid
import random
from urllib.parse import quote

from fpdf import FPDF
from bidi.algorithm import get_display

import dropbox
try:
    from dropbox.common import PathRoot
except ImportError:
    PathRoot = None
import streamlit as st
from streamlit_calendar import calendar
import holidays
from docxtpl import DocxTemplate
import pandas as pd
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

# --- התחברות לגוגל שיטס ---
# init_connection אינו משתמש ב-@st.cache_data - חיבור נוצר מחדש בעת הצורך
def init_connection():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_data = st.secrets["gcp_service_account"]
    creds_dict = json.loads(creds_data) if isinstance(creds_data, str) else creds_data
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds)


def _get_spreadsheet():
    """מחזיר את אובייקט הגיליון. אם אין חיבור, מנסה להתחבר מחדש לפני כתיבה."""
    global spreadsheet, _sheets_init_error
    if spreadsheet is not None:
        return spreadsheet
    try:
        client = init_connection()
        spreadsheet = client.open_by_key(SHEET_ID)
        _sheets_init_error = None
        return spreadsheet
    except Exception as e:
        _sheets_init_error = e
        return None


SHEET_ID = '1ZvAtkWaXpf9zZRgXY2HUcRB6QWpUMe6KWNjPu-eyzdo'
_sheets_init_error = None
try:
    client = init_connection()
    spreadsheet = client.open_by_key(SHEET_ID)
except Exception as e:
    spreadsheet = None
    _sheets_init_error = e
    st.error(f"שגיאה בהתחברות לגוגל שיטס: {e}")

# נתיב LibreOffice להמרת DOCX ל-PDF (חלופה ל-Word)
LIBREOFFICE_PATH = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")


def convert_to_pdf(docx_path: str | Path, pdf_path: str | Path) -> bool:
    """
    ממיר קובץ DOCX ל-PDF באמצעות LibreOffice.
    מחזיר True אם ההמרה הצליחה, False אחרת.
    """
    docx_path = Path(docx_path)
    pdf_path = Path(pdf_path)
    if not docx_path.exists():
        raise FileNotFoundError(f"קובץ המקור לא נמצא: {docx_path}")

    if not LIBREOFFICE_PATH.exists():
        raise RuntimeError("לא נמצא LibreOffice. אנא וודא שהתקנת אותו.")

    outdir = pdf_path.parent
    outdir.mkdir(parents=True, exist_ok=True)

    try:
        result = subprocess.run(
            [
                str(LIBREOFFICE_PATH),
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(outdir),
                str(docx_path),
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice נכשל: {result.stderr or result.stdout or 'שגיאה לא ידועה'}"
            )
        # LibreOffice יוצר קובץ עם אותו שם בסיס + .pdf
        expected_pdf = outdir / (docx_path.stem + ".pdf")
        if expected_pdf.exists():
            if expected_pdf != pdf_path:
                expected_pdf.rename(pdf_path)
            return True
        return False
    except subprocess.TimeoutExpired:
        raise RuntimeError("המרה ל-PDF נכשלה: פג הזמן.")
import plotly.express as px

# --- נתיבי בסיס דינמיים ---
# כל שמירות הקבצים ותיקיות העבודה יהיו יחסית למיקום הקובץ app.py,
# כך שהאפליקציה תעבוד באופן זהה אצל כל עובד בלי קשר לאות הכונן.
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
# All data in Google Sheets - no local CSV files
PROJECTS_ROOT = Path(os.path.join(CURRENT_DIR, "Projects"))
QUOTES_ROOT = BASE_DIR / "Quotes"
TEMPLATE_PATH = BASE_DIR / "quote_template.docx"

# Legacy paths (for backward compatibility when searching)
QUOTES_PENDING = QUOTES_ROOT / "Pending"
QUOTES_APPROVED = QUOTES_ROOT / "Approved"
QUOTES_REJECTED = QUOTES_ROOT / "Rejected"

# תיקייה מקומית להצעות - שומרת את הקובץ בין שמירה לשליחת מייל (לאחר rerun של Streamlit)
TEMP_PROPOSALS_DIR = BASE_DIR / "temp_proposals"
# עותק קשיח של ה-PDF האחרון בתיקיית הפרויקט – לצירוף למייל (עמיד ב-rerun של Streamlit)
CURRENT_QUOTE_TEMP_PDF = BASE_DIR / "current_quote_temp.pdf"


def _persist_current_quote_temp_pdf(source_path: str | Path) -> None:
    """מעתיק את קובץ ה-PDF לשם קבוע בשורש הפרויקט לשליחה אמינה."""
    try:
        p = Path(source_path)
        if p.is_file() and p.suffix.lower() == ".pdf":
            shutil.copyfile(str(p), str(CURRENT_QUOTE_TEMP_PDF))
    except Exception:
        pass


MESSAGES_COLUMNS = ["Timestamp", "Sender", "Recipient", "Type", "Message"]


def _normalize_records_to_columns(records: list[dict], expected_columns: list[str]) -> list[dict]:
    """מנרמל רשומות מגוגל שיטס כך ששמות העמודות יתאימו בדיוק (גם אם יש הבדל באותיות גדולות/קטנות)."""
    if not records:
        return []
    key_lower_map = {col.lower(): col for col in expected_columns}
    result = []
    for r in records:
        normalized = {}
        for k, v in r.items():
            if k and isinstance(k, str):
                canonical = key_lower_map.get(k.strip().lower())
                if canonical:
                    normalized[canonical] = (str(v or "").strip())
        for col in expected_columns:
            if col not in normalized:
                normalized[col] = ""
        result.append(normalized)
    return result


def _read_worksheet_safe(worksheet, expected_columns: list[str]) -> pd.DataFrame:
    """קריאה בטוחה מגיליון גוגל שיטס - תומך בכותרות בעברית, מונע קריסות של gspread."""
    raw_data = worksheet.get_all_values()
    if len(raw_data) > 1:
        raw_headers = [str(h).strip() for h in raw_data[0]]
        df = pd.DataFrame(raw_data[1:], columns=raw_headers)
        # מיפוי כותרות גולמיות לעמודות הצפויות (case-insensitive) - למקרה של רווחים/אותיות
        rename_map = {}
        for col in df.columns:
            col_clean = (col or "").strip()
            for exp in expected_columns:
                if exp.strip().lower() == col_clean.lower():
                    rename_map[col] = exp
                    break
        df = df.rename(columns=rename_map)
        df = df.reindex(columns=expected_columns, fill_value='')
    else:
        df = pd.DataFrame(columns=expected_columns)
    # ניקוי נתונים אגרסיבי - כולל הסרת רווחים משמות עמודות (כמו ב-quotes)
    try:
        df.columns = [str(c).strip() for c in df.columns]
    except Exception:
        pass
    df = df.dropna(how='all')
    df = df.fillna('')
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
    return df


def _ensure_messages_csv() -> None:
    """וידוא שקיים גיליון messages בגוגל שיטס. לא יוצר גיליון חדש - תציג שגיאה אם חסר."""
    if spreadsheet is None:
        return
    spreadsheet.worksheet('messages')  # יקרוס ויציג שגיאה אם הגיליון חסר


@st.cache_data(ttl=600)
def _read_messages_df() -> pd.DataFrame:
    """קריאת הודעות תקשורת מהירה מגוגל שיטס."""
    if spreadsheet is None:
        return pd.DataFrame(columns=MESSAGES_COLUMNS)
    _ensure_messages_csv()
    try:
        worksheet = spreadsheet.worksheet('messages')
        return _read_worksheet_safe(worksheet, MESSAGES_COLUMNS)
    except Exception as e:
        st.error(f"שגיאה בשליפת נתונים (messages): {e}")
        return pd.DataFrame(columns=MESSAGES_COLUMNS)


def _write_messages_df(df: pd.DataFrame) -> None:
    """שמירת הודעות תקשורת מהירה לגוגל שיטס."""
    if spreadsheet is None:
        st.error("אין חיבור לגוגל שיטס. לא ניתן לשמור.")
        return
    try:
        worksheet = spreadsheet.worksheet('messages')
        worksheet.clear()
        df_safe = df.reindex(columns=MESSAGES_COLUMNS, fill_value="").fillna("").astype(str)
        data = [df_safe.columns.values.tolist()] + df_safe.values.tolist()
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        time.sleep(1.5)
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת messages: {e}")


def clean_name_for_match(name) -> str:
    """פונקציית עזר לניקוי שמות - מתעלמת מגרשיים, מקפים, רווחים ופיסוק."""
    if pd.isna(name):
        return ""
    return re.sub(r"[^a-zA-Zא-ת0-9]", "", str(name))


def _render_quick_comm_sidebar_form() -> None:
    """מציג טופס תקשורת מהירה בתפריט הצידי."""
    team_list = [n for n in TEAM_DISPLAY_NAMES if n and str(n).strip()]
    st.sidebar.markdown("---")
    st.sidebar.subheader("💬 תקשורת מהירה")
    with st.sidebar.form("quick_comm_form", clear_on_submit=True):
        msg_type = st.selectbox(
            "סוג הודעה:",
            ["⚡ משימה מהירה", "🤝 זימון ישיבה", "📢 עדכון כללי"],
        )
        msg_recipients = st.multiselect("נמען:", ["כולם"] + team_list, default=["כולם"])
        msg_content = st.text_input("תוכן ההודעה:")
        submit_msg = st.form_submit_button("שלח הודעה")
    if submit_msg and msg_content:
        if not msg_recipients:
            st.sidebar.warning("נא לבחור לפחות נמען אחד.")
        else:
            _ensure_messages_csv()
            current_user = (st.session_state.get("current_user") or "צוות").strip()
            recipient_str = "כולם" if "כולם" in msg_recipients else ", ".join(msg_recipients)
            new_row = {
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Sender": current_user,
                "Recipient": recipient_str,
                "Type": msg_type,
                "Message": msg_content.strip(),
            }
            df = _read_messages_df()
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            _write_messages_df(df)
            st.rerun()


def _get_dropbox_access_token() -> str | None:
    """מחזיר Access Token לדרופבוקס מ-Secrets: DROPBOX_ACCESS_TOKEN או dropbox.access_token (או מפתחות חלופיים תחת dropbox)."""
    try:
        flat = st.secrets.get("DROPBOX_ACCESS_TOKEN")
        if flat is not None and str(flat).strip():
            return str(flat).strip()
    except Exception:
        pass
    try:
        nested = st.secrets.get("dropbox")
        if isinstance(nested, dict):
            for key in ("access_token", "DROPBOX_ACCESS_TOKEN", "token"):
                v = nested.get(key)
                if v is not None and str(v).strip():
                    return str(v).strip()
    except Exception:
        pass
    return None


def _render_dropbox_access_token_hint_sidebar() -> None:
    """מציג אזהרה בתפריט הצד רק כאשר אין DROPBOX_ACCESS_TOKEN (או מקביל תחת dropbox)."""
    if _get_dropbox_access_token():
        return
    st.sidebar.markdown("---")
    st.sidebar.warning(
        "נדרש חיבור לדרופבוקס – הגדר ב-Secrets את DROPBOX_ACCESS_TOKEN "
        "או מקטע [dropbox] עם access_token."
    )


def _render_quick_comm_notifications() -> None:
    """מציג התראות תקשורת מהירה בראש המסך - הודעות מ-24 השעות האחרונות."""
    _ensure_messages_csv()
    df = _read_messages_df()
    if df.empty or "Timestamp" not in df.columns:
        return
    current_user = (st.session_state.get("current_user") or "צוות").strip()
    safe_current_user = clean_name_for_match(current_user)
    cutoff = datetime.now() - timedelta(days=1)
    for _, row in df.iterrows():
        try:
            ts = pd.to_datetime(row.get("Timestamp"), errors="coerce")
            if pd.isna(ts) or ts < cutoff:
                continue
        except Exception:
            continue
        recipient = (row.get("Recipient") or "").strip()
        if recipient != "כולם":
            # נמענים מרובים נשמרים כמחרוזת מופרדת בפסיקים - clean_name_for_match מסיר פסיקים
            if not safe_current_user or safe_current_user not in clean_name_for_match(recipient):
                continue
        msg_type = (row.get("Type") or "").strip()
        sender = (row.get("Sender") or "").strip()
        message = (row.get("Message") or "").strip()
        time_str = ts.strftime("%H:%M") if hasattr(ts, "strftime") else str(ts)[:5]
        display_text = f"**{msg_type} מ-{sender}:** {message} _(נשלח ב-{time_str})_"
        if "ישיבה" in msg_type:
            st.warning(display_text)
        elif "משימה" in msg_type:
            st.success(display_text)
        else:
            st.info(display_text)


def _status_to_folder_name(status: str) -> str:
    """Map quote status to folder name: Pending / Approved / Rejected / Signed."""
    s = (status or "").strip()
    if s == "Approved":
        return "Approved"
    if s == "Rejected":
        return "Rejected"
    if s == "Signed":
        return "Signed"
    return "Pending"  # Draft, Sent, Revision Needed


def _status_to_quotes_folder(status: str, year: int, month: int) -> Path:
    """Return hierarchical path: Quotes/{Year}/{Month}/{Status}."""
    status_name = _status_to_folder_name(status)
    return QUOTES_ROOT / str(year) / f"{month:02d}" / status_name


def find_proposal_file(filename: str) -> Path | None:
    """
    Smart File Finder: recursively search for a proposal file (PDF or DOCX) in Quotes folder.
    Returns full path if found, None otherwise.
    """
    if not filename or not str(filename).strip():
        return None
    name = str(filename).strip()
    # Normalize to PDF name for search (File Path in CSV may store PDF path)
    if name.lower().endswith(".docx"):
        search_name = name
    else:
        search_name = name.replace(".docx", ".pdf") if not name.lower().endswith(".pdf") else name
    if not QUOTES_ROOT.exists():
        return None
    for root, _, files in os.walk(QUOTES_ROOT):
        for f in files:
            if f == search_name or f.lower() == search_name.lower():
                return Path(root) / f
    # Also search for the other extension (e.g. if asked for PDF, return DOCX path's sibling)
    alt_name = search_name.replace(".pdf", ".docx") if search_name.lower().endswith(".pdf") else search_name.replace(".docx", ".pdf")
    for root, _, files in os.walk(QUOTES_ROOT):
        for f in files:
            if f == alt_name or f.lower() == alt_name.lower():
                return Path(root) / f
    return None


def _find_pdf_in_quotes_folders(filename: str, extra_dir: Path | None = None) -> Path | None:
    """Search for PDF by filename - uses find_proposal_file for hierarchical structure."""
    base = filename if filename.lower().endswith(".pdf") else filename.replace(".docx", ".pdf")
    found = find_proposal_file(base)
    if found:
        return found
    if extra_dir and extra_dir.exists():
        candidate = extra_dir / base
        if candidate.exists():
            return candidate
    return None


def _find_pdf_by_client_project(client: str, project: str) -> Path | None:
    """Search recursively in Quotes for any PDF containing client and project names."""
    safe_client = sanitize_filename_part(client)
    safe_project = sanitize_filename_part(project)
    if not safe_client or not safe_project:
        return None
    if not QUOTES_ROOT.exists():
        return None
    for root, _, files in os.walk(QUOTES_ROOT):
        for f in files:
            if f.lower().endswith(".pdf"):
                stem = Path(f).stem
                if safe_client in stem and safe_project in stem:
                    return Path(root) / f
    return None


# נתיב Quotes ב-Dropbox API (תואם להעלאה ביצירת הצעה)
DROPBOX_QUOTES_ROOT = "/Studio84/StudioManager/Quotes"


def _normalize_quote_version_str(version_val: str) -> str:
    v = (version_val or "").strip() or "1"
    if not v.upper().startswith("V"):
        v = f"V{v}"
    return v


def find_physical_quote_pdf(r: dict) -> Path | None:
    """
    מאתר קובץ PDF תחת Quotes או temp_proposals לפי שורת ההצעה:
    נתיב File Path (אם קיים), או סריקה לפי שם לקוח/פרויקט/תאריך בשם הקובץ.
    """
    client = (r.get("Client") or "").strip()
    project = (r.get("Project") or "").strip()
    date_val = (r.get("Date") or "").strip()
    file_path = (r.get("File Path") or "").strip()
    version_norm = _normalize_quote_version_str((r.get("Version") or "").strip() or "1")

    safe_client = sanitize_filename_part(client)
    safe_project = sanitize_filename_part(project)

    if file_path:
        p = Path(file_path.strip())
        if p.suffix.lower() == ".pdf" and p.is_file():
            return p
        pdf_name = p.name.replace(".docx", ".pdf") if p.suffix.lower() == ".docx" else p.name
        if not pdf_name.lower().endswith(".pdf"):
            pdf_name = f"{p.stem}.pdf"
        extra_dir = p.parent if p.parent.exists() else None
        found = _find_pdf_in_quotes_folders(pdf_name, extra_dir)
        if found and found.is_file():
            return found

    if not safe_client or not safe_project:
        return None

    date_tokens: list[str] = []
    parsed = _parse_edit_date(date_val) if date_val else None
    if parsed:
        date_tokens.append(parsed.strftime("%Y-%m-%d"))
        date_tokens.append(f"{parsed.day:02d}/{parsed.month:02d}/{parsed.year}")
        date_tokens.append(f"{parsed.day:02d}-{parsed.month:02d}-{parsed.year}")
    if date_val:
        s = date_val.strip()
        if s not in date_tokens:
            date_tokens.append(s)

    def row_pdf_filename_matches(fn: str) -> bool:
        if not fn.lower().endswith(".pdf"):
            return False
        if safe_client not in fn or safe_project not in fn:
            return False
        if date_tokens:
            return any(tok for tok in date_tokens if tok and tok in fn)
        return True

    candidates: list[Path] = []
    for root_dir in (QUOTES_ROOT, TEMP_PROPOSALS_DIR):
        if not root_dir.exists():
            continue
        for walk_root, _, files in os.walk(root_dir):
            for f in files:
                if row_pdf_filename_matches(f):
                    candidates.append(Path(walk_root) / f)

    if not candidates:
        return _find_pdf_by_client_project(client, project)

    if len(candidates) == 1:
        return candidates[0]

    def score(path: Path) -> int:
        fn = path.name
        sc = 0
        if version_norm in fn:
            sc += 10
        for tok in date_tokens:
            if tok and len(tok) >= 8 and tok in fn:
                sc += len(tok)
        return sc

    return max(candidates, key=score)


def _extract_year_month_from_path_or_filename(path_or_name: str, base_name: str) -> tuple[int, int]:
    """
    Extract year and month from path (Quotes/2026/02/...) or filename (Quote_X_V1_2026-02-18).
    Falls back to current date if not found.
    """
    now = date.today()
    # Try path: .../Quotes/2026/02/Pending/... (hierarchical) or .../Quotes/Pending/... (legacy)
    p = Path(path_or_name)
    parts = p.parts
    try:
        idx = parts.index("Quotes")
        if idx + 2 < len(parts):
            y, m = parts[idx + 1], parts[idx + 2]
            if y.isdigit() and len(y) == 4 and m.isdigit():
                year, month = int(y), int(m)
                if 1 <= month <= 12 and 2000 <= year <= 2100:
                    return year, month
    except (ValueError, IndexError):
        pass
    # Try filename: Quote_Client_Project_V1_2026-02-18
    match = re.search(r"(\d{4})-(\d{2})-\d{2}", base_name)
    if match:
        try:
            year, month = int(match.group(1)), int(match.group(2))
            if 1 <= month <= 12 and 2000 <= year <= 2100:
                return year, month
        except (ValueError, IndexError):
            pass
    return now.year, now.month


def _move_quote_files_to_status_folder(
    file_path: str, old_status: str, new_status: str, date_val: str | None = None
) -> str | None:
    """
    Move PDF and DOCX to Quotes/{Year}/{Month}/{New_Status}.
    Uses proposal creation date (date_val from dataframe) for year/month.
    Falls back to current date if date_val is missing or invalid.
    Uses find_proposal_file to locate files in hierarchical structure.
    Returns new PDF path for CSV, or None if nothing moved.
    """
    if not file_path or not file_path.strip():
        return None
    if _status_to_folder_name(old_status) == _status_to_folder_name(new_status):
        return None

    p = Path(file_path.strip())
    base_name = p.stem
    pdf_name = f"{base_name}.pdf"
    docx_name = f"{base_name}.docx"

    # Find files using smart finder
    pdf_src = find_proposal_file(pdf_name)
    docx_src = find_proposal_file(docx_name)
    if not pdf_src and not docx_src:
        return None

    # Target folder: Quotes/{Year}/{Month}/{New_Status} - לפי תאריך יצירת ההצעה
    parsed = _parse_edit_date(date_val or "") if (date_val or "").strip() else None
    if parsed:
        year, month = parsed.year, parsed.month
    else:
        now = datetime.now()
        year, month = now.year, now.month
    new_folder = _status_to_quotes_folder(new_status, year, month)
    new_folder.mkdir(parents=True, exist_ok=True)

    new_pdf_path = None
    try:
        if pdf_src and pdf_src.parent != new_folder:
            dest = new_folder / pdf_name
            shutil.move(str(pdf_src), str(dest))
            new_pdf_path = str(dest.resolve())
        elif pdf_src:
            new_pdf_path = str(pdf_src.resolve())
        if docx_src and docx_src.parent != new_folder:
            shutil.move(str(docx_src), str(new_folder / docx_name))
        if not new_pdf_path and docx_src:
            new_pdf_path = str((new_folder / pdf_name).resolve()) if (new_folder / pdf_name).exists() else str((new_folder / docx_name).resolve())
    except Exception as e:
        st.warning(f"שגיאה בהעברת קבצים: {e}")
        return new_pdf_path if new_pdf_path else (str((new_folder / pdf_name).resolve()) if pdf_src else None)

    return new_pdf_path or (str((new_folder / pdf_name).resolve()) if (pdf_src or docx_src) else None)


def _delete_proposal_and_move_to_trash(
    date_val: str, client_val: str, project_val: str, version_val: str, file_path: str
) -> None:
    """
    Move proposal PDF and DOCX files to Quotes/Trash.
    Uses find_proposal_file to locate files. Does not modify CSV.
    """
    trash_dir = QUOTES_ROOT / "Trash"
    trash_dir.mkdir(parents=True, exist_ok=True)

    if file_path and str(file_path).strip():
        p = Path(file_path.strip())
        base_name = p.stem
    else:
        safe_client = sanitize_filename_part(client_val or "")
        safe_project = sanitize_filename_part(project_val or "")
        version_norm = (version_val or "1").strip()
        if not version_norm.upper().startswith("V"):
            version_norm = f"V{version_norm}"
        parsed = _parse_edit_date(date_val or "") if (date_val or "").strip() else None
        file_date = parsed.strftime("%Y-%m-%d") if parsed else date.today().strftime("%Y-%m-%d")
        base_name = f"Quote_{safe_client}_{safe_project}_{version_norm}_{file_date}"

    pdf_name = f"{base_name}.pdf"
    docx_name = f"{base_name}.docx"

    for fname in (pdf_name, docx_name):
        src = find_proposal_file(fname)
        if src:
            dest = trash_dir / fname
            if dest.exists():
                dest.unlink()
            shutil.move(str(src), str(dest))


# --- כתובות מייל קבועות ---
# הגדרות מייל קבועות
EMAIL_ACCOUNTING = 'account@studio84.co.il'
EMAIL_ERAN = 'eran@studio84.co.il'
# המייל שלך (השלם את הכתובת החסרה או השאר למילוי)
EMAIL_MYSELF = 'talalcheh84@gmail.com'

# --- אנשי צוות למרכז התנעה (Kickoff) ---
TEAM_MEMBERS = {
    'ערן - מנהל קריאייטיב': 'eran@studio84.co.il',
    'טלי - מנהלת קריאייטיב': 'talalcheh84@studio84.co.il',
    'ליאור - תלת מימד': 'lior@studio84.co.il',
    'אחיעד - תלת מימד': 'achiad@studio84.co.il',
    'אור - עורך ובמאי': 'or@studio84.co.il',
    "ג׳ורג׳ - תלת מימד": 'George.berdichevsky@gmail.com',
    'מודליסט 1': 'lior@studio84.co.il',
    'מודליסט 2': 'achiad@studio84.co.il',
}
# שמות קצרים לתצוגה (ללא תפקידים) - משמש ב-selectbox, multiselect, וטקסט המייל
TEAM_DISPLAY_NAMES = [name.split('-')[0].strip() for name in TEAM_MEMBERS.keys()]
# מיפוי שם קצר -> מייל (לשליחת מייל)
TEAM_EMAIL_BY_SHORT = {name.split('-')[0].strip(): email for name, email in TEAM_MEMBERS.items()}
# מספר וואצאפ של ערן (placeholder - ניתן למלא/לשנות)
WHATSAPP_ERAN = "972547641984"

# שמות צוות + מיילים (שליחת מיילים; current_user נקבע בהתחברות בלבד)
TEAM_EMAILS = {
    "טל": "talalcheh84@gmail.com",
    "ערן": "eran@studio84.co.il",
    "מיה": "maya@studio84.co.il",
    "ליאור": "lior@studio84.co.il",
    "אחיעד": "achiad@studio84.co.il",
    "אור": "or@studio84.co.il",
    "ג׳ורג׳": "George.berdichevsky@gmail.com",
    "הנהלת חשבונות": "accounts@studio84.co.il",
}
MANAGEMENT_USERS = frozenset({"טל", "ערן"})

# ניווט ראשי — תוויות יחידות לרשימת ה-radio ול-branch של if/elif
NAV_MAIN_PROJECT_ROOM = "📊 חדר מצב (קנבן)"
NAV_MY_TASKS = "🖥️ המשימות שלי (מוניטור צוות)"
NAV_PROJECT_FOLDERS = "📁 תיקי פרויקטים (מידע וקשר)"
NAV_MGMT_SEPARATOR = "--- אזור ניהול ---"
NAV_QUOTES_FINANCE = "📝 הצעות מחיר ופיננסים"
NAV_TASKS_PRODUCTION = "🎯 הפקת פרויקטים והקצאת משימות"
NAV_CRM = "📞 לקוחות ואנשי קשר"

# מיפוי שם משתמש ממסך ההתחברות (אנגלית, lower) → שם קנוני ב-TEAM_EMAILS / משימות.
# לאחר הסרת בורר הזהות בסיידבר, current_user נקבע רק בכניסה — יש להרחיב כאן לפי שמות ב-secrets.
LOGIN_USERNAME_TO_CURRENT_USER: dict[str, str] = {
    "tal": "טל",
    "tali": "טל",
    "eran": "ערן",
    "maya": "מיה",
    "lior": "ליאור",
    "achiad": "אחיעד",
    "or": "אור",
    "george": "ג׳ורג׳",
}


def _session_current_user_from_login(username_normalized: str) -> str:
    """שם המשתמש לריצה והרשאות — תואם למפתחות TEAM_EMAILS; לא מגיע מקלט UI אחר."""
    u = (username_normalized or "").strip().lower()
    if not u:
        return ""
    if u in LOGIN_USERNAME_TO_CURRENT_USER:
        return LOGIN_USERNAME_TO_CURRENT_USER[u]
    if u in TEAM_EMAILS:
        return u
    su = (username_normalized or "").strip()
    if su in TEAM_EMAILS:
        return su
    return u


# צבעים קבועים לגאנט (לפי מפתח TEAM_EMAILS)
TEAM_GANTT_COLOR_HEX: dict[str, str] = {
    "טל": "#1f77b4",  # כחול
    "ערן": "#2ca02c",  # ירוק
    "מיה": "#e377c2",  # ורוד
    "ליאור": "#ff7f0e",
    "אחיעד": "#9467bd",
    "אור": "#17becf",
    "ג׳ורג׳": "#8c564b",
    "הנהלת חשבונות": "#bcbd22",
    "(ללא אחראי)": "#7f7f7f",
    # חופשה/היעדרות — צהוב מובהק (לא לפי עובד)
    "חופשה/היעדרות": "#e6b800",
}


def _task_team_key_for_color(task_assignee: str) -> str:
    """מחזיר מפתח מ-TEAM_EMAILS להתאמת צבע בגאנט."""
    for k in TEAM_EMAILS:
        if _assignee_matches_team_key(task_assignee or "", k):
            return k
    return (task_assignee or "").strip() or "(ללא אחראי)"


def _is_out_of_office_task(row_or_name) -> bool:
    """משימת חופשה/היעדרות — לפי שם המשימה (סוג המשימה שנשמר בשם)."""
    if isinstance(row_or_name, dict):
        name = (row_or_name.get("שם משימה") or "").strip()
    elif hasattr(row_or_name, "get") and not isinstance(row_or_name, (str, bytes)):
        name = str(row_or_name.get("שם משימה", "") or "").strip()
    else:
        name = str(row_or_name or "").strip()
    return name == TASK_TYPE_OOO


def _ooo_event_title_from_row(row) -> str:
    """כותרת תצוגה לחופשה בלוח שנה ובגאנט — לפי איש צוות (עמודת הוקצה ל)."""
    assignee_raw = str(row.get("הוקצה ל", "") or "").strip() if hasattr(row, "get") else ""
    team_key = _task_team_key_for_color(assignee_raw)
    if team_key and team_key != "(ללא אחראי)":
        return f"🌴 חופשה: {team_key}"
    if assignee_raw:
        return f"🌴 חופשה: {assignee_raw}"
    return "🌴 חופשה/היעדרות"


def _format_task_date_cell_for_edit_label(val) -> str:
    """מחרוזת תאריך להצגה ב-selectbox של עריכת משימה (dd/mm/yyyy)."""
    p = _parse_task_date(str(val or "").strip()) if val else None
    if p is None:
        return (str(val).strip() if val is not None else "") or "—"
    return p.strftime("%d/%m/%Y")


def _format_edit_task_select_label(r: dict) -> str:
    """תווית לשורת בחירה ב-selectbox 'עריכת משימה קיימת' — חופשה בולטת, אחרת פרויקט|סוג|צוות."""
    task_name = (r.get("שם משימה") or "").strip()
    assignee_raw = str(r.get("הוקצה ל", "") or "")
    team_key = _task_team_key_for_color(assignee_raw)
    if "חופשה" in task_name:
        start_str = _format_task_date_cell_for_edit_label(r.get("תאריך התחלה"))
        end_str = _format_task_date_cell_for_edit_label(r.get("תאריך יעד"))
        return f"🌴 חופשה: {team_key} | מ-{start_str} עד-{end_str}"
    proj = _task_project_stored_to_label(str(r.get("פרויקט") or "").strip())
    return f"{proj} | {task_name} | {team_key}"


def _gantt_opacity_for_status(status_val) -> float:
    if (str(status_val or "").strip().lower()) == "הסתיים":
        return 0.5
    return 1.0


def _find_task_row_index_in_full_list(full_rows: list[dict], selected: dict) -> int | None:
    """מאתר אינדקס שורה ברשימת המשימות המלאה — לפי מזהה משימה או מפתח שורה."""
    tid = (selected.get("מזהה משימה") or "").strip()
    if tid:
        for i, r in enumerate(full_rows):
            if (r.get("מזהה משימה") or "").strip() == tid:
                return i
        return None

    def _rk(r: dict) -> tuple:
        return (
            str(r.get("פרויקט", "") or ""),
            str(r.get("שם משימה", "") or ""),
            str(r.get("הוקצה ל", "") or ""),
            str(r.get("תאריך התחלה", "") or ""),
            str(r.get("תאריך יעד", "") or ""),
        )

    key = _rk(selected)
    for i, r in enumerate(full_rows):
        if _rk(r) == key:
            return i
    return None


def _ensure_task_rows_have_ids(rows: list[dict]) -> list[dict]:
    """מבטיח שכל משימה שנשמרת לגיליון כוללת מזהה ייחודי."""
    out: list[dict] = []
    for r in rows:
        rr = dict(r)
        if not (rr.get("מזהה משימה") or "").strip():
            rr["מזהה משימה"] = str(uuid.uuid4())
        out.append(rr)
    return out


def _assignee_cell_matches_login(task_assignee: str, selected_user: str) -> bool:
    """התאמת ערך עמודת 'הוקצה ל' (איש צוות / אחראי) למשתמש הנבחר בצד."""
    a = (task_assignee or "").strip()
    s = (selected_user or "").strip()
    if not a or not s:
        return False

    def _matches(sel: str) -> bool:
        if not sel:
            return False
        return a == sel or a.startswith(sel + " ") or a.startswith(sel + "-")

    if _matches(s):
        return True
    if s == "טל" and _matches("טלי"):
        return True
    if s == "טלי" and _matches("טל"):
        return True
    return False


def _task_row_matches_view_filter(task: dict, view_filter: str) -> bool:
    """סינון תצוגת משימות לפי אחראי — 'הצג הכל' מציג הכל; לא משנה current_user או הרשאות."""
    vf = (view_filter or "").strip()
    if not vf or vf == "הצג הכל":
        return True
    cell = task.get("הוקצה ל") or ""
    if _assignee_matches_team_key(cell, vf):
        return True
    return _assignee_cell_matches_login(cell, vf)


def _assignee_matches_team_key(task_assignee: str, team_key: str) -> bool:
    """התאמת 'הוקצה ל' למפתח ב-TEAM_EMAILS: טל/טלי + איחוד גרש (ג'ורג' / ג׳ורג׳)."""
    if _assignee_cell_matches_login(task_assignee, team_key):
        return True
    a = (task_assignee or "").strip()
    k = (team_key or "").strip()
    if not a or not k:
        return False

    def _norm(x: str) -> str:
        return x.replace("'", "\u05f3").replace("\u2019", "\u05f3")

    an, kn = _norm(a), _norm(k)
    if an == kn or an.startswith(kn + " ") or an.startswith(kn + "-"):
        return True
    return False


def _team_email_for_task_assignee(assignee: str) -> str | None:
    """מחזיר כתובת מייל לפי ערך 'הוקצה ל' (TASK_TEAM) מתוך TEAM_EMAILS."""
    for key in TEAM_EMAILS:
        if _assignee_matches_team_key(assignee, key):
            em = TEAM_EMAILS.get(key)
            if em and str(em).strip():
                return str(em).strip()
    return None


def send_task_assignment_email_tali(
    to_email: str,
    project_label: str,
    task_name: str,
    task_description: str,
    due_str: str,
    assignee_name: str,
    cc_emails: list[str] | None = None,
) -> bool:
    """
    מייל על הקצאת משימה — כולל סוג המשימה ופירוט/הערות אם הוזנו.
    """
    try:
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = str(st.secrets["email_tali"]["sender_email"]).strip()
        password = str(st.secrets["email_tali"]["password"]).strip()
        if not smtp_server or not sender_email or not password:
            st.error("חסרים נתוני אימייל ב-secrets (מקטע [email_tali])")
            return False
        to_clean = (to_email or "").strip()
        if not to_clean:
            st.warning("אין כתובת נמען תקינה לשליחת מייל משימה.")
            return False
        cc_clean = [e.strip() for e in (cc_emails or []) if e and str(e).strip() and str(e).strip() != to_clean]
        desc = (task_description or "").strip()
        name = (task_name or "").strip()
        detail_for_email = desc if desc else name
        html_body = f"""
<!DOCTYPE html>
<html dir="rtl" lang="he">
<head><meta charset="utf-8"></head>
<body style="font-family: Arial, sans-serif; direction: rtl;">
<h2>🎯 משימה חדשה הוקצתה</h2>
<table style="border-collapse: collapse;">
<tr><td style="padding:6px;"><strong>פרויקט:</strong></td><td>{html.escape(project_label)}</td></tr>
<tr><td style="padding:6px;vertical-align:top;"><strong>תיאור המשימה:</strong></td><td style="white-space:pre-wrap;">{html.escape(detail_for_email)}</td></tr>
<tr><td style="padding:6px;"><strong>הוקצה ל:</strong></td><td>{html.escape(assignee_name)}</td></tr>
<tr><td style="padding:6px;"><strong>תאריך יעד:</strong></td><td>{html.escape(due_str)}</td></tr>
</table>
<p>בהצלחה!</p>
</body>
</html>
"""
        plain_body = (
            f"פרויקט: {project_label}\n"
            f"תיאור המשימה: {detail_for_email}\n"
            f"הוקצה ל: {assignee_name}\n"
            f"תאריך יעד: {due_str}\n"
        )
        msg = EmailMessage()
        subj_task = (name[:60] + "…") if len(name) > 63 else name
        msg["Subject"] = f"משימה חדשה: {subj_task} — {project_label}"
        msg["From"] = sender_email
        msg["To"] = to_clean
        if cc_clean:
            msg["Cc"] = ", ".join(cc_clean)
        msg.set_content(plain_body)
        msg.add_alternative(html_body, subtype="html")

        def _try_send(server) -> bool:
            try:
                server.send_message(msg)
            except Exception as e:
                st.warning(f"שגיאה בשליחת מייל משימה: {e}")
                return False
            return True

        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, password)
                if not _try_send(server):
                    return False
        else:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, password)
                if not _try_send(server):
                    return False
        return True
    except smtplib.SMTPAuthenticationError as e:
        st.warning(f"שגיאת אימות SMTP במייל משימה: {e}")
        return False
    except Exception as e:
        st.warning(f"שגיאה בשליחת מייל משימה: {e}")
        return False


# סיסמת מנהל (Admin)
ADMIN_PASSWORD = "8484"
# שמות/תפקידים שמוגדרים כמנהלים - לצורכי הרשאות וסינון תצוגה
ADMIN_NAMES = {"Admin (טלי / ערן)", "Admin", "טלי", "ערן"}


st.set_page_config(page_title="דשבורד ניהול סטודיו", layout="wide")


QUOTES_LOG_COLUMNS = [
    "Date",
    "Client",
    "Project",
    "Version",
    "Total Price",
    "Status",
    "File Path",
    "Signed File Path",
    "Custom Item Desc",
    "Custom Item Price",
    "מקדמה שולמה",
    "סכום מקדמה",
    "גמר חשבון שולם",
    "freelancer_cost",
    "misc_expenses",
]

# Full quote form data for edit/pre-fill (quotes tab in Google Sheets)
# Includes Status, File Path, Signed File Path, Total Price for quote management
QUOTES_CSV_COLUMNS = [
    "Date", "Client", "Project", "Version", "Quote Number", "Contact Person", "Client Email",
    "Quote Subject", "show_exterior", "show_interior", "show_drone", "show_video", "show_shots",
    "scope_of_work", "work_process", "delivery_time", "video_terms",
    "price_exterior", "base_views_count", "price_ext_extra",
    "price_interior", "price_int_extra", "price_int_space",
    "price_drone", "price_video", "price_shot_unit", "shots_count",
    "include_price_exterior", "include_price_ext_extra", "include_price_interior",
    "include_price_int_extra", "include_price_int_space", "include_price_drone",
    "include_price_video", "include_total_shots_price",
    "model_update_val", "view_update_val", "extra_view_val",
    "custom_item_desc", "custom_item_price",
    "Status", "File Path", "Signed File Path", "Total Price",
    "Client Phone", "Architect Contact", "Project Special Notes",
    "שלב עבודה",
    "מקדמה שולמה",
    "סכום מקדמה",
    "גמר חשבון שולם",
    "freelancer_cost",
    "misc_expenses",
]

ALLOWED_QUOTE_STATUSES = ["Draft", "Sent", "Approved", "Revision Needed", "Rejected", "Signed", "הומר לפרויקט"]
DEFAULT_QUOTE_STATUS = "Draft"
PROJECTS_DB_COLUMNS = [
    "Project ID",
    "Client",
    "Project Name",
    "Manager",
    "Team",
    "Status",
    "Start Date",
    "Dropbox_Main",
    "Dropbox_Upload",
    "Dropbox_Deliverables",
    'היקף כספי (₪)',
]
ALLOWED_PROJECT_STATUSES = [
    "ממתין להתחלה",
    "בעבודה",
    "נשלח לסבב הערות 1",
    "נשלח לסבב הערות 2",
    "ממתין לאדריכל/לקוח",
    "הוקפא",
    "הסתיים",
    "חשבונית נשלחה",
    "שולם",
]
DEFAULT_PROJECT_STATUS = "ממתין להתחלה"

# שלבי עבודה — לוח קנבן במסך חדר מצב (מוניטור פרויקטים)
KANBAN_WORK_STAGES = [
    "התקבל",
    "במידול",
    "חומרים ותאורה",
    "ממתין לפידבק",
    "רינדור סופי",
    "אחר",
    "נמסר",
]

# סטטוסים למוניטור צוות תלת-מימד (זרימת עבודה)
MONITOR_3D_STATUS_OPTIONS = [
    "ממתין להתחלה",
    "במידול",
    "חומרים ותאורה",
    "ברינדור",
    "פוסט-פרודקשן",
    "מוכן לאישור לקוח",
]

# --- projects.csv (פרויקטים פעילים) ---
PROJECTS_CSV_COLUMNS = ["ID", "Client", "Project", "Deadline", "Team", "Status", "Budget_Hours", 'היקף כספי (₪)', "אנשי קשר מקושרים"]
PROJECTS_CSV_STATUSES = ["Active", "Done"]

PROJECT_MANAGERS = ["ערן", "טלי"]
PROJECT_TEAM_MEMBERS = ["ג'ורג'", "מיה", "ליאור", "אור", "אחיעד"]
TASK_TEAM = ["ערן", "טלי", "ג'ורג'", "מיה", "ליאור", "אור", "אחיעד"]
# עמודות משימות מסונכרנות: טופס, מוניטור, גאנט, לוח שנה
TASKS_LOG_COLUMNS = [
    "מזהה משימה",
    "פרויקט",
    "שם משימה",
    "תיאור המשימה",
    "הוקצה ל",
    "תאריך התחלה",
    "תאריך יעד",
    "סטטוס",
]
TASK_STATUSES = ["To Do", "In Progress", "Done", "Stuck"]
# סטטוסים לעריכת משימה (טאב מוניטור) ולטבלת המשימות
TASK_EDIT_STATUS_OPTIONS = [
    "ממתין",
    "בעבודה",
    "ממתין לפידבק לקוח",
    "הסתיים",
    "הושלם",
]
TASK_PRIORITIES = ["רגיל", "דחוף", "קריטי"]

# סוגי משימות להקצאה לצוות (טאב 'הקצאת משימות לצוות')
TASK_TYPE_OOO = "🌴 חופשה/היעדרות"
TASKS_PROJECT_OOO_DEFAULT = "סטודיו"
# משימות סטודיו כלליות (ללא פרויקט לקוח) — נשמר בעמודת «פרויקט» בגיליון tasks
TASKS_PROJECT_GENERAL_STUDIO = "כללי / סטודיו"
TASKS_PROJECT_GENERAL_STUDIO_SELECT = "🏢 כללי / סטודיו"
TASK_TYPE_OPTIONS = [
    "מידול (Modeling)",
    "חומרים ותאורה (Texturing & Lighting)",
    "רינדור (Rendering)",
    "פוסט-פרודקשן (Post-Production)",
    "הערות לקוח / תיקונים",
    TASK_TYPE_OOO,
    "אחר",
]
CALENDAR_TASK_COLOR_DEFAULT = "#4b8bbe"
CALENDAR_TASK_COLOR_OOO = "#e6b800"

# סטטוסים שנחשבים "הושלם" - משימות עם סטטוס כזה לא יוצגו ברשימה
DONE_STATUSES = ("done", "בוצע", "הושלם", "completed", "הסתיים")

def _parse_date_safe(val, date_column_name: str = "Date") -> date | None:
    """ממיר ערך תאריך (מחרוזת או אחר) ל-date. מחזיר None אם ריק/לא תקין."""
    if val is None or (isinstance(val, str) and not (val or "").strip()):
        return None
    if isinstance(val, date):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        dt = pd.to_datetime(val, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.date() if hasattr(dt, "date") else dt
    except Exception:
        return None

def _filter_tasks_by_status_and_date(
    tasks: list[dict],
    date_col: str,
    status_col: str = "Status",
    done_statuses: tuple[str, ...] = DONE_STATUSES,
) -> list[dict]:
    """
    מסנן משימות: סטטוס לא 'בוצע' ותאריך <= היום (או תאריך ריק - יוצגו באיחור).
    מחזיר רשימה ממוינת: באיחור ראשון, אחר כך היום.
    תאים ריקים: מטפלים ב-pd.to_datetime(errors='coerce') - תאריך לא תקין יוצג באיחור.
    """
    today_dt = date.today()
    result = []
    for t in tasks:
        status = (t.get(status_col) or "").strip().lower()
        if status in done_statuses:
            continue
        parsed = _parse_date_safe(t.get(date_col), date_col)
        if parsed is not None and parsed > today_dt:
            continue  # תאריך עתידי - לא מציגים
        # תאריך ריק/לא תקין: מציגים (כאילו באיחור) כדי לא לאבד משימות
        sort_date = parsed if parsed is not None else today_dt - timedelta(days=1)
        result.append((t, sort_date))
    # מיון: באיחור (תאריך < היום) ראשון, אחר כך היום
    result.sort(key=lambda x: (0 if x[1] < today_dt else 1, x[1]))
    return [r[0] for r in result]


def _filter_tasks_open_not_done(
    tasks: list[dict],
    status_col: str = "סטטוס",
    done_statuses: tuple[str, ...] = DONE_STATUSES,
) -> list[dict]:
    """
    משימות פתוחות לפי סטטוס בלבד (לא מסנן תאריכי יעד עתידיים).
    משמש את מוניטור תמונת המצב הצוותית כדי שמשימות שהוקצו יופיעו גם לפני תאריך היעד.
    """
    done_lower = {s.lower() for s in done_statuses}
    result = []
    for t in tasks:
        status = (t.get(status_col) or "").strip().lower()
        if status in done_lower:
            continue
        result.append(t)
    return result


def _is_done_daily(val) -> bool:
    """בודק אם משימה יומית סומנה כהושלמה (Is Done)."""
    if val is True or val == 1:
        return True
    if isinstance(val, str) and (val or "").strip().lower() in ("true", "1", "yes"):
        return True
    return False

def _filter_daily_tasks_by_date(
    tasks: list[dict],
    date_col: str = "Date",
    status_col: str = "Status",
) -> list[dict]:
    """
    מסנן משימות יומיות: לא הושלמו (Is Done=False), סטטוס לא בוצע, תאריך <= היום.
    מחזיר רשימה ממוינת: באיחור ראשון, אחר כך היום.
    תאים ריקים: תאריך לא תקין יוצג באיחור כדי לא לאבד משימות.
    """
    today_dt = date.today()
    result = []
    for t in tasks:
        if _is_done_daily(t.get("Is Done")):
            continue
        status = (t.get(status_col) or "").strip().lower()
        if status in DONE_STATUSES:
            continue
        parsed = _parse_date_safe(t.get(date_col), date_col)
        if parsed is not None and parsed > today_dt:
            continue  # תאריך עתידי - לא מציגים
        sort_date = parsed if parsed is not None else today_dt - timedelta(days=1)
        result.append((t, sort_date))
    result.sort(key=lambda x: (0 if x[1] < today_dt else 1, x[1]))
    return [r[0] for r in result]

# רשימת משתמשים במסך הכניסה (ניתן לעדכן שמות עובדים)
TEAM_MEMBERS_LOGIN = ['Admin (טלי / ערן)', "ג'ורג'", 'אור', 'ליאור', 'אחיעד', 'מיה']
# רשימת Assignees למשימות יומיות (לבחירת מנהל)
DAILY_TASKS_ASSIGNEES = ["טלי", "ערן", "ג'ורג'", "עובד 2", "עובד 3"]
DAILY_TASKS_COLUMNS = ["Task Name", "Project", "Assignee", "Date", "Status", "Is Done", "Flexible"]
PROJECT_TEMPLATE_OPTIONS = [
    "1. סקיצות לאישור מבטים וזוויות",
    "2. מודל מלא (חומרים, צמחייה ותאורה)",
    "3. סבב הערות 1",
    "4. רינדור סופי והפקת קבצים",
]
PROJECT_MANAGER_EMAILS = {
    "ערן": EMAIL_ERAN,
    "טלי": EMAIL_MYSELF,  # ניתן לעדכן לכתובת ייעודית אם קיימת
}

# --- אנשי קשר לפרויקט (גיליון project_contacts בגוגל שיטס) ---
PROJECT_CONTACTS_COLUMNS = [
    "Project",
    "Role Category",
    "Office/Company Name",
    "Contact Name",
    "Email",
    "Phone",
    "Notes",
]
ROLE_CATEGORIES = [
    "משרד אדריכלים",
    "עיצוב פנים",
    "אדריכלות נוף",
    "לקוח",
    "פיקוח",
    "ספקים",
    "אחר",
]

# --- CRM - אנשי קשר (גיליון contacts בגוגל שיטס) ---
CONTACTS_COLUMNS = [
    "שם מלא",
    "חברה / משרד אדריכלים",
    "תפקיד",
    "טלפון",
    "אימייל",
    "סוג איש קשר",
]
CONTACT_TYPE_OPTIONS = ["לקוח", "אדריכל", "יזם", "מפקח", "קבלן", "אחר"]


DROPBOX_BASE_FOLDER = "/Studio84/StudioManager/Projects"


def create_studio_dropbox_structure(project_name: str) -> tuple[str, str, str] | None:
    """
    יוצר מבנה תיקיות בדרופבוקס לפרויקט: תיקייה ראשית + 2 תת-תיקיות.
    מחזיר (main_link, upload_link, deliverables_link) או None במקרה של שגיאה.
    - main_link: קישור שיתוף לתיקייה הראשית
    - upload_link: קישור File Request לתת-תיקיית חומרים נכנסים
    - deliverables_link: קישור שיתוף לתת-תיקיית תוצרים
    """
    if not (project_name or "").strip():
        return None
    clean_name = (
        project_name.strip()
        .replace(" ", "_")
        .replace("/", "-")
        .replace("\\", "-")
    )
    clean_name = re.sub(r'[/\\:*?"<>|]', "_", clean_name) or "project"
    base_path = "/Studio84/StudioManager/Projects"
    main_folder = f"{base_path}/{clean_name}"
    materials_folder = f"{main_folder}/01_Client_Materials"
    deliverables_folder = f"{main_folder}/02_Studio_Deliverables"

    try:
        access_token = _get_dropbox_access_token()
        if not access_token:
            return None
        dbx = dropbox.Dropbox(access_token)
        try:
            dbx.users_get_current_account()
        except Exception:
            return None

        # Dropbox Business - שורש ל-Team Space במקום Member Space
        path_root = None
        if PathRoot:
            try:
                ns_id = st.secrets.get("DROPBOX_NAMESPACE_ID")
                if ns_id and str(ns_id).strip():
                    path_root = PathRoot.namespace_id(str(ns_id).strip())
                else:
                    acc = dbx.users_get_current_account()
                    if acc and getattr(acc, "root_info", None):
                        root_ns = getattr(acc.root_info, "root_namespace_id", None)
                        if root_ns:
                            path_root = PathRoot.root(root_ns)
            except Exception:
                pass
        if path_root:
            dbx = dbx.with_path_root(path_root)

        # יצירת תיקיות הורה אם הנתיב מקונן
        parts = [p for p in main_folder.split("/") if p]
        for i in range(1, len(parts)):
            parent = "/" + "/".join(parts[:i])
            try:
                dbx.files_create_folder_v2(parent)
            except dropbox.exceptions.ApiError:
                pass

        # תיקייה ראשית
        try:
            dbx.files_create_folder_v2(main_folder)
        except dropbox.exceptions.ApiError:
            pass

        # תת-תיקייה לחומרים נכנסים
        try:
            dbx.files_create_folder_v2(materials_folder)
        except dropbox.exceptions.ApiError:
            pass

        # תת-תיקייה לתוצרים
        try:
            dbx.files_create_folder_v2(deliverables_folder)
        except dropbox.exceptions.ApiError:
            pass

        def _get_shared_link(path: str) -> str:
            try:
                link_metadata = dbx.sharing_create_shared_link_with_settings(path)
                return link_metadata.url
            except dropbox.exceptions.ApiError as e:
                err = getattr(e, "error", None)
                if err is not None and getattr(err, "is_shared_link_already_exists", lambda: False)():
                    meta = getattr(err, "get_shared_link_already_exists", lambda: None)()
                    if meta and getattr(meta, "is_metadata", lambda: False)():
                        link_meta = meta.get_metadata()
                        if link_meta:
                            return link_meta.url
                    try:
                        links_result = dbx.sharing_list_shared_links(path=path, direct_only=True)
                        if links_result.links:
                            return links_result.links[0].url
                    except Exception:
                        pass
                raise

        main_link = _get_shared_link(main_folder)
        deliverables_link = _get_shared_link(deliverables_folder)

        # File Request לחומרים נכנסים - try-except נפרד כדי שלא יכשיל את יתר הלינקים
        upload_link = ""
        try:
            dest = f"{base_path}/{clean_name}/01_Client_Materials"
            if not dest.startswith("/"):
                dest = "/" + dest
            req = dbx.file_requests_create(
                title=f"Upload Materials - {clean_name}",
                destination=dest,
            )
            if req is not None and hasattr(req, "url"):
                url_val = req.url
                if url_val and str(url_val).strip() and str(url_val).strip() != "0":
                    upload_link = str(url_val).strip()
        except Exception:
            upload_link = ""

        return (main_link, upload_link, deliverables_link)
    except Exception:
        return None


def create_dropbox_folder_and_link(project_name: str, folder_path: str | None = None) -> str | None:
    """
    יוצר תיקייה בדרופבוקס ומחזיר קישור שיתוף.
    כל התיקיות נוצרות בתיקייה המשותפת של הצוות (Studio84/StudioManager/Projects).
    מחזיר את ה-URL או None/מחרוזת ריקה במקרה של שגיאה. לא קורא ל-st.error כדי לא לעצור את הסקריפט.
    """
    if not (project_name or "").strip():
        return ""
    clean_project_name = (
        project_name.strip()
        .replace(" ", "_")
        .replace("/", "-")
        .replace("\\", "-")
    )
    clean_project_name = re.sub(r'[/\\:*?"<>|]', "_", clean_project_name) or "project"
    folder_path = f"{DROPBOX_BASE_FOLDER}/{clean_project_name}"
    try:
        access_token = _get_dropbox_access_token()
        if not access_token:
            return ""
        dbx = dropbox.Dropbox(access_token)
        path_root = None
        if PathRoot:
            try:
                ns_id = st.secrets.get("DROPBOX_NAMESPACE_ID")
                if ns_id and str(ns_id).strip():
                    path_root = PathRoot.namespace_id(str(ns_id).strip())
                else:
                    acc = dbx.users_get_current_account()
                    if acc and getattr(acc, "root_info", None):
                        root_ns = getattr(acc.root_info, "root_namespace_id", None)
                        if root_ns:
                            path_root = PathRoot.root(root_ns)
            except Exception:
                pass
        if path_root:
            dbx = dbx.with_path_root(path_root)
        # יצירת תיקיות הורה אם הנתיב מקונן (למשל /Studio84/StudioManager/Projects)
        parts = [p for p in folder_path.split("/") if p]
        for i in range(1, len(parts)):
            parent = "/" + "/".join(parts[:i])
            try:
                dbx.files_create_folder_v2(parent)
            except dropbox.exceptions.ApiError:
                pass  # תיקייה כבר קיימת
        try:
            dbx.files_create_folder_v2(folder_path)
        except dropbox.exceptions.ApiError:
            pass
        try:
            link_metadata = dbx.sharing_create_shared_link_with_settings(folder_path)
            return link_metadata.url
        except dropbox.exceptions.ApiError as e:
            err = getattr(e, 'error', None)
            if err is not None and getattr(err, 'is_shared_link_already_exists', lambda: False)():
                meta = getattr(err, 'get_shared_link_already_exists', lambda: None)()
                if meta and getattr(meta, 'is_metadata', lambda: False)():
                    link_meta = meta.get_metadata()
                    if link_meta:
                        return link_meta.url
                try:
                    links_result = dbx.sharing_list_shared_links(path=folder_path, direct_only=True)
                    if links_result.links:
                        return links_result.links[0].url
                except Exception:
                    pass
            raise
    except Exception:
        return ""


def send_kickoff_email(
    project_name: str,
    client: str,
    deadline_str: str,
    main_link: str = "",
    upload_link: str = "",
    deliverables_link: str = "",
) -> bool:
    """
    שולח מייל התנעת פרויקט לצוות עם לינקים ישירים לתיקיות הדרופבוקס.
    משתמש ב-st.secrets['email']: smtp_server, smtp_port, sender_email, password.
    תומך ב-Webmail (SMTP פרטי) עם פורט 465 (SSL) או 587 (STARTTLS).
    מחזיר True אם השליחה הצליחה, False אחרת.
    """
    try:
        email_config = st.secrets.get("email", {}) or {}
        smtp_server = (email_config.get("smtp_server") or "").strip()
        smtp_port = int(email_config.get("smtp_port", 465))
        sender_email = (email_config.get("sender_email") or "").strip()
        password = (email_config.get("password") or "").strip()
        recipient = (email_config.get("recipient") or st.secrets.get("TEAM_EMAIL", "") or "").strip()
        if not smtp_server or not sender_email or not password or not recipient:
            st.error("חסרים נתוני אימייל ב-secrets (smtp_server, sender_email, password, recipient)")
            return False
        links_html = []
        if main_link and main_link.startswith("http"):
            links_html.append(f'<li><a href="{html.escape(main_link)}">📂 תיקייה ראשית</a></li>')
        if upload_link and upload_link.startswith("http"):
            links_html.append(f'<li><a href="{html.escape(upload_link)}">📥 בקשת חומרים (העלאה ללקוח)</a></li>')
        if deliverables_link and deliverables_link.startswith("http"):
            links_html.append(f'<li><a href="{html.escape(deliverables_link)}">📤 תיקיית תוצרים</a></li>')
        links_section = f"<ul>{''.join(links_html)}</ul>" if links_html else "<p>לינקים לא זמינים.</p>"
        html_body = f"""
<!DOCTYPE html>
<html dir="rtl" lang="he">
<head><meta charset="utf-8"></head>
<body style="font-family: Arial, sans-serif; direction: rtl;">
<h2>🚀 התנעת פרויקט חדש</h2>
<table style="border-collapse: collapse;">
<tr><td style="padding:6px;"><strong>פרויקט:</strong></td><td>{html.escape(project_name)}</td></tr>
<tr><td style="padding:6px;"><strong>לקוח:</strong></td><td>{html.escape(client)}</td></tr>
<tr><td style="padding:6px;"><strong>דדליין:</strong></td><td>{html.escape(deadline_str)}</td></tr>
</table>
<h3>לינקים לתיקיות הדרופבוקס</h3>
{links_section}
<p>בהצלחה!</p>
</body>
</html>
"""
        msg = EmailMessage()
        msg["Subject"] = f"התנעת פרויקט: {project_name} - {client}"
        msg["From"] = sender_email
        msg["To"] = recipient
        msg.set_content(f"פרויקט: {project_name}\nלקוח: {client}\nדדליין: {deadline_str}\n\nלינקים:\n{main_link}\n{upload_link}\n{deliverables_link}")
        msg.add_alternative(html_body, subtype="html")
        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as smtp:
                smtp.login(sender_email, password)
                smtp.send_message(msg)
        else:
            with smtplib.SMTP(smtp_server, smtp_port) as smtp:
                smtp.starttls()
                smtp.login(sender_email, password)
                smtp.send_message(msg)
        return True
    except smtplib.SMTPAuthenticationError as e:
        st.error(f"שגיאת אימות SMTP: {e}")
        return False
    except smtplib.SMTPConnectError as e:
        st.error(f"שגיאת התחברות לשרת SMTP: {e}")
        return False
    except smtplib.SMTPException as e:
        st.error(f"שגיאת SMTP: {e}")
        return False
    except Exception as e:
        st.error(f"שגיאה בשליחת המייל: {e}")
        return False


def send_quote_email_via_smtp(
    to_email: str,
    subject: str,
    body: str,
    cc_list: list[str] | None = None,
    smtp_profile: str = "eran",
) -> bool:
    """
    שולח מייל הצעת מחיר ללקוח עם PDF מצורף מקובץ פיזי קבוע (current_quote_temp.pdf).
    מחזיר True אם השליחה הצליחה, False אחרת.
    """
    try:
        if smtp_profile == "tali":
            smtp_server = "smtp.gmail.com"
            smtp_port = 587
            sender_email = str(st.secrets["email_tali"]["sender_email"]).strip()
            password = str(st.secrets["email_tali"]["password"]).strip()
        elif smtp_profile == "eran":
            email_eran = dict(st.secrets.get("email_eran", {}) or {})
            smtp_server = (email_eran.get("smtp_server") or "").strip()
            smtp_port = int(email_eran.get("smtp_port", 465))
            sender_email = (email_eran.get("sender_email") or "").strip()
            password = (email_eran.get("password") or "").strip()
        else:
            st.error("פרופיל SMTP לא נתמך.")
            return False
        if not smtp_server or not sender_email or not password:
            st.error(
                "חסרים נתוני אימייל ב-secrets (בדוק מקטע [email_tali] או [email_eran] בהתאם לבחירה)"
            )
            return False
        to_email = (to_email or "").strip()
        if not to_email:
            st.error("חסרה כתובת אימייל ללקוח.")
            return False

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = to_email
        msg["Subject"] = subject
        if cc_list:
            msg["Cc"] = ", ".join(cc for cc in cc_list if cc and str(cc).strip())
        msg.attach(MIMEText(body, "plain"))

        file_to_attach = str(CURRENT_QUOTE_TEMP_PDF)
        if os.path.exists(file_to_attach):
            with open(file_to_attach, "rb") as f:
                pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")
                pdf_attachment.add_header(
                    "Content-Disposition", "attachment", filename="Quote_Studio84.pdf"
                )
                msg.attach(pdf_attachment)
        else:
            st.error(
                "הקובץ הפיזי לא נמצא. אנא לחץ שוב על 'שמור / עדכן הצעת מחיר' כדי לייצר אותו."
            )
            st.stop()

        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, password)
                server.send_message(msg)
        return True
    except smtplib.SMTPAuthenticationError as e:
        st.error(f"שגיאת אימות SMTP: {e}")
        return False
    except smtplib.SMTPException as e:
        st.error(f"שגיאת SMTP: {e}")
        return False
    except Exception as e:
        st.error(f"שגיאה בשליחת המייל: {e}")
        return False


def _ensure_sheet(sheet_name: str, columns: list[str]):
    """וידוא שקיים גיליון בגוגל שיטס. יוצר גיליון חדש אוטומטית אם חסר."""
    if spreadsheet is None:
        return None
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        return worksheet
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=sheet_name,
            rows=1000,
            cols=max(len(columns), 20),
        )
        worksheet.append_row(columns)
        return worksheet


@st.cache_data(ttl=600)
def read_contacts_sheet() -> pd.DataFrame:
    """משיכת נתוני אנשי קשר מגיליון contacts ל-DataFrame (כמו quotes / tasks)."""
    if spreadsheet is None:
        return pd.DataFrame(columns=CONTACTS_COLUMNS)
    _ensure_sheet('contacts', CONTACTS_COLUMNS)
    try:
        worksheet = spreadsheet.worksheet('contacts')
        df = _read_worksheet_safe(worksheet, CONTACTS_COLUMNS)
        return df.fillna('')
    except Exception as e:
        st.warning(f"שגיאה בקריאת contacts: {e}")
        return pd.DataFrame(columns=CONTACTS_COLUMNS)


def save_contacts(df: pd.DataFrame) -> None:
    """שומר מאגר אנשי קשר לגיליון contacts בגוגל שיטס."""
    if spreadsheet is None:
        st.error("אין חיבור לגוגל שיטס. לא ניתן לשמור.")
        return
    _ensure_sheet('contacts', CONTACTS_COLUMNS)
    try:
        worksheet = spreadsheet.worksheet('contacts')
        worksheet.clear()
        df_safe = df.reindex(columns=CONTACTS_COLUMNS, fill_value="").fillna("").astype(str)
        data = [df_safe.columns.values.tolist()] + df_safe.values.tolist()
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        time.sleep(1.5)
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת contacts: {e}")


@st.cache_data(ttl=600)
def read_project_contacts() -> list[dict]:
    """קריאת אנשי קשר לפרויקטים מגיליון project_contacts בגוגל שיטס."""
    if spreadsheet is None:
        return []
    _ensure_sheet('project_contacts', PROJECT_CONTACTS_COLUMNS)
    try:
        worksheet = spreadsheet.worksheet('project_contacts')
        df = _read_worksheet_safe(worksheet, PROJECT_CONTACTS_COLUMNS)
        return df.to_dict(orient='records')
    except Exception as e:
        st.warning(f"שגיאה בקריאת project_contacts: {e}")
        return []


def write_project_contacts(rows: list[dict]) -> None:
    """שמירת אנשי קשר לפרויקטים לגיליון project_contacts בגוגל שיטס."""
    if spreadsheet is None:
        st.error("אין חיבור לגוגל שיטס. לא ניתן לשמור.")
        return
    _ensure_sheet('project_contacts', PROJECT_CONTACTS_COLUMNS)
    try:
        worksheet = spreadsheet.worksheet('project_contacts')
        worksheet.clear()
        data = [PROJECT_CONTACTS_COLUMNS] + [[str(r.get(c, "") or "") for c in PROJECT_CONTACTS_COLUMNS] for r in rows]
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        time.sleep(1.5)
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת project_contacts: {e}")


def format_number(value: float) -> str:
    """Format numbers with thousand separators (e.g. 12,000)."""
    try:
        as_int = int(round(float(value or 0)))
    except (TypeError, ValueError):
        as_int = 0
    return f"{as_int:,.0f}"


def _payment_marked_yes(val) -> bool:
    """כן / True / yes — נחשב כחיובי לשדות גבייה."""
    if val is True:
        return True
    s = str(val or "").strip().lower()
    return s in ("כן", "yes", "true", "1")


def _normalize_payment_cell(val) -> str:
    """ערך אחיד לגיליון: 'כן' או 'לא'."""
    return "כן" if _payment_marked_yes(val) else "לא"


def _parse_currency_amount(val) -> float:
    """מחרוזת מספרית (₪, פסיקים) -> float; ריק או לא תקין -> 0."""
    if val is None:
        return 0.0
    s = str(val).strip().replace(",", "").replace(" ", "").replace("₪", "")
    if not s:
        return 0.0
    try:
        return max(0.0, float(s))
    except (TypeError, ValueError):
        return 0.0


def _format_advance_amount_storage(val) -> str:
    """שמירת סכום מקדמה לגיליון — מחרוזת עם ספרות עשרוניות או ריק."""
    a = _parse_currency_amount(val)
    return f"{a:.2f}" if a > 0 else ""


def _extract_total_from_quote_row(row: dict) -> float:
    """Extract total amount from a quote row. Tries: Total Price, סה"כ, Total, מחיר סופי, היקף כספי."""
    AMOUNT_COLS = ["Total Price", 'סה"כ', "Total", "מחיר סופי", "היקף כספי"]
    for col in AMOUNT_COLS:
        val = row.get(col)
        if val is None or (isinstance(val, str) and not str(val).strip()):
            continue
        try:
            s = str(val).strip().replace(",", "").replace(" ", "")
            if s:
                return float(s)
        except (TypeError, ValueError):
            continue
    return 0.0


def sanitize_filename_part(text: str) -> str:
    """Sanitize a string for safe use in filenames."""
    if not text:
        return "Client"
    # Allow Hebrew, English letters and digits, replace others with underscore
    safe = re.sub(r"[^A-Za-z0-9\u0590-\u05FF]+", "_", text).strip("_")
    return safe or "Client"


def load_config() -> dict:
    """Load configuration for auto-increment quote numbers."""
    config_path = CONFIG_PATH
    default_config = {"next_quote_number": 1}

    if not config_path.exists():
        return default_config

    try:
        with config_path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return default_config
        if "next_quote_number" not in data:
            data["next_quote_number"] = 1
        return data
    except Exception:
        return default_config


def save_config(config: dict) -> None:
    """Save configuration (e.g. next quote number) to config.json."""
    config_path = CONFIG_PATH
    try:
        with config_path.open("w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.warning(f"שגיאה בשמירת config.json: {e}")


# מיפוי סטטוסים ישנים לחדשים (למען תאימות לאחור)
_LEGACY_PROJECT_STATUS_MAP = {
    "Active": "בעבודה",
    "Completed": "הסתיים",
    "Done": "הסתיים",
}

# כותרות מאוחדות לקריאה אחת מגיליון projects (חוסך קריאת API כפולה)
_PROJECTS_SHEET_MERGED_COLUMNS = list(
    dict.fromkeys(list(PROJECTS_DB_COLUMNS) + list(PROJECTS_CSV_COLUMNS))
)


@st.cache_data(ttl=600)
def _read_projects_sheet_merged_df() -> pd.DataFrame:
    """קריאה אחת לגיליון projects — משותף ל-read_projects ו-read_projects_csv."""
    if spreadsheet is None:
        return pd.DataFrame()
    _ensure_sheet('projects', PROJECTS_DB_COLUMNS)
    worksheet = spreadsheet.worksheet('projects')
    df = _read_worksheet_safe(worksheet, _PROJECTS_SHEET_MERGED_COLUMNS)
    if not df.empty and hasattr(df.columns, 'str'):
        df.columns = df.columns.str.strip()
    return df


def read_projects() -> list[dict]:
    """קריאת פרויקטים מגיליון projects בגוגל שיטס (מוניטור, Task Board)."""
    if spreadsheet is None:
        err_msg = str(_sheets_init_error) if _sheets_init_error else "אין חיבור לגוגל שיטס"
        st.error(f"שגיאת קריאה: {err_msg}")
        return []
    try:
        df = _read_projects_sheet_merged_df()
        if df.empty:
            return []
        df = df.reindex(columns=PROJECTS_DB_COLUMNS, fill_value="")
        # ניקוי שורות רפאים - שורות ללא שם פרויקט
        if "Project Name" in df.columns and not df.empty:
            df = df[df["Project Name"].astype(str).str.strip() != ""]
        rows = df.to_dict(orient='records')
        for r in rows:
            status = (r.get("Status") or DEFAULT_PROJECT_STATUS).strip()
            if status in _LEGACY_PROJECT_STATUS_MAP:
                r["Status"] = _LEGACY_PROJECT_STATUS_MAP[status]
            elif status not in ALLOWED_PROJECT_STATUSES:
                r["Status"] = DEFAULT_PROJECT_STATUS
        return rows
    except Exception as e:
        st.error(f"שגיאת קריאה: {e}")
        return []


def write_projects(rows: list[dict], skip_rerun: bool = False) -> None:
    """שמירת פרויקטים לגיליון projects בגוגל שיטס."""
    if spreadsheet is None:
        err_msg = str(_sheets_init_error) if _sheets_init_error else "אין חיבור לגוגל שיטס"
        st.error(f"שגיאת שמירה: {err_msg}")
        return
    _ensure_sheet('projects', PROJECTS_DB_COLUMNS)
    try:
        worksheet = spreadsheet.worksheet('projects')
        worksheet.clear()
        data = [PROJECTS_DB_COLUMNS] + [[str(r.get(c, "") or "") for c in PROJECTS_DB_COLUMNS] for r in rows]
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        if not skip_rerun:
            time.sleep(1.5)
            st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת projects: {e}")


# מיפוי עמודות ישנות (אנגלית) לעמודות חדשות (עברית) - תאימות לאחור
_TASKS_LEGACY_COL_MAP = {
    "Project": "פרויקט", "Task Name": "שם משימה", "Assignee": "הוקצה ל",
    "Start Date": "תאריך התחלה", "Due Date": "תאריך יעד", "Status": "סטטוס",
    "Task Description": "תיאור המשימה",
    "Task ID": "מזהה משימה",
}


@st.cache_data(ttl=600)
def read_tasks() -> list[dict]:
    """קריאת משימות מגיליון tasks בגוגל שיטס (Task Board). כולל Retry ו-time.sleep למניעת 429."""
    if spreadsheet is None:
        return []
    _ensure_sheet('tasks', TASKS_LOG_COLUMNS)
    max_retries = 3
    last_error = None
    for attempt in range(max_retries):
        try:
            worksheet = spreadsheet.worksheet('tasks')
            df = _read_worksheet_safe(worksheet, TASKS_LOG_COLUMNS)
            if not df.empty and hasattr(df.columns, 'str'):
                df.columns = df.columns.str.strip()
            # תאימות לאחור: המרת עמודות אנגלית לעברית
            legacy_rename = {k: v for k, v in _TASKS_LEGACY_COL_MAP.items() if k in df.columns}
            if legacy_rename:
                df = df.rename(columns=legacy_rename)
            df = df.reindex(columns=TASKS_LOG_COLUMNS, fill_value="").fillna("")
            return df.to_dict(orient='records')
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                time.sleep(1)
    st.warning(f"שגיאה בקריאת tasks: {last_error}")
    return []


def write_tasks(rows: list[dict], skip_rerun: bool = False) -> None:
    """שמירת משימות לגיליון tasks בגוגל שיטס."""
    if spreadsheet is None:
        err_msg = str(_sheets_init_error) if _sheets_init_error else "אין חיבור לגוגל שיטס"
        st.error(f"שגיאת שמירה: {err_msg}")
        return
    _ensure_sheet('tasks', TASKS_LOG_COLUMNS)
    rows = _ensure_task_rows_have_ids(rows)
    try:
        worksheet = spreadsheet.worksheet('tasks')
        worksheet.clear()
        data = [TASKS_LOG_COLUMNS] + [[str(r.get(c, "") or "") for c in TASKS_LOG_COLUMNS] for r in rows]
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        if not skip_rerun:
            time.sleep(1.5)
            st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת tasks: {e}")


def _ensure_tasks_csv_schema() -> None:
    """וידוא שקיים גיליון tasks בגוגל שיטס. לא יוצר גיליון חדש - תציג שגיאה אם חסר."""
    if spreadsheet is None:
        return
    spreadsheet.worksheet('tasks')  # יקרוס ויציג שגיאה אם הגיליון חסר


@st.cache_data(ttl=600)
def read_daily_tasks() -> list[dict]:
    """קריאת כל המשימות היומיות מגוגל שיטס (גיליון tasks)."""
    if spreadsheet is None:
        return []
    _ensure_tasks_csv_schema()
    try:
        worksheet = spreadsheet.worksheet('tasks')  # שם גיליון: tasks (lowercase)
        df = _read_worksheet_safe(worksheet, DAILY_TASKS_COLUMNS)
        if not df.empty and hasattr(df.columns, 'str'):
            df.columns = df.columns.str.strip()
        return df.to_dict(orient='records')
    except Exception as e:
        st.error(f"שגיאה בשליפת נתונים (tasks): {e}")
        return []


def write_daily_tasks(rows: list[dict], skip_rerun: bool = False) -> None:
    """שמירת משימות יומיות לגוגל שיטס (גיליון tasks)."""
    if spreadsheet is None:
        st.error("אין חיבור לגוגל שיטס. לא ניתן לשמור.")
        return
    try:
        worksheet = spreadsheet.worksheet('tasks')
        worksheet.clear()
        data = [DAILY_TASKS_COLUMNS] + [[str(r.get(c, "") or "") for c in DAILY_TASKS_COLUMNS] for r in rows]
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        if not skip_rerun:
            time.sleep(1.5)
            st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת tasks: {e}")


def append_kickoff_tasks_to_csv(
    project_display: str,
    assigned_team: list[str],
    project_template: list[str],
    task_deadline: date,
    skip_rerun: bool = False,
) -> None:
    """יוצר שרשרת משימות ב-tasks.csv עבור כל עובד וכל שלב שנבחרו בהזנקה."""
    _ensure_tasks_csv_schema()
    existing = read_daily_tasks()
    deadline_str = task_deadline.strftime("%d/%m/%Y")
    for assignee in assigned_team:
        for stage_name in project_template:
            row = {
                "Task Name": (stage_name or "").strip(),
                "Project": (project_display or "").strip(),
                "Assignee": (assignee or "").strip(),
                "Date": deadline_str,
                "Status": "לביצוע",
                "Is Done": "0",
                "Flexible": "0",
            }
            existing.append(row)
    write_daily_tasks(existing, skip_rerun=skip_rerun)


def next_task_id(existing_rows: list[dict]) -> int:
    """Return the next numeric task ID based on existing rows."""
    max_id = 0
    for r in existing_rows or []:
        try:
            v = int((r.get("Task ID") or "0").strip() or "0")
        except ValueError:
            continue
        if v > max_id:
            max_id = v
    return max_id + 1


def next_project_id(existing_rows: list[dict]) -> int:
    """Return the next numeric project ID based on existing rows."""
    max_id = 0
    for r in existing_rows or []:
        try:
            v = int((r.get("Project ID") or "0").strip() or "0")
        except ValueError:
            continue
        if v > max_id:
            max_id = v
    return max_id + 1


def _safe_link(val: str) -> str:
    """Safe link string: never save 0, use empty string if missing or invalid."""
    if not val or not str(val).strip() or str(val).strip() == "0":
        return ""
    return str(val).strip()


def find_project_dropbox_links_for_client(client: str, project_name: str) -> tuple[str, str, str]:
    """קישורי דרופבוקס לפרויקט לפי לקוח ושם פרויקט (גיליון projects)."""
    c = (client or "").strip()
    p = (project_name or "").strip()
    if not c or not p:
        return ("", "", "")
    for r in read_projects():
        if (r.get("Client") or "").strip() != c:
            continue
        if (r.get("Project Name") or "").strip() != p:
            continue
        return (
            (r.get("Dropbox_Main") or "").strip(),
            (r.get("Dropbox_Upload") or "").strip(),
            (r.get("Dropbox_Deliverables") or "").strip(),
        )
    return ("", "", "")


def append_project_record(
    client: str,
    project_name: str,
    manager: str,
    team_members: list[str],
    status: str,
    start_date_str: str,
    budget_amount: str | float = "",
    dropbox_main: str = "",
    dropbox_upload: str = "",
    dropbox_deliverables: str = "",
    skip_rerun: bool = False,
) -> None:
    """Append a new project row to projects (Google Sheets)."""
    _ensure_sheet('projects', PROJECTS_DB_COLUMNS)
    existing_rows = read_projects()
    project_id = next_project_id(existing_rows)

    clean_status = status if status in ALLOWED_PROJECT_STATUSES else DEFAULT_PROJECT_STATUS

    amt_str = ""
    if budget_amount is not None and budget_amount != "":
        try:
            amt_str = str(int(round(float(budget_amount))))
        except (TypeError, ValueError):
            amt_str = ""

    row = {
        "Project ID": str(project_id),
        "Client": (client or "").strip(),
        "Project Name": (project_name or "").strip(),
        "Manager": (manager or "").strip(),
        "Team": ", ".join(team_members or []),
        "Status": clean_status,
        "Start Date": start_date_str,
        "Dropbox_Main": _safe_link(dropbox_main),
        "Dropbox_Upload": _safe_link(dropbox_upload),
        "Dropbox_Deliverables": _safe_link(dropbox_deliverables),
        'היקף כספי (₪)': amt_str,
    }

    existing_rows.append(row)
    write_projects(existing_rows, skip_rerun=skip_rerun)


def _ensure_projects_csv_schema() -> None:
    """וידוא שקיים גיליון projects בגוגל שיטס. לא יוצר גיליון חדש - תציג שגיאה אם חסר."""
    if spreadsheet is None:
        return
    spreadsheet.worksheet('projects')  # יקרוס ויציג שגיאה אם הגיליון חסר


def read_projects_csv() -> list[dict]:
    """קריאת כל הפרויקטים מגוגל שיטס (גיליון projects) — נגזר מ-_read_projects_sheet_merged_df."""
    if spreadsheet is None:
        return []
    _ensure_projects_csv_schema()
    try:
        df = _read_projects_sheet_merged_df()
        if df.empty:
            return []
        df = df.reindex(columns=PROJECTS_CSV_COLUMNS, fill_value="")
        rows = df.to_dict(orient='records')
        statuses_normalized = [s.strip().lower() for s in PROJECTS_CSV_STATUSES]
        for r in rows:
            status = (r.get("Status") or "Active").strip()
            if status.lower() not in statuses_normalized:
                status = "Active"
            else:
                status = PROJECTS_CSV_STATUSES[statuses_normalized.index(status.lower())]
            r["Status"] = status
        return rows
    except Exception as e:
        st.error(f"שגיאה בשליפת נתונים (projects): {e}")
        return []


def write_projects_csv(rows: list[dict], skip_rerun: bool = False) -> None:
    """שמירת פרויקטים פעילים לגוגל שיטס (גיליון projects)."""
    if spreadsheet is None:
        st.error("אין חיבור לגוגל שיטס. לא ניתן לשמור.")
        return
    try:
        worksheet = spreadsheet.worksheet('projects')
        worksheet.clear()
        data = [PROJECTS_CSV_COLUMNS] + [[str(r.get(c, "") or "") for c in PROJECTS_CSV_COLUMNS] for r in rows]
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        if not skip_rerun:
            time.sleep(1.5)
            st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת projects: {e}")


def next_projects_csv_id(existing_rows: list[dict]) -> int:
    """Return next ID for projects.csv."""
    max_id = 0
    for r in existing_rows or []:
        try:
            v = int((r.get("ID") or "0").strip() or "0")
        except ValueError:
            continue
        if v > max_id:
            max_id = v
    return max_id + 1


def append_to_projects_csv(
    client: str,
    project: str,
    deadline: str,
    team: str,
    budget_hours: str = "",
    budget_amount: str | float = "",
    project_contacts: str = "",
    skip_rerun: bool = False,
) -> None:
    """Add a new row to projects.csv (from Kickoff on signed proposal)."""
    _ensure_projects_csv_schema()
    existing = read_projects_csv()
    new_id = next_projects_csv_id(existing)
    amt_str = ""
    if budget_amount is not None and budget_amount != "":
        try:
            amt_str = str(int(round(float(budget_amount))))
        except (TypeError, ValueError):
            amt_str = ""
    row = {
        "ID": str(new_id),
        "Client": (client or "").strip(),
        "Project": (project or "").strip(),
        "Deadline": (deadline or "").strip(),
        "Team": (team or "").strip(),
        "Status": "Active",
        "Budget_Hours": (budget_hours or "").strip(),
        'היקף כספי (₪)': amt_str,
        "אנשי קשר מקושרים": (project_contacts or "").strip(),
    }
    existing.append(row)
    write_projects_csv(existing, skip_rerun=skip_rerun)


def _project_exists_in_projects_csv(client: str, project: str) -> bool:
    """Check if client+project already exists in projects.csv."""
    rows = read_projects_csv()
    c = (client or "").strip()
    p = (project or "").strip()
    for r in rows:
        if (r.get("Client") or "").strip() == c and (r.get("Project") or "").strip() == p:
            return True
    return False


def _project_exists_in_projects(client: str, project_name: str) -> bool:
    """Check if client+project already exists in projects (מוניטור וטבלת פרויקטים)."""
    rows = read_projects()
    c = (client or "").strip()
    p = (project_name or "").strip()
    for r in rows:
        if (r.get("Client") or "").strip() == c and (r.get("Project Name") or "").strip() == p:
            return True
    return False


def _ensure_project_active_in_projects(
    client: str,
    project_name: str,
    status: str = "בעבודה",
) -> bool:
    """Update project status in projects to active status. Returns True if updated."""
    rows = read_projects()
    c = (client or "").strip()
    p = (project_name or "").strip()
    for r in rows:
        if (r.get("Client") or "").strip() == c and (r.get("Project Name") or "").strip() == p:
            if (r.get("Status") or "").strip() != status:
                r["Status"] = status
                write_projects(rows, skip_rerun=True)
                return True
            return False
    return False


def _quote_csv_to_log_row(r: dict) -> dict:
    """Map a quote row from Google Sheets (QUOTES_CSV_COLUMNS) to QUOTES_LOG_COLUMNS format."""
    total = (r.get("Total Price") or "").strip()
    if not total:
        extracted = _extract_total_from_quote_row(r)
        total = f"{extracted:.2f}" if extracted else ""
    status = (r.get("Status") or DEFAULT_QUOTE_STATUS).strip()
    if status not in ALLOWED_QUOTE_STATUSES:
        status = DEFAULT_QUOTE_STATUS
    return {
        "Date": (r.get("Date") or "").strip(),
        "Client": (r.get("Client") or "").strip(),
        "Project": (r.get("Project") or "").strip(),
        "Version": (r.get("Version") or "V1").strip(),
        "Total Price": total,
        "Status": status,
        "File Path": (r.get("File Path") or "").strip(),
        "Signed File Path": (r.get("Signed File Path") or "").strip(),
        "Custom Item Desc": (r.get("custom_item_desc") or r.get("Custom Item Desc") or "").strip(),
        "Custom Item Price": (r.get("custom_item_price") or r.get("Custom Item Price") or "").strip(),
        "מקדמה שולמה": _normalize_payment_cell(r.get("מקדמה שולמה")),
        "סכום מקדמה": _format_advance_amount_storage(r.get("סכום מקדמה")),
        "גמר חשבון שולם": _normalize_payment_cell(r.get("גמר חשבון שולם")),
        "freelancer_cost": (r.get("freelancer_cost") or "").strip(),
        "misc_expenses": (r.get("misc_expenses") or "").strip(),
    }


def read_quotes_log() -> list[dict]:
    """קריאת הצעות מגיליון quotes — נגזר מ-read_quotes_csv() (מטמון) ללא שכבת מטמון כפולה."""
    rows = read_quotes_csv()
    return [_quote_csv_to_log_row(r) for r in rows]


def write_quotes_log(rows: list[dict]) -> None:
    """עדכון הצעות בגיליון quotes בגוגל שיטס - מיזוג שינויים (Status, File Path וכו') לרשומות הקיימות."""
    if spreadsheet is None:
        st.error("אין חיבור לגוגל שיטס. לא ניתן לשמור.")
        return
    current = read_quotes_csv()
    log_by_key = {(r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version")): r for r in rows}
    for i, q in enumerate(current):
        k = (q.get("Date"), q.get("Client"), q.get("Project"), q.get("Version"))
        log_row = log_by_key.get(k)
        if log_row:
            q["Status"] = (log_row.get("Status") or DEFAULT_QUOTE_STATUS).strip()
            q["File Path"] = (log_row.get("File Path") or "").strip()
            q["Signed File Path"] = (log_row.get("Signed File Path") or "").strip()
            q["Total Price"] = (log_row.get("Total Price") or "").strip()
            q["custom_item_desc"] = (log_row.get("Custom Item Desc") or "").strip()
            q["custom_item_price"] = (log_row.get("Custom Item Price") or "").strip()
            q["מקדמה שולמה"] = _normalize_payment_cell(log_row.get("מקדמה שולמה", q.get("מקדמה שולמה")))
            q["סכום מקדמה"] = _format_advance_amount_storage(
                log_row.get("סכום מקדמה", q.get("סכום מקדמה"))
            )
            q["גמר חשבון שולם"] = _normalize_payment_cell(log_row.get("גמר חשבון שולם", q.get("גמר חשבון שולם")))
            q["freelancer_cost"] = _format_advance_amount_storage(
                log_row.get("freelancer_cost", q.get("freelancer_cost"))
            )
            q["misc_expenses"] = _format_advance_amount_storage(
                log_row.get("misc_expenses", q.get("misc_expenses"))
            )
    # Remove quotes that were deleted (in rows we have the current set - if a quote is in current but not in rows, it was deleted)
    keys_in_rows = {(r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version")) for r in rows}
    current = [q for q in current if (q.get("Date"), q.get("Client"), q.get("Project"), q.get("Version")) in keys_in_rows]
    write_quotes_csv(current)


def parse_version_number(version_value: str) -> int:
    """
    Parse 'V2' / 'v10' / '2' -> 2/10/2. Returns 0 if unknown.
    """
    if not version_value:
        return 0
    m = re.search(r"(\d+)", str(version_value))
    return int(m.group(1)) if m else 0


def next_quote_version(client_name: str, project_name: str, existing_rows: list[dict]) -> str:
    client_key = (client_name or "").strip()
    project_key = (project_name or "").strip()

    if not client_key or not project_key:
        return "V1"

    max_v = 0
    for r in existing_rows or []:
        if (r.get("Client") or "").strip() == client_key and (r.get("Project") or "").strip() == project_key:
            max_v = max(max_v, parse_version_number(r.get("Version") or ""))
    return f"V{max_v + 1 if max_v > 0 else 1}"




def _quote_key(client: str, project: str, version: str) -> tuple[str, str, str]:
    """Unique key for a quote: (Client, Project, Version)."""
    return ((client or "").strip(), (project or "").strip(), (version or "V1").strip())


def _ensure_quotes_csv_schema() -> None:
    """וידוא שקיים גיליון quotes בגוגל שיטס. לא יוצר גיליון חדש - תציג שגיאה אם חסר."""
    if spreadsheet is None:
        return
    spreadsheet.worksheet('quotes')  # יקרוס ויציג שגיאה אם הגיליון חסר


@st.cache_data(ttl=600)
def read_quotes_csv() -> list[dict]:
    """קריאת נתוני טופס הצעות מחיר מגוגל שיטס (גיליון quotes)."""
    if spreadsheet is None:
        return []
    _ensure_quotes_csv_schema()
    try:
        worksheet = spreadsheet.worksheet('quotes')
        df = _read_worksheet_safe(worksheet, QUOTES_CSV_COLUMNS)
        result = []
        for _, row in df.iterrows():
            normalized = {}
            for c in QUOTES_CSV_COLUMNS:
                v = row.get(c)
                if v is None or (isinstance(v, str) and v.strip().lower() in ("nan", "none")):
                    normalized[c] = ""
                else:
                    s = str(v).strip()
                    normalized[c] = "" if s.lower() in ("nan", "none") else s
            for pay_col in ("מקדמה שולמה", "גמר חשבון שולם"):
                normalized[pay_col] = _normalize_payment_cell(normalized.get(pay_col))
            normalized["סכום מקדמה"] = _format_advance_amount_storage(normalized.get("סכום מקדמה"))
            normalized["freelancer_cost"] = _format_advance_amount_storage(normalized.get("freelancer_cost"))
            normalized["misc_expenses"] = _format_advance_amount_storage(normalized.get("misc_expenses"))
            result.append(normalized)
        return result
    except Exception as e:
        st.error(f"שגיאה בשליפת נתונים (quotes): {e}")
        return []


def write_quotes_csv(rows: list[dict], *, skip_rerun: bool = False) -> None:
    """Write full quote form data to Google Sheets (quotes tab). If skip_rerun, only updates sheet + clears cache."""
    sheet = _get_spreadsheet()
    if sheet is None:
        st.error("אין חיבור לגוגל שיטס. לא ניתן לשמור.")
        return
    _ensure_quotes_csv_schema()
    try:
        worksheet = spreadsheet.worksheet('quotes')
        worksheet.clear()
        data = [QUOTES_CSV_COLUMNS] + [[str(r.get(c, "") or "") for c in QUOTES_CSV_COLUMNS] for r in rows]
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        if skip_rerun:
            return
        st.success("הצעת המחיר נשמרה בהצלחה בגוגל שיטס!")
        time.sleep(1.5)
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת quotes: {e}")


@st.cache_data(ttl=600)
def get_quote_from_csv(client: str, project: str, version: str) -> dict | None:
    """Get full quote row from quotes (Google Sheets) by (Client, Project, Version)."""
    key = _quote_key(client, project, version)
    for r in read_quotes_csv():
        if _quote_key(r.get("Client", ""), r.get("Project", ""), r.get("Version", "")) == key:
            return r
    return None


def _quote_versions_equivalent_for_match(v_log: str, v_csv: str) -> bool:
    """Match Version between UI log row and raw CSV (V1 vs 1, empty vs V1)."""
    a = (v_log or "").strip() or "V1"
    b = (v_csv or "").strip() or "V1"
    if a == b:
        return True
    na, nb = parse_version_number(a), parse_version_number(b)
    return na > 0 and nb > 0 and na == nb


def get_quote_from_csv_by_log_row(log_row: dict) -> dict | None:
    """
    שליפת שורת הצעה מלאה לפי Date + Client + Project + Version (כמו במפתח ב-write_quotes_log).
    נדרש כשיש יותר משורה עם אותה שלישיית Client/Project/Version — get_quote_from_csv מחזירה רק את הראשונה.
    """
    td = (log_row.get("Date") or "").strip()
    tc = (log_row.get("Client") or "").strip()
    tp = (log_row.get("Project") or "").strip()
    tv = (log_row.get("Version") or "").strip() or "V1"
    for r in read_quotes_csv():
        if (r.get("Date") or "").strip() != td:
            continue
        if (r.get("Client") or "").strip() != tc:
            continue
        if (r.get("Project") or "").strip() != tp:
            continue
        if not _quote_versions_equivalent_for_match(tv, (r.get("Version") or "").strip()):
            continue
        return r
    return None


def append_quote_to_csv(row: dict) -> None:
    """Append a new quote row to quotes (Google Sheets)."""
    rows = read_quotes_csv()
    rows.append(row)
    write_quotes_csv(rows)


def update_quote_in_csv(
    client: str,
    project: str,
    version: str,
    updated_row: dict,
    *,
    skip_rerun: bool = False,
) -> bool:
    """Update existing quote in quotes (Google Sheets). Returns True if found and updated."""
    rows = read_quotes_csv()
    key = _quote_key(client, project, version)
    for i, r in enumerate(rows):
        if _quote_key(r.get("Client", ""), r.get("Project", ""), r.get("Version", "")) == key:
            rows[i] = {c: (updated_row.get(c) or "") for c in QUOTES_CSV_COLUMNS}
            write_quotes_csv(rows, skip_rerun=skip_rerun)
            return True
    return False


def _status_allows_convert_to_project(status: str) -> bool:
    """Signed / Approved / אושר / חתום — ללא תלות ברישיות."""
    s = (status or "").strip().lower()
    return s in ("signed", "approved", "אושר", "חתום")


def convert_quote_to_project(
    client: str,
    project: str,
    version: str,
    *,
    log_row: dict | None = None,
) -> tuple[bool, str, tuple[str, str, str], bool]:
    """
    המרת הצעת מחיר מאושרת לפרויקט פעיל.
    מבצע: משיכת נתונים, יצירת שורה ב-projects, תיקיות דרופבוקס (אם אפשר), עדכון סטטוס הצעה.
    מייל התנעה נשלח מהממשק בנפרד.
    מחזיר (הצלחה, הודעת שגיאה, קישורי דרופבוקס שנשמרו, dropbox_failed).
    dropbox_failed=True כשיצירת תיקיות/לינקים בדרופבוקס נכשלה (למשל טוקן פג תוקף).
    """
    try:
        # א. משיכת נתוני הצעת המחיר — לפי שורת ניהול מלאה (תאריך+מפתח) כדי לא לבלבל בין כפילויות Client/Project/Version
        if log_row is not None:
            full_row = get_quote_from_csv_by_log_row(log_row)
        else:
            full_row = get_quote_from_csv(client, project, version)
        if not full_row:
            return False, "ההצעה לא נמצאה במערכת.", ("", "", ""), True
        client_name = (full_row.get("Client") or "").strip()
        project_name = (full_row.get("Project") or "").strip()
        total_amount = _extract_total_from_quote_row(full_row)
        total_str = f"{total_amount:.2f}" if total_amount else ""
        if not client_name or not project_name:
            return False, "חסרים שם לקוח או שם פרויקט בהצעה.", ("", "", ""), True
        status = (full_row.get("Status") or "").strip()
        if status == "הומר לפרויקט":
            return False, "ההצעה כבר הומרה לפרויקט בעבר.", ("", "", ""), True
        if not _status_allows_convert_to_project(status):
            return (
                False,
                f"ניתן להמיר רק הצעות בסטטוס 'אושר' או 'חתום'. הסטטוס הנוכחי: {status}",
                ("", "", ""),
                True,
            )

        today_str = date.today().strftime("%d/%m/%Y")
        dropbox_project_id = f"{client_name}_{project_name}".replace(" ", "_")

        main_link, upload_link, deliverables_link = "", "", ""
        dropbox_failed = False
        try:
            result = create_studio_dropbox_structure(dropbox_project_id)
            if result is None:
                dropbox_failed = True
            else:
                main_link, upload_link, deliverables_link = result
        except Exception:
            dropbox_failed = True

        append_project_record(
            client=client_name,
            project_name=project_name,
            manager="",
            team_members=[],
            status="ממתין להתחלה",
            start_date_str=today_str,
            budget_amount=total_str,
            dropbox_main=main_link,
            dropbox_upload=upload_link,
            dropbox_deliverables=deliverables_link,
            skip_rerun=True,
        )

        # ה. עדכון סטטוס הצעת המחיר
        updated = {c: (full_row.get(c) or "") for c in QUOTES_CSV_COLUMNS}
        updated["Status"] = "הומר לפרויקט"
        if not update_quote_in_csv(
            full_row.get("Client", ""),
            full_row.get("Project", ""),
            full_row.get("Version", ""),
            updated,
        ):
            return False, "לא ניתן לעדכן את סטטוס ההצעה.", (main_link, upload_link, deliverables_link), dropbox_failed

        return True, "", (main_link, upload_link, deliverables_link), dropbox_failed
    except Exception as e:
        return False, str(e), ("", "", ""), True


def send_project_kickoff_email_eran(
    to_emails: list[str],
    project_name: str,
    client: str,
    deadline_str: str,
    main_link: str = "",
    upload_link: str = "",
    deliverables_link: str = "",
    brief_notes: str = "",
) -> bool:
    """
    מייל התנעת פרויקט לנמענים שנבחרו — אותן הגדרות SMTP כמו send_quote_email_via_smtp (פרופיל tali / email_tali).
    """
    try:
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = str(st.secrets["email_tali"]["sender_email"]).strip()
        password = str(st.secrets["email_tali"]["password"]).strip()
        if not smtp_server or not sender_email or not password:
            st.error("חסרים נתוני אימייל ב-secrets (מקטע [email_tali])")
            return False
        cleaned = [e.strip() for e in (to_emails or []) if e and str(e).strip()]
        if not cleaned:
            st.warning("אין כתובות נמען תקינות.")
            return False
        links_html = []
        if main_link and main_link.startswith("http"):
            links_html.append(f'<li><a href="{html.escape(main_link)}">📂 תיקייה ראשית</a></li>')
        if upload_link and upload_link.startswith("http"):
            links_html.append(f'<li><a href="{html.escape(upload_link)}">📥 בקשת חומרים (העלאה ללקוח)</a></li>')
        if deliverables_link and deliverables_link.startswith("http"):
            links_html.append(f'<li><a href="{html.escape(deliverables_link)}">📤 תיקיית תוצרים</a></li>')
        links_section = f"<ul>{''.join(links_html)}</ul>" if links_html else "<p>קישורי דרופבוקס לא זמינים.</p>"
        brief_block = ""
        if (brief_notes or "").strip():
            brief_block = (
                f'<h3>בריף / הערות</h3><p style="white-space:pre-wrap;">'
                f"{html.escape(brief_notes.strip())}</p>"
            )
        html_body = f"""
<!DOCTYPE html>
<html dir="rtl" lang="he">
<head><meta charset="utf-8"></head>
<body style="font-family: Arial, sans-serif; direction: rtl;">
<h2>🚀 התנעת פרויקט וחלוקת משימות</h2>
<table style="border-collapse: collapse;">
<tr><td style="padding:6px;"><strong>פרויקט:</strong></td><td>{html.escape(project_name)}</td></tr>
<tr><td style="padding:6px;"><strong>לקוח:</strong></td><td>{html.escape(client)}</td></tr>
<tr><td style="padding:6px;"><strong>דדליין:</strong></td><td>{html.escape(deadline_str)}</td></tr>
</table>
<h3>לינקים לתיקיות הדרופבוקס</h3>
{links_section}
{brief_block}
<p>בהצלחה!</p>
</body>
</html>
"""
        def _kickoff_plain_link(url: str) -> str:
            u = (url or "").strip()
            return u if u.startswith("http") else "(לא זמין)"

        brief_text = (brief_notes or "").strip()
        email_body = (
            "היי צוות, מצורף בריף וקובץ הצעת מחיר.\n"
            f"הערות למשימה: {brief_text if brief_text else '(ללא)'}\n\n"
            "קישורי דרופבוקס לעבודה:\n"
            f"📁 תיקייה ראשית: {_kickoff_plain_link(main_link)}\n"
            f"📥 בקשת חומרים: {_kickoff_plain_link(upload_link)}\n"
            f"📤 תוצרים: {_kickoff_plain_link(deliverables_link)}\n\n"
            "בהצלחה!"
        )
        msg = EmailMessage()
        msg["Subject"] = f"התנעת פרויקט: {project_name} - {client}"
        msg["From"] = sender_email
        msg["To"] = ", ".join(cleaned)
        msg.set_content(email_body)
        msg.add_alternative(html_body, subtype="html")

        def _kickoff_try_send(server) -> bool:
            try:
                server.send_message(msg)
            except Exception as e:
                st.error(f"שגיאה בשליחת המייל: {e}")
                return False
            st.success("המייל נשלח לצוות.")
            return True

        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, password)
                if not _kickoff_try_send(server):
                    return False
        else:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, password)
                if not _kickoff_try_send(server):
                    return False
        return True
    except smtplib.SMTPAuthenticationError as e:
        st.error(f"שגיאת אימות SMTP: {e}")
        return False
    except smtplib.SMTPException as e:
        st.error(f"שגיאת SMTP: {e}")
        return False
    except Exception as e:
        st.error(f"שגיאה בשליחת המייל: {e}")
        return False


def build_mailto_link(to: str, cc_list, subject: str, body: str) -> str:
    """
    Create a mailto: link with encoded subject & body and optional CC.

    Hebrew and other non-ASCII characters are safely encoded using urllib.parse.quote.
    """
    to = (to or "").strip()
    # בסיס ה-mailto עם כתובת הלקוח (אם קיימת)
    if to:
        mailto = f"mailto:{quote(to, safe='@.+-_')}"
    else:
        mailto = "mailto:"

    params = []

    # CC: רשימת כתובות מופרדת בפסיקים
    if cc_list:
        cc_clean = [addr.strip() for addr in cc_list if addr and addr.strip()]
        if cc_clean:
            cc_str = ",".join(cc_clean)
            params.append(f"cc={quote(cc_str, safe=',@.+-_')}")

    if subject:
        params.append(f"subject={quote(subject)}")
    if body:
        params.append(f"body={quote(body)}")

    if params:
        mailto += "?" + "&".join(params)

    return mailto


def build_gmail_link(to: str, cc_list, subject: str, body: str) -> str:
    """
    Create a Gmail compose URL with encoded subject, body and optional CC.

    Hebrew and other non-ASCII characters are safely encoded using urllib.parse.quote.
    """
    base_url = "https://mail.google.com/mail/?view=cm&fs=1&tf=1"

    params = []

    # To
    to = (to or "").strip()
    if to:
        params.append(f"to={quote(to, safe='@.+-_')}")

    # CC
    if cc_list:
        cc_clean = [addr.strip() for addr in cc_list if addr and addr.strip()]
        if cc_clean:
            cc_str = ",".join(cc_clean)
            params.append(f"cc={quote(cc_str, safe=',@.+-_')}")

    # Subject & body
    if subject:
        params.append(f"su={quote(subject)}")
    if body:
        params.append(f"body={quote(body)}")

    if params:
        return base_url + "&" + "&".join(params)
    return base_url


def ensure_project_folders_for_approved_quote(client_name: str, project_name: str) -> Path:
    """
    Ensure project folder exists at Projects/Client/Project with subfolders:
    Client_Material, Client_Materials and Studio_Material.
    """
    safe_client = sanitize_filename_part(client_name)
    safe_project = sanitize_filename_part(project_name)

    # יצירת תיקיות הפרויקט יחסית לתיקיית הבסיס של האפליקציה (BASE_DIR),
    # כך שהנתיב יעבוד באופן זהה בכל מחשב.
    project_path = PROJECTS_ROOT / safe_client / safe_project
    (project_path / "Client_Material").mkdir(parents=True, exist_ok=True)
    (project_path / "Client_Materials").mkdir(parents=True, exist_ok=True)
    (project_path / "Studio_Material").mkdir(parents=True, exist_ok=True)
    return project_path

def _parse_edit_date(date_str: str) -> date | None:
    """Parse date string from CSV (dd/mm/yyyy or yyyy-mm-dd) to date object."""
    if not date_str or not str(date_str).strip():
        return None
    s = str(date_str).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _show_import_past_quote_form() -> None:
    """טופס ייבוא הצעת עבר - העלאת PDF ושמירה למערכת."""
    st.header("ייבוא הצעת עבר")
    st.caption("הזן פרטי הצעה קיימת והעלה את קובץ ה-PDF. ההצעה תישמר כחתומה (Signed).")

    quote_date = st.date_input("תאריך ההצעה המקורית", key="import_quote_date")
    client = st.text_input("שם הלקוח", key="import_quote_client")
    project = st.text_input("שם הפרויקט", key="import_quote_project")
    file = st.file_uploader("העלה את קובץ ה-PDF של ההצעה", type=["pdf"], key="import_quote_file")

    if st.button("שמור וייבא למערכת", key="import_quote_btn"):
        if not client or not project:
            st.error("יש למלא את שם הלקוח ואת שם הפרויקט.")
        elif file is None:
            st.error("יש להעלות קובץ PDF.")
        else:
            try:
                year = quote_date.year
                month = quote_date.month
                display_date = quote_date.strftime("%d/%m/%Y")

                # נתיב: Quotes/{Year}/{Month}/Signed
                target_dir = _status_to_quotes_folder("Signed", year, month)
                target_dir.mkdir(parents=True, exist_ok=True)

                existing_rows = read_quotes_log()
                version = next_quote_version(client, project, existing_rows)

                safe_client = sanitize_filename_part(client)
                safe_project = sanitize_filename_part(project)
                filename = f"Signed_{safe_client}_{safe_project}.pdf"
                target_path = target_dir / filename

                with target_path.open("wb") as f:
                    f.write(file.getvalue())

                # הוספת שורה לגיליון quotes בגוגל שיטס
                new_row = {c: "" for c in QUOTES_CSV_COLUMNS}
                new_row.update({
                    "Date": display_date,
                    "Client": (client or "").strip(),
                    "Project": (project or "").strip(),
                    "Version": version,
                    "Total Price": "0.00",
                    "Status": "Signed",
                    "File Path": "",
                    "Signed File Path": str(target_path.resolve()),
                    "מקדמה שולמה": "לא",
                    "סכום מקדמה": "",
                    "גמר חשבון שולם": "לא",
                })
                append_quote_to_csv(new_row)

                st.success(
                    "ההצעה יובאה בהצלחה! ניתן כעת להזניק את הפרויקט בלשונית ניהול הצעות."
                )
            except Exception as e:
                st.error(f"שגיאה בייבוא: {e}")


def show_quote_page() -> None:
    st.title("הפקת הצעת מחיר לאדריכלות")

    action_type = st.radio(
        "בחר סוג פעולה",
        ["📝 יצירת מסמך חדש", "📥 ייבוא הצעת עבר (PDF קיים)"],
        key="quote_action_type",
    )

    if action_type == "📥 ייבוא הצעת עבר (PDF קיים)":
        _show_import_past_quote_form()
        return

    # --- בחירת מצב עבודה: חדש / עריכה ---
    quotes_log_rows = read_quotes_log()
    quote_mode_options = ["צור הצעת מחיר חדשה"]
    for r in quotes_log_rows:
        client_val = (r.get("Client") or "").strip()
        project_val = (r.get("Project") or "").strip()
        version_val = (r.get("Version") or "").strip() or "V1"
        date_val = (r.get("Date") or "").strip()
        quote_mode_options.append(f"{client_val} | {project_val} | גרסה {version_val} | {date_val}")
    selected_mode = st.selectbox(
        "בחר פעולה: הצעה חדשה / עריכת הצעה קיימת",
        quote_mode_options,
        key="quote_mode_select",
    )
    is_edit_mode = selected_mode != "צור הצעת מחיר חדשה"
    edit_quote_row: dict | None = None
    if is_edit_mode and selected_mode in quote_mode_options:
        idx = quote_mode_options.index(selected_mode)
        if idx > 0 and idx <= len(quotes_log_rows):
            edit_quote_row = quotes_log_rows[idx - 1]

    # --- טעינת נתוני עריכה (Pre-fill) ---
    prefill: dict = {}
    if edit_quote_row:
        client_key = (edit_quote_row.get("Client") or "").strip()
        project_key = (edit_quote_row.get("Project") or "").strip()
        version_key = (edit_quote_row.get("Version") or "").strip() or "V1"
        full_row = get_quote_from_csv(client_key, project_key, version_key)
        if full_row:
            prefill = full_row
        else:
            prefill = {
                "Date": (edit_quote_row.get("Date") or "").strip(),
                "Client": client_key,
                "Project": project_key,
                "Version": version_key,
                "custom_item_desc": (edit_quote_row.get("Custom Item Desc") or "").strip(),
                "custom_item_price": (edit_quote_row.get("Custom Item Price") or "0").strip(),
            }

    # --- מצב יצירת מסמך חדש ---
    today = date.today()
    date_str_from_prefill = (prefill.get("Date") or "").strip()
    parsed_edit_date = _parse_edit_date(date_str_from_prefill) if date_str_from_prefill else None
    date_to_use = parsed_edit_date if parsed_edit_date else today
    display_date = date_to_use.strftime("%d/%m/%Y")
    file_date = date_to_use.strftime("%Y-%m-%d")

    # --- טעינת קונפיג למספור אוטומטי ---
    config = load_config()
    next_quote_number = str(config.get("next_quote_number", 1))

    # --- הודעת עריכה ---
    if is_edit_mode and edit_quote_row:
        edit_total = (edit_quote_row.get("Total Price") or "").strip()
        msg = "✏️ **מצב עריכה:** נתונים נטענו מההצעה שנבחרה."
        if edit_total:
            msg += f" סה\"כ הצעה מקורית: {edit_total} ₪"
        st.info(msg)

    def _prefill_str(key: str, default: str = "") -> str:
        """Safe string from prefill; handles NaN and empty."""
        v = prefill.get(key, default)
        if v is None or (isinstance(v, float) and v != v):
            return default or ""
        s = str(v).strip()
        if s.lower() in ("nan", "none", "inf", "-inf"):
            return default or ""
        return s or (default or "")

    def _prefill_bool(key: str) -> bool:
        v = prefill.get(key, "")
        return str(v).strip().lower() in ("true", "1", "yes")

    def _prefill_float(key: str, default: float = 0.0) -> float:
        try:
            v = prefill.get(key, default) or default
            if isinstance(v, str) and str(v).strip().lower() in ("nan", "inf", "-inf"):
                return default
            f = float(v)
            return default if (f != f) else f  # NaN check
        except (TypeError, ValueError):
            return default

    def _prefill_int(key: str, default: int = 0) -> int:
        try:
            v = prefill.get(key, default) or default
            if isinstance(v, str) and str(v).strip().lower() in ("nan", "inf", "-inf"):
                return default
            return int(float(v))
        except (TypeError, ValueError):
            return default

    # --- עדכון session_state כשעוברים למצב עריכה עם prefill (כדי שהשדות ייטענו) ---
    _prev_selection = st.session_state.get("_quote_edit_selection", "")
    if prefill and selected_mode != _prev_selection:
        st.session_state["_quote_edit_selection"] = selected_mode
        st.session_state["client_name"] = _prefill_str("Client")
        st.session_state["contact_person"] = _prefill_str("Contact Person")
        st.session_state["client_email"] = _prefill_str("Client Email")
        st.session_state["project_name"] = _prefill_str("Project")
        st.session_state["quote_subject"] = _prefill_str("Quote Subject")
        st.session_state["quote_number"] = _prefill_str("Quote Number") or next_quote_number
        st.session_state["show_exterior"] = _prefill_bool("show_exterior")
        st.session_state["show_interior"] = _prefill_bool("show_interior")
        st.session_state["show_drone"] = _prefill_bool("show_drone")
        st.session_state["show_video"] = _prefill_bool("show_video")
        st.session_state["show_shots"] = _prefill_bool("show_shots")
        for key, val in [
            ("scope_of_work", _prefill_str("scope_of_work")),
            ("work_process", _prefill_str("work_process")),
            ("delivery_time", _prefill_str("delivery_time")),
            ("video_terms", _prefill_str("video_terms")),
        ]:
            if val:
                st.session_state[key] = val
        st.session_state["price_exterior"] = _prefill_float("price_exterior")
        st.session_state["base_views_count"] = _prefill_int("base_views_count", 5)
        st.session_state["price_ext_extra"] = _prefill_float("price_ext_extra")
        st.session_state["include_price_exterior"] = _prefill_bool("include_price_exterior") if prefill.get("include_price_exterior") != "" else True
        st.session_state["include_price_ext_extra"] = _prefill_bool("include_price_ext_extra") if prefill.get("include_price_ext_extra") != "" else True
        st.session_state["price_drone"] = _prefill_float("price_drone")
        st.session_state["include_price_drone"] = _prefill_bool("include_price_drone") if prefill.get("include_price_drone") != "" else True
        st.session_state["price_shot_unit"] = _prefill_float("price_shot_unit")
        st.session_state["shots_count"] = _prefill_int("shots_count")
        st.session_state["include_total_shots_price"] = _prefill_bool("include_total_shots_price") if prefill.get("include_total_shots_price") != "" else True
        st.session_state["price_interior"] = _prefill_float("price_interior")
        st.session_state["price_int_extra"] = _prefill_float("price_int_extra")
        st.session_state["price_int_space"] = _prefill_float("price_int_space")
        st.session_state["include_price_interior"] = _prefill_bool("include_price_interior") if prefill.get("include_price_interior") != "" else True
        st.session_state["include_price_int_extra"] = _prefill_bool("include_price_int_extra") if prefill.get("include_price_int_extra") != "" else True
        st.session_state["include_price_int_space"] = _prefill_bool("include_price_int_space") if prefill.get("include_price_int_space") != "" else True
        st.session_state["price_video"] = _prefill_float("price_video")
        st.session_state["include_price_video"] = _prefill_bool("include_price_video") if prefill.get("include_price_video") != "" else True
        st.session_state["model_update_val"] = _prefill_float("model_update_val")
        st.session_state["view_update_val"] = _prefill_float("view_update_val")
        st.session_state["extra_view_val"] = _prefill_float("extra_view_val")
        st.session_state["custom_item_desc"] = _prefill_str("custom_item_desc")
        _custom_price = _prefill_float("custom_item_price")
        st.session_state["custom_item_price"] = max(0, int(round(_custom_price)))
    elif not is_edit_mode:
        st.session_state["_quote_edit_selection"] = ""

    # --- בחירת סקופ ההצעה (Checkboxes) ---
    st.header("בחירת מרכיבי ההצעה")
    cb1, cb2, cb3, cb4, cb5 = st.columns(5)
    with cb1:
        show_exterior = st.checkbox("הדמיות חוץ", value=_prefill_bool("show_exterior"), key="show_exterior")
    with cb2:
        show_interior = st.checkbox("הדמיות פנים", value=_prefill_bool("show_interior"), key="show_interior")
    with cb3:
        show_drone = st.checkbox("צילום רחפן", value=_prefill_bool("show_drone"), key="show_drone")
    with cb4:
        show_video = st.checkbox("סרטון תדמית", value=_prefill_bool("show_video"), key="show_video")
    with cb5:
        show_shots = st.checkbox("הפקת שוטים בלבד", value=_prefill_bool("show_shots"), key="show_shots")

    # --- פרטי פרויקט (תמיד מוצגים) ---
    st.header("פרטי הפרויקט")
    default_client = prefill.get("Client", "") or st.session_state.get("edit_client_name", "")
    default_project = prefill.get("Project", "") or st.session_state.get("edit_project_name", "")
    default_contact = prefill.get("Contact Person", "") or st.session_state.get("edit_contact_person", "")
    default_email = prefill.get("Client Email", "") or st.session_state.get("edit_client_email", "")
    default_subject = prefill.get("Quote Subject", "") or st.session_state.get("edit_quote_subject", "")
    default_quote_num = prefill.get("Quote Number", "") or st.session_state.get("edit_quote_number") or next_quote_number
    client_name = st.text_input("שם הלקוח (client_name)", value=default_client, key="client_name")
    contact_person = st.text_input("איש קשר (contact_person)", value=default_contact, key="contact_person")
    client_email = st.text_input("אימייל לקוח (client_email)", value=default_email, key="client_email")
    project_name = st.text_input("שם הפרויקט (project_name)", value=default_project, key="project_name")
    quote_subject = st.text_input("נושא ההצעה (quote_subject)", value=default_subject, key="quote_subject")
    quote_number = st.text_input(
        "מספר הצעה (quote_number)", value=default_quote_num, key="quote_number"
    )

    st.markdown(f"**תאריך (date):** {display_date}")

    # --- טקסטים לעריכה (תמיד מוצגים) ---
    st.header("טקסטים לעריכה")

    scope_of_work_default = (
        "מידול הפרויקט הכולל בניית מודל תלת־ממדי מלא של המבנה, "
        "התאמת חומרים, תאורה וסביבה, על בסיס תכניות האדריכל והחומרים "
        "שיסופקו על ידי הלקוח. ההצעה כוללת הפקת הדמיות סטילס ברזולוציה גבוהה "
        "לשימושי שיווק, מצגות והגשות."
    )
    scope_of_work = st.text_area(
        "תיאור העבודה (scope_of_work)",
        value=prefill.get("scope_of_work", "") or scope_of_work_default,
        height=180,
        key="scope_of_work",
    )

    work_process_default = (
        "המודל יבנה על פי תכניות אדריכל, מידות וחומרים שימסרו על ידי הלקוח. "
        "בשלב הראשון יישלחו סקיצות ראשוניות לאישור זוויות ותאורה. "
        "לאחר קבלת הערות מהלקוח יבוצעו תיקונים נדרשים, ולאחר מכן יוצגו "
        "הדמיות סופיות ברזולוציה מלאה. התהליך יכלול עד שני סבבי תיקונים "
        "במסגרת ההצעה."
    )
    work_process = st.text_area(
        "תהליך העבודה (work_process)",
        value=prefill.get("work_process", "") or work_process_default,
        height=180,
        key="work_process",
    )

    delivery_time_default = (
        "הזמן הדרוש הוא 14 ימי עסקים ממועד קבלת כל החומרים הנדרשים "
        "ומהתשלום המקדמי, בכפוף למענה מהיר על סקיצות ותיקונים ביניים. "
        "שינויים מהותיים בהיקף העבודה או בתכניות עלולים להשפיע על לוחות הזמנים."
    )
    delivery_time = st.text_area(
        "לוחות זמנים (delivery_time)",
        value=prefill.get("delivery_time", "") or delivery_time_default,
        height=160,
        key="delivery_time",
    )

    # תנאי טקסט לסרטון - יוצגו ויישלחו רק אם נבחר סרטון
    video_terms_default = (
        "הפקת סרטון תדמית כולל תכנון תסריט, יום צילום בשטח, עריכה, הנפשות בסיסיות "
        "ותוספת מוזיקה מותרת לשימוש. שינוי מהותי בתסריט או בתוכן לאחר תחילת העבודה "
        "עלול לגרור עלויות נוספות וארכה בלוחות הזמנים."
    )
    video_terms = ""
    if show_video:
        video_terms = st.text_area(
            "תנאים לסרטון (video_terms)",
            value=prefill.get("video_terms", "") or video_terms_default,
            height=160,
            key="video_terms",
        )

    # --- תמחור (שדות דינמיים בהתאם לבחירה) ---
    st.header("תמחור")

    # אתחול משתני מחירים
    price_exterior = 0.0
    base_views_count = 5
    price_ext_extra = 0.0
    price_interior = 0.0
    price_int_extra = 0.0
    price_int_space = 0.0
    price_drone = 0.0
    price_video = 0.0
    price_shot_unit = 0.0
    price_shots_total = 0.0
    shots_count = 0

    # אתחול משתני בחירה לחישוב סה"כ (לכל שדה מחיר בנפרד)
    include_price_exterior = True
    include_price_ext_extra = True
    include_price_interior = True
    include_price_int_extra = True
    include_price_int_space = True
    include_price_drone = True
    include_price_video = True
    include_total_shots_price = True

    col1, col2 = st.columns(2)

    with col1:
        if show_exterior:
            st.subheader("הדמיות חוץ")

            # מחיר בסיס חוץ
            ext_base_price_col, ext_base_check_col = st.columns([3, 1])
            with ext_base_price_col:
                price_exterior = st.number_input(
                    "מחיר בסיס חוץ (price_exterior)",
                    min_value=0.0,
                    step=1000.0,
                    value=_prefill_float("price_exterior"),
                    key="price_exterior",
                )
                base_views_count = st.number_input(
                    "כמות מבטים הכלולים במחיר הבסיס",
                    min_value=0,
                    step=1,
                    value=_prefill_int("base_views_count", 5),
                    key="base_views_count",
                )
            with ext_base_check_col:
                include_price_exterior = st.checkbox(
                    "כלול בסה\"כ",
                    value=_prefill_bool("include_price_exterior") if "include_price_exterior" in prefill else True,
                    key="include_price_exterior",
                )

            # תוספת למבט חוץ
            ext_extra_price_col, ext_extra_check_col = st.columns([3, 1])
            with ext_extra_price_col:
                price_ext_extra = st.number_input(
                    "תוספת למבט חוץ (price_ext_extra)",
                    min_value=0.0,
                    step=1000.0,
                    value=_prefill_float("price_ext_extra"),
                    key="price_ext_extra",
                )
            with ext_extra_check_col:
                include_price_ext_extra = st.checkbox(
                    "כלול בסה\"כ",
                    value=_prefill_bool("include_price_ext_extra") if "include_price_ext_extra" in prefill else True,
                    key="include_price_ext_extra",
                )

        if show_drone:
            st.subheader("צילום רחפן")
            drone_price_col, drone_check_col = st.columns([3, 1])
            with drone_price_col:
                price_drone = st.number_input(
                    "מחיר רחפן (price_drone)",
                    min_value=0.0,
                    step=500.0,
                    value=_prefill_float("price_drone"),
                    key="price_drone",
                )
            with drone_check_col:
                include_price_drone = st.checkbox(
                    "כלול בסה\"כ",
                    value=_prefill_bool("include_price_drone") if "include_price_drone" in prefill else True,
                    key="include_price_drone",
                )

        if show_shots:
            st.subheader("הפקת שוטים בלבד")

            # הזנת מחיר לשוט וכמות שוטים
            shots_inputs_col, _ = st.columns([3, 1])
            with shots_inputs_col:
                price_shot_unit = st.number_input(
                    "מחיר לשוט (price_shot_unit)",
                    min_value=0.0,
                    step=100.0,
                    value=_prefill_float("price_shot_unit"),
                    key="price_shot_unit",
                )
                shots_count = st.number_input(
                    "כמות שוטים (shots_count)",
                    min_value=0,
                    step=1,
                    value=_prefill_int("shots_count"),
                    key="shots_count",
                )

            # חישוב סה\"כ שוטים והצגת תיבה לשליטה על הכללה בסה\"כ
            price_shots_total = price_shot_unit * shots_count
            shots_total_col, shots_check_col = st.columns([3, 1])
            with shots_total_col:
                st.markdown(
                    f"סה\"כ שוטים (total_shots_price): {format_number(price_shots_total)} ₪"
                )
            with shots_check_col:
                include_total_shots_price = st.checkbox(
                    "כלול בסה\"כ",
                    value=_prefill_bool("include_total_shots_price") if "include_total_shots_price" in prefill else True,
                    key="include_total_shots_price",
                )

    with col2:
        if show_interior:
            st.subheader("הדמיות פנים")

            # מחיר בסיס פנים
            int_base_price_col, int_base_check_col = st.columns([3, 1])
            with int_base_price_col:
                price_interior = st.number_input(
                    "מחיר בסיס פנים (price_interior)",
                    min_value=0.0,
                    step=1000.0,
                    value=_prefill_float("price_interior"),
                    key="price_interior",
                )
            with int_base_check_col:
                include_price_interior = st.checkbox(
                    "כלול בסה\"כ",
                    value=_prefill_bool("include_price_interior") if "include_price_interior" in prefill else True,
                    key="include_price_interior",
                )

            # תוספת למבט פנים
            int_extra_price_col, int_extra_check_col = st.columns([3, 1])
            with int_extra_price_col:
                price_int_extra = st.number_input(
                    "תוספת למבט פנים (price_int_extra)",
                    min_value=0.0,
                    step=1000.0,
                    value=_prefill_float("price_int_extra"),
                    key="price_int_extra",
                )
            with int_extra_check_col:
                include_price_int_extra = st.checkbox(
                    "כלול בסה\"כ",
                    value=_prefill_bool("include_price_int_extra") if "include_price_int_extra" in prefill else True,
                    key="include_price_int_extra",
                )

            # תוספת לחלל
            int_space_price_col, int_space_check_col = st.columns([3, 1])
            with int_space_price_col:
                price_int_space = st.number_input(
                    "תוספת לחלל (price_int_space)",
                    min_value=0.0,
                    step=1000.0,
                    value=_prefill_float("price_int_space"),
                    key="price_int_space",
                )
            with int_space_check_col:
                include_price_int_space = st.checkbox(
                    "כלול בסה\"כ",
                    value=_prefill_bool("include_price_int_space") if "include_price_int_space" in prefill else True,
                    key="include_price_int_space",
                )

        if show_video:
            st.subheader("סרטון תדמית")
            video_price_col, video_check_col = st.columns([3, 1])
            with video_price_col:
                price_video = st.number_input(
                    "מחיר סרטון (price_video)",
                    min_value=0.0,
                    step=1000.0,
                    value=_prefill_float("price_video"),
                    key="price_video",
                )
            with video_check_col:
                include_price_video = st.checkbox(
                    "כלול בסה\"כ",
                    value=_prefill_bool("include_price_video") if "include_price_video" in prefill else True,
                    key="include_price_video",
                )

    # מחירון שינויים ותוספות (אופציונלי)
    with st.expander("💰 תמחור שינויים ותוספות (אופציונלי)"):
        model_update_val = st.number_input(
            "מחיר לעדכון מודל",
            min_value=0.0,
            step=50.0,
            value=_prefill_float("model_update_val"),
            key="model_update_val",
        )
        view_update_val = st.number_input(
            "מחיר לעדכון מבט קיים",
            min_value=0.0,
            step=50.0,
            value=_prefill_float("view_update_val"),
            key="view_update_val",
        )
        extra_view_val = st.number_input(
            "מחיר לתוספת מבט חדש",
            min_value=0.0,
            step=50.0,
            value=_prefill_float("extra_view_val"),
            key="extra_view_val",
        )

    # תנאי שורת השוטים
    if show_shots and shots_count > 0 and price_shot_unit > 0:
        shots_line = (
            f"{int(shots_count)} שוטים × {format_number(price_shot_unit)} ₪ לשוט = "
            f"{format_number(price_shots_total)} ₪"
        )
    else:
        shots_line = ""

    custom_item_desc = st.text_input(
        'תיאור שירות נוסף / אחר (אופציונלי)',
        value=prefill.get("custom_item_desc", ""),
        key="custom_item_desc",
    )
    custom_item_price = st.number_input(
        'תמחור לשירות הנוסף (₪)',
        min_value=0,
        value=_prefill_int("custom_item_price"),
        step=100,
    )

    # --- חישובים ---
    total_price = 0.0
    if show_exterior and include_price_exterior:
        total_price += price_exterior
    if show_exterior and include_price_ext_extra:
        total_price += price_ext_extra
    if show_interior and include_price_interior:
        total_price += price_interior
    if show_interior and include_price_int_extra:
        total_price += price_int_extra
    if show_interior and include_price_int_space:
        total_price += price_int_space
    if show_drone and include_price_drone:
        total_price += price_drone
    if show_video and include_price_video:
        total_price += price_video
    if show_shots and include_total_shots_price:
        total_price += price_shots_total

    total_price += custom_item_price

    vat = total_price * 0.17
    total_inc_vat = total_price + vat

    st.subheader("סיכום נתונים לבדיקה")

    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("סה\"כ לפני מע\"מ", format_number(total_price))
    with m2:
        st.metric("מע\"מ 17%", format_number(vat))
    with m3:
        st.metric("סה\"כ לתשלום (כולל מע\"מ)", format_number(total_inc_vat))

    st.caption("אנא בדקו שהסכומים והמע\"מ נכונים לפני יצירת המסמך.")

    # --- יצירת הקובץ ---
    st.divider()
    st.subheader("יצירת מסמך הצעת מחיר")
    create_button = st.button("שמור / עדכן הצעת מחיר")

    if create_button:
        if not client_name or not project_name or not quote_number:
            st.warning("נא למלא לפחות שם לקוח, שם פרויקט ומספר הצעה לפני יצירת הקובץ.")
            return

        template_path = TEMPLATE_PATH
        if not template_path.exists():
            st.error("הקובץ 'quote_template.docx' לא נמצא בתיקייה הנוכחית.")
            return

        # --- CRM Versioning: חדש = גרסה חדשה, עריכה = גרסה קיימת ---
        existing_rows = read_quotes_log()
        if is_edit_mode and edit_quote_row:
            quote_version = (edit_quote_row.get("Version") or "").strip() or "V1"
            orig_client = (edit_quote_row.get("Client") or "").strip()
            orig_project = (edit_quote_row.get("Project") or "").strip()
            orig_version = quote_version
        else:
            quote_version = next_quote_version(client_name, project_name, existing_rows)
            orig_client = orig_project = orig_version = None

        # תנאי טקסט הסרטון לפי בחירה
        video_terms_context = video_terms if show_video else ""

        # בדיקה אם יש עלויות נוספות (להסתרת סעיף תמחור שינויים ותוספות בוורד)
        has_extra_costs = model_update_val > 0 or view_update_val > 0 or extra_view_val > 0

        # שירות נוסף - להצגה לפני סה"כ לתשלום (בסגנון הנקודות כמו שאר השירותים)
        custom_item_line = ""
        if custom_item_price > 0 or (custom_item_desc and str(custom_item_desc).strip()):
            custom_item_line = f"{custom_item_desc or ''} .................................................. {format_number(custom_item_price)}\n\n"

        context = {
            # פרטים כלליים
            "client_name": client_name,
            "contact_person": contact_person,
            "project_name": project_name,
            "quote_subject": quote_subject,
            "quote_number": quote_number,
            "date": display_date,
            # בוליאנים לסקופ
            "show_exterior": show_exterior,
            "show_interior": show_interior,
            "show_drone": show_drone,
            "show_video": show_video,
            "show_shots": show_shots,
            # טקסטים
            "scope_of_work": scope_of_work,
            "work_process": work_process,
            "delivery_time": delivery_time,
            "video_terms": video_terms_context,
            "shots_line": shots_line,
            # מחירים בפורמט עם פסיקים
            "price_exterior": format_number(price_exterior),
            "base_views": base_views_count,
            "price_ext_extra": format_number(price_ext_extra),
            "price_drone": format_number(price_drone),
            "price_interior": format_number(price_interior),
            "price_int_extra": format_number(price_int_extra),
            "price_int_space": format_number(price_int_space),
            "price_video": format_number(price_video),
            "price_shot_unit": format_number(price_shot_unit),
            "price_shots_total": format_number(price_shots_total),
            "total_shots_price": format_number(price_shots_total),
            "shots_count": shots_count,
            "total_price": format_number(total_price),
            "vat": format_number(vat),
            "total_inc_vat": format_number(total_inc_vat),
            # מחירון שינויים ותוספות
            "model_update_price": f"{model_update_val:,.0f} ₪",
            "view_update_price": f"{view_update_val:,.0f} ₪",
            "extra_view_price": f"{extra_view_val:,.0f} ₪",
            "has_extra_costs": has_extra_costs,
            "custom_item_line": custom_item_line,
        }

        try:
            doc = DocxTemplate(str(template_path))
            doc.render(context)

            safe_client = sanitize_filename_part(client_name)
            safe_project = sanitize_filename_part(project_name)
            filename_base = f"Quote_{safe_client}_{safe_project}_{quote_version}_{file_date}"
            filename_docx = f"{filename_base}.docx"
            filename_pdf = f"{filename_base}.pdf"

            # שמירה במבנה היררכי: Quotes/{Year}/{Month}/Pending (לפי תאריך ההצעה)
            year, month = date_to_use.year, date_to_use.month
            quotes_dir = _status_to_quotes_folder(DEFAULT_QUOTE_STATUS, year, month)
            quotes_dir.mkdir(parents=True, exist_ok=True)
            output_path = quotes_dir / filename_docx

            # Save Word to disk (גיבוי לעריכה)
            doc.save(str(output_path))

            # המרת Word ל-PDF באמצעות LibreOffice (לענן לינוקס)
            pdf_filename = str(output_path).replace(".docx", ".pdf")
            pdf_path = Path(pdf_filename)
            outdir = str(quotes_dir)
            try:
                result = subprocess.run(
                    ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, str(output_path)],
                    capture_output=True,
                    text=True,
                    timeout=60,
                )
                if not os.path.exists(pdf_filename):
                    st.error(f"המרת PDF נכשלה. פלט: {result.stderr or ''} {result.stdout or ''}")
                    st.session_state.pop("final_pdf_bytes", None)
                else:
                    with open(pdf_filename, "rb") as f:
                        st.session_state["final_pdf_bytes"] = f.read()
                    _persist_current_quote_temp_pdf(pdf_path)
            except Exception as pdf_err:
                st.error(f"המסמך נוצר, אך המרה ל-PDF נכשלה: {pdf_err}")

            # העלאת קובצי ה-Word וה-PDF לדרופבוקס (בענן אין סנכרון אוטומטי)
            dropbox_shared_link = ""
            try:
                access_token = _get_dropbox_access_token()
                if not access_token:
                    raise RuntimeError("אין DROPBOX_ACCESS_TOKEN ב-Secrets")
                dbx = dropbox.Dropbox(access_token)
                if PathRoot:
                    try:
                        ns_id = st.secrets.get("DROPBOX_NAMESPACE_ID")
                        if ns_id and str(ns_id).strip():
                            dbx = dbx.with_path_root(PathRoot.namespace_id(str(ns_id).strip()))
                        else:
                            acc = dbx.users_get_current_account()
                            if acc and getattr(acc, "root_info", None):
                                root_ns = getattr(acc.root_info, "root_namespace_id", None)
                                if root_ns:
                                    dbx = dbx.with_path_root(PathRoot.root(root_ns))
                    except Exception:
                        pass
                status_name = _status_to_folder_name(DEFAULT_QUOTE_STATUS)
                dbx_quotes_path_docx = f"/Studio84/StudioManager/Quotes/{year}/{month:02d}/{status_name}/{filename_docx}"
                dbx_quotes_path_pdf = f"/Studio84/StudioManager/Quotes/{year}/{month:02d}/{status_name}/{filename_pdf}"
                # יצירת תיקיות הורה במידת הצורך
                path_parts = [p for p in dbx_quotes_path_docx.split("/") if p][:-1]  # ללא שם הקובץ
                for i in range(1, len(path_parts) + 1):
                    parent = "/" + "/".join(path_parts[:i])
                    try:
                        dbx.files_create_folder_v2(parent)
                    except dropbox.exceptions.ApiError:
                        pass
                with open(str(output_path), "rb") as f:
                    dbx.files_upload(f.read(), dbx_quotes_path_docx, mode=dropbox.files.WriteMode.overwrite)
                if pdf_path.exists():
                    with open(str(pdf_path), "rb") as f:
                        dbx.files_upload(f.read(), dbx_quotes_path_pdf, mode=dropbox.files.WriteMode.overwrite)
                    st.success("✅ קובצי ה-Word וה-PDF נוצרו והועלו לדרופבוקס בהצלחה!")
                else:
                    st.success("✅ קובץ ה-Word נוצר והועלה לדרופבוקס בהצלחה!")
                # יצירת קישור שיתוף לפתיחה בדפדפן (PDF עדיף, אחרת Word)
                try:
                    def _get_quote_shared_link(path: str) -> str:
                        try:
                            link_metadata = dbx.sharing_create_shared_link_with_settings(path)
                            return link_metadata.url
                        except dropbox.exceptions.ApiError as e:
                            err = getattr(e, "error", None)
                            if err is not None and getattr(err, "is_shared_link_already_exists", lambda: False)():
                                meta = getattr(err, "get_shared_link_already_exists", lambda: None)()
                                if meta and getattr(meta, "is_metadata", lambda: False)():
                                    link_meta = meta.get_metadata()
                                    if link_meta:
                                        return link_meta.url
                                try:
                                    links_result = dbx.sharing_list_shared_links(path=path, direct_only=True)
                                    if links_result.links:
                                        return links_result.links[0].url
                                except Exception:
                                    pass
                            return ""
                    dropbox_shared_link = _get_quote_shared_link(dbx_quotes_path_pdf) if pdf_path.exists() else _get_quote_shared_link(dbx_quotes_path_docx)
                except Exception:
                    pass
            except Exception as dbx_err:
                st.warning(f"הקובץ נוצר בהצלחה, אך העלאה לדרופבוקס נכשלה: {dbx_err}")

            # עדכון מספור אוטומטי ב-config.json (רק בהצעה חדשה)
            if not is_edit_mode:
                try:
                    current_next = int(config.get("next_quote_number", 1))
                except (TypeError, ValueError):
                    current_next = 1
                try:
                    entered_number = int(quote_number)
                    new_next = max(current_next, entered_number + 1)
                except ValueError:
                    new_next = current_next + 1
                config["next_quote_number"] = new_next
                save_config(config)

            # שמירה / עדכון: quotes (Google Sheets בלבד)
            full_path_pdf = pdf_path.resolve() if pdf_path.exists() else output_path.resolve()
            quote_csv_row = {
                "Date": display_date,
                "Client": (client_name or "").strip(),
                "Project": (project_name or "").strip(),
                "Version": quote_version,
                "Quote Number": (quote_number or "").strip(),
                "Contact Person": (contact_person or "").strip(),
                "Client Email": (client_email or "").strip(),
                "Quote Subject": (quote_subject or "").strip(),
                "show_exterior": str(show_exterior),
                "show_interior": str(show_interior),
                "show_drone": str(show_drone),
                "show_video": str(show_video),
                "show_shots": str(show_shots),
                "scope_of_work": scope_of_work or "",
                "work_process": work_process or "",
                "delivery_time": delivery_time or "",
                "video_terms": video_terms if show_video else "",
                "price_exterior": str(price_exterior),
                "base_views_count": str(base_views_count),
                "price_ext_extra": str(price_ext_extra),
                "price_interior": str(price_interior),
                "price_int_extra": str(price_int_extra),
                "price_int_space": str(price_int_space),
                "price_drone": str(price_drone),
                "price_video": str(price_video),
                "price_shot_unit": str(price_shot_unit),
                "shots_count": str(shots_count),
                "include_price_exterior": str(include_price_exterior),
                "include_price_ext_extra": str(include_price_ext_extra),
                "include_price_interior": str(include_price_interior),
                "include_price_int_extra": str(include_price_int_extra),
                "include_price_int_space": str(include_price_int_space),
                "include_price_drone": str(include_price_drone),
                "include_price_video": str(include_price_video),
                "include_total_shots_price": str(include_total_shots_price),
                "model_update_val": str(model_update_val),
                "view_update_val": str(view_update_val),
                "extra_view_val": str(extra_view_val),
                "custom_item_desc": (custom_item_desc or "").strip(),
                "custom_item_price": str(custom_item_price),
                "Status": DEFAULT_QUOTE_STATUS,
                "File Path": str(full_path_pdf),
                "Signed File Path": "",
                "Total Price": f"{total_inc_vat:.2f}",
            }
            merged_csv = {c: "" for c in QUOTES_CSV_COLUMNS}
            if is_edit_mode and edit_quote_row:
                prev_full = get_quote_from_csv_by_log_row(edit_quote_row)
                if prev_full:
                    merged_csv = {c: (prev_full.get(c) or "") for c in QUOTES_CSV_COLUMNS}
            merged_csv.update(quote_csv_row)
            merged_csv["מקדמה שולמה"] = _normalize_payment_cell(merged_csv.get("מקדמה שולמה"))
            merged_csv["סכום מקדמה"] = _format_advance_amount_storage(merged_csv.get("סכום מקדמה"))
            merged_csv["גמר חשבון שולם"] = _normalize_payment_cell(merged_csv.get("גמר חשבון שולם"))
            merged_csv["freelancer_cost"] = _format_advance_amount_storage(merged_csv.get("freelancer_cost"))
            merged_csv["misc_expenses"] = _format_advance_amount_storage(merged_csv.get("misc_expenses"))
            quote_csv_row = merged_csv
            if is_edit_mode and orig_client is not None and orig_project is not None and orig_version:
                if not update_quote_in_csv(orig_client, orig_project, orig_version, quote_csv_row):
                    append_quote_to_csv(quote_csv_row)
            else:
                append_quote_to_csv(quote_csv_row)

            # שמירת פרטי ההצעה האחרונה שנוצרה עבור כפתור עדכון הסטטוס
            st.session_state["last_quote_client"] = (client_name or "").strip()
            st.session_state["last_quote_project"] = (project_name or "").strip()
            st.session_state["last_quote_version"] = (quote_version or "").strip()

            # ניקוי מפתחות עריכה כדי שהטופס יתאפס ליצירה חדשה
            for k in list(st.session_state.keys()):
                if k.startswith("edit_"):
                    del st.session_state[k]

            # שמירת נתיבים ופרטי המייל ב-session state - המסך יוצג מחוץ ל-if
            display_path = pdf_path if pdf_path.exists() else output_path
            st.session_state['current_pdf_path'] = str(full_path_pdf)
            st.session_state['current_docx_path'] = str(output_path)
            st.session_state['current_file_path'] = str(full_path_pdf)  # נתיב לקובץ לצירוף למייל
            st.session_state['current_dropbox_link'] = dropbox_shared_link
            # שמירה פיזית ב-temp_proposals כדי שהקובץ יהיה זמין אחרי rerun (לפני שליחת מייל)
            TEMP_PROPOSALS_DIR.mkdir(parents=True, exist_ok=True)
            src_file = full_path_pdf if full_path_pdf.exists() else output_path
            temp_copy = TEMP_PROPOSALS_DIR / src_file.name
            shutil.copy2(str(src_file), str(temp_copy))
            st.session_state['proposal_file_path'] = os.path.abspath(str(temp_copy))
            st.session_state['current_quotes_dir'] = str(quotes_dir)
            st.session_state['current_display_path'] = str(display_path)
            # שמירת תוכן הקבצים בבתים - נשמר גם אחרי רענון מסך, לצירוף למייל ולהורדה
            try:
                pdf_exists = full_path_pdf.exists()
                if pdf_path.exists():
                    with open(str(pdf_path), 'rb') as f:
                        _pdf_data = f.read()
                        st.session_state['pdf_bytes'] = _pdf_data
                        st.session_state['final_pdf_bytes'] = _pdf_data
                    _persist_current_quote_temp_pdf(pdf_path)
                else:
                    st.session_state['pdf_bytes'] = None
                    st.session_state['final_pdf_bytes'] = None
                main_bytes = full_path_pdf.read_bytes() if pdf_exists else output_path.read_bytes()
                main_name = full_path_pdf.name if pdf_exists else output_path.name
                st.session_state['current_file_bytes'] = main_bytes
                st.session_state['current_file_name'] = main_name
                st.session_state['current_file_bytes_pdf'] = main_bytes if pdf_exists else None
                st.session_state['current_file_bytes_docx'] = output_path.read_bytes()
                # שמירה ייעודית לצירוף למייל - מונע אובדן הקובץ ב-rerun (אותו משתנה שכפתור ההורדה משתמש בו)
                st.session_state['pdf_bytes_for_email'] = main_bytes if pdf_exists else None
            except Exception:
                st.session_state['current_file_bytes'] = None
                st.session_state['current_file_name'] = None
                st.session_state['current_file_bytes_pdf'] = None
                st.session_state['current_file_bytes_docx'] = None
                st.session_state['pdf_bytes_for_email'] = None
                st.session_state['pdf_bytes'] = None
                st.session_state['final_pdf_bytes'] = None
            st.session_state['current_email_client'] = client_email or ""
            st.session_state['current_contact_person'] = contact_person or ""
            st.session_state['current_project_name'] = project_name or ""
            st.session_state['current_quote_version'] = quote_version or ""
            st.session_state['current_client_name'] = client_name or ""
        except Exception as e:
            st.error(f'שגיאה מפורטת: {str(e)}')

    # סינכרון session state במצב עריכה - כדי שאזור התצוגה המקדימה והשליחה יוצג
    if is_edit_mode and edit_quote_row:
        file_path = (edit_quote_row.get("File Path") or "").strip()
        if file_path:
            path_abs = str(Path(file_path).resolve())
            fp = Path(file_path)
            if not fp.exists():
                fp = find_proposal_file(fp.name)
            if fp and fp.exists():
                # שמירה פיזית ב-temp_proposals כדי שהקובץ יהיה זמין אחרי rerun (לפני שליחת מייל)
                TEMP_PROPOSALS_DIR.mkdir(parents=True, exist_ok=True)
                temp_copy = TEMP_PROPOSALS_DIR / fp.name
                shutil.copy2(str(fp), str(temp_copy))
                st.session_state["proposal_file_path"] = os.path.abspath(str(temp_copy))
            else:
                st.session_state["proposal_file_path"] = path_abs  # fallback לנתיב המקורי
            st.session_state["current_pdf_path"] = path_abs
            docx_path_candidate = path_abs.replace(".pdf", ".docx")
            st.session_state["current_docx_path"] = docx_path_candidate
            st.session_state["current_file_path"] = path_abs
            st.session_state["current_dropbox_link"] = ""  # במצב עריכה אין קישור Dropbox (נוצר רק בהעלאה)
            st.session_state["current_quotes_dir"] = str(Path(file_path).parent)
            st.session_state["current_display_path"] = file_path
            st.session_state["current_email_client"] = client_email or ""
            st.session_state["current_contact_person"] = contact_person or ""
            st.session_state["current_project_name"] = project_name or ""
            st.session_state["current_quote_version"] = (edit_quote_row.get("Version") or "").strip() or "V1"
            st.session_state["current_client_name"] = client_name or ""
            # טעינת bytes מהקובץ במצב עריכה (אם הקובץ קיים על הדיסק)
            try:
                fp = Path(file_path)
                if not fp.exists():
                    fp = find_proposal_file(fp.name)
                if fp and fp.exists():
                    main_bytes = fp.read_bytes()
                    st.session_state["current_file_bytes"] = main_bytes
                    st.session_state["current_file_name"] = fp.name
                    st.session_state["current_file_bytes_pdf"] = main_bytes if (fp.suffix or "").lower() == ".pdf" else None
                    docx_p = Path(docx_path_candidate)
                    st.session_state["current_file_bytes_docx"] = docx_p.read_bytes() if docx_p.exists() else None
                    # שמירה ייעודית לצירוף למייל - מונע אובדן הקובץ ב-rerun (אותו משתנה שכפתור ההורדה משתמש בו)
                    _is_pdf_edit = (fp.suffix or "").lower() == ".pdf"
                    st.session_state["pdf_bytes_for_email"] = main_bytes if _is_pdf_edit else None
                    st.session_state["pdf_bytes"] = main_bytes if _is_pdf_edit else None
                    st.session_state["final_pdf_bytes"] = main_bytes if _is_pdf_edit else None
                    if _is_pdf_edit:
                        _persist_current_quote_temp_pdf(fp)
                else:
                    st.session_state["current_file_bytes"] = None
                    st.session_state["current_file_name"] = None
                    st.session_state["current_file_bytes_pdf"] = None
                    st.session_state["current_file_bytes_docx"] = None
                    st.session_state["pdf_bytes_for_email"] = None
                    st.session_state["pdf_bytes"] = None
                    st.session_state["final_pdf_bytes"] = None
            except Exception:
                st.session_state["current_file_bytes"] = None
                st.session_state["current_file_name"] = None
                st.session_state["current_file_bytes_pdf"] = None
                st.session_state["current_file_bytes_docx"] = None
                st.session_state["pdf_bytes_for_email"] = None
                st.session_state["pdf_bytes"] = None
                st.session_state["final_pdf_bytes"] = None

    # הצגת אזור תצוגה מקדימה ושליחה - תמיד כשנבחרה הצעה (מעריכה או מיצירה)
    has_quote_for_preview = "current_pdf_path" in st.session_state
    if has_quote_for_preview:
        # עדכון פרטי המייל מהטופס הנוכחי (במצב עריכה)
        if is_edit_mode and edit_quote_row:
            st.session_state["current_email_client"] = client_email or ""
            st.session_state["current_contact_person"] = contact_person or ""
            st.session_state["current_project_name"] = project_name or ""
            st.session_state["current_client_name"] = client_name or ""

        quotes_dir_str = st.session_state.get("current_quotes_dir", "")
        if not is_edit_mode:
            st.success(f"הקובץ נוצר בהצלחה ונשמר ב:\n{quotes_dir_str}\n(Word + PDF)")

        # כפתורי הורדה - מיד לאחר שמירת ההצעה וההעלאה לדרופבוקס (משתמשים ב-bytes מ-session_state אם הקובץ לא על הדיסק)
        client_name_dl = st.session_state.get("current_client_name", "Client")
        safe_client = sanitize_filename_part(client_name_dl) if client_name_dl else "Client"

        # כפתור הורדת קובץ Word
        docx_bytes_data = st.session_state.get("current_file_bytes_docx")
        if not docx_bytes_data:
            docx_path_str = st.session_state.get("current_docx_path", "")
            if docx_path_str:
                docx_path = Path(docx_path_str)
                if docx_path.exists():
                    docx_bytes_data = docx_path.read_bytes()
        if docx_bytes_data:
            st.download_button(
                label="📥 הורדת קובץ Word",
                data=docx_bytes_data,
                file_name=f"Quote_{safe_client}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_quote_docx_btn",
            )

        # כפתור הורדת קובץ PDF
        pdf_bytes_data = st.session_state.get("current_file_bytes_pdf")
        if not pdf_bytes_data:
            display_path_str = st.session_state.get("current_display_path", st.session_state.get("current_pdf_path", ""))
            target = Path(display_path_str) if display_path_str else None
            if target and not target.exists():
                target = find_proposal_file(target.name) if target else None
            if target and target.exists():
                pdf_bytes_data = target.read_bytes()
        if pdf_bytes_data:
            st.download_button(
                label="📥 הורדת קובץ PDF",
                data=pdf_bytes_data,
                file_name=f"Quote_{safe_client}.pdf",
                mime="application/pdf",
                key="download_quote_btn",
            )

        # קישור לפתיחת המסמך בדפדפן (מקישור Dropbox)
        dropbox_link = st.session_state.get("current_dropbox_link", "")
        if dropbox_link and dropbox_link.startswith("http"):
            st.link_button("📄 פתח את המסמך בדפדפן", dropbox_link, type="primary")

        # --- אזור תצוגה מקדימה ושליחה 📨 ---
        st.subheader("תצוגה מקדימה ושליחה 📨")
        project_name_mail = st.session_state.get("current_project_name", "")
        quote_version_mail = st.session_state.get("current_quote_version", "")
        contact_person_mail = st.session_state.get("current_contact_person", "")
        client_email_mail = st.session_state.get("current_email_client", "")
        email_subject = f"הצעת מחיר: {project_name_mail} - סטודיו 84 (גרסה {quote_version_mail})"
        email_body = (
            f"היי {contact_person_mail},\n"
            f"בהמשך לשיחתנו, מצורפת הצעת מחיר עבור פרויקט {project_name_mail}.\n"
            "אשמח לעמוד לרשותך לכל שאלה.\n\n"
            "בברכה,\n"
            "סטודיו 84"
        )
        email_tali_str = EMAIL_MYSELF
        email_eran_str = EMAIL_ERAN
        cc_base = [EMAIL_ACCOUNTING]
        file_ready = CURRENT_QUOTE_TEMP_PDF.exists()
        if file_ready:
            st.success("📎 הקובץ מצורף ומוכן לשליחה ללקוח")
        else:
            st.warning("⚠️ הקובץ לצירוף לא נמצא – ודא שההצעה נשמרה או הורד את הקובץ מחדש")
        if file_ready:
            st.download_button(
                label="📥 הורד והצג את ה-PDF לפני שליחה (current_quote_temp.pdf)",
                data=CURRENT_QUOTE_TEMP_PDF.read_bytes(),
                file_name="current_quote_temp.pdf",
                mime="application/pdf",
                key="preview_quote_temp_pdf_download",
            )
        else:
            st.caption("לאחר שמירת ההצעה יופיע כאן כפתור להורדת ה-PDF לבדיקה.")
        st.caption("תצוגה מקדימה של המייל:")
        st.text_input("נושא", value=email_subject, key="quote_email_preview_subject")
        st.text_area("תוכן", value=email_body, height=120, key="quote_email_preview_body")
        send_via = st.radio(
            "בחר שיטת שליחה",
            ["שליחה מ-Gmail (טלי)", "שליחה מ-Webmail (ערן)"],
            key="quote_send_via_radio",
            horizontal=True,
        )
        cc_list = cc_base + [email_tali_str, email_eran_str]
        if send_via == "שליחה מ-Gmail (טלי)":
            if st.button("שלח הצעת מחיר ללקוח 🚀", key="send_quote_btn_gmail", type="primary"):
                if not client_email_mail or not client_email_mail.strip():
                    st.error("חסרה כתובת אימייל ללקוח. הזן אימייל בשדה 'אימייל לקוח'.")
                elif not CURRENT_QUOTE_TEMP_PDF.exists():
                    st.error(
                        "הקובץ הפיזי לא נמצא. אנא לחץ שוב על 'שמור / עדכן הצעת מחיר' כדי לייצר אותו."
                    )
                else:
                    with st.spinner("שולח מייל..."):
                        ok = send_quote_email_via_smtp(
                            client_email_mail,
                            st.session_state.get(
                                "quote_email_preview_subject", email_subject
                            ),
                            st.session_state.get("quote_email_preview_body", email_body),
                            cc_list=cc_list,
                            smtp_profile="tali",
                        )
                    if ok:
                        st.success("המייל נשלח בהצלחה ללקוח!")
                        time.sleep(2)
                        st.rerun()
        else:
            if st.button("שלח הצעת מחיר ללקוח 🚀", key="send_quote_btn_webmail", type="primary"):
                if not client_email_mail or not client_email_mail.strip():
                    st.error("חסרה כתובת אימייל ללקוח. הזן אימייל בשדה 'אימייל לקוח'.")
                elif not CURRENT_QUOTE_TEMP_PDF.exists():
                    st.error(
                        "הקובץ הפיזי לא נמצא. אנא לחץ שוב על 'שמור / עדכן הצעת מחיר' כדי לייצר אותו."
                    )
                else:
                    with st.spinner("שולח מייל..."):
                        ok = send_quote_email_via_smtp(
                            client_email_mail,
                            st.session_state.get(
                                "quote_email_preview_subject", email_subject
                            ),
                            st.session_state.get("quote_email_preview_body", email_body),
                            cc_list=cc_list,
                            smtp_profile="eran",
                        )
                    if ok:
                        st.success("המייל נשלח בהצלחה ללקוח!")
                        time.sleep(2)
                        st.rerun()

        if st.button("🔄 התחל הצעה חדשה", key="reset_quote_btn"):
            for k in [
                "current_pdf_path",
                "current_docx_path",
                "current_file_path",
                "proposal_file_path",
                "current_dropbox_link",
                "current_quotes_dir",
                "current_display_path",
                "current_email_client",
                "current_contact_person",
                "current_project_name",
                "current_quote_version",
                "current_client_name",
                "current_file_bytes",
                "current_file_name",
                "current_file_bytes_pdf",
                "current_file_bytes_docx",
                "pdf_bytes_for_email",
                "pdf_bytes",
                "final_pdf_bytes",
                "quote_email_preview_subject",
                "quote_email_preview_body",
            ]:
                st.session_state.pop(k, None)
            st.rerun()


FINANCE_QUOTE_CANCELLED_STATUSES = frozenset({"Rejected"})

# דרישת תשלום / חשבון עסקה — מע"מ (חשבונית עסקה בישראל)
TRANSACTION_INVOICE_VAT_RATE = 0.17


def _resolve_hebrew_ttf_font_pair() -> tuple[Path | None, Path | None]:
    """נתיבים לפונט רגיל ומודגש לעברית (Arial/Segoe ב-Windows, DejaVu בלינוקס)."""
    windir = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    candidates_reg = [
        windir / "arial.ttf",
        windir / "segoeui.ttf",
        windir / "david.ttf",
    ]
    candidates_bold = [
        windir / "arialbd.ttf",
        windir / "segoeuib.ttf",
        windir / "davidbd.ttf",
    ]
    for reg in candidates_reg:
        if reg.exists():
            for bd in candidates_bold:
                if bd.exists():
                    return reg, bd
            return reg, reg
    linux_reg = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
    linux_bd = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")
    if linux_reg.exists():
        return linux_reg, linux_bd if linux_bd.exists() else linux_reg
    return None, None


def _pdf_bidi_text(s: str) -> str:
    if not s:
        return ""
    try:
        return get_display(str(s))
    except Exception:
        return str(s)


def _studio_transaction_invoice_footer_text() -> str:
    """ניתן לעקוף ב-.streamlit/secrets.toml: transaction_invoice_footer = '''...'''"""
    try:
        ft = st.secrets.get("transaction_invoice_footer", "")
        if isinstance(ft, str) and ft.strip():
            return ft.strip()
    except Exception:
        pass
    return (
        "בנק הפועלים | סניף 123 | חשבון 123456\n"
        "עוסק מורשה: 123456789"
    )


def build_transaction_invoice_pdf_bytes(
    line_df: pd.DataFrame,
    client_name: str,
    project_name: str,
    serial: str,
    subtotal_ex_vat: float,
    vat_amount: float,
    total_with_vat: float,
    footer_text: str,
) -> bytes:
    """יוצר PDF של חשבון עסקה עם טקסט עברי (TTF + bidi)."""
    font_reg, font_bold = _resolve_hebrew_ttf_font_pair()
    if not font_reg:
        raise RuntimeError(
            "לא נמצא פונט TTF לעברית. התקן Arial/Segoe ב-Windows או DejaVu בלינוקס."
        )

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()
    pdf.add_font("Hebrew", "", str(font_reg))
    pdf.add_font("Hebrew", "B", str(font_bold))

    left_m = 12
    pdf.set_margins(left_m, 12, left_m)
    pdf.set_xy(left_m, 12)

    pdf.set_font("Hebrew", "B", 17)
    pdf.cell(95, 9, "STUDIO 84", align="L")
    pdf.set_font("Hebrew", "B", 14)
    pdf.cell(0, 9, _pdf_bidi_text(f"חשבון עסקה מס׳ {serial}"), align="R", ln=1)

    pdf.set_font("Hebrew", "", 10)
    pdf.ln(2)
    pdf.cell(0, 7, _pdf_bidi_text(f"לכבוד: {client_name}"), align="R", ln=1)
    pdf.cell(0, 7, _pdf_bidi_text(f"הנדון: {project_name}"), align="R", ln=1)
    pdf.cell(0, 6, _pdf_bidi_text(f"תאריך: {date.today().strftime('%d/%m/%Y')}"), align="R", ln=1)
    pdf.ln(4)

    # כותרות טבלה
    w_desc, w_qty, w_price, w_line = 88, 22, 28, 32
    pdf.set_font("Hebrew", "B", 10)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(w_desc, 8, _pdf_bidi_text("תיאור"), border=1, align="C", fill=True)
    pdf.cell(w_qty, 8, _pdf_bidi_text("כמות"), border=1, align="C", fill=True)
    pdf.cell(w_price, 8, _pdf_bidi_text("מחיר (₪)"), border=1, align="C", fill=True)
    pdf.cell(w_line, 8, _pdf_bidi_text("סכום (₪)"), border=1, align="C", fill=True, ln=1)

    pdf.set_font("Hebrew", "", 9)
    for _, row in line_df.iterrows():
        desc = str(row.get("תיאור") or "").strip()
        try:
            qty = float(row.get("כמות") or 0)
        except (TypeError, ValueError):
            qty = 0.0
        try:
            unit_price = float(row.get("מחיר") or 0)
        except (TypeError, ValueError):
            unit_price = 0.0
        line_total = qty * unit_price
        if not desc and line_total == 0:
            continue
        pdf.cell(w_desc, 8, _pdf_bidi_text(desc[:120]), border=1, align="R")
        pdf.cell(w_qty, 8, f"{qty:g}", border=1, align="C")
        pdf.cell(w_price, 8, f"{unit_price:,.0f}", border=1, align="C")
        pdf.cell(w_line, 8, f"{line_total:,.0f}", border=1, align="C", ln=1)

    pdf.ln(3)
    pdf.set_font("Hebrew", "", 10)
    pdf.cell(0, 7, _pdf_bidi_text(f"סה״כ לפני מע״מ: {subtotal_ex_vat:,.2f} ₪"), align="R", ln=1)
    pdf.cell(0, 7, _pdf_bidi_text(f"מע״מ ({TRANSACTION_INVOICE_VAT_RATE * 100:.0f}%): {vat_amount:,.2f} ₪"), align="R", ln=1)
    pdf.set_font("Hebrew", "B", 11)
    pdf.cell(0, 8, _pdf_bidi_text(f"סה״כ לתשלום כולל מע״מ: {total_with_vat:,.2f} ₪"), align="R", ln=1)

    pdf.ln(6)
    pdf.set_font("Hebrew", "", 9)
    pdf.multi_cell(0, 5, _pdf_bidi_text(footer_text.strip()), align="R")

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin-1")
    return bytes(out)


def _finance_dedupe_quotes_df(quotes_df: pd.DataFrame) -> pd.DataFrame:
    """שורה אחת לכל (Client, Project) — הגרסה הגבוהה ביותר; רק הצעות שאינן מבוטלות."""
    if quotes_df.empty or "Status" not in quotes_df.columns:
        return quotes_df.iloc[0:0].copy()
    st_col = quotes_df["Status"].fillna("").astype(str).str.strip()
    active = quotes_df[~st_col.isin(FINANCE_QUOTE_CANCELLED_STATUSES)].copy()
    if active.empty:
        return active
    best_idx: dict[tuple[str, str], int] = {}
    best_ver: dict[tuple[str, str], int] = {}
    for i in active.index:
        row = active.loc[i]
        c = str(row.get("Client") or "").strip()
        p = str(row.get("Project") or "").strip()
        if not p:
            continue
        key = (c, p)
        pv = parse_version_number(str(row.get("Version") or ""))
        if key not in best_ver or pv > best_ver[key]:
            best_ver[key] = pv
            best_idx[key] = i
    keep = list(best_idx.values())
    out = active.loc[keep].copy()
    return out.sort_values(by=["Client", "Project"]).reset_index(drop=True)


def _show_finance_collection_dashboard() -> None:
    """דשבורד כספים וגבייה — מדדים, טבלת מעקב ועדכון סטטוס תשלום."""
    st.subheader("💰 דשבורד פיננסי וגבייה")

    rows = read_quotes_csv()
    if not rows:
        st.info("אין נתוני הצעות בגיליון quotes.")
        return

    quotes_df = pd.DataFrame(rows)
    quotes_df = quotes_df.reindex(columns=QUOTES_CSV_COLUMNS, fill_value="")
    for pay_col in ("מקדמה שולמה", "גמר חשבון שולם"):
        if pay_col in quotes_df.columns:
            quotes_df[pay_col] = quotes_df[pay_col].map(_normalize_payment_cell)

    deduped = _finance_dedupe_quotes_df(quotes_df)
    if deduped.empty:
        st.info("אין פרויקטים להצגה (לאחר סינון הצעות מבוטלות).")
        return

    projects_rows = read_projects()
    proj_status: dict[tuple[str, str], str] = {}
    for pr in projects_rows:
        ck = ((pr.get("Client") or "").strip(), (pr.get("Project Name") or "").strip())
        if ck[1]:
            proj_status[ck] = (pr.get("Status") or "").strip()

    totals: list[float] = []
    collected_full_final: list[float] = []
    total_received_parts: list[float] = []
    total_expected_gross = 0.0
    for _, row in deduped.iterrows():
        rdict = row.to_dict()
        price = _extract_total_from_quote_row(rdict)
        adv = min(_parse_currency_amount(rdict.get("סכום מקדמה")), price)
        fc_g = _parse_currency_amount(rdict.get("freelancer_cost"))
        me_g = _parse_currency_amount(rdict.get("misc_expenses"))
        total_expected_gross += price - (fc_g + me_g)
        totals.append(price)
        if _payment_marked_yes(rdict.get("גמר חשבון שולם")):
            collected_full_final.append(price)
            total_received_parts.append(price)
        else:
            collected_full_final.append(0.0)
            total_received_parts.append(adv)

    total_volume = sum(totals)
    total_collected_full = sum(collected_full_final)
    total_received = sum(total_received_parts)
    open_balance = total_volume - total_received

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("סך היקף עסקאות (פעילים, לפי גרסה אחרונה)", f"₪{total_volume:,.0f}")
    c2.metric("סך שנגבה במלואו (גמר חשבון שולם)", f"₪{total_collected_full:,.0f}")
    c3.metric("יתרת חוב פתוחה / צפי הכנסה", f"₪{open_balance:,.0f}")
    c4.metric("סך הרווח הגולמי הצפוי (פרויקטים פעילים)", f"₪{total_expected_gross:,.0f}")

    with st.expander("📄 הפקת דרישת תשלום (חשבון עסקה) ללקוח", expanded=False):
        active_for_demand = deduped[
            deduped["Status"].fillna("").astype(str).str.strip().isin(["Signed", "הומר לפרויקט"])
        ].copy().reset_index(drop=True)
        if active_for_demand.empty:
            st.caption("אין פרויקטים פעילים (Signed / הומר לפרויקט) לבחירה.")
        else:
            demand_labels: list[str] = []
            for _, r in active_for_demand.iterrows():
                c_l = str(r.get("Client") or "").strip()
                p_l = str(r.get("Project") or "").strip()
                demand_labels.append(f"{p_l} — {c_l}")
            demand_ix = st.selectbox(
                "בחר פרויקט פעיל",
                options=list(range(len(active_for_demand))),
                format_func=lambda i: demand_labels[i] if i < len(demand_labels) else str(i),
                key="finance_payment_demand_project_select",
            )
            r_dem = active_for_demand.iloc[int(demand_ix)]
            client_dem = str(r_dem.get("Client") or "").strip()
            project_dem = str(r_dem.get("Project") or "").strip()
            rdict_dem = r_dem.to_dict()
            price_dem = _extract_total_from_quote_row(rdict_dem)
            adv_dem = min(_parse_currency_amount(r_dem.get("סכום מקדמה")), price_dem)
            balance_dem = price_dem - adv_dem
            _main_u, _up_u, drop_deliv_dem = find_project_dropbox_links_for_client(client_dem, project_dem)
            link_line = (drop_deliv_dem or "").strip() or "לא הוזן קישור"

            st.caption(
                "ערכו את הסעיפים לפי הצורך; ניתן להוסיף שורות (למשל תוספת מבטים או שעות עריכה). "
                "הסכומים מחושבים ככמות × מחיר; המע״מ מחושב על הסה״כ לפני מע״מ."
            )
            demand_df = pd.DataFrame(
                [
                    {
                        "תיאור": f"תשלום עבור פרויקט {project_dem}",
                        "כמות": 1.0,
                        "מחיר": float(balance_dem),
                    },
                    {"תיאור": "", "כמות": 1.0, "מחיר": 0.0},
                    {"תיאור": "", "כמות": 1.0, "מחיר": 0.0},
                ]
            )
            edited_demand = st.data_editor(
                demand_df,
                column_config={
                    "תיאור": st.column_config.TextColumn("תיאור", width="large"),
                    "כמות": st.column_config.NumberColumn("כמות", min_value=0.0, step=0.5, format="%.2f"),
                    "מחיר": st.column_config.NumberColumn("מחיר (₪)", min_value=0.0, step=50.0, format="%.0f"),
                },
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                key=f"finance_demand_lines_{demand_ix}",
            )

            subtotal = 0.0
            if edited_demand is not None and not edited_demand.empty:
                ed = edited_demand.copy()
                ed["כמות"] = pd.to_numeric(ed["כמות"], errors="coerce").fillna(0.0)
                ed["מחיר"] = pd.to_numeric(ed["מחיר"], errors="coerce").fillna(0.0)
                subtotal = float((ed["כמות"] * ed["מחיר"]).sum())
            vat_amt = subtotal * TRANSACTION_INVOICE_VAT_RATE
            total_pay = subtotal + vat_amt

            tot1, tot2, tot3 = st.columns(3)
            tot1.metric('סה"כ לפני מע"מ', f"₪{subtotal:,.0f}")
            tot2.metric(
                f'מע"מ ({TRANSACTION_INVOICE_VAT_RATE * 100:.0f}%)',
                f"₪{vat_amt:,.0f}",
            )
            tot3.metric('סה"כ לתשלום (כולל מע"מ)', f"₪{total_pay:,.0f}")

            ss_serial = f"txn_inv_serial_{demand_ix}"
            if ss_serial not in st.session_state:
                st.session_state[ss_serial] = (
                    f"{datetime.now().strftime('%Y%m%d')}-{random.randint(1000, 9999)}"
                )
            serial = st.session_state[ss_serial]

            pdf_bytes = b""
            try:
                footer_t = _studio_transaction_invoice_footer_text()
                edf = edited_demand if edited_demand is not None else demand_df
                pdf_bytes = build_transaction_invoice_pdf_bytes(
                    edf,
                    client_dem,
                    project_dem,
                    serial,
                    subtotal,
                    vat_amt,
                    total_pay,
                    footer_t,
                )
            except Exception as e:
                st.error(f"שגיאה ביצירת PDF: {e}")

            if pdf_bytes:
                st.download_button(
                    label="הורד PDF — חשבון עסקה",
                    data=pdf_bytes,
                    file_name=f"hesbon_iska_{serial}.pdf",
                    mime="application/pdf",
                    key=f"finance_demand_pdf_dl_{demand_ix}",
                )

            st.markdown("###### קישור לתיקיית תוצרים (דרופבוקס)")
            st.text_area(
                "קישור להעתקה ללקוח",
                value=link_line,
                height=100,
                disabled=True,
                help="בחרו את כל הטקסט (Ctrl+A) והעתיקו לגוף המייל.",
                key=f"finance_demand_dbx_{demand_ix}",
            )

    st.markdown("##### טבלת מעקב גבייה")
    table_rows = []
    for _, row in deduped.iterrows():
        client = str(row.get("Client") or "").strip()
        project = str(row.get("Project") or "").strip()
        rdict = row.to_dict()
        price = _extract_total_from_quote_row(rdict)
        st_proj = proj_status.get((client, project), "")
        if not st_proj:
            st_proj = str(row.get("Status") or "").strip()
        adv_amt = min(_parse_currency_amount(row.get("סכום מקדמה")), price)
        fc = _parse_currency_amount(rdict.get("freelancer_cost"))
        me = _parse_currency_amount(rdict.get("misc_expenses"))
        gross = price - (fc + me)
        table_rows.append(
            {
                "שם פרויקט": project,
                "לקוח": client,
                "סטטוס": st_proj,
                "שלב בקנבן": _normalized_kanban_stage(str(row.get("שלב עבודה") or "")),
                "מחיר בהצעה": price,
                "freelancer_cost": fc,
                "misc_expenses": me,
                "רווח גולמי": gross,
                "מקדמה שולמה": _normalize_payment_cell(row.get("מקדמה שולמה")),
                "סכום מקדמה": adv_amt,
                "גמר חשבון שולם": _normalize_payment_cell(row.get("גמר חשבון שולם")),
            }
        )
    display_df = pd.DataFrame(table_rows)
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "מחיר בהצעה": st.column_config.NumberColumn("מחיר בהצעה (₪)", format="%.0f"),
            "freelancer_cost": st.column_config.NumberColumn(
                "עלות קבלני משנה / פרילנסרים (₪)", format="%.0f"
            ),
            "misc_expenses": st.column_config.NumberColumn(
                "הוצאות פרויקט נוספות (₪)", format="%.0f"
            ),
            "רווח גולמי": st.column_config.NumberColumn("רווח גולמי (₪)", format="%.0f"),
            "סכום מקדמה": st.column_config.NumberColumn("סכום מקדמה (₪)", format="%.0f"),
        },
    )

    st.markdown("##### עדכון סטטוס תשלום")
    row_labels: list[str] = []
    for _, row in deduped.iterrows():
        client = str(row.get("Client") or "").strip()
        project = str(row.get("Project") or "").strip()
        version = str(row.get("Version") or "").strip() or "V1"
        row_labels.append(f"{project} — {client} (גרסה {version})")
    n_dedup = len(deduped)
    selected_ix = st.selectbox(
        "בחר פרויקט פעיל",
        options=list(range(n_dedup)),
        format_func=lambda i: row_labels[i] if i < len(row_labels) else str(i),
        key="finance_pay_project_select",
    )
    r0 = deduped.iloc[int(selected_ix)]
    adv_default = _payment_marked_yes(r0.get("מקדמה שולמה"))
    adv_amt_default = _parse_currency_amount(r0.get("סכום מקדמה"))
    final_default = _payment_marked_yes(r0.get("גמר חשבון שולם"))
    freelancer_default = _parse_currency_amount(r0.get("freelancer_cost"))
    misc_default = _parse_currency_amount(r0.get("misc_expenses"))

    adv1, adv2, cb_final = st.columns([1, 1, 1])
    with adv1:
        adv_amt_in = st.number_input(
            "סכום מקדמה ששולם (₪)",
            min_value=0.0,
            value=float(adv_amt_default),
            step=100.0,
            format="%.0f",
            key=f"finance_pay_adv_amt_{selected_ix}",
        )
    with adv2:
        adv_ok = st.checkbox(
            "מקדמה שולמה",
            value=adv_default,
            key=f"finance_pay_adv_{selected_ix}",
        )
    with cb_final:
        final_ok = st.checkbox(
            "גמר חשבון שולם",
            value=final_default,
            key=f"finance_pay_final_{selected_ix}",
        )

    ex1, ex2 = st.columns(2)
    with ex1:
        freelancer_cost_in = st.number_input(
            "עלות קבלני משנה / פרילנסרים (₪)",
            min_value=0.0,
            value=float(freelancer_default),
            step=100.0,
            format="%.0f",
            key=f"finance_pay_freelancer_{selected_ix}",
        )
    with ex2:
        misc_expenses_in = st.number_input(
            "הוצאות פרויקט נוספות — מודלים, פלאגינים וכו׳ (₪)",
            min_value=0.0,
            value=float(misc_default),
            step=100.0,
            format="%.0f",
            key=f"finance_pay_misc_{selected_ix}",
        )

    if st.button("💾 עדכן סטטוס תשלום", key="finance_pay_save", type="primary"):
        date_v = str(r0.get("Date") or "").strip()
        client_v = str(r0.get("Client") or "").strip()
        project_v = str(r0.get("Project") or "").strip()
        version_v = str(r0.get("Version") or "").strip() or "V1"
        all_rows = read_quotes_csv()
        updated = False
        for i, qr in enumerate(all_rows):
            if (qr.get("Date") or "").strip() != date_v:
                continue
            if (qr.get("Client") or "").strip() != client_v:
                continue
            if (qr.get("Project") or "").strip() != project_v:
                continue
            if not _quote_versions_equivalent_for_match(version_v, (qr.get("Version") or "").strip()):
                continue
            merged = {c: (qr.get(c) or "") for c in QUOTES_CSV_COLUMNS}
            merged["סכום מקדמה"] = _format_advance_amount_storage(adv_amt_in)
            merged["מקדמה שולמה"] = "כן" if (adv_ok or adv_amt_in > 0) else "לא"
            merged["גמר חשבון שולם"] = "כן" if final_ok else "לא"
            merged["freelancer_cost"] = _format_advance_amount_storage(freelancer_cost_in)
            merged["misc_expenses"] = _format_advance_amount_storage(misc_expenses_in)
            all_rows[i] = merged
            updated = True
            break
        if updated:
            write_quotes_csv(all_rows, skip_rerun=True)
            st.cache_data.clear()
            st.rerun()
        else:
            st.error("לא נמצאה שורה מתאימה לעדכון.")


def show_quotes_management_page() -> None:
    st.title("ניהול הצעות")

    rows = read_quotes_log()
    if not rows:
        st.info("אין עדיין הצעות בגיליון quotes. צרו הצעה חדשה כדי להתחיל.")
        return

    # Prefer pandas for a nicer editor experience, but keep a fallback.
    try:
        import pandas as pd  # type: ignore

        df = pd.DataFrame(rows, columns=QUOTES_LOG_COLUMNS)
        df = df.fillna('')

        # --- מנוע חיפוש וסינון ---
        search_term = st.text_input(
            "🔍 חיפוש הצעה (לפי לקוח, פרויקט או תאריך)",
            key="quote_search_term",
            placeholder="הקלד לחיפוש...",
        )
        if search_term and str(search_term).strip():
            term = str(search_term).strip().lower()
            mask = (
                df["Client"].fillna("").astype(str).str.lower().str.contains(term, na=False)
                | df["Project"].fillna("").astype(str).str.lower().str.contains(term, na=False)
                | df["Date"].fillna("").astype(str).str.contains(term, na=False)
            )
            df_filtered = df[mask].copy()
        else:
            df_filtered = df.copy()

        st.caption(f"מוצגות {len(df_filtered)} מתוך {len(df)} הצעות")
        edited_df = st.data_editor(
            df_filtered,
            hide_index=True,
            use_container_width=True,
            disabled=[c for c in QUOTES_LOG_COLUMNS if c != "Status"],
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=ALLOWED_QUOTE_STATUSES,
                    required=True,
                )
            },
            key="quotes_editor",
        )

        # --- הפיכת הצעת מחיר מאושרת לפרויקט חדש ---
        def _quote_log_row_unique_key(row_dict: dict) -> str:
            return "|||".join(
                [
                    (row_dict.get("Date") or "").strip(),
                    (row_dict.get("Client") or "").strip(),
                    (row_dict.get("Project") or "").strip(),
                    (row_dict.get("Version") or "").strip() or "V1",
                ]
            )

        convertible = [
            (i, r)
            for i, r in edited_df.iterrows()
            if _status_allows_convert_to_project((r.get("Status") or "").strip())
        ]
        if convertible:
            st.subheader("הפוך לפרויקט חדש")
            convert_options = []
            for _, r in convertible:
                row_dict = r.to_dict() if hasattr(r, "to_dict") else dict(r)
                v = (row_dict.get("Version") or "").strip() or "1"
                lbl = f"{row_dict.get('Date', '')} | {row_dict.get('Client', '')} | {row_dict.get('Project', '')} | גרסה: {v}"
                row_key = _quote_log_row_unique_key(row_dict)
                convert_options.append((row_key, lbl, row_dict))
            convert_labels = {k: lbl for k, lbl, _ in convert_options}
            selected_convert_key = st.selectbox(
                "בחר הצעה מאושרת להמרה לפרויקט",
                options=[k for k, _, _ in convert_options],
                format_func=lambda k: convert_labels.get(k, str(k)),
                key="convert_quote_select",
            )
            convert_clicked = st.button(
                "הפוך לפרויקט חדש 🚀",
                key="convert_quote_btn",
                type="primary",
            )
            if convert_clicked and selected_convert_key is not None:
                sel = next(
                    (rd for k, _, rd in convert_options if k == selected_convert_key),
                    None,
                )
                if sel is not None:
                    client_c = (sel.get("Client") or "").strip()
                    project_c = (sel.get("Project") or "").strip()
                    version_c = (sel.get("Version") or "").strip() or "V1"
                    with st.spinner("מייצר פרויקט ותיקיות דרופבוקס..."):
                        ok, err, _links_tuple, dropbox_failed = convert_quote_to_project(
                            client_c,
                            project_c,
                            version_c,
                            log_row=sel,
                        )
                    if ok:
                        if dropbox_failed:
                            st.warning(
                                "תיקיות דרופבוקס לא נוצרו עקב שגיאת התחברות (למשל טוקן פג תוקף). "
                                "הסטטוס עודכן ל'הומר לפרויקט' והפרויקט נוסף לרשימה."
                            )
                        else:
                            st.success("הפרויקט הוקם בהצלחה. הסטטוס עודכן ל'הומר לפרויקט'.")
                        st.session_state["quote_mgmt_kickoff_default_key"] = selected_convert_key
                        time.sleep(1.5)
                        st.rerun()
                    else:
                        st.error(f"שגיאה בהמרה: {err}")

        # --- התנעת פרויקט (הצעות במצב 'הומר לפרויקט') ---
        converted_kickoff = []
        for _, r in edited_df.iterrows():
            row_dict = r.to_dict() if hasattr(r, "to_dict") else dict(r)
            if (row_dict.get("Status") or "").strip() == "הומר לפרויקט":
                row_key = _quote_log_row_unique_key(row_dict)
                v = (row_dict.get("Version") or "").strip() or "1"
                lbl = f"{row_dict.get('Date', '')} | {row_dict.get('Client', '')} | {row_dict.get('Project', '')} | גרסה: {v}"
                converted_kickoff.append((row_key, lbl, row_dict))

        if converted_kickoff:
            st.divider()
            st.markdown("### 🚀 התנעת פרויקט וחלוקת משימות לצוות")
            kickoff_labels = {k: lbl for k, lbl, _ in converted_kickoff}
            kickoff_keys = [k for k, _, _ in converted_kickoff]
            default_ix = 0
            dk_pending = st.session_state.pop("quote_mgmt_kickoff_default_key", None)
            if dk_pending is not None and dk_pending in kickoff_keys:
                default_ix = kickoff_keys.index(dk_pending)
            default_ix = min(default_ix, max(0, len(kickoff_keys) - 1))
            selected_kickoff_key = st.selectbox(
                "בחר הצעה לתיאום התנעה והודעה לצוות",
                options=kickoff_keys,
                format_func=lambda k: kickoff_labels.get(k, str(k)),
                index=default_ix,
                key="quote_mgmt_kickoff_select",
            )
            sel_k = next((rd for k, _, rd in converted_kickoff if k == selected_kickoff_key), None)
            if sel_k is not None:
                k_client = (sel_k.get("Client") or "").strip()
                k_project = (sel_k.get("Project") or "").strip()
                k_main, k_upload, k_del = find_project_dropbox_links_for_client(k_client, k_project)
                st.markdown("**קישורי דרופבוקס**")
                if (k_main and k_main.startswith("http")) or (k_upload and k_upload.startswith("http")) or (
                    k_del and k_del.startswith("http")
                ):
                    col_k1, col_k2, col_k3 = st.columns(3)
                    if k_main and k_main.startswith("http"):
                        col_k1.link_button("📂 תיקייה ראשית", k_main)
                    if k_upload and k_upload.startswith("http"):
                        col_k2.link_button("📥 בקשת חומרים", k_upload)
                    if k_del and k_del.startswith("http"):
                        col_k3.link_button("📤 תיקיית תוצרים", k_del)
                else:
                    st.caption("לא נוצרו קישורי דרופבוקס לפרויקט זה (או שנכשלו בעת ההמרה).")

                kickoff_recipients = st.multiselect(
                    "נמעני מייל (צוות)",
                    options=list(TEAM_EMAILS.keys()),
                    default=[],
                    key="quote_mgmt_kickoff_recipients",
                )
                kickoff_brief = st.text_area(
                    "בריף / הערות למשימה",
                    key="quote_mgmt_kickoff_brief",
                    placeholder="הערות לצוות...",
                )
                if st.button("שלח מייל התנעה לצוות", key="quote_mgmt_kickoff_send", type="primary"):
                    actual_emails_list = []
                    seen = set()
                    for name in kickoff_recipients or []:
                        em = TEAM_EMAILS.get(name)
                        if em and str(em).strip():
                            addr = str(em).strip()
                            if addr not in seen:
                                seen.add(addr)
                                actual_emails_list.append(addr)
                    if not actual_emails_list:
                        st.warning(
                            "לא נמצאו כתובות מייל: בחרו חברי צוות שמופיעים במילון TEAM_EMAILS, או הוסיפו את השם והכתובת שם."
                        )
                    else:
                        deadline_k = date.today().strftime("%d/%m/%Y")
                        send_project_kickoff_email_eran(
                            actual_emails_list,
                            k_project,
                            k_client,
                            deadline_k,
                            k_main,
                            k_upload,
                            k_del,
                            kickoff_brief or "",
                        )

        # --- בחירת הצעה למחיקה ---
        st.subheader("מחיקת הצעה")
        delete_options = []
        for idx, r in edited_df.iterrows():
            version_val = (r.get("Version") or "").strip() or "1"
            label = f"{r.get('Date','')} | {r.get('Client','')} | {r.get('Project','')} | גרסה: {version_val}"
            delete_options.append((idx, label))
        delete_idx_to_label = {i: label for i, label in delete_options}
        selected_delete_idx = st.selectbox(
            "בחר הצעה למחיקה",
            options=[i for i, _ in delete_options],
            format_func=lambda i: delete_idx_to_label.get(i, str(i)),
            key="delete_quote_select",
        )
        delete_clicked = st.button(
            "🗑️ מחק הצעה",
            key="delete_quote_btn",
            type="secondary",
        )
        if delete_clicked and selected_delete_idx is not None:
            sel_row = edited_df.loc[selected_delete_idx]
            if isinstance(sel_row, pd.Series):
                sel_row = sel_row.to_dict()
            else:
                sel_row = dict(sel_row)
            date_val = (sel_row.get("Date") or "").strip()
            client_val = (sel_row.get("Client") or "").strip()
            project_val = (sel_row.get("Project") or "").strip()
            version_val = (sel_row.get("Version") or "").strip() or "1"
            file_path_val = (sel_row.get("File Path") or "").strip()
            try:
                st.warning("הקבצים יועברו לתיקיית Quotes/Trash (סל המחזור). ניתן לשחזר אותם משם.")
                _delete_proposal_and_move_to_trash(
                    date_val, client_val, project_val, version_val, file_path_val
                )
                updated_rows = [
                    r for r in rows
                    if (r.get("Date") or "").strip() != date_val
                    or (r.get("Client") or "").strip() != client_val
                    or (r.get("Project") or "").strip() != project_val
                    or ((r.get("Version") or "").strip() or "1") != version_val
                ]
                write_quotes_log(updated_rows)
                st.success("ההצעה נמחקה והקבצים הועברו לסל המחזור.")
                st.rerun()
            except Exception as e:
                st.error(f"שגיאה במחיקת ההצעה: {e}")

        save_clicked = st.button("שמור סטטוסים", type="primary")
        if save_clicked:
            # מיזוג השינויים מהטבלה המסוננת חזרה לרשימה המלאה
            edited_by_key = {
                (r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version")): r
                for r in edited_df.to_dict(orient="records")
            }
            updated_rows = []
            for r in rows:
                k = (r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version"))
                updated_rows.append(edited_by_key.get(k, r))

            original_by_key = {
                (r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version")): r
                for r in rows
            }
            approved_new = []
            for r in updated_rows:
                k = (r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version"))
                old_status = (original_by_key.get(k, {}) or {}).get("Status") or DEFAULT_QUOTE_STATUS
                new_status = (r.get("Status") or DEFAULT_QUOTE_STATUS).strip()
                if new_status not in ALLOWED_QUOTE_STATUSES:
                    new_status = DEFAULT_QUOTE_STATUS
                    r["Status"] = new_status
                if new_status == "Approved" and old_status != "Approved":
                    approved_new.append((r.get("Client") or "", r.get("Project") or ""))
                # העברת קבצים לתיקייה לפי סטטוס (Quotes/{Year}/{Month}/{New_Status})
                if old_status != new_status:
                    new_path = _move_quote_files_to_status_folder(
                        r.get("File Path") or "", old_status, new_status,
                        date_val=r.get("Date") or ""
                    )
                    if new_path:
                        r["File Path"] = new_path

            write_quotes_log(updated_rows)

            created_paths = []
            for client, project in approved_new:
                if client.strip() and project.strip():
                    try:
                        created_paths.append(
                            str(ensure_project_folders_for_approved_quote(client, project).resolve())
                        )
                    except Exception as e:
                        st.warning(f"שגיאה ביצירת תיקיית פרויקט עבור Approved ({client}/{project}): {e}")

            st.success("הסטטוסים נשמרו בהצלחה.")
            if created_paths:
                st.caption("נוצרו/אושרו תיקיות פרויקט עבור הצעות Approved:")
                for p in created_paths:
                    st.code(p)

        # --- העלאת חוזה חתום לכל שורה (Approved/Sent) ---
        st.subheader("📎 העלאת חוזה חתום (Signed Quote)")
        signed_col = "Signed File Path"
        has_signed_path = edited_df[signed_col].fillna("").astype(str).str.strip() != ""
        upload_rows = edited_df[
            (edited_df["Status"].isin(["Approved", "Sent"]))
            | ((edited_df["Status"] == "Signed") & has_signed_path)
        ]
        if upload_rows.empty:
            st.caption("אין הצעות במצב Approved או Sent להעלאת חוזה חתום.")
        else:
            for idx, row in upload_rows.iterrows():
                status_val = (row.get("Status") or "").strip()
                signed_path_val = (row.get("Signed File Path") or "").strip()
                date_val = (row.get("Date") or "").strip()
                client_val = (row.get("Client") or "").strip()
                project_val = (row.get("Project") or "").strip()
                version_val = (row.get("Version") or "").strip() or "1"
                quote_key = f"{date_val}_{client_val}_{project_val}_{version_val}".replace(" ", "_").replace("/", "-")

                with st.container():
                    col_info, col_action = st.columns([3, 1])
                    with col_info:
                        st.caption(f"{date_val} | {client_val} | {project_val} | גרסה: {version_val}")
                    with col_action:
                        if status_val == "Signed" and signed_path_val:
                            st.success("✅ חוזה חתום קיים")
                            sp_signed = Path(signed_path_val)
                            if sp_signed.is_file():
                                try:
                                    _sig_bytes = sp_signed.read_bytes()
                                    _mime = (
                                        "application/pdf"
                                        if sp_signed.suffix.lower() == ".pdf"
                                        else "application/octet-stream"
                                    )
                                    st.download_button(
                                        label="📥 הורד חוזה חתום",
                                        data=_sig_bytes,
                                        file_name=sp_signed.name,
                                        mime=_mime,
                                        key=f"dl_signed_{quote_key}",
                                    )
                                except OSError:
                                    st.warning("הקובץ לא נמצא בנתיב השמור.")
                            else:
                                st.warning("הקובץ לא נמצא בנתיב השמור.")
                        elif status_val in ("Approved", "Sent"):
                            uploaded_file = st.file_uploader(
                                "📎 העלה חוזה חתום",
                                type=["pdf", "jpg", "jpeg", "png"],
                                key=f"signed_upload_{quote_key}",
                            )
                            if uploaded_file is not None:
                                try:
                                    if not client_val or not project_val:
                                        st.warning("נתוני הצעה חסרים (לקוח/פרויקט).")
                                    else:
                                        safe_client = sanitize_filename_part(client_val)
                                        safe_project = sanitize_filename_part(project_val)
                                        parsed_dt = _parse_edit_date(date_val) if date_val else None
                                        parsed_dt = parsed_dt or date.today()
                                        year, month = parsed_dt.year, parsed_dt.month
                                        signed_dir = QUOTES_ROOT / str(year) / f"{month:02d}" / "Signed"
                                        signed_dir.mkdir(parents=True, exist_ok=True)
                                        original_suffix = Path(uploaded_file.name).suffix.lower() or ".pdf"
                                        filename = f"Signed_{safe_client}_{safe_project}{original_suffix}"
                                        signed_path = signed_dir / filename
                                        file_bytes = uploaded_file.read()
                                        with signed_path.open("wb") as f:
                                            f.write(file_bytes)
                                        rows_all = read_quotes_log()
                                        for r in rows_all:
                                            if (
                                                (r.get("Date") or "").strip() == date_val
                                                and (r.get("Client") or "").strip() == client_val
                                                and (r.get("Project") or "").strip() == project_val
                                                and (r.get("Version") or "").strip() == version_val
                                            ):
                                                r["Status"] = "Signed"
                                                r["Signed File Path"] = str(signed_path.resolve())
                                                break
                                        write_quotes_log(rows_all)
                                        st.success("החוזה החתום נשמר והסטטוס עודכן!")
                                        st.rerun()
                                except Exception as e:
                                    st.error(f"שגיאה בשמירת הקובץ: {e}")

        st.divider()
        st.subheader("פתיחת פרויקט מהצעה מאושרת")

        approved_df = edited_df[edited_df["Status"].isin(["Approved", "Signed"])]
        if approved_df.empty:
            st.caption("אין הצעות במצב Approved או Signed לפתיחת פרויקט.")
        else:
            for idx, row in approved_df.iterrows():
                client = (row.get("Client") or "").strip()
                project = (row.get("Project") or "").strip()
                version = (row.get("Version") or "").strip()

                # --- מרכז התנעה (Kickoff) ---
                kickoff_key = f"kickoff_{idx}"
                with st.expander(f"🚀 הזנק פרויקט (Kickoff) - {client} | {project}", expanded=False):
                    safe_client = sanitize_filename_part(client)
                    safe_project = sanitize_filename_part(project)
                    local_project_path = PROJECTS_ROOT / safe_client / safe_project
                    materials_path = local_project_path / "חומרים מהלקוח"
                    if not local_project_path.exists():
                        try:
                            local_project_path.mkdir(parents=True, exist_ok=True)
                            if not materials_path.exists():
                                materials_path.mkdir(parents=True, exist_ok=True)
                            st.success("נוצרו תיקיות פרויקט + חומרים מהלקוח")
                        except Exception as e:
                            st.error(f"לא הצלחתי ליצור תיקייה: {e}")

                    # קריאת אנשי קשר מגיליון contacts בגוגל שיטס
                    contacts_list: list[str] = []
                    try:
                        df_contacts = read_contacts_sheet()
                        if not df_contacts.empty:
                            for _, c in df_contacts.iterrows():
                                name = (c.get("שם מלא") or "").strip()
                                company = (c.get("חברה / משרד אדריכלים") or "").strip()
                                contact_type = (c.get("סוג איש קשר") or "").strip()
                                if name:
                                    parts = [name]
                                    if company:
                                        parts.append(company)
                                    suffix = f" ({contact_type})" if contact_type else ""
                                    contacts_list.append(f"{' - '.join(parts)}{suffix}")
                    except Exception:
                        contacts_list = []

                    assigned_team = st.multiselect(
                        "בחר עובדים לפרויקט",
                        options=TEAM_DISPLAY_NAMES,
                        default=TEAM_DISPLAY_NAMES,
                        key=f"kickoff_team_{kickoff_key}",
                    )
                    project_contacts = st.multiselect(
                        "בחר אנשי קשר לפרויקט (אדריכל, לקוח, הנהלת חשבונות)",
                        options=contacts_list,
                        default=[],
                        key=f"kickoff_contacts_{kickoff_key}",
                    )
                    project_template = st.multiselect(
                        "בחר שלבי עבודה אוטומטיים",
                        options=PROJECT_TEMPLATE_OPTIONS,
                        default=PROJECT_TEMPLATE_OPTIONS[:2],
                        key=f"kickoff_template_{kickoff_key}",
                    )
                    task_deadline = st.date_input(
                        "תאריך יעד לסקיצות ראשוניות",
                        value=date.today() + timedelta(days=14),
                        key=f"kickoff_task_deadline_{kickoff_key}",
                    )
                    personal_message = st.text_area(
                        "הודעה אישית (אופציונלי)",
                        key=f"kickoff_message_{kickoff_key}",
                        placeholder="הערות נוספות למייל...",
                    )

                    # הוספה ל-projects.csv (פרויקטים פעילים)
                    st.caption("הוספה לרשימת הפרויקטים הפעילים")
                    col_deadline, col_budget = st.columns(2)
                    with col_deadline:
                        default_deadline = date.today() + timedelta(days=30)
                        kickoff_deadline = st.date_input("דדליין לסיום", value=default_deadline, key=f"kickoff_deadline_{kickoff_key}")
                    with col_budget:
                        kickoff_budget = st.text_input("תקציב שעות (אופציונלי)", key=f"kickoff_budget_{kickoff_key}", placeholder="")
                    if st.button("הזנק פרויקט 🚀", key=f"kickoff_add_projects_{kickoff_key}", type="primary"):
                        exists_csv = _project_exists_in_projects_csv(client, project)
                        exists_db = _project_exists_in_projects(client, project)
                        if exists_db:
                            # הפרויקט כבר ב-projects – עדכן סטטוס ל'בעבודה' כדי שיופיע במוניטור וב-Task Board
                            _ensure_project_active_in_projects(client, project, status="בעבודה")
                            time.sleep(1.5)
                            st.success("הסטטוס עודכן ל'בעבודה'. הפרויקט יופיע במוניטור וברשימת המשימות.")
                            st.rerun()
                        elif exists_csv:
                            # הפרויקט ב-projects.csv בלבד – הוסף ל-projects כדי שיופיע במוניטור וב-Task Board
                            row_dict = row.to_dict() if hasattr(row, "to_dict") else dict(row)
                            budget_amt = _extract_total_from_quote_row(row_dict)
                            today_str = date.today().strftime("%d/%m/%Y")
                            manager_default = PROJECT_MANAGERS[0] if PROJECT_MANAGERS else ""
                            main_link, upload_link, deliverables_link = "", "", ""
                            try:
                                with st.spinner('מייצר מבנה תיקיות ב-Dropbox ופותח פרויקט...'):
                                    result = create_studio_dropbox_structure(project)
                                    if result:
                                        main_link, upload_link, deliverables_link = result
                            except Exception as e:
                                st.error(f'🚨 שגיאת דרופבוקס: {e}')
                            append_project_record(
                                client=client,
                                project_name=project,
                                manager=manager_default,
                                team_members=assigned_team or [],
                                status="בעבודה",
                                start_date_str=today_str,
                                budget_amount=budget_amt,
                                dropbox_main=main_link,
                                dropbox_upload=upload_link,
                                dropbox_deliverables=deliverables_link,
                                skip_rerun=True,
                            )
                            time.sleep(1.5)
                            st.session_state["kickoff_success_project"] = f"{client}|{project}"
                            st.session_state["kickoff_success_links"] = (main_link, upload_link, deliverables_link)
                            deadline_str_csv = kickoff_deadline.strftime('%d/%m/%Y')
                            send_kickoff_email(project, client, deadline_str_csv, main_link, upload_link, deliverables_link)
                            st.success('✅ התיקיות נוצרו ומייל התנעה עם הלינקים נשלח לצוות!')
                            col1, col2, col3 = st.columns(3)
                            if main_link and main_link.startswith('http'):
                                col1.link_button("📂 תיקייה ראשית", main_link)
                            if upload_link and upload_link.startswith('http'):
                                col2.link_button("📥 בקשת חומרים", upload_link)
                            if deliverables_link and deliverables_link.startswith('http'):
                                col3.link_button("📤 תיקיית תוצרים", deliverables_link)
                            time.sleep(8)
                            st.rerun()
                        else:
                            team_str = ", ".join(assigned_team) if assigned_team else ""
                            contacts_str = ", ".join(project_contacts) if project_contacts else ""
                            deadline_str = kickoff_deadline.strftime('%d/%m/%Y')
                            row_dict = row.to_dict() if hasattr(row, "to_dict") else dict(row)
                            budget_amt = _extract_total_from_quote_row(row_dict)
                            today_str = date.today().strftime("%d/%m/%Y")
                            manager_default = PROJECT_MANAGERS[0] if PROJECT_MANAGERS else ""
                            main_link, upload_link, deliverables_link = "", "", ""
                            try:
                                with st.spinner('מייצר מבנה תיקיות ב-Dropbox ופותח פרויקט...'):
                                    result = create_studio_dropbox_structure(project)
                                    if result:
                                        main_link, upload_link, deliverables_link = result
                            except Exception as e:
                                st.error(f'🚨 שגיאת דרופבוקס: {e}')
                            append_to_projects_csv(client, project, deadline_str, team_str, kickoff_budget, budget_amt, project_contacts=contacts_str, skip_rerun=True)
                            # הוספה ל-projects.csv עם סטטוס 'בעבודה' – זהה למה שהמוניטור וה-Task Board מחפשים
                            append_project_record(
                                client=client,
                                project_name=project,
                                manager=manager_default,
                                team_members=assigned_team or [],
                                status="בעבודה",
                                start_date_str=today_str,
                                budget_amount=budget_amt,
                                dropbox_main=main_link,
                                dropbox_upload=upload_link,
                                dropbox_deliverables=deliverables_link,
                                skip_rerun=True,
                            )
                            project_display = f"{client} | {project}"
                            if assigned_team and project_template:
                                append_kickoff_tasks_to_csv(
                                    project_display=project_display,
                                    assigned_team=assigned_team,
                                    project_template=project_template,
                                    task_deadline=task_deadline,
                                    skip_rerun=True,
                                )
                            time.sleep(1.5)
                            st.session_state["kickoff_success_project"] = f"{client}|{project}"
                            st.session_state["kickoff_success_links"] = (main_link, upload_link, deliverables_link)
                            send_kickoff_email(project, client, deadline_str, main_link, upload_link, deliverables_link)
                            st.success('✅ התיקיות נוצרו ומייל התנעה עם הלינקים נשלח לצוות!')
                            col1, col2, col3 = st.columns(3)
                            if main_link and main_link.startswith('http'):
                                col1.link_button("📂 תיקייה ראשית", main_link)
                            if upload_link and upload_link.startswith('http'):
                                col2.link_button("📥 בקשת חומרים", upload_link)
                            if deliverables_link and deliverables_link.startswith('http'):
                                col3.link_button("📤 תיקיית תוצרים", deliverables_link)
                            time.sleep(8)
                            st.rerun()

                    # מייל לצוות - תוכן אוניברסלי (שם תיקייה + לינק דרופבוקס, בלי נתיבים מקומיים)
                    team_names_str = ", ".join(assigned_team) if assigned_team else "צוות"
                    emails = [TEAM_EMAIL_BY_SHORT.get(name, "") for name in assigned_team]
                    to_str = ",".join(emails) if emails else ""
                    subject = f"פרויקט חדש: {project} - {client}"
                    kickoff_project_key = f"{client}|{project}"
                    stored_links = st.session_state.get("kickoff_success_links") if st.session_state.get("kickoff_success_project") == kickoff_project_key else None
                    main_link_body = stored_links[0] if stored_links and len(stored_links) >= 1 else ""
                    upload_link_body = stored_links[1] if stored_links and len(stored_links) >= 2 else ""
                    body = f"""היי {team_names_str},
נכנס פרויקט חדש: {project}
לקוח: {client}

📂 שם התיקייה בדרופבוקס: Projects / {client} / {project}"""
                    if main_link_body and main_link_body.startswith("http"):
                        body += f"\n\n📂 תיקיית פרויקט בדרופבוקס: {main_link_body}"
                    if upload_link_body and upload_link_body.startswith("http"):
                        body += f"\n📥 לינק לשליחה ללקוח להעלאת חומרים: {upload_link_body}"
                    if not (main_link_body and main_link_body.startswith("http")) and not (upload_link_body and upload_link_body.startswith("http")):
                        body += "\n\n☁️ לינק ישיר לחומרים: (נוצר אוטומטית בעת הוספת הפרויקט)"
                    if project_template:
                        body += "\n\n📋 שלבי עבודה שנבחרו:\n" + "\n".join(f"• {step}" for step in project_template)
                    body += "\n\nבהצלחה!"
                    if (personal_message or "").strip():
                        body += f"\n\nהערה: {personal_message.strip()}"

                    st.text_area("תוכן המייל (להעתקה)", value=body, height=200, key=f"kickoff_email_preview_{kickoff_key}", disabled=True)

                    # כפתור 1: Outlook (mailto)
                    mailto_url = f"mailto:{quote(to_str, safe=',@.+-_')}?subject={quote(subject)}&body={quote(body)}"
                    # כפתור 2: Gmail (פורמט דפדפן)
                    gmail_url = f"https://mail.google.com/mail/?view=cm&fs=1&to={quote(to_str, safe=',@.+-_')}&su={quote(subject)}&body={quote(body)}"
                    col_outlook, col_gmail = st.columns(2)
                    with col_outlook:
                        st.markdown(
                            f'<a href="{mailto_url}" target="_blank" style="display:inline-block;padding:8px 16px;background:#0078D4;color:white;text-decoration:none;border-radius:4px;">✉️ פתח טיוטה ב-Outlook</a>',
                            unsafe_allow_html=True,
                        )
                    with col_gmail:
                        st.markdown(
                            f'<a href="{gmail_url}" target="_blank" style="display:inline-block;padding:8px 16px;background:#EA4335;color:white;text-decoration:none;border-radius:4px;">📧 פתח טיוטה ב-Gmail</a>',
                            unsafe_allow_html=True,
                        )

                    # כפתור וואצאפ לערן
                    wa_links = ""
                    if main_link_body:
                        wa_links += f" תיקייה: {main_link_body}"
                    if upload_link_body:
                        wa_links += f" העלאה: {upload_link_body}"
                    wa_message = f"היי ערן, נכנס פרויקט חדש: {project} עבור {client}.{wa_links if wa_links else ' הלינק לחומרים נשלח במייל.'}"
                    wa_link = f"https://wa.me/{WHATSAPP_ERAN}?text={quote(wa_message)}"
                    st.markdown(
                        f'<a href="{wa_link}" target="_blank" style="display:inline-block;padding:8px 16px;background:#25D366;color:white;text-decoration:none;border-radius:4px;">📱 וואצאפ לערן</a>',
                        unsafe_allow_html=True,
                    )

        # הורדת PDF — קובץ פיזי בשרת בלבד (סביבת Web)
        st.divider()
        st.subheader("פתיחת קובץ PDF")
        st.caption("הורדת עותק PDF מהשרת אם הקובץ קיים בתיקיית המערכת (Quotes / temp_proposals).")
        options = []
        idx_to_row = {}
        for i, r in enumerate(rows):
            label = f"{r.get('Client','')} | {r.get('Project','')} | {r.get('Version','')} | {r.get('Date','')}"
            options.append((i, label))
            idx_to_row[i] = r

        selected_idx = st.selectbox(
            "בחר הצעה להורדה",
            options=[i for i, _ in options],
            format_func=lambda i: dict(options).get(i, str(i)),
        )
        r_pdf = idx_to_row.get(selected_idx, {})
        pdf_path_dl = find_physical_quote_pdf(r_pdf)
        if pdf_path_dl and pdf_path_dl.is_file():
            try:
                pdf_bytes = pdf_path_dl.read_bytes()
                st.download_button(
                    label="📥 הורד PDF",
                    data=pdf_bytes,
                    file_name=pdf_path_dl.name,
                    mime="application/pdf",
                    key=f"quote_mgmt_dl_pdf_{selected_idx}",
                )
            except OSError:
                st.error("קובץ ה-PDF הפיזי לא נמצא בתיקיית המערכת")
        else:
            st.error("קובץ ה-PDF הפיזי לא נמצא בתיקיית המערכת")

    except Exception:
        # Fallback without pandas (less control, but still works)
        search_term_fb = st.text_input(
            "🔍 חיפוש הצעה (לפי לקוח, פרויקט או תאריך)",
            key="quote_search_term_fallback",
            placeholder="הקלד לחיפוש...",
        )
        if search_term_fb and str(search_term_fb).strip():
            term_fb = str(search_term_fb).strip().lower()
            rows_filtered_fb = [
                r for r in rows
                if term_fb in (r.get("Client") or "").lower()
                or term_fb in (r.get("Project") or "").lower()
                or term_fb in (r.get("Date") or "")
            ]
        else:
            rows_filtered_fb = rows
        st.caption(f"מוצגות {len(rows_filtered_fb)} מתוך {len(rows)} הצעות")
        edited_rows_raw = st.data_editor(
            rows_filtered_fb,
            use_container_width=True,
            disabled=[c for c in QUOTES_LOG_COLUMNS if c != "Status"],
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=ALLOWED_QUOTE_STATUSES,
                    required=True,
                )
            },
            key="quotes_editor_fallback",
        )

        # --- בחירת הצעה למחיקה - fallback ---
        st.subheader("מחיקת הצעה")
        delete_options_fb = []
        for i, r in enumerate(rows_filtered_fb):
            version_val = (r.get("Version") or "").strip() or "1"
            label = f"{r.get('Date','')} | {r.get('Client','')} | {r.get('Project','')} | גרסה: {version_val}"
            delete_options_fb.append((i, label))
        delete_idx_to_label_fb = {i: label for i, label in delete_options_fb}
        selected_delete_idx_fb = st.selectbox(
            "בחר הצעה למחיקה",
            options=[i for i, _ in delete_options_fb],
            format_func=lambda i: delete_idx_to_label_fb.get(i, str(i)),
            key="delete_quote_select_fallback",
        )
        delete_fb = st.button(
            "🗑️ מחק הצעה",
            key="delete_quote_btn_fallback",
            type="secondary",
        )
        if delete_fb and selected_delete_idx_fb is not None:
            sel_row = rows_filtered_fb[selected_delete_idx_fb]
            date_val = (sel_row.get("Date") or "").strip()
            client_val = (sel_row.get("Client") or "").strip()
            project_val = (sel_row.get("Project") or "").strip()
            version_val = (sel_row.get("Version") or "").strip() or "1"
            file_path_val = (sel_row.get("File Path") or "").strip()
            try:
                st.warning("הקבצים יועברו לתיקיית Quotes/Trash (סל המחזור). ניתן לשחזר אותם משם.")
                _delete_proposal_and_move_to_trash(
                    date_val, client_val, project_val, version_val, file_path_val
                )
                updated_rows = [
                    r for r in rows
                    if (r.get("Date") or "").strip() != date_val
                    or (r.get("Client") or "").strip() != client_val
                    or (r.get("Project") or "").strip() != project_val
                    or ((r.get("Version") or "").strip() or "1") != version_val
                ]
                write_quotes_log(updated_rows)
                st.success("ההצעה נמחקה והקבצים הועברו לסל המחזור.")
                st.rerun()
            except Exception as e:
                st.error(f"שגיאה במחיקת ההצעה: {e}")

        save_clicked = st.button("שמור סטטוסים", type="primary")
        if save_clicked:
            edited_list = list(edited_rows_raw) if hasattr(edited_rows_raw, "__iter__") else []
            edited_by_key_fb = {
                (r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version")): r
                for r in edited_list
            }
            merged_rows = []
            original_by_key_fb = {(r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version")): r for r in rows}
            for r in rows:
                k = (r.get("Date"), r.get("Client"), r.get("Project"), r.get("Version"))
                merged = edited_by_key_fb.get(k, r)
                old_status = (original_by_key_fb.get(k) or {}).get("Status") or DEFAULT_QUOTE_STATUS
                new_status = (merged.get("Status") or DEFAULT_QUOTE_STATUS).strip()
                if old_status != new_status:
                    new_path = _move_quote_files_to_status_folder(
                        merged.get("File Path") or "", old_status, new_status,
                        date_val=merged.get("Date") or ""
                    )
                    if new_path:
                        merged["File Path"] = new_path
                merged_rows.append(merged)
            write_quotes_log(merged_rows)
            st.success("הסטטוסים נשמרו בהצלחה.")

        # --- הורדת PDF (fallback) ---
        st.divider()
        st.subheader("פתיחת קובץ PDF")
        st.caption("הורדת עותק PDF מהשרת אם הקובץ קיים בתיקיית המערכת (Quotes / temp_proposals).")
        options_fb = [(i, f"{r.get('Client','')} | {r.get('Project','')} | {r.get('Version','')} | {r.get('Date','')}") for i, r in enumerate(rows)]
        idx_to_label_fb = {i: lbl for i, lbl in options_fb}
        idx_to_row_fb = {i: r for i, r in enumerate(rows)}
        selected_fb = st.selectbox(
            "בחר הצעה להורדה",
            options=[i for i, _ in options_fb],
            format_func=lambda i: idx_to_label_fb.get(i, str(i)),
            key="open_pdf_select_fb",
        )
        r_pdf_fb = idx_to_row_fb.get(selected_fb, {})
        pdf_path_fb = find_physical_quote_pdf(r_pdf_fb)
        if pdf_path_fb and pdf_path_fb.is_file():
            try:
                pdf_bytes_fb = pdf_path_fb.read_bytes()
                st.download_button(
                    label="📥 הורד PDF",
                    data=pdf_bytes_fb,
                    file_name=pdf_path_fb.name,
                    mime="application/pdf",
                    key=f"quote_mgmt_dl_pdf_fb_{selected_fb}",
                )
            except OSError:
                st.error("קובץ ה-PDF הפיזי לא נמצא בתיקיית המערכת")
        else:
            st.error("קובץ ה-PDF הפיזי לא נמצא בתיקיית המערכת")

        # --- העלאת חוזה חתום (fallback ללא pandas) ---
        st.subheader("📎 העלאת חוזה חתום (Signed Quote)")
        fallback_upload_rows = [
            r for r in rows_filtered_fb
            if (r.get("Status") or "").strip() in ("Approved", "Sent")
            or ((r.get("Status") or "").strip() == "Signed" and (r.get("Signed File Path") or "").strip())
        ]
        if not fallback_upload_rows:
            st.caption("אין הצעות במצב Approved או Sent להעלאת חוזה חתום.")
        else:
            for r in fallback_upload_rows:
                date_val = (r.get("Date") or "").strip()
                client_val = (r.get("Client") or "").strip()
                project_val = (r.get("Project") or "").strip()
                version_val = (r.get("Version") or "").strip() or "1"
                status_val = (r.get("Status") or "").strip()
                signed_path_val = (r.get("Signed File Path") or "").strip()
                quote_key = f"{date_val}_{client_val}_{project_val}_{version_val}".replace(" ", "_").replace("/", "-")

                with st.container():
                    col_info, col_action = st.columns([3, 1])
                    with col_info:
                        st.caption(f"{date_val} | {client_val} | {project_val} | גרסה: {version_val}")
                    with col_action:
                        if status_val == "Signed" and signed_path_val:
                            st.success("✅ חוזה חתום קיים")
                            sp_signed_fb = Path(signed_path_val)
                            if sp_signed_fb.is_file():
                                try:
                                    _sig_b_fb = sp_signed_fb.read_bytes()
                                    _mime_fb = (
                                        "application/pdf"
                                        if sp_signed_fb.suffix.lower() == ".pdf"
                                        else "application/octet-stream"
                                    )
                                    st.download_button(
                                        label="📥 הורד חוזה חתום",
                                        data=_sig_b_fb,
                                        file_name=sp_signed_fb.name,
                                        mime=_mime_fb,
                                        key=f"dl_signed_fb_{quote_key}",
                                    )
                                except OSError:
                                    st.warning("הקובץ לא נמצא בנתיב השמור.")
                            else:
                                st.warning("הקובץ לא נמצא בנתיב השמור.")
                        elif status_val in ("Approved", "Sent"):
                            uploaded_file = st.file_uploader(
                                "📎 העלה חוזה חתום",
                                type=["pdf", "jpg", "jpeg", "png"],
                                key=f"signed_upload_fb_{quote_key}",
                            )
                            if uploaded_file is not None:
                                try:
                                    if not client_val or not project_val:
                                        st.warning("נתוני הצעה חסרים (לקוח/פרויקט).")
                                    else:
                                        safe_client = sanitize_filename_part(client_val)
                                        safe_project = sanitize_filename_part(project_val)
                                        parsed_dt = _parse_edit_date(date_val) if date_val else None
                                        parsed_dt = parsed_dt or date.today()
                                        year, month = parsed_dt.year, parsed_dt.month
                                        signed_dir = QUOTES_ROOT / str(year) / f"{month:02d}" / "Signed"
                                        signed_dir.mkdir(parents=True, exist_ok=True)
                                        original_suffix = Path(uploaded_file.name).suffix.lower() or ".pdf"
                                        filename = f"Signed_{safe_client}_{safe_project}{original_suffix}"
                                        signed_path = signed_dir / filename
                                        file_bytes = uploaded_file.read()
                                        with signed_path.open("wb") as f:
                                            f.write(file_bytes)
                                        rows_all = read_quotes_log()
                                        for row in rows_all:
                                            if (
                                                (row.get("Date") or "").strip() == date_val
                                                and (row.get("Client") or "").strip() == client_val
                                                and (row.get("Project") or "").strip() == project_val
                                                and (row.get("Version") or "").strip() == version_val
                                            ):
                                                row["Status"] = "Signed"
                                                row["Signed File Path"] = str(signed_path.resolve())
                                                break
                                        write_quotes_log(rows_all)
                                        st.success("החוזה החתום נשמר והסטטוס עודכן!")
                                        st.rerun()
                                except Exception as e:
                                    st.error(f"שגיאה בשמירת הקובץ: {e}")

        # --- מרכז התנעה (Kickoff) - fallback ---
        st.divider()
        st.subheader("🚀 מרכז התנעה (Kickoff)")
        fallback_kickoff_rows = [
            r for r in rows_filtered_fb
            if (r.get("Status") or "").strip() in ("Approved", "Signed")
        ]
        if not fallback_kickoff_rows:
            st.caption("אין הצעות במצב Approved או Signed למרכז התנעה.")
        else:
            for i, r in enumerate(fallback_kickoff_rows):
                date_val = (r.get("Date") or "").strip()
                client_val = (r.get("Client") or "").strip()
                project_val = (r.get("Project") or "").strip()
                version_val = (r.get("Version") or "").strip() or "1"
                kickoff_key_fb = f"kickoff_fb_{i}_{date_val}_{client_val}_{project_val}_{version_val}".replace(" ", "_").replace("/", "-")
                with st.expander(f"🚀 הזנק פרויקט (Kickoff) - {client_val} | {project_val}", expanded=False):
                    safe_client_fb = sanitize_filename_part(client_val)
                    safe_project_fb = sanitize_filename_part(project_val)
                    local_project_path_fb = PROJECTS_ROOT / safe_client_fb / safe_project_fb
                    materials_path_fb = local_project_path_fb / "חומרים מהלקוח"
                    if not local_project_path_fb.exists():
                        try:
                            local_project_path_fb.mkdir(parents=True, exist_ok=True)
                            if not materials_path_fb.exists():
                                materials_path_fb.mkdir(parents=True, exist_ok=True)
                            st.success("נוצרו תיקיות פרויקט + חומרים מהלקוח")
                        except Exception as e:
                            st.error(f"לא הצלחתי ליצור תיקייה: {e}")

                    # קריאת אנשי קשר מגיליון contacts בגוגל שיטס
                    contacts_list_fb: list[str] = []
                    try:
                        df_contacts_fb = read_contacts_sheet()
                        if not df_contacts_fb.empty:
                            for _, c in df_contacts_fb.iterrows():
                                name = (c.get("שם מלא") or "").strip()
                                company = (c.get("חברה / משרד אדריכלים") or "").strip()
                                contact_type = (c.get("סוג איש קשר") or "").strip()
                                if name:
                                    parts = [name]
                                    if company:
                                        parts.append(company)
                                    suffix = f" ({contact_type})" if contact_type else ""
                                    contacts_list_fb.append(f"{' - '.join(parts)}{suffix}")
                    except Exception:
                        contacts_list_fb = []

                    assigned_team_fb = st.multiselect(
                        "בחר עובדים לפרויקט",
                        options=TEAM_DISPLAY_NAMES,
                        default=TEAM_DISPLAY_NAMES,
                        key=f"kickoff_team_{kickoff_key_fb}",
                    )
                    project_contacts_fb = st.multiselect(
                        "בחר אנשי קשר לפרויקט (אדריכל, לקוח, הנהלת חשבונות)",
                        options=contacts_list_fb,
                        default=[],
                        key=f"kickoff_contacts_{kickoff_key_fb}",
                    )
                    project_template_fb = st.multiselect(
                        "בחר שלבי עבודה אוטומטיים",
                        options=PROJECT_TEMPLATE_OPTIONS,
                        default=PROJECT_TEMPLATE_OPTIONS[:2],
                        key=f"kickoff_template_{kickoff_key_fb}",
                    )
                    task_deadline_fb = st.date_input(
                        "תאריך יעד לסקיצות ראשוניות",
                        value=date.today() + timedelta(days=14),
                        key=f"kickoff_task_deadline_{kickoff_key_fb}",
                    )
                    personal_message = st.text_area(
                        "הודעה אישית (אופציונלי)",
                        key=f"kickoff_message_{kickoff_key_fb}",
                        placeholder="הערות נוספות למייל...",
                    )

                    # הוספה ל-projects.csv (פרויקטים פעילים)
                    st.caption("הוספה לרשימת הפרויקטים הפעילים")
                    col_deadline_fb, col_budget_fb = st.columns(2)
                    with col_deadline_fb:
                        default_deadline_fb = date.today() + timedelta(days=30)
                        kickoff_deadline_fb = st.date_input("דדליין לסיום", value=default_deadline_fb, key=f"kickoff_deadline_{kickoff_key_fb}")
                    with col_budget_fb:
                        kickoff_budget_fb = st.text_input("תקציב שעות (אופציונלי)", key=f"kickoff_budget_{kickoff_key_fb}", placeholder="")
                    if st.button("הזנק פרויקט 🚀", key=f"kickoff_add_projects_{kickoff_key_fb}", type="primary"):
                        exists_csv_fb = _project_exists_in_projects_csv(client_val, project_val)
                        exists_db_fb = _project_exists_in_projects(client_val, project_val)
                        if exists_db_fb:
                            # הפרויקט כבר ב-projects – עדכן סטטוס ל'בעבודה' כדי שיופיע במוניטור וב-Task Board
                            _ensure_project_active_in_projects(client_val, project_val, status="בעבודה")
                            time.sleep(1.5)
                            st.success("הסטטוס עודכן ל'בעבודה'. הפרויקט יופיע במוניטור וברשימת המשימות.")
                            st.rerun()
                        elif exists_csv_fb:
                            # הפרויקט ב-projects.csv בלבד – הוסף ל-projects כדי שיופיע במוניטור וב-Task Board
                            budget_amt_fb = _extract_total_from_quote_row(r)
                            today_str_fb = date.today().strftime("%d/%m/%Y")
                            manager_default_fb = PROJECT_MANAGERS[0] if PROJECT_MANAGERS else ""
                            main_link_fb, upload_link_fb, deliverables_link_fb = "", "", ""
                            try:
                                with st.spinner('מייצר מבנה תיקיות ב-Dropbox ופותח פרויקט...'):
                                    result_fb = create_studio_dropbox_structure(project_val)
                                    if result_fb:
                                        main_link_fb, upload_link_fb, deliverables_link_fb = result_fb
                            except Exception as e:
                                st.error(f'🚨 שגיאת דרופבוקס: {e}')
                            append_project_record(
                                client=client_val,
                                project_name=project_val,
                                manager=manager_default_fb,
                                team_members=assigned_team_fb or [],
                                status="בעבודה",
                                start_date_str=today_str_fb,
                                budget_amount=budget_amt_fb,
                                dropbox_main=main_link_fb,
                                dropbox_upload=upload_link_fb,
                                dropbox_deliverables=deliverables_link_fb,
                                skip_rerun=True,
                            )
                            time.sleep(1.5)
                            st.session_state["kickoff_success_project"] = f"{client_val}|{project_val}"
                            st.session_state["kickoff_success_links"] = (main_link_fb, upload_link_fb, deliverables_link_fb)
                            deadline_str_csv_fb = kickoff_deadline_fb.strftime('%d/%m/%Y')
                            send_kickoff_email(project_val, client_val, deadline_str_csv_fb, main_link_fb, upload_link_fb, deliverables_link_fb)
                            st.success('✅ התיקיות נוצרו ומייל התנעה עם הלינקים נשלח לצוות!')
                            col1_fb, col2_fb, col3_fb = st.columns(3)
                            if main_link_fb and main_link_fb.startswith('http'):
                                col1_fb.link_button("📂 תיקייה ראשית", main_link_fb)
                            if upload_link_fb and upload_link_fb.startswith('http'):
                                col2_fb.link_button("📥 בקשת חומרים", upload_link_fb)
                            if deliverables_link_fb and deliverables_link_fb.startswith('http'):
                                col3_fb.link_button("📤 תיקיית תוצרים", deliverables_link_fb)
                            time.sleep(8)
                            st.rerun()
                        else:
                            team_str_fb = ", ".join(assigned_team_fb) if assigned_team_fb else ""
                            contacts_str_fb = ", ".join(project_contacts_fb) if project_contacts_fb else ""
                            deadline_str_fb = kickoff_deadline_fb.strftime('%d/%m/%Y')
                            budget_amt_fb = _extract_total_from_quote_row(r)
                            today_str_fb = date.today().strftime("%d/%m/%Y")
                            manager_default_fb = PROJECT_MANAGERS[0] if PROJECT_MANAGERS else ""
                            main_link_fb, upload_link_fb, deliverables_link_fb = "", "", ""
                            try:
                                with st.spinner('מייצר מבנה תיקיות ב-Dropbox ופותח פרויקט...'):
                                    result_fb = create_studio_dropbox_structure(project_val)
                                    if result_fb:
                                        main_link_fb, upload_link_fb, deliverables_link_fb = result_fb
                            except Exception as e:
                                st.error(f'🚨 שגיאת דרופבוקס: {e}')
                            append_to_projects_csv(client_val, project_val, deadline_str_fb, team_str_fb, kickoff_budget_fb, budget_amt_fb, project_contacts=contacts_str_fb, skip_rerun=True)
                            # הוספה ל-projects.csv עם סטטוס 'בעבודה' – זהה למה שהמוניטור וה-Task Board מחפשים
                            append_project_record(
                                client=client_val,
                                project_name=project_val,
                                manager=manager_default_fb,
                                team_members=assigned_team_fb or [],
                                status="בעבודה",
                                start_date_str=today_str_fb,
                                budget_amount=budget_amt_fb,
                                dropbox_main=main_link_fb,
                                dropbox_upload=upload_link_fb,
                                dropbox_deliverables=deliverables_link_fb,
                                skip_rerun=True,
                            )
                            project_display_fb = f"{client_val} | {project_val}"
                            if assigned_team_fb and project_template_fb:
                                append_kickoff_tasks_to_csv(
                                    project_display=project_display_fb,
                                    assigned_team=assigned_team_fb,
                                    project_template=project_template_fb,
                                    task_deadline=task_deadline_fb,
                                    skip_rerun=True,
                                )
                            time.sleep(1.5)
                            st.session_state["kickoff_success_project"] = f"{client_val}|{project_val}"
                            st.session_state["kickoff_success_links"] = (main_link_fb, upload_link_fb, deliverables_link_fb)
                            send_kickoff_email(project_val, client_val, deadline_str_fb, main_link_fb, upload_link_fb, deliverables_link_fb)
                            st.success('✅ התיקיות נוצרו ומייל התנעה עם הלינקים נשלח לצוות!')
                            col1_fb, col2_fb, col3_fb = st.columns(3)
                            if main_link_fb and main_link_fb.startswith('http'):
                                col1_fb.link_button("📂 תיקייה ראשית", main_link_fb)
                            if upload_link_fb and upload_link_fb.startswith('http'):
                                col2_fb.link_button("📥 בקשת חומרים", upload_link_fb)
                            if deliverables_link_fb and deliverables_link_fb.startswith('http'):
                                col3_fb.link_button("📤 תיקיית תוצרים", deliverables_link_fb)
                            time.sleep(8)
                            st.rerun()

                    # מייל לצוות - תוכן אוניברסלי (שם תיקייה + לינק דרופבוקס, בלי נתיבים מקומיים)
                    team_names_str_fb = ", ".join(assigned_team_fb) if assigned_team_fb else "צוות"
                    emails = [TEAM_EMAIL_BY_SHORT.get(name, "") for name in assigned_team_fb]
                    to_str = ",".join(emails) if emails else ""
                    subject = f"פרויקט חדש: {project_val} - {client_val}"
                    kickoff_project_key_fb = f"{client_val}|{project_val}"
                    stored_links_fb = st.session_state.get("kickoff_success_links") if st.session_state.get("kickoff_success_project") == kickoff_project_key_fb else None
                    main_link_body_fb = stored_links_fb[0] if stored_links_fb and len(stored_links_fb) >= 1 else ""
                    upload_link_body_fb = stored_links_fb[1] if stored_links_fb and len(stored_links_fb) >= 2 else ""
                    body = f"""היי {team_names_str_fb},
נכנס פרויקט חדש: {project_val}
לקוח: {client_val}

📂 שם התיקייה בדרופבוקס: Projects / {client_val} / {project_val}"""
                    if main_link_body_fb and main_link_body_fb.startswith("http"):
                        body += f"\n\n📂 תיקיית פרויקט בדרופבוקס: {main_link_body_fb}"
                    if upload_link_body_fb and upload_link_body_fb.startswith("http"):
                        body += f"\n📥 לינק לשליחה ללקוח להעלאת חומרים: {upload_link_body_fb}"
                    if not (main_link_body_fb and main_link_body_fb.startswith("http")) and not (upload_link_body_fb and upload_link_body_fb.startswith("http")):
                        body += "\n\n☁️ לינק ישיר לחומרים: (נוצר אוטומטית בעת הוספת הפרויקט)"
                    if project_template_fb:
                        body += "\n\n📋 שלבי עבודה שנבחרו:\n" + "\n".join(f"• {step}" for step in project_template_fb)
                    body += "\n\nבהצלחה!"
                    if (personal_message or "").strip():
                        body += f"\n\nהערה: {personal_message.strip()}"

                    st.text_area("תוכן המייל (להעתקה)", value=body, height=200, key=f"kickoff_email_preview_{kickoff_key_fb}", disabled=True)

                    # כפתור 1: Outlook (mailto)
                    mailto_url = f"mailto:{quote(to_str, safe=',@.+-_')}?subject={quote(subject)}&body={quote(body)}"
                    # כפתור 2: Gmail (פורמט דפדפן)
                    gmail_url = f"https://mail.google.com/mail/?view=cm&fs=1&to={quote(to_str, safe=',@.+-_')}&su={quote(subject)}&body={quote(body)}"
                    col_outlook_fb, col_gmail_fb = st.columns(2)
                    with col_outlook_fb:
                        st.markdown(
                            f'<a href="{mailto_url}" target="_blank" style="display:inline-block;padding:8px 16px;background:#0078D4;color:white;text-decoration:none;border-radius:4px;">✉️ פתח טיוטה ב-Outlook</a>',
                            unsafe_allow_html=True,
                        )
                    with col_gmail_fb:
                        st.markdown(
                            f'<a href="{gmail_url}" target="_blank" style="display:inline-block;padding:8px 16px;background:#EA4335;color:white;text-decoration:none;border-radius:4px;">📧 פתח טיוטה ב-Gmail</a>',
                            unsafe_allow_html=True,
                        )

                    # כפתור וואצאפ לערן
                    wa_links_fb = ""
                    if main_link_body_fb:
                        wa_links_fb += f" תיקייה: {main_link_body_fb}"
                    if upload_link_body_fb:
                        wa_links_fb += f" העלאה: {upload_link_body_fb}"
                    wa_message_fb = f"היי ערן, נכנס פרויקט חדש: {project_val} עבור {client_val}.{wa_links_fb if wa_links_fb else ' הלינק לחומרים נשלח במייל.'}"
                    wa_link_fb = f"https://wa.me/{WHATSAPP_ERAN}?text={quote(wa_message_fb)}"
                    st.markdown(
                        f'<a href="{wa_link_fb}" target="_blank" style="display:inline-block;padding:8px 16px;background:#25D366;color:white;text-decoration:none;border-radius:4px;">📱 וואצאפ לערן</a>',
                        unsafe_allow_html=True,
                    )


# סטטוסים שנחשבים "פעילים" (לבחירת פרויקט להוספת משימות, אנשי קשר וכו')
ACTIVE_PROJECT_STATUSES = ("בעבודה", "ממתין להתחלה")


def _get_active_projects_options() -> list[str]:
    """Return list of 'Client | Project Name' for projects. זמנית: מציג את כל הפרויקטים (ללא סינון סטטוס)."""
    rows = read_projects()
    options = []
    # זמנית: הצגת כל הפרויקטים (ביטול סינון Status == Active)
    # active_normalized = [s.strip().lower() for s in ACTIVE_PROJECT_STATUSES]
    for r in rows:
        # status = (r.get("Status") or "").strip().lower()
        # if status in active_normalized:
        client = (r.get("Client") or "").strip()
        project_name = (r.get("Project Name") or "").strip()
        if client and project_name:
            options.append(f"{client} | {project_name}")
    return options


def _task_project_label_to_stored(label: str) -> str:
    """ממיר בחירה מה-selectbox לערך לשמירה בעמודת «פרויקט» בגיליון tasks."""
    s = (label or "").strip()
    if s == TASKS_PROJECT_GENERAL_STUDIO_SELECT or s == TASKS_PROJECT_GENERAL_STUDIO:
        return TASKS_PROJECT_GENERAL_STUDIO
    return s


def _task_project_stored_to_label(stored: str) -> str:
    """תווית ב-selectbox מתוך ערך שמור (תאימות לאחור לערך ללא אימוג'י)."""
    s = (stored or "").strip()
    if s == TASKS_PROJECT_GENERAL_STUDIO:
        return TASKS_PROJECT_GENERAL_STUDIO_SELECT
    return s


def _task_project_select_options() -> list[str]:
    """פרויקטים לטופס משימה: «כללי / סטודיו» ראשון, אחר כך רשימת הפרויקטים ממסד."""
    return [TASKS_PROJECT_GENERAL_STUDIO_SELECT] + _get_active_projects_options()


def _is_task_project_non_client_row(project_cell: str) -> bool:
    """משימה ללא פרויקט לקוח — לא לפרש כ-'לקוח | פרויקט' ולא למשוך נתוני הצעה/דרופבוקס."""
    s = (project_cell or "").strip()
    return s in (TASKS_PROJECT_GENERAL_STUDIO, TASKS_PROJECT_OOO_DEFAULT)


def _compute_project_monitor_stats() -> dict[str, int | float]:
    """Compute project counts and budget sums for sidebar monitor by status category."""
    rows = read_projects()
    active_statuses = ("בעבודה", "ממתין להתחלה")
    feedback_statuses = ("נשלח לסבב הערות 1", "נשלח לסבב הערות 2", "ממתין לאדריכל/לקוח")
    frozen_statuses = ("הוקפא",)
    completed_statuses = ("הסתיים", "חשבונית נשלחה", "שולם")

    stats = {
        "active": 0,
        "active_sum": 0.0,
        "feedback": 0,
        "feedback_sum": 0.0,
        "frozen": 0,
        "frozen_sum": 0.0,
        "completed": 0,
        "completed_sum": 0.0,
    }

    def _safe_amount(r: dict) -> float:
        val = r.get('היקף כספי (₪)')
        if val is None or (isinstance(val, str) and not str(val).strip()):
            return 0.0
        try:
            return float(str(val).replace(",", "").replace(" ", ""))
        except (TypeError, ValueError):
            return 0.0

    active_norm = [x.strip().lower() for x in active_statuses]
    feedback_norm = [x.strip().lower() for x in feedback_statuses]
    frozen_norm = [x.strip().lower() for x in frozen_statuses]
    completed_norm = [x.strip().lower() for x in completed_statuses]
    for r in rows:
        s = (r.get("Status") or "").strip()
        s_lower = s.lower()
        amt = _safe_amount(r)
        if s_lower in active_norm:
            stats["active"] += 1
            stats["active_sum"] += amt
        elif s_lower in feedback_norm:
            stats["feedback"] += 1
            stats["feedback_sum"] += amt
        elif s_lower in frozen_norm:
            stats["frozen"] += 1
            stats["frozen_sum"] += amt
        elif s_lower in completed_norm:
            stats["completed"] += 1
            stats["completed_sum"] += amt
    return stats


def _parse_task_date(date_str: str):
    """Parse date string (dd/mm/yyyy or yyyy-mm-dd) to datetime for plotly."""
    if not date_str or not str(date_str).strip():
        return None
    s = str(date_str).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _build_gantt_dataframe_for_timeline(
    active_tasks: list[dict],
    y_axis_mode: str,
) -> tuple[pd.DataFrame | None, str | None, str | None]:
    """
    מכין DataFrame ל-plotly.express.timeline.
    y_axis_mode: 'assignee' — אחראי (מפתח צבע); 'project' — פרויקט; 'task_detail' — פרויקט | משימה.
    מחזיר (df, y_column_name, error_message).
    """
    if not active_tasks:
        return None, None, "אין משימות"
    gantt_df = pd.DataFrame(active_tasks)
    if hasattr(gantt_df.columns, "str"):
        gantt_df.columns = gantt_df.columns.str.strip()
    end_col = "תאריך יעד" if "תאריך יעד" in gantt_df.columns else None
    if not end_col:
        return None, None, "חסרה עמודת תאריך יעד"
    # ציר X: תאריך התחלה (או תאריך יעד אם חסרה התחלה) → סיום בתאריך יעד
    gantt_df["_start_dt"] = gantt_df.apply(
        lambda r: (
            _parse_task_date(str(r.get("תאריך התחלה", "") or ""))
            or _parse_task_date(str(r.get("תאריך יעד", "") or ""))
            or datetime.today()
        ),
        axis=1,
    )
    gantt_df["_end_dt"] = gantt_df.apply(
        lambda r: (
            _parse_task_date(str(r.get("תאריך יעד", "") or ""))
            or _parse_task_date(str(r.get("תאריך התחלה", "") or ""))
            or datetime.today()
        ),
        axis=1,
    )
    mask = gantt_df["_start_dt"] > gantt_df["_end_dt"]
    gantt_df.loc[mask, "_end_dt"] = gantt_df.loc[mask, "_start_dt"] + timedelta(days=1)
    gantt_df = gantt_df[gantt_df["_end_dt"].notna()]
    if gantt_df.empty:
        return None, None, "אין שורות תקינות לגאנט"

    acol = "הוקצה ל" if "הוקצה ל" in gantt_df.columns else None
    if acol:
        gantt_df[acol] = gantt_df[acol].fillna("").astype(str).replace("nan", "")
        gantt_df["_assignee_color_key"] = gantt_df[acol].apply(_task_team_key_for_color)
        gantt_df["_task_color_key"] = gantt_df["_assignee_color_key"]
    else:
        gantt_df["_assignee_color_key"] = "(ללא אחראי)"
        gantt_df["_task_color_key"] = "(ללא אחראי)"

    if "שם משימה" in gantt_df.columns:
        _ooo_mask = gantt_df["שם משימה"].apply(_is_out_of_office_task)
        gantt_df.loc[_ooo_mask, "_task_color_key"] = "חופשה/היעדרות"

    if "סטטוס" in gantt_df.columns:
        gantt_df["_opacity"] = gantt_df["סטטוס"].apply(_gantt_opacity_for_status)
    else:
        gantt_df["_opacity"] = 1.0

    if y_axis_mode == "assignee":
        gantt_df["_y_gantt"] = gantt_df["_assignee_color_key"]
    elif y_axis_mode == "project":

        def _proj_y(row) -> str:
            if _is_out_of_office_task(row):
                return _ooo_event_title_from_row(row)
            _p = str(row.get("פרויקט", "") or "").strip()
            return _p if _p else "(ללא פרויקט)"

        gantt_df["_y_gantt"] = gantt_df.apply(_proj_y, axis=1)
    else:
        if "פרויקט" in gantt_df.columns and "שם משימה" in gantt_df.columns:

            def _task_detail_y(row) -> str:
                if _is_out_of_office_task(row):
                    return _ooo_event_title_from_row(row)
                return (
                    str(row.get("פרויקט", "") or "") + " | " + str(row.get("שם משימה", "") or "")
                ).strip(" |")

            gantt_df["_y_gantt"] = gantt_df.apply(_task_detail_y, axis=1)
        else:
            gantt_df["_y_gantt"] = gantt_df.get("שם משימה", pd.Series([""] * len(gantt_df))).fillna("")

    gantt_df["_gantt_bar_text"] = gantt_df.apply(
        lambda r: (
            _ooo_event_title_from_row(r)
            if _is_out_of_office_task(r)
            else str(r.get("_assignee_color_key", "") or "")
        ),
        axis=1,
    )

    return gantt_df, "_y_gantt", None


def _apply_gantt_bar_opacity(fig, gantt_df: pd.DataFrame, y_col: str) -> None:
    """שקיפות לפי עמודת _opacity (משימות בסטטוס 'הסתיים' מעומעמות)."""
    if "_opacity" not in gantt_df.columns:
        return
    for tr in fig.data:
        if getattr(tr, "type", None) != "bar":
            continue
        name = tr.name
        if name is None:
            continue
        sub = gantt_df[gantt_df["_task_color_key"].astype(str) == str(name)]
        if sub.empty:
            continue
        ys = tr.y
        if ys is None or len(ys) == 0:
            continue
        opacities: list[float] = []
        for y_val in ys:
            yv = str(y_val) if y_val is not None else ""
            match = sub[sub[y_col].astype(str) == yv]
            if match.empty:
                match = sub[sub[y_col] == y_val]
            if not match.empty:
                opacities.append(float(match["_opacity"].iloc[0]))
            else:
                opacities.append(1.0)
        if len(opacities) == len(ys):
            tr.marker.opacity = opacities


def _style_gantt_timeline_figure(fig, gantt_df: pd.DataFrame, y_col: str) -> None:
    """עיצוב דק לגאנט: רוחב פסים, גובה דינמי (עד תקרה), שקיפות למשימות שהושלמו."""
    fig.update_traces(width=0.4)
    bar_height = max(400, min(1200, len(gantt_df) * 24))
    fig.update_layout(height=bar_height)
    _apply_gantt_bar_opacity(fig, gantt_df, y_col)


def _signed_quote_rows_for_project_hub() -> list[dict]:
    """שורות quotes בסטטוסים פעילים (Signed / הומר לפרויקט) — שורה אחת לכל (Client, Project) לפי גרסה גבוהה ביותר."""
    rows = read_quotes_csv()
    if not rows:
        return []
    quotes_df = pd.DataFrame(rows)
    if quotes_df.empty or "Status" not in quotes_df.columns:
        return []
    active_projects = quotes_df[
        quotes_df["Status"].fillna("").astype(str).str.strip().isin(["Signed", "הומר לפרויקט"])
    ]
    best: dict[tuple[str, str], dict] = {}
    for _, row in active_projects.iterrows():
        r = {}
        for c in QUOTES_CSV_COLUMNS:
            v = row.get(c)
            if v is None or (isinstance(v, str) and v.strip().lower() in ("nan", "none")):
                r[c] = ""
            else:
                s = str(v).strip()
                r[c] = "" if s.lower() in ("nan", "none") else s
        c = (r.get("Client") or "").strip()
        p = (r.get("Project") or "").strip()
        if not p:
            continue
        key = (c, p)
        pv = parse_version_number(r.get("Version") or "")
        prev = best.get(key)
        if prev is None or pv > parse_version_number(prev.get("Version") or ""):
            best[key] = r
    return list(best.values())


def _kanban_stage_lookup_by_client_project() -> dict[tuple[str, str], str]:
    """מפת (Client, Project) לשלב קנבן מעודכן מהגיליון quotes (הצעות Signed / הומר לפרויקט)."""
    out: dict[tuple[str, str], str] = {}
    for r in _signed_quote_rows_for_project_hub():
        c = (r.get("Client") or "").strip()
        p = (r.get("Project") or "").strip()
        if not p:
            continue
        out[(c, p)] = _normalized_kanban_stage(r.get("שלב עבודה") or "")
    return out


def _normalized_kanban_stage(val: str) -> str:
    """מחזיר שלב קנבן תקין; ערך ריק → 'התקבל'; לא מוכר → 'אחר' (זרימה לא סטנדרטית)."""
    v = (val or "").strip()
    if not v:
        return KANBAN_WORK_STAGES[0]
    if v in KANBAN_WORK_STAGES:
        return v
    return "אחר"


def _kanban_work_stage_changed(key: str, client: str, project: str, version: str) -> None:
    """שמירת שלב עבודה לגיליון quotes לאחר שינוי ב-selectbox בלוח הקנבן."""
    new_stage = st.session_state.get(key)
    if new_stage is None:
        return
    rows = read_quotes_csv()
    qk = _quote_key(client, project, version)
    for i, r in enumerate(rows):
        if _quote_key(r.get("Client", ""), r.get("Project", ""), r.get("Version", "")) == qk:
            merged = {c: (r.get(c) or "") for c in QUOTES_CSV_COLUMNS}
            merged["שלב עבודה"] = new_stage
            rows[i] = merged
            write_quotes_csv(rows, skip_rerun=True)
            # לא לקרוא st.rerun() כאן — בתוך on_change זה no-op; Streamlit מריץ מחדש אחרי ה-callback.
            return


def _render_project_kanban_board() -> None:
    """
    לוח קנבן — פרויקטים פעילים מהצעות (Signed / הומר לפרויקט), עמודת 'שלב עבודה'.
    העמודה ב-QUOTES_CSV_COLUMNS; ערך ריק בגיליון מוצג כ-'התקבל' (_normalized_kanban_stage).
    """
    if not read_quotes_csv():
        st.info("אין נתוני הצעות ללוח הקנבן.")
        return
    active_rows = _signed_quote_rows_for_project_hub()
    if not active_rows:
        st.info(
            "אין פרויקטים פעילים (Signed / הומר לפרויקט) ללוח הקנבן."
        )
        return

    st.subheader("לוח שלבי עבודה (קנבן)")
    by_stage: dict[str, list[tuple[int, dict]]] = {s: [] for s in KANBAN_WORK_STAGES}
    for idx, row in enumerate(active_rows):
        stg = _normalized_kanban_stage(row.get("שלב עבודה") or "")
        by_stage[stg].append((idx, row))

    cols = st.columns(len(KANBAN_WORK_STAGES))
    for col_idx, stage in enumerate(KANBAN_WORK_STAGES):
        with cols[col_idx]:
            st.markdown(f"### {stage}")
            for idx, row in by_stage[stage]:
                client = (row.get("Client") or "").strip()
                project = (row.get("Project") or "").strip()
                version = (row.get("Version") or "").strip() or "V1"
                current = _normalized_kanban_stage(row.get("שלב עבודה") or "")
                key = f"kanban_work_stage_{idx}"
                with st.container(border=True):
                    st.markdown(f"**{project}**" if project else "**(ללא שם פרויקט)**")
                    st.caption(client if client else "—")
                    st.selectbox(
                        "שלב עבודה",
                        KANBAN_WORK_STAGES,
                        index=KANBAN_WORK_STAGES.index(current),
                        key=key,
                        label_visibility="collapsed",
                        on_change=_kanban_work_stage_changed,
                        args=(key, client, project, version),
                    )


def _find_project_row_for_hub(client: str, project_name: str) -> dict | None:
    """רשומת projects לפי לקוח ושם פרויקט (התאמה לשדות ב-quotes)."""
    c_key = (client or "").strip()
    p_key = (project_name or "").strip()
    for pr in read_projects():
        if (pr.get("Client") or "").strip() == c_key and (pr.get("Project Name") or "").strip() == p_key:
            return pr
    return None


def show_project_folders_page() -> None:
    """תיקי פרויקטים — מידע מרוכז מהצעה (Signed / הומר לפרויקט) וקישורי דרופבוקס ממסד הפרויקטים."""
    st.title("📁 תיקי פרויקטים (מידע וקשר)")
    st.caption("פרויקטים פעילים בגיליון ההצעות (חתום או הומר לפרויקט) — פרטי לקוח וקישורים לצוות.")

    active_rows = _signed_quote_rows_for_project_hub()
    if not active_rows:
        st.info(
            "אין פרויקטים בסטטוס 'Signed' או 'הומר לפרויקט' בגיליון ההצעות. "
            "עדכנו סטטוס הצעה בהתאם כדי שיופיעו כאן."
        )
        return

    def _hub_label(r: dict) -> str:
        c = (r.get("Client") or "").strip()
        p = (r.get("Project") or "").strip()
        return f"{p} — {c}" if c else p

    active_rows.sort(key=lambda r: _hub_label(r).lower())
    labels = [_hub_label(r) for r in active_rows]
    pick = st.selectbox(
        "בחר פרויקט פעיל",
        options=list(range(len(active_rows))),
        format_func=lambda i: labels[i],
        key="project_hub_pick",
    )
    row = active_rows[pick]
    client = (row.get("Client") or "").strip()
    project = (row.get("Project") or "").strip()

    client_email = (row.get("Client Email") or "").strip()
    contact_person = (row.get("Contact Person") or "").strip()
    architect_phone = (
        (
            row.get("Architect Contact")
            or row.get("Architect Phone")
            or row.get("טלפון אדריכל")
            or row.get("ArchitectPhone")
            or ""
        ).strip()
    )
    client_phone = (
        (row.get("Client Phone") or row.get("Phone") or row.get("טלפון לקוח") or row.get("טלפון") or "").strip()
    )

    with st.container(border=True):
        st.subheader("תעודת זהות — פרויקט")
        st.markdown(f"**שם פרויקט:** {project or '—'}")
        st.markdown(f"**לקוח:** {client or '—'}")
        st.markdown(f"**איש קשר (מההצעה):** {contact_person or '—'}")
        st.markdown(f"**אימייל לקוח:** {client_email or '—'}")
        st.markdown(f"**טלפון לקוח:** {client_phone or '—'}")
        st.markdown(f"**שם/טלפון אדריכל:** {architect_phone or '—'}")
        ver = (row.get("Version") or "").strip()
        st.markdown(f"**גרסת הצעה:** {ver or '—'}")

    with st.expander("✏️ עדכון פרטי פרויקט והערות", expanded=False):
        with st.form("hub_edit_quote_form"):
            f_client_email = st.text_input(
                "אימייל לקוח",
                value=str(row.get("Client Email", "") or ""),
                key=f"hub_f_client_email_{pick}",
            )
            f_client_phone = st.text_input(
                "טלפון לקוח",
                value=str(row.get("Client Phone", "") or ""),
                key=f"hub_f_client_phone_{pick}",
            )
            f_architect = st.text_input(
                "שם/טלפון אדריכל",
                value=str(row.get("Architect Contact", "") or ""),
                key=f"hub_f_architect_{pick}",
            )
            f_notes = st.text_area(
                "הערות פרויקט מיוחדות",
                value=str(row.get("Project Special Notes", "") or ""),
                height=120,
                key=f"hub_f_notes_{pick}",
            )
            submitted = st.form_submit_button("שמור 💾", type="primary", use_container_width=True)

        if submitted:
            ver_save = (row.get("Version") or "").strip() or "V1"
            base = get_quote_from_csv(client, project, ver_save)
            if base is None:
                base = {c: (row.get(c, "") or "") for c in QUOTES_CSV_COLUMNS}
            else:
                base = {c: (base.get(c, "") or "") for c in QUOTES_CSV_COLUMNS}
            base["Client Email"] = (f_client_email or "").strip()
            base["Client Phone"] = (f_client_phone or "").strip()
            base["Architect Contact"] = (f_architect or "").strip()
            base["Project Special Notes"] = (f_notes or "").strip()
            if update_quote_in_csv(client, project, ver_save, base, skip_rerun=True):
                st.success("הפרטים נשמרו בהצלחה.")
                time.sleep(0.5)
                st.rerun()
            else:
                st.error("לא נמצאה שורת הצעה לעדכון (Client / Project / Version).")

    st.subheader("קישורי דרופבוקס")
    proj_row = _find_project_row_for_hub(client, project)
    if proj_row is None:
        st.info(
            "לא נמצאה רשומת פרויקט תואמת בגיליון projects (אותו לקוח ושם פרויקט). "
            "לאחר יצירת הפרויקט והגדרת קישורים — הם יופיעו כאן."
        )
    else:
        dm = (proj_row.get("Dropbox_Main") or "").strip()
        du = (proj_row.get("Dropbox_Upload") or "").strip()
        dd = (proj_row.get("Dropbox_Deliverables") or "").strip()
        if not dm and not du and not dd:
            st.caption("אין קישורי דרופבוקס שמורים לפרויקט זה במסד.")
        else:
            c1, c2, c3 = st.columns(3)
            if dm:
                c1.link_button("דרופבוקס — ראשי", dm, use_container_width=True)
            if du:
                c2.link_button("דרופבוקס — בקשת קבצים / העלאה", du, use_container_width=True)
            if dd:
                c3.link_button("דרופבוקס — תוצרים", dd, use_container_width=True)


def show_monitor_3d_page() -> None:
    """מסך מוניטור צוות / משימות: פרויקטים פעילים עם עריכת סטטוס, גאנט ומשימות."""
    st.title("מוניטור צוות / משימות 🖥️")

    projects_rows = read_projects()
    if not projects_rows:
        st.info("אין פרויקטים. נתוני הפרויקטים נמשכים מגוגל שיטס.")
        return

    # סינון: רק פרויקטים פעילים (התעלם מ'הסתיים' ו'בוטל')
    EXCLUDED_STATUSES = ("הסתיים", "בוטל")
    excluded_lower = [s.strip().lower() for s in EXCLUDED_STATUSES]
    active_rows = [
        r for r in projects_rows
        if (r.get("Status") or "").strip().lower() not in excluded_lower
    ]

    if not active_rows:
        st.info("לא נמצאו פרויקטים פעילים (כל הפרויקטים בסטטוס 'הסתיים' או 'בוטל').")
        return

    # עמודות להצגה: שם פרויקט, לקוח, תאריך יעד, סטטוס
    df_full = pd.DataFrame(active_rows, columns=PROJECTS_DB_COLUMNS).fillna("")
    df_display = df_full[["Project Name", "Client", "Start Date", "Status"]].copy()
    df_display = df_display.rename(columns={"Start Date": "תאריך יעד (דד-ליין)", "Project Name": "שם פרויקט", "Client": "לקוח", "Status": "סטטוס"})
    df_display = df_display[["שם פרויקט", "לקוח", "תאריך יעד (דד-ליין)", "סטטוס"]]  # סדר לפי הדרישה

    # אפשרויות סטטוס למוניטור תלת-מימד + סטטוסים קיימים בנתונים
    status_options = list(dict.fromkeys(MONITOR_3D_STATUS_OPTIONS + [s for s in df_display["סטטוס"].unique() if s and str(s).strip()]))

    edited_df = st.data_editor(
        df_display,
        hide_index=True,
        use_container_width=True,
        disabled=["שם פרויקט", "לקוח", "תאריך יעד (דד-ליין)"],
        column_config={
            "סטטוס": st.column_config.SelectboxColumn(
                "סטטוס",
                options=status_options,
                required=True,
            ),
        },
        key="monitor_3d_editor",
    )

    if st.button("שמור עדכוני סטטוס 💾", type="primary", key="save_monitor_3d_btn", use_container_width=True):
        # מיזוג הסטטוס המעודכן חזרה לרשימה המלאה
        edited_by_key = {(str(r.get("לקוח", "")), str(r.get("שם פרויקט", ""))): r.get("סטטוס", "") for _, r in edited_df.iterrows()}
        full_rows = read_projects()
        changed = False
        for i, row in enumerate(full_rows):
            key = (str(row.get("Client", "")), str(row.get("Project Name", "")))
            if key in edited_by_key:
                new_status = (edited_by_key[key] or "").strip()
                if new_status and (row.get("Status") or "").strip() != new_status:
                    full_rows[i] = {**row, "Status": new_status}
                    changed = True

        if changed:
            write_projects(full_rows, skip_rerun=True)
            time.sleep(1.5)
            st.success("העדכונים נשמרו בהצלחה! ✅")
            st.rerun()
        else:
            st.info("לא בוצעו שינויים בשמירה.")

    # --- תרשים גאנט וטבלת משימות ---
    st.divider()
    st.subheader("תרשים גאנט וטבלת משימות")
    is_management = st.session_state.get("is_management", False)
    current_user = (st.session_state.get("current_user") or "").strip()
    view_filter = (st.session_state.get("view_filter") or "").strip()

    if not is_management:
        st.caption(f"מוצגות רק משימות שבהן **איש צוות / אחראי** הוא **{current_user}**.")
    elif view_filter and view_filter != "הצג הכל":
        st.caption(f"תצוגת מנהל: מוצגות משימות של **{view_filter}** (סינון תצוגה בלבד).")

    tasks_rows_raw = read_tasks()
    _cancel_lower = {s.strip().lower() for s in ("בוטל",)}
    gantt_tasks = [
        t for t in tasks_rows_raw
        if (t.get("סטטוס") or "").strip().lower() not in _cancel_lower
    ]

    if view_filter and view_filter != "הצג הכל":
        gantt_tasks = [
            t for t in gantt_tasks
            if _task_row_matches_view_filter(t, view_filter)
        ]
    elif not is_management and not current_user:
        gantt_tasks = []

    active_tasks = [
        t for t in gantt_tasks
        if (t.get("סטטוס") or "").strip().lower() != "הסתיים"
    ]

    if not gantt_tasks:
        st.info("אין משימות להצגה (לפי הסינון).")
        return

    if not active_tasks:
        st.success("אין משימות פתוחות כרגע! 🎉")
        st.caption("להלן גאנט כולל משימות שהושלמו (מעומעמות).")

    # תרשים גאנט (מעל טבלת המשימות)
    try:
        gantt_df, y_col, _err = _build_gantt_dataframe_for_timeline(gantt_tasks, "task_detail")
        if gantt_df is not None and y_col:
            _gantt_color_map = {**TEAM_GANTT_COLOR_HEX}
            for _k in gantt_df["_task_color_key"].unique():
                if _k not in _gantt_color_map:
                    _gantt_color_map[_k] = "#888888"
            fig = px.timeline(
                gantt_df,
                x_start="_start_dt",
                x_end="_end_dt",
                y=y_col,
                color="_task_color_key",
                color_discrete_map=_gantt_color_map,
                text="_gantt_bar_text",
                title="תרשים גאנט - משימות (פעילות והסתיים בבהירות נמוכה)",
            )
            fig.update_yaxes(autorange="reversed")  # משימות חדשות בראש
            fig.update_layout(xaxis_title="תאריך", yaxis_title="משימה")
            _style_gantt_timeline_figure(fig, gantt_df, y_col)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("אין משימות פעילות להצגת תרשים גאנט.")
    except Exception:
        st.warning("אין מספיק נתוני תאריכים להצגת תרשים גאנט")

    # טבלת משימות — רק משימות פתוחות
    if not active_tasks:
        return

    df_tasks = pd.DataFrame(active_tasks, columns=TASKS_LOG_COLUMNS).reindex(columns=TASKS_LOG_COLUMNS, fill_value="").fillna("")
    if hasattr(df_tasks.columns, 'str'):
        df_tasks.columns = df_tasks.columns.str.strip()

    task_status_options = ['ממתין', 'בעבודה', 'הסתיים']
    existing_statuses = [s for s in df_tasks["סטטוס"].unique() if s and str(s).strip()]
    status_options = list(dict.fromkeys(task_status_options + existing_statuses))
    disabled_cols = [c for c in TASKS_LOG_COLUMNS if c != "סטטוס"]

    def _task_key(r):
        return (str(r.get("פרויקט", "") or ""), str(r.get("שם משימה", "") or ""), str(r.get("הוקצה ל", "") or ""),
                str(r.get("תאריך התחלה", "") or ""), str(r.get("תאריך יעד", "") or ""))

    st.data_editor(
        df_tasks,
        hide_index=True,
        use_container_width=True,
        key="task_editor",
        disabled=disabled_cols,
        column_config={
            "הוקצה ל": st.column_config.TextColumn("איש צוות / אחראי"),
            "סטטוס": st.column_config.SelectboxColumn(
                "סטטוס",
                options=status_options,
                required=True,
            ),
        },
    )

    if st.button("שמור עדכוני סטטוס 💾", type="primary", key="save_task_status_btn", use_container_width=True):
        edited = st.session_state.get("task_editor", {}).get("edited_rows", {})
        if not edited:
            st.info("לא בוצעו שינויים בשמירה.")
        elif spreadsheet is None:
            st.error("אין חיבור לגוגל שיטס. לא ניתן לשמור.")
        else:
            # מיזוג עדכוני סטטוס לפי מפתח שורה (פרויקט, שם משימה, הוקצה ל, תאריכים)
            edits_by_key = {}
            for row_idx, changes in edited.items():
                new_status = (changes.get("סטטוס") or "").strip()
                if not new_status:
                    continue
                try:
                    display_row = df_tasks.iloc[int(row_idx)]
                    edits_by_key[_task_key(display_row)] = new_status
                except Exception:
                    pass
            if edits_by_key:
                any_completed = "הסתיים" in edits_by_key.values()
                updated_rows = []
                for r in tasks_rows_raw:
                    k = _task_key(r)
                    if k in edits_by_key:
                        r = dict(r)
                        r["סטטוס"] = edits_by_key[k]
                    updated_rows.append(r)
                write_tasks(updated_rows, skip_rerun=True)
                st.success("הסטטוס עודכן בהצלחה!")
                if any_completed:
                    st.balloons()
            if "task_editor" in st.session_state:
                del st.session_state["task_editor"]
            time.sleep(1)
            st.rerun()


def _render_edit_existing_task_block() -> None:
    """אזור עריכת משימה קיימת — משימות פעילות, שמירה לגיליון tasks בגוגל שיטס."""
    st.subheader("✏️ עריכת משימה קיימת")
    tasks_rows = read_tasks()
    if not tasks_rows:
        st.info("אין משימות במערכת.")
        return

    tasks_df = pd.DataFrame(tasks_rows, columns=TASKS_LOG_COLUMNS).reindex(
        columns=TASKS_LOG_COLUMNS, fill_value=""
    ).fillna("")
    if hasattr(tasks_df.columns, "str"):
        tasks_df.columns = tasks_df.columns.str.strip()

    excluded_status = ("הסתיים", "בוטל")
    ex_lower = {s.strip().lower() for s in excluded_status}

    def _is_active_status(s: str) -> bool:
        return (s or "").strip().lower() not in ex_lower

    mask = tasks_df["סטטוס"].astype(str).apply(_is_active_status)
    # כל המשימות הפעילות מהגיליון — כולל «כללי / סטודיו», «סטודיו» (חופשה) וסוג חופשה/היעדרות
    active_records = tasks_df.loc[mask].to_dict(orient="records")

    is_management_edit = st.session_state.get("is_management", False)
    current_user_edit = (st.session_state.get("current_user") or "").strip()
    view_filter_edit = (st.session_state.get("view_filter") or "").strip()
    if not is_management_edit and current_user_edit:
        active_records = [
            r
            for r in active_records
            if _assignee_cell_matches_login(r.get("הוקצה ל"), current_user_edit)
        ]
        st.caption(f"מוצגות רק משימות שבהן **אחראי** הוא **{current_user_edit}**.")
    elif not is_management_edit and not current_user_edit:
        active_records = []
    elif is_management_edit and view_filter_edit and view_filter_edit != "הצג הכל":
        active_records = [
            r for r in active_records if _task_row_matches_view_filter(r, view_filter_edit)
        ]
        st.caption(f"תצוגת מנהל: משימות של **{view_filter_edit}** (סינון תצוגה בלבד).")

    if not active_records:
        st.info("אין משימות פעילות לעריכה.")
        return
    if spreadsheet is None:
        st.warning("אין חיבור לגוגל שיטס — לא ניתן לערוך משימות עד שהחיבור יוחזר.")
        return

    labels = [_format_edit_task_select_label(r) for r in active_records]
    pick = st.selectbox(
        "בחר משימה פעילה",
        options=list(range(len(active_records))),
        format_func=lambda i: labels[i],
        key="edit_existing_task_pick_main",
    )
    sel = active_records[pick]
    tid = (sel.get("מזהה משימה") or "").strip() or f"idx_{pick}"

    team_keys = list(TEAM_EMAILS.keys())

    def _assignee_key_for_row(t: dict) -> str:
        cell = str(t.get("הוקצה ל") or "")
        for k in TEAM_EMAILS:
            if _assignee_matches_team_key(cell, k):
                return k
        return team_keys[0]

    default_team = _assignee_key_for_row(sel)
    start_parsed = _parse_task_date(str(sel.get("תאריך התחלה") or ""))
    start_val = start_parsed.date() if start_parsed else date.today()
    due_parsed = _parse_task_date(str(sel.get("תאריך יעד") or ""))
    due_val = due_parsed.date() if due_parsed else date.today()

    cur_status = (sel.get("סטטוס") or "").strip()
    existing_statuses = sorted(
        {
            str(r.get("סטטוס") or "").strip()
            for r in tasks_rows
            if str(r.get("סטטוס") or "").strip()
        }
    )
    status_opts = list(dict.fromkeys(list(TASK_EDIT_STATUS_OPTIONS) + existing_statuses))

    project_opts = _task_project_select_options()
    cur_proj = (sel.get("פרויקט") or "").strip()
    display_project = _task_project_stored_to_label(cur_proj)
    if display_project not in project_opts:
        project_opts = list(dict.fromkeys(project_opts + [display_project]))

    cur_task_type = (sel.get("שם משימה") or "").strip()
    type_opts = list(TASK_TYPE_OPTIONS)
    if cur_task_type and cur_task_type not in type_opts:
        type_opts = list(dict.fromkeys(type_opts + [cur_task_type]))
    type_index = type_opts.index(cur_task_type) if cur_task_type in type_opts else 0

    with st.form("edit_task_google_sheet_form"):
        st.text_input(
            "תיאור משימה",
            value=(sel.get("תיאור המשימה") or ""),
            key=f"edit_gs_desc_{tid}",
        )
        st.selectbox(
            "סוג משימה",
            options=type_opts,
            index=type_index,
            key=f"edit_gs_type_{tid}",
        )
        if st.session_state.get(f"edit_gs_type_{tid}") == TASK_TYPE_OOO:
            st.caption("חופשה/היעדרות נשמרת תחת פרויקט «סטודיו» (כמו בהקצאת משימה חדשה).")
        proj_pick_index = (
            project_opts.index(display_project)
            if display_project in project_opts
            else 0
        )
        st.selectbox(
            "פרויקט",
            options=project_opts,
            index=proj_pick_index,
            key=f"edit_gs_project_{tid}",
        )
        assignee_index = team_keys.index(default_team) if default_team in team_keys else 0
        st.selectbox(
            "איש צוות",
            options=team_keys,
            index=assignee_index,
            key=f"edit_gs_assignee_{tid}",
        )
        st.date_input(
            "תאריך התחלה",
            value=start_val,
            key=f"edit_gs_start_{tid}",
        )
        st.date_input(
            "תאריך יעד",
            value=due_val,
            key=f"edit_gs_due_{tid}",
        )
        status_index = status_opts.index(cur_status) if cur_status in status_opts else 0
        st.selectbox(
            "סטטוס",
            options=status_opts,
            index=status_index,
            key=f"edit_gs_status_{tid}",
        )
        submitted = st.form_submit_button("עדכן משימה בגוגל שיטס")

    if st.button("🗑️ מחק משימה זו", type="primary", key=f"delete_task_gs_{tid}"):
        full_rows_del = read_tasks()
        if not full_rows_del:
            st.error("אין משימות במערכת.")
        else:
            tasks_df_del = pd.DataFrame(full_rows_del, columns=TASKS_LOG_COLUMNS).reindex(
                columns=TASKS_LOG_COLUMNS, fill_value=""
            ).fillna("")
            if hasattr(tasks_df_del.columns, "str"):
                tasks_df_del.columns = tasks_df_del.columns.str.strip()
            row_idx_del = _find_task_row_index_in_full_list(full_rows_del, sel)
            if row_idx_del is None:
                st.error("לא נמצאה המשימה בגיליון (אולי נמחקה או עודכנה). רענן ונסה שוב.")
            else:
                updated_del = tasks_df_del.drop(index=row_idx_del).reset_index(drop=True)
                write_tasks(updated_del.to_dict(orient="records"), skip_rerun=True)
                st.success("המשימה נמחקה מהמערכת.")
                time.sleep(1)
                st.rerun()

    if submitted:
        new_desc = (st.session_state.get(f"edit_gs_desc_{tid}") or "").strip()
        new_assignee = st.session_state.get(f"edit_gs_assignee_{tid}")
        new_start = st.session_state.get(f"edit_gs_start_{tid}")
        new_due = st.session_state.get(f"edit_gs_due_{tid}")
        new_status = st.session_state.get(f"edit_gs_status_{tid}")
        new_project_raw = st.session_state.get(f"edit_gs_project_{tid}")
        new_task_type = (st.session_state.get(f"edit_gs_type_{tid}") or "").strip()
        new_project = _task_project_label_to_stored(str(new_project_raw or ""))
        if new_task_type == TASK_TYPE_OOO:
            new_project = TASKS_PROJECT_OOO_DEFAULT
        if new_assignee is None or new_status is None or new_due is None or new_start is None:
            st.error("חסרים ערכים בטופס — נסה שוב.")
        elif not new_task_type:
            st.error("חסר סוג משימה — נסה שוב.")
        else:
            full_rows = read_tasks()
            row_idx = _find_task_row_index_in_full_list(full_rows, sel)
            if row_idx is None:
                st.error("לא נמצאה המשימה בגיליון (אולי נמחקה או עודכנה). רענן ונסה שוב.")
            else:
                updated = dict(full_rows[row_idx])
                updated["תיאור המשימה"] = new_desc
                updated["שם משימה"] = new_task_type
                updated["פרויקט"] = new_project
                updated["הוקצה ל"] = new_assignee
                updated["תאריך התחלה"] = new_start.strftime("%Y-%m-%d")
                updated["תאריך יעד"] = new_due.strftime("%Y-%m-%d")
                updated["סטטוס"] = new_status
                full_rows[row_idx] = updated
                write_tasks(full_rows, skip_rerun=True)
                st.success("המשימה עודכנה בגוגל שיטס!")
                time.sleep(1)
                st.rerun()


def show_tasks_page() -> None:
    st.title("ניהול פרויקטים ומשימות")
    sub_nav = st.radio(
        "בחר תצוגה:",
        ["הקצאת משימה חדשה 🎯", "מוניטור צוות (עדכון סטטוסים) 📋", "לוח עומסים (גאנט ויומן) 📊"],
        horizontal=True,
    )

    if sub_nav == "הקצאת משימה חדשה 🎯":
        st.subheader("הקצאת משימה חדשה לצוות")
        if not _get_active_projects_options():
            st.info(
                "אין פרויקטי לקוח רשומים בגיליון projects — ניתן עדיין להקצות משימות **כלליות / סטודיו** מהרשימה."
            )
        projects_options = _task_project_select_options()
        with st.form("assign_task_form", clear_on_submit=True):
            selected_project = st.selectbox("פרויקט", options=projects_options, key="assign_task_project")
            task_type = st.selectbox("סוג משימה", options=TASK_TYPE_OPTIONS, key="assign_task_type")
            if st.session_state.get("assign_task_type") == TASK_TYPE_OOO:
                st.caption("חופשה/היעדרות נשמרת תחת פרויקט «סטודיו» (ללא צורך בפרויקט לקוח).")
            st.text_area(
                "פירוט המשימה / הערות נוספות (אופציונלי)",
                key="assign_task_notes",
                placeholder="ניתן להרחיב — הטקסט יישמר בעמודת תיאור המשימה ויישלח במייל לצוות",
                height=120,
            )
            assignee = st.selectbox("הוקצה ל:", options=TASK_TEAM, key="assign_task_assignee")
            start_date = st.date_input("תאריך התחלה", value=date.today(), key="assign_task_start")
            due_date = st.date_input("תאריך יעד למשימה", value=date.today(), key="assign_task_due")
            submitted = st.form_submit_button("הקצה משימה 🚀")

        if submitted:
            notes = (st.session_state.get("assign_task_notes") or "").strip()
            task_name = task_type
            project_for_row = (
                TASKS_PROJECT_OOO_DEFAULT
                if task_type == TASK_TYPE_OOO
                else _task_project_label_to_stored(selected_project)
            )
            # גיליון: סוג ב«שם משימה»; ב«תיאור» — שילוב סוג + פירוט כשיש טקסט חופשי
            task_desc = f"{task_type}\n\n{notes}" if notes else ""

            existing = read_tasks()
            row = {
                "מזהה משימה": str(uuid.uuid4()),
                "פרויקט": project_for_row,
                "שם משימה": task_name,
                "תיאור המשימה": task_desc,
                "הוקצה ל": assignee,
                "תאריך התחלה": start_date.strftime("%Y-%m-%d"),
                "תאריך יעד": due_date.strftime("%Y-%m-%d"),
                "סטטוס": "ממתין",
            }
            existing.append(row)
            write_tasks(existing, skip_rerun=True)
            assignee_email = _team_email_for_task_assignee(assignee)
            if assignee_email:
                cc_list: list[str] = []
                for k in ("ערן", "טל"):
                    em = TEAM_EMAILS.get(k)
                    if em and str(em).strip() and str(em).strip() != assignee_email:
                        cc_list.append(str(em).strip())
                send_task_assignment_email_tali(
                    assignee_email,
                    project_for_row,
                    task_name,
                    task_desc,
                    due_date.strftime("%d/%m/%Y"),
                    assignee,
                    cc_emails=cc_list,
                )
            else:
                st.warning(
                    "המשימה נשמרה בגיליון, אך לא נמצאה כתובת מייל לנמען שנבחר — לא נשלח מייל לצוות."
                )
            st.success("המשימה הוקצתה בהצלחה! ✅")
            time.sleep(1)
            st.rerun()

    _render_edit_existing_task_block()
    st.divider()

    if sub_nav == "מוניטור צוות (עדכון סטטוסים) 📋":
        # --- מוניטור סטודיו - תמונת מצב צוותית ---
        st.subheader("🎯 מוניטור סטודיו - תמונת מצב צוותית")

        # נתונים מ-read_tasks() (גיליון tasks בגוגל שיטס) — לא רשימה זמנית אחרת
        tasks_rows_monitor = read_tasks()
        # משימות פתוחות: סטטוס לא הושלם; כולל תאריך יעד עתידי (בניגוד לטבלה למטה שמסננת לפי תאריך)
        open_tasks = _filter_tasks_open_not_done(tasks_rows_monitor)

        today_dt = date.today()

        def _get_task_due_date(task: dict) -> date | None:
            due_str = task.get("תאריך יעד") or ""
            if not due_str or not str(due_str).strip():
                return None
            dt = pd.to_datetime(due_str, errors="coerce")
            if pd.isna(dt):
                return None
            return dt.date() if hasattr(dt, "date") else dt

        def _mark_task_done_monitor(task: dict) -> None:
            idx_done = _find_task_row_index_in_full_list(tasks_rows_monitor, task)
            if idx_done is not None:
                tasks_rows_monitor[idx_done] = {**tasks_rows_monitor[idx_done], "סטטוס": "הסתיים"}
            write_tasks(tasks_rows_monitor, skip_rerun=True)
            st.rerun()

        vf_mon = (st.session_state.get("view_filter") or "").strip()
        is_mgmt_mon = st.session_state.get("is_management", False)
        if is_mgmt_mon and vf_mon == "הצג הכל":
            assignees_iter = list(TEAM_EMAILS.keys())
        elif vf_mon and vf_mon != "הצג הכל":
            assignees_iter = [vf_mon] if vf_mon in TEAM_EMAILS else []
        else:
            assignees_iter = (
                [k for k in TEAM_EMAILS if _assignee_cell_matches_login(k, vf_mon)]
                if vf_mon
                else []
            )

        for assignee in assignees_iter:
            user_tasks = [
                t
                for t in open_tasks
                if _assignee_matches_team_key(t.get("הוקצה ל") or "", assignee)
            ]
            try:
                with st.expander(
                    f"{assignee} | משימות פתוחות: {len(user_tasks)}",
                    expanded=False,
                ):
                    if not user_tasks:
                        st.info("אין משימות פתוחות")
                    else:
                        st.caption("סיום מהיר — לחיצה אחת מעדכנת את הסטטוס בגיליון.")
                        for idx, r in enumerate(user_tasks):
                            task_name = (r.get("שם משימה") or "").strip()
                            project = (r.get("פרויקט") or "").strip()
                            due_str = (r.get("תאריך יעד") or "").strip()
                            due_parsed = _get_task_due_date(r)
                            tid = (r.get("מזהה משימה") or "").strip() or f"fallback_{assignee}_{idx}"
                            col_msg, col_btn = st.columns([4, 1])
                            with col_msg:
                                if due_parsed is None or due_parsed < today_dt:
                                    st.error(f"🔴 **{task_name}** — {project} — {due_str}")
                                elif due_parsed == today_dt:
                                    st.info(f"🔵 **{task_name}** — {project} — {due_str}")
                                else:
                                    st.success(f"🟢 **{task_name}** — {project} — {due_str}")
                            with col_btn:
                                if st.button("✔️ סיימתי", key=f"quick_done_{tid}"):
                                    _mark_task_done_monitor(r)
            except Exception as e:
                st.error(f"שגיאה בהצגת נתונים: {e}")

        st.divider()
        with st.expander("🗑️ ניקוי ומחיקת פרויקטים מהמערכת"):
            delete_options = _get_active_projects_options()
            if not delete_options:
                st.info("אין פרויקטים פעילים למחיקה.")
            else:
                project_to_delete = st.selectbox(
                    "בחר פרויקט למחיקה",
                    options=delete_options,
                    key="delete_project_select",
                )
                if st.button("מחק פרויקט לצמיתות", type="primary", key="delete_project_btn"):
                    try:
                        client, project_name = project_to_delete.split(" | ", 1)
                        client = client.strip()
                        project_name = project_name.strip()
                        if not client or not project_name:
                            st.error("לא ניתן לחלץ לקוח/פרויקט מהבחירה.")
                        else:
                            project_key_pipe = f"{client} | {project_name}"
                            project_key_dash = f"{client} - {project_name}"

                            # 1. projects (גוגל שיטס)
                            db_rows = read_projects()
                            db_updated = [
                                r for r in db_rows
                                if not (
                                    (r.get("Client") or "").strip() == client
                                    and (r.get("Project Name") or "").strip() == project_name
                                )
                            ]
                            if len(db_updated) < len(db_rows):
                                write_projects(db_updated)

                            # 2. projects.csv
                            csv_rows = read_projects_csv()
                            csv_updated = [
                                r for r in csv_rows
                                if not (
                                    (r.get("Client") or "").strip() == client
                                    and (r.get("Project") or "").strip() == project_name
                                )
                            ]
                            if len(csv_updated) < len(csv_rows):
                                write_projects_csv(csv_updated)

                            # 3. quotes (גוגל שיטס)
                            quotes_rows = read_quotes_log()
                            quotes_updated = [
                                r for r in quotes_rows
                                if not (
                                    (r.get("Client") or "").strip() == client
                                    and (r.get("Project") or "").strip() == project_name
                                )
                            ]
                            if len(quotes_updated) < len(quotes_rows):
                                write_quotes_log(quotes_updated)

                            # 4. tasks (גוגל שיטס)
                            tasks_rows = read_tasks()
                            tasks_updated = [
                                t for t in tasks_rows
                                if (t.get("פרויקט") or "").strip() not in (project_key_pipe, project_key_dash)
                            ]
                            if len(tasks_updated) < len(tasks_rows):
                                write_tasks(tasks_updated)

                            # 5. project_contacts (גוגל שיטס)
                            contacts_rows = read_project_contacts()
                            contacts_updated = [
                                c for c in contacts_rows
                                if (c.get("Project") or "").strip() not in (project_key_pipe, project_key_dash)
                            ]
                            if len(contacts_updated) < len(contacts_rows):
                                write_project_contacts(contacts_updated)

                            st.success("הפרויקט נמחק בהצלחה מכל קבצי המערכת!")
                            st.rerun()
                    except ValueError:
                        st.error("לא ניתן לפרש את הבחירה. הפורמט הצפוי: 'לקוח | שם פרויקט'.")

        st.divider()
        st.subheader("טבלת משימות")
        tasks_rows = read_tasks()
        if not tasks_rows:
            st.info("אין משימות. השתמש בטאב 'הקצאת משימה חדשה' להוספת משימות.")
        else:
            # סינון חכם: סטטוס לא בוצע, תאריך <= היום. מיון: באיחור ראשון, אחר כך היום.
            filtered_rows = _filter_tasks_by_status_and_date(
                tasks_rows, date_col="תאריך יעד", status_col="סטטוס"
            )
            df = pd.DataFrame(filtered_rows, columns=TASKS_LOG_COLUMNS)
            df = df.reindex(columns=TASKS_LOG_COLUMNS, fill_value='')
            df = df.fillna('')
            if hasattr(df.columns, 'str'):
                df.columns = df.columns.str.strip()
            # הוספת עמודת אינדיקציה לאיחור (🔴 למשימות שנגררו)
            today_dt = date.today()
            def _overdue_indicator(row):
                d = _parse_date_safe(row.get("תאריך יעד"), "תאריך יעד")
                return "🔴" if d is None or d < today_dt else ""
            df.insert(0, "איחור", df.apply(_overdue_indicator, axis=1))
            # עמודת סימון למחיקה מרובה
            DELETE_COL = "סמן למחיקה 🗑️"
            df.insert(0, DELETE_COL, False)
            vf_tbl = (st.session_state.get("view_filter") or "").strip()
            if vf_tbl and vf_tbl != "הצג הכל":
                df = df[
                    df["הוקצה ל"].apply(
                        lambda cell: _task_row_matches_view_filter({"הוקצה ל": cell}, vf_tbl)
                    )
                ]
            if df.empty:
                st.info("אין משימות להצגה.")
            else:
                existing_s = [s for s in df["סטטוס"].unique() if s and str(s).strip()]
                status_options_table = list(dict.fromkeys(list(TASK_EDIT_STATUS_OPTIONS) + existing_s))
                try:
                    editable_cols = ["סטטוס"]
                    # איחור וסמן למחיקה - עריכה מותרת רק בסמן למחיקה (תיבת סימון)
                    disabled_cols = ["איחור"] + [c for c in TASKS_LOG_COLUMNS if c not in editable_cols]
                    edited_df = st.data_editor(
                        df,
                        hide_index=True,
                        use_container_width=True,
                        disabled=disabled_cols,
                        column_config={
                            DELETE_COL: st.column_config.CheckboxColumn(
                                DELETE_COL,
                                help="סמן משימות למחיקה",
                                default=False,
                            ),
                            "סטטוס": st.column_config.SelectboxColumn(
                                "סטטוס",
                                options=status_options_table,
                                required=True,
                            ),
                        },
                        key="tasks_editor",
                    )
                except Exception as e:
                    st.error(f"שגיאה בהצגת נתונים: {e}")
                    edited_df = df
                if st.button("שמור שינויים", type="primary", key="save_tasks_btn"):
                    # הסרת עמודות עזר לפני שמירה (איחור, סמן למחיקה)
                    save_df = edited_df.drop(columns=["איחור", DELETE_COL], errors="ignore")
                    def _row_key(r):
                        return (str(r.get("פרויקט", "") or ""), str(r.get("שם משימה", "") or ""),
                               str(r.get("הוקצה ל", "") or ""), str(r.get("תאריך התחלה", "") or ""),
                               str(r.get("תאריך יעד", "") or ""))
                    if is_admin():
                        updated = save_df.to_dict(orient="records")
                    else:
                        # משתמש לא-מנהל: מיזוג השינויים חזרה לרשימת המשימות המלאה
                        full_rows = read_tasks()
                        edited_records = save_df.to_dict(orient="records")
                        edited_by_key = {_row_key(r): r for r in edited_records}
                        for i, row in enumerate(full_rows):
                            k = _row_key(row)
                            if k in edited_by_key:
                                full_rows[i] = edited_by_key[k]
                        updated = full_rows
                    write_tasks(updated, skip_rerun=True)
                    st.success("השינויים נשמרו בהצלחה!")
                    time.sleep(1)
                    st.rerun()

                # כפתור מחיקה מרובה
                if st.button("מחק משימות שסומנו 🚨", type="primary", key="delete_tasks_btn"):
                    to_delete = edited_df[edited_df[DELETE_COL] == True]
                    if to_delete.empty:
                        st.warning("לא סומנו משימות למחיקה. סמן תיבות בעמודה 'סמן למחיקה 🗑️' ולחץ שוב.")
                    else:
                        def _del_key(r):
                            return (str(r.get("פרויקט", "") or ""), str(r.get("שם משימה", "") or ""),
                                    str(r.get("הוקצה ל", "") or ""), str(r.get("תאריך התחלה", "") or ""),
                                    str(r.get("תאריך יעד", "") or ""))
                        to_remove_keys = {_del_key(r) for _, r in to_delete.iterrows()}
                        updated_rows = [r for r in tasks_rows if _del_key(r) not in to_remove_keys]
                        write_tasks(updated_rows, skip_rerun=True)
                        st.success("המשימות נמחקו בהצלחה!")
                        time.sleep(1)
                        st.rerun()

    elif sub_nav == "לוח עומסים (גאנט ויומן) 📊":
        st.subheader("לוח עומסים צוותי - לוח שנה אינטראקטיבי")

        is_management_wl = st.session_state.get("is_management", False)
        current_user_wl = (st.session_state.get("current_user") or "").strip()
        vf_wl = (st.session_state.get("view_filter") or "").strip()
        if is_management_wl and vf_wl == "הצג הכל":
            st.caption("תרשים הגאנט מציג את כל משימות הצוות (מנהלים: טל וערן).")
        elif is_management_wl and vf_wl and vf_wl != "הצג הכל":
            st.caption(f"תצוגת מנהל: הגאנט מציג משימות של **{vf_wl}** (סינון תצוגה בלבד).")
        elif current_user_wl:
            st.caption(f"תרשים הגאנט מציג רק את **המשימות שלך** (אחראי: {current_user_wl}).")
        else:
            st.caption("לא זוהה משתמש — הגאנט עשוי להיות ריק.")

        st.markdown("#### תרשים גאנט אינטראקטיבי")
        wl_axis_mode = st.radio(
            "ציר אנכי (בחירת תצוגה)",
            ["task_detail", "assignee", "project"],
            format_func=lambda x: {
                "task_detail": "משימה (פרויקט | שם)",
                "assignee": "עובד (אחראי)",
                "project": "פרויקט",
            }[x],
            horizontal=True,
            key="workload_gantt_y_mode",
        )
        tasks_wl_all = read_tasks()
        tasks_for_gantt_wl = [
            t for t in tasks_wl_all
            if (t.get("סטטוס") or "").strip().lower() != "בוטל"
        ]
        if vf_wl and vf_wl != "הצג הכל":
            tasks_for_gantt_wl = [t for t in tasks_for_gantt_wl if _task_row_matches_view_filter(t, vf_wl)]
        elif not is_management_wl and not current_user_wl:
            tasks_for_gantt_wl = []

        try:
            gantt_wl, y_wl, err_wl = _build_gantt_dataframe_for_timeline(tasks_for_gantt_wl, wl_axis_mode)
            if gantt_wl is not None and y_wl:
                _cm_wl = {**TEAM_GANTT_COLOR_HEX}
                for _k in gantt_wl["_task_color_key"].unique():
                    if _k not in _cm_wl:
                        _cm_wl[_k] = "#888888"
                _y_title = {"task_detail": "משימה", "assignee": "עובד", "project": "פרויקט"}.get(
                    wl_axis_mode, "משימה"
                )
                fig_wl = px.timeline(
                    gantt_wl,
                    x_start="_start_dt",
                    x_end="_end_dt",
                    y=y_wl,
                    color="_task_color_key",
                    color_discrete_map=_cm_wl,
                    text="_gantt_bar_text",
                    title="תרשים גאנט — עומסי עבודה (תאריך התחלה עד תאריך יעד)",
                )
                fig_wl.update_yaxes(autorange="reversed")
                fig_wl.update_layout(
                    xaxis_title="זמן",
                    yaxis_title=_y_title,
                )
                _style_gantt_timeline_figure(fig_wl, gantt_wl, y_wl)
                st.plotly_chart(fig_wl, use_container_width=True)
            else:
                st.info(err_wl or "אין משימות פעילות להצגת גאנט.")
        except Exception as e:
            st.warning(f"לא ניתן להציג גאנט: {e}")

        st.divider()
        st.subheader("לוח שנה — חגים ומשימות")
        calendar_events = []

        # טעינת חגים וחופשות בטוחה
        try:
            current_year = datetime.today().year
            # 1. טעינת חגי ישראל מהספרייה
            il_holidays = holidays.IL(years=[current_year, current_year+1])
            for hol_date, hol_name in il_holidays.items():
                calendar_events.append({
                    "title": f"חג: {hol_name}",
                    "start": hol_date.strftime("%Y-%m-%d"),
                    "color": "#ffb6c1",  # ורוד עדין
                    "allDay": True,
                    "display": "background"
                })

            # 2. הוספת החופש הגדול (בתי ספר - 1 ביולי עד 31 באוגוסט)
            for y in [current_year, current_year+1]:
                calendar_events.append({
                    "title": "החופש הגדול (בתי ספר)",
                    "start": f"{y}-07-01",
                    "end": f"{y}-09-01",  # יום אחרי הסיום כדי שיצבע את כל אוגוסט
                    "color": "#fffacd",  # צהוב בהיר ועדין שלא מסתיר טקסט
                    "allDay": True,
                    "display": "background"
                })
        except Exception as e:
            st.warning(f"שגיאה בטעינת חגים: {e}")

        # טעינת משימות מוגנת - זמנית: הצגת כל המשימות (ללא סינון Status)
        try:
            tasks_rows = read_tasks()
            df_tasks = pd.DataFrame(tasks_rows) if tasks_rows else pd.DataFrame()
            if not df_tasks.empty and hasattr(df_tasks.columns, 'str'):
                df_tasks.columns = df_tasks.columns.str.strip()
            if df_tasks.empty:
                df_open = pd.DataFrame()
            else:
                # זמנית: הצגת כל המשימות (ביטול סינון Status != Done)
                df_open = df_tasks.copy()
            for index, row in df_open.iterrows():
                try:
                    start_val = row.get("תאריך התחלה", row.get("תאריך יעד"))
                    if pd.isna(start_val) or str(start_val).strip() == "":
                        start_val = row.get("תאריך יעד")
                    if pd.isna(start_val):
                        continue

                    start_str = pd.to_datetime(start_val).strftime("%Y-%m-%d")
                    end_val = row.get("תאריך יעד")
                    end_str = pd.to_datetime(end_val).strftime("%Y-%m-%d") if not pd.isna(end_val) else start_str
                    task_name = str(row.get("שם משימה", "") or "").strip()
                    project = str(row.get("פרויקט", "") or "").strip()
                    _cal_color = (
                        CALENDAR_TASK_COLOR_OOO
                        if _is_out_of_office_task(row)
                        else CALENDAR_TASK_COLOR_DEFAULT
                    )
                    if _is_out_of_office_task(row):
                        _ev_title = _ooo_event_title_from_row(row)
                    else:
                        _ev_title = f"{task_name} - {project}" if project else task_name or "משימה"

                    calendar_events.append({
                        "title": _ev_title,
                        "start": start_str,
                        "end": end_str,
                        "color": _cal_color,
                    })
                except Exception:
                    pass  # דילוג שקט על משימות עם תאריך לא תקין
        except Exception as e:
            st.warning(f"בעיה בשליפת המשימות: {e}")

        # רינדור הלוח - עם טיפול בשגיאות
        try:
            cal_options = {
                "headerToolbar": {
                    "left": "today prev,next",
                    "center": "title",
                    "right": "dayGridMonth,timeGridWeek,timeGridDay",
                },
                "initialView": "dayGridMonth",
                "direction": "rtl",
            }
            st.markdown("<br><br><br>", unsafe_allow_html=True)
            calendar(events=calendar_events, options=cal_options)
        except Exception as e:
            st.error(f"שגיאה בהצגת נתונים: {e}")


def show_contacts_page() -> None:
    """לשונית CRM - ניהול אנשי קשר (לקוחות, אדריכלים וכו')."""
    st.title("👥 לקוחות ואנשי קשר")
    st.markdown(
        "כאן מנהלים את ספר הטלפונים של הסטודיו. הנתונים מכאן ישמשו לשליחת חומרים וחשבוניות."
    )
    st.markdown("---")

    # --- אנשי קשר לפרויקט (גיליון project_contacts) ---
    st.divider()
    st.subheader("📞 אנשי קשר לפרויקט")
    projects_options = _get_active_projects_options()
    if projects_options:
        with st.expander("➕ הוסף איש קשר", expanded=True):
            with st.form("add_contact_form"):
                contact_project = st.selectbox(
                    "פרויקט",
                    options=projects_options,
                    key="contact_project",
                )
                role_category = st.selectbox(
                    "סוג הגורם",
                    options=ROLE_CATEGORIES,
                    key="contact_role",
                )
                office_name = st.text_input(
                    "שם המשרד/החברה",
                    placeholder="למשל: יסקי מור סיון אדריכלים (לקוח פרטי: השאר ריק או כתוב 'פרטי')",
                    key="contact_office",
                )
                contact_name = st.text_input(
                    "שם איש הקשר",
                    placeholder="למשל: דנה מנהלת הפרויקט",
                    key="contact_name",
                )
                contact_email = st.text_input("מייל", key="contact_email")
                contact_phone = st.text_input("טלפון", key="contact_phone")
                contact_notes = st.text_area("הערות", key="contact_notes", height=80)
                add_contact_btn = st.form_submit_button("הוסף איש קשר")

            if add_contact_btn:
                if not contact_name.strip():
                    st.warning("נא להזין שם איש קשר.")
                elif not contact_project:
                    st.warning("נא לבחור פרויקט.")
                else:
                    existing = read_project_contacts()
                    row = {
                        "Project": contact_project,
                        "Role Category": role_category,
                        "Office/Company Name": (office_name or "").strip(),
                        "Contact Name": contact_name.strip(),
                        "Email": (contact_email or "").strip(),
                        "Phone": (contact_phone or "").strip(),
                        "Notes": (contact_notes or "").strip(),
                    }
                    existing.append(row)
                    write_project_contacts(existing)
                    st.success("איש הקשר נוסף בהצלחה!")
                    st.rerun()

        # תצוגה מקובצת לפי סוג הגורם
        contacts_rows = read_project_contacts()
        had_any_contacts = bool(contacts_rows)
        if contacts_rows:
            filter_project = st.selectbox(
                "סנן לפי פרויקט",
                options=["הכל"] + projects_options,
                key="contacts_filter_project",
            )
            if filter_project != "הכל":
                contacts_rows = [r for r in contacts_rows if r.get("Project") == filter_project]

        if contacts_rows:
            # מיון וקיבוץ לפי Role Category
            df_contacts = pd.DataFrame(contacts_rows, columns=PROJECT_CONTACTS_COLUMNS)
            df_contacts = df_contacts.sort_values(
                ["Role Category", "Office/Company Name", "Contact Name"],
                ascending=[True, True, True],
            )
            for role in ROLE_CATEGORIES:
                group_df = df_contacts[df_contacts["Role Category"] == role]
                if not group_df.empty:
                    with st.expander(f"**{role}** ({len(group_df)} אנשי קשר)", expanded=True):
                        st.dataframe(
                            group_df,
                            hide_index=True,
                            use_container_width=True,
                            column_config={
                                "Project": st.column_config.TextColumn("פרויקט"),
                                "Role Category": st.column_config.TextColumn("סוג הגורם"),
                                "Office/Company Name": st.column_config.TextColumn("משרד/חברה"),
                                "Contact Name": st.column_config.TextColumn("איש קשר"),
                                "Email": st.column_config.TextColumn("מייל"),
                                "Phone": st.column_config.TextColumn("טלפון"),
                                "Notes": st.column_config.TextColumn("הערות"),
                            },
                        )
        else:
            if had_any_contacts:
                st.info("אין אנשי קשר לפרויקט שנבחר. נסה לבחור 'הכל' או פרויקט אחר.")
            else:
                st.info("אין אנשי קשר. הוסף איש קשר חדש למעלה.")
    else:
        st.info("אין פרויקטים פעילים. הוסף פרויקטים כדי לנהל אנשי קשר.")

    st.divider()

    with st.form("הוספת איש קשר חדש", clear_on_submit=False):
        col1, col2 = st.columns(2)
        with col1:
            form_name = st.text_input("שם מלא", key="contact_form_full_name")
            form_company = st.text_input("חברה / משרד אדריכלים", key="contact_form_company")
            form_role = st.text_input("תפקיד", key="contact_form_role")
        with col2:
            form_phone = st.text_input("טלפון", key="contact_form_phone")
            form_email = st.text_input("אימייל", key="contact_form_email")
            form_type = st.selectbox(
                "סוג איש קשר",
                options=CONTACT_TYPE_OPTIONS,
                key="contact_form_type",
            )
        submitted = st.form_submit_button("שמור איש קשר", type="primary")
        if submitted:
            if not (form_name or "").strip():
                st.warning("נא להזין לפחות שם מלא.")
            else:
                try:
                    df = read_contacts_sheet()
                    new_row = {
                        "שם מלא": (form_name or "").strip(),
                        "חברה / משרד אדריכלים": (form_company or "").strip(),
                        "תפקיד": (form_role or "").strip(),
                        "טלפון": (form_phone or "").strip(),
                        "אימייל": (form_email or "").strip(),
                        "סוג איש קשר": (form_type or "").strip() or CONTACT_TYPE_OPTIONS[0],
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    save_contacts(df)
                except Exception as e:
                    st.error(f"שגיאה בהוספה: {e}")

    df_for_edit = read_contacts_sheet()
    if not df_for_edit.empty:

        def _crm_contact_row_label(i: int) -> str:
            name = str(df_for_edit.iloc[i].get("שם מלא", "") or "").strip()
            company = str(df_for_edit.iloc[i].get("חברה / משרד אדריכלים", "") or "").strip()
            if name and company:
                return f"{name} - {company}"
            if name:
                return name
            if company:
                return company
            return f"(שורה {i + 1})"

        st.markdown("### ✏️ עריכת איש קשר קיים")
        _row_options = list(range(len(df_for_edit)))
        selected_edit_idx = st.selectbox(
            "בחר איש קשר לעריכה",
            options=_row_options,
            format_func=_crm_contact_row_label,
            key="crm_edit_contact_select",
        )
        _edit_row = df_for_edit.iloc[selected_edit_idx]
        _current_type = str(_edit_row.get("סוג איש קשר", "") or "").strip()
        _type_opts = list(CONTACT_TYPE_OPTIONS)
        if _current_type and _current_type not in _type_opts:
            _type_opts = [_current_type] + _type_opts
        _type_default_idx = _type_opts.index(_current_type) if _current_type in _type_opts else 0

        with st.form("crm_edit_contact_form"):
            _ec1, _ec2 = st.columns(2)
            with _ec1:
                edit_name = st.text_input(
                    "שם מלא",
                    value=str(_edit_row.get("שם מלא", "") or ""),
                    key=f"crm_edit_full_name_{selected_edit_idx}",
                )
                edit_company = st.text_input(
                    "חברה / משרד אדריכלים",
                    value=str(_edit_row.get("חברה / משרד אדריכלים", "") or ""),
                    key=f"crm_edit_company_{selected_edit_idx}",
                )
                edit_role = st.text_input(
                    "תפקיד",
                    value=str(_edit_row.get("תפקיד", "") or ""),
                    key=f"crm_edit_role_{selected_edit_idx}",
                )
            with _ec2:
                edit_phone = st.text_input(
                    "טלפון",
                    value=str(_edit_row.get("טלפון", "") or ""),
                    key=f"crm_edit_phone_{selected_edit_idx}",
                )
                edit_email = st.text_input(
                    "אימייל",
                    value=str(_edit_row.get("אימייל", "") or ""),
                    key=f"crm_edit_email_{selected_edit_idx}",
                )
                edit_type = st.selectbox(
                    "סוג איש קשר",
                    options=_type_opts,
                    index=_type_default_idx,
                    key=f"crm_edit_type_{selected_edit_idx}",
                )
            update_submitted = st.form_submit_button("עדכן פרטי איש קשר")

        if update_submitted:
            if not (edit_name or "").strip():
                st.warning("נא להזין לפחות שם מלא.")
            else:
                try:
                    df_upd = read_contacts_sheet()
                    if df_upd.empty or selected_edit_idx >= len(df_upd):
                        st.error("איש הקשר לא נמצא. רענן את הדף ונסה שוב.")
                    else:
                        _c = df_upd.columns.get_loc
                        df_upd.iloc[selected_edit_idx, _c("שם מלא")] = (edit_name or "").strip()
                        df_upd.iloc[selected_edit_idx, _c("חברה / משרד אדריכלים")] = (edit_company or "").strip()
                        df_upd.iloc[selected_edit_idx, _c("תפקיד")] = (edit_role or "").strip()
                        df_upd.iloc[selected_edit_idx, _c("טלפון")] = (edit_phone or "").strip()
                        df_upd.iloc[selected_edit_idx, _c("אימייל")] = (edit_email or "").strip()
                        df_upd.iloc[selected_edit_idx, _c("סוג איש קשר")] = (edit_type or "").strip() or CONTACT_TYPE_OPTIONS[0]
                        save_contacts(df_upd)
                except Exception as e:
                    st.error(f"שגיאה בעדכון: {e}")

    st.subheader("ספר טלפונים")
    df_display = read_contacts_sheet()
    search_q = st.text_input(
        "🔍 חיפוש לפי שם, חברה או תפקיד",
        key="contacts_phonebook_search",
        placeholder="הקלד לסינון…",
    )
    if df_display.empty:
        st.info("אין עדיין אנשי קשר. הוסף איש קשר באמצעות הטופס למעלה.")
    else:
        q = (search_q or "").strip().lower()
        if q:
            mask = (
                df_display["שם מלא"].astype(str).str.lower().str.contains(q, na=False)
                | df_display["חברה / משרד אדריכלים"].astype(str).str.lower().str.contains(q, na=False)
                | df_display["תפקיד"].astype(str).str.lower().str.contains(q, na=False)
            )
            df_filtered = df_display[mask].copy()
        else:
            df_filtered = df_display
        if df_filtered.empty and q:
            st.caption("לא נמצאו תוצאות לחיפוש. נסה מילה אחרת או נקה את החיפוש.")
        st.dataframe(df_filtered, hide_index=True, use_container_width=True)

    st.markdown("---")

    # ייבוא מרוכז מקובץ CSV
    with st.expander("📥 ייבוא אנשי קשר מקובץ (Outlook / Excel)"):
        st.markdown(
            "העלה קובץ CSV או Excel עם אנשי קשר. ודא ששמות העמודות בקובץ תואמים לעמודות במערכת "
            "(שם מלא, חברה / משרד אדריכלים, תפקיד, טלפון, אימייל, סוג איש קשר)."
        )
        uploaded_file = st.file_uploader(
            "בחר קובץ CSV או Excel", type=["csv", "xlsx"], key="contacts_import_uploader"
        )
        if uploaded_file is not None:
            try:
                if uploaded_file.name.lower().endswith(".xlsx"):
                    df_temp = pd.read_excel(uploaded_file, engine="openpyxl")
                else:
                    try:
                        df_temp = pd.read_csv(uploaded_file, encoding="utf-8")
                    except UnicodeDecodeError:
                        df_temp = pd.read_csv(uploaded_file, encoding="utf-8-sig")
                if df_temp.empty:
                    st.warning("הקובץ ריק. אין נתונים לייבא.")
                else:
                    # נרמול שמות עמודות (הסרת רווחים)
                    df_temp.columns = [str(c).strip() for c in df_temp.columns]
                    # מיפוי עמודות נפוצות מ-Outlook/Excel
                    _COL_MAP = {
                        "Display Name": "שם מלא",
                        "Name": "שם מלא",
                        "Email Address": "אימייל",
                        "Email": "אימייל",
                        "Company": "חברה / משרד אדריכלים",
                        "Job Title": "תפקיד",
                        "Business Phone": "טלפון",
                        "Phone": "טלפון",
                        "Mobile Phone": "טלפון",
                    }
                    for old, new in _COL_MAP.items():
                        if old in df_temp.columns and new not in df_temp.columns:
                            df_temp = df_temp.rename(columns={old: new})
                    # שמירה רק על עמודות שמוכרות במערכת
                    valid_cols = [c for c in CONTACTS_COLUMNS if c in df_temp.columns]
                    if not valid_cols:
                        st.error(
                            "לא נמצאו עמודות תואמות. נא לוודא שהקובץ מכיל לפחות אחת מהעמודות: "
                            "שם מלא, אימייל, חברה / משרד אדריכלים, תפקיד, טלפון, סוג איש קשר."
                        )
                    else:
                        df_temp = df_temp[valid_cols].copy()
                        for col in CONTACTS_COLUMNS:
                            if col not in df_temp.columns:
                                df_temp[col] = ""
                        df_temp = df_temp.reindex(columns=CONTACTS_COLUMNS, fill_value="").fillna("").astype(str)
                        st.caption("תצוגה מקדימה:")
                        st.dataframe(df_temp.head(), use_container_width=True)
                        if st.button("ייבא נתונים לספר הטלפונים", type="primary", key="contacts_import_btn"):
                            df_existing = read_contacts_sheet()
                            df_combined = pd.concat([df_existing, df_temp], ignore_index=True)
                            # הסרת כפילויות: שורות עם אימייל - לפי אימייל; שורות בלי אימייל - לפי שם מלא
                            has_email = (df_combined["אימייל"].astype(str).str.strip() != "")
                            df_with_email = df_combined[has_email].drop_duplicates(subset=["אימייל"], keep="first")
                            df_no_email = df_combined[~has_email].drop_duplicates(subset=["שם מלא"], keep="first")
                            df_combined = pd.concat([df_with_email, df_no_email], ignore_index=True)
                            save_contacts(df_combined)
                            st.success(f"יובאו {len(df_combined) - len(df_existing)} אנשי קשר בהצלחה! הטבלה מתעדכנת.")
                            st.rerun()
            except pd.errors.EmptyDataError:
                st.warning("הקובץ ריק או פגום. אין נתונים לייבא.")
            except Exception as e:
                st.error(f"שגיאה בקריאת הקובץ: {e}")


def _validate_credentials(username_input: str, password_input: str) -> tuple[bool, str | None, str | None]:
    """
    מאמת שם משתמש וסיסמה מול st.secrets (מבנה שטוח: passwords, roles).
    מחזיר (success, role, error_message).
    """
    username = username_input.strip().lower()
    if not username:
        return (False, None, "שם משתמש או סיסמה שגויים")

    passwords = st.secrets.get("passwords", {}) or {}
    roles = st.secrets.get("roles", {}) or {}
    correct_password = passwords.get(username, None)
    user_role = roles.get(username, None)

    if correct_password and str(password_input).strip() == str(correct_password).strip():
        return (True, (user_role or "team").strip() or "team", None)
    return (False, None, "שם משתמש או סיסמה שגויים")


def show_login_screen() -> None:
    """
    מסך התחברות - טופס שם משתמש וסיסמה.
    מאמת מול st.secrets (מבנה שטוח: passwords, roles) ושומר ב-session_state: logged_in, username, role.
    """
    if st.session_state.get("logged_in"):
        return

    st.title("🔐 כניסה למערכת ניהול סטודיו")
    st.markdown("---")

    username_input = st.text_input("שם משתמש", key="login_username", placeholder="הזן שם משתמש")
    password_input = st.text_input("סיסמה", type="password", key="login_password", placeholder="הזן סיסמה")

    if st.button("התחבר", type="primary", key="login_submit"):
        try:
            success, role, _ = _validate_credentials(username_input, password_input)
            if success:
                username_normalized = username_input.strip().lower()
                st.session_state["logged_in"] = True
                st.session_state["username"] = username_normalized
                st.session_state["role"] = role or "team"
                st.session_state["current_user"] = _session_current_user_from_login(username_normalized)
                st.rerun()
            else:
                st.error("שם משתמש או סיסמה שגויים")
        except Exception:
            st.error("שם משתמש או סיסמה שגויים")


def _get_assignee_for_current_user() -> str:
    """מחזיר את ה-Assignee המתאים למשתמש הנוכחי (לסינון משימות) — רק מ-session current_user (כניסה)."""
    return (st.session_state.get("current_user") or "").strip() or "צוות"


def is_admin() -> bool:
    """מזהה אם המשתמש המחובר הוא מנהל (role == 'manager')."""
    return st.session_state.get("role") == "manager"


def _user_in_team(team_value: str, current_user: str) -> bool:
    """בודק אם שם המשתמש מופיע בעמודת Team (תומך ברשימה מופרדת בפסיקים)."""
    if not current_user or not team_value:
        return False
    safe_user = clean_name_for_match(current_user)
    safe_team = clean_name_for_match(team_value)
    return bool(safe_user) and safe_user in safe_team


def _render_daily_tasks_editor(assignee: str, key_prefix: str = "daily_tasks") -> None:
    """
    מציג data_editor למשימות יומיות של assignee מסוים.
    מסנן משימות שלא הושלמו (Is Done == False/0), תאריך <= היום.
    בעת סימון וי - מעדכן את הקובץ ומריץ rerun.
    """
    all_tasks = read_daily_tasks()
    today_str = date.today().strftime("%Y-%m-%d")

    # סינון: Assignee תואם, ואז סינון חכם (תאריך <= היום, לא הושלמו)
    assignee_tasks = [r for r in all_tasks if (r.get("Assignee") or "").strip() == assignee]
    pending = _filter_daily_tasks_by_date(assignee_tasks, date_col="Date", status_col="Status")

    if not pending:
        st.info("אין משימות ממתינות להיום. כל הכבוד! 🎉")
        return

    # המרת Is Done לבוליאני לתצוגה (CheckboxColumn דורש bool)
    today_dt = date.today()
    for r in pending:
        r["Is Done"] = False

    df = pd.DataFrame(pending)
    df = df.fillna('')
    # הוספת אינדיקציה לאיחור (🔴 למשימות שנגררו)
    def _ov(row):
        d = _parse_date_safe(row.get("Date"), "Date")
        return "🔴" if d is None or d < today_dt else ""
    df.insert(0, "איחור", df.apply(_ov, axis=1))
    if "Is Done" not in df.columns:
        df["Is Done"] = False

    # המרת Flexible לבוליאני (תאימות לאחור: עמודה חסרה = False)
    def _is_flexible(val) -> bool:
        if val is True or val == 1:
            return True
        if isinstance(val, str) and (val or "").strip().lower() in ("true", "1", "yes"):
            return True
        return False

    for r in pending:
        r["Flexible"] = _is_flexible(r.get("Flexible"))

    column_config = {
        "איחור": st.column_config.TextColumn("איחור", disabled=True),
        "Task Name": st.column_config.TextColumn("שם המשימה", disabled=True),
        "Project": st.column_config.TextColumn("פרויקט", disabled=True),
        "Assignee": st.column_config.TextColumn("אחראי", disabled=True),
        "Date": st.column_config.TextColumn("תאריך", disabled=True),
        "Status": st.column_config.TextColumn("סטטוס", disabled=True),
        "Is Done": st.column_config.CheckboxColumn("בוצע", default=False),
        "Flexible": st.column_config.CheckboxColumn("גמיש", disabled=True, default=False),
    }

    edited_df = st.data_editor(
        df,
        hide_index=True,
        use_container_width=True,
        column_config=column_config,
        key=f"{key_prefix}_{assignee}",
    )

    # בדיקה אם שונה Is Done ל-True - עדכון הקובץ
    if not edited_df.equals(df):
        for idx, row in edited_df.iterrows():
            if row.get("Is Done") is True:
                task_name = (row.get("Task Name") or "").strip()
                proj = (row.get("Project") or "").strip()
                orig_assignee = (row.get("Assignee") or "").strip()
                dt = (row.get("Date") or today_str).strip()
                # עדכון המשימה ב-all_tasks ל-Is Done = 1
                for t in all_tasks:
                    if (
                        (t.get("Task Name") or "").strip() == task_name
                        and (t.get("Project") or "").strip() == proj
                        and (t.get("Assignee") or "").strip() == orig_assignee
                        and (t.get("Date") or "").strip() == dt
                    ):
                        t["Is Done"] = "1"
                        break
                write_daily_tasks(all_tasks)
                st.success("המשימה סומנה כהושלמה!")
                st.rerun()


def show_my_work_page() -> None:
    """לשונית 'העבודה שלי' - לעובדים בלבד. מציג משימות פתוחות אישיות בכרטיסיות."""
    _render_quick_comm_notifications()
    current_user = (st.session_state.get("current_user") or "צוות").strip()
    st.title("העבודה שלי")
    st.subheader(f"👋 בוקר טוב {current_user}! הנה המשימות שלך:")
    st.markdown("---")

    all_tasks = read_daily_tasks()
    if not all_tasks:
        st.info("אין משימות. כל הכבוד! 🎉")
        return

    safe_current_user = clean_name_for_match(current_user)

    # סינון: משימות של המשתמש הנוכחי
    user_tasks_raw = [
        t for t in all_tasks
        if bool(safe_current_user) and safe_current_user in clean_name_for_match(t.get("Assignee") or "")
    ]
    # סינון חכם: תאריך <= היום, לא הושלמו. מיון: באיחור ראשון, אחר כך היום.
    user_tasks_filtered = _filter_daily_tasks_by_date(user_tasks_raw, date_col="Date", status_col="Status")

    if not user_tasks_filtered:
        st.info("אין משימות פתוחות. כל הכבוד! 🎉")
        return

    df_user_tasks = pd.DataFrame(user_tasks_filtered)

    def _mark_task_done(task_name: str, project: str, assignee: str, date_str: str) -> None:
        for t in all_tasks:
            if (
                (t.get("Task Name") or "").strip() == (task_name or "").strip()
                and (t.get("Project") or "").strip() == (project or "").strip()
                and (t.get("Assignee") or "").strip() == (assignee or "").strip()
                and (t.get("Date") or "").strip() == (date_str or "").strip()
                and not _is_done_daily(t.get("Is Done"))
            ):
                t["Is Done"] = "1"
                break
        write_daily_tasks(all_tasks)
        st.success("המשימה סומנה כהושלמה!")
        st.rerun()

    # תצוגה בכרטיסיות (באיחור ראשון, אחר כך היום)
    today_dt = date.today()
    for i, (_, row) in enumerate(df_user_tasks.iterrows()):
        task_name = (row.get("Task Name") or "").strip()
        project = (row.get("Project") or "").strip()
        assignee = (row.get("Assignee") or "").strip()
        date_str = (row.get("Date") or "").strip()
        status = (row.get("Status") or "").strip()
        task_date = _parse_date_safe(row.get("Date"), "Date")
        is_overdue = task_date is None or task_date < today_dt
        overdue_badge = "🔴 " if is_overdue else ""

        with st.container():
            col_btn, col_content = st.columns([0.1, 0.9])
            with col_btn:
                if st.button("✅ סיום", key=f"my_work_done_{i}"):
                    _mark_task_done(task_name, project, assignee, date_str)
            with col_content:
                border_color = "#dc3545" if is_overdue else "#0d6efd"
                st.markdown(
                    f"""
                <div style="
                    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
                    border-radius: 8px;
                    padding: 12px 16px;
                    margin-bottom: 8px;
                    border-right: 4px solid {border_color};
                    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
                ">
                    <strong>{overdue_badge}{html.escape(task_name)}</strong><br>
                    <span style="color:#495057; font-size:0.9em;">📁 {html.escape(project)}</span><br>
                    <span style="color:#6c757d; font-size:0.85em;">📅 תאריך יעד: {html.escape(date_str)}</span>
                    {f'<br><span style="color:#6c757d; font-size:0.85em;">סטטוס: {html.escape(status)}</span>' if status else ''}
                </div>
                """,
                    unsafe_allow_html=True,
                )


def show_daily_tasks_page() -> None:
    """לשונית 'משימות יומיות' - למנהלים. בחירת איש צוות לצפייה במשימות שלו."""
    st.title("משימות יומיות")
    st.markdown("בחר איש צוות כדי לראות את המשימות שלו:")

    selected_assignee = st.selectbox(
        "בחר איש צוות להצגת משימות:",
        options=TEAM_DISPLAY_NAMES,
        key="daily_tasks_assignee_select",
    )

    st.subheader(f"המוניטור של {selected_assignee}")

    # --- מוניטור משימות אישי (טפול בתאריכים יציב עם pd.to_datetime) ---
    def _assignee_matches(task_assignee: str, sel: str) -> bool:
        a = (task_assignee or "").strip()
        return a == sel or a.startswith(sel + " ") or a.startswith(sel + "-")

    def _is_done_monitor(val) -> bool:
        if val is True or val == 1 or (isinstance(val, str) and (val or "").strip().lower() in ("true", "1", "yes")):
            return True
        return False

    all_tasks_monitor = read_daily_tasks()
    today = pd.to_datetime("today").normalize()

    # סינון: משימות של assignee
    assignee_tasks = [
        t for t in all_tasks_monitor
        if _assignee_matches(t.get("Assignee") or "", selected_assignee)
    ]
    # סינון חכם: תאריך <= היום, לא הושלמו. מיון: באיחור ראשון, אחר כך היום.
    filtered_list = _filter_daily_tasks_by_date(assignee_tasks, date_col="Date", status_col="Status")

    # חלוקה ל-2 קבוצות: overdue (תאריך < היום), today (תאריך == היום)
    overdue_tasks = []
    today_tasks = []
    for t in filtered_list:
        r = t.copy()
        task_date = _parse_date_safe(t.get("Date"), "Date")
        if task_date is None:
            overdue_tasks.append((r, 0))
        elif task_date < date.today():
            days_overdue = (date.today() - task_date).days
            overdue_tasks.append((r, days_overdue))
        else:
            today_tasks.append(r)

    def _mark_task_done(task_name: str, project: str, date_str: str) -> None:
        for t in all_tasks_monitor:
            if (
                (t.get("Task Name") or "").strip() == (task_name or "").strip()
                and (t.get("Project") or "").strip() == (project or "").strip()
                and _assignee_matches(t.get("Assignee") or "", selected_assignee)
                and (t.get("Date") or "").strip() == (date_str or "").strip()
                and not _is_done_monitor(t.get("Is Done"))
            ):
                t["Is Done"] = "1"
                break
        write_daily_tasks(all_tasks_monitor)
        st.rerun()

    # תצוגה: overdue גמיש ב-st.warning, overdue רגיל ב-st.error, today ב-st.info
    if overdue_tasks:
        st.markdown("#### באיחור")
        for idx, (r, days_overdue) in enumerate(overdue_tasks):
            task_name = (r.get("Task Name") or "").strip()
            project = (r.get("Project") or "").strip()
            date_str = (r.get("Date") or "").strip()
            days_text = f"באיחור של {days_overdue} ימים!" if days_overdue > 1 else "באיחור של יום אחד!"
            is_flex = (r.get("Flexible") or "").strip().lower() in ("1", "true", "yes")
            col_btn, col_msg = st.columns([0.12, 0.88])
            with col_btn:
                if st.button("✅ סיום", key=f"done_ov_{idx}"):
                    _mark_task_done(task_name, project, date_str)
            with col_msg:
                if is_flex:
                    st.warning(f"🔴 ⏳ גמיש: **{task_name}** – {project} – {date_str} ({days_text})")
                else:
                    st.error(f"🔴 **{task_name}** – {project} – {date_str} ({days_text})")

    if today_tasks:
        st.markdown("#### להיום")
        for idx, r in enumerate(today_tasks):
            task_name = (r.get("Task Name") or "").strip()
            project = (r.get("Project") or "").strip()
            date_str = (r.get("Date") or "").strip()
            col_btn, col_msg = st.columns([0.12, 0.88])
            with col_btn:
                if st.button("✅ סיום", key=f"done_today_{idx}"):
                    _mark_task_done(task_name, project, date_str)
            with col_msg:
                st.info(f"**{task_name}** – {project} – {date_str}")

    if not overdue_tasks and not today_tasks:
        st.success(f"אין משימות פתוחות ל־{selected_assignee}. כל הכבוד! 🎉")

    st.markdown("---")

    # גשר ממוניטור: מילוי מראש מהפרויקט שנבחר
    if "bridge_project_name" in st.session_state:
        st.session_state["new_daily_task_project"] = st.session_state["bridge_project_name"]
    assignee_index = 0
    if "bridge_assignee" in st.session_state:
        bridge_assignee = st.session_state["bridge_assignee"]
        bridge_short = (bridge_assignee or "").split('-')[0].strip()
        assignee_index = next((i for i, name in enumerate(TEAM_DISPLAY_NAMES) if name == bridge_short or (bridge_short and bridge_short in name)), 0)

    with st.expander("➕ הוסף משימה יומית", expanded=bool(st.session_state.get("bridge_project_name"))):
        with st.form("add_daily_task_form"):
            new_task_name = st.text_input("שם המשימה", key="new_daily_task_name")
            new_project = st.text_input("פרויקט", key="new_daily_task_project")
            new_assignee = st.selectbox("אחראי", options=TEAM_DISPLAY_NAMES, key="new_daily_task_assignee", index=assignee_index)
            new_date = st.date_input("תאריך", value=date.today(), key="new_daily_task_date")
            new_flexible = st.checkbox("דדליין גמיש (ניתן לדחייה בעת עומס)", key="new_daily_task_flexible")
            if st.form_submit_button("הוסף משימה"):
                if (new_task_name or "").strip():
                    all_tasks = read_daily_tasks()
                    all_tasks.append({
                        "Task Name": (new_task_name or "").strip(),
                        "Project": (new_project or "").strip(),
                        "Assignee": new_assignee,
                        "Date": new_date.strftime("%Y-%m-%d"),
                        "Status": "",
                        "Is Done": "0",
                        "Flexible": "1" if new_flexible else "0",
                    })
                    write_daily_tasks(all_tasks)
                    for k in ("bridge_project_name", "bridge_assignee"):
                        st.session_state.pop(k, None)
                    st.success("המשימה נוספה!")
                    st.rerun()
                else:
                    st.warning("נא להזין שם משימה.")

    st.markdown("---")
    _render_daily_tasks_editor(selected_assignee, key_prefix="admin_daily")


def main() -> None:
    # בדיקת התחברות - אם לא מחובר: הסתר תפריט צד והצג רק טופס התחברות
    if not st.session_state.get("logged_in"):
        show_login_screen()
        return

    # אתחול משתני מצב למוניטור (Drill-down)
    if "monitor_filter" not in st.session_state:
        st.session_state.monitor_filter = "__ALL__"  # זמנית: הצגת כל הפרויקטים כברירת מחדל
    if "monitor_title" not in st.session_state:
        st.session_state.monitor_title = "כל הפרויקטים (ללא סינון)"

    st.sidebar.title("תפריט ניהול")
    current_user = (st.session_state.get("current_user") or "").strip()
    st.sidebar.markdown(f"**מחובר/ת כעת:** {current_user or '—'}")
    is_management = current_user in MANAGEMENT_USERS
    st.session_state["is_management"] = is_management

    if is_management:
        view_filter = st.sidebar.selectbox(
            "👁️ סינון משימות (תצוגת מנהל):",
            options=["הצג הכל"] + list(TEAM_EMAILS.keys()),
            index=0,
            key="sidebar_view_filter_task",
        )
    else:
        view_filter = (st.session_state.get("current_user") or "").strip()
    st.session_state["view_filter"] = view_filter

    st.sidebar.caption(
        "נתוני גוגל שיטס נטענים ממטמון עד 10 דקות; לאחר כל שמירה במערכת המטמון מתעדכן אוטומטית."
    )

    menu_options = [
        NAV_MAIN_PROJECT_ROOM,
        NAV_MY_TASKS,
        NAV_PROJECT_FOLDERS,
    ]
    if is_management:
        menu_options.extend(
            [
                NAV_MGMT_SEPARATOR,
                NAV_QUOTES_FINANCE,
                NAV_TASKS_PRODUCTION,
                NAV_CRM,
            ]
        )

    selected_page = st.sidebar.radio(
        "ניווט ראשי",
        menu_options,
        key="nav_main_primary",
    )
    if selected_page == NAV_MGMT_SEPARATOR:
        st.warning("אנא בחר באחת מהאפשרויות מתחת לאזור הניהול.")
        selected_page = st.session_state.get("last_valid_main_nav", menu_options[0])
    else:
        st.session_state["last_valid_main_nav"] = selected_page

    # טעינת נתוני פרויקטים לדיבאג (מצב 'רנטגן')
    projects_rows_debug = read_projects()
    df_projects = pd.DataFrame(projects_rows_debug, columns=PROJECTS_DB_COLUMNS)
    df_projects = df_projects.fillna("")

    if selected_page == NAV_MAIN_PROJECT_ROOM:
        _render_quick_comm_notifications()
        # אזור חדר המצב: מוניטור פרויקטים - כפתורים צידיים + תצוגת טבלאות
        stats = _compute_project_monitor_stats()
        st.sidebar.markdown("---")
        st.sidebar.markdown("### סינון לפי סטטוס (מוניטור)")
        active_sum = int(stats.get("active_sum", 0) or 0)
        feedback_sum = int(stats.get("feedback_sum", 0) or 0)
        frozen_sum = int(stats.get("frozen_sum", 0) or 0)
        completed_sum = int(stats.get("completed_sum", 0) or 0)
        if st.sidebar.button(f"🟢 פעילים: {stats['active']} (₪{active_sum:,})", use_container_width=True, key="monitor_active"):
            st.session_state.monitor_filter = ["בעבודה", "ממתין להתחלה"]
            st.session_state.monitor_title = "פרויקטים פעילים"
            st.rerun()
        if st.sidebar.button(f"🟠 בסבב פידבק: {stats['feedback']} (₪{feedback_sum:,})", use_container_width=True, key="monitor_feedback"):
            st.session_state.monitor_filter = ["נשלח לסבב הערות 1", "נשלח לסבב הערות 2", "ממתין לאדריכל/לקוח"]
            st.session_state.monitor_title = "בסבב פידבק / המתנה"
            st.rerun()
        if st.sidebar.button(f"❄️ הוקפאו: {stats['frozen']} (₪{frozen_sum:,})", use_container_width=True, key="monitor_frozen"):
            st.session_state.monitor_filter = ["הוקפא"]
            st.session_state.monitor_title = "פרויקטים שהוקפאו"
            st.rerun()
        if st.sidebar.button(f"🔵 הסתיימו (בגבייה): {stats['completed']} (₪{completed_sum:,})", use_container_width=True, key="monitor_completed"):
            st.session_state.monitor_filter = ["הסתיים", "חשבונית נשלחה", "שולם"]
            st.session_state.monitor_title = "הסתיימו (בגבייה)"
            st.rerun()
        if st.sidebar.button("📋 הצג הכל (Show All) - ללא סינון", use_container_width=True, key="monitor_show_all"):
            st.session_state.monitor_filter = "__ALL__"
            st.session_state.monitor_title = "כל הפרויקטים (ללא סינון)"
            st.rerun()
        st.sidebar.markdown("---")
        _render_quick_comm_sidebar_form()
        if st.sidebar.button('הצג נתונים גולמיים'):
            st.write(df_projects)
        _render_dropbox_access_token_hint_sidebar()

        st.markdown("---")
        _render_project_kanban_board()
        st.markdown("---")

        # תצוגת חתך ממוקדת (Drill-down) - במסך הראשי
        if st.session_state.monitor_filter is not None:
            with st.container(border=True):
                st.subheader(f"🔎 תצוגת חתך: {st.session_state.monitor_title}")
                projects_rows = read_projects()
                df_projects = pd.DataFrame(projects_rows, columns=PROJECTS_DB_COLUMNS)
                df_projects = df_projects.fillna('')
                # זמנית: תמיכה ב-Show All - הצגת כל הפרויקטים ללא סינון
                if st.session_state.monitor_filter == "__ALL__":
                    filtered_df = df_projects.copy()
                else:
                    # סינון סלחני - חסין לרווחים ואותיות (strip + case-insensitive לעמודת Status)
                    status_col = df_projects.get("Status")
                    if status_col is not None:
                        status_series = status_col.fillna("").astype(str).str.strip().str.lower()
                        filter_normalized = [s.strip().lower() for s in st.session_state.monitor_filter]
                        mask = status_series.isin(filter_normalized)
                    else:
                        mask = pd.Series([False] * len(df_projects), index=df_projects.index)
                    filtered_df = df_projects[mask]
                # זמנית: ביטול סינון לפי Team - הצגת כל הפרויקטים
                # if not is_admin():
                #     current_user = _get_assignee_for_current_user()
                #     team_col = filtered_df.get("Team")
                #     if team_col is not None:
                #         team_mask = team_col.fillna("").apply(lambda t: _user_in_team(str(t), current_user))
                #         filtered_df = filtered_df[team_mask]
                if filtered_df.empty:
                    st.info("לא נמצאו פרויקטים." if (st.session_state.monitor_filter == "__ALL__") else "לא נמצאו פרויקטים פעילים." if (st.session_state.monitor_filter == ["בעבודה", "ממתין להתחלה"]) else f"לא נמצאו פרויקטים בקטגוריה '{st.session_state.monitor_title}'.")
                else:
                    _kanban_map = _kanban_stage_lookup_by_client_project()

                    def _row_kanban_stage(row: pd.Series) -> str:
                        k = (
                            (row.get("Client") or "").strip(),
                            (row.get("Project Name") or "").strip(),
                        )
                        return _kanban_map.get(k, "")

                    filtered_df = filtered_df.copy()
                    filtered_df["שלב בקנבן"] = filtered_df.apply(_row_kanban_stage, axis=1)
                    _pi = PROJECTS_DB_COLUMNS.index("Project Name")
                    _slice_col_order = (
                        PROJECTS_DB_COLUMNS[: _pi + 1]
                        + ["שלב בקנבן"]
                        + PROJECTS_DB_COLUMNS[_pi + 1 :]
                    )
                    filtered_df = filtered_df[_slice_col_order]
                    filtered_for_edit = filtered_df.drop(columns=["Status"], errors="ignore")
                    _cols_compare = [c for c in filtered_for_edit.columns if c != "שלב בקנבן"]
                    try:
                        edited_filtered_df = st.data_editor(
                            filtered_for_edit,
                            hide_index=True,
                            use_container_width=True,
                            key="drilldown_editor",
                            column_config={
                                "שלב בקנבן": st.column_config.TextColumn(
                                    "שלב בקנבן",
                                    help="שלב עבודה מעודכן מלוח הקנבן (גיליון הצעות).",
                                    disabled=True,
                                ),
                                "Dropbox_Main": st.column_config.LinkColumn(
                                    "דרופבוקס - ראשי",
                                    display_text="תיקייה ראשית",
                                ),
                                "Dropbox_Upload": st.column_config.LinkColumn(
                                    "דרופבוקס - העלאה",
                                    display_text="בקשת קבצים",
                                ),
                                "Dropbox_Deliverables": st.column_config.LinkColumn(
                                    "דרופבוקס - תוצרים",
                                    display_text="תוצרים",
                                ),
                            },
                        )
                    except Exception as e:
                        st.error(f"שגיאה בהצגת נתונים: {e}")
                        edited_filtered_df = filtered_for_edit
                    if not edited_filtered_df[_cols_compare].equals(filtered_for_edit[_cols_compare]):
                        merged_edit = edited_filtered_df.copy()
                        merged_edit["Status"] = filtered_df["Status"].values
                        df_projects.update(merged_edit[PROJECTS_DB_COLUMNS])
                        updated = df_projects.to_dict(orient="records")
                        write_projects(updated)
                        st.success("הנתונים עודכנו בהצלחה!")
                        st.rerun()
        if st.button("✖️ סגור תצוגה ממוקדת", key="close_monitor_drill"):
            st.session_state.monitor_filter = None
            st.session_state.monitor_title = ""
            st.rerun()

    elif selected_page == NAV_PROJECT_FOLDERS:
        _render_quick_comm_notifications()
        _render_quick_comm_sidebar_form()
        _render_dropbox_access_token_hint_sidebar()
        show_project_folders_page()

    elif selected_page == NAV_MY_TASKS:
        _render_quick_comm_notifications()
        _render_quick_comm_sidebar_form()
        _render_dropbox_access_token_hint_sidebar()
        show_monitor_3d_page()

    elif selected_page == NAV_QUOTES_FINANCE:
        _render_quick_comm_notifications()
        _render_quick_comm_sidebar_form()
        if st.sidebar.button("הצג נתונים גולמיים"):
            st.write(df_projects)
        _render_dropbox_access_token_hint_sidebar()
        tab_quotes, tab_finance = st.tabs(["📝 ניהול הצעות מחיר", "💰 דשבורד פיננסי וגבייה"])
        with tab_quotes:
            quotes_mode = st.radio(
                "הצעות מחיר",
                ["יצירת הצעה חדשה", "ניהול הצעות"],
                horizontal=True,
                key="quotes_mode_inline",
            )
            if quotes_mode == "יצירת הצעה חדשה":
                show_quote_page()
            else:
                show_quotes_management_page()
        with tab_finance:
            _show_finance_collection_dashboard()

    elif selected_page == NAV_TASKS_PRODUCTION:
        _render_quick_comm_notifications()
        _render_quick_comm_sidebar_form()
        if st.sidebar.button("הצג נתונים גולמיים"):
            st.write(df_projects)
        _render_dropbox_access_token_hint_sidebar()
        show_tasks_page()

    elif selected_page == NAV_CRM:
        _render_quick_comm_notifications()
        _render_quick_comm_sidebar_form()
        if st.sidebar.button("הצג נתונים גולמיים"):
            st.write(df_projects)
        _render_dropbox_access_token_hint_sidebar()
        show_contacts_page()

    st.sidebar.markdown("---")
    if st.sidebar.button("התנתק 🚪", key="logout_btn_manager", use_container_width=True):
        for k in ("logged_in", "username", "role", "current_user", "is_management", "sidebar_current_user"):
            st.session_state.pop(k, None)
        st.rerun()


if __name__ == "__main__":
    main()

