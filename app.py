import html
import io
import re
import requests
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
import os
import csv
import json
import subprocess
import platform
from urllib.parse import quote

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
from oauth2client.service_account import ServiceAccountCredentials

# --- התחברות לגוגל שיטס ---
def init_connection():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_data = st.secrets["gcp_service_account"]
    creds_dict = json.loads(creds_data) if isinstance(creds_data, str) else creds_data
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds)

_sheets_init_error = None
try:
    client = init_connection()
    SHEET_ID = '1ZvAtkWaXpf9zZRgXY2HUcRB6QWpUMe6KWNjPu-eyzdo'
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


@st.cache_data(ttl=60)
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


def _render_dropbox_refresh_token_sidebar() -> None:
    """מציג בלוק בתחתית התפריט הצד לחילוץ Refresh Token - רק כאשר DROPBOX_REFRESH_TOKEN ריק או חסר."""
    refresh_token = st.secrets.get("DROPBOX_REFRESH_TOKEN", "")
    if refresh_token and str(refresh_token).strip():
        return
    app_key = st.secrets.get("DROPBOX_APP_KEY", "")
    app_secret = st.secrets.get("DROPBOX_APP_SECRET", "")
    if not app_key or not app_secret:
        st.sidebar.warning("נדרש חיבור קבוע לדרופבוקס – הגדר DROPBOX_APP_KEY ו-DROPBOX_APP_SECRET ב-Secrets.")
        return
    st.sidebar.markdown("---")
    st.sidebar.warning("נדרש חיבור קבוע לדרופבוקס")
    auth_url = f"https://www.dropbox.com/oauth2/authorize?client_id={app_key}&token_access_type=offline&response_type=code"
    st.sidebar.markdown(f"[1. לחץ כאן לקבלת קוד גישה לדרופבוקס]({auth_url})")
    auth_code = st.sidebar.text_input("2. הדבק את הקוד שקיבלת כאן:", key="dropbox_auth_code")
    if st.sidebar.button("3. הפק מפתח קבוע", key="dropbox_fetch_refresh"):
        if not auth_code or not str(auth_code).strip():
            st.sidebar.error("נא להדביק את קוד הגישה שקיבלת.")
        else:
            try:
                r = requests.post(
                    "https://api.dropboxapi.com/oauth2/token",
                    data={
                        "grant_type": "authorization_code",
                        "code": auth_code.strip(),
                        "client_id": app_key,
                        "client_secret": app_secret,
                    },
                )
                r.raise_for_status()
                data = r.json()
                rt = data.get("refresh_token")
                if rt:
                    st.sidebar.success("המפתח הקבוע הופק בהצלחה. העתק אותו ל-Streamlit Secrets:")
                    st.sidebar.code(rt, language=None)
                    st.sidebar.info("הוסף ל-Secrets: DROPBOX_REFRESH_TOKEN = \"<המפתח שהעתקת>\"")
                else:
                    st.sidebar.error("לא נמצא refresh_token בתשובה. ייתכן שהקוד פג תוקף – נסה שוב.")
            except requests.RequestException as e:
                st.sidebar.error(f"שגיאה בבקשה: {e}")


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
}
# שמות קצרים לתצוגה (ללא תפקידים) - משמש ב-selectbox, multiselect, וטקסט המייל
TEAM_DISPLAY_NAMES = [name.split('-')[0].strip() for name in TEAM_MEMBERS.keys()]
# מיפוי שם קצר -> מייל (לשליחת מייל)
TEAM_EMAIL_BY_SHORT = {name.split('-')[0].strip(): email for name, email in TEAM_MEMBERS.items()}
# מספר וואצאפ של ערן (placeholder - ניתן למלא/לשנות)
WHATSAPP_ERAN = "972547641984"

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
]

ALLOWED_QUOTE_STATUSES = ["Draft", "Sent", "Approved", "Revision Needed", "Rejected", "Signed"]
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

# --- projects.csv (פרויקטים פעילים) ---
PROJECTS_CSV_COLUMNS = ["ID", "Client", "Project", "Deadline", "Team", "Status", "Budget_Hours", 'היקף כספי (₪)', "אנשי קשר מקושרים"]
PROJECTS_CSV_STATUSES = ["Active", "Done"]

PROJECT_MANAGERS = ["ערן", "טלי"]
PROJECT_TEAM_MEMBERS = ["ג'ורג'", "מיה", "ליאור", "אור", "אחיעד"]
TASK_TEAM = ["ערן", "טלי", "ג'ורג'", "מיה", "ליאור", "אור", "אחיעד"]
TASKS_LOG_COLUMNS = [
    "Task ID",
    "Project",
    "Assignee",
    "Task Name",
    "Start Date",
    "Due Date",
    "Status",
    "Priority",
    "Notes",
    "Flexible",
]
TASK_STATUSES = ["To Do", "In Progress", "Done", "Stuck"]
TASK_PRIORITIES = ["רגיל", "דחוף", "קריטי"]

# סטטוסים שנחשבים "הושלם" - משימות עם סטטוס כזה לא יוצגו ברשימה
DONE_STATUSES = ("done", "בוצע", "הושלם", "completed")

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
CONTACT_TYPE_OPTIONS = ["אדריכל", "יזם/לקוח", "הנהלת חשבונות", "מפקח/אחר"]


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
        refresh_token = st.secrets.get("DROPBOX_REFRESH_TOKEN", "")
        if not refresh_token or not str(refresh_token).strip():
            return None
        dbx = dropbox.Dropbox(
            app_key=st.secrets["DROPBOX_APP_KEY"],
            app_secret=st.secrets["DROPBOX_APP_SECRET"],
            oauth2_refresh_token=refresh_token,
        )

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
        refresh_token = st.secrets.get("DROPBOX_REFRESH_TOKEN", "")
        if not refresh_token or not str(refresh_token).strip():
            return ""
        dbx = dropbox.Dropbox(
            app_key=st.secrets["DROPBOX_APP_KEY"],
            app_secret=st.secrets["DROPBOX_APP_SECRET"],
            oauth2_refresh_token=refresh_token,
        )
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


def _ensure_sheet(sheet_name: str, columns: list[str]) -> None:
    """וידוא שקיים גיליון בגוגל שיטס. לא יוצר גיליון חדש - תציג שגיאה אם חסר."""
    if spreadsheet is None:
        return
    spreadsheet.worksheet(sheet_name)  # יקרוס ויציג שגיאה אם הגיליון חסר


@st.cache_data(ttl=60)
def load_contacts() -> pd.DataFrame:
    """טוען אנשי קשר מגיליון contacts בגוגל שיטס."""
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
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת contacts: {e}")


@st.cache_data(ttl=60)
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


@st.cache_data(ttl=60)
def read_projects() -> list[dict]:
    """קריאת פרויקטים מגיליון projects בגוגל שיטס (מוניטור, Task Board)."""
    if spreadsheet is None:
        err_msg = str(_sheets_init_error) if _sheets_init_error else "אין חיבור לגוגל שיטס"
        st.error(f"שגיאת קריאה: {err_msg}")
        return []
    _ensure_sheet('projects', PROJECTS_DB_COLUMNS)
    try:
        worksheet = spreadsheet.worksheet('projects')
        df = _read_worksheet_safe(worksheet, PROJECTS_DB_COLUMNS)
        df.columns = df.columns.str.strip()
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


def write_projects(rows: list[dict]) -> None:
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
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת projects: {e}")


@st.cache_data(ttl=60)
def read_tasks() -> list[dict]:
    """קריאת משימות מגיליון tasks בגוגל שיטס (Task Board)."""
    if spreadsheet is None:
        return []
    _ensure_sheet('tasks', TASKS_LOG_COLUMNS)
    try:
        worksheet = spreadsheet.worksheet('tasks')
        df = _read_worksheet_safe(worksheet, TASKS_LOG_COLUMNS)
        if not df.empty and hasattr(df.columns, 'str'):
            df.columns = df.columns.str.strip()
        return df.to_dict(orient='records')
    except Exception as e:
        st.warning(f"שגיאה בקריאת tasks: {e}")
        return []


def write_tasks(rows: list[dict]) -> None:
    """שמירת משימות לגיליון tasks בגוגל שיטס."""
    if spreadsheet is None:
        err_msg = str(_sheets_init_error) if _sheets_init_error else "אין חיבור לגוגל שיטס"
        st.error(f"שגיאת שמירה: {err_msg}")
        return
    _ensure_sheet('tasks', TASKS_LOG_COLUMNS)
    try:
        worksheet = spreadsheet.worksheet('tasks')
        worksheet.clear()
        data = [TASKS_LOG_COLUMNS] + [[str(r.get(c, "") or "") for c in TASKS_LOG_COLUMNS] for r in rows]
        if data:
            worksheet.update(data, 'A1')
        st.cache_data.clear()
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת tasks: {e}")


def _ensure_tasks_csv_schema() -> None:
    """וידוא שקיים גיליון tasks בגוגל שיטס. לא יוצר גיליון חדש - תציג שגיאה אם חסר."""
    if spreadsheet is None:
        return
    spreadsheet.worksheet('tasks')  # יקרוס ויציג שגיאה אם הגיליון חסר


@st.cache_data(ttl=60)
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


def write_daily_tasks(rows: list[dict]) -> None:
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
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת tasks: {e}")


def append_kickoff_tasks_to_csv(
    project_display: str,
    assigned_team: list[str],
    project_template: list[str],
    task_deadline: date,
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
    write_daily_tasks(existing)


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
    write_projects(existing_rows)


def _ensure_projects_csv_schema() -> None:
    """וידוא שקיים גיליון projects בגוגל שיטס. לא יוצר גיליון חדש - תציג שגיאה אם חסר."""
    if spreadsheet is None:
        return
    spreadsheet.worksheet('projects')  # יקרוס ויציג שגיאה אם הגיליון חסר


@st.cache_data(ttl=60)
def read_projects_csv() -> list[dict]:
    """קריאת כל הפרויקטים מגוגל שיטס (גיליון projects)."""
    if spreadsheet is None:
        return []
    _ensure_projects_csv_schema()
    try:
        worksheet = spreadsheet.worksheet('projects')  # שם גיליון: projects (lowercase)
        df = _read_worksheet_safe(worksheet, PROJECTS_CSV_COLUMNS)
        if not df.empty and hasattr(df.columns, 'str'):
            df.columns = df.columns.str.strip()
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


def write_projects_csv(rows: list[dict]) -> None:
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
    write_projects_csv(existing)


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
                write_projects(rows)
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
    }


def read_quotes_log() -> list[dict]:
    """קריאת הצעות מגיליון quotes בגוגל שיטס - באותה צורה בטוחה כמו projects (fillna, ניקוי רווחים)."""
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


@st.cache_data(ttl=60)
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
            result.append(normalized)
        return result
    except Exception as e:
        st.error(f"שגיאה בשליפת נתונים (quotes): {e}")
        return []


def write_quotes_csv(rows: list[dict]) -> None:
    """Write full quote form data to Google Sheets (quotes tab)."""
    if spreadsheet is None:
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
        st.rerun()
    except Exception as e:
        st.warning(f"שגיאה בשמירת quotes: {e}")


def get_quote_from_csv(client: str, project: str, version: str) -> dict | None:
    """Get full quote row from quotes (Google Sheets) by (Client, Project, Version)."""
    key = _quote_key(client, project, version)
    for r in read_quotes_csv():
        if _quote_key(r.get("Client", ""), r.get("Project", ""), r.get("Version", "")) == key:
            return r
    return None


def append_quote_to_csv(row: dict) -> None:
    """Append a new quote row to quotes (Google Sheets)."""
    rows = read_quotes_csv()
    rows.append(row)
    write_quotes_csv(rows)


def update_quote_in_csv(client: str, project: str, version: str, updated_row: dict) -> bool:
    """Update existing quote in quotes (Google Sheets). Returns True if found and updated."""
    rows = read_quotes_csv()
    key = _quote_key(client, project, version)
    for i, r in enumerate(rows):
        if _quote_key(r.get("Client", ""), r.get("Project", ""), r.get("Version", "")) == key:
            rows[i] = {c: (updated_row.get(c) or "") for c in QUOTES_CSV_COLUMNS}
            write_quotes_csv(rows)
            return True
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


def open_folder_in_windows(path):
    """פותח תיקייה ב-Windows Explorer באמצעות os.startfile."""
    try:
        path = os.path.normpath(path)
        if os.path.exists(path):
            os.startfile(path)
        else:
            st.error(f"הנתיב לא נמצא: {path}")
    except Exception as e:
        st.error(f"לא ניתן לפתוח את התיקייה: {e}")


def open_file_in_explorer(file_path: str) -> None:
    """פותח את התיקייה ב-Windows Explorer ומסמן את הקובץ. משתמש ב-explorer /select."""
    open_folder_with_selection(file_path)


def open_folder_with_selection(file_path):
    """פותח את התיקייה ב-Windows Explorer ומסמן את הקובץ. תומך בנתיבים עם רווחים."""
    try:
        abs_path = os.path.abspath(file_path)
        abs_path = os.path.normpath(abs_path)
        if not os.path.exists(abs_path):
            st.error(f"הקובץ לא נמצא: {abs_path}")
            return
        # שימוש ב-f-string עם מרכאות עוטפות לנתיב - תומך בנתיבים עם רווחים
        subprocess.Popen(f'explorer /select,"{abs_path}"', shell=True)
    except Exception:
        st.error("לא הצלחתי לפתוח את הקובץ. מנסה לפתוח את התיקייה בלבד...")
        try:
            folder = os.path.dirname(os.path.abspath(file_path))
            os.startfile(folder)
        except Exception:
            st.error("שגיאה בפתיחת התיקייה.")


def open_folder(path):
    """Open the folder containing the file in the system file explorer."""
    folder_path = os.path.dirname(path)
    if platform.system() == "Windows":
        open_folder_in_windows(folder_path)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", folder_path])
    else:
        subprocess.Popen(["xdg-open", folder_path])


def open_file(path: str) -> None:
    """Open a file with the default system application."""
    file_path = (path or "").strip()
    if not file_path:
        raise ValueError("לא נמצא נתיב קובץ לפתיחה.")

    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(f"הקובץ לא נמצא: {p}")

    if platform.system() == "Windows":
        os.startfile(str(p))
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", str(p)])
    else:
        subprocess.Popen(["xdg-open", str(p)])


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

            # המרת Word ל-PDF (LibreOffice)
            pdf_path = quotes_dir / filename_pdf
            try:
                convert_to_pdf(str(output_path), str(pdf_path))
            except Exception as pdf_err:
                st.warning(f"המסמך נוצר, אך המרה ל-PDF נכשלה: {pdf_err}")

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
            st.session_state['current_quotes_dir'] = str(quotes_dir)
            st.session_state['current_display_path'] = str(display_path)
            st.session_state['current_email_client'] = client_email or ""
            st.session_state['current_contact_person'] = contact_person or ""
            st.session_state['current_project_name'] = project_name or ""
            st.session_state['current_quote_version'] = quote_version or ""
            st.session_state['current_client_name'] = client_name or ""
        except Exception as e:
            st.error(f"שגיאה ביצירת קובץ ה-Word: {e}")

    # הצגת התוצאה מחוץ ל-if - נשארת גם אחרי לחיצה על כפתורים
    if 'current_pdf_path' in st.session_state:
        quotes_dir_str = st.session_state.get('current_quotes_dir', '')
        st.success(f"הקובץ נוצר בהצלחה ונשמר ב:\n{quotes_dir_str}\n(Word + PDF)")

        # כפתור הורדת קובץ Word
        docx_path_str = st.session_state.get('current_docx_path', '')
        client_name_dl = st.session_state.get('current_client_name', 'Client')
        safe_client = sanitize_filename_part(client_name_dl) if client_name_dl else "Client"
        download_filename = f"Quote_{safe_client}.docx"
        if docx_path_str:
            docx_path = Path(docx_path_str)
            if docx_path.exists():
                docx_bytes = io.BytesIO(docx_path.read_bytes())
                st.download_button(
                    label="הורדת קובץ Word",
                    data=docx_bytes.getvalue(),
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_quote_docx_btn",
                )

        display_path_str = st.session_state.get('current_display_path', st.session_state['current_pdf_path'])
        target = Path(display_path_str)
        if not target.exists():
            target = find_proposal_file(target.name)
        if target and target.exists():
            file_bytes = target.read_bytes()
            mime_type = "application/pdf" if target.suffix.lower() == ".pdf" else (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.download_button(
                label="הורדת הקובץ PDF",
                data=file_bytes,
                file_name=target.name,
                mime=mime_type,
                key="download_quote_btn",
            )

            if st.button('📂 פתח את מיקום הקובץ', key='open_folder_button'):
                open_folder_with_selection(str(target.resolve()))

            # הכנת תוכן המייל
            project_name = st.session_state.get('current_project_name', '')
            quote_version = st.session_state.get('current_quote_version', '')
            contact_person = st.session_state.get('current_contact_person', '')
            client_email = st.session_state.get('current_email_client', '')
            email_subject = f"הצעת מחיר: {project_name} - סטודיו 84 (גרסה {quote_version})"
            email_body = (
                f"היי {contact_person},\n"
                f"בהמשך לשיחתנו, מצורפת הצעת מחיר עבור פרויקט {project_name}.\n"
                "אשמח לעמוד לרשותך לכל שאלה.\n\n"
                "בברכה,\n"
                "סטודיו 84"
            )
            cc_list = [EMAIL_ACCOUNTING, EMAIL_ERAN, EMAIL_MYSELF]
            gmail_url = build_gmail_link(client_email, cc_list, email_subject, email_body)
            mailto_url = build_mailto_link(client_email, cc_list, email_subject, email_body)
            col_gmail, col_outlook = st.columns(2)
            with col_gmail:
                st.link_button("📧 פתח טיוטה ב-Gmail", gmail_url)
            with col_outlook:
                st.link_button("✉️ פתח ב-Outlook / תוכנה אחרת", mailto_url)

        if st.button('🔄 התחל הצעה חדשה', key='reset_quote_btn'):
            for k in ['current_pdf_path', 'current_docx_path', 'current_quotes_dir', 'current_display_path',
                      'current_email_client', 'current_contact_person', 'current_project_name', 'current_quote_version', 'current_client_name']:
                st.session_state.pop(k, None)
            st.rerun()

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
                            if st.button("📄 פתח קובץ", key=f"open_signed_{quote_key}"):
                                try:
                                    if Path(signed_path_val).exists():
                                        open_file_in_explorer(signed_path_val)
                                        st.success("חלון Explorer נפתח עם מיקום הקובץ.")
                                    else:
                                        st.warning("הקובץ לא נמצא בנתיב השמור.")
                                except Exception as e:
                                    st.error(f"שגיאה בפתיחת הקובץ: {e}")
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
                label = f"🚀 פתח פרויקט - {client} | {project} | {version}"
                if st.button(label, key=f"open_project_from_quote_{idx}"):
                    st.session_state["open_project_from_quote"] = {
                        "Client": client,
                        "Project": project,
                        "Version": version,
                    }

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
                        df_contacts = load_contacts()
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
                    if st.button("➕ הוסף לפרויקטים פעילים", key=f"kickoff_add_projects_{kickoff_key}"):
                        exists_csv = _project_exists_in_projects_csv(client, project)
                        exists_db = _project_exists_in_projects(client, project)
                        if exists_db:
                            # הפרויקט כבר ב-projects – עדכן סטטוס ל'בעבודה' כדי שיופיע במוניטור וב-Task Board
                            _ensure_project_active_in_projects(client, project, status="בעבודה")
                            st.cache_data.clear()
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
                            except BaseException as e:
                                st.error(f"⚠️ שגיאה ביצירת תיקיית דרופבוקס (הפרויקט יוקם ללא קישור): {e}")
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
                            )
                            st.cache_data.clear()
                            st.session_state["kickoff_success_project"] = f"{client}|{project}"
                            st.session_state["kickoff_success_links"] = (main_link, upload_link, deliverables_link)
                            st.success("✅ הפרויקט, התיקיות ובקשת הקבצים הוקמו בהצלחה!")
                            col1, col2, col3 = st.columns(3)
                            if (main_link or "").startswith("http"):
                                col1.link_button("📂 תיקיית פרויקט ראשית", main_link)
                            if (upload_link or "").startswith("http"):
                                col2.link_button("📥 לינק לבקשת חומרים", upload_link)
                            if (deliverables_link or "").startswith("http"):
                                col3.link_button("📤 תיקיית תוצרים", deliverables_link)
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
                            except BaseException as e:
                                st.error(f"⚠️ שגיאה ביצירת תיקיית דרופבוקס (הפרויקט יוקם ללא קישור): {e}")
                            append_to_projects_csv(client, project, deadline_str, team_str, kickoff_budget, budget_amt, project_contacts=contacts_str)
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
                            )
                            st.cache_data.clear()
                            project_display = f"{client} | {project}"
                            if assigned_team and project_template:
                                append_kickoff_tasks_to_csv(
                                    project_display=project_display,
                                    assigned_team=assigned_team,
                                    project_template=project_template,
                                    task_deadline=task_deadline,
                                )
                            st.session_state["kickoff_success_project"] = f"{client}|{project}"
                            st.session_state["kickoff_success_links"] = (main_link, upload_link, deliverables_link)
                            st.success("✅ הפרויקט, התיקיות ובקשת הקבצים הוקמו בהצלחה!")
                            col1, col2, col3 = st.columns(3)
                            if (main_link or "").startswith("http"):
                                col1.link_button("📂 תיקיית פרויקט ראשית", main_link)
                            if (upload_link or "").startswith("http"):
                                col2.link_button("📥 לינק לבקשת חומרים", upload_link)
                            if (deliverables_link or "").startswith("http"):
                                col3.link_button("📤 תיקיית תוצרים", deliverables_link)

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
                        body += f"\n\n📂 תיקיית פרויקט (דרופבוקס): {main_link_body}"
                    if upload_link_body and upload_link_body.startswith("http"):
                        body += f"\n📥 להעלאת חומרים ע\"י הלקוח: {upload_link_body}"
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

        project_ctx = st.session_state.get("open_project_from_quote")
        if project_ctx:
            st.markdown("### פתיחת פרויקט מתוך הצעה מאושרת")
            client = project_ctx.get("Client", "")
            project = project_ctx.get("Project", "")
            version = project_ctx.get("Version", "")

            with st.form("open_project_form"):
                st.text_input("לקוח", value=client, disabled=True, key="project_form_client")
                st.text_input("פרויקט", value=project, disabled=True, key="project_form_project")
                manager = st.selectbox(
                    "מנהל פרויקט",
                    PROJECT_MANAGERS,
                    key="project_form_manager",
                )
                team = st.multiselect(
                    "צוות",
                    PROJECT_TEAM_MEMBERS,
                    key="project_form_team",
                )
                submit_project = st.form_submit_button("הקם פרויקט ושלח לצוות")

            if submit_project:
                if not manager:
                    st.warning("נא לבחור מנהל פרויקט.")
                else:
                    try:
                        today_str = date.today().strftime("%d/%m/%Y")
                        project_path = ensure_project_folders_for_approved_quote(client, project)
                        # נתיב Dropbox יישמר ויישלח במייל כנתיב יחסי, זהה אצל כל עובד:
                        # Projects/{Client}/{Project}
                        budget_amt = 0.0
                        for q in read_quotes_log():
                            if (
                                (q.get("Client") or "").strip() == client
                                and (q.get("Project") or "").strip() == project
                                and (q.get("Version") or "").strip() == version
                            ):
                                budget_amt = _extract_total_from_quote_row(q)
                                break

                        main_link_ap, upload_link_ap, deliverables_link_ap = "", "", ""
                        try:
                            with st.spinner('מייצר מבנה תיקיות ב-Dropbox ופותח פרויקט...'):
                                result_ap = create_studio_dropbox_structure(project)
                                if result_ap:
                                    main_link_ap, upload_link_ap, deliverables_link_ap = result_ap
                        except BaseException as e:
                            st.error(f"⚠️ שגיאה ביצירת תיקיית דרופבוקס (הפרויקט יוקם ללא קישור): {e}")
                        append_project_record(
                            client=client,
                            project_name=project,
                            manager=manager,
                            team_members=team,
                            status=DEFAULT_PROJECT_STATUS,
                            start_date_str=today_str,
                            budget_amount=budget_amt,
                            dropbox_main=main_link_ap,
                            dropbox_upload=upload_link_ap,
                            dropbox_deliverables=deliverables_link_ap,
                        )
                        st.cache_data.clear()

                        # עדכון סטטוס ההצעה ל-Approved בגיליון quotes
                        try:
                            rows = read_quotes_log()
                            updated = False
                            for r in rows:
                                if (
                                    (r.get("Client") or "").strip() == client
                                    and (r.get("Project") or "").strip() == project
                                    and (r.get("Version") or "").strip() == version
                                ):
                                    if (r.get("Status") or "").strip() != "Signed":
                                        r["Status"] = "Approved"
                                    updated = True
                                    break
                            if updated:
                                write_quotes_log(rows)
                            else:
                                st.warning("לא נמצאה הצעה תואמת לעדכון סטטוס בגיליון quotes.")
                        except Exception as e:
                            st.error(f"שגיאה בעדכון הסטטוס ל-Approved: {e}")

                        manager_email = PROJECT_MANAGER_EMAILS.get(manager, "")
                        team_str = ", ".join(team) if team else "הצוות"
                        email_subject = f"פרויקט חדש לביצוע: {project}"
                        email_body = f"""היי {team_str},
נפתח פרויקט חדש עבור {client}.
מנהל הפרויקט: {manager}.
לינק לתיקייה:

{project_path}
"""
                        if main_link_ap or upload_link_ap:
                            email_body += "\n"
                            if main_link_ap and main_link_ap.startswith("http"):
                                email_body += f"📂 תיקיית פרויקט (דרופבוקס): {main_link_ap}\n"
                            if upload_link_ap and upload_link_ap.startswith("http"):
                                email_body += f"📥 להעלאת חומרים ע\"י הלקוח: {upload_link_ap}\n"
                        email_body += "\nבהצלחה!"

                        gmail_url = build_gmail_link(
                            manager_email,
                            [],
                            email_subject,
                            email_body,
                        )
                        mailto_url = build_mailto_link(
                            manager_email,
                            [],
                            email_subject,
                            email_body,
                        )

                        st.success("✅ הפרויקט, התיקיות ובקשת הקבצים הוקמו בהצלחה!")
                        col1_ap, col2_ap, col3_ap = st.columns(3)
                        if (main_link_ap or "").startswith("http"):
                            col1_ap.link_button("📂 תיקיית פרויקט ראשית", main_link_ap)
                        if (upload_link_ap or "").startswith("http"):
                            col2_ap.link_button("📥 לינק לבקשת חומרים", upload_link_ap)
                        if (deliverables_link_ap or "").startswith("http"):
                            col3_ap.link_button("📤 תיקיית תוצרים", deliverables_link_ap)
                        col_gmail_proj, col_outlook_proj = st.columns(2)
                        with col_gmail_proj:
                            st.link_button("📧 פתח טיוטה ב-Gmail למנהל הפרויקט", gmail_url)
                        with col_outlook_proj:
                            st.link_button("✉️ פתח טיוטה ב-Outlook למנהל הפרויקט", mailto_url)
                    except Exception as e:
                        st.error(f"שגיאה בפתיחת הפרויקט ושמירתו: {e}")

        # פתיחת PDF - חיפוש בכל התיקיות (Pending/Approved/Rejected) + נתיב legacy
        st.divider()
        st.subheader("פתיחת קובץ PDF")
        options = []
        idx_to_row = {}
        for i, r in enumerate(rows):
            label = f"{r.get('Client','')} | {r.get('Project','')} | {r.get('Version','')} | {r.get('Date','')}"
            options.append((i, label))
            idx_to_row[i] = r

        selected_idx = st.selectbox(
            "בחר הצעה לפתיחה",
            options=[i for i, _ in options],
            format_func=lambda i: dict(options).get(i, str(i)),
        )
        open_clicked = st.button("📄 פתח PDF", help="מחפש את ה-PDF בתיקיות Pending/Approved ופותח חלון Explorer.")
        if open_clicked:
            try:
                r = idx_to_row.get(selected_idx, {})
                client = (r.get("Client") or "").strip()
                project = (r.get("Project") or "").strip()
                file_path = r.get("File Path") or ""
                found = None
                if file_path:
                    p = Path(file_path)
                    pdf_name = p.name.replace(".docx", ".pdf") if p.suffix.lower() == ".docx" else p.name
                    extra_dir = p.parent if p.parent.exists() else None
                    found = _find_pdf_in_quotes_folders(pdf_name, extra_dir)
                if not found and client and project:
                    found = _find_pdf_by_client_project(client, project)
                if found:
                    open_file_in_explorer(str(found))
                    st.success("חלון Explorer נפתח עם מיקום הקובץ.")
                else:
                    quotes_path = str(QUOTES_ROOT.resolve())
                    if platform.system() == "Windows":
                        subprocess.run(["explorer", quotes_path])
                    elif platform.system() == "Darwin":
                        subprocess.Popen(["open", quotes_path])
                    else:
                        subprocess.Popen(["xdg-open", quotes_path])
                    st.info("הקובץ לא נמצא. נפתחה תיקיית Quotes לחיפוש ידני.")
            except Exception as e:
                st.error(f"שגיאה בפתיחת הקובץ: {e}")

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

        # --- פתיחת PDF (fallback) ---
        st.divider()
        st.subheader("פתיחת קובץ PDF")
        options_fb = [(i, f"{r.get('Client','')} | {r.get('Project','')} | {r.get('Version','')} | {r.get('Date','')}") for i, r in enumerate(rows)]
        idx_to_label_fb = {i: lbl for i, lbl in options_fb}
        idx_to_row_fb = {i: r for i, r in enumerate(rows)}
        selected_fb = st.selectbox("בחר הצעה לפתיחה", options=[i for i, _ in options_fb], format_func=lambda i: idx_to_label_fb.get(i, str(i)), key="open_pdf_select_fb")
        if st.button("📄 פתח PDF", key="open_pdf_btn_fb"):
            try:
                r = idx_to_row_fb.get(selected_fb, {})
                client = (r.get("Client") or "").strip()
                project = (r.get("Project") or "").strip()
                file_path = r.get("File Path") or ""
                found = None
                if file_path:
                    p = Path(file_path)
                    pdf_name = p.name.replace(".docx", ".pdf") if p.suffix.lower() == ".docx" else p.name
                    extra_dir = p.parent if p.parent.exists() else None
                    found = _find_pdf_in_quotes_folders(pdf_name, extra_dir)
                if not found and client and project:
                    found = _find_pdf_by_client_project(client, project)
                if found:
                    open_file_in_explorer(str(found))
                    st.success("חלון Explorer נפתח עם מיקום הקובץ.")
                else:
                    quotes_path = str(QUOTES_ROOT.resolve())
                    if platform.system() == "Windows":
                        subprocess.run(["explorer", quotes_path])
                    elif platform.system() == "Darwin":
                        subprocess.Popen(["open", quotes_path])
                    else:
                        subprocess.Popen(["xdg-open", quotes_path])
                    st.info("הקובץ לא נמצא. נפתחה תיקיית Quotes לחיפוש ידני.")
            except Exception as e:
                st.error(f"שגיאה בפתיחת הקובץ: {e}")

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
                            if st.button("📄 פתח קובץ", key=f"open_signed_fb_{quote_key}"):
                                try:
                                    if Path(signed_path_val).exists():
                                        open_file_in_explorer(signed_path_val)
                                        st.success("חלון Explorer נפתח עם מיקום הקובץ.")
                                    else:
                                        st.warning("הקובץ לא נמצא בנתיב השמור.")
                                except Exception as e:
                                    st.error(f"שגיאה בפתיחת הקובץ: {e}")
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
                        df_contacts_fb = load_contacts()
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
                    if st.button("➕ הוסף לפרויקטים פעילים", key=f"kickoff_add_projects_{kickoff_key_fb}"):
                        exists_csv_fb = _project_exists_in_projects_csv(client_val, project_val)
                        exists_db_fb = _project_exists_in_projects(client_val, project_val)
                        if exists_db_fb:
                            # הפרויקט כבר ב-projects – עדכן סטטוס ל'בעבודה' כדי שיופיע במוניטור וב-Task Board
                            _ensure_project_active_in_projects(client_val, project_val, status="בעבודה")
                            st.cache_data.clear()
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
                            except BaseException as e:
                                st.error(f"⚠️ שגיאה ביצירת תיקיית דרופבוקס (הפרויקט יוקם ללא קישור): {e}")
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
                            )
                            st.cache_data.clear()
                            st.session_state["kickoff_success_project"] = f"{client_val}|{project_val}"
                            st.session_state["kickoff_success_links"] = (main_link_fb, upload_link_fb, deliverables_link_fb)
                            st.success("✅ הפרויקט, התיקיות ובקשת הקבצים הוקמו בהצלחה!")
                            col1_fb, col2_fb, col3_fb = st.columns(3)
                            if (main_link_fb or "").startswith("http"):
                                col1_fb.link_button("📂 תיקיית פרויקט ראשית", main_link_fb)
                            if (upload_link_fb or "").startswith("http"):
                                col2_fb.link_button("📥 לינק לבקשת חומרים", upload_link_fb)
                            if (deliverables_link_fb or "").startswith("http"):
                                col3_fb.link_button("📤 תיקיית תוצרים", deliverables_link_fb)
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
                            except BaseException as e:
                                st.error(f"⚠️ שגיאה ביצירת תיקיית דרופבוקס (הפרויקט יוקם ללא קישור): {e}")
                            append_to_projects_csv(client_val, project_val, deadline_str_fb, team_str_fb, kickoff_budget_fb, budget_amt_fb, project_contacts=contacts_str_fb)
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
                            )
                            st.cache_data.clear()
                            project_display_fb = f"{client_val} | {project_val}"
                            if assigned_team_fb and project_template_fb:
                                append_kickoff_tasks_to_csv(
                                    project_display=project_display_fb,
                                    assigned_team=assigned_team_fb,
                                    project_template=project_template_fb,
                                    task_deadline=task_deadline_fb,
                                )
                            st.session_state["kickoff_success_project"] = f"{client_val}|{project_val}"
                            st.session_state["kickoff_success_links"] = (main_link_fb, upload_link_fb, deliverables_link_fb)
                            st.success("✅ הפרויקט, התיקיות ובקשת הקבצים הוקמו בהצלחה!")
                            col1_fb, col2_fb, col3_fb = st.columns(3)
                            if (main_link_fb or "").startswith("http"):
                                col1_fb.link_button("📂 תיקיית פרויקט ראשית", main_link_fb)
                            if (upload_link_fb or "").startswith("http"):
                                col2_fb.link_button("📥 לינק לבקשת חומרים", upload_link_fb)
                            if (deliverables_link_fb or "").startswith("http"):
                                col3_fb.link_button("📤 תיקיית תוצרים", deliverables_link_fb)

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
                        body += f"\n\n📂 תיקיית פרויקט (דרופבוקס): {main_link_body_fb}"
                    if upload_link_body_fb and upload_link_body_fb.startswith("http"):
                        body += f"\n📥 להעלאת חומרים ע\"י הלקוח: {upload_link_body_fb}"
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


def show_tasks_page() -> None:
    st.title("ניהול פרויקטים ומשימות")

    sub_nav = st.radio(
        "בחר תצוגה:",
        ["טבלת פרויקטים", "רשימת משימות (Task Board)", "לוח עומסים מנהלים (גאנט)"],
        horizontal=True,
    )

    if sub_nav == "טבלת פרויקטים":
        st.subheader("טבלת פרויקטים")
        projects_rows = read_projects()
        if not projects_rows:
            st.info("אין פרויקטים. פתח פרויקט מתוך הצעה מאושרת בלשונית ניהול הצעות.")
        else:
            df_projects = pd.DataFrame(projects_rows, columns=PROJECTS_DB_COLUMNS)
            df_projects = df_projects.fillna('')
            if df_projects.empty:
                st.info("אין פרויקטים להצגה.")
            else:
                try:
                    edited_projects = st.data_editor(
                        df_projects,
                        hide_index=True,
                        use_container_width=True,
                        disabled=[c for c in PROJECTS_DB_COLUMNS if c not in ("Status", "Manager", "Team")],
                        column_config={
                            "Status": st.column_config.SelectboxColumn(
                                "סטטוס",
                                options=ALLOWED_PROJECT_STATUSES,
                                required=True,
                            ),
                            "Manager": st.column_config.SelectboxColumn(
                                "מנהל",
                                options=PROJECT_MANAGERS,
                            ),
                            "Team": st.column_config.SelectboxColumn(
                                "צוות",
                                options=PROJECT_TEAM_MEMBERS,
                            ),
                            'היקף כספי (₪)': st.column_config.NumberColumn(
                                "היקף כספי (₪)",
                                format="₪%d",
                                default=0,
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
                        key="projects_editor",
                    )
                except Exception as e:
                    st.error(f"שגיאה בהצגת נתונים: {e}")
                    edited_projects = df_projects
                if st.button("שמור שינויים בפרויקטים", type="primary", key="save_projects_btn"):
                    if is_admin():
                        updated = edited_projects.to_dict(orient="records")
                    else:
                        # משתמש לא-מנהל: מיזוג השינויים חזרה לרשימה המלאה
                        full_rows = read_projects()
                        edited_records = edited_projects.to_dict(orient="records")
                        edited_by_key = {(str(r.get("Client", "")), str(r.get("Project Name", ""))): r for r in edited_records}
                        for i, row in enumerate(full_rows):
                            key = (str(row.get("Client", "")), str(row.get("Project Name", "")))
                            if key in edited_by_key:
                                full_rows[i] = edited_by_key[key]
                        updated = full_rows
                    write_projects(updated)
                    st.success("השינויים נשמרו בהצלחה!")
                    st.rerun()

    elif sub_nav == "רשימת משימות (Task Board)":
        # --- מוניטור סטודיו - תמונת מצב צוותית (לפני טופס הוספת משימה) ---
        st.subheader("🎯 מוניטור סטודיו - תמונת מצב צוותית")

        def _assignee_matches_task(task_assignee: str, sel: str) -> bool:
            a = (task_assignee or "").strip()
            return a == sel or a.startswith(sel + " ") or a.startswith(sel + "-")

        tasks_rows_monitor = read_tasks()
        today_dt = date.today()
        # סינון חכם: סטטוס לא בוצע, תאריך <= היום. מיון: באיחור ראשון, אחר כך היום.
        open_tasks = _filter_tasks_by_status_and_date(
            tasks_rows_monitor, date_col="Due Date", status_col="Status"
        )
        # משתמש שאינו מנהל - זמנית מציגים הכל (ביטול סינון לפי Assignee)
        # if not is_admin():
        #     current_user = _get_assignee_for_current_user()
        #     open_tasks = [t for t in open_tasks if _assignee_matches_task(t.get("Assignee") or "", current_user)]
        df_tasks = pd.DataFrame(open_tasks) if open_tasks else pd.DataFrame()
        due_col = df_tasks.get("Due Date") if not df_tasks.empty else None
        if due_col is not None:
            df_tasks["due_date_parsed"] = pd.to_datetime(due_col, errors="coerce").dt.date

        def _get_task_due_date(task: dict) -> date | None:
            due_str = task.get("Due Date") or ""
            if not due_str or not str(due_str).strip():
                return None
            dt = pd.to_datetime(due_str, errors="coerce")
            if pd.isna(dt):
                return None
            return dt.date() if hasattr(dt, "date") else dt

        def _mark_task_done_monitor(task_id: str) -> None:
            for t in tasks_rows_monitor:
                if str(t.get("Task ID") or "").strip() == str(task_id or "").strip():
                    t["Status"] = "Done"
                    break
            write_tasks(tasks_rows_monitor)
            st.rerun()

        team_list = [n for n in TEAM_DISPLAY_NAMES if n and str(n).strip()]
        for assignee in team_list:
            user_tasks = [
                t for t in open_tasks
                if _assignee_matches_task(t.get("Assignee") or "", assignee)
            ]
            try:
                with st.expander(f"👤 {assignee} | משימות פתוחות: {len(user_tasks)}", expanded=False):
                    if not user_tasks:
                        st.success("השולחן נקי!")
                    else:
                        for idx, r in enumerate(user_tasks):
                            task_name = (r.get("Task Name") or "").strip()
                            project = (r.get("Project") or "").strip()
                            due_str = (r.get("Due Date") or "").strip()
                            due_parsed = _get_task_due_date(r)
                            col_btn, col_msg = st.columns([0.12, 0.88])
                            with col_btn:
                                if st.button("✅ סמן כבוצע", key=f"monitor_done_{r.get('Task ID')}_{assignee}_{idx}"):
                                    _mark_task_done_monitor(r.get("Task ID"))
                            with col_msg:
                                if due_parsed is None or due_parsed < today_dt:
                                    st.error(f"🔴 **{task_name}** – {project} – {due_str}")
                                elif due_parsed == today_dt:
                                    st.info(f"🔵 **{task_name}** – {project} – {due_str}")
                                else:
                                    st.success(f"🟢 **{task_name}** – {project} – {due_str}")
            except Exception as e:
                st.error(f"שגיאה בהצגת נתונים: {e}")

        st.divider()
        st.subheader("הוספת משימה")
        projects_options = _get_active_projects_options()
        if not projects_options:
            st.info("אין פרויקטים פעילים. הוסף פרויקטים בגיליון projects עם סטטוס 'בעבודה' או 'ממתין להתחלה'.")
        else:
            col1, col2, col3 = st.columns(3)
            with col1:
                selected_project = st.selectbox("פרויקט", options=projects_options, key="task_project")
                assignee = st.selectbox("אחראי", options=TEAM_DISPLAY_NAMES, key="task_assignee")
            with col2:
                task_name = st.text_input("שם המשימה", key="task_name")
                start_date = st.date_input("תאריך התחלה", value=date.today(), key="task_start")
                due_date = st.date_input("תאריך סיום", value=date.today(), key="task_due")
                is_flexible = st.checkbox("דדליין גמיש (ניתן לדחייה בעת עומס)", key="task_flexible")
            with col3:
                priority = st.selectbox("עדיפות", options=TASK_PRIORITIES, key="task_priority")
                add_btn = st.button("הוסף משימה", type="primary", key="add_task_btn")

            if add_btn:
                if not task_name.strip():
                    st.warning("נא להזין שם משימה.")
                elif not selected_project:
                    st.warning("נא לבחור פרויקט.")
                else:
                    existing = read_tasks()
                    task_id = next_task_id(existing)
                    row = {
                        "Task ID": str(task_id),
                        "Project": selected_project,
                        "Assignee": assignee,
                        "Task Name": task_name.strip(),
                        "Start Date": start_date.strftime("%d/%m/%Y"),
                        "Due Date": due_date.strftime("%d/%m/%Y"),
                        "Status": "To Do",
                        "Priority": priority,
                        "Notes": "",
                        "Flexible": "1" if is_flexible else "0",
                    }
                    existing.append(row)
                    write_tasks(existing)
                    st.success("המשימה נוספה בהצלחה!")
                    st.rerun()

        st.divider()
        with st.expander("🗑️ ניקוי ומחיקת פרויקטים מהמערכת"):
            delete_options = _get_active_projects_options()
            if not delete_options:
                st.info("אין פרויקטים פעילים למחיקה. הרשימה זהה לרשימת 'הוספת משימה'.")
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
                                if (t.get("Project") or "").strip() not in (project_key_pipe, project_key_dash)
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
            st.info("אין משימות. הוסף משימה חדשה למעלה.")
        else:
            # סינון חכם: סטטוס לא בוצע, תאריך <= היום. מיון: באיחור ראשון, אחר כך היום.
            filtered_rows = _filter_tasks_by_status_and_date(
                tasks_rows, date_col="Due Date", status_col="Status"
            )
            df = pd.DataFrame(filtered_rows, columns=TASKS_LOG_COLUMNS)
            df = df.reindex(columns=TASKS_LOG_COLUMNS, fill_value='')
            df = df.fillna('')
            if hasattr(df.columns, 'str'):
                df.columns = df.columns.str.strip()
            # הוספת עמודת אינדיקציה לאיחור (🔴 למשימות שנגררו)
            today_dt = date.today()
            def _overdue_indicator(row):
                d = _parse_date_safe(row.get("Due Date"), "Due Date")
                return "🔴" if d is None or d < today_dt else ""
            df.insert(0, "איחור", df.apply(_overdue_indicator, axis=1))
            # זמנית: ביטול סינון לפי Assignee - הצגת כל המשימות
            # if not is_admin():
            #     current_user = _get_assignee_for_current_user()
            #     assignee_col = df.get("Assignee")
            #     if assignee_col is not None:
            #         df = df[assignee_col.fillna("").apply(lambda a: _assignee_matches_task(str(a), current_user))]
            if df.empty:
                st.info("אין משימות להצגה.")
            else:
                try:
                    editable_cols = ["Status", "Priority", "Notes"]
                    disabled_cols = ["איחור"] + [c for c in TASKS_LOG_COLUMNS if c not in editable_cols]
                    edited_df = st.data_editor(
                        df,
                        hide_index=True,
                        use_container_width=True,
                        disabled=disabled_cols,
                        column_config={
                            "Status": st.column_config.SelectboxColumn(
                                "Status",
                                options=TASK_STATUSES,
                                required=True,
                            ),
                            "Priority": st.column_config.SelectboxColumn(
                                "עדיפות",
                                options=TASK_PRIORITIES,
                                required=True,
                            ),
                        },
                        key="tasks_editor",
                    )
                except Exception as e:
                    st.error(f"שגיאה בהצגת נתונים: {e}")
                    edited_df = df
                if st.button("שמור שינויים", type="primary", key="save_tasks_btn"):
                    # הסרת עמודת האינדיקציה לפני שמירה
                    save_df = edited_df.drop(columns=["איחור"], errors="ignore")
                    if is_admin():
                        updated = save_df.to_dict(orient="records")
                    else:
                        # משתמש לא-מנהל: מיזוג השינויים חזרה לרשימת המשימות המלאה
                        full_rows = read_tasks()
                        edited_records = save_df.to_dict(orient="records")
                        edited_by_id = {str(r.get("Task ID", "")): r for r in edited_records}
                        for i, row in enumerate(full_rows):
                            tid = str(row.get("Task ID", ""))
                            if tid in edited_by_id:
                                full_rows[i] = edited_by_id[tid]
                        updated = full_rows
                    write_tasks(updated)
                    st.success("השינויים נשמרו בהצלחה!")

        # --- אנשי קשר לפרויקט ---
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

    elif sub_nav == "לוח עומסים מנהלים (גאנט)":
        st.subheader("לוח עומסים צוותי - לוח שנה אינטראקטיבי")
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
                    start_val = row.get("Start Date", row.get("Due Date"))
                    if pd.isna(start_val) or str(start_val).strip() == "":
                        start_val = row.get("Due Date")
                    if pd.isna(start_val):
                        continue

                    start_str = pd.to_datetime(start_val).strftime("%Y-%m-%d")
                    end_val = row.get("Due Date")
                    end_str = pd.to_datetime(end_val).strftime("%Y-%m-%d") if not pd.isna(end_val) else start_str
                    assignee = str(row.get("Assignee", "") or "").split("-")[0].strip()
                    task_name = row.get("Task Name", "") or ""

                    calendar_events.append({
                        "title": f"{assignee}: {task_name}",
                        "start": start_str,
                        "end": end_str,
                        "color": "#4b8bbe",
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

    # טופס הוספה מהירה
    with st.expander("➕ הוספת איש קשר חדש", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            form_name = st.text_input("שם מלא", key="contact_form_name")
            form_company = st.text_input("חברה / משרד אדריכלים", key="contact_form_company")
        with col2:
            form_role = st.text_input("תפקיד", key="contact_form_role")
            form_phone = st.text_input("טלפון", key="contact_form_phone")
        with col3:
            form_email = st.text_input("אימייל", key="contact_form_email")
            form_type = st.selectbox(
                "סוג איש קשר",
                options=CONTACT_TYPE_OPTIONS,
                key="contact_form_type",
            )
        if st.button("הוסף איש קשר", type="primary", key="add_contact_btn"):
            if not (form_name or "").strip():
                st.warning("נא להזין לפחות שם מלא.")
            else:
                try:
                    df = load_contacts()
                    new_row = {
                        "שם מלא": (form_name or "").strip(),
                        "חברה / משרד אדריכלים": (form_company or "").strip(),
                        "תפקיד": (form_role or "").strip(),
                        "טלפון": (form_phone or "").strip(),
                        "אימייל": (form_email or "").strip(),
                        "סוג איש קשר": form_type or CONTACT_TYPE_OPTIONS[0],
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    save_contacts(df)
                    st.success("איש הקשר נוסף בהצלחה!")
                    st.rerun()
                except Exception as e:
                    st.error(f"שגיאה בהוספה: {e}")

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
                            df_existing = load_contacts()
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

    st.subheader("מאגר אנשי קשר")
    df = load_contacts()
    if df.empty:
        st.info("אין עדיין אנשי קשר. הוסף איש קשר באמצעות הטופס למעלה.")
        return

    column_config = {
        "שם מלא": st.column_config.TextColumn("שם מלא"),
        "חברה / משרד אדריכלים": st.column_config.TextColumn("חברה / משרד אדריכלים"),
        "תפקיד": st.column_config.TextColumn("תפקיד"),
        "טלפון": st.column_config.TextColumn("טלפון"),
        "אימייל": st.column_config.TextColumn("אימייל"),
        "סוג איש קשר": st.column_config.SelectboxColumn(
            "סוג איש קשר",
            options=CONTACT_TYPE_OPTIONS,
            required=True,
        ),
    }

    edited_df = st.data_editor(
        df,
        hide_index=True,
        use_container_width=True,
        column_config=column_config,
        key="contacts_data_editor",
    )

    if not edited_df.equals(df):
        try:
            save_contacts(edited_df)
            st.success("השינויים נשמרו בהצלחה!")
            st.rerun()
        except Exception as e:
            st.error(f"שגיאה בשמירה: {e}")


def show_login_screen() -> bool:
    """
    מסך כניסה - Admin דורש סיסמה, Team Member לא.
    שומר את שם המשתמש הנבחר ב-st.session_state.current_user.
    מחזיר True אם המשתמש התחבר בהצלחה.
    """
    if st.session_state.get("authenticated"):
        return True

    st.title("🔐 כניסה למערכת ניהול סטודיו")
    st.markdown("---")

    selected_user = st.selectbox(
        "בחר משתמש",
        TEAM_MEMBERS_LOGIN,
        key="login_user_select",
    )

    if selected_user == "Admin (טלי / ערן)":
        password = st.text_input("סיסמה", type="password", key="login_password", placeholder="הזן סיסמה")
        if st.button("התחבר", type="primary", key="login_submit"):
            if password == ADMIN_PASSWORD:
                st.session_state["authenticated"] = True
                st.session_state["user_type"] = "admin"
                st.session_state["current_user"] = selected_user
                st.success("התחברת בהצלחה!")
                st.rerun()
            else:
                st.error("סיסמה שגויה. נסה שוב.")
    else:
        if st.button("התחבר", type="primary", key="login_submit_team"):
            st.session_state["authenticated"] = True
            st.session_state["user_type"] = "team_member"
            st.session_state["current_user"] = selected_user
            st.success("שלום צוות!")
            st.rerun()

    return False


def _get_assignee_for_current_user() -> str:
    """מחזיר את ה-Assignee המתאים למשתמש הנוכחי (לסינון משימות)."""
    return (st.session_state.get("current_user") or "צוות").strip()


def is_admin() -> bool:
    """מזהה אם המשתמש המחובר הוא מנהל (Admin או שמות המנהלים)."""
    user_type = st.session_state.get("user_type")
    if user_type == "admin":
        return True
    current_user = (st.session_state.get("current_user") or "").strip()
    return current_user in ADMIN_NAMES


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
    # בדיקת התחברות
    if not st.session_state.get("authenticated"):
        show_login_screen()
        return

    # אתחול משתני מצב למוניטור (Drill-down)
    if "monitor_filter" not in st.session_state:
        st.session_state.monitor_filter = "__ALL__"  # זמנית: הצגת כל הפרויקטים כברירת מחדל
    if "monitor_title" not in st.session_state:
        st.session_state.monitor_title = "כל הפרויקטים (ללא סינון)"

    user_type = st.session_state.get("user_type", "team_member")

    st.sidebar.title("תפריט ניהול")
    if st.sidebar.button("🔄 רענן נתונים", key="refresh_data_btn", use_container_width=True):
        st.rerun()
    if st.sidebar.button("🚪 התנתק", key="logout_btn"):
        st.session_state["authenticated"] = False
        st.session_state["user_type"] = None
        st.session_state["current_user"] = None
        st.rerun()

    if user_type == "team_member":
        # עובד - רק לשונית 'העבודה שלי'
        if st.sidebar.button("🔄 רענן נתונים", key="refresh_data_btn", use_container_width=True):
            st.rerun()
        _render_quick_comm_sidebar_form()
        _render_dropbox_refresh_token_sidebar()
        show_my_work_page()
        return

    # מנהל - ניווט ראשי: חדר מצב vs ניהול שוטף
    main_nav = st.sidebar.radio("ניווט ראשי:", ["📊 חדר מצב (מוניטור פרויקטים)", "⚙️ ניהול שוטף (הצעות, משימות, לקוחות)"])

    # טעינת נתוני פרויקטים לדיבאג (מצב 'רנטגן')
    projects_rows_debug = read_projects()
    df_projects = pd.DataFrame(projects_rows_debug, columns=PROJECTS_DB_COLUMNS)
    df_projects = df_projects.fillna('')

    if main_nav == "📊 חדר מצב (מוניטור פרויקטים)":
        _render_quick_comm_notifications()
        # אזור חדר המצב: מוניטור פרויקטים - כפתורים צידיים + תצוגת טבלאות
        stats = _compute_project_monitor_stats()
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📊 מוניטור פרויקטים")
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
        _render_dropbox_refresh_token_sidebar()

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
                    try:
                        edited_filtered_df = st.data_editor(
                            filtered_df,
                            hide_index=True,
                            use_container_width=True,
                            key="drilldown_editor",
                            column_config={
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
                        edited_filtered_df = filtered_df
                    if not edited_filtered_df.equals(filtered_df):
                        df_projects.update(edited_filtered_df)
                        updated = df_projects.to_dict(orient="records")
                        write_projects(updated)
                        st.success("הנתונים עודכנו בהצלחה!")
                        st.rerun()
                if st.button("✖️ סגור תצוגה ממוקדת", key="close_monitor_drill"):
                    st.session_state.monitor_filter = None
                    st.session_state.monitor_title = ""
                    st.rerun()

    elif main_nav == "⚙️ ניהול שוטף (הצעות, משימות, לקוחות)":
        _render_quick_comm_notifications()
        # אזור הניהול השוטף: תפריט 'בחר פעולה' + כל המסכים המשויכים
        page = st.sidebar.radio(
            "בחר פעולה",
            (
                "יצירת הצעה חדשה",
                "ניהול הצעות",
                "ניהול פרויקטים ומשימות",
                "👥 לקוחות ואנשי קשר",
            ),
            index=0,
        )
        _render_quick_comm_sidebar_form()
        if st.sidebar.button('הצג נתונים גולמיים'):
            st.write(df_projects)
        _render_dropbox_refresh_token_sidebar()

        if page == "יצירת הצעה חדשה":
            show_quote_page()
        elif page == "ניהול הצעות":
            show_quotes_management_page()
        elif page == "👥 לקוחות ואנשי קשר":
            show_contacts_page()
        else:
            show_tasks_page()


if __name__ == "__main__":
    main()

