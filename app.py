import base64
import io
import json
import os
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Electronic Component Procurement System — Application Shell
# ----------------------------------------------------------------------------

st.set_page_config(
    page_title="Electronic Component Procurement System",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# Color palette (matches the Figma theme tokens)
# ----------------------------------------------------------------------------
PRIMARY_BLUE = "#5B7DB1"      # --primary
LIGHT_BLUE = "#EAF1FA"        # --secondary / sidebar bg
SIDEBAR_ACCENT = "#D9EAF7"    # --sidebar-accent (active pill)
MUTED_GOLD = "#D9C38A"        # --accent
MINT = "#E8F3EE"
BACKGROUND = "#F4F7FB"        # --background
CARD_BG = "#FFFFFF"           # --card
TEXT = "#1E2A3A"             # --foreground
MUTED_TEXT = "#5A6A80"        # --muted-foreground
BORDER = "rgba(91, 125, 177, 0.15)"

INPUT_BG = "#EAF1FA"          # --input-background
GOLD_BG = "#F7EFD7"
GOLD_TEXT = "#9A7B2E"

# Bump this whenever code changes, so the deployed build is identifiable at a
# glance (shown in the sidebar). If the cloud shows an older value than this,
# it has NOT redeployed the latest commit yet.
APP_VERSION = "build 2026-06-25 #16 (procurement-outputs)"

# ----------------------------------------------------------------------------
# Data layer — two tables: Wishlist + SupplierOptions
#
# Storage backend is chosen automatically:
#   * Google Sheets  — when a [gcp_service_account] secret is configured
#                      (this is the shared, persistent store used in the cloud).
#   * Local Excel    — otherwise (handy for local development).
#
# NOTE: This is a workflow-management / record-keeping tool. It does NOT scrape
# supplier websites and does NOT auto-compare prices. Users enter all supplier
# details manually after checking vendor sites themselves.
# ----------------------------------------------------------------------------
EXCEL_FILE = Path(__file__).parent / "wishlist.xlsx"

WL_COLUMNS = ["#", "Component", "Model", "Specifications", "Quantity",
              "Date Added", "Status", "Selected Supplier"]
OPT_COLUMNS = ["Component ID", "Supplier", "Supplier Type", "Price", "Stock",
               "ETA", "Shopping Cart Available", "URL", "Selected"]

TRUSTED_SUPPLIERS = ["RS Components", "Communica", "Mantech", "Mintech", "Other"]


# ---- Backend selection -----------------------------------------------------
def _secrets_file_exists() -> bool:
    """True if a secrets.toml exists in either location Streamlit reads.

    Checked first so we never touch st.secrets when there's no file — otherwise
    Streamlit shows a 'No secrets files found' error in the app. On Streamlit
    Cloud, secrets configured in the dashboard ARE written to one of these.
    """
    for p in (Path.home() / ".streamlit" / "secrets.toml",
              Path(__file__).parent / ".streamlit" / "secrets.toml"):
        if p.exists():
            return True
    return False


def _use_gsheets() -> bool:
    if not _secrets_file_exists():
        return False
    try:
        return ("gcp_service_account_json" in st.secrets
                or "gcp_service_account" in st.secrets)
    except Exception:
        return False


def _service_account_info() -> dict:
    """Return the service-account credentials from secrets.

    Two accepted formats:
      * gcp_service_account_json = '''<paste the whole .json file here>'''  (easy)
      * [gcp_service_account] TOML table with each field  (advanced)
    """
    if "gcp_service_account_json" in st.secrets:
        return json.loads(st.secrets["gcp_service_account_json"])
    return dict(st.secrets["gcp_service_account"])


@st.cache_resource(show_spinner=False)
def _get_worksheets():
    """Open (and lazily create) the two worksheets in the shared spreadsheet."""
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        _service_account_info(), scopes=scopes
    )
    gc = gspread.authorize(creds)

    # Accept either a bare Sheet ID or a full URL in either secret field.
    raw = str(st.secrets.get("spreadsheet_key", "")
              or st.secrets.get("spreadsheet_url", "")).strip()
    m = re.search(r"/d/([a-zA-Z0-9\-_]+)", raw)
    key = m.group(1) if m else raw
    sh = gc.open_by_key(key)

    def ws_or_create(title, headers):
        try:
            ws = sh.worksheet(title)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title, rows=200, cols=len(headers))
        if ws.row_values(1) != headers:
            if not ws.row_values(1):
                ws.append_row(headers)
        return ws

    return ws_or_create("Wishlist", WL_COLUMNS), ws_or_create("SupplierOptions", OPT_COLUMNS)


@st.cache_data(ttl=5, show_spinner=False)
def _gs_records(title: str):
    wl_ws, opt_ws = _get_worksheets()
    ws = wl_ws if title == "Wishlist" else opt_ws
    return ws.get_all_records()


def _refresh():
    """Drop cached reads so the next load hits the live sheet."""
    try:
        _gs_records.clear()
    except Exception:
        pass


def _as_text(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce every column to object dtype with NaN -> "".

    Newer pandas (on Streamlit Cloud) refuses to assign a string into a column
    it inferred as numeric/empty. Keeping everything as text avoids that and
    also gives clean display (no NaN).
    """
    df = df.astype(object)
    return df.where(pd.notna(df), "")


def _frame(records, columns) -> pd.DataFrame:
    df = pd.DataFrame(records)
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    if df.empty:
        return pd.DataFrame(columns=columns)
    return _as_text(df[columns])


# ---- Loads -----------------------------------------------------------------
def load_wishlist() -> pd.DataFrame:
    if _use_gsheets():
        return _frame(_gs_records("Wishlist"), WL_COLUMNS)
    if EXCEL_FILE.exists():
        xls = pd.ExcelFile(EXCEL_FILE)
        sheet = "Wishlist" if "Wishlist" in xls.sheet_names else xls.sheet_names[0]
        df = xls.parse(sheet)
        if "Status" not in df.columns:
            df["Status"] = "Pending"
        for col in WL_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return _as_text(df[WL_COLUMNS])
    return pd.DataFrame(columns=WL_COLUMNS)


def load_options() -> pd.DataFrame:
    if _use_gsheets():
        return _frame(_gs_records("SupplierOptions"), OPT_COLUMNS)
    if EXCEL_FILE.exists():
        xls = pd.ExcelFile(EXCEL_FILE)
        if "SupplierOptions" in xls.sheet_names:
            df = xls.parse("SupplierOptions")
            for col in OPT_COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            return _as_text(df[OPT_COLUMNS])
    return pd.DataFrame(columns=OPT_COLUMNS)


def write_workbook(wishlist_df: pd.DataFrame, options_df: pd.DataFrame) -> None:
    """Write both sheets together so neither clobbers the other (Excel only)."""
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        wishlist_df.to_excel(writer, sheet_name="Wishlist", index=False)
        options_df.to_excel(writer, sheet_name="SupplierOptions", index=False)


# ---- Writes ----------------------------------------------------------------
def _next_id(records_or_df) -> int:
    """Next component id = max(existing) + 1, so ids never collide after deletes."""
    nums = []
    if isinstance(records_or_df, pd.DataFrame):
        values = records_or_df["#"] if "#" in records_or_df.columns else []
    else:
        values = [r.get("#") for r in records_or_df]
    for v in values:
        try:
            nums.append(int(float(v)))
        except (ValueError, TypeError):
            pass
    return (max(nums) + 1) if nums else 1


def add_component(component, model, specs, qty) -> None:
    if _use_gsheets():
        wl_ws, _ = _get_worksheets()
        records = wl_ws.get_all_records()
        next_num = _next_id(records)
        wl_ws.append_row([
            next_num, component, model, specs, int(qty),
            datetime.now().strftime("%Y-%m-%d"), "Pending", "",
        ], value_input_option="USER_ENTERED")
        _refresh()
        return
    wl = load_wishlist()
    opts = load_options()
    new_row = {
        "#": _next_id(wl),
        "Component": component,
        "Model": model,
        "Specifications": specs,
        "Quantity": int(qty),
        "Date Added": datetime.now().strftime("%Y-%m-%d"),
        "Status": "Pending",
        "Selected Supplier": "",
    }
    wl = pd.concat([wl, pd.DataFrame([new_row])], ignore_index=True)
    write_workbook(wl, opts)


def add_option(component_id, supplier, supplier_type, price, stock, eta, cart, url):
    """Record a manually-entered supplier option for a component."""
    if _use_gsheets():
        _, opt_ws = _get_worksheets()
        opt_ws.append_row([
            component_id, supplier, supplier_type, price, stock, eta,
            cart, url, "FALSE",
        ], value_input_option="USER_ENTERED")
        _refresh()
        return
    wl = load_wishlist()
    opts = load_options()
    new_row = {
        "Component ID": component_id,
        "Supplier": supplier,
        "Supplier Type": supplier_type,
        "Price": price,
        "Stock": stock,
        "ETA": eta,
        "Shopping Cart Available": cart,   # "Yes" / "No"
        "URL": url,
        "Selected": False,
    }
    opts = pd.concat([opts, pd.DataFrame([new_row])], ignore_index=True)
    write_workbook(wl, opts)


def select_option(component_id, row_index) -> None:
    """Mark one option as the chosen supplier (only one per component).

    `row_index` is the 0-based position within the full options list, which
    matches the spreadsheet row order (sheet row = row_index + 2, after header).
    """
    if _use_gsheets():
        _, opt_ws = _get_worksheets()
        records = opt_ws.get_all_records()
        sel_col = OPT_COLUMNS.index("Selected") + 1
        for i, rec in enumerate(records):
            if str(rec.get("Component ID")) == str(component_id):
                opt_ws.update_cell(i + 2, sel_col,
                                   "TRUE" if i == row_index else "FALSE")
        _refresh()
        return
    wl = load_wishlist()
    opts = load_options()
    mask = opts["Component ID"].astype(str) == str(component_id)
    opts.loc[mask, "Selected"] = False
    opts.loc[row_index, "Selected"] = True
    write_workbook(wl, opts)


def confirm_source(component_id) -> str:
    """Lock the sourcing decision: Pending -> Sourced, store chosen supplier."""
    if _use_gsheets():
        wl_ws, opt_ws = _get_worksheets()
        chosen = [r for r in opt_ws.get_all_records()
                  if str(r.get("Component ID")) == str(component_id)
                  and is_true(r.get("Selected"))]
        if not chosen:
            return ""
        supplier = str(chosen[0]["Supplier"])
        status_col = WL_COLUMNS.index("Status") + 1
        supp_col = WL_COLUMNS.index("Selected Supplier") + 1
        for i, rec in enumerate(wl_ws.get_all_records()):
            if str(rec.get("#")) == str(component_id):
                wl_ws.update_cell(i + 2, status_col, "Sourced")
                wl_ws.update_cell(i + 2, supp_col, supplier)
                break
        _refresh()
        return supplier
    wl = load_wishlist()
    opts = load_options()
    chosen = opts[(opts["Component ID"].astype(str) == str(component_id))
                  & (opts["Selected"].apply(is_true))]
    if chosen.empty:
        return ""
    supplier = str(chosen.iloc[0]["Supplier"])
    wmask = wl["#"].astype(str) == str(component_id)
    wl.loc[wmask, "Status"] = "Sourced"
    wl.loc[wmask, "Selected Supplier"] = supplier
    write_workbook(wl, opts)
    return supplier


def delete_component(component_id) -> None:
    """Remove a component AND every supplier option recorded for it."""
    if _use_gsheets():
        wl_ws, opt_ws = _get_worksheets()
        # Delete supplier-option rows (bottom-up so row numbers stay valid)
        opt_records = opt_ws.get_all_records()
        for i in range(len(opt_records) - 1, -1, -1):
            if str(opt_records[i].get("Component ID")) == str(component_id):
                opt_ws.delete_rows(i + 2)
        # Delete the wishlist row
        wl_records = wl_ws.get_all_records()
        for i in range(len(wl_records) - 1, -1, -1):
            if str(wl_records[i].get("#")) == str(component_id):
                wl_ws.delete_rows(i + 2)
        _refresh()
        return
    wl = load_wishlist()
    opts = load_options()
    wl = wl[wl["#"].astype(str) != str(component_id)].reset_index(drop=True)
    opts = opts[opts["Component ID"].astype(str) != str(component_id)].reset_index(drop=True)
    write_workbook(wl, opts)


def route_for(supplier, cart) -> str:
    """Decide which procurement bucket a sourced component belongs to.

    Routing is rule-based on user-entered data only — there is no automatic
    price comparison or 'cheapest supplier' logic (per the lecturer's brief).
    """
    supplier = clean(supplier)
    cart = clean(cart)
    if not supplier:
        return "Unresolved Components"
    if cart.lower() == "yes":
        return "Shopping Cart Queue"
    if supplier == "RS Components":
        return "Veronica Procurement List"
    if supplier in ("Communica", "Mintech", "Mantech"):
        return "Tan Procurement List"
    return "Unresolved Components"


def is_true(v) -> bool:
    """Robust truthy check for values that may be bool or string from Excel."""
    return str(v).strip().lower() in ("true", "1", "yes")


def clean(value) -> str:
    """Display helper: turn NaN / None into an empty string."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    s = str(value)
    return "" if s.lower() == "nan" else s


def _pill(text, bg, fg) -> str:
    return (
        f'<span style="background:{bg};color:{fg};padding:4px 14px;'
        f'border-radius:999px;font-size:13px;font-weight:600;'
        f'white-space:nowrap;">{text}</span>'
    )


def status_badge(status: str) -> str:
    s = (clean(status) or "Pending").strip()
    styles = {
        "Pending": (GOLD_BG, GOLD_TEXT),
        "Sourcing": (LIGHT_BLUE, PRIMARY_BLUE),
        "Sourced": (MINT, "#3E7A63"),
        "Ordered": (MINT, "#3E7A63"),
    }
    bg, fg = styles.get(s, (GOLD_BG, GOLD_TEXT))
    return _pill(s, bg, fg)


def type_badge(t: str) -> str:
    t = (clean(t) or "Trusted").strip()
    if t.lower() == "trusted":
        return _pill("Trusted", MINT, "#3E7A63")
    return _pill("External", "#EEF1F5", MUTED_TEXT)


def cart_badge(v: str) -> str:
    v = (clean(v) or "No").strip()
    if v.lower() == "yes":
        return _pill("Yes", MINT, "#3E7A63")
    return _pill("No", "#F7E0E0", "#C94B4B")


def cmp_id(num) -> str:
    """Format the row number as a component id, e.g. 17 -> CMP-0017."""
    try:
        return f"CMP-{int(float(num)):04d}"
    except (ValueError, TypeError):
        return clean(num)


def sourcing_badge(status: str) -> str:
    s = (clean(status) or "Pending").strip()
    if s.lower() == "sourced":
        return _pill("Sourced", MINT, "#3E7A63")
    return _pill("Pending Sourcing", GOLD_BG, GOLD_TEXT)


# ----------------------------------------------------------------------------
# Logo: use your real file if present, otherwise an SVG recreation
# ----------------------------------------------------------------------------
def get_logo_html() -> str:
    """Return an <img> for a local logo file, or an SVG fallback."""
    here = os.path.dirname(os.path.abspath(__file__))
    for fname in ("logo.png", "wits_logo.png", "logo.jpg", "logo.jpeg", "logo.svg"):
        path = os.path.join(here, fname)
        if os.path.exists(path):
            with open(path, "rb") as f:
                data = base64.b64encode(f.read()).decode()
            mime = "image/png"
            if fname.endswith(".svg"):
                mime = "image/svg+xml"
            elif fname.endswith((".jpg", ".jpeg")):
                mime = "image/jpeg"
            return f'<img src="data:{mime};base64,{data}" class="logo-img" />'

    # --- SVG fallback recreation of the Wits EIE emblem ---------------------
    return f"""
    <svg class="logo-img" viewBox="0 0 220 150" xmlns="http://www.w3.org/2000/svg">
        <g fill="none" stroke="{PRIMARY_BLUE}" stroke-width="6">
            <circle cx="150" cy="45" r="38"/>
        </g>
        <g fill="{PRIMARY_BLUE}">
            <rect x="132" y="27" width="18" height="18" transform="rotate(45 141 36)"/>
            <rect x="150" y="45" width="18" height="18" transform="rotate(45 159 54)"/>
            <rect x="150" y="27" width="14" height="14" transform="rotate(45 157 34)" opacity="0.55"/>
        </g>
        <text x="6" y="108" font-family="Inter, sans-serif" font-size="11"
              font-weight="600" fill="{PRIMARY_BLUE}" letter-spacing="1">SCHOOL OF</text>
        <text x="6" y="124" font-family="Inter, sans-serif" font-size="15"
              font-weight="800" fill="{TEXT}">ELECTRICAL <tspan fill="{PRIMARY_BLUE}">AND</tspan></text>
        <text x="6" y="140" font-family="Inter, sans-serif" font-size="15"
              font-weight="800" fill="{TEXT}">INFORMATION</text>
        <text x="6" y="156" font-family="Inter, sans-serif" font-size="15"
              font-weight="800" fill="{TEXT}">ENGINEERING</text>
    </svg>
    """


# ----------------------------------------------------------------------------
# Global CSS
# ----------------------------------------------------------------------------
st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"], .stApp, p, span, div {{
            font-family: 'Inter', sans-serif;
            color: {TEXT};
        }}
        html {{ font-size: 14.5px; }}

        .stApp {{ background-color: {BACKGROUND}; }}

        /* Hide default Streamlit chrome */
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        header {{visibility: hidden;}}
        [data-testid="collapsedControl"] {{display: none;}}

        .block-container {{
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            max-width: 100%;
        }}

        /* -------------------------------------------------------------- */
        /* Native sidebar — fixed 300px, light blue                      */
        /* -------------------------------------------------------------- */
        section[data-testid="stSidebar"] {{
            width: 300px !important;
            min-width: 300px !important;
            max-width: 300px !important;
            background-color: {LIGHT_BLUE};
            border-right: 1px solid {BORDER};
        }}
        section[data-testid="stSidebar"] > div {{
            background-color: {LIGHT_BLUE};
        }}
        [data-testid="stSidebarUserContent"] {{
            padding: 4px 14px 8px 14px !important;
        }}

        /* Logo box (white card) */
        .logo-box {{
            background-color: {CARD_BG};
            border-radius: 14px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
            padding: 8px;
            text-align: center;
            margin: 0 24px 6px 24px;
        }}
        .logo-img {{
            width: 100%;
            max-width: 100px;
            height: auto;
            display: inline-block;
        }}

        /* Caption under logo */
        .sidebar-caption {{
            text-align: center;
            font-size: 11px;
            letter-spacing: 2px;
            color: {PRIMARY_BLUE};
            font-weight: 600;
            text-transform: uppercase;
            margin: 0 0 8px 0;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}

        .sidebar-divider {{
            border: none;
            border-top: 1px solid {BORDER};
            margin: 2px 6px 8px 6px;
        }}

        /* Bottom admin card */
        .admin-card {{
            background-color: {CARD_BG};
            border-radius: 14px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
            padding: 10px 12px;
            display: flex;
            align-items: center;
            gap: 8px;
            margin: 8px 4px 4px 4px;
            overflow: hidden;            /* keep contents inside the card */
        }}
        .admin-card > div:last-child {{
            flex: 1;                     /* take remaining space */
            min-width: 0;                /* allow shrinking below content */
        }}
        .admin-avatar {{
            width: 34px;
            height: 34px;
            border-radius: 50%;
            background-color: {PRIMARY_BLUE};
            color: #fff;
            font-weight: 700;
            font-size: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
        }}
        .admin-name {{
            font-size: 16px;
            font-weight: 600;
            color: {TEXT};
        }}
        .admin-email {{
            font-size: 12px;
            color: {MUTED_TEXT};
            display: block;
            white-space: normal;         /* wrap (never truncate an email) */
            word-break: break-word;      /* graceful fallback if font is wide */
        }}

        /* -------------------------------------------------------------- */
        /* Nav buttons — scoped to the sidebar                            */
        /* In Streamlit 1.35 the active button is marked kind="primary".  */
        /* -------------------------------------------------------------- */
        section[data-testid="stSidebar"] .stButton > button {{
            width: 100%;
            text-align: left;
            justify-content: flex-start;
            padding-left: 20px;
            background-color: transparent;   /* same as sidebar bg */
            color: {TEXT};
            border: none;
            border-radius: 12px;
            padding: 7px 16px;
            margin-bottom: 0px;
            box-shadow: none;
            transition: all 0.15s ease;
        }}
        /* Make the label text (and emoji icon) larger */
        section[data-testid="stSidebar"] .stButton > button p {{
            font-size: 17px !important;
            font-weight: 500 !important;
        }}
        section[data-testid="stSidebar"] .stButton > button:hover {{
            background-color: {SIDEBAR_ACCENT};
            color: {PRIMARY_BLUE};
        }}
        section[data-testid="stSidebar"] .stButton > button:hover p {{
            color: {PRIMARY_BLUE};
        }}
        section[data-testid="stSidebar"] .stButton > button:focus {{
            box-shadow: none;
            color: {TEXT};
        }}

        /* ACTIVE item (primary) — soft light-blue pill, blue BOLD text */
        section[data-testid="stSidebar"] .stButton > button[kind="primary"] {{
            background-color: {SIDEBAR_ACCENT};
            box-shadow: 0 2px 8px rgba(91,125,177,0.18);
        }}
        section[data-testid="stSidebar"] .stButton > button[kind="primary"] p {{
            color: {PRIMARY_BLUE} !important;
            font-weight: 700 !important;
        }}
        section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover p {{
            color: {PRIMARY_BLUE} !important;
        }}

        /* -------------------------------------------------------------- */
        /* Main content                                                   */
        /* -------------------------------------------------------------- */
        .card {{
            background-color: {CARD_BG};
            border-radius: 16px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
            border: 1px solid {BORDER};
            padding: 20px 24px;
            margin-bottom: 18px;
        }}
        .page-title {{
            font-size: 26px;
            font-weight: 700;
            color: {TEXT};
            margin-bottom: 4px;
        }}
        .page-subtitle {{
            font-size: 14px;
            color: {MUTED_TEXT};
            margin-bottom: 18px;
        }}
        .placeholder {{
            font-size: 15px;
            color: {MUTED_TEXT};
            line-height: 1.6;
        }}

        /* -------------------------------------------------------------- */
        /* Form inputs (Wishlist) — light-blue fill, soft blue border     */
        /* -------------------------------------------------------------- */
        .stTextInput input,
        .stNumberInput input {{
            background-color: {INPUT_BG} !important;
            border: 1px solid {BORDER} !important;
            border-radius: 10px !important;
            color: {TEXT} !important;
            font-size: 14px !important;
            padding: 9px 12px !important;
            min-height: 40px !important;
        }}
        .stTextInput input:focus,
        .stNumberInput input:focus {{
            border-color: {PRIMARY_BLUE} !important;
            box-shadow: 0 0 0 3px rgba(91,125,177,0.18) !important;
        }}
        .stTextInput label, .stNumberInput label {{
            font-size: 13px !important;
            font-weight: 600 !important;
            color: {TEXT} !important;
        }}
        /* Hide the number-input +/- steppers for a cleaner look */
        .stNumberInput button {{
            display: none;
        }}

        /* "Add Component" submit button + download button (main area) */
        [data-testid="stFormSubmitButton"] button,
        .stDownloadButton button {{
            background-color: {PRIMARY_BLUE} !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 10px !important;
            padding: 9px 20px !important;
            box-shadow: 0 4px 12px rgba(91,125,177,0.30) !important;
        }}
        [data-testid="stFormSubmitButton"] button p,
        .stDownloadButton button p {{
            color: #ffffff !important;
            font-size: 14px !important;
            font-weight: 600 !important;
        }}
        [data-testid="stFormSubmitButton"] button:hover,
        .stDownloadButton button:hover {{
            background-color: #4E6D9C !important;
        }}

        /* Make the form render as a white card, identical to .card */
        [data-testid="stForm"] {{
            background-color: {CARD_BG} !important;
            border: 1px solid {BORDER} !important;
            border-radius: 16px !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06) !important;
            padding: 20px 24px !important;
            margin-bottom: 18px !important;
        }}

        /* Cards: primary path = our OWN stable class from st.container(key=...).
           Fallback path (older Streamlit) = innermost bordered wrapper holding
           a .card-marker. Both are listed so styling holds on any version. */
        [class*="st-key-appcard"],
        [data-testid="stVerticalBlockBorderWrapper"]:has(.card-marker):not(:has([data-testid="stVerticalBlockBorderWrapper"] .card-marker)) {{
            background-color: {CARD_BG} !important;
            border: 1px solid {BORDER} !important;
            border-radius: 16px !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06) !important;
            padding: 18px 20px !important;
            margin-bottom: 16px !important;
        }}
        .card-marker {{ display: none; }}

        /* Red "Delete" button — stable class (newer) OR .del-marker column
           (older). Both selectors listed so it works on any Streamlit version. */
        [class*="st-key-delcell"] button,
        [data-testid="column"]:has(.del-marker) button,
        [data-testid="stColumn"]:has(.del-marker) button {{
            background: #F7E0E0 !important;
            border: 1px solid #F0C9C9 !important;
            border-radius: 10px !important;
            box-shadow: none !important;
            min-height: 38px !important;
            padding: 4px 8px !important;
            white-space: nowrap !important;
        }}
        [class*="st-key-delcell"] button p,
        [data-testid="column"]:has(.del-marker) button p,
        [data-testid="stColumn"]:has(.del-marker) button p {{
            color: #C94B4B !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            white-space: nowrap !important;
        }}
        [class*="st-key-delcell"] button:hover,
        [data-testid="column"]:has(.del-marker) button:hover,
        [data-testid="stColumn"]:has(.del-marker) button:hover {{
            background: #F2C9C9 !important;
            border-color: #C94B4B !important;
        }}
        .del-marker {{ display: none; }}

        .card-heading {{
            font-size: 18px;
            font-weight: 700;
            color: {TEXT};
            margin-bottom: 16px;
        }}

        /* Sourcing card headers */
        .src-card-title {{
            font-size: 16px;
            font-weight: 700;
            color: {TEXT};
        }}
        .src-card-sub {{
            font-size: 13px;
            color: {MUTED_TEXT};
            margin-bottom: 6px;
        }}

        /* Supplier comparison "table" built from columns */
        .cmp-h {{
            font-size: 11px;
            letter-spacing: 1px;
            text-transform: uppercase;
            color: {MUTED_TEXT};
            font-weight: 600;
            padding: 6px 4px 10px 4px;
            border-bottom: 1px solid {BORDER};
        }}
        .cmp-c {{
            font-size: 14px;
            color: {TEXT};
            padding: 11px 4px;
            border-bottom: 1px solid {BORDER};
        }}
        .cmp-strong {{ font-weight: 700; }}
        .cmp-muted {{ color: {MUTED_TEXT}; }}
        .cmp-price {{ color: {PRIMARY_BLUE}; font-weight: 600; }}

        /* Sourcing Progress card */
        .progress-card {{
            background-color: {CARD_BG};
            border: 1px solid {BORDER};
            border-radius: 16px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
            padding: 16px 20px;
        }}
        .progress-top {{
            display: flex; justify-content: space-between; align-items: baseline;
        }}
        .progress-label {{ font-size: 13px; font-weight: 700; color: {TEXT}; }}
        .progress-pct {{ font-size: 13px; font-weight: 700; color: {PRIMARY_BLUE}; }}
        .progress-track {{
            height: 7px; background: {INPUT_BG}; border-radius: 999px;
            margin: 12px 0 8px 0; overflow: hidden;
        }}
        .progress-fill {{
            height: 100%; background: {PRIMARY_BLUE}; border-radius: 999px;
        }}
        .progress-sub {{ font-size: 13px; color: {MUTED_TEXT}; }}

        /* Component Queue */
        .queue-title {{ font-size: 16px; font-weight: 700; color: {TEXT}; }}
        .queue-sub {{
            font-size: 13px; color: {MUTED_TEXT};
            margin: 2px 0 12px 0; padding-bottom: 12px;
            border-bottom: 1px solid {BORDER};
        }}

        /* CMP id chip */
        .id-chip {{
            font-size: 12px; font-weight: 600; color: {MUTED_TEXT};
            background: {INPUT_BG}; border: 1px solid {BORDER};
            padding: 2px 10px; border-radius: 999px; vertical-align: middle;
        }}

        /* Supplier options header */
        .opt-head {{
            display: flex; justify-content: space-between; align-items: baseline;
        }}
        .opt-count {{ font-size: 13px; color: {MUTED_TEXT}; }}
        .opt-sub {{ font-size: 13px; color: {MUTED_TEXT}; margin: 2px 0 18px 0; }}

        /* Routing hint box */
        .route-hint {{
            background: {INPUT_BG};
            border: 1px dashed {BORDER};
            border-radius: 12px;
            padding: 12px 16px;
            text-align: center;
            font-size: 13px;
            color: {MUTED_TEXT};
        }}
        .route-active {{
            border-style: solid;
            background: {MINT};
            color: #3E7A63;
        }}

        /* Wishlist table */
        .wl-table {{
            width: 100%;
            border-collapse: collapse;
        }}
        .wl-table th {{
            text-align: left;
            font-size: 12px;
            letter-spacing: 1px;
            text-transform: uppercase;
            color: {MUTED_TEXT};
            font-weight: 600;
            padding: 12px 12px;
            border-bottom: 1px solid {BORDER};
        }}
        .wl-table td {{
            padding: 16px 12px;
            font-size: 16px;
            color: {TEXT};
            border-bottom: 1px solid {BORDER};
        }}
        .wl-table tr:last-child td {{ border-bottom: none; }}
        .wl-comp {{ font-weight: 600; }}
        .wl-muted {{ color: {MUTED_TEXT}; }}
        .wl-empty {{
            font-size: 14px;
            color: {MUTED_TEXT};
            padding: 8px 4px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Navigation state
# ----------------------------------------------------------------------------
PAGES = [
    ("Dashboard", "▦"),
    ("Wishlist", "☆"),
    ("Sourcing", "🔍"),
    ("Procurement Outputs", "🛒"),
]

if "active_page" not in st.session_state:
    st.session_state.active_page = "Dashboard"


def set_page(name: str) -> None:
    st.session_state.active_page = name


_CARD_SEQ = 0


def card():
    """A white 'card' container, styled two crash-safe ways.

    Newer Streamlit (>=1.39): st.container(key=...) stamps a stable, app-owned
    class `st-key-appcard_*` that we target directly — no :has(), no testids.

    Older Streamlit (no `key` arg): we fall back to a hidden .card-marker and a
    :has() selector. Either way the same white-card styling applies, and we
    never crash on `key` not being supported.
    """
    global _CARD_SEQ
    _CARD_SEQ += 1
    try:
        c = st.container(border=True, key=f"appcard_{_CARD_SEQ}")
    except TypeError:
        c = st.container(border=True)
    c.markdown('<span class="card-marker"></span>', unsafe_allow_html=True)
    return c


# ----------------------------------------------------------------------------
# Wishlist page
# ----------------------------------------------------------------------------
def render_wishlist() -> None:
    st.markdown(
        """
        <div class="page-title">Component Wishlist</div>
        <div class="page-subtitle">Add components needed for upcoming lab
        experiments and projects.</div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Add New Component form (the st.form itself is the white card) ---
    with st.form("add_component_form", clear_on_submit=True):
        st.markdown('<div class="card-heading">Add New Component</div>',
                    unsafe_allow_html=True)

        c1, c2 = st.columns(2, gap="large")
        with c1:
            component = st.text_input("Component Name *",
                                      placeholder="e.g. NE555 Timer IC")
        with c2:
            model = st.text_input("Model Number", placeholder="e.g. NE555P")

        c3, c4 = st.columns(2, gap="large")
        with c3:
            specs = st.text_input("Specification",
                                  placeholder="e.g. Single, 4.5–16V, DIP-8")
        with c4:
            qty = st.number_input("Quantity *", min_value=1, max_value=100000,
                                  value=1, step=1)

        submitted = st.form_submit_button("+  Add Component")
        if submitted:
            if not component.strip():
                st.warning("Please enter a component name.")
            else:
                add_component(component.strip(), model.strip(),
                              specs.strip(), qty)
                st.success(f"“{component.strip()}” added to the wishlist.")

    # show any delete confirmation message from the previous run
    if st.session_state.get("wl_msg"):
        st.success(st.session_state.pop("wl_msg"))

    # ---- Wishlist table -------------------------------------------------
    df = load_wishlist()
    count = len(df)

    with card():
        st.markdown(
            f'<div class="card-heading">Wishlist — {count} '
            f'item{"s" if count != 1 else ""}</div>',
            unsafe_allow_html=True,
        )

        # Pending-delete confirmation banner
        pending = st.session_state.get("wl_pending_delete")
        if pending is not None:
            match = df[df["#"].astype(str) == str(pending)]
            if match.empty:
                st.session_state.pop("wl_pending_delete", None)
            else:
                pname = clean(match.iloc[0]["Component"])
                st.markdown(
                    f'<div class="route-hint" style="border-color:#C94B4B;'
                    f'background:#F7E0E0;color:#C94B4B;">Delete '
                    f'<b>{pname}</b> ({cmp_id(pending)}) and all of its sourcing '
                    f'options? This cannot be undone.</div>',
                    unsafe_allow_html=True,
                )
                bc1, bc2, _ = st.columns([1.2, 1, 4])
                with bc1:
                    if st.button("Yes, delete", key="wl_confirm_del",
                                 type="primary", use_container_width=True):
                        delete_component(pending)
                        st.session_state.pop("wl_pending_delete", None)
                        st.session_state.wl_msg = f"“{pname}” removed from the wishlist."
                        st.rerun()
                with bc2:
                    if st.button("Cancel", key="wl_cancel_del",
                                 use_container_width=True):
                        st.session_state.pop("wl_pending_delete", None)
                        st.rerun()

        if count == 0:
            st.markdown(
                '<div class="wl-empty">No components yet — add your first one '
                'above.</div>',
                unsafe_allow_html=True,
            )
        else:
            widths = [2.1, 1.2, 2.2, 0.7, 1.4, 1.2, 1.0]
            heads = ["COMPONENT", "MODEL", "SPECIFICATION", "QTY",
                     "DATE ADDED", "STATUS", "ACTION"]
            hc = st.columns(widths)
            for col, h in zip(hc, heads):
                col.markdown(f'<div class="cmp-h">{h}</div>',
                             unsafe_allow_html=True)

            for _, r in df.iterrows():
                row = st.columns(widths)
                row[0].markdown(
                    f'<div class="cmp-c cmp-strong">{clean(r["Component"])}</div>',
                    unsafe_allow_html=True)
                row[1].markdown(
                    f'<div class="cmp-c cmp-muted">{clean(r["Model"])}</div>',
                    unsafe_allow_html=True)
                row[2].markdown(f'<div class="cmp-c">{clean(r["Specifications"])}</div>',
                                unsafe_allow_html=True)
                row[3].markdown(f'<div class="cmp-c">{clean(r["Quantity"])}</div>',
                                unsafe_allow_html=True)
                row[4].markdown(
                    f'<div class="cmp-c cmp-muted">{clean(r["Date Added"])}</div>',
                    unsafe_allow_html=True)
                row[5].markdown(f'<div class="cmp-c">{status_badge(r["Status"])}</div>',
                                unsafe_allow_html=True)
                with row[6]:
                    try:
                        cont = st.container(key=f"delcell_{r['#']}")
                    except TypeError:
                        cont = st.container()
                    with cont:
                        st.markdown('<span class="del-marker"></span>',
                                    unsafe_allow_html=True)
                        if st.button("Delete", key=f"wl_del_{r['#']}",
                                     use_container_width=True):
                            st.session_state.wl_pending_delete = r["#"]
                            st.rerun()

    # ---- Download -------------------------------------------------------
    if EXCEL_FILE.exists():
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                "⬇  Download wishlist (.xlsx)",
                data=f,
                file_name="wishlist.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ----------------------------------------------------------------------------
# Sourcing page
#
# Records supplier options the USER has checked manually. No web scraping, no
# automatic price comparison — the system stores options and the user chooses.
# ----------------------------------------------------------------------------
def render_sourcing() -> None:
    wl = load_wishlist()
    total = len(wl)
    sourced = int(sum(clean(s).lower() == "sourced" for s in wl["Status"])) if total else 0
    pending = total - sourced
    pct = int(round(sourced / total * 100)) if total else 0

    # --- Header row: title + progress card ------------------------------
    head_l, head_r = st.columns([2.4, 1], gap="medium")
    with head_l:
        st.markdown(
            """
            <div class="page-title">Sourcing</div>
            <div class="page-subtitle">Select a component → review supplier options
            → confirm source.</div>
            """,
            unsafe_allow_html=True,
        )
    with head_r:
        st.markdown(
            f"""
            <div class="progress-card">
                <div class="progress-top">
                    <span class="progress-label">Sourcing Progress</span>
                    <span class="progress-pct">{pct}%</span>
                </div>
                <div class="progress-track">
                    <div class="progress-fill" style="width:{pct}%;"></div>
                </div>
                <div class="progress-sub">{sourced} of {total} components sourced</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if wl.empty:
        st.markdown(
            '<div class="card"><div class="wl-empty">No components yet — add some '
            'on the <b>Wishlist</b> page first.</div></div>',
            unsafe_allow_html=True,
        )
        return

    if st.session_state.get("sourcing_msg"):
        st.success(st.session_state.pop("sourcing_msg"))

    n = total
    if st.session_state.get("sourcing_idx", 0) >= n:
        st.session_state.sourcing_idx = 0
    idx = st.session_state.get("sourcing_idx", 0)

    # Selection state for the CURRENT component (computed up-front)
    comp = wl.iloc[idx]
    comp_id = comp["#"]
    opts = load_options()
    comp_opts = opts[opts["Component ID"].astype(str) == str(comp_id)]
    sel_rows = comp_opts[comp_opts["Selected"].apply(is_true)]
    has_selection = not sel_rows.empty

    # --- Main row: component queue (left) + sourcing content (right) -----
    qcol, ccol = st.columns([1.15, 3.3], gap="medium")

    # ---- Component Queue ----
    with qcol:
        with card():
            st.markdown(
                f'<div class="queue-title">Component Queue</div>'
                f'<div class="queue-sub">{pending} pending · {sourced} sourced</div>',
                unsafe_allow_html=True,
            )
            for i, (_, r) in enumerate(wl.iterrows()):
                status = clean(r["Status"]) or "Pending"
                is_cur = (i == idx)
                if status.lower() == "sourced":
                    dot = "🟢"
                elif is_cur:
                    dot = "🔵"
                else:
                    dot = "⚪"
                label = f"{dot}  {clean(r['Component'])}   ·   {cmp_id(r['#'])}"
                if st.button(label, key=f"q_{i}", use_container_width=True,
                             type="primary" if is_cur else "secondary"):
                    st.session_state.sourcing_idx = i
                    st.rerun()

    # ---- Sourcing content ----
    with ccol:
        # Component card: info | Skip | Confirm  (single level of columns)
        with card():
            ci, cs, cc = st.columns([6.4, 1.2, 1.7])
            with ci:
                st.markdown(
                    f"""
                    <div style="display:flex;align-items:center;gap:14px;">
                        <div style="width:40px;height:40px;border-radius:10px;
                            background:{INPUT_BG};display:flex;align-items:center;
                            justify-content:center;font-size:18px;">🔌</div>
                        <div>
                            <div style="font-size:17px;font-weight:700;color:{TEXT};">
                                {clean(comp["Component"])}
                                &nbsp;{sourcing_badge(comp["Status"])}
                                &nbsp;<span class="id-chip">{cmp_id(comp["#"])}</span>
                            </div>
                            <div style="font-size:13px;color:{MUTED_TEXT};margin-top:5px;">
                                <b style="color:{MUTED_TEXT};font-weight:600;">MODEL</b>
                                &nbsp;{clean(comp["Model"])} &nbsp;&nbsp;|&nbsp;&nbsp;
                                <b style="color:{MUTED_TEXT};font-weight:600;">SPEC</b>
                                &nbsp;{clean(comp["Specifications"])} &nbsp;&nbsp;|&nbsp;&nbsp;
                                <b style="color:{MUTED_TEXT};font-weight:600;">QTY</b>
                                &nbsp;<b style="color:{TEXT};">{clean(comp["Quantity"])}</b>
                            </div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with cs:
                if st.button("Skip", key="src_skip", use_container_width=True):
                    st.session_state.sourcing_idx = (idx + 1) % n
                    st.rerun()
            with cc:
                if st.button("Confirm Source", key="src_confirm", type="primary",
                             use_container_width=True, disabled=not has_selection):
                    supplier = confirm_source(comp_id)
                    route = route_for(sel_rows.iloc[0]["Supplier"],
                                      sel_rows.iloc[0]["Shopping Cart Available"])
                    st.session_state.sourcing_msg = (
                        f"“{clean(comp['Component'])}” sourced from {supplier} "
                        f"→ routed to {route}."
                    )
                    st.session_state.sourcing_idx = min(idx + 1, n - 1)
                    st.rerun()

        # Two sourcing-entry cards (side by side; inner fields stack)
        left, right = st.columns(2, gap="large")

        with left:
            with card():
                st.markdown(
                    '<div class="src-card-title">Trusted Supplier</div>'
                    '<div class="src-card-sub">Pre-approved local vendors</div>',
                    unsafe_allow_html=True,
                )
                t_on = st.toggle("Enable trusted supplier entry", value=True,
                                 key="t_on")
                supplier = st.selectbox("Supplier", TRUSTED_SUPPLIERS,
                                        disabled=not t_on, key="t_supplier")
                t_stock = st.text_input("Stock", placeholder="1,240",
                                        disabled=not t_on, key="t_stock")
                t_eta = st.text_input("ETA", placeholder="3–5 days",
                                      disabled=not t_on, key="t_eta")
                t_price = st.text_input("Unit Price", placeholder="R 4.20",
                                        disabled=not t_on, key="t_price")
                t_cart = st.toggle("Shopping Cart Available", value=True,
                                   disabled=not t_on, key="t_cart")
                if st.button("+  Add to comparison", key="t_add",
                             disabled=not t_on, use_container_width=True):
                    if not clean(t_stock) and not clean(t_price):
                        st.warning("Enter at least the stock or unit price.")
                    else:
                        add_option(comp_id, supplier, "Trusted", t_price, t_stock,
                                   t_eta, "Yes" if t_cart else "No", "")
                        st.rerun()

        with right:
            with card():
                st.markdown(
                    '<div class="src-card-title">URL Search</div>'
                    '<div class="src-card-sub">Manually record details from any '
                    'vendor page</div>',
                    unsafe_allow_html=True,
                )
                u_on = st.toggle("Enable URL search entry", value=False, key="u_on")
                u_url = st.text_input("Product URL",
                                      placeholder="https://www.digikey.com/...",
                                      disabled=not u_on, key="u_url")
                u_vendor = st.text_input("Vendor", placeholder="e.g. DigiKey",
                                         disabled=not u_on, key="u_vendor")
                u_price = st.text_input("Price", placeholder="R 2.90",
                                        disabled=not u_on, key="u_price")
                u_stock = st.text_input("Stock", placeholder="9,800",
                                        disabled=not u_on, key="u_stock")
                u_eta = st.text_input("ETA", placeholder="7–10 days",
                                      disabled=not u_on, key="u_eta")
                u_cart = st.toggle("Shopping Cart Available", value=False,
                                   disabled=not u_on, key="u_cart")
                if st.button("+  Add to comparison", key="u_add",
                             disabled=not u_on, use_container_width=True):
                    if not clean(u_vendor):
                        st.warning("Enter a vendor name.")
                    else:
                        add_option(comp_id, u_vendor, "External", u_price, u_stock,
                                   u_eta, "Yes" if u_cart else "No", u_url)
                        st.rerun()

        # Supplier options table
        with card():
            st.markdown(
                f'<div class="opt-head">'
                f'<span class="card-heading" style="margin:0;">Supplier Options for '
                f'<span style="color:{PRIMARY_BLUE};">{clean(comp["Component"])}</span></span>'
                f'<span class="opt-count">{len(comp_opts)} option'
                f'{"s" if len(comp_opts) != 1 else ""} found</span></div>'
                f'<div class="opt-sub">{cmp_id(comp["#"])} &nbsp;·&nbsp; '
                f'{clean(comp["Model"])}</div>',
                unsafe_allow_html=True,
            )

            if comp_opts.empty:
                st.markdown(
                    '<div class="wl-empty">No supplier options recorded yet — add one '
                    'using the cards above.</div>',
                    unsafe_allow_html=True,
                )
            else:
                widths = [2.4, 1.3, 1, 1.3, 1.2, 1.2, 1.1]
                heads = ["SUPPLIER", "TYPE", "STOCK", "ETA", "UNIT PRICE",
                         "CART", "ACTION"]
                hc = st.columns(widths)
                for col, h in zip(hc, heads):
                    col.markdown(f'<div class="cmp-h">{h}</div>',
                                 unsafe_allow_html=True)

                for ri, r in comp_opts.iterrows():
                    row = st.columns(widths)
                    row[0].markdown(
                        f'<div class="cmp-c cmp-strong">{clean(r["Supplier"])}</div>',
                        unsafe_allow_html=True)
                    row[1].markdown(
                        f'<div class="cmp-c">{type_badge(r["Supplier Type"])}</div>',
                        unsafe_allow_html=True)
                    row[2].markdown(f'<div class="cmp-c">{clean(r["Stock"])}</div>',
                                    unsafe_allow_html=True)
                    row[3].markdown(
                        f'<div class="cmp-c cmp-muted">{clean(r["ETA"])}</div>',
                        unsafe_allow_html=True)
                    row[4].markdown(
                        f'<div class="cmp-c cmp-price">{clean(r["Price"])}</div>',
                        unsafe_allow_html=True)
                    row[5].markdown(
                        f'<div class="cmp-c">{cart_badge(r["Shopping Cart Available"])}</div>',
                        unsafe_allow_html=True)
                    is_sel = is_true(r["Selected"])
                    with row[6]:
                        if st.button("✓ Selected" if is_sel else "Select",
                                     key=f"sel_{ri}",
                                     type="primary" if is_sel else "secondary",
                                     use_container_width=True):
                            select_option(comp_id, ri)
                            st.rerun()

        # Routing preview / hint
        if has_selection:
            chosen = sel_rows.iloc[0]
            route = route_for(chosen["Supplier"], chosen["Shopping Cart Available"])
            st.markdown(
                f'<div class="route-hint route-active"><b>{clean(chosen["Supplier"])}</b>'
                f' selected → routes to <b>{route}</b>. '
                f'Click <b>Confirm Source</b> to lock the decision.</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="route-hint">Select a supplier above to preview routing '
                'and confirm the source.</div>',
                unsafe_allow_html=True,
            )


# ----------------------------------------------------------------------------
# Procurement Outputs page (step 3)
# ----------------------------------------------------------------------------
def _chosen_option(opts, comp_id):
    sel = opts[(opts["Component ID"].astype(str) == str(comp_id))
               & (opts["Selected"].apply(is_true))]
    return sel.iloc[0] if not sel.empty else None


def _missing_fields(comp, chosen):
    """List the blank fields for a component (wishlist + chosen supplier)."""
    missing = []
    if not clean(comp["Model"]):
        missing.append("Model")
    if not clean(comp["Specifications"]):
        missing.append("Specification")
    if not clean(comp["Quantity"]):
        missing.append("Quantity")
    if clean(comp["Status"]).lower() != "sourced":
        missing.append("Not sourced")
    elif chosen is None:
        missing.append("No supplier selected")
    else:
        if not clean(chosen["Price"]):
            missing.append("Unit Price")
        if not clean(chosen["Stock"]):
            missing.append("Stock")
        if not clean(chosen["ETA"]):
            missing.append("ETA")
    return missing


def _table(headers, rows_html):
    head = "".join(f'<th class="cmp-h">{h}</th>' for h in headers)
    return (f'<table style="width:100%;border-collapse:collapse;">'
            f'<thead><tr>{head}</tr></thead><tbody>{rows_html}</tbody></table>')


def render_procurement():
    st.markdown(
        '<div class="page-title">Procurement Outputs</div>'
        '<div class="page-subtitle">Review completeness, then route sourced '
        'components to carts and order lists.</div>',
        unsafe_allow_html=True,
    )
    wl = load_wishlist()
    opts = load_options()
    if wl.empty:
        st.markdown(
            '<div class="card"><div class="wl-empty">No components yet — add some '
            'on the <b>Wishlist</b> page first.</div></div>',
            unsafe_allow_html=True,
        )
        return

    # ---- Completeness reminder (non-blocking) --------------------------
    incomplete = []
    for _, comp in wl.iterrows():
        miss = _missing_fields(comp, _chosen_option(opts, comp["#"]))
        if miss:
            incomplete.append((comp, miss))

    if incomplete:
        rows = ""
        for comp, miss in incomplete:
            rows += (
                "<tr>"
                f'<td class="cmp-c cmp-strong">{clean(comp["Component"])}</td>'
                f'<td class="cmp-c cmp-muted">{cmp_id(comp["#"])}</td>'
                f'<td class="cmp-c" style="color:{GOLD_TEXT};">{", ".join(miss)}</td>'
                "</tr>"
            )
        st.markdown(
            f'<div class="card" style="background:#FEFBF2;border-color:#EAD9A6;">'
            f'<div class="card-heading" style="color:{GOLD_TEXT};margin-bottom:6px;">'
            f'⚠ Reminder — {len(incomplete)} item'
            f'{"s" if len(incomplete) != 1 else ""} still have missing data</div>'
            f'<div class="page-subtitle" style="margin-bottom:14px;">This is only a '
            f'reminder — you can still continue and use the outputs below.</div>'
            + _table(["COMPONENT", "ID", "MISSING FIELDS"], rows)
            + '</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="card"><div style="color:#3E7A63;font-weight:600;">'
            '✓ All components have complete data.</div></div>',
            unsafe_allow_html=True,
        )

    # ---- Route sourced components into buckets -------------------------
    buckets = {"Shopping Cart Queue": [], "Veronica Procurement List": [],
               "Tan Procurement List": [], "Unresolved Components": []}
    for _, comp in wl.iterrows():
        chosen = _chosen_option(opts, comp["#"])
        if clean(comp["Status"]).lower() == "sourced" and chosen is not None:
            supplier = clean(chosen["Supplier"])
            cart = clean(chosen["Shopping Cart Available"])
            route = route_for(supplier, cart)
        else:
            supplier, cart, route = "", "", "Unresolved Components"
        if route not in buckets:
            route = "Unresolved Components"
        buckets[route].append((comp, supplier, cart))

    # ---- Shopping Cart Queue -------------------------------------------
    cart_items = buckets["Shopping Cart Queue"]
    rows = "".join(
        "<tr>"
        f'<td class="cmp-c cmp-strong">{clean(c["Component"])}</td>'
        f'<td class="cmp-c">{sup}</td>'
        f'<td class="cmp-c">{clean(c["Quantity"])}</td>'
        f'<td class="cmp-c">{cart_badge("Yes")}</td>'
        "</tr>"
        for (c, sup, cart) in cart_items
    ) or '<tr><td colspan="4"><div class="wl-empty">Nothing here yet.</div></td></tr>'
    st.markdown(
        f'<div class="card"><div class="card-heading">🛒 Shopping Cart Queue '
        f'<span class="opt-count">({len(cart_items)})</span></div>'
        + _table(["COMPONENT", "SUPPLIER", "QTY", "CART"], rows) + '</div>',
        unsafe_allow_html=True,
    )
    lect = st.text_input("Lecturer email", placeholder="lecturer@wits.ac.za",
                         key="lect_email")
    if cart_items:
        body = "Shopping cart for procurement:\n\n" + "\n".join(
            f'- {clean(c["Component"])} (qty {clean(c["Quantity"])}) — {sup}'
            for (c, sup, cart) in cart_items)
        mailto = (f'mailto:{quote(lect or "")}?subject='
                  f'{quote("Procurement shopping cart")}&body={quote(body)}')
        st.markdown(
            f'<a href="{mailto}" target="_blank" style="display:inline-block;'
            f'background:{PRIMARY_BLUE};color:#fff;padding:9px 20px;border-radius:10px;'
            f'font-size:14px;font-weight:600;text-decoration:none;">'
            f'✉ Email Cart to Lecturer</a>',
            unsafe_allow_html=True,
        )

    # ---- Manual Order Lists (per buyer) --------------------------------
    st.markdown('<div class="card-heading" style="margin-top:6px;">📋 Manual Order '
                'Lists</div>', unsafe_allow_html=True)
    for route_name, subtitle in [
        ("Veronica Procurement List", "RS Components"),
        ("Tan Procurement List", "Communica · Mintech · Mantech"),
    ]:
        items = buckets[route_name]
        rows = "".join(
            "<tr>"
            f'<td class="cmp-c cmp-strong">{clean(c["Component"])}</td>'
            f'<td class="cmp-c">{sup}</td>'
            f'<td class="cmp-c">{clean(c["Quantity"])}</td>'
            "</tr>"
            for (c, sup, cart) in items
        ) or '<tr><td colspan="3"><div class="wl-empty">Nothing here yet.</div></td></tr>'
        st.markdown(
            f'<div class="card"><div class="card-heading" style="margin-bottom:2px;">'
            f'{route_name} <span class="opt-count">({len(items)})</span></div>'
            f'<div class="opt-sub">{subtitle}</div>'
            + _table(["COMPONENT", "SUPPLIER", "QTY"], rows) + '</div>',
            unsafe_allow_html=True,
        )

    # ---- Unresolved ----------------------------------------------------
    unresolved = buckets["Unresolved Components"]
    rows = "".join(
        "<tr>"
        f'<td class="cmp-c cmp-strong">{clean(c["Component"])}</td>'
        f'<td class="cmp-c cmp-muted">{cmp_id(c["#"])}</td>'
        f'<td class="cmp-c cmp-muted">'
        f'{"Not sourced" if clean(c["Status"]).lower() != "sourced" else "No route"}</td>'
        "</tr>"
        for (c, sup, cart) in unresolved
    ) or '<tr><td colspan="3"><div class="wl-empty">Nothing here yet.</div></td></tr>'
    st.markdown(
        f'<div class="card"><div class="card-heading">❓ Unresolved Components '
        f'<span class="opt-count">({len(unresolved)})</span></div>'
        + _table(["COMPONENT", "ID", "REASON"], rows) + '</div>',
        unsafe_allow_html=True,
    )

    # ---- Export to Excel ----------------------------------------------
    summary = []
    for _, comp in wl.iterrows():
        chosen = _chosen_option(opts, comp["#"])
        sup = clean(chosen["Supplier"]) if chosen is not None else ""
        cart = clean(chosen["Shopping Cart Available"]) if chosen is not None else ""
        sourced = clean(comp["Status"]).lower() == "sourced"
        route = route_for(sup, cart) if (sourced and chosen is not None) \
            else "Unresolved Components"
        summary.append({
            "ID": cmp_id(comp["#"]),
            "Component": clean(comp["Component"]),
            "Model": clean(comp["Model"]),
            "Specification": clean(comp["Specifications"]),
            "Quantity": clean(comp["Quantity"]),
            "Status": clean(comp["Status"]),
            "Supplier": sup,
            "Unit Price": clean(chosen["Price"]) if chosen is not None else "",
            "Cart Available": cart,
            "Route": route,
            "Missing": ", ".join(_missing_fields(comp, chosen)),
        })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(summary).to_excel(writer, sheet_name="Procurement Outputs",
                                       index=False)
    st.download_button("⬇  Export to Excel", data=buf.getvalue(),
                       file_name="procurement_outputs.xlsx",
                       mime="application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet")


# ----------------------------------------------------------------------------
# Layout: native sidebar + top-level main content
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown(
        f"""
        <div class="logo-box">{get_logo_html()}</div>
        <div class="sidebar-caption">Component&nbsp;&nbsp;Procurement</div>
        <hr class="sidebar-divider" />
        """,
        unsafe_allow_html=True,
    )

    for name, icon in PAGES:
        is_active = st.session_state.active_page == name
        st.button(
            name,
            key=f"nav_{name}",
            use_container_width=True,
            type="primary" if is_active else "secondary",
            on_click=set_page,
            args=(name,),
        )

    st.markdown(
        """
        <div class="admin-card">
            <div class="admin-avatar">AK</div>
            <div>
                <div class="admin-name">Admin</div>
                <div class="admin-email">procurement@wits.ac.za</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if _use_gsheets():
        _store_label, _store_bg, _store_fg = "Shared storage · Google Sheets", MINT, "#3E7A63"
    else:
        _store_label, _store_bg, _store_fg = "Temporary storage · local file", GOLD_BG, GOLD_TEXT
    st.markdown(
        f'<div style="text-align:center;margin-top:8px;">'
        f'<span style="background:{_store_bg};color:{_store_fg};font-size:11px;'
        f'font-weight:600;padding:3px 10px;border-radius:999px;">{_store_label}</span>'
        f'</div>'
        f'<div style="text-align:center;font-size:11px;color:{MUTED_TEXT};'
        f'margin-top:4px;">{APP_VERSION}</div>',
        unsafe_allow_html=True,
    )

# ---- Main content ------------------------------------------------------------
PAGE_CONTENT = {
    "Dashboard": (
        "Procurement Dashboard",
        "Faculty of Engineering & the Built Environment — June 2026",
        "Dashboard placeholder — overview metrics, workflow progress and recent "
        "components will appear here.",
    ),
    "Wishlist": (
        "Wishlist",
        "Components identified for the active procurement cycle",
        "Wishlist placeholder — the list of requested components will appear here.",
    ),
    "Sourcing": (
        "Sourcing",
        "Suppliers assigned and quotes gathered",
        "Sourcing placeholder — supplier assignment and quote comparison will "
        "appear here.",
    ),
    "Procurement Outputs": (
        "Procurement Outputs",
        "Orders placed and deliveries confirmed",
        "Procurement Outputs placeholder — purchase orders and delivery status "
        "will appear here.",
    ),
}

page = st.session_state.active_page
if page == "Wishlist":
    render_wishlist()
elif page == "Sourcing":
    render_sourcing()
elif page == "Procurement Outputs":
    render_procurement()
else:
    title, subtitle, body = PAGE_CONTENT[page]
    st.markdown(
        f"""
        <div class="page-title">{title}</div>
        <div class="page-subtitle">{subtitle}</div>
        <div class="card">
            <div class="placeholder">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
