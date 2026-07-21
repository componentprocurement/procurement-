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
APP_VERSION = "build 2026-06-25 #17 (outputs-cart-vs-supplier)"

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

SUPPLIERS = [
    "RS Components", "Communica", "Mantech Electronics", "Micro Robotics",
    "Netram Technologies", "DIY Electronics", "PiShop", "Electrocomp",
    "Actum Electronics", "Mouser", "DigiKey", "Other",
]

# Default recipient for the "Email to Lecturer" buttons (editable in the app).
LECTURER_EMAIL = "2433692@students.wits.ac.za"


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
def _open_spreadsheet():
    """Authorize and open the shared spreadsheet (cached)."""
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
    return gc.open_by_key(key)


@st.cache_resource(show_spinner=False)
def _get_worksheets():
    """Open (and lazily create) the two worksheets in the shared spreadsheet."""
    import gspread
    sh = _open_spreadsheet()

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


@st.cache_data(ttl=30, show_spinner=False)
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


def add_components_bulk(items) -> int:
    """Add many components in ONE write. items: list of (comp, model, spec, qty).

    Much faster than adding one at a time (a single Sheets API call).
    """
    date = datetime.now().strftime("%Y-%m-%d")
    if _use_gsheets():
        wl_ws, _ = _get_worksheets()
        start = _next_id(wl_ws.get_all_records())
        rows = [[start + i, c, m, s, int(q), date, "Pending", ""]
                for i, (c, m, s, q) in enumerate(items)]
        if rows:
            wl_ws.append_rows(rows, value_input_option="USER_ENTERED")
        _refresh()
        return len(rows)
    wl = load_wishlist()
    opts = load_options()
    start = _next_id(wl)
    new_rows = [{
        "#": start + i, "Component": c, "Model": m, "Specifications": s,
        "Quantity": int(q), "Date Added": date, "Status": "Pending",
        "Selected Supplier": "",
    } for i, (c, m, s, q) in enumerate(items)]
    wl = pd.concat([wl, pd.DataFrame(new_rows)], ignore_index=True)
    write_workbook(wl, opts)
    return len(new_rows)


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
        import gspread
        _, opt_ws = _get_worksheets()
        records = opt_ws.get_all_records()
        sel_col = OPT_COLUMNS.index("Selected") + 1
        cells = [gspread.Cell(i + 2, sel_col, "TRUE" if i == row_index else "FALSE")
                 for i, rec in enumerate(records)
                 if str(rec.get("Component ID")) == str(component_id)]
        if cells:
            opt_ws.update_cells(cells)   # one API call instead of one per row
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
        import gspread
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
                wl_ws.update_cells([gspread.Cell(i + 2, status_col, "Sourced"),
                                    gspread.Cell(i + 2, supp_col, supplier)])
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


def update_component(component_id, component, model, specs, qty) -> None:
    """Edit an existing component's wishlist fields."""
    if _use_gsheets():
        wl_ws, _ = _get_worksheets()
        cols = {name: WL_COLUMNS.index(name) + 1 for name in
                ("Component", "Model", "Specifications", "Quantity")}
        for i, rec in enumerate(wl_ws.get_all_records()):
            if str(rec.get("#")) == str(component_id):
                r = i + 2
                wl_ws.update_cell(r, cols["Component"], component)
                wl_ws.update_cell(r, cols["Model"], model)
                wl_ws.update_cell(r, cols["Specifications"], specs)
                wl_ws.update_cell(r, cols["Quantity"], int(qty))
                break
        _refresh()
        return
    wl = load_wishlist()
    opts = load_options()
    mask = wl["#"].astype(str) == str(component_id)
    wl.loc[mask, "Component"] = component
    wl.loc[mask, "Model"] = model
    wl.loc[mask, "Specifications"] = specs
    wl.loc[mask, "Quantity"] = int(qty)
    write_workbook(wl, opts)


ORDER_HEADERS = ["Component", "Model", "Specification", "Quantity",
                 "Unit Price", "Line Total", "URL"]


def save_company_order_sheets(orders_by_company) -> list:
    """Save each company's aggregated manual order to its own sheet/tab.

    orders_by_company: {company_name: [ {Component, Model, Spec, Qty, Price,
    Total}, ... ]}. Returns the list of sheet/tab names written.
    """
    def _rows(rows):
        return [ORDER_HEADERS] + [
            [r["Component"], r["Model"], r["Spec"], r["Qty"],
             _money(r["Price"]), _money(r["Total"]), r.get("URL", "")]
            for r in rows]

    saved = []
    if _use_gsheets():
        sh = _open_spreadsheet()
        for company, rows in orders_by_company.items():
            title = ("Order - " + (company or "Unknown"))[:99]
            try:
                ws = sh.worksheet(title)
                ws.clear()
            except Exception:
                ws = sh.add_worksheet(title, rows=max(20, len(rows) + 5),
                                      cols=len(ORDER_HEADERS))
            ws.update(range_name="A1", values=_rows(rows))
            saved.append(title)
        return saved

    # Local Excel fallback → one workbook with a sheet per company
    out = EXCEL_FILE.parent / "manual_orders.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for company, rows in orders_by_company.items():
            data = _rows(rows)
            pd.DataFrame(data[1:], columns=data[0]).to_excel(
                writer, sheet_name=("Order - " + (company or "Unknown"))[:31],
                index=False)
            saved.append("Order - " + (company or "Unknown"))
    return saved


def reset_cycle() -> str:
    """Archive the current batch, then clear the working Wishlist + options.

    Returns the timestamp label used for the archive.
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    wl = load_wishlist()
    opts = load_options()

    if _use_gsheets():
        sh = _open_spreadsheet()

        def _archive(title, headers, df):
            try:
                ws = sh.worksheet(title)
                if ws.row_values(1) != headers:
                    ws.clear()
                    ws.append_row(headers)
            except Exception:
                ws = sh.add_worksheet(title, rows=200, cols=len(headers))
                ws.append_row(headers)
            rows = [[str(x) for x in r] + [ts] for r in df.values.tolist()]
            if rows:
                ws.append_rows(rows, value_input_option="USER_ENTERED")

        _archive("Archive Wishlist", WL_COLUMNS + ["Archived"], wl)
        _archive("Archive Options", OPT_COLUMNS + ["Archived"], opts)
        wl_ws, opt_ws = _get_worksheets()
        wl_ws.clear(); wl_ws.append_row(WL_COLUMNS)
        opt_ws.clear(); opt_ws.append_row(OPT_COLUMNS)
        _refresh()
        return ts

    # Local Excel: append this batch to archive.xlsx, then clear the workbook
    arch = EXCEL_FILE.parent / "archive.xlsx"

    def _read(sheet, cols):
        if arch.exists():
            try:
                return pd.read_excel(arch, sheet_name=sheet)
            except Exception:
                pass
        return pd.DataFrame(columns=cols + ["Archived"])

    awl = _read("Wishlist", WL_COLUMNS)
    aopt = _read("Options", OPT_COLUMNS)
    wl2 = wl.copy(); wl2["Archived"] = ts
    opt2 = opts.copy(); opt2["Archived"] = ts
    with pd.ExcelWriter(arch, engine="openpyxl") as writer:
        pd.concat([awl, wl2], ignore_index=True).to_excel(
            writer, sheet_name="Wishlist", index=False)
        pd.concat([aopt, opt2], ignore_index=True).to_excel(
            writer, sheet_name="Options", index=False)
    write_workbook(pd.DataFrame(columns=WL_COLUMNS),
                   pd.DataFrame(columns=OPT_COLUMNS))
    return ts


def route_for(supplier, cart) -> str:
    """Decide the procurement bucket for a sourced component.

    Only two routes (per the lecturer's brief), based purely on user-entered
    data — no price comparison:
      * Shopping Cart Available = Yes -> Shopping Cart Queue
      * Shopping Cart Available = No  -> Manual Order List (grouped by supplier)
    """
    supplier = clean(supplier)
    cart = clean(cart)
    if not supplier:
        return "Unresolved"
    return "Shopping Cart Queue" if cart.lower() == "yes" else "Manual Order List"


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

        /* Yellow reminder card (Procurement Outputs) */
        [class*="st-key-reminder_card"] {{
            background: #FEFBF2 !important;
            border: 1px solid #EAD9A6 !important;
            border-radius: 16px !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05) !important;
            padding: 18px 20px !important;
            margin-bottom: 16px !important;
        }}

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

        /* Blue "Edit" button (Wishlist rows) */
        [class*="st-key-editcell"] button {{
            background: {INPUT_BG} !important;
            border: 1px solid #C9D8EE !important;
            border-radius: 10px !important;
            box-shadow: none !important;
            min-height: 38px !important;
            padding: 4px 8px !important;
            white-space: nowrap !important;
        }}
        [class*="st-key-editcell"] button p {{
            color: {PRIMARY_BLUE} !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            white-space: nowrap !important;
        }}
        [class*="st-key-editcell"] button:hover {{
            background: {SIDEBAR_ACCENT} !important;
            border-color: {PRIMARY_BLUE} !important;
        }}

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

    # ---- Bulk add (fast entry of many components at once) ---------------
    with st.expander("➕  Add many at once (paste a list)"):
        st.caption("One component per line: **Component, Model, Specification, "
                   "Quantity** (only Component is required).")
        with st.form("bulk_add_form", clear_on_submit=True):
            bulk = st.text_area(
                "Paste components", height=140,
                placeholder=("LM358 Op-Amp, LM358N, Dual 3-32V DIP-8, 50\n"
                             "10k Resistor, CF1/4W, 0.25W 5%, 200\n"
                             "Arduino Mega, 2560 R3, 54 I/O, 10"))
            if st.form_submit_button("Add all"):
                items = []
                for line in bulk.splitlines():
                    if not line.strip():
                        continue
                    parts = [p.strip() for p in line.split(",")]
                    comp = parts[0]
                    if not comp:
                        continue
                    model = parts[1] if len(parts) > 1 else ""
                    spec = parts[2] if len(parts) > 2 else ""
                    try:
                        qty = int(parts[3]) if len(parts) > 3 and parts[3] else 1
                    except ValueError:
                        qty = 1
                    items.append((comp, model, spec, qty))
                if items:
                    n_added = add_components_bulk(items)
                    st.success(f"Added {n_added} component(s).")
                    st.rerun()
                else:
                    st.warning("Nothing to add — paste at least one line.")

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

        # Inline edit form (to fill in / correct missing data)
        editing = st.session_state.get("wl_edit_id")
        if editing is not None:
            match = df[df["#"].astype(str) == str(editing)]
            if match.empty:
                st.session_state.pop("wl_edit_id", None)
            else:
                er = match.iloc[0]
                with st.form("edit_component_form"):
                    st.markdown(
                        f'<div class="src-card-title">Edit {clean(er["Component"])} '
                        f'<span class="opt-count">({cmp_id(editing)})</span></div>',
                        unsafe_allow_html=True)
                    e1, e2 = st.columns(2, gap="medium")
                    with e1:
                        e_comp = st.text_input("Component Name *",
                                               value=clean(er["Component"]))
                        e_spec = st.text_input("Specification",
                                               value=clean(er["Specifications"]))
                    with e2:
                        e_model = st.text_input("Model Number",
                                                value=clean(er["Model"]))
                        e_qty = st.number_input(
                            "Quantity *", min_value=1, max_value=100000,
                            value=max(1, _qty(er["Quantity"])), step=1)
                    s1, s2, _ = st.columns([1, 1, 3])
                    with s1:
                        if st.form_submit_button("Save changes"):
                            if not e_comp.strip():
                                st.warning("Component name can't be empty.")
                            else:
                                update_component(editing, e_comp.strip(),
                                                 e_model.strip(), e_spec.strip(), e_qty)
                                st.session_state.pop("wl_edit_id", None)
                                st.session_state.wl_msg = \
                                    f"“{e_comp.strip()}” updated."
                                st.rerun()
                    with s2:
                        if st.form_submit_button("Cancel"):
                            st.session_state.pop("wl_edit_id", None)
                            st.rerun()

        if count == 0:
            st.markdown(
                '<div class="wl-empty">No components yet — add your first one '
                'above.</div>',
                unsafe_allow_html=True,
            )
        else:
            widths = [1.9, 1.0, 1.8, 0.6, 1.25, 1.0, 1.2, 1.2]
            heads = ["COMPONENT", "MODEL", "SPECIFICATION", "QTY",
                     "DATE ADDED", "STATUS", "", ""]
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
                        econt = st.container(key=f"editcell_{r['#']}")
                    except TypeError:
                        econt = st.container()
                    with econt:
                        if st.button("Edit", key=f"wl_edit_{r['#']}",
                                     use_container_width=True):
                            st.session_state.wl_edit_id = r["#"]
                            st.session_state.pop("wl_pending_delete", None)
                            st.rerun()
                with row[7]:
                    try:
                        cont = st.container(key=f"delcell_{r['#']}")
                        need_marker = False
                    except TypeError:
                        cont = st.container()
                        need_marker = True
                    with cont:
                        # marker only needed as the old-Streamlit CSS fallback;
                        # skipping it on modern Streamlit keeps Delete aligned
                        # with Edit (no extra element pushing the button down)
                        if need_marker:
                            st.markdown('<span class="del-marker"></span>',
                                        unsafe_allow_html=True)
                        if st.button("Delete", key=f"wl_del_{r['#']}",
                                     use_container_width=True):
                            st.session_state.wl_pending_delete = r["#"]
                            st.rerun()

    # ---- Download a snapshot of the LIVE data (Google Sheets or Excel) --
    if count:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            load_wishlist().to_excel(writer, sheet_name="Wishlist", index=False)
            load_options().to_excel(writer, sheet_name="SupplierOptions", index=False)
        st.download_button(
            "⬇  Download a snapshot (.xlsx)",
            data=buf.getvalue(),
            file_name="wishlist.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="A one-off copy of the current data. The live data lives in "
                 "Google Sheets when shared storage is on.",
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
    show_sourced = bool(st.session_state.get("show_sourced", False))

    def _is_sourced(i):
        return clean(wl.iloc[i]["Status"]).lower() == "sourced"

    pending_idx = [i for i in range(n) if not _is_sourced(i)]
    # Queue lists components still to source. Sourced ones drop off (done),
    # unless the user ticks "Show sourced", or everything is already sourced.
    visible_idx = list(range(n)) if (show_sourced or not pending_idx) else pending_idx

    idx = st.session_state.get("sourcing_idx", 0)
    if idx not in visible_idx:
        idx = visible_idx[0] if visible_idx else 0
        st.session_state.sourcing_idx = idx

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
            st.toggle("Show sourced", key="show_sourced")
            if not pending_idx:
                st.markdown(
                    '<div style="color:#3E7A63;font-weight:600;font-size:13px;'
                    'margin:6px 0;">✓ All sourced — head to Procurement Outputs.</div>',
                    unsafe_allow_html=True,
                )
            for i in visible_idx:
                r = wl.iloc[i]
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
                    nxt = [j for j in pending_idx if j != idx]
                    st.session_state.sourcing_idx = nxt[0] if nxt else idx
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
                    # jump to the next still-pending component
                    nxt = [j for j in pending_idx if j != idx]
                    st.session_state.sourcing_idx = nxt[0] if nxt else idx
                    st.rerun()

        # Single sourcing-entry card (choose type; URL always available)
        with card():
            st.markdown(
                '<div class="src-card-title">Add Supplier</div>',
                unsafe_allow_html=True,
            )
            s_type = st.radio(
                "Source type",
                ["Trusted supplier", "Found via search"],
                horizontal=True, key="s_type",
            )
            supplier_type = "Trusted" if s_type == "Trusted supplier" else "External"
            supplier = st.selectbox("Supplier", SUPPLIERS, key="s_supplier")
            if supplier == "Other":
                supplier = st.text_input(
                    "Supplier name", placeholder="Type the supplier / vendor name",
                    key="s_other").strip()

            s_url = st.text_input("Product URL (optional)",
                                  placeholder="https://www.digikey.com/...",
                                  key="s_url")
            f1, f2, f3 = st.columns(3)
            with f1:
                s_stock = st.text_input("Stock", placeholder="1,240", key="s_stock")
            with f2:
                s_eta = st.text_input("ETA", placeholder="3–5 days", key="s_eta")
            with f3:
                s_price = st.text_input("Unit Price", placeholder="R 4.20",
                                        key="s_price")
            s_cart = st.toggle("Shopping Cart Available", value=True, key="s_cart")
            if st.button("+  Add option", key="s_add", use_container_width=True):
                if not clean(supplier):
                    st.warning("Choose a supplier or enter a vendor name.")
                else:
                    add_option(comp_id, supplier, supplier_type, s_price, s_stock,
                               s_eta, "Yes" if s_cart else "No", s_url)
                    st.rerun()

        # Supplier options table
        with card():
            st.markdown(
                f'<div class="opt-head">'
                f'<span class="card-heading" style="margin:0;">Recorded options for '
                f'<span style="color:{PRIMARY_BLUE};">{clean(comp["Component"])}</span></span>'
                f'<span class="opt-count">{len(comp_opts)} option'
                f'{"s" if len(comp_opts) != 1 else ""}</span></div>'
                f'<div class="opt-sub">Select the one you\'ll use — add another as a '
                f'backup for no-stock / issues. &nbsp;·&nbsp; {cmp_id(comp["#"])} '
                f'&nbsp;·&nbsp; {clean(comp["Model"])}</div>',
                unsafe_allow_html=True,
            )

            if comp_opts.empty:
                st.markdown(
                    '<div class="wl-empty">No options recorded yet — add one '
                    'using the card above.</div>',
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


def _price(s) -> float:
    m = re.search(r"[-+]?\d*\.?\d+", clean(s).replace(",", ""))
    return float(m.group()) if m else 0.0


def _qty(x) -> int:
    try:
        return int(float(clean(x)))
    except (ValueError, TypeError):
        return 0


def _money(v) -> str:
    return f"R {v:,.2f}"


def _group_by_company(items):
    """Group sourced items by supplier company, aggregating identical
    component+specification rows (summing quantity, keeping a URL).

    items: list of (comp_row, supplier, chosen_option). Returns
    {company: [ {Component, Model, Spec, Qty, Price, Total, URL} ]}.
    """
    by = {}
    for (comp, sup, chosen) in items:
        agg = by.setdefault(sup or "(no supplier)", {})
        k = (clean(comp["Component"]).strip().lower(),
             clean(comp["Specifications"]).strip().lower())
        if k not in agg:
            agg[k] = {"Component": clean(comp["Component"]),
                      "Model": clean(comp["Model"]),
                      "Spec": clean(comp["Specifications"]),
                      "Qty": 0, "Price": _price(chosen["Price"]),
                      "URL": clean(chosen["URL"])}
        agg[k]["Qty"] += _qty(comp["Quantity"])
        if not agg[k]["URL"] and clean(chosen["URL"]):
            agg[k]["URL"] = clean(chosen["URL"])
    out = {}
    for sup, agg in by.items():
        rows = []
        for a in agg.values():
            a["Total"] = a["Price"] * a["Qty"]
            rows.append(a)
        out[sup] = rows
    return out


def _link_cell(url):
    url = clean(url)
    if not url:
        return '<td class="cmp-c cmp-muted">—</td>'
    return (f'<td class="cmp-c"><a href="{url}" target="_blank" '
            f'style="color:#5B7DB1;font-weight:600;">🔗 open</a></td>')


def _company_tables_html(by_co):
    """Render one sub-card table per company (with a URL/link column)."""
    html = ""
    for sup, rows in by_co.items():
        sub_total = sum(a["Total"] for a in rows)
        trows = "".join(
            "<tr>"
            f'<td class="cmp-c cmp-strong">{a["Component"]}</td>'
            f'<td class="cmp-c cmp-muted">{a["Model"]}</td>'
            f'<td class="cmp-c cmp-muted">{a["Spec"]}</td>'
            f'<td class="cmp-c">{a["Qty"]}</td>'
            f'<td class="cmp-c cmp-muted">{_money(a["Price"])}</td>'
            f'<td class="cmp-c cmp-price">{_money(a["Total"])}</td>'
            + _link_cell(a["URL"]) + "</tr>"
            for a in rows
        )
        html += (
            f'<div class="card"><div class="card-heading" style="margin-bottom:12px;">'
            f'{sup} <span class="opt-count">{len(rows)} line'
            f'{"s" if len(rows) != 1 else ""} · {_money(sub_total)}</span></div>'
            + _table(["COMPONENT", "MODEL", "SPECIFICATION", "QTY", "UNIT PRICE",
                      "TOTAL", "LINK"], trows) + '</div>')
    return html


def render_procurement():
    st.markdown(
        '<div class="page-title">Procurement Outputs</div>'
        '<div class="page-subtitle">Components are automatically routed based on '
        'Shopping Cart availability.</div>',
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
        try:
            rc = st.container(key="reminder_card")
        except TypeError:
            rc = st.container()
        with rc:
            st.markdown(
                f'<div class="card-heading" style="color:{GOLD_TEXT};margin-bottom:6px;">'
                f'⚠ Reminder — {len(incomplete)} item'
                f'{"s" if len(incomplete) != 1 else ""} still have missing data</div>'
                f'<div class="page-subtitle" style="margin-bottom:14px;">This is only a '
                f'reminder — fix it here, or continue and use the outputs below.</div>'
                + _table(["COMPONENT", "ID", "MISSING FIELDS"], rows),
                unsafe_allow_html=True,
            )
            fx1, fx2, _ = st.columns([1.5, 1.5, 2])
            with fx1:
                if st.button("✏  Edit on Wishlist", key="fix_wishlist",
                             use_container_width=True):
                    st.session_state.active_page = "Wishlist"
                    st.rerun()
            with fx2:
                if st.button("🔍  Fix on Sourcing", key="fix_sourcing",
                             use_container_width=True):
                    st.session_state.active_page = "Sourcing"
                    st.rerun()
    else:
        st.markdown(
            '<div class="card"><div style="color:#3E7A63;font-weight:600;">'
            '✓ All components have complete data.</div></div>',
            unsafe_allow_html=True,
        )

    # ---- Split sourced components: Cart=Yes vs Cart=No -----------------
    cart_items, manual_items, unresolved = [], [], []
    for _, comp in wl.iterrows():
        chosen = _chosen_option(opts, comp["#"])
        if clean(comp["Status"]).lower() == "sourced" and chosen is not None \
                and clean(chosen["Supplier"]):
            supplier = clean(chosen["Supplier"])
            if clean(chosen["Shopping Cart Available"]).lower() == "yes":
                cart_items.append((comp, supplier, chosen))
            else:
                manual_items.append((comp, supplier, chosen))
        else:
            unresolved.append(comp)

    cart_total = sum(_price(ch["Price"]) * _qty(c["Quantity"])
                     for (c, s, ch) in cart_items)
    manual_total = sum(_price(ch["Price"]) * _qty(c["Quantity"])
                       for (c, s, ch) in manual_items)

    # ---- Summary cards -------------------------------------------------
    st.markdown(
        f'''
        <div style="display:flex;gap:16px;margin-bottom:14px;">
          <div class="card" style="flex:1;margin-bottom:0;border:1.5px solid #C9D8EE;">
            <div style="display:flex;justify-content:space-between;align-items:flex-start;">
              <div style="font-size:20px;">🛒</div>
              <div style="font-size:24px;font-weight:700;color:{TEXT};">{len(cart_items)}</div>
            </div>
            <div style="font-size:15px;font-weight:700;color:{TEXT};margin-top:6px;">Shopping Cart Queue</div>
            <div style="font-size:13px;color:{MUTED_TEXT};">Cart available — send via email</div>
            <div style="font-size:13px;color:{PRIMARY_BLUE};font-weight:700;margin-top:6px;">{_money(cart_total)} total</div>
          </div>
          <div class="card" style="flex:1;margin-bottom:0;">
            <div style="display:flex;justify-content:space-between;align-items:flex-start;">
              <div style="font-size:20px;">📄</div>
              <div style="font-size:24px;font-weight:700;color:{TEXT};">{len(manual_items)}</div>
            </div>
            <div style="font-size:15px;font-weight:700;color:{TEXT};margin-top:6px;">Manual Order List</div>
            <div style="font-size:13px;color:{MUTED_TEXT};">No cart — grouped by supplier</div>
            <div style="font-size:13px;color:{GOLD_TEXT};font-weight:700;margin-top:6px;">{_money(manual_total)} total</div>
          </div>
        </div>
        <div class="card" style="padding:9px 16px;margin-bottom:14px;font-size:13px;color:{MUTED_TEXT};">
          Routing logic: &nbsp;<b style="color:{PRIMARY_BLUE};">●</b> Cart Available = Yes → Shopping Cart Queue
          &nbsp;&nbsp;<b style="color:{GOLD_TEXT};">●</b> Cart Available = No → Manual Order List
        </div>
        ''',
        unsafe_allow_html=True,
    )

    # Group both buckets by company (aggregated, with URL)
    cart_by_co = _group_by_company(cart_items)
    manual_by_co = _group_by_company(manual_items)
    orders_by_company = manual_by_co   # used by Save/email below

    st.session_state.setdefault("lect_email", LECTURER_EMAIL)
    lect = st.text_input("Lecturer email", key="lect_email")

    def _email_body(title, by_co):
        lines = []
        for company, rows in by_co.items():
            lines.append(f"== {company} ==")
            for a in rows:
                u = f"  {a['URL']}" if a["URL"] else ""
                lines.append(f'- {a["Component"]} ({a["Model"]}) x{a["Qty"]} '
                             f'@ {_money(a["Price"])}{u}')
            lines.append("")
        return f"{title}\n\n" + "\n".join(lines)

    # ---- Shopping Cart Queue (grouped by company, with URL) ------------
    cart_mailto = (
        f'mailto:{quote(lect or "")}?subject={quote("Procurement shopping cart")}'
        f'&body={quote(_email_body("Shopping cart to build online:", cart_by_co))}')
    cart_email_btn = (
        f'<a href="{cart_mailto}" target="_blank" style="background:{PRIMARY_BLUE};'
        f'color:#fff;padding:8px 16px;border-radius:10px;font-size:13px;'
        f'font-weight:600;text-decoration:none;white-space:nowrap;">'
        f'✉ Email Cart to Lecturer</a>') if cart_by_co else ""
    st.markdown(
        f'<div style="display:flex;justify-content:space-between;align-items:center;'
        f'margin:6px 0 12px 0;">'
        f'<div class="card-heading" style="margin:0;">🛒 Shopping Cart Queue '
        f'<span class="opt-count">grouped by company · click 🔗 to open</span></div>'
        f'{cart_email_btn}</div>',
        unsafe_allow_html=True,
    )
    if cart_by_co:
        st.markdown(_company_tables_html(cart_by_co), unsafe_allow_html=True)
    else:
        st.markdown('<div class="card"><div class="wl-empty">Nothing here yet.</div>'
                    '</div>', unsafe_allow_html=True)

    # ---- Manual Order List (grouped by company, with URL) --------------
    st.markdown('<div class="card-heading" style="margin:6px 0 12px 0;">'
                '📄 Manual Order List <span class="opt-count">grouped by company</span>'
                '</div>', unsafe_allow_html=True)
    if manual_by_co:
        st.markdown(_company_tables_html(manual_by_co), unsafe_allow_html=True)
    else:
        st.markdown('<div class="card"><div class="wl-empty">Nothing here yet — no '
                    'sourced components without an online cart.</div></div>',
                    unsafe_allow_html=True)

    # ---- Save per-company order sheets + email the lecturer ------------
    if orders_by_company:
        if st.session_state.get("orders_saved_msg"):
            st.success(st.session_state.pop("orders_saved_msg"))
        sc1, sc2, _ = st.columns([1.4, 1.6, 2])
        with sc1:
            if st.button("💾  Save order sheets", key="save_orders",
                         type="primary", use_container_width=True):
                saved = save_company_order_sheets(orders_by_company)
                where = "Google Sheet" if _use_gsheets() else "manual_orders.xlsx"
                st.session_state.orders_saved_msg = (
                    f"Saved {len(saved)} company order sheet(s) to the {where}: "
                    + ", ".join(saved))
                st.rerun()
        with sc2:
            mbody = _email_body("Manual order lists (grouped by supplier):",
                                orders_by_company)
            mmailto = (f'mailto:{quote(lect or "")}?subject='
                       f'{quote("Manual procurement orders")}&body={quote(mbody)}')
            st.markdown(
                f'<a href="{mmailto}" target="_blank" style="display:block;'
                f'text-align:center;background:{PRIMARY_BLUE};color:#fff;padding:8px 16px;'
                f'border-radius:10px;font-size:14px;font-weight:600;'
                f'text-decoration:none;">✉ Email Manual Orders to Lecturer</a>',
                unsafe_allow_html=True,
            )

    # ---- Unresolved (not sourced) --------------------------------------
    if unresolved:
        rows = "".join(
            "<tr>"
            f'<td class="cmp-c cmp-strong">{clean(c["Component"])}</td>'
            f'<td class="cmp-c cmp-muted">{cmp_id(c["#"])}</td>'
            f'<td class="cmp-c cmp-muted">Not sourced yet</td>'
            "</tr>"
            for c in unresolved
        )
        st.markdown(
            f'<div class="card"><div class="card-heading">❓ Unresolved '
            f'<span class="opt-count">({len(unresolved)})</span></div>'
            + _table(["COMPONENT", "ID", "REASON"], rows) + '</div>',
            unsafe_allow_html=True,
        )

    # ---- Export to Excel: one sheet per company (cart + manual) --------
    def _co_df(rows):
        return pd.DataFrame([{
            "Component": a["Component"], "Model": a["Model"],
            "Specification": a["Spec"], "Quantity": a["Qty"],
            "Unit Price": _money(a["Price"]), "Line Total": _money(a["Total"]),
            "URL": a["URL"],
        } for a in rows])

    used = set()

    def _sheet_name(prefix, company):
        base = f"{prefix} - {company}"[:31]
        name, k = base, 2
        while name in used:
            name = f"{base[:28]}~{k}"
            k += 1
        used.add(name)
        return name

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        wrote = False
        for company, rows in cart_by_co.items():
            _co_df(rows).to_excel(writer, index=False,
                                  sheet_name=_sheet_name("Cart", company))
            wrote = True
        for company, rows in manual_by_co.items():
            _co_df(rows).to_excel(writer, index=False,
                                  sheet_name=_sheet_name("Manual", company))
            wrote = True
        if not wrote:
            pd.DataFrame(columns=["Component", "Model", "Specification",
                                  "Quantity", "Unit Price", "Line Total", "URL"]
                         ).to_excel(writer, index=False, sheet_name="Procurement")
    st.download_button("⬇  Export to Excel (per company)", data=buf.getvalue(),
                       file_name="procurement_outputs.xlsx",
                       mime="application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet")

    # ---- Start a new procurement cycle --------------------------------
    st.markdown('<div class="card-heading" style="margin-top:16px;">'
                '🔄 Start a new cycle</div>', unsafe_allow_html=True)
    if st.session_state.get("cycle_reset_msg"):
        st.success(st.session_state.pop("cycle_reset_msg"))
    if not st.session_state.get("confirm_reset"):
        st.markdown(
            '<div class="opt-sub">Finished this batch (sourced, emailed, orders '
            'saved)? This archives everything, then clears the Wishlist &amp; '
            'Sourcing so you can start fresh.</div>',
            unsafe_allow_html=True,
        )
        if st.button("Start new cycle…", key="start_reset"):
            st.session_state.confirm_reset = True
            st.rerun()
    else:
        st.markdown(
            '<div class="route-hint" style="border-color:#C94B4B;'
            'background:#F7E0E0;color:#C94B4B;">This will <b>archive</b> the current '
            'batch and then <b>clear</b> the Wishlist and all sourcing. Continue?</div>',
            unsafe_allow_html=True,
        )
        rc1, rc2, _ = st.columns([1.5, 1, 3])
        with rc1:
            if st.button("Yes, archive & clear", key="do_reset", type="primary",
                         use_container_width=True):
                ts = reset_cycle()
                st.session_state.confirm_reset = False
                st.session_state.sourcing_idx = 0
                st.session_state.cycle_reset_msg = (
                    f"New cycle started — previous batch archived ({ts}).")
                st.rerun()
        with rc2:
            if st.button("Cancel", key="cancel_reset", use_container_width=True):
                st.session_state.confirm_reset = False
                st.rerun()


# ----------------------------------------------------------------------------
# Dashboard page
# ----------------------------------------------------------------------------
def render_dashboard():
    wl = load_wishlist()
    year = datetime.now().year
    month = datetime.now().strftime("%B %Y")

    components = len(wl)
    models = len({clean(m).lower() for _, m in wl["Model"].items() if clean(m)}) \
        if components else 0
    total_qty = sum(_qty(q) for q in wl["Quantity"]) if components else 0
    sourced = int(sum(clean(s).lower() == "sourced" for s in wl["Status"])) \
        if components else 0
    pending = components - sourced

    st.markdown(
        f'<div class="page-title">Procurement Dashboard</div>'
        f'<div class="page-subtitle">School of Electrical &amp; Information '
        f'Engineering — {month}</div>',
        unsafe_allow_html=True,
    )

    # ---- Four metric cards --------------------------------------------
    metrics = [("⚙", components, "Components"), ("▤", models, "Models"),
               ("#", total_qty, "Total Quantity"), ("⏱", pending, "Pending Items")]
    cards = ""
    for icon, value, label in metrics:
        cards += (
            f'<div class="card" style="flex:1;margin-bottom:0;">'
            f'<div style="display:flex;justify-content:space-between;align-items:center;">'
            f'<div style="width:34px;height:34px;border-radius:9px;background:{INPUT_BG};'
            f'display:flex;align-items:center;justify-content:center;font-size:16px;">{icon}</div>'
            f'<span style="font-size:12px;color:{MUTED_TEXT};">{year}</span></div>'
            f'<div style="font-size:30px;font-weight:800;color:{TEXT};margin-top:8px;">{value}</div>'
            f'<div style="font-size:13px;color:{MUTED_TEXT};">{label}</div></div>')
    st.markdown(f'<div style="display:flex;gap:14px;margin-bottom:16px;">{cards}</div>',
                unsafe_allow_html=True)

    # ---- Workflow progress --------------------------------------------
    steps = [("Wishlist", "Components identified"),
             ("Sourcing", "Suppliers assigned"),
             ("Review", "Quotes compared"),
             ("Procurement", "Orders placed"),
             ("Received", "Delivery confirmed")]
    if components == 0:
        current = 1
    elif sourced < components:
        current = 2
    else:
        current = 3

    steps_html = '<div style="display:flex;align-items:flex-start;">'
    for i, (name, sub) in enumerate(steps):
        num = i + 1
        if num < current:
            circle = (f'<div style="width:38px;height:38px;border-radius:50%;'
                      f'background:{PRIMARY_BLUE};color:#fff;display:flex;'
                      f'align-items:center;justify-content:center;font-weight:700;">✓</div>')
            nm_col = TEXT
        elif num == current:
            circle = (f'<div style="width:38px;height:38px;border-radius:50%;'
                      f'background:#fff;color:{PRIMARY_BLUE};border:2px solid {PRIMARY_BLUE};'
                      f'display:flex;align-items:center;justify-content:center;'
                      f'font-weight:700;">{num}</div>')
            nm_col = TEXT
        else:
            circle = (f'<div style="width:38px;height:38px;border-radius:50%;'
                      f'background:#fff;color:{MUTED_TEXT};border:2px solid {BORDER};'
                      f'display:flex;align-items:center;justify-content:center;'
                      f'font-weight:700;">{num}</div>')
            nm_col = MUTED_TEXT
        steps_html += (
            f'<div style="flex:1;text-align:center;">'
            f'<div style="display:flex;justify-content:center;">{circle}</div>'
            f'<div style="font-weight:700;font-size:14px;color:{nm_col};margin-top:8px;">{name}</div>'
            f'<div style="font-size:11px;color:{MUTED_TEXT};">{sub}</div></div>')
        if i < len(steps) - 1:
            line_col = PRIMARY_BLUE if num < current else BORDER
            steps_html += (f'<div style="flex:0.5;height:2px;background:{line_col};'
                           f'margin-top:18px;"></div>')
    steps_html += '</div>'

    st.markdown(
        f'<div class="card">'
        f'<div style="display:flex;justify-content:space-between;align-items:center;'
        f'margin-bottom:20px;">'
        f'<div class="card-heading" style="margin:0;">Workflow Progress</div>'
        f'<span style="background:{INPUT_BG};color:{PRIMARY_BLUE};font-size:12px;'
        f'font-weight:600;padding:4px 12px;border-radius:999px;">Active Cycle</span></div>'
        + steps_html + '</div>',
        unsafe_allow_html=True,
    )

    # ---- Recent components --------------------------------------------
    if components:
        rows = "".join(
            "<tr>"
            f'<td class="cmp-c cmp-strong">{clean(r["Component"])}</td>'
            f'<td class="cmp-c cmp-muted">{clean(r["Model"])}</td>'
            f'<td class="cmp-c">{clean(r["Quantity"])}</td>'
            f'<td class="cmp-c">{status_badge(r["Status"])}</td>'
            "</tr>"
            for _, r in wl.tail(5).iloc[::-1].iterrows()
        )
    else:
        rows = ('<tr><td colspan="4"><div class="wl-empty">No components yet — '
                'add some on the Wishlist page.</div></td></tr>')
    st.markdown(
        '<div class="card"><div class="card-heading">Recent Components</div>'
        + _table(["COMPONENT", "MODEL", "QUANTITY", "STATUS"], rows) + '</div>',
        unsafe_allow_html=True,
    )


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
        f'</div>',
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
    render_dashboard()
