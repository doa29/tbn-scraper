# ─────────────────────────────────────────────────────────────────────────────
# Streamlit TBN Scraper App  —  UPDATED
#
# Fixes added:
# 1) Robust login: submit the <form> element so hidden CSRF tokens are sent.
# 2) Manual-login mode now copies session cookies into a headless driver.
# 3) Hardened datepicker selectors and waits.
# 4) Better diagnostics: cookie names + HTML snippet on failure.
#
# Features:
# - Secure credential entry (no plaintext storage) for TBN portal
# - Supports single years, ranges (e.g., 2020-2024), and comma lists (e.g., 2023,2025)
# - Optional upload of prior scraped JSON/CSV for YoY comparisons
# - Headless Selenium scrape, robust error handling & progress UI
# - Generates Excel with ADA notes & weekend shading
# - Download results in-app and (optionally) email via Outlook or SMTP
#
# Dependencies (pip):
#   streamlit pandas beautifulsoup4 selenium python-dotenv xlsxwriter lxml html5lib openpyxl
#   (Optional) pywin32 (Outlook), yagmail (or use stdlib smtplib)
#
# Respect website terms. Scrape only content you are authorized to access.
# ─────────────────────────────────────────────────────────────────────────────

import os
import re
import time
import json
import tempfile
from datetime import datetime
from io import StringIO
from typing import List, Dict, Optional, Tuple

import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup

# Selenium imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

# Excel writer
import xlsxwriter  # noqa: F401

# Outlook optional
try:
    import win32com.client as win32  # type: ignore
    HAS_OUTLOOK = True
except Exception:
    HAS_OUTLOOK = False

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
LOGIN_URL  = 'https://portal.thebusnetwork.com/login'
REPORT_URL = 'https://portal.thebusnetwork.com/salesman/reports/daily-usage'

def EXCEL_PATH(tmpdir: str, y: int) -> str:
    return os.path.join(tmpdir, f"TBN_Report_{y}.xlsx")

def RAW_JSON_PATH(tmpdir: str, y: int) -> str:
    return os.path.join(tmpdir, f"TBN_RawData_{y}.json")

# ─────────────────────────────────────────────────────────────────────────────
# Utilities
# ─────────────────────────────────────────────────────────────────────────────
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def parse_years_input(text: str) -> List[int]:
    years: List[int] = []
    for part in [p.strip() for p in text.split(",") if p.strip()]:
        if "-" in part:
            a, b = [x.strip() for x in part.split("-", 1)]
            if a.isdigit() and b.isdigit():
                lo, hi = int(a), int(b)
                if lo <= hi:
                    years.extend(list(range(lo, hi + 1)))
        else:
            if part.isdigit():
                years.append(int(part))
    years = [y for y in years if 2000 <= y <= 2100]
    return sorted(list(set(years)))

def validate_email_list(raw: str) -> List[str]:
    emails = [e.strip() for e in re.split(r"[;,]", raw) if e.strip()]
    bad = [e for e in emails if not EMAIL_RE.match(e)]
    if bad:
        raise ValueError(f"Invalid email(s): {', '.join(bad)}")
    return emails

# ─────────────────────────────────────────────────────────────────────────────
# Selenium setup & auth
# ─────────────────────────────────────────────────────────────────────────────
def build_chrome(manual_visible: bool = False) -> webdriver.Chrome:
    opts = ChromeOptions()
    if not manual_visible:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-blink-features=AutomationControlled")

    chrome_path = os.environ.get("CHROME_PATH")
    if chrome_path and os.path.exists(chrome_path):
        opts.binary_location = chrome_path

    try:
        drv = webdriver.Chrome(options=opts)
        drv.set_page_load_timeout(45)
        return drv
    except WebDriverException as e:
        raise RuntimeError(f"Failed to start Chrome. Details: {e}") from e

def copy_cookies(src_drv: webdriver.Chrome, dst_drv: webdriver.Chrome, domain_hint: str):
    """Copy cookies from a logged-in visible session to headless."""
    cookies = src_drv.get_cookies()
    # Must be on same domain to add cookies
    dst_drv.get(domain_hint)
    for c in cookies:
        c.pop("sameSite", None)  # selenium may reject unknown fields
        try:
            dst_drv.add_cookie(c)
        except Exception:
            pass

def login_and_get_driver(username: str, password: str, log_area=None) -> webdriver.Chrome:
    """Headless automated login with robust form submit & diagnostics."""
    drv = build_chrome(manual_visible=False)
    drv.get(LOGIN_URL)
    wait = WebDriverWait(drv, 30)

    # Enter iframe if present
    frames = drv.find_elements(By.TAG_NAME, "iframe")
    in_iframe = False
    if frames:
        try:
            drv.switch_to.frame(frames[0])
            in_iframe = True
            if log_area: log_area.info("Switched into login iframe.")
        except Exception:
            in_iframe = False

    def find_el(css):
        return wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css)))

    try:
        # Try to locate a form node so hidden inputs (CSRF) go with submission
        try:
            form = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "form")))
        except Exception:
            form = None

        # Fill fields
        email = find_el("input[name='email'], input#email, input#username, input[type='email']")
        email.clear(); email.send_keys(username)

        pwd = find_el("input[name='password'], input#password, input[type='password']")
        pwd.clear(); pwd.send_keys(password)

        # Submit the form element first; fall back to clickable button; then Enter
        if form:
            drv.execute_script("arguments[0].submit();", form)
        else:
            try:
                btn = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, "button[type='submit'], input[type='submit'], button"
                )))
                drv.execute_script("arguments[0].click();", btn)
            except Exception:
                pwd.send_keys(Keys.ENTER)

        if in_iframe:
            drv.switch_to.default_content()

        # Wait for either success or a bounce to login
        WebDriverWait(drv, 15).until(EC.any_of(
            EC.url_contains("/login"),
            EC.url_contains("/salesman"),
            EC.presence_of_element_located((By.CSS_SELECTOR, "table, .table, div[data-report], #report"))
        ))

        # Try to visit report page; if not authenticated, most apps redirect to /login
        drv.get(REPORT_URL)
        time.sleep(2)

        current = drv.current_url.lower()
        authed = ("login" not in current) and bool(
            drv.find_elements(By.CSS_SELECTOR, "table, .table, div[data-report], #report")
        )
        if not authed:
            # Diagnostics
            snippet = (drv.page_source or "")[:1500]
            cookies = [c.get("name") for c in drv.get_cookies()]
            if log_area:
                log_area.error("Login appears to have failed; showing diagnostics.")
                log_area.code(f"URL: {drv.current_url}\nCookies: {cookies}\n\nHTML (first 1500):\n{snippet}")
                try:
                    st.image(drv.get_screenshot_as_png(), caption="Post-login state (screenshot)")
                except Exception:
                    pass
            raise RuntimeError("Login appears to have failed; redirected back to login page or no report found.")

        if log_area: log_area.success("Authenticated successfully (headless).")
        return drv

    except Exception:
        # Surface page snippet for debugging then re-raise
        try:
            snippet = drv.page_source[:1500]
            cookies = [c.get("name") for c in drv.get_cookies()]
            if log_area:
                log_area.code(f"URL: {drv.current_url}\nCookies: {cookies}\n\nHTML (first 1500):\n{snippet}")
        finally:
            drv.quit()
        raise

# ─────────────────────────────────────────────────────────────────────────────
# Scraping helpers
# ─────────────────────────────────────────────────────────────────────────────
def set_datepicker(drv: webdriver.Chrome, month: int, year: int):
    wait = WebDriverWait(drv, 20)
    selectors = [
        'div.styles_dateTimePickerInputGroup__2urdc input.datetimepicker-input',
        'input[name*="date"][type="text"]',
        'input.form-control.datetimepicker-input'
    ]
    inp = None
    for css in selectors:
        try:
            inp = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
            break
        except Exception:
            continue
    if not inp:
        raise RuntimeError("Could not find the report date input.")
    inp.clear()
    inp.send_keys(f"{month:02d}/{year}")
    inp.send_keys(Keys.TAB)
    time.sleep(1.5)

def scrape_month_data(drv: webdriver.Chrome, year: int, month: int) -> Optional[pd.DataFrame]:
    drv.get(REPORT_URL)
    time.sleep(1.5)
    set_datepicker(drv, month, year)
    time.sleep(2)
    soup = BeautifulSoup(drv.page_source, 'html.parser')
    tbl  = soup.find('table')
    if not tbl:
        return None
    df = pd.read_html(StringIO(str(tbl)))[0]
    df['Month'] = month
    df['Year']  = year
    return df

def collect_all_data_for_year(drv: webdriver.Chrome, year: int, progress=None) -> pd.DataFrame:
    all_df = []
    for m in range(1, 13):
        if progress:
            progress.progress(m/12.0, text=f"Scraping {year} – Month {m}/12")
        df = scrape_month_data(drv, year, m)
        if df is not None:
            all_df.append(df)
    if not all_df:
        return pd.DataFrame()
    return pd.concat(all_df, ignore_index=True)

# ─────────────────────────────────────────────────────────────────────────────
# Excel builder
# ─────────────────────────────────────────────────────────────────────────────
def generate_daily_totals_excel(new_data: pd.DataFrame, year: int, excel_path: str, old_raw_df: Optional[pd.DataFrame] = None):
    monthly_totals_new: Dict[int, Dict[int, int]] = {}
    monthly_ada_new: Dict[int, Dict[int, int]] = {}

    for m in range(1, 13):
        sub = new_data[new_data['Month'] == m]

        # Totals
        tot = sub[sub['Vehicle Types'].str.upper().str.contains("TOTAL", na=False)]
        if not tot.empty:
            tot = tot.iloc[0]
        else:
            days = [str(d) for d in range(1, 32) if str(d) in sub.columns]
            tot = sub[days].sum(numeric_only=True)
        monthly_totals_new[m] = {
            d: int(tot.get(str(d), 0)) if pd.notnull(tot.get(str(d), 0)) else 0
            for d in range(1, 32)
        }

        # ADA
        wc = sub[sub['Vehicle Types'].str.contains("Wheelchair", case=False, na=False)]
        if not wc.empty:
            wc = wc.iloc[0]
            monthly_ada_new[m] = {
                d: int(wc.get(str(d), 0)) if pd.notnull(wc.get(str(d), 0)) else 0
                for d in range(1, 32)
            }
        else:
            monthly_ada_new[m] = {d: 0 for d in range(1, 32)}

    # Old totals for arrows
    if old_raw_df is not None and not old_raw_df.empty:
        monthly_totals_old: Optional[Dict[int, Dict[int, int]]] = {}
        for m in range(1, 13):
            sub = old_raw_df[old_raw_df['Month'] == m]
            tot = sub[sub['Vehicle Types'].str.upper().str.contains("TOTAL", na=False)]
            if not tot.empty:
                tot = tot.iloc[0]
            else:
                days = [str(d) for d in range(1, 32) if str(d) in sub.columns]
                tot = sub[days].sum(numeric_only=True)
            monthly_totals_old[m] = {
                d: int(tot.get(str(d), 0)) if pd.notnull(tot.get(str(d), 0)) else 0
                for d in range(1, 32)
            }
    else:
        monthly_totals_old = None

    month_names = [
        "January","February","March","April","May","June",
        "July","August","September","October","November","December"
    ]
    new_df = pd.DataFrame({month_names[m-1]: pd.Series(monthly_totals_new[m]) for m in range(1, 13)})
    new_df.index.name = "Day"
    s = new_df.sum(numeric_only=True); s.name = "Monthly Totals"
    new_df = pd.concat([new_df, s.to_frame().T])
    new_df["Total"] = new_df.sum(axis=1, numeric_only=True)

    if monthly_totals_old is not None:
        old_df = pd.DataFrame({month_names[m-1]: pd.Series(monthly_totals_old[m]) for m in range(1, 13)})
        old_df.index.name = "Day"
        s2 = old_df.sum(numeric_only=True); s2.name = "Monthly Totals"
        old_df = pd.concat([old_df, s2.to_frame().T])
        old_df["Total"] = old_df.sum(axis=1, numeric_only=True)
    else:
        old_df = None

    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet("Daily Totals")
        writer.sheets["Daily Totals"] = ws

        title_fmt   = wb.add_format({'bold':True,'align':'center','valign':'vcenter','font_size':14})
        left_fmt    = wb.add_format({'align':'left','bold':True})
        hdr_fmt     = wb.add_format({'bold':True,'align':'center','border':1})
        data_fmt    = wb.add_format({'align':'center'})
        weekend_fmt = wb.add_format({'align':'center','bg_color':'#D3D3D3'})

        ncols = len(month_names) + 2
        ws.merge_range(0, 0, 0, ncols - 1, "Starr Transit Company, Inc.", title_fmt)
        ws.merge_range(1, 0, 1, ncols - 1, "Coach Requirements Calendar Based on Number of Moves Per Day", title_fmt)
        ws.write(2, 0, "Garage: Starr Garage", left_fmt)
        ws.merge_range(2, 1, 2, ncols - 1, f"January 1 {year} to December 31 {year}", title_fmt)

        headers = ["Day"] + month_names + ["Total"]
        for c, h in enumerate(headers):
            ws.write(4, c, h, hdr_fmt)

        ws.set_column(0, 0, 10)
        for c in range(1, ncols):
            ws.set_column(c, c, 15)

        first_row = 5
        for i, day in enumerate(new_df.index):
            r = first_row + i
            if day == "Monthly Totals":
                ws.write(r, 0, day, left_fmt)
            else:
                ws.write(r, 0, int(day), left_fmt)

            for m_idx, mname in enumerate(month_names, start=1):
                c = m_idx
                new_val = int(new_df.at[day, mname])
                if day == "Monthly Totals":
                    ws.write(r, c, new_val, data_fmt)
                else:
                    arrow = ""
                    if old_df is not None and day in old_df.index:
                        diff = new_val - int(old_df.at[day, mname])
                        if diff > 0:
                            arrow = f" ↑{diff}"
                        elif diff < 0:
                            arrow = f" ↓{abs(diff)}"
                    txt = f"{new_val}{arrow}"

                    ada_val = monthly_ada_new.get(m_idx, {}).get(day, 0)
                    if ada_val and int(ada_val) > 0:
                        txt += "*"
                    try:
                        wd = datetime(year, m_idx, int(day)).weekday()
                        fmt = weekend_fmt if wd >= 5 else data_fmt
                    except Exception:
                        fmt = data_fmt
                    ws.write_string(r, c, txt, fmt)
            tot_val = int(new_df.at[day, "Total"])
            ws.write(r, ncols - 1, tot_val, data_fmt)

        foot_r = first_row + len(new_df.index) + 1
        ws.write(foot_r, 0, "* = ADA (Wheelchair) job(s)", left_fmt)

# ─────────────────────────────────────────────────────────────────────────────
# Email senders
# ─────────────────────────────────────────────────────────────────────────────
def send_email_outlook(subject: str, body: str, to_emails: List[str], attachments: List[Tuple[str, bytes]]):
    if not HAS_OUTLOOK:
        raise RuntimeError("Outlook COM is not available on this system.")
    outlook = win32.Dispatch('Outlook.Application')
    mail    = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body    = body
    for addr in to_emails:
        mail.Recipients.Add(addr)
    if not mail.Recipients.ResolveAll():
        raise RuntimeError("One or more recipients could not be resolved by Outlook.")
    with tempfile.TemporaryDirectory() as tmpd:
        for fname, blob in attachments:
            path = os.path.join(tmpd, fname)
            with open(path, "wb") as f:
                f.write(blob)
            mail.Attachments.Add(os.path.abspath(path))
        mail.Send()

def send_email_smtp(
    subject: str,
    body: str,
    to_emails: List[str],
    attachments: List[Tuple[str, bytes]],
    smtp_host: str,
    smtp_port: int,
    smtp_username: Optional[str] = None,
    smtp_password: Optional[str] = None,
    use_tls: bool = True
):
    import smtplib
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email import encoders

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = smtp_username if smtp_username else "no-reply@example.com"
    msg["To"] = ", ".join(to_emails)
    msg.attach(MIMEText(body, "plain"))

    for fname, blob in attachments:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(blob)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{fname}"')
        msg.attach(part)

    if use_tls:
        server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
        server.starttls()
    else:
        server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30)

    if smtp_username and smtp_password:
        server.login(smtp_username, smtp_password)

    server.sendmail(msg["From"], to_emails, msg.as_string())
    server.quit()

# ─────────────────────────────────────────────────────────────────────────────
# Orchestration
# ─────────────────────────────────────────────────────────────────────────────
def run_year_job(
    drv: webdriver.Chrome,
    year: int,
    tmpdir: str,
    uploaded_past_df: Optional[pd.DataFrame],
    status_area=None,
    progress=None
) -> Tuple[str, str, pd.DataFrame]:
    if status_area: status_area.write(f"Scraping year **{year}**…")
    data = collect_all_data_for_year(drv, year, progress=progress)
    if data.empty:
        raise RuntimeError(f"No data scraped for {year}.")

    old_raw_df = None
    if uploaded_past_df is not None and not uploaded_past_df.empty:
        old_raw_df = uploaded_past_df[uploaded_past_df["Year"] == year]
        if old_raw_df.empty:
            old_raw_df = None

    raw_json_path = RAW_JSON_PATH(tmpdir, year)
    data.to_json(raw_json_path, orient='records', indent=2)

    excel_path = EXCEL_PATH(tmpdir, year)
    if status_area: status_area.write(f"Generating Excel for **{year}**…")
    generate_daily_totals_excel(data, year, excel_path, old_raw_df)

    return excel_path, raw_json_path, data

# ─────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="TBN Scraper", page_icon="🚌", layout="wide")

st.title("🚌 TBN Daily Usage Scraper")
st.caption("Securely authenticate, scrape specified years, compare with past data, and export/email your reports.")

with st.expander("Security & Ethics Notes", expanded=False):
    st.markdown(
        "- Credentials are held in memory only for this session and are not written to disk.\n"
        "- Consider OAuth/SSO for production deployments.\n"
        "- Scrape only content you are authorized to access and respect the website’s Terms of Service."
    )

# Sidebar: Authentication
st.sidebar.header("Authentication")
manual_login = st.sidebar.checkbox(
    "Manual login (open a real browser window)",
    value=False,
    help="Use this if headless login fails or MFA/captcha is required."
)
auth_method = st.sidebar.selectbox("Method", ["Username & Password"], help="For production, prefer OAuth/SSO.")
username = st.sidebar.text_input("TBN Username (email)", type="default", placeholder="you@company.com")
password = st.sidebar.text_input("TBN Password", type="password", placeholder="••••••••")

# Sidebar: Parameters
st.sidebar.header("Parameters")
years_text = st.sidebar.text_input(
    "Years to scrape",
    value=str(datetime.now().year),
    help='Examples: "2025", "2021-2024", or "2023, 2025-2026"'
)
email_text = st.sidebar.text_input("Send results to (optional)", value="", help="Comma- or semicolon-separated emails")

# Sidebar: Past Data Upload
st.sidebar.header("Past Data (Optional)")
uploaded = st.sidebar.file_uploader("Upload prior TBN data (JSON or CSV)", type=["json","csv"])

uploaded_past_df: Optional[pd.DataFrame] = None
if uploaded is not None:
    try:
        if uploaded.type.endswith("json") or uploaded.name.lower().endswith(".json"):
            raw = uploaded.read()
            arr = json.loads(raw.decode("utf-8"))
            uploaded_past_df = pd.DataFrame(arr)
        else:
            uploaded.seek(0)
            uploaded_past_df = pd.read_csv(uploaded)
        required_cols = {"Year", "Month"}
        if not required_cols.issubset(set(uploaded_past_df.columns)):
            st.sidebar.warning(f"Uploaded file is missing expected columns: {required_cols}. Continuing without comparisons.")
            uploaded_past_df = None
        else:
            st.sidebar.success("Past data loaded. Comparisons enabled when years overlap.")
    except Exception as e:
        st.sidebar.error(f"Could not read uploaded file: {e}")
        uploaded_past_df = None

# Sidebar: Email
st.sidebar.header("Email (Optional)")
send_email = st.sidebar.checkbox("Email the results after scraping")
email_backend = st.sidebar.selectbox("Email backend", ["Outlook (Windows only)", "SMTP"], disabled=not send_email)
smtp_host = smtp_port = smtp_user = smtp_pass = None
smtp_tls = True
if send_email:
    if email_backend == "SMTP":
        smtp_host = st.sidebar.text_input("SMTP Host", placeholder="smtp.gmail.com")
        smtp_port = st.sidebar.number_input("SMTP Port", value=587, min_value=1, max_value=65535, step=1)
        smtp_tls  = st.sidebar.checkbox("Use STARTTLS (uncheck for SSL)", value=True)
        smtp_user = st.sidebar.text_input("SMTP Username (from address)")
        smtp_pass = st.sidebar.text_input("SMTP Password / App Password", type="password")
    elif email_backend == "Outlook (Windows only)" and not HAS_OUTLOOK:
        st.sidebar.warning("Outlook COM not available on this system. Choose SMTP instead.")

# Main action
colA, colB = st.columns([1,1])
with colA:
    run_button = st.button("🚀 Run Scrape", type="primary")
with colB:
    st.download_button = st.empty()  # placeholder

# Status & containers
status = st.empty()
log_area = st.empty()
progress_bar = st.progress(0.0, text="Idle")
results_container = st.container()

if run_button:
    try:
        years = parse_years_input(years_text)
        if not years:
            st.error("Please provide at least one valid year (e.g., 2025 or 2023-2024).")
            st.stop()

        to_emails: List[str] = []
        if email_text.strip():
            to_emails = validate_email_list(email_text)

        if send_email and not to_emails:
            st.error("You selected 'Email the results' but did not provide any valid recipient emails.")
            st.stop()

        if send_email and email_backend == "Outlook (Windows only)" and not HAS_OUTLOOK:
            st.error("Outlook COM is not available. Please switch to SMTP.")
            st.stop()

        if send_email and email_backend == "SMTP":
            if not smtp_host or not smtp_port or not smtp_user or not smtp_pass:
                st.error("Please complete all SMTP fields (host, port, username, password).")
                st.stop()

        attachments: List[Tuple[str, bytes]] = []
        summary_rows: List[Dict] = []

        # ── Authentication ──────────────────────────────────────────────────
        if manual_login:
            status.info("Manual login mode: a Chrome window will open. Please log in, then click the button to continue.")
            drv_vis = build_chrome(manual_visible=True)
            drv_vis.get(LOGIN_URL)
            st.info("Complete login in the Chrome window. When finished, click the button below.")
            proceed = st.button("✅ I'm logged in — continue")
            if not proceed:
                st.stop()

            drv_vis.get(REPORT_URL); time.sleep(2)
            if "login" in drv_vis.current_url.lower():
                st.error("Still looks unauthenticated. Finish login in the browser window, then click the button again.")
                st.stop()

            # Try cookie handoff into headless for faster scraping
            drv_headless = build_chrome(manual_visible=False)
            copy_cookies(drv_vis, drv_headless, "https://portal.thebusnetwork.com/")
            drv_headless.get(REPORT_URL); time.sleep(1.5)

            if "login" in drv_headless.current_url.lower():
                st.warning("Session cookies did not persist to headless. Continuing with visible browser for scraping.")
                drv = drv_vis
            else:
                try:
                    drv_vis.quit()
                except Exception:
                    pass
                drv = drv_headless

            log_area.success("Authenticated via manual mode.")
        else:
            if not username or not password:
                st.error("Please provide your TBN username and password (or enable Manual login).")
                st.stop()
            status.info("Starting headless browser and authenticating…")
            log_area.info("Launching browser and attempting automated login.")
            drv = login_and_get_driver(username, password, log_area=log_area)

        # ── Scrape years ────────────────────────────────────────────────────
        with tempfile.TemporaryDirectory() as tmpdir:
            for idx, yr in enumerate(years, start=1):
                progress_bar.progress((idx-1)/max(len(years),1), text=f"Processing {yr}")
                try:
                    excel_path, json_path, scraped_df = run_year_job(
                        drv=drv,
                        year=yr,
                        tmpdir=tmpdir,
                        uploaded_past_df=uploaded_past_df,
                        status_area=log_area,
                        progress=progress_bar
                    )

                    with open(excel_path, "rb") as f:
                        excel_bytes = f.read()
                    with open(json_path, "rb") as f:
                        json_bytes = f.read()

                    attachments.append((os.path.basename(excel_path), excel_bytes))
                    attachments.append((os.path.basename(json_path), json_bytes))

                    # Monthly summary for display
                    month_map = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
                    ms = []
                    for m in range(1,13):
                        sub = scraped_df[scraped_df["Month"]==m]
                        if sub.empty:
                            ms.append({"Year": yr, "Month": month_map[m], "Total Moves": 0})
                            continue
                        tot_row = sub[sub["Vehicle Types"].str.upper().str.contains("TOTAL", na=False)]
                        if not tot_row.empty:
                            tr = tot_row.iloc[0]
                            days = [str(d) for d in range(1,32) if str(d) in tr.index]
                            total_moves = int(pd.to_numeric(tr[days], errors="coerce").fillna(0).sum())
                        else:
                            days = [c for c in sub.columns if c.isdigit()]
                            total_moves = int(pd.to_numeric(sub[days], errors="coerce").fillna(0).sum().sum())
                        ms.append({"Year": yr, "Month": month_map[m], "Total Moves": total_moves})
                    month_df = pd.DataFrame(ms)
                    total_year_moves = int(month_df["Total Moves"].sum())
                    summary_rows.append({"Year": yr, "Total Moves": total_year_moves})

                    with results_container:
                        st.subheader(f"Results for {yr}")
                        st.dataframe(month_df, use_container_width=True)
                        st.bar_chart(month_df.set_index("Month")["Total Moves"])

                        c1, c2 = st.columns(2)
                        with c1:
                            st.download_button(
                                label=f"⬇️ Download Excel ({yr})",
                                data=excel_bytes,
                                file_name=os.path.basename(excel_path),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        with c2:
                            st.download_button(
                                label=f"⬇️ Download Raw JSON ({yr})",
                                data=json_bytes,
                                file_name=os.path.basename(json_path),
                                mime="application/json"
                            )

                except Exception as e:
                    st.error(f"Failed to process {yr}: {e}")
                    continue

        # Close driver
        try:
            drv.quit()
        except Exception:
            pass

        # Yearly summary
        if summary_rows:
            summary_df = pd.DataFrame(summary_rows).sort_values("Year")
            st.subheader("Yearly Summary")
            st.dataframe(summary_df, use_container_width=True)
            st.line_chart(summary_df.set_index("Year")["Total Moves"])

        progress_bar.progress(1.0, text="Completed.")
        status.success("Scraping and report generation complete.")

        # Email step
        if send_email and attachments:
            st.divider()
            st.subheader("Email Delivery")
            try:
                subj = "TBN Daily Totals Reports" if len(years) > 1 else f"TBN Daily Totals Report – {years[0]}"
                body = (
                    "Hi,\n\nAttached are the updated TBN reports, including comparisons and ADA notes.\n\n"
                    "Some values may be affected by data sync or processing delays. Let me know if anything looks off.\n\nThanks!"
                )
                if email_backend == "Outlook (Windows only)":
                    send_email_outlook(subj, body, to_emails, attachments)
                else:
                    send_email_smtp(
                        subject=subj,
                        body=body,
                        to_emails=to_emails,
                        attachments=attachments,
                        smtp_host=str(smtp_host),
                        smtp_port=int(smtp_port),
                        smtp_username=smtp_user,
                        smtp_password=smtp_pass,
                        use_tls=bool(smtp_tls)
                    )
                st.success(f"Email sent to: {', '.join(to_emails)}")
            except Exception as e:
                st.error(f"Email failed: {e}")

    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()

# Footer
st.caption(
    "Tip: For production, consider an OAuth/SSO login flow that redirects users through "
    "TBN’s official login page and returns a session token for scraping. Avoid persisting passwords."
)
