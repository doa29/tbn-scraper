# app.py
# ─────────────────────────────────────────────────────────────────────────────
# Streamlit TBN Scraper App (Production-ready, cross-platform)
#
# Highlights:
# - Auto-installs Chromium (via Playwright) if Chrome is missing
# - Manual-login fallback for MFA/CAPTCHA
# - Headless by default; visible window only if manual login is enabled
# - Responsive Streamlit UI (works on phones when hosted)
# - SMTP / Outlook email options
# - Past-data comparisons, Excel export, downloads
# ─────────────────────────────────────────────────────────────────────────────

import os
import re
import sys
import time
import json
import shutil
import pathlib
import tempfile
import subprocess
from datetime import datetime
from io import StringIO
from typing import List, Dict, Optional, Tuple

import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

# Excel writer
import xlsxwriter  # noqa: F401

# Outlook (optional, Windows)
try:
    import win32com.client as win32  # type: ignore
    HAS_OUTLOOK = True
except Exception:
    HAS_OUTLOOK = False

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
LOGIN_URL  = os.getenv("TBN_LOGIN_URL", "https://portal.thebusnetwork.com/login")
REPORT_URL = os.getenv("TBN_REPORT_URL", "https://portal.thebusnetwork.com/salesman/reports/daily-usage")

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
                    years.extend(range(lo, hi + 1))
        else:
            if part.isdigit():
                years.append(int(part))
    years = [y for y in years if 2000 <= y <= 2100]
    return sorted(set(years))

def validate_email_list(raw: str) -> List[str]:
    emails = [e.strip() for e in re.split(r"[;,]", raw) if e.strip()]
    bad = [e for e in emails if not EMAIL_RE.match(e)]
    if bad:
        raise ValueError(f"Invalid email(s): {', '.join(bad)}")
    return emails

# ─────────────────────────────────────────────────────────────────────────────
# Chromium bootstrap (self-healing)
# ─────────────────────────────────────────────────────────────────────────────
def _ensure_playwright_chromium() -> str:
    """
    Ensure a Chromium binary exists (Playwright). Returns path to the binary.
    Works without root; installs into ~/.cache/ms-playwright on first use.
    """
    try:
        import playwright  # noqa: F401
    except Exception:
        subprocess.run([sys.executable, "-m", "pip", "install", "playwright>=1.46"], check=True)
    # Install browser (idempotent)
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium", "--with-deps"], check=True)

    home = pathlib.Path.home()
    # Typical locations
    candidates = list(home.glob(".cache/ms-playwright/chromium-*/chrome-linux/chrome"))
    if not candidates:
        candidates = list(home.glob(".cache/ms-playwright/chromium-*/chromium/chrome"))
    if not candidates:
        raise RuntimeError("Playwright Chromium install succeeded but the chrome binary was not found.")
    return str(candidates[0])

def build_chrome(manual_visible: bool = False) -> webdriver.Chrome:
    """
    Build a Chrome driver that 'just works' on any device/host.
    - Uses system Chrome/Chromium if present (or CHROME_PATH env).
    - Otherwise auto-installs Playwright Chromium and retries.
    - Headless unless manual_visible=True (for manual login/MFA).
    """
    opts = ChromeOptions()
    if not manual_visible:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-blink-features=AutomationControlled")

    # Try CHROME_PATH; else look for system chrome/chromium
    chrome_path = os.getenv("CHROME_PATH")
    if chrome_path and os.path.exists(chrome_path):
        opts.binary_location = chrome_path
    else:
        for candidate in ["google-chrome", "chromium", "chromium-browser"]:
            p = shutil.which(candidate)
            if p:
                opts.binary_location = p
                break

    try:
        driver = webdriver.Chrome(options=opts)  # Selenium Manager fetches driver
        driver.set_page_load_timeout(45)
        return driver
    except WebDriverException as e:
        msg = str(e)
        # Typical in containers: missing Chrome -> chromedriver exit code 127 or 'cannot find'
        if "Status code was: 127" in msg or "cannot find" in msg.lower() or "no such file" in msg.lower():
            chromium_bin = _ensure_playwright_chromium()
            opts.binary_location = chromium_bin
            driver = webdriver.Chrome(options=opts)
            driver.set_page_load_timeout(45)
            return driver
        raise RuntimeError(f"Failed to start Chrome. Details: {e}") from e

# ─────────────────────────────────────────────────────────────────────────────
# Selenium login & scraping
# ─────────────────────────────────────────────────────────────────────────────
def login_and_get_driver(username: str, password: str, log_area=None) -> webdriver.Chrome:
    """
    Automated login. If your portal enforces CAPTCHA/MFA, use the manual login mode.
    """
    drv = build_chrome(manual_visible=False)
    drv.get(LOGIN_URL)
    wait = WebDriverWait(drv, 30)

    # If the login is within an iframe, switch in
    frames = drv.find_elements(By.TAG_NAME, "iframe")
    if frames:
        try:
            drv.switch_to.frame(frames[0])
            if log_area: log_area.info("Switched into login iframe.")
        except Exception:
            pass

    def find_el(css):
        return wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css)))

    try:
        # Email/username
        email = find_el("input[name='email'], input#email, input#username, input[type='email']")
        email.clear(); email.send_keys(username)

        # Password
        pwd = find_el("input[name='password'], input#password, input[type='password']")
        pwd.clear(); pwd.send_keys(password)

        # Submit button + Enter (cover both patterns)
        try:
            btn = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[@type='submit' or contains(., 'Log') or contains(., 'Sign')]"
            )))
            time.sleep(0.5)
            btn.click()
        except Exception:
            pass
        time.sleep(0.5)
        pwd.send_keys(Keys.ENTER)

        drv.switch_to.default_content()
        time.sleep(2)

        # Probe the report page to confirm auth
        drv.get(REPORT_URL)
        time.sleep(2)
        if ("login" in drv.current_url.lower()) or ("signin" in drv.current_url.lower()):
            drv.refresh()
            time.sleep(2)

        # Check for the report/table presence
        try:
            WebDriverWait(drv, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table, .table, div[data-report], #report"))
            )
        except Exception:
            raise RuntimeError("Login failed or report not found.")

        if log_area: log_area.success("Authenticated successfully.")
        return drv
    except Exception:
        snippet = drv.page_source[:800] if drv.page_source else "No page source available."
        if log_area:
            log_area.code(f"URL: {drv.current_url}\n\nHTML (first 800 chars):\n{snippet}")
        drv.quit()
        raise

def set_datepicker(drv, month: int, year: int):
    wait = WebDriverWait(drv, 20)
    inp  = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, 'div.styles_dateTimePickerInputGroup__2urdc input.datetimepicker-input')
    ))
    inp.clear()
    inp.send_keys(f"{month:02d}/{year}")
    inp.send_keys(Keys.TAB)
    time.sleep(2)

def scrape_month_data(drv, year: int, month: int) -> Optional[pd.DataFrame]:
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

def collect_all_data_for_year(drv, year: int, progress=None) -> pd.DataFrame:
    all_df = []
    for m in range(1, 13):
        if progress:
            progress.progress(m/12.0, text=f"Scraping {year} – Month {m}/12")
        df = scrape_month_data(drv, year, m)
        if df is not None:
            all_df.append(df)
    return pd.concat(all_df, ignore_index=True) if all_df else pd.DataFrame()

# ─────────────────────────────────────────────────────────────────────────────
# Report builder (full Excel, ADA markers + weekend shading + comparisons)
# ─────────────────────────────────────────────────────────────────────────────
def generate_daily_totals_excel(new_data: pd.DataFrame, year: int, excel_path: str, old_raw_df: Optional[pd.DataFrame] = None):
    monthly_totals_new: Dict[int, Dict[int, int]] = {}
    monthly_ada_new: Dict[int, Dict[int, int]] = {}

    for m in range(1, 13):
        sub = new_data[new_data['Month'] == m]

        # Totals row
        tot = sub[sub['Vehicle Types'].str.upper().str.contains("TOTAL", na=False)]
        if not tot.empty:
            tot = tot.iloc[0]
        else:
            days = [str(d) for d in range(1, 32) if str(d) in sub.columns]
            tot = sub[days].sum(numeric_only=True)
        monthly_totals_new[m] = {d: int(tot.get(str(d), 0) or 0) for d in range(1, 32)}

        # ADA (Wheelchair)
        wc = sub[sub['Vehicle Types'].str.contains("Wheelchair", case=False, na=False)]
        if not wc.empty:
            wc = wc.iloc[0]
            monthly_ada_new[m] = {d: int(wc.get(str(d), 0) or 0) for d in range(1, 32)}
        else:
            monthly_ada_new[m] = {d: 0 for d in range(1, 32)}

    # Old totals (for comparisons)
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
            monthly_totals_old[m] = {d: int(tot.get(str(d), 0) or 0) for d in range(1, 32)}
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
                        if diff > 0: arrow = f" ↑{diff}"
                        elif diff < 0: arrow = f" ↓{abs(diff)}"
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
# Email
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
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="TBN Scraper", page_icon="🚌", layout="wide")

st.title("🚌 TBN Daily Usage Scraper")
st.caption("Authenticate → scrape → compare → export/email. Mobile-friendly when hosted.")

with st.expander("Security & Ethics Notes", expanded=False):
    st.markdown(
        "- Credentials are held only in memory for this session (not written to disk).\n"
        "- Prefer OAuth/SSO for production if available.\n"
        "- Scrape only data you’re authorized to access; respect Terms of Service."
    )

# Sidebar: Authentication
st.sidebar.header("Authentication")
manual_login = st.sidebar.checkbox(
    "Manual login (open a visible browser)",
    value=False,
    help="Enable this if automated (headless) login fails or MFA/CAPTCHA is required."
)
username = st.sidebar.text_input("TBN Username (email)", placeholder="you@company.com")
password = st.sidebar.text_input("TBN Password", type="password", placeholder="••••••••")

# Sidebar: Parameters
st.sidebar.header("Parameters")
years_text = st.sidebar.text_input("Years to scrape", value=str(datetime.now().year),
                                   help='Examples: "2025", "2021-2024", or "2023, 2025-2026"')
email_text = st.sidebar.text_input("Send results to (optional)", value="",
                                   help="Comma- or semicolon-separated emails")

# Past Data Upload
st.sidebar.header("Past Data (Optional)")
uploaded = st.sidebar.file_uploader("Upload prior TBN data (JSON or CSV)", type=["json", "csv"])

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
            st.sidebar.warning(f"Uploaded file missing columns: {required_cols}. Comparisons disabled.")
            uploaded_past_df = None
        else:
            st.sidebar.success("Past data loaded. Comparisons enabled for overlapping years.")
    except Exception as e:
        st.sidebar.error(f"Could not read uploaded file: {e}")
        uploaded_past_df = None

# Email options
st.sidebar.header("Email (Optional)")
send_email = st.sidebar.checkbox("Email results after scraping")
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
    else:
        if not HAS_OUTLOOK:
            st.sidebar.warning("Outlook COM is not available on this system. Choose SMTP instead.")

# Action button
colA, colB = st.columns([1,1])
with colA:
    run_button = st.button("🚀 Run Scrape", type="primary")
with colB:
    st.download_button = st.empty()

# Status UI
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
            st.error("You selected email delivery but no valid recipient emails were provided.")
            st.stop()

        if send_email and email_backend == "Outlook (Windows only)" and not HAS_OUTLOOK:
            st.error("Outlook COM is not available. Switch to SMTP.")
            st.stop()

        if send_email and email_backend == "SMTP":
            if not smtp_host or not smtp_port or not smtp_user or not smtp_pass:
                st.error("Please complete all SMTP fields (host, port, username, password).")
                st.stop()

        # ── Authenticate ────────────────────────────────────────────────────
        if manual_login:
            status.info("Manual login: a visible browser will open. Complete login, then continue.")
            drv = build_chrome(manual_visible=True)
            drv.get(LOGIN_URL)
            st.info("After finishing login in the browser window, click the button below.")
            proceed = st.button("✅ I'm logged in — continue")
            if not proceed:
                st.stop()
            drv.get(REPORT_URL)
            time.sleep(2)
            if "login" in drv.current_url.lower():
                st.error("Still looks unauthenticated. Finish login and click the button again.")
                st.stop()
            log_area.success("Authenticated (manual mode).")
        else:
            if not username or not password:
                st.error("Provide username and password (or enable Manual login).")
                st.stop()
            status.info("Starting headless browser and authenticating…")
            log_area.info("Attempting automated login…")
            drv = login_and_get_driver(username, password, log_area=log_area)

        # ── Scrape + Build ──────────────────────────────────────────────────
        attachments: List[Tuple[str, bytes]] = []
        summary_rows: List[Dict] = []

        with tempfile.TemporaryDirectory() as tmpdir:
            for idx, yr in enumerate(years, start=1):
                progress_bar.progress((idx-1)/max(len(years),1), text=f"Processing {yr}")
                try:
                    # Scrape
                    data_df = collect_all_data_for_year(drv, yr, progress=progress_bar)
                    if data_df.empty:
                        st.warning(f"No data scraped for {yr}.")
                        continue

                    # Past data year slice (if uploaded)
                    old_df = None
                    if uploaded_past_df is not None and not uploaded_past_df.empty:
                        slice_df = uploaded_past_df[uploaded_past_df["Year"] == yr]
                        old_df = slice_df if not slice_df.empty else None

                    # Save raw JSON
                    raw_json_path = RAW_JSON_PATH(tmpdir, yr)
                    data_df.to_json(raw_json_path, orient='records', indent=2)
                    with open(raw_json_path, "rb") as f:
                        json_bytes = f.read()
                    attachments.append((os.path.basename(raw_json_path), json_bytes))

                    # Build Excel
                    excel_path = EXCEL_PATH(tmpdir, yr)
                    generate_daily_totals_excel(data_df, yr, excel_path, old_df)
                    with open(excel_path, "rb") as f:
                        excel_bytes = f.read()
                    attachments.append((os.path.basename(excel_path), excel_bytes))

                    # Summary table (monthly totals quick view)
                    month_map = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
                    ms = []
                    for m in range(1,13):
                        sub = data_df[data_df["Month"] == m]
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
                                file_name=os.path.basename(raw_json_path),
                                mime="application/json"
                            )

                except Exception as e:
                    st.error(f"Failed to process {yr}: {e}")
                    continue

        try:
            drv.quit()
        except Exception:
            pass

        if summary_rows:
            summary_df = pd.DataFrame(summary_rows).sort_values("Year")
            st.subheader("Yearly Summary")
            st.dataframe(summary_df, use_container_width=True)
            st.line_chart(summary_df.set_index("Year")["Total Moves"])

        progress_bar.progress(1.0, text="Completed.")
        status.success("Scraping and report generation complete.")

        # ── Email ───────────────────────────────────────────────────────────
        if send_email and attachments:
            st.divider()
            st.subheader("Email Delivery")
            try:
                subj = "TBN Daily Totals Reports" if len(years) > 1 else f"TBN Daily Totals Report – {years[0]}"
                body = (
                    "Hi,\n\nAttached are the updated TBN reports, including comparisons and ADA notes.\n\n"
                    "Some values may be affected by data sync or processing delays.\n\nThanks!"
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
st.caption("For production, consider hosting on a server (EC2/Render/Streamlit Cloud) and accessing from your phone.")
