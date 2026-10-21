"""
Multi-account GST portal status check (runs on YOUR computer, free).

Setup (once):
    pip install playwright pandas openpyxl
    playwright install chromium

Input file (Excel/CSV, no header needed):
    Column A = username, Column B = password

Run:
    python gst_multi_login.py accounts.xlsx

Output: gst_accounts_status.xlsx  (passwords are NOT written to it)

For each account the script opens a fresh browser window, fills in the username
and password, then WAITS for you to type the captcha (and OTP, if the portal
asks) and click Login. After login it reads the dashboard and moves to the next
account. Results are saved after every account.
"""
import re, sys, time
import pandas as pd
from playwright.sync_api import sync_playwright

LOGIN_URL = "https://services.gst.gov.in/services/login"
OUT = "gst_accounts_status.xlsx"
WAIT_SECONDS = 240          # time allowed per account for captcha / OTP
GSTIN_RE = re.compile(r"\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b")
FAIL_RE = re.compile(r"invalid username|invalid password|incorrect|locked|"
                     r"enter valid captcha|invalid captcha", re.I)


def read_accounts(path):
    df = pd.read_csv(path, header=None, dtype=str) if path.lower().endswith(".csv") \
        else pd.read_excel(path, header=None, dtype=str)
    rows = []
    for _, r in df.iterrows():
        u, p = str(r.iloc[0]).strip(), str(r.iloc[1]).strip()
        if u and p and u.lower() not in ("nan", "username") and p.lower() != "nan":
            rows.append((u, p))
    return rows


def after_label(lines, label):
    for i, l in enumerate(lines):
        if l.lower().startswith(label.lower()):
            tail = l[len(label):].strip(" :")
            return tail or (lines[i + 1] if i + 1 < len(lines) else "")
    return ""


def read_dashboard(page):
    text = page.inner_text("body")
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    m = GSTIN_RE.search(text)
    st = re.search(r"\b(Active|Cancelled|Suspended)\b", text)
    return {
        "GSTIN": m.group(0) if m else "",
        "Legal Name": after_label(lines, "Legal Name"),
        "Trade Name": after_label(lines, "Trade Name"),
        "Status": st.group(1) if st else "",
        "Dashboard Text (first 300 chars)": " | ".join(lines)[:300],
    }


def try_fill(page, selectors, value):
    for s in selectors:
        try:
            el = page.locator(s).first
            if el.is_visible(timeout=1500):
                el.fill(value)
                return True
        except Exception:
            continue
    return False


def run_account(browser, user, pwd):
    ctx = browser.new_context()          # fresh session per account
    page = ctx.new_page()
    try:
        page.goto(LOGIN_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        ok1 = try_fill(page, ["#username", "input[name='username']"], user)
        ok2 = try_fill(page, ["#user_pass", "input[type='password']"], pwd)
        if not (ok1 and ok2):
            print("   could not find login fields; fill them in manually.")
        print("   type the captcha (and OTP if asked) and click Login...")
        deadline = time.time() + WAIT_SECONDS
        while time.time() < deadline:
            url = page.url.lower()
            if "dashboard" in url or "/auth/" in url:
                page.wait_for_timeout(3000)
                d = read_dashboard(page)
                return {"Login": "Success", **d}
            body = page.inner_text("body")
            if FAIL_RE.search(body):
                return {"Login": "Failed (wrong credentials / captcha?)"}
            time.sleep(1.5)
        return {"Login": "Timed out"}
    except Exception as e:
        return {"Login": f"Error: {type(e).__name__}"}
    finally:
        ctx.close()


def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: python gst_multi_login.py accounts.xlsx")
    accounts = read_accounts(sys.argv[1])
    print(f"{len(accounts)} accounts found.")
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        for n, (u, pw) in enumerate(accounts, 1):
            print(f"[{n}/{len(accounts)}] {u}")
            r = run_account(browser, u, pw)
            results.append({"Username": u, **r})
            pd.DataFrame(results).to_excel(OUT, index=False)
            print(f"   -> {r['Login']}  {r.get('Status', '')}")
        browser.close()
    print(f"\nDone. Results saved to {OUT}")


if __name__ == "__main__":
    main()
