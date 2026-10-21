"""
Bulk GSTIN status lookup (runs on YOUR computer, free).

Setup (once):
    pip install playwright pandas openpyxl
    playwright install chromium

Run:
    python gstin_status.py gstins.xlsx

Input: Excel/CSV with GSTINs in the first column.
Output: gstin_status_results.xlsx (saved after every GSTIN, so nothing is lost).

You log in manually in the browser window. Your credentials never go into this script.
If the search page asks for a captcha, type it and click Search yourself;
the script waits and reads the result.
"""
import re, sys, time
import pandas as pd
from playwright.sync_api import sync_playwright

LOGIN_URL = "https://services.gst.gov.in/services/login"
SEARCH_URL = "https://services.gst.gov.in/services/searchtp"
OUT = "gstin_status_results.xlsx"
WAIT_SECONDS = 180  # time allowed per GSTIN (for captcha)

LABELS = {
    "Legal Name": "Legal Name of Business",
    "Trade Name": "Trade Name",
    "Status": "GSTIN / UIN Status",
    "Registration Date": "Effective Date of registration",
    "Constitution": "Constitution of Business",
    "Taxpayer Type": "Taxpayer Type",
    "Cancellation Date": "Date of Cancellation",
}
PATTERN = re.compile(r"^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")


def read_gstins(path):
    df = pd.read_csv(path, header=None) if path.lower().endswith(".csv") else pd.read_excel(path, header=None)
    out, seen = [], set()
    for v in df.iloc[:, 0].dropna():
        g = str(v).strip().upper().replace(" ", "")
        if PATTERN.match(g) and g not in seen:
            seen.add(g)
            out.append(g)
    return out


def parse(text):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    res = {}
    for key, label in LABELS.items():
        for i, l in enumerate(lines):
            if l.lower().startswith(label.lower()):
                tail = l[len(label):].strip(" :")
                res[key] = tail or (lines[i + 1] if i + 1 < len(lines) else "")
                break
        else:
            res[key] = ""
    if not res["Status"]:
        m = re.search(r"\b(Active|Cancelled|Suspended)\b", text)
        res["Status"] = m.group(1) if m else ""
    return res


def lookup(page, gstin):
    page.goto(SEARCH_URL, wait_until="domcontentloaded")
    box = page.locator("#for_gstin")
    box.wait_for(timeout=30000)
    box.fill(gstin)
    # If a captcha is shown, the user types it and clicks Search.
    # If no captcha is needed, we click Search ourselves.
    try:
        if not page.locator("#imgcaptcha").is_visible():
            page.locator("#lotsearch").click()
    except Exception:
        pass
    print(f"   waiting for result (solve captcha / click Search if asked)...")
    deadline = time.time() + WAIT_SECONDS
    while time.time() < deadline:
        text = page.inner_text("body")
        if "Legal Name of Business" in text:
            return parse(text)
        if re.search(r"no record|invalid GSTIN|does not exist", text, re.I):
            return {"Status": "Not found / invalid", **{k: "" for k in LABELS if k != "Status"}}
        time.sleep(1.5)
    return {"Status": "Timed out", **{k: "" for k in LABELS if k != "Status"}}


def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: python gstin_status.py gstins.xlsx")
    gstins = read_gstins(sys.argv[1])
    print(f"{len(gstins)} unique valid-format GSTINs found.")
    if not gstins:
        return
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_context().new_page()
        page.goto(LOGIN_URL)
        input("\nLog in to the GST portal in the browser window, then press Enter here... ")
        for n, g in enumerate(gstins, 1):
            print(f"[{n}/{len(gstins)}] {g}")
            try:
                r = lookup(page, g)
            except Exception as e:
                r = {"Status": f"Error: {type(e).__name__}", **{k: "" for k in LABELS if k != "Status"}}
            results.append({"GSTIN": g, **r})
            pd.DataFrame(results).to_excel(OUT, index=False)
            print(f"   -> {r['Status']}")
        browser.close()
    print(f"\nDone. Results saved to {OUT}")


if __name__ == "__main__":
    main()
