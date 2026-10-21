"""
Download monthly returns for multiple GST accounts (runs on YOUR computer, free).

Needs gst_multi_login.py in the same folder.

Run:
    python gst_returns.py accounts.xlsx
(column A = username, column B = password)

The script asks which financial year and which returns you want.
Your only action: type the captcha (and OTP, if asked) once per account.


         GSTR1  : GSTR1_<GSTIN>_<MMYYYY>.pdf
         GSTR3B : GSTR3B_<GSTIN>_<MMYYYY>.pdf
         GSTR2B : <MMYYYY>_<GSTIN>_GSTR2B.xlsx
Log:   gst_returns_log.xlsx   (failed steps also save a screenshot in the folder)
"""
import json, re, sys, time
from pathlib import Path
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from gst_multi_login import read_accounts, try_fill, read_dashboard, LOGIN_URL, FAIL_RE

MONTHS = ["April", "May", "June", "July", "August", "September",
          "October", "November", "December", "January", "February", "March"]
RETURN_PATTERNS = {"GSTR1": r"GSTR-?1\b", "GSTR3B": r"GSTR-?3B", "GSTR2B": r"GSTR-?2B"}
RETURNS_URL = "https://return.gst.gov.in/returns/auth/dashboard"
LOG = "gst_returns_log.xlsx"
LOGIN_WAIT = 240
BTN_RE = re.compile("download", re.I)
STEP_RE = re.compile(r"generate|download|pdf|json|excel|summary", re.I)


def ask_settings():
    fy = input("Financial year (e.g. 2025-26): ").strip() or "2025-26"
    names = list(RETURN_PATTERNS)
    print("Returns:", ", ".join(f"{i}={n}" for i, n in enumerate(names, 1)))
    pick = input("Which ones? (e.g. 1,2 or 'all'): ").strip().lower()
    if pick in ("", "all"):
        chosen = names
    else:
        chosen = []
        for tok in re.split(r"[,\s]+", pick):
            t = re.sub(r"^gstr-?", "", tok.strip().lower())      # "gstr2b", "gstr-2b" -> "2b"
            by_name = {"1": "GSTR1", "3b": "GSTR3B", "2b": "GSTR2B"}
            if t in by_name:                                      # typed a return name: 2b, 3b
                n = by_name[t]
            elif t.isdigit() and 0 < int(t) <= len(names):        # typed a menu number: 1, 2, 3
                n = names[int(t) - 1]
            else:
                print(f"   (ignored '{tok}')")
                continue
            if n not in chosen:
                chosen.append(n)
        if not chosen:
            sys.exit("No valid return chosen.")
    return fy, {n: RETURN_PATTERNS[n] for n in chosen}


def login(page, user, pwd):
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    try_fill(page, ["#username", "input[name='username']"], user)
    try_fill(page, ["#user_pass", "input[type='password']"], pwd)
    print("   >> type captcha (and OTP if asked) and click Login")
    end = time.time() + LOGIN_WAIT
    while time.time() < end:
        if "dashboard" in page.url.lower() or "/auth/" in page.url.lower():
            page.wait_for_timeout(3000)
            return read_dashboard(page)
        if FAIL_RE.search(page.inner_text("body")):
            return None
        time.sleep(1.5)
    return None


DIRECT_OK = False                     # the portal answers direct URL loads with "Access Denied", so always click through menus


def goto_returns_dashboard(page):
    """Load the Returns Dashboard URL directly. Returns True when the dropdowns are showing."""
    global DIRECT_OK
    if not DIRECT_OK:
        return False
    try:
        page.goto(RETURNS_URL, wait_until="domcontentloaded", timeout=20000)
        page.wait_for_selector("select[name='fin']", timeout=15000)
        return True
    except Exception:
        DIRECT_OK = False
        print("   [nav] direct load of Returns Dashboard failed; using click routes from now on")
        return False


class SessionEnded(RuntimeError):
    pass


def open_returns_dashboard(page):
    """The portal blocks direct URLs (Access Denied), so click through its menus."""
    if page.locator("select[name='fin']").count():
        return
    if goto_returns_dashboard(page):
        return
    def vis_link(rx):
        return page.locator("a:visible").filter(has_text=re.compile(rx, re.I))

    def via_dropdown():                # Services menu > Returns > "Returns Dashboard" (the master link)
        def txt(t):                    # any visible element whose own text is exactly t
            return page.locator(f":text-matches('^\\\\s*{t}\\\\s*$', 'i'):visible")
        svc = vis_link(r"^\s*Services\s*$").first
        svc.hover(timeout=4000)
        page.wait_for_timeout(600)
        rd = txt("Returns Dashboard")
        if not rd.count():                         # submenu not open yet: hover the 'Returns' category
            txt("Returns").first.hover(timeout=3000)
            page.wait_for_timeout(600)
        try:
            rd.first.click(timeout=4000)
        except Exception:                          # last resort: fire the link's click from the page itself
            ok = page.evaluate("""() => { const a = [...document.querySelectorAll('a')]
                .find(e => /^\\s*Returns Dashboard\\s*$/i.test(e.textContent)); if (a) { a.click(); return true } return false }""")
            if not ok:
                raise LookupError("Returns Dashboard menu link not found")
        page.wait_for_load_state("domcontentloaded", timeout=30000)

    def via_quicklinks():              # click the 'Returns Dashboard' link in the page content
        link = vis_link(r"^\s*Returns Dashboard\s*$").last   # .last = the page-body link, not the menu
        if not link.count():
            raise LookupError("no Returns Dashboard link on this page")
        try:
            link.click(timeout=6000)
        except Exception:
            page.evaluate("""() => { const a = [...document.querySelectorAll('a')]
                .filter(e => /^\\s*Returns Dashboard\\s*$/i.test(e.textContent)).pop(); if (a) a.click(); }""")
        page.wait_for_load_state("domcontentloaded", timeout=30000)

    def via_dashboard_button():        # dashboard home -> RETURN DASHBOARD button/link in the navigation panel
        btn = page.locator("button:visible, a:visible").filter(has_text=re.compile(r"return\s+dashboard", re.I))
        if not btn.count():
            vis_link(r"^\s*Dashboard\s*$").first.click(timeout=8000)      # go to dashboard home first
            page.wait_for_load_state("domcontentloaded", timeout=30000)
            btn.first.wait_for(state="visible", timeout=25000)
        btn.first.scroll_into_view_if_needed(timeout=5000)
        btn.first.click(timeout=8000)
        page.wait_for_load_state("domcontentloaded", timeout=30000)

    if "quicklinks" in page.url.lower():
        routes = [via_dropdown, via_quicklinks, via_dashboard_button]
    else:
        routes = [via_dropdown, via_dashboard_button]
    for route in routes:
        try:
            if "successfully logged out" in page.inner_text("body").lower():
                raise SessionEnded("Portal session ended (logged out)")
            if page.locator("select[name='fin']").count():
                return
            print(f"   [nav] trying {route.__name__} (page: {page.url[-45:]})")
            route()
            page.wait_for_selector("select[name='fin']", timeout=40000)
            return
        except SessionEnded:
            raise
        except Exception as e:
            print(f"   [nav] {route.__name__} failed: {type(e).__name__}")
            page.wait_for_timeout(2000)
    page.screenshot(path="DEBUG_returns_nav.png")
    Path("DEBUG_returns_nav.html").write_text(page.content(), encoding="utf-8")
    raise RuntimeError("Could not open Returns Dashboard (see DEBUG_returns_nav.png/.html)")


def dump_selects(page):
    try:
        info = page.evaluate("""() => [...document.querySelectorAll('select')].map(s => ({
            name: s.name, id: s.id,
            options: [...s.options].map(o => o.text.trim()).slice(0, 15)}))""")
        print("   [debug] dropdowns on page:", info)
    except Exception as e:
        print("   [debug] could not read dropdowns:", e)


def select_period(page, fy, month):
    open_returns_dashboard(page)
    # Use the three VISIBLE dropdowns (FY, Quarter, Period), not every <select> on the page
    fin = page.locator("select[name='fin']").first
    vis = page.locator("select:visible")
    try:
        fin.select_option(label=fy, timeout=8000)                      # Financial Year
        page.wait_for_timeout(300)
        q_no = MONTHS.index(month) // 3 + 1
        q_sel = vis.nth(1)
        q_label = q_sel.evaluate(
            "(s, n) => { const o=[...s.options].find(o => o.text.includes('Quarter '+n)); return o ? o.text : null }",
            q_no)
        q_sel.select_option(label=q_label, timeout=8000)               # Quarter
        page.wait_for_timeout(300)
        vis.nth(2).select_option(label=month, timeout=8000)            # Period
    except Exception:
        dump_selects(page)
        page.screenshot(path="DEBUG_select_period.png")
        raise
    page.mouse.move(5, 600)                                       # close any open hover menu
    search_btn = page.locator("button, input[type=submit], input[type=button]").filter(
        has_text=re.compile(r"^\s*search\s*$", re.I)).first
    # only VISIBLE View/Download buttons count (a hidden first match made the old check time out)
    tile_txt = page.locator("button:visible, a:visible").filter(
        has_text=re.compile(r"^\s*(view|download)\s*$", re.I))
    for attempt in (1, 2):                                        # click Search, retry once if nothing appears
        search_btn.click()
        try:
            tile_txt.first.wait_for(state="visible", timeout=20000)
            break
        except Exception:
            if attempt == 2:
                try:
                    sel = page.evaluate("() => [...document.querySelectorAll('select')]"
                                        ".filter(s => s.offsetParent).map(s => s.options[s.selectedIndex]?.text)")
                    Path("DEBUG_no_tiles.txt").write_text(f"selected: {sel}\n\n" + page.inner_text("body"),
                                                          encoding="utf-8")
                except Exception:
                    pass
                page.screenshot(path="DEBUG_no_tiles.png", full_page=True)
                raise RuntimeError(f"No return tiles after Search for {month} (see DEBUG_no_tiles.png/.txt)")
    page.wait_for_timeout(300)


def tiles_showing(page, pattern):
    try:
        return page.locator("button, a").filter(has_text=re.compile(r"^\s*view\s*$", re.I)).count() > 0 \
            and page.get_by_text(re.compile(pattern)).count() > 0
    except Exception:
        return False


def back_to_dashboard(page):
    """Cheap way back to the Returns Dashboard (browser Back) instead of menu clicking."""
    if not page.locator("select[name='fin']").count() and goto_returns_dashboard(page):
        return
    if "gstr2b" in page.url.lower():
        return                                   # 2B page: next step uses the Services > Returns Dashboard menu
    for _ in range(3):
        if page.locator("select[name='fin']").count():
            return
        if "gstr" not in page.url.lower():       # never Back past the dashboard (would hit the login page)
            return
        try:                                     # breadcrumb "Returns" is the quickest way back
            crumb = page.locator("a:visible").filter(has_text=re.compile(r"^\s*Returns\s*$", re.I))
            if crumb.count():
                crumb.first.click(timeout=4000)
                page.locator("select[name='fin']").first.wait_for(timeout=15000)
                return
        except Exception:
            pass
        if "gstr2b" in page.url.lower():
            return                               # no browser-Back on the 2B page; open_returns_dashboard handles it
        try:
            page.go_back(wait_until="domcontentloaded", timeout=10000)
            page.locator("select[name='fin']").first.wait_for(timeout=4000)
            return
        except Exception:
            continue


def file_stem(who, fy, month, ret):
    """File name (without extension) in the form the GST Tool reads: no 'File name is not in ... form' warning."""
    mn = MONTHS.index(month)
    cal = (mn + 4 - 1) % 12 + 1                       # April -> 4 ... March -> 3
    y0 = int(fy[:4])
    year = y0 + 1 if cal <= 3 else y0
    mmyyyy = f"{cal:02d}{year}"
    if ret == "GSTR2B":
        return f"{mmyyyy}_{who}_GSTR2B"
    return f"{ret}_{who}_{mmyyyy}"


def _manifest(folder):
    p = folder / ".done.json"
    try:
        return p, json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return p, {}


def is_done(folder, stem):
    """Resume check: was this month's file saved earlier (and still exists)?"""
    _, m = _manifest(folder)
    return stem in m and (folder / m[stem]).exists()


def save(dl, folder, stem):
    """Keep the file name exactly as the GST portal supplies it."""
    name = dl.suggested_filename or (stem + ".dat")
    target = folder / name
    if target.exists():                      # portal used the same name for another month: don't overwrite
        n = Path(name)
        target = folder / f"{n.stem}__{stem}{n.suffix}"
    dl.save_as(str(target))
    p, m = _manifest(folder)
    m[stem] = target.name
    p.write_text(json.dumps(m, indent=1), encoding="utf-8")
    return target.name


def next_step(page, counts):
    """Find a visible Generate/Download/PDF/JSON/Excel button not clicked twice yet."""
    for role in ("button", "link"):
        for el in page.get_by_role(role, name=STEP_RE).all():
            try:
                label = el.inner_text().strip().lower()
                if label and el.is_visible() and counts.get(label, 0) < 2:
                    counts[label] = counts.get(label, 0) + 1
                    return el, label
            except Exception:
                continue
    return None, None


DL_ANY = re.compile("download", re.I)


def dump_page(page, folder, stem):
    try:
        items = page.evaluate("""() => [...document.querySelectorAll('button,a,input[type=button]')]
            .filter(e => e.offsetParent !== null)
            .map(e => (e.innerText || e.value || '').trim()).filter(t => t)""")
        print("   [debug] visible buttons/links:", items[:40])
        (folder / f"FAILED_{stem}.txt").write_text(page.inner_text("body"), encoding="utf-8")
        page.screenshot(path=str(folder / f"FAILED_{stem}.png"), full_page=True)
    except Exception as e:
        print("   [debug] dump failed:", e)


def view_pdf(page, pattern, folder, stem):
    """VIEW on the return's tile -> (VIEW SUMMARY if shown) -> DOWNLOAD (PDF)."""
    tile = None
    for css in ("div[class*='col-']", "div[class*='card']", "div"):
        cand = page.locator(css).filter(has_text=re.compile(pattern)).filter(
            has=page.locator("button, a").filter(has_text=re.compile(r"^\s*view\s*$", re.I)))
        if cand.count():
            tile = cand.last
            break
    if tile is None:
        return None
    tile.locator("button, a").filter(has_text=re.compile(r"^\s*view\s*$", re.I)).first.click(timeout=8000)
    pdf_btn = page.locator("button, a").filter(has_text=re.compile(r"download\s*\(pdf\)", re.I))
    summ_btn = page.locator("button, a").filter(has_text=re.compile(r"view\s*summary", re.I))
    either = pdf_btn.or_(summ_btn)
    for _ in range(3):
        try:
            either.first.wait_for(state="visible", timeout=15000)   # proceeds the moment it appears
        except Exception:
            break
        if pdf_btn.count() and pdf_btn.first.is_visible():
            break
        summ_btn.first.click(timeout=8000)
        page.wait_for_timeout(500)
    if not (pdf_btn.count() and pdf_btn.first.is_visible()):
        dump_page(page, folder, stem)
        return None
    with page.expect_download(timeout=30000) as di:
        pdf_btn.first.click(timeout=8000)
    return save(di.value, folder, stem)


def excel_2b(page, pattern, folder, stem):
    """GSTR-2B tile -> DOWNLOAD -> 'GENERATE EXCEL FILE TO DOWNLOAD' -> grab the generated file."""
    tile = None
    for css in ("div[class*='col-']", "div[class*='card']", "div"):
        cand = page.locator(css).filter(has_text=re.compile(pattern)).filter(
            has=page.locator("button, a").filter(has_text=DL_ANY))
        if cand.count():
            tile = cand.last
            break
    if tile is None:
        return None
    tile.locator("button, a").filter(has_text=DL_ANY).first.click(timeout=8000)
    gen = page.locator("button, a").filter(has_text=re.compile(r"generate\s+excel", re.I)).first
    gen.wait_for(state="visible", timeout=20000)
    try:                                              # sometimes the file starts downloading at once
        with page.expect_download(timeout=6000) as di:
            gen.click(timeout=8000)
        return save(di.value, folder, stem)
    except PWTimeout:
        pass
    cand = page.locator("button:visible, a:visible").filter(has_text=re.compile(r"download(?!s)", re.I)).filter(
        has_not_text=re.compile(r"generate", re.I))
    end = time.time() + 240                           # big 2B files can take a few minutes to build
    last_note = 0
    while time.time() < end:
        if time.time() - last_note > 20:              # show what the portal is saying while we wait
            last_note = time.time()
            try:
                body = page.inner_text("body")
                hint = [l.strip() for l in body.splitlines()
                        if re.search(r"generat|progress|wait|minute|available|link|error|try again|already", l, re.I)]
                print("   [2B wait]", " | ".join(hint)[:200] or "(no status text)")
            except Exception:
                pass
        for i in range(cand.count()):
            el = cand.nth(i)
            try:
                if not el.is_visible():
                    continue
                with page.expect_download(timeout=8000) as di:
                    el.click(timeout=4000)
                return save(di.value, folder, stem)
            except Exception:
                continue
        page.wait_for_timeout(1500)
    dump_page(page, folder, stem)
    return None


def get_return(page, pattern, folder, stem):
    if "2B" in pattern:                               # GSTR-2B: Excel file, not PDF
        try:
            fname = excel_2b(page, pattern, folder, stem)
        except Exception as e:
            print(f"   [2B excel failed: {type(e).__name__}: {str(e)[:80]}]")
            fname = None
        return ("Downloaded (Excel)", fname) if fname else ("2B Excel not downloaded (see FAILED_ files)", "")
    try:
        fname = view_pdf(page, pattern, folder, stem)
        if fname:
            return "Downloaded (View > PDF)", fname
    except Exception as e:
        print(f"   [view/pdf route failed: {type(e).__name__}; trying tile Download]")
        try:
            open_returns_dashboard(page)
        except Exception:
            pass
    tile = None
    for css in ("div[class*='col-']", "div[class*='card']", "div"):
        cand = page.locator(css).filter(has_text=re.compile(pattern)).filter(
            has=page.locator("button, a").filter(has_text=DL_ANY))
        if cand.count():
            tile = cand.last
            break
    if tile is None:
        dump_page(page, folder, stem)
        return "Return tile not found", ""
    BTN = page.locator("button, a").filter(has_text=DL_ANY)
    try:
        with page.expect_download(timeout=15000) as di:
            tile.locator("button, a").filter(has_text=DL_ANY).first.click(timeout=8000)
        return "Downloaded", save(di.value, folder, stem)
    except PWTimeout:
        pass
    counts = {}
    for _ in range(8):                       # follow Generate -> Download steps
        el, label = next_step(page, counts)
        if el is None:
            break
        try:
            with page.expect_download(timeout=15000) as di:
                el.click(timeout=8000)
            return f"Downloaded (via '{label}')", save(di.value, folder, stem)
        except PWTimeout:
            page.wait_for_timeout(3000)      # file may still be generating
        except Exception:
            continue
    try:
        page.screenshot(path=str(folder / f"FAILED_{stem}.png"))
    except Exception:
        pass
    return "Not downloaded (see screenshot)", ""


def main():
    print("== GST Returns Downloader ==")
    path = sys.argv[1] if len(sys.argv) > 1 else input("Accounts file path: ").strip()
    accounts = read_accounts(path)
    fy, returns = ask_settings()
    print(f"\n{len(accounts)} accounts | FY {fy} | returns: {', '.join(returns)}\n")
    log = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        for n, (u, pw) in enumerate(accounts, 1):
            print(f"[{n}/{len(accounts)}] {u}")
            ctx = browser.new_context(accept_downloads=True)
            page = ctx.new_page()
            info = login(page, u, pw)
            if not info:
                log.append({"Username": u, "Result": "Login failed / timed out"})
                pd.DataFrame(log).to_excel(LOG, index=False)
                ctx.close()
                continue
            who = info["GSTIN"] or u
            base_folder = Path("downloads") / who / fy
            for m in MONTHS:
                need_select = True
                for ret, pat in returns.items():
                    stem = file_stem(who, fy, m, ret)
                    folder = base_folder / ret                # one sub-folder per return, so each can be dropped into its own tool
                    folder.mkdir(parents=True, exist_ok=True)
                    done = is_done(folder, stem)
                    if done:                                  # resume: skip what is already saved
                        print(f"   {m} {ret}: already downloaded, skipped")
                        continue
                    try:
                        if need_select or not tiles_showing(page, pat):
                            select_period(page, fy, m)        # also gets us back from View/Summary pages
                        res, fname = get_return(page, pat, folder, stem)
                        need_select = False
                        back_to_dashboard(page)
                    except Exception as e:
                        need_select = True
                        res, fname = f"Error: {type(e).__name__}: {str(e)[:80]}", ""
                    print(f"   {m} {ret}: {res}")
                    log.append({"Username": u, "GSTIN": who, "Month": m, "Return": ret,
                                "Result": res, "File": fname})
                    pd.DataFrame(log).to_excel(LOG, index=False)
            ctx.close()
        browser.close()
    print(f"\nDone. Files in ./downloads, log in {LOG}")


if __name__ == "__main__":
    main()
