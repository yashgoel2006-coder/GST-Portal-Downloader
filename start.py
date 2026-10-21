"""Menu launcher. Installs what is missing on first run, then starts the chosen tool."""
import subprocess, sys, importlib.util, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def ensure():
    need = [m for m in ("playwright", "pandas", "openpyxl") if importlib.util.find_spec(m) is None]
    if need:
        print("First-time setup: installing", ", ".join(need), "...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", *need])
    flag = ".chromium_installed"
    if not os.path.exists(flag):
        print("Downloading browser (one time, ~150 MB)...")
        subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
        open(flag, "w").close()

def main():
    ensure()
    print("\n=== GST Toolkit ===")
    print("1  Download returns (GSTR-1 / 3B / 2B) for all accounts")
    print("2  Check login + dashboard status for all accounts")
    print("3  Bulk GSTIN status lookup (portal search)")
    print("q  Quit")
    c = input("\nChoose: ").strip().lower()
    scripts = {"1": ("gst_returns.py", "accounts.xlsx"),
               "2": ("gst_multi_login.py", "accounts.xlsx"),
               "3": ("gstin_status.py", "gstins.xlsx")}
    if c not in scripts:
        return
    script, default = scripts[c]
    f = input(f"Input file [{default}]: ").strip().strip('"') or default
    subprocess.call([sys.executable, script, f])

main()
input("\nPress Enter to close...")
