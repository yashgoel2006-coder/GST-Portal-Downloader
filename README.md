# GST Toolkit

Bulk helpers for the GST portal. Runs on your own computer; **you** type the captcha/OTP.

## Use without installing anything
**GSTIN checker** (format + checksum validation, in your browser):
`https://yashgoel2006-coder.github.io/GST-Portal-Downloader/`

## Desktop tools (returns download, multi-login, GSTIN status)
1. Install Python 3 (python.org, tick "Add to PATH" on Windows).
2. Click **Code > Download ZIP**, unzip.
3. Put `accounts.xlsx` in the folder (column A = username, column B = password).
4. Double-click **Start-Windows.bat** (or **Start-Mac-Linux.command**), pick a tool.

First run installs Playwright + Chromium automatically.

Output: `downloads/<GSTIN>/<FY>/<Return>/...` and `gst_returns_log.xlsx`. Re-running skips files already saved.

## Security
`accounts.xlsx` holds passwords. `.gitignore` blocks it from being committed. Keep the repo **private** or never add that file. Use only your own / your clients' authorised credentials.

## Publish
1. Create a repo on GitHub, upload these files.
2. Settings > Pages > Source: `main` branch, folder `/docs`.
3. Replace `yashgoel2006-coder` in `README.md` and `docs/index.html`.
