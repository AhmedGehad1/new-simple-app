# Certificate Excel Builder

Combines the ECG & NIBP certificates and pass/fail test sheets of many devices into **two Excel files**.

## ⬇️ Download (Windows)

<p align="center">
  <a href="https://github.com/AhmedGehad1/new-simple-app/raw/HEAD/download/CertificateExcelBuilder.zip">
    <img src="https://img.shields.io/badge/%E2%AC%87%20Download-CertificateExcelBuilder.zip-1E7D45?style=for-the-badge" alt="Download CertificateExcelBuilder.zip">
  </a>
</p>

**Direct link:** <https://github.com/AhmedGehad1/new-simple-app/raw/HEAD/download/CertificateExcelBuilder.zip>

1. **Click the Download button above.** `CertificateExcelBuilder.zip` (about 20 MB) downloads straight away. You don't need a GitHub account.
2. **Unzip it:** right-click the downloaded zip → **Extract All…** → **Extract**.
3. **Open the extracted `CertificateExcelBuilder` folder and double-click `CertificateExcelBuilder.exe`.** There's nothing to install, and you don't need Python.
4. If Windows says *"Windows protected your PC"*, click **More info → Run anyway**. It appears only because the app isn't code-signed.

The zip also contains **How to use.txt** with short instructions.

---

## What it does

You pick the Excel workbooks (one per device). The app makes two Excel files with two sheets each:

| File | Sheets |
| --- | --- |
| **Certificates.xlsx** | **ECG certificates** (the `ECG.certificate` tab of every device) · **NIBP certificates** (`NIBP.certificate`) |
| **Pass-Fail Test Sheets.xlsx** | **ECG reports** (`ECG.pass-fail`) · **NIBP reports** (`NIBP.pass-fail`) |

* Each sheet has that tab of every device one after another, in the order of your list, with a page break between
  devices.
* A sheet only holds copies of one template, so it keeps that template's print scale, margins, paper and footer.
  Every device's page prints exactly like its original tab.
* Column widths, row heights, merged cells, borders and conditional formatting are kept. Column widths stay exact
  even when one device's copy of the template has different widths.
* Formulas are replaced by the values from the original, so the files no longer depend on the other tabs.
* To print, open a file in Excel (any version, 2007 included) and choose **Print → Print entire workbook** for both
  sheets.

**The app never prints anything.** It only reads your Excel files and never changes them. It doesn't need Excel
itself; old `.xls` files are converted automatically if the free
[LibreOffice](https://www.libreoffice.org/download/download-libreoffice/) is installed (or save them as `.xlsx`).

---

## How to use

1. **Choose the Excel files.** Click **＋ Add files…** (or **Add folder…**).
   The table shows ✔ / ✖ for each of the four tabs, so you can see straight away if a tab is missing.
   The output follows the order of the list. Use **▲ ▼** or **Sort A–Z** to change it.
2. **Where to save.** Pick the folder and, if you want, change the two file names (the app adds `.xlsx`).
3. Click **Create files**. The table shows the progress of each file. At the end a summary appears and the folder
   opens (untick *Open the folder when finished* if you don't want that). The **Open:** links open the files right
   away.

### Tab names

The app looks for tabs named `ECG.certificate`, `NIBP.certificate`, `ECG.pass-fail` and `NIBP.pass-fail`.
Capital letters, spaces, dots and dashes don't matter, and extra words are fine. For example,
`ECG.pass-fail test sheet` and `NIBP.pass-fail sheet` are found too.
If your workbooks use different names, click **Tab names…** and type them in. The app remembers them.

Files with a missing tab are still processed: the tabs that exist are included, and the app tells you what was left
out.

---

## Troubleshooting

| Problem | What to do |
| --- | --- |
| A tab shows ✖ | Select the file: the line under the table lists the tabs it has. Fix the name under **Tab names…**. |
| *"Could not save … is it open?"* | Close the file in Excel and click **Create files** again. |
| *"this is an old .xls file"* | Open it in Excel and save it as `.xlsx`, or install LibreOffice (free). |
| Page layout is wrong (cut off, several pages) | Fix the tab's *Page Layout* (print area, *Fit to 1 page*) in the original Excel file; the app keeps that setup. |

---

## For developers

Run from source (Windows, macOS or Linux, Python 3.9+):

```bash
pip install -r requirements.txt
python -m certpdf            # or double-click CertificateExcelBuilder.pyw on Windows
```

Linux needs the `python3-tk` package.

**Build the Windows .exe yourself:** on a Windows PC with Python installed, double-click `build_windows.bat`.
The program appears in `dist\CertificateExcelBuilder.exe`.
GitHub Actions does the same on every push (`.github/workflows/build.yml`).

**Update the download:** the Download button serves `download/CertificateExcelBuilder.zip` from this repository.
After changing the app, download the `CertificateExcelBuilder-windows` artifact from the latest successful Actions
run. It holds the .exe and *How to use.txt*. Put both in a `CertificateExcelBuilder` folder, zip that folder, and
replace `download/CertificateExcelBuilder.zip`. Optionally, publishing a GitHub release also gets the zip attached
to it automatically.

**Tests:** `pip install -r requirements-dev.txt` then `python -m pytest`. The `.xls` tests run only when LibreOffice
is installed.

| Path | What it is |
| --- | --- |
| `certpdf/gui.py` | The window (tkinter) |
| `certpdf/excel_output.py` | Reads the tabs of every workbook and writes the two Excel files |
| `certpdf/stacked_sheet.py` | Lays one tab of every device out on one sheet, keeping each one's widths and print setup |
| `certpdf/sheet_copy.py` | Reading the values Excel saved; small helpers for the copies |
| `certpdf/libreoffice.py` | Converting old `.xls` files with LibreOffice, when installed |
| `certpdf/workbook_info.py` | Reading tab names and matching them without opening Excel |
| `certpdf/config.py` | Tab definitions and saved settings |
| `tools/build_exe.py` | PyInstaller build of the single-file .exe |
| `tools/make_icon.py` | Draws the app icon |
