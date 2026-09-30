# Certificate PDF Builder

Combines the ECG & NIBP certificates and pass/fail test sheets of many devices into **two print-ready PDF files**.

## ⬇️ Download (Windows)

<p align="center">
  <a href="https://github.com/AhmedGehad1/new-simple-app/releases/latest/download/CertificatePDFBuilder.zip">
    <img src="https://img.shields.io/badge/%E2%AC%87%20Download-CertificatePDFBuilder.zip-1F5FA8?style=for-the-badge" alt="Download CertificatePDFBuilder.zip">
  </a>
</p>

**Direct link:** <https://github.com/AhmedGehad1/new-simple-app/releases/latest/download/CertificatePDFBuilder.zip>

1. **Click the Download button above.** `CertificatePDFBuilder.zip` (about 23 MB) downloads straight away. You don't need a GitHub account.
2. **Unzip it:** right-click the downloaded zip → **Extract All…** → **Extract**.
3. **Open the extracted `CertificatePDFBuilder` folder and double-click `CertificatePDFBuilder.exe`.** There's nothing to install, and you don't need Python.
4. If Windows says *"Windows protected your PC"*, click **More info → Run anyway**. It appears only because the app isn't code-signed.

The zip also contains **How to use.txt** with short instructions.
To see all versions, go to the [Releases page](https://github.com/AhmedGehad1/new-simple-app/releases).

---

## What it does

You pick the Excel workbooks (one per device). For every workbook, the app prints these tabs exactly as they would come out of the printer:

| PDF | Pages, for every device in your list |
| --- | --- |
| **Certificates.pdf** | `ECG.certificate`, then `NIBP.certificate` |
| **Pass-Fail Test Sheets.pdf** | `ECG.pass-fail`, then `NIBP.pass-fail` |

Device 1's pages come first, then device 2's, and so on. Each PDF also gets bookmarks (one per device) so you can jump to any device.

## What the PC needs

The app uses the PC's spreadsheet program to print the tabs, so it needs **one** of these:

* **Microsoft Excel** (recommended): gives exactly the same pages as *File → Print* in Excel.
* **LibreOffice** (free, [download](https://www.libreoffice.org/download/download-libreoffice/)) for PCs without Excel.

The app finds whichever is installed on its own. With *Convert with: Automatic*, Excel is used when available.

---

## How to use

1. **Choose the Excel files.** Click **＋ Add files…** (or **Add folder…**).
   The table shows ✔ / ✖ for each of the four tabs, so you can see straight away if a tab is missing.
   The PDFs follow the order of the list. Use **▲ ▼** or **Sort A–Z** to change it.
2. **Where to save.** Pick the folder and, if you want, change the two PDF file names.
3. Click **Create PDFs**. The table shows the progress of each file. At the end a summary appears and the folder
   opens (untick *Open the folder when finished* if you don't want that). The **Open:** links show the PDFs right away.

### Tab names

The app looks for tabs named `ECG.certificate`, `NIBP.certificate`, `ECG.pass-fail` and `NIBP.pass-fail`.
Capital letters, spaces, dots and dashes don't matter, so `ECG Certificate` or `ecg_certificate` are found too.
If your workbooks use different names, click **Tab names…** and type them in. The app remembers them.

Files with a missing tab are still processed: the tabs that exist go into the PDFs, and the app tells you what was left out.

---

## Troubleshooting

| Problem | What to do |
| --- | --- |
| *"Excel / LibreOffice not found"* | Install LibreOffice (free) or run the app on a PC with Excel. |
| A tab shows ✖ | Select the file: the line under the table lists the tabs it has. Fix the name under **Tab names…**. |
| *"Could not save … is it open in a PDF viewer?"* | Close the PDF in Acrobat, Edge, etc. and click **Create PDFs** again. |
| Pages look different from Excel's print preview | Choose **Convert with: Microsoft Excel**. LibreOffice can differ slightly with some fonts. |
| Page layout is wrong (cut off, several pages) | Fix the tab's *Page Layout* (print area, *Fit to 1 page*) in the Excel file; the app prints what Excel prints. |

The app never changes your Excel files; it only reads them.

---

## For developers

Run from source (Windows, macOS or Linux, Python 3.9+):

```bash
pip install -r requirements.txt
python -m certpdf            # or double-click CertificatePDFBuilder.pyw on Windows
```

On macOS/Linux the app uses LibreOffice. Linux needs the `python3-tk` package.

**Build the Windows .exe yourself:** on a Windows PC with Python installed, double-click `build_windows.bat`.
The program appears in `dist\CertificatePDFBuilder.exe`.
GitHub Actions does the same on every push (`.github/workflows/build.yml`). To publish a new download, raise `__version__`
in `certpdf/__init__.py` and push a tag such as `v1.0.1`. The workflow then creates a release with `CertificatePDFBuilder.zip`,
and the Download button always points to the newest one.

**Tests:** `pip install -r requirements-dev.txt` then `python -m pytest`. The end-to-end tests run only when LibreOffice is installed.

| Path | What it is |
| --- | --- |
| `certpdf/gui.py` | The window (tkinter) |
| `certpdf/builder.py` | Prints the tabs of every workbook and merges them into the two PDFs |
| `certpdf/engines.py` | Printing a tab with Microsoft Excel (COM automation) or LibreOffice (headless) |
| `certpdf/workbook_info.py` | Reading tab names and matching them without opening Excel |
| `certpdf/config.py` | Tab definitions and saved settings |
| `tools/build_exe.py` | PyInstaller build of the single-file .exe |
| `tools/make_icon.py` | Draws the app icon |
