# Certificate PDF Builder

Combines the ECG & NIBP certificates and pass/fail test sheets of many devices into **Excel and PDF files**.

## ⬇️ Download (Windows)

<p align="center">
  <a href="https://github.com/AhmedGehad1/new-simple-app/raw/HEAD/download/CertificatePDFBuilder.zip">
    <img src="https://img.shields.io/badge/%E2%AC%87%20Download-CertificatePDFBuilder.zip-1F5FA8?style=for-the-badge" alt="Download CertificatePDFBuilder.zip">
  </a>
</p>

**Direct link:** <https://github.com/AhmedGehad1/new-simple-app/raw/HEAD/download/CertificatePDFBuilder.zip>

1. **Click the Download button above.** `CertificatePDFBuilder.zip` (about 23 MB) downloads straight away. You don't need a GitHub account.
2. **Unzip it:** right-click the downloaded zip → **Extract All…** → **Extract**.
3. **Open the extracted `CertificatePDFBuilder` folder and double-click `CertificatePDFBuilder.exe`.** There's nothing to install, and you don't need Python.
4. If Windows says *"Windows protected your PC"*, click **More info → Run anyway**. It appears only because the app isn't code-signed.

The zip also contains **How to use.txt** with short instructions.

---

## What it does

You pick the Excel workbooks (one per device). For every workbook, the app collects these tabs, in the order of your list:

| File | Contents, for every device in your list |
| --- | --- |
| **Certificates** (.xlsx and .pdf) | `ECG.certificate`, then `NIBP.certificate` |
| **Pass-Fail Test Sheets** (.xlsx and .pdf) | `ECG.pass-fail`, then `NIBP.pass-fail` |

* **Excel files:** one tab per sheet (for example `F23-AGH019-0626 ECG cert`, `F23-AGH019-0626 NIBP cert`, …).
  Each tab keeps its layout and page setup and holds the values from the original, so it no longer
  depends on the other tabs. To get everything on paper or in one PDF yourself, open the file in any Excel
  version and choose **Print → Print entire workbook**. Making the Excel files doesn't need Excel at all.
* **PDF files:** the same pages as printing each tab, with a bookmark for every device.

**The app never prints anything.** It only reads your Excel files and never changes them.

## What the PC needs

* **Excel files:** nothing extra.
* **PDF files:** one of these.
  * **Microsoft Excel 2010 or newer.** Excel 2007 cannot save PDFs, because Microsoft's "Save as PDF"
    add-in for it is no longer available.
  * **LibreOffice** (free, [download](https://www.libreoffice.org/download/download-libreoffice/)).
    Once it's installed, the app uses it automatically; you don't need to open it.

With *PDFs made with: Automatic* (under **Options…**), the app checks whether Excel can save PDFs and uses
LibreOffice when it can't. If neither can, the Excel files are still created and the app tells you why no PDFs
were made.

---

## How to use

1. **Choose the Excel files.** Click **＋ Add files…** (or **Add folder…**).
   The table shows ✔ / ✖ for each of the four tabs, so you can see straight away if a tab is missing.
   The output follows the order of the list. Use **▲ ▼** or **Sort A–Z** to change it.
2. **Where to save.** Pick the folder and, if you want, change the two file names (the app adds `.xlsx` / `.pdf`).
   Under **Save as**, tick **Excel files**, **PDF files** or both.
3. Click **Create files**. The table shows the progress of each file. At the end a summary appears and the folder
   opens (untick *Open the folder when finished* if you don't want that). The **Open:** links show the files right away.

### Tab names

The app looks for tabs named `ECG.certificate`, `NIBP.certificate`, `ECG.pass-fail` and `NIBP.pass-fail`.
Capital letters, spaces, dots and dashes don't matter, and extra words are fine. For example,
`ECG.pass-fail test sheet` and `NIBP.pass-fail sheet` are found too.
If your workbooks use different names, click **Options…** and type them in. The app remembers them.

Files with a missing tab are still processed: the tabs that exist are included, and the app tells you what was left out.

---

## Troubleshooting

| Problem | What to do |
| --- | --- |
| *"PDF files were not made"* | Your Excel can't save PDFs (e.g. Excel 2007). Install LibreOffice (free), or use the Excel files: open one and choose *Print → Print entire workbook*. |
| A tab shows ✖ | Select the file: the line under the table lists the tabs it has. Fix the name under **Options…**. |
| *"Could not save … is it open?"* | Close the file in Excel / the PDF viewer and click **Create files** again. |
| Page layout is wrong (cut off, several pages) | Fix the tab's *Page Layout* (print area, *Fit to 1 page*) in the original Excel file; the app keeps that setup. |

---

## For developers

Run from source (Windows, macOS or Linux, Python 3.9+):

```bash
pip install -r requirements.txt
python -m certpdf            # or double-click CertificatePDFBuilder.pyw on Windows
```

On macOS/Linux, PDFs are made with LibreOffice. Linux needs the `python3-tk` package.

**Build the Windows .exe yourself:** on a Windows PC with Python installed, double-click `build_windows.bat`.
The program appears in `dist\CertificatePDFBuilder.exe`.
GitHub Actions does the same on every push (`.github/workflows/build.yml`).

**Update the download:** the Download button serves `download/CertificatePDFBuilder.zip` from this repository.
After changing the app, download the `CertificatePDFBuilder-windows` artifact from the latest successful Actions run.
It holds the .exe and *How to use.txt*. Put both in a `CertificatePDFBuilder` folder, zip that folder, and replace
`download/CertificatePDFBuilder.zip`. Optionally, publishing a GitHub release also gets the zip attached to it automatically.

**Tests:** `pip install -r requirements-dev.txt` then `python -m pytest`. The end-to-end tests run only when LibreOffice is installed.

| Path | What it is |
| --- | --- |
| `certpdf/gui.py` | The window (tkinter) |
| `certpdf/excel_output.py` | Copies the tabs of every workbook into the two Excel files |
| `certpdf/sheet_copy.py` | Copies one tab (values, formatting, page setup) into another workbook |
| `certpdf/builder.py` | Turns the tabs of every workbook into PDF and merges them into the two PDFs |
| `certpdf/engines.py` | Saving a tab as PDF with Microsoft Excel (COM, never printing) or LibreOffice (headless) |
| `certpdf/workbook_info.py` | Reading tab names and matching them without opening Excel |
| `certpdf/config.py` | Tab definitions and saved settings |
| `tools/build_exe.py` | PyInstaller build of the single-file .exe |
| `tools/make_icon.py` | Draws the app icon |
