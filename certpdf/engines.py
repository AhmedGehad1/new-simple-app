"""Print a worksheet tab to PDF with Microsoft Excel or LibreOffice.

Excel gives exactly the same result as File > Print in Excel, so it is used
whenever it is installed. LibreOffice (free) is the fallback on PCs without
Excel, and the only option on macOS and Linux.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Callable, Iterator

from openpyxl import Workbook

from .config import ENGINE_AUTO, ENGINE_EXCEL, ENGINE_LIBREOFFICE
from .sheet_copy import adopt_default_font, copy_sheet, load_values
from .workbook_info import auto_height_rows

LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"


class EngineError(Exception):
    """A conversion engine could not start or failed on a workbook."""


# --------------------------------------------------------------------------
# Microsoft Excel (Windows)
# --------------------------------------------------------------------------
#
# Only ExportAsFixedFormat ("Save as PDF") is used. Nothing is ever sent to a
# printer: Excel 2007 sends print jobs to the real printer even when told to
# use "Microsoft Print to PDF".

XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_SHEET_VISIBLE = -1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


class EngineUnusable(EngineError):
    """The engine cannot make PDFs on this PC at all."""


def _com_message(exc: Exception) -> str:
    """Readable text from a pywintypes.com_error, including Excel's error code."""
    args = getattr(exc, "args", ())
    text, code = "", None
    try:
        excepinfo = args[2] if len(args) > 2 else None
        if excepinfo:
            text = str(excepinfo[2] or "").strip()
            code = excepinfo[5] if len(excepinfo) > 5 else None
        if not text and len(args) > 1 and args[1]:
            text = str(args[1]).strip()
        if not code and args and isinstance(args[0], int):
            code = args[0]
    except (IndexError, TypeError):
        pass
    text = text or str(exc) or type(exc).__name__
    if isinstance(code, int) and code:
        text += f" (code 0x{code & 0xFFFFFFFF:08X})"
    return text


def _copy_file_contents(src: Path, dst: Path) -> None:
    """Copy only the file's data, so Windows' "downloaded from the internet" mark
    (which makes Excel open files in Protected View) is not copied along."""
    with open(src, "rb") as reader, open(dst, "wb") as writer:
        shutil.copyfileobj(reader, writer, 1024 * 1024)


def _has_content(path: Path) -> bool:
    return path.exists() and path.stat().st_size > 0


class ExcelEngine:
    name = "Microsoft Excel"

    def __init__(self) -> None:
        self._app = None
        self._com_ready = False
        self._tmp: Path | None = None

    @staticmethod
    def is_available() -> bool:
        if sys.platform != "win32":
            return False
        try:
            import win32com.client  # noqa: F401
            import winreg
        except ImportError:
            return False
        try:
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"))
            return True
        except OSError:
            return False

    def start(self) -> None:
        try:
            import pythoncom
            import win32com.client
        except ImportError:
            raise EngineError("Excel automation is only available on Windows") from None
        pythoncom.CoInitialize()
        self._com_ready = True
        try:
            # DispatchEx starts a private Excel so the user's open workbooks are left alone.
            app = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            self.stop()
            raise EngineError(f"Excel could not be started ({_com_message(exc)})") from None
        for attr, value in (
            ("Visible", False),
            ("DisplayAlerts", False),
            ("ScreenUpdating", False),
            ("EnableEvents", False),
            ("AskToUpdateLinks", False),
            ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),  # never run macros
        ):
            try:
                setattr(app, attr, value)
            except Exception:
                pass
        self._app = app
        self._tmp = Path(tempfile.mkdtemp(prefix="certpdf-xl-"))
        try:
            self.check_pdf_export()
        except EngineUnusable:
            self.stop()
            raise

    def check_pdf_export(self) -> None:
        """Save a scratch workbook as PDF to find out whether this Excel can make PDFs.

        Excel 2007 can't without Microsoft's separate "Save as PDF" add-in.
        """
        probe = self._tmp / "pdf-check.pdf"
        book = None
        try:
            book = self._app.Workbooks.Add()
            sheet = book.Worksheets(1)
            sheet.Range("A1").Value = "PDF check"
            sheet.ExportAsFixedFormat(XL_TYPE_PDF, str(probe), XL_QUALITY_STANDARD, False, False)
            if not _has_content(probe):
                raise EngineError("no PDF was created")
        except Exception as exc:
            reason = str(exc) if isinstance(exc, EngineError) else _com_message(exc)
            raise EngineUnusable(
                f"this version of Excel cannot save PDFs ({reason}). Excel 2007 needs Microsoft's "
                "“Save as PDF” add-in, which is no longer available") from None
        finally:
            if book is not None:
                try:
                    book.Close(SaveChanges=False)
                except Exception:
                    pass

    def stop(self) -> None:
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None
        if self._com_ready:
            import pythoncom
            pythoncom.CoUninitialize()
            self._com_ready = False
        if self._tmp is not None:
            shutil.rmtree(self._tmp, ignore_errors=True)
            self._tmp = None

    @contextmanager
    def open(self, path: Path) -> Iterator["_ExcelWorkbook"]:
        # Excel works on a local copy: files on network shares, files marked as
        # downloaded, and files someone else has open all cause export errors.
        # The copy keeps the file name, which formulas such as CELL("filename") show.
        work = Path(tempfile.mkdtemp(dir=self._tmp))
        local = work / Path(path).name
        try:
            try:
                _copy_file_contents(Path(path), local)
            except OSError as exc:
                raise EngineError(f"the file could not be read ({exc.strerror or exc})") from None
            try:
                workbook = self._app.Workbooks.Open(
                    str(local),
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True,
                    # A dummy password makes Excel fail on protected files instead of
                    # waiting for someone to type the password into a hidden window.
                    Password="-",
                    Notify=False,
                    AddToMru=False,
                )
            except Exception as exc:
                raise EngineError(f"Excel could not open the file ({_com_message(exc)})") from None
            try:
                yield _ExcelWorkbook(workbook)
            finally:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
        finally:
            shutil.rmtree(work, ignore_errors=True)


class _ExcelWorkbook:
    def __init__(self, workbook) -> None:
        self._workbook = workbook
        self.sheet_names = [sheet.Name for sheet in workbook.Worksheets]

    def export_many(self, jobs: list[tuple[str, Path]]) -> dict[str, str]:
        errors = {}
        for sheet_name, pdf_path in jobs:
            pdf_path = Path(pdf_path)
            try:
                sheet = self._workbook.Worksheets(sheet_name)
                if sheet.Visible != XL_SHEET_VISIBLE:
                    try:
                        sheet.Visible = XL_SHEET_VISIBLE  # hidden tabs cannot be exported
                    except Exception:
                        pass
                # Same pages as printing the tab (print area, scaling, footer...), saved as a file.
                sheet.ExportAsFixedFormat(XL_TYPE_PDF, str(pdf_path), XL_QUALITY_STANDARD, True, False)
                if not _has_content(pdf_path):
                    raise EngineError("Excel did not create the PDF (is the tab empty?)")
            except EngineError as exc:
                errors[sheet_name] = str(exc)
            except Exception as exc:
                errors[sheet_name] = _com_message(exc)
        return errors


# --------------------------------------------------------------------------
# LibreOffice (Windows, macOS, Linux)
# --------------------------------------------------------------------------

class LibreOfficeEngine:
    name = "LibreOffice"

    def __init__(self, soffice: str | None = None) -> None:
        self.soffice = soffice
        self._tmp: Path | None = None

    @staticmethod
    def find() -> str | None:
        candidates: list[Path] = []
        override = os.environ.get("CERTPDF_SOFFICE")
        if override:
            candidates.append(Path(override))
        if sys.platform == "win32":
            candidates += _windows_soffice_candidates()
        elif sys.platform == "darwin":
            for base in (Path("/Applications"), Path.home() / "Applications"):
                candidates.append(base / "LibreOffice.app" / "Contents" / "MacOS" / "soffice")
        for name in ("soffice", "libreoffice"):
            found = shutil.which(name)
            if found:
                candidates.append(Path(found))
        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)
        return None

    @classmethod
    def is_available(cls) -> bool:
        return cls.find() is not None

    def start(self) -> None:
        self.soffice = self.soffice or self.find()
        if not self.soffice:
            raise EngineError("LibreOffice is not installed")
        self._tmp = Path(tempfile.mkdtemp(prefix="certpdf-lo-"))

    def stop(self) -> None:
        if self._tmp is not None:
            shutil.rmtree(self._tmp, ignore_errors=True)
            self._tmp = None

    @contextmanager
    def open(self, path: Path) -> Iterator["_LibreOfficeWorkbook"]:
        path = Path(path).resolve()
        work = Path(tempfile.mkdtemp(dir=self._tmp))
        try:
            source = self.convert_to_xlsx(path, work) if path.suffix.lower() == ".xls" else path
            try:
                book = load_values(source)
            except Exception as exc:
                raise EngineError(f"the file could not be read ({exc or type(exc).__name__})") from None
            yield _LibreOfficeWorkbook(self, book, auto_height_rows(source, book.sheetnames), work)
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def convert_to_xlsx(self, path: Path, work: Path | None = None) -> Path:
        """Convert an old .xls workbook to .xlsx; the result lives until stop()."""
        work = work or Path(tempfile.mkdtemp(dir=self._tmp))
        shutil.copyfile(path, work / "source.xls")
        self.run(["--convert-to", "xlsx", "--outdir", str(work), str(work / "source.xls")], timeout=180)
        converted = work / "source.xlsx"
        if not converted.exists():
            raise EngineError("LibreOffice could not read this .xls file")
        return converted

    def run(self, args: list[str], timeout: float) -> str:
        profile = (self._tmp / "profile").as_uri()
        command = [
            self.soffice, f"-env:UserInstallation={profile}",
            "--headless", "--invisible", "--nologo", "--norestore", "--nodefault", "--nolockcheck",
            *args,
        ]
        kwargs = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        try:
            result = subprocess.run(command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                                    stderr=subprocess.STDOUT, timeout=timeout, **kwargs)
        except subprocess.TimeoutExpired:
            raise EngineError("LibreOffice took too long and was stopped") from None
        except OSError as exc:
            raise EngineError(f"LibreOffice could not be started ({exc})") from None
        return result.stdout.decode("utf-8", "replace")


class _LibreOfficeWorkbook:
    """Converts tabs through one-tab copies holding the values Excel saved.

    Converting the original workbook would not work: LibreOffice prints hidden
    tabs that have a print area, and it recalculates formulas (a certificate
    number taken from the file name would come out wrong).
    """

    def __init__(self, engine: LibreOfficeEngine, book, auto_rows: dict[str, set[int]], work: Path) -> None:
        self._engine = engine
        self._book = book
        self._auto_rows = auto_rows
        self._work = work
        self.sheet_names = list(book.sheetnames)

    def export_many(self, jobs: list[tuple[str, Path]]) -> dict[str, str]:
        errors: dict[str, str] = {}
        copies = []
        for number, (sheet_name, pdf_path) in enumerate(jobs):
            copy = self._work / f"tab{number}.xlsx"
            try:
                single = Workbook()
                adopt_default_font(single, self._book)
                single.active.title = sheet_name
                # Rows Excel sized itself are left for LibreOffice to size, as with the original.
                copy_sheet(self._book[sheet_name], single.active, self._auto_rows.get(sheet_name, set()))
                single.save(copy)
            except Exception as exc:
                errors[sheet_name] = f"could not prepare the tab ({exc or type(exc).__name__})"
                continue
            copies.append((sheet_name, copy, Path(pdf_path)))
        if not copies:
            return errors

        out_dir = self._work / "pdf"
        # One LibreOffice run prints all the tabs of this workbook.
        self._engine.run(["--convert-to", "pdf", "--outdir", str(out_dir), *(str(c) for _, c, _ in copies)],
                         timeout=120 + 60 * len(copies))
        for sheet_name, copy, pdf_path in copies:
            produced = out_dir / f"{copy.stem}.pdf"
            if produced.exists():
                shutil.move(str(produced), str(pdf_path))
            else:
                errors[sheet_name] = "LibreOffice did not create the PDF"
        return errors


def _windows_soffice_candidates() -> list[Path]:
    folders: list[Path] = []
    try:
        import winreg
        for view in (0, getattr(winreg, "KEY_WOW64_32KEY", 0)):
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\LibreOffice\UNO\InstallPath",
                                    0, winreg.KEY_READ | view) as key:
                    folders.append(Path(winreg.QueryValueEx(key, "")[0]))
            except OSError:
                pass
    except ImportError:
        pass
    for var in ("ProgramFiles", "ProgramW6432", "ProgramFiles(x86)"):
        root = os.environ.get(var)
        if root:
            folders += [p / "program" for p in sorted(Path(root).glob("LibreOffice*"), reverse=True)]
    local = os.environ.get("LOCALAPPDATA")
    if local:
        folders.append(Path(local) / "Programs" / "LibreOffice" / "program")
    # soffice.com waits for the conversion to finish; soffice.exe is the fallback.
    return [folder / exe for folder in folders for exe in ("soffice.com", "soffice.exe")]


# --------------------------------------------------------------------------
# Choosing an engine
# --------------------------------------------------------------------------

ENGINE_CLASSES = {ENGINE_EXCEL: ExcelEngine, ENGINE_LIBREOFFICE: LibreOfficeEngine}


def available_engines() -> dict[str, bool]:
    return {key: cls.is_available() for key, cls in ENGINE_CLASSES.items()}


@contextmanager
def open_engine(preference: str, log: Callable[[str, str], None] = lambda level, text: None):
    """Start the preferred engine; in automatic mode fall back from Excel to LibreOffice."""
    if preference == ENGINE_AUTO:
        keys = [key for key in (ENGINE_EXCEL, ENGINE_LIBREOFFICE) if ENGINE_CLASSES[key].is_available()]
        if not keys:
            raise EngineError("Neither Microsoft Excel nor LibreOffice is installed on this PC")
    else:
        keys = [preference]

    errors = []
    for key in keys:
        engine = ENGINE_CLASSES[key]()
        try:
            engine.start()
        except EngineError as exc:
            errors.append(f"{engine.name}: {exc}")
            log("warning", f"{engine.name} could not be used: {exc}")
            continue
        try:
            yield engine
        finally:
            engine.stop()
        return
    raise EngineError("; ".join(errors) or "no conversion engine available")
