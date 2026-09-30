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
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Callable, Iterator

from .config import ENGINE_AUTO, ENGINE_EXCEL, ENGINE_LIBREOFFICE
from .workbook_info import WorkbookReadError, isolate_sheet, read_sheet_names

LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"


class EngineError(Exception):
    """A conversion engine could not start or failed on a workbook."""


# --------------------------------------------------------------------------
# Microsoft Excel (Windows)
# --------------------------------------------------------------------------

XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_SHEET_VISIBLE = -1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
PDF_PRINTER = "Microsoft Print to PDF"


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


def _pdf_printer_names() -> list[str]:
    """Names Excel may accept for Windows' built-in PDF printer."""
    names = [PDF_PRINTER]
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows NT\CurrentVersion\Devices") as key:
            port = str(winreg.QueryValueEx(key, PDF_PRINTER)[0]).split(",")[-1]
        names.append(f"{PDF_PRINTER} on {port}")
    except (ImportError, OSError, IndexError):
        pass
    return names


def _has_content(path: Path) -> bool:
    return path.exists() and path.stat().st_size > 0


def _wait_for_pdf(path: Path, timeout: float) -> bool:
    """Printing finishes in the background; wait until the PDF is fully written."""
    deadline = time.monotonic() + timeout
    last_size = -1
    while time.monotonic() < deadline:
        size = path.stat().st_size if path.exists() else -1
        if size > 0 and size == last_size:
            return True
        last_size = size
        time.sleep(0.5)
    return _has_content(path)


class ExcelEngine:
    name = "Microsoft Excel"

    def __init__(self) -> None:
        self._app = None
        self._com_ready = False
        self._tmp: Path | None = None
        self.log: Callable[[str, str], None] = lambda level, text: None
        # Ways to turn a tab into a PDF, best first. The first one that works
        # on this PC is moved to the front and used from then on.
        self.methods = [
            ("export", self._export),
            ("export from a copy of the tab", self._export_copy),
            (f"print to “{PDF_PRINTER}”", self._print_to_pdf_printer),
            ("export ignoring the print area", self._export_ignoring_print_area),
        ]

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
        work = Path(tempfile.mkdtemp(dir=self._tmp))
        local = work / f"workbook{Path(path).suffix.lower()}"
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
                yield _ExcelWorkbook(self, workbook)
            finally:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def convert(self, sheet, pdf_path: Path) -> None:
        """Turn one tab into a PDF, trying each method until one works."""
        failures = []
        for position, (label, method) in enumerate(self.methods):
            _remove_quietly(pdf_path)
            try:
                method(sheet, pdf_path)
                printed = method == self._print_to_pdf_printer
                if not (_wait_for_pdf(pdf_path, timeout=60) if printed else _has_content(pdf_path)):
                    raise EngineError("no PDF was created")
            except Exception as exc:
                message = str(exc) if isinstance(exc, EngineError) else _com_message(exc)
                failures.append(f"{label}: {message}")
                continue
            if position:
                self.methods.insert(0, self.methods.pop(position))
                self.log("warning" if "ignoring" in label else "info",
                         f"Excel's normal PDF export did not work on this PC; using “{label}” instead.")
            return
        raise EngineError("Excel could not convert it. " + " | ".join(failures))

    # Methods, best first -------------------------------------------------

    @staticmethod
    def _export(sheet, pdf_path: Path) -> None:
        # Same as printing the tab: uses its print area, orientation, scaling, headers...
        sheet.ExportAsFixedFormat(XL_TYPE_PDF, str(pdf_path), XL_QUALITY_STANDARD, True, False)

    def _export_copy(self, sheet, pdf_path: Path) -> None:
        # Copying a tab on its own into a new workbook keeps its page setup.
        count = self._app.Workbooks.Count
        sheet.Copy()
        if self._app.Workbooks.Count <= count:
            raise EngineError("the tab could not be copied")
        book = self._app.Workbooks(self._app.Workbooks.Count)
        try:
            book.ExportAsFixedFormat(XL_TYPE_PDF, str(pdf_path), XL_QUALITY_STANDARD, True, False)
        finally:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass

    @staticmethod
    def _print_to_pdf_printer(sheet, pdf_path: Path) -> None:
        errors = []
        for printer in _pdf_printer_names():
            try:
                sheet.PrintOut(ActivePrinter=printer, PrintToFile=True, PrToFileName=str(pdf_path))
                return
            except Exception as exc:
                errors.append(_com_message(exc))
        raise EngineError(f"the “{PDF_PRINTER}” printer is not available ({errors[-1]})")

    @staticmethod
    def _export_ignoring_print_area(sheet, pdf_path: Path) -> None:
        # Last resort for tabs whose print area is broken: prints the used range.
        sheet.ExportAsFixedFormat(XL_TYPE_PDF, str(pdf_path), XL_QUALITY_STANDARD, True, True)


class _ExcelWorkbook:
    def __init__(self, engine: ExcelEngine, workbook) -> None:
        self._engine = engine
        self._workbook = workbook
        self.sheet_names = [sheet.Name for sheet in workbook.Worksheets]

    def export_many(self, jobs: list[tuple[str, Path]]) -> dict[str, str]:
        errors = {}
        for sheet_name, pdf_path in jobs:
            try:
                sheet = self._workbook.Worksheets(sheet_name)
                if sheet.Visible != XL_SHEET_VISIBLE:
                    try:
                        sheet.Visible = XL_SHEET_VISIBLE  # hidden tabs cannot be printed
                    except Exception:
                        pass
                self._engine.convert(sheet, Path(pdf_path))
            except EngineError as exc:
                errors[sheet_name] = str(exc)
            except Exception as exc:
                errors[sheet_name] = _com_message(exc)
        return errors


def _remove_quietly(path: Path) -> None:
    try:
        os.remove(path)
    except OSError:
        pass


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
            suffix = path.suffix.lower()
            if suffix == ".xls":
                # Old-style workbooks are converted once so single tabs can be picked out.
                shutil.copyfile(path, work / "source.xls")
                self.run(["--convert-to", "xlsx", "--outdir", str(work), str(work / "source.xls")], timeout=180)
                source = work / "source.xlsx"
                if not source.exists():
                    raise EngineError("LibreOffice could not read this .xls file")
            else:
                source = path
            try:
                names = read_sheet_names(source)
            except WorkbookReadError as exc:
                raise EngineError(f"the file could not be read ({exc})") from None
            yield _LibreOfficeWorkbook(self, source, names, work)
        finally:
            shutil.rmtree(work, ignore_errors=True)

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
    def __init__(self, engine: LibreOfficeEngine, source: Path, sheet_names: list[str], work: Path) -> None:
        self._engine = engine
        self._source = source
        self._work = work
        self.sheet_names = sheet_names

    def export_many(self, jobs: list[tuple[str, Path]]) -> dict[str, str]:
        errors: dict[str, str] = {}
        copies = []
        for number, (sheet_name, pdf_path) in enumerate(jobs):
            copy = self._work / f"tab{number}{self._source.suffix.lower()}"
            try:
                isolate_sheet(self._source, copy, sheet_name)
            except (WorkbookReadError, OSError, KeyError, ValueError) as exc:
                errors[sheet_name] = f"could not prepare the tab ({exc})"
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
        engine.log = log
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
