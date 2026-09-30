"""Convert old .xls workbooks to .xlsx with LibreOffice, when it is installed.

Only needed for .xls files: .xlsx and .xlsm files are read directly.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"


class ConversionError(Exception):
    """LibreOffice could not start or could not convert a file."""


class LibreOffice:
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
            raise ConversionError("LibreOffice is not installed")
        self._tmp = Path(tempfile.mkdtemp(prefix="certxl-lo-"))

    def stop(self) -> None:
        if self._tmp is not None:
            shutil.rmtree(self._tmp, ignore_errors=True)
            self._tmp = None

    def convert_to_xlsx(self, path: Path) -> Path:
        """Convert an old .xls workbook to .xlsx; the result lives until stop()."""
        work = Path(tempfile.mkdtemp(dir=self._tmp))
        shutil.copyfile(path, work / "source.xls")
        self._run(["--convert-to", "xlsx", "--outdir", str(work), str(work / "source.xls")], timeout=180)
        converted = work / "source.xlsx"
        if not converted.exists():
            raise ConversionError("LibreOffice could not read this .xls file")
        return converted

    def _run(self, args: list[str], timeout: float) -> str:
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
            raise ConversionError("LibreOffice took too long and was stopped") from None
        except OSError as exc:
            raise ConversionError(f"LibreOffice could not be started ({exc})") from None
        return result.stdout.decode("utf-8", "replace")


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
