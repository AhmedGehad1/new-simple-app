"""Sheet definitions and persisted user settings."""

from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path

from . import APP_ID


@dataclass(frozen=True)
class SheetSpec:
    key: str
    label: str          # shown in the UI
    default_name: str   # tab name looked for in each workbook
    group: str          # which output PDF the tab goes into


CERTIFICATES = "certificates"
PASS_FAIL = "passfail"

SHEET_SPECS = (
    SheetSpec("ecg_cert", "ECG certificate", "ECG.certificate", CERTIFICATES),
    SheetSpec("nibp_cert", "NIBP certificate", "NIBP.certificate", CERTIFICATES),
    SheetSpec("ecg_pf", "ECG pass/fail", "ECG.pass-fail", PASS_FAIL),
    SheetSpec("nibp_pf", "NIBP pass/fail", "NIBP.pass-fail", PASS_FAIL),
)
SPEC_BY_KEY = {spec.key: spec for spec in SHEET_SPECS}

# Order of the tabs inside each output PDF (per workbook).
GROUPS = {
    CERTIFICATES: ("ecg_cert", "nibp_cert"),
    PASS_FAIL: ("ecg_pf", "nibp_pf"),
}
GROUP_TITLES = {
    CERTIFICATES: "Certificates",
    PASS_FAIL: "Pass/Fail test sheets",
}

ENGINE_AUTO = "auto"
ENGINE_EXCEL = "excel"
ENGINE_LIBREOFFICE = "libreoffice"


def default_sheet_names() -> dict[str, str]:
    return {spec.key: spec.default_name for spec in SHEET_SPECS}


def app_data_dir() -> Path:
    """Per-user folder for settings and the LibreOffice profile."""
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / APP_ID


@dataclass
class Settings:
    output_dir: str = ""
    certificates_name: str = "Certificates.pdf"
    passfail_name: str = "Pass-Fail Test Sheets.pdf"
    sheet_names: dict = field(default_factory=default_sheet_names)
    engine: str = ENGINE_AUTO
    open_folder_when_done: bool = True
    last_browse_dir: str = ""

    @classmethod
    def path(cls) -> Path:
        return app_data_dir() / "settings.json"

    @classmethod
    def load(cls) -> "Settings":
        settings = cls()
        try:
            data = json.loads(cls.path().read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return settings
        known = {f.name for f in fields(cls)}
        for name, value in data.items():
            if name in known and isinstance(value, type(getattr(settings, name))):
                setattr(settings, name, value)
        # Keep only known sheet keys and fill in any that are missing.
        names = default_sheet_names()
        names.update({k: v for k, v in settings.sheet_names.items()
                      if k in names and isinstance(v, str) and v.strip()})
        settings.sheet_names = names
        if settings.engine not in (ENGINE_AUTO, ENGINE_EXCEL, ENGINE_LIBREOFFICE):
            settings.engine = ENGINE_AUTO
        return settings

    def save(self) -> None:
        try:
            path = self.path()
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(asdict(self), indent=2), encoding="utf-8")
        except OSError:
            pass  # settings are a convenience; never fail because of them
