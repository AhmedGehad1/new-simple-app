"""Start the application:  python -m certpdf"""

from __future__ import annotations

import sys

from . import APP_ID


def _prepare_windows() -> None:
    import ctypes
    try:
        # Sharp text on high-resolution screens instead of blurry, stretched pixels.
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass
    try:
        # Show this app's own icon in the taskbar.
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except (AttributeError, OSError):
        pass


def _missing_modules() -> list[str]:
    """Modules the packaged app must contain (checked by --smoke-test)."""
    names = ["pypdf", "xlrd"]
    if sys.platform == "win32":
        names += ["pythoncom", "win32com.client"]
    missing = []
    for name in names:
        try:
            __import__(name)
        except ImportError:
            missing.append(name)
    return missing


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if sys.platform == "win32":
        _prepare_windows()

    from .gui import App

    app = App()
    if "--smoke-test" in argv:
        # Used by the build to check that the packaged app starts; closes by itself.
        app.after(2000, app.destroy)
        app.mainloop()
        return 1 if _missing_modules() else 0
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
