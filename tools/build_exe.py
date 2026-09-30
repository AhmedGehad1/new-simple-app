"""Build the stand-alone program with PyInstaller.

    python tools/build_exe.py

On Windows this creates dist/CertificatePDFBuilder.exe, a single file that
runs on any Windows 10/11 PC without installing Python.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from certpdf import APP_ID, APP_NAME, __version__  # noqa: E402

VERSION_INFO = """\
VSVersionInfo(
  ffi=FixedFileInfo(filevers={nums}, prodvers={nums}, mask=0x3f, flags=0x0, OS=0x40004,
                    fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('040904B0', [
      StringStruct('FileDescription', '{name}'),
      StringStruct('FileVersion', '{version}'),
      StringStruct('InternalName', '{app_id}'),
      StringStruct('OriginalFilename', '{app_id}.exe'),
      StringStruct('ProductName', '{name}'),
      StringStruct('ProductVersion', '{version}')])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""


def main() -> None:
    import PyInstaller.__main__

    build = ROOT / "build"
    build.mkdir(exist_ok=True)
    args = [
        str(ROOT / "CertificatePDFBuilder.pyw"),
        "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", APP_ID,
        "--icon", str(ROOT / "assets" / "app.ico"),
        "--distpath", str(ROOT / "dist"),
        "--workpath", str(build / "pyinstaller"),
        "--specpath", str(build),
    ]
    if sys.platform == "win32":
        nums = tuple(int(part) for part in __version__.split("."))
        nums = (nums + (0, 0, 0, 0))[:4]
        version_file = build / "version_info.txt"
        version_file.write_text(VERSION_INFO.format(nums=nums, name=APP_NAME, version=__version__, app_id=APP_ID),
                                encoding="utf-8")
        args += ["--version-file", str(version_file), "--hidden-import", "win32timezone"]
    PyInstaller.__main__.run(args)


if __name__ == "__main__":
    main()
