import json

from certpdf import config
from certpdf.config import Settings


def test_settings_round_trip(tmp_path, monkeypatch):
    monkeypatch.setattr(Settings, "path", classmethod(lambda cls: tmp_path / "s" / "settings.json"))
    settings = Settings.load()
    assert settings.sheet_names == config.default_sheet_names()
    settings.output_dir = "C:/Reports"
    settings.sheet_names["ecg_cert"] = "ECG Cert"
    settings.save()

    loaded = Settings.load()
    assert loaded.output_dir == "C:/Reports"
    assert loaded.sheet_names["ecg_cert"] == "ECG Cert"


def test_settings_ignore_bad_values(tmp_path, monkeypatch):
    path = tmp_path / "settings.json"
    monkeypatch.setattr(Settings, "path", classmethod(lambda cls: path))
    path.write_text(json.dumps({"engine": "word", "sheet_names": {"ecg_cert": "", "junk": "x"},
                                "open_folder_when_done": "yes", "unknown": 1}))
    loaded = Settings.load()
    assert loaded.sheet_names == config.default_sheet_names()
    assert loaded.open_folder_when_done is True

    path.write_text("{not json")
    assert Settings.load() == Settings()


def test_settings_from_earlier_versions_are_migrated(tmp_path, monkeypatch):
    path = tmp_path / "settings.json"
    monkeypatch.setattr(Settings, "path", classmethod(lambda cls: path))
    path.write_text(json.dumps({"certificates_name": "Certificates.pdf", "passfail_name": "PF sheets.PDF",
                                "make_pdf": True, "engine": "excel"}))
    loaded = Settings.load()
    assert loaded.certificates_name == "Certificates"
    assert loaded.passfail_name == "PF sheets"
