import json

import pytest

from certpdf import config, engines
from certpdf.config import ENGINE_AUTO, ENGINE_EXCEL, ENGINE_LIBREOFFICE, Settings


class FakeEngine:
    fail = False
    available = True
    started = stopped = 0

    def __init__(self):
        pass

    @classmethod
    def is_available(cls):
        return cls.available

    def start(self):
        if self.fail:
            raise engines.EngineError("broken")
        type(self).started += 1

    def stop(self):
        type(self).stopped += 1


def make_engine(name, fail=False, available=True):
    return type(name, (FakeEngine,), {"name": name, "fail": fail, "available": available, "started": 0, "stopped": 0})


def test_auto_prefers_excel(monkeypatch):
    excel, libre = make_engine("Excel"), make_engine("LibreOffice")
    monkeypatch.setattr(engines, "ENGINE_CLASSES", {ENGINE_EXCEL: excel, ENGINE_LIBREOFFICE: libre})
    with engines.open_engine(ENGINE_AUTO) as engine:
        assert isinstance(engine, excel)
    assert excel.stopped == 1 and libre.started == 0


def test_auto_falls_back_to_libreoffice(monkeypatch):
    excel, libre = make_engine("Excel", fail=True), make_engine("LibreOffice")
    monkeypatch.setattr(engines, "ENGINE_CLASSES", {ENGINE_EXCEL: excel, ENGINE_LIBREOFFICE: libre})
    logs = []
    with engines.open_engine(ENGINE_AUTO, lambda level, text: logs.append((level, text))) as engine:
        assert isinstance(engine, libre)
    assert logs and logs[0][0] == "warning" and "Excel" in logs[0][1]
    assert libre.stopped == 1


def test_engine_is_stopped_after_an_error(monkeypatch):
    libre = make_engine("LibreOffice")
    monkeypatch.setattr(engines, "ENGINE_CLASSES", {ENGINE_EXCEL: make_engine("Excel", available=False),
                                                    ENGINE_LIBREOFFICE: libre})
    with pytest.raises(RuntimeError):
        with engines.open_engine(ENGINE_AUTO):
            raise RuntimeError("boom")
    assert libre.stopped == 1


def test_no_engine(monkeypatch):
    monkeypatch.setattr(engines, "ENGINE_CLASSES", {ENGINE_EXCEL: make_engine("Excel", available=False),
                                                    ENGINE_LIBREOFFICE: make_engine("LO", available=False)})
    with pytest.raises(engines.EngineError, match="Neither"):
        with engines.open_engine(ENGINE_AUTO):
            pass


def test_explicit_engine_failure(monkeypatch):
    monkeypatch.setattr(engines, "ENGINE_CLASSES", {ENGINE_EXCEL: make_engine("Excel", fail=True),
                                                    ENGINE_LIBREOFFICE: make_engine("LO")})
    with pytest.raises(engines.EngineError, match="broken"):
        with engines.open_engine(ENGINE_EXCEL):
            pass


def test_settings_round_trip(tmp_path, monkeypatch):
    monkeypatch.setattr(Settings, "path", classmethod(lambda cls: tmp_path / "s" / "settings.json"))
    settings = Settings.load()
    assert settings.sheet_names == config.default_sheet_names()
    settings.output_dir = "C:/Reports"
    settings.sheet_names["ecg_cert"] = "ECG Cert"
    settings.engine = ENGINE_LIBREOFFICE
    settings.save()

    loaded = Settings.load()
    assert loaded.output_dir == "C:/Reports"
    assert loaded.sheet_names["ecg_cert"] == "ECG Cert"
    assert loaded.engine == ENGINE_LIBREOFFICE


def test_settings_ignore_bad_values(tmp_path, monkeypatch):
    path = tmp_path / "settings.json"
    monkeypatch.setattr(Settings, "path", classmethod(lambda cls: path))
    path.write_text(json.dumps({"engine": "word", "sheet_names": {"ecg_cert": "", "junk": "x"},
                                "open_folder_when_done": "yes", "unknown": 1}))
    loaded = Settings.load()
    assert loaded.engine == ENGINE_AUTO
    assert loaded.sheet_names == config.default_sheet_names()
    assert loaded.open_folder_when_done is True

    path.write_text("{not json")
    assert Settings.load() == Settings()
