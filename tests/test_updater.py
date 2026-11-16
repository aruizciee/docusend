"""Tests del actualizador (sin red y sin reiniciar nada)."""
import hashlib
import types

import pytest

import auto_updater


def release(**asset):
    base = {"name": "docusend.exe", "browser_download_url": "https://example.invalid/docusend.exe"}
    base.update(asset)
    return {"tag_name": "v2026.11.16-2", "assets": [{"name": "otro.txt"}, base]}


class TestFindAsset:
    def test_with_digest(self):
        found = auto_updater._find_asset(release(digest="sha256:ABCDEF", size=123))
        assert found == {"url": "https://example.invalid/docusend.exe", "sha256": "abcdef", "size": 123}

    def test_without_digest(self):
        found = auto_updater._find_asset(release())
        assert found["sha256"] == "" and found["size"] == 0

    def test_missing_executable(self):
        assert auto_updater._find_asset({"assets": [{"name": "otro.txt"}]}) is None


class TestVerifyDownload:
    @pytest.fixture
    def exe(self, tmp_path):
        path = tmp_path / "docusend.exe.new"
        path.write_bytes(b"x" * (2 * 1024 * 1024))
        return path

    def test_ok(self, exe):
        digest = hashlib.sha256(exe.read_bytes()).hexdigest()
        auto_updater._verify_download(str(exe), digest.upper(), exe.stat().st_size)

    def test_wrong_hash(self, exe):
        with pytest.raises(ValueError, match="SHA-256"):
            auto_updater._verify_download(str(exe), "0" * 64, exe.stat().st_size)

    def test_incomplete(self, exe):
        with pytest.raises(ValueError, match="incompleta"):
            auto_updater._verify_download(str(exe), "", exe.stat().st_size + 10)

    def test_too_small_to_be_the_app(self, tmp_path):
        path = tmp_path / "pagina_de_error.html"
        path.write_bytes(b"<html>error</html>")
        with pytest.raises(ValueError, match="ejecutable"):
            auto_updater._verify_download(str(path))


class TestRestart:
    def test_environment_does_not_leak_pyinstaller_state(self, monkeypatch):
        # Estas variables hacen que la app reiniciada busque python3xx.dll en la
        # carpeta temporal de la version anterior ("Failed to load Python DLL").
        monkeypatch.setenv("_PYI_APPLICATION_HOME_DIR", r"C:\Temp\_MEI123")
        monkeypatch.setenv("_PYI_ARCHIVE_FILE", r"C:\app\docusend.exe")
        monkeypatch.setenv("_MEIPASS2", r"C:\Temp\_MEI123")
        monkeypatch.setenv("OTRA_VARIABLE", "se conserva")
        env = auto_updater._clean_environment()
        assert not [k for k in env if k.upper().startswith(("_PYI_", "_MEIPASS"))]
        assert env["PYINSTALLER_RESET_ENVIRONMENT"] == "1"
        assert env["OTRA_VARIABLE"] == "se conserva"

    def test_script(self):
        new = r"C:\Users\José\AppData\Local\DocuSend\docusend.exe.new"
        original = r"C:\Users\José\OneDrive - CIEE\Descargas\docusend.exe"
        script = auto_updater._build_updater_script(4321, new, original)
        lines = script.splitlines()
        assert "chcp 65001 >nul" in lines
        assert "set PYINSTALLER_RESET_ENVIRONMENT=1" in lines
        assert "set _PYI_APPLICATION_HOME_DIR=" in lines
        assert f'move /y "{new}" "{original}" >nul 2>&1' in lines
        assert f'start "" "{original}"' in lines
        # Nunca se borra el ejecutable actual antes de tener el nuevo en su sitio
        assert not any(line.startswith("del") and original in line for line in lines)
        # El entorno se limpia antes de arrancar la nueva version
        assert lines.index("set PYINSTALLER_RESET_ENVIRONMENT=1") < lines.index(f'start "" "{original}"')


class TestManualCheck:
    """El boton "Buscar actualizaciones" siempre responde algo."""

    @pytest.fixture
    def app(self):
        calls = []
        fake = types.SimpleNamespace(after=lambda _ms, fn: fn(), calls=calls)
        return fake

    @pytest.fixture
    def boxes(self, monkeypatch):
        shown = []
        for kind in ("showinfo", "showwarning", "askyesno"):
            monkeypatch.setattr(auto_updater.messagebox, kind,
                                lambda *a, _k=kind, **kw: shown.append((_k, a)) or False)
        return shown

    def _respond(self, monkeypatch, payload=None, error=None):
        def fake_get(*_a, **_k):
            if error:
                raise error
            return types.SimpleNamespace(raise_for_status=lambda: None, json=lambda: payload)
        monkeypatch.setattr(auto_updater.requests, "get", fake_get)

    def test_up_to_date(self, app, boxes, monkeypatch):
        monkeypatch.setattr(auto_updater, "VERSION", "v2026.11.16-2")
        self._respond(monkeypatch, release())
        auto_updater._check(app, manual=True)
        assert boxes and boxes[0][0] == "showinfo" and "ultima version" in boxes[0][1][1]

    def test_network_error(self, app, boxes, monkeypatch):
        self._respond(monkeypatch, error=OSError("sin red"))
        auto_updater._check(app, manual=True)
        assert boxes and boxes[0][0] == "showwarning"

    def test_automatic_check_is_silent(self, app, boxes, monkeypatch):
        monkeypatch.setattr(auto_updater, "VERSION", "v2026.11.16-2")
        self._respond(monkeypatch, release())
        auto_updater._check(app, manual=False)
        self._respond(monkeypatch, error=OSError("sin red"))
        auto_updater._check(app, manual=False)
        assert boxes == []

    def test_new_version_is_offered(self, app, boxes, monkeypatch):
        monkeypatch.setattr(auto_updater, "VERSION", "v2026.11.16-1")
        self._respond(monkeypatch, release(digest="sha256:" + "a" * 64))
        auto_updater._check(app, manual=False)
        assert boxes and boxes[0][0] == "askyesno" and "v2026.11.16-2" in boxes[0][1][1]
