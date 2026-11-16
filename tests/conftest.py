"""Configuración común de los tests.

DocuSend es una app de Windows (Word, Outlook, AutoFirma). Para poder ejecutar
los tests en cualquier sistema se sustituyen los módulos de Windows por dobles
cuando no están instalados; en Windows se usan los reales.
"""
import ctypes
import os
import sys
import types

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import win32com.client  # noqa: F401
except ImportError:
    _client = types.ModuleType("win32com.client")
    _client.Dispatch = lambda *a, **k: None
    _pkg = types.ModuleType("win32com")
    _pkg.client = _client
    sys.modules["win32com"] = _pkg
    sys.modules["win32com.client"] = _client

try:
    import pythoncom  # noqa: F401
except ImportError:
    _pc = types.ModuleType("pythoncom")
    _pc.CoInitialize = _pc.CoUninitialize = lambda: None
    sys.modules["pythoncom"] = _pc

if not hasattr(ctypes, "windll"):
    ctypes.windll = types.SimpleNamespace(shell32=types.SimpleNamespace(
        SetCurrentProcessExplicitAppUserModelID=lambda _id: None))


class FakeMail:
    def __init__(self, log):
        self._log = log
        self.Subject = self.HTMLBody = self.To = self.CC = self.BCC = ""
        self.SendUsingAccount = None
        self.attachments = []

    @property
    def Attachments(self):
        return types.SimpleNamespace(Add=self.attachments.append)

    def Send(self):
        self._log["sent"].append(self)

    def Save(self):
        self._log["saved"].append(self)

    def Display(self):
        self._log["shown"].append(self)


class FakeOutlook:
    """Outlook de mentira: registra lo que se enviaría, guardaría o mostraría."""
    def __init__(self, accounts=()):
        self.log = {"sent": [], "saved": [], "shown": []}
        items = [types.SimpleNamespace(DisplayName=n, SmtpAddress=a) for n, a in accounts]
        self.Session = types.SimpleNamespace(Accounts=types.SimpleNamespace(
            Count=len(items), Item=lambda i: items[i - 1]))

    def CreateItem(self, _kind):
        return FakeMail(self.log)


@pytest.fixture
def outlook(monkeypatch):
    import docusend
    fake = FakeOutlook(accounts=[("Ana", "ana@x.org"), ("Ana Maria", "anamaria@x.org")])
    monkeypatch.setattr(docusend.win32, "Dispatch", lambda *a, **k: fake)
    return fake


@pytest.fixture
def dialogs(monkeypatch):
    """Sustituye los cuadros de diálogo y registra los que se habrían mostrado."""
    from tkinter import messagebox
    shown = []
    for kind in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(messagebox, kind,
                            lambda *a, _k=kind, **kw: shown.append((_k, a)))
    answers = {"yes": True}
    monkeypatch.setattr(messagebox, "askyesno",
                        lambda *a, **kw: (shown.append(("askyesno", a)), answers["yes"])[1])
    shown_answers = types.SimpleNamespace(shown=shown, answers=answers)
    return shown_answers


@pytest.fixture
def app(tmp_path, monkeypatch, dialogs):
    """La aplicación real, sin configuración previa. Se omite si no hay pantalla."""
    import tkinter
    import docusend
    monkeypatch.setattr(docusend, "CONFIG_PATH", str(tmp_path / "config.json"))
    monkeypatch.setattr(docusend, "check_for_updates", lambda _app: None)
    monkeypatch.setattr(docusend.App, "_load_outlook_accounts", lambda self: None)
    opened = []
    monkeypatch.setattr(docusend.App, "_open_path", lambda self, p: opened.append(p))
    try:
        application = docusend.App()
    except tkinter.TclError as exc:
        pytest.skip(f"No hay entorno gráfico disponible: {exc}")
    application.opened = opened
    application.dialogs = dialogs
    yield application
    try:
        application.destroy()
    except tkinter.TclError:
        pass
