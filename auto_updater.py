"""
Auto-updater para DocuSend.
Comprueba la ultima release en GitHub y, si hay version nueva,
pregunta al usuario si desea actualizar. La descarga e instala en
%LOCALAPPDATA%\DocuSend\ (carpeta local, fuera de OneDrive)
para evitar interferencias con la sincronizacion al cargar DLLs.
"""
import os
import sys
import subprocess
import threading
import requests
from tkinter import messagebox

try:
    from version import VERSION
except ImportError:
    VERSION = "v0.0.0-dev"

GITHUB_REPO  = "aruizciee/docusend"
GITHUB_API   = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
EXE_NAME     = "docusend.exe"


def _get_install_dir():
    """Devuelve %LOCALAPPDATA%\DocuSend, creandolo si no existe."""
    local_appdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    d = os.path.join(local_appdata, "DocuSend")
    os.makedirs(d, exist_ok=True)
    return d


def check_for_updates(app):
    """Lanza la comprobacion en un hilo daemon para no bloquear la UI."""
    threading.Thread(target=_check, args=(app,), daemon=True).start()


def _check(app):
    try:
        resp = requests.get(GITHUB_API, timeout=10)
        resp.raise_for_status()
        data   = resp.json()
        latest = data.get("tag_name", "")
        if not latest or _parse_version(latest) <= _parse_version(VERSION):
            return

        asset_url = _find_asset(data)
        if not asset_url:
            return

        app.after(0, lambda: _prompt_update(app, latest, asset_url))

    except Exception:
        pass


def _parse_version(tag):
    """Convierte 'v2026.03.23-4' -> (2026, 3, 23, 4) para comparacion numerica."""
    try:
        tag = tag.lstrip("v")
        date_part, _, build = tag.partition("-")
        parts = [int(x) for x in date_part.split(".")]
        parts.append(int(build) if build.isdigit() else 0)
        return tuple(parts)
    except Exception:
        return (0,)


def _find_asset(release_data):
    for asset in release_data.get("assets", []):
        if asset["name"].lower() == EXE_NAME:
            return asset["browser_download_url"]
    return None


def _prompt_update(app, latest, asset_url):
    answer = messagebox.askyesno(
        "Actualizacion disponible",
        f"Hay una nueva version disponible: {latest}\n"
        f"Version actual: {VERSION}\n\n"
        "Deseas actualizar ahora?\n"
        "(La aplicacion se cerrara y se reiniciara automaticamente.)"
    )
    if answer:
        threading.Thread(target=_download_and_restart, args=(app, asset_url), daemon=True).start()


def _download_and_restart(app, asset_url):
    """
    Descarga el nuevo .exe en un directorio temporal y lanza
    un script batch que reemplaza el ejecutable original y lo reinicia.
    """
    if not getattr(sys, "frozen", False):
        messagebox.showinfo("Dev mode", "Auto-update desactivado en modo desarrollo.")
        return

    original_exe = sys.executable
    temp_dir = _get_install_dir()
    new_exe  = os.path.join(temp_dir, EXE_NAME + ".new")
    bat_path = os.path.join(temp_dir, "_updater.bat")

    try:
        app.after(0, lambda: _show_downloading(app))

        resp = requests.get(asset_url, stream=True, timeout=120)
        resp.raise_for_status()
        with open(new_exe, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)

        pid = os.getpid()
        bat_content = (
            f"@echo off\n"
            f"setlocal enabledelayedexpansion\n"
            f":wait_process\n"
            f'tasklist /fi "PID eq {pid}" 2>nul | find /i "{pid}" >nul\n'
            f"if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait_process)\n"
            f"timeout /t 2 /nobreak >nul\n"
            f"set MAX_RETRIES=15\n"
            f"set RETRY_COUNT=0\n"
            f":retry_update\n"
            f'del /f /q "{original_exe}" >nul 2>&1\n'
            f'move /y "{new_exe}" "{original_exe}" >nul 2>&1\n'
            f'if exist "{new_exe}" (\n'
            f"    set /a RETRY_COUNT+=1\n"
            f"    if !RETRY_COUNT! geq !MAX_RETRIES! goto on_error\n"
            f"    timeout /t 2 /nobreak >nul\n"
            f"    goto retry_update\n"
            f")\n"
            f'start "" "{original_exe}"\n'
            f"goto end\n"
            f":on_error\n"
            f'mshta vbscript:Execute("msgbox ""Error: El archivo original de la app esta bloqueado (posiblemente por OneDrive o un antivirus). Intentalo de nuevo mas tarde."",16,""Error de actualizacion"":close")\n'
            f'start "" "{original_exe}"\n'
            f":end\n"
            f'del "%~f0"\n'
        )
        with open(bat_path, "w") as f:
            f.write(bat_content)

        subprocess.Popen(bat_path, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        app.after(0, app.destroy)

    except Exception as e:
        for path in (new_exe, bat_path):
            try:
                os.remove(path)
            except OSError:
                pass
        app.after(0, lambda: messagebox.showerror(
            "Error de actualizacion",
            f"No se pudo descargar la actualizacion:\n{e}"
        ))


def _show_downloading(app):
    import customtkinter as ctk
    win = ctk.CTkToplevel(app)
    win.title("Actualizando")
    win.geometry("300x80")
    win.resizable(False, False)
    win.grab_set()
    ctk.CTkLabel(win, text="Descargando actualizacion, por favor espera...").pack(
        expand=True, padx=20, pady=20
    )
