"""
Auto-updater para GeneradorContratos.
Comprueba la última release en GitHub y, si hay versión nueva,
pregunta al usuario si desea actualizar. La descarga e instala en
%LOCALAPPDATA%\GeneradorContratos\ (carpeta local, fuera de OneDrive)
para evitar interferencias con la sincronización al cargar DLLs.
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

GITHUB_REPO  = "aruizciee/GeneradorContratos"
GITHUB_API   = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
EXE_NAME     = "generador_contratos.exe"


# ---------------------------------------------------------------------------
# Directorio de instalación local (fuera de OneDrive)
# ---------------------------------------------------------------------------

def _get_install_dir():
    """Devuelve %LOCALAPPDATA%\GeneradorContratos, creándolo si no existe."""
    local_appdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    d = os.path.join(local_appdata, "GeneradorContratos")
    os.makedirs(d, exist_ok=True)
    return d


# ---------------------------------------------------------------------------
# Punto de entrada público
# ---------------------------------------------------------------------------

def check_for_updates(app):
    """Lanza la comprobación en un hilo daemon para no bloquear la UI."""
    threading.Thread(target=_check, args=(app,), daemon=True).start()


# ---------------------------------------------------------------------------
# Lógica interna
# ---------------------------------------------------------------------------

def _check(app):
    try:
        resp = requests.get(GITHUB_API, timeout=10)
        resp.raise_for_status()
        data   = resp.json()
        latest = data.get("tag_name", "")
        if not latest or _parse_version(latest) <= _parse_version(VERSION):
            return                               # ya tenemos la última versión

        asset_url = _find_asset(data)
        if not asset_url:
            return

        # Volvemos al hilo principal para mostrar la UI
        app.after(0, lambda: _prompt_update(app, latest, asset_url))

    except Exception:
        pass    # sin internet o error de red → silencioso


def _parse_version(tag):
    """Convierte 'v2026.03.23-4' → (2026, 3, 23, 4) para comparación numérica."""
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
        "Actualización disponible",
        f"Hay una nueva versión disponible: {latest}\n"
        f"Versión actual: {VERSION}\n\n"
        "¿Deseas actualizar ahora?\n"
        "(La aplicación se cerrará y se reiniciará automáticamente.)"
    )
    if answer:
        threading.Thread(target=_download_and_restart, args=(app, asset_url), daemon=True).start()


def _download_and_restart(app, asset_url):
    """
    Descarga el nuevo .exe en %LOCALAPPDATA%\\GeneradorContratos\\ y lanza
    un script batch que lo activa. Usar esta carpeta local (no OneDrive)
    evita que la sincronización en tiempo real bloquee la carga de DLLs.
    """
    if not getattr(sys, "frozen", False):
        messagebox.showinfo("Dev mode", "Auto-update desactivado en modo desarrollo.")
        return

    install_dir = _get_install_dir()
    install_exe = os.path.join(install_dir, EXE_NAME)
    new_exe     = install_exe + ".new"
    bat_path    = os.path.join(install_dir, "_updater.bat")

    try:
        # Mostrar aviso de descarga en hilo principal
        app.after(0, lambda: _show_downloading(app))

        # Descargar nuevo ejecutable en la carpeta local
        resp = requests.get(asset_url, stream=True, timeout=120)
        resp.raise_for_status()
        with open(new_exe, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)

        # Script batch: espera a que este proceso termine, sustituye el exe
        # y lo relanza desde la carpeta local (fuera de OneDrive).
        pid = os.getpid()
        bat_content = (
            f"@echo off\n"
            f":wait\n"
            f'tasklist /fi "PID eq {pid}" 2>nul | find /i "generador_contratos.exe" >nul\n'
            f"if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)\n"
            f"timeout /t 2 /nobreak >nul\n"
            f'move /y "{new_exe}" "{install_exe}"\n'
            f"timeout /t 1 /nobreak >nul\n"
            f'start "" "{install_exe}"\n'
            f'del "%~f0"\n'
        )
        with open(bat_path, "w") as f:
            f.write(bat_content)

        subprocess.Popen(
            bat_path,
            shell=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        app.after(0, app.destroy)

    except Exception as e:
        for path in (new_exe, bat_path):
            try:
                os.remove(path)
            except OSError:
                pass
        app.after(0, lambda: messagebox.showerror(
            "Error de actualización",
            f"No se pudo descargar la actualización:\n{e}"
        ))


def _show_downloading(app):
    """Muestra una ventana modal sencilla de 'Descargando…'"""
    import customtkinter as ctk
    win = ctk.CTkToplevel(app)
    win.title("Actualizando")
    win.geometry("300x80")
    win.resizable(False, False)
    win.grab_set()
    ctk.CTkLabel(win, text="Descargando actualización, por favor espera…").pack(
        expand=True, padx=20, pady=20
    )
    # La ventana se cerrará sola cuando app.destroy() se llame
