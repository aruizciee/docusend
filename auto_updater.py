"""
Auto-updater para GeneradorContratos.
Comprueba la última release en GitHub y, si hay versión nueva,
pregunta al usuario si desea actualizar. La descarga y reemplaza
el .exe usando un script batch auxiliar (necesario en Windows porque
no se puede sobreescribir un ejecutable mientras está en uso).
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
        data     = resp.json()
        latest   = data.get("tag_name", "")
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
    """Descarga el nuevo .exe y lanza un script batch que lo reemplaza."""
    if not getattr(sys, "frozen", False):
        # Modo desarrollo: no tiene sentido reemplazar nada
        messagebox.showinfo("Dev mode", "Auto-update desactivado en modo desarrollo.")
        return

    current_exe = sys.executable
    new_exe     = current_exe + ".new"
    bat_path    = os.path.join(os.path.dirname(current_exe), "_updater.bat")

    try:
        # Mostrar aviso de descarga en hilo principal
        app.after(0, lambda: _show_downloading(app))

        # Descargar nuevo ejecutable
        resp = requests.get(asset_url, stream=True, timeout=120)
        resp.raise_for_status()
        with open(new_exe, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)

        # Crear script batch que espera a que este proceso (por PID) termine
        # antes de reemplazar el exe — evita el error de DLL de PyInstaller.
        pid = os.getpid()
        bat_content = (
            f"@echo off\n"
            f":wait\n"
            f'tasklist /fi "PID eq {pid}" 2>nul | find /i "generador_contratos.exe" >nul\n'
            f"if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)\n"
            f"timeout /t 5 /nobreak >nul\n"
            f'move /y "{new_exe}" "{current_exe}"\n'
            f"timeout /t 2 /nobreak >nul\n"
            f'start "" "{current_exe}"\n'
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
        # Limpiar archivos temporales si algo falló
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
