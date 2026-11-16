"""
Auto-updater para DocuSend.

Comprueba la ultima release en GitHub y, si hay version nueva, pregunta al
usuario si desea actualizar. Descarga el nuevo ejecutable en
%LOCALAPPDATA%\\DocuSend\\ (carpeta local, fuera de OneDrive), verifica su
integridad y lanza un script que sustituye el ejecutable y reinicia la app.
"""
import hashlib
import os
import sys
import subprocess
import threading
from tkinter import messagebox


class _LazyRequests:
    """requests se importa en el primer uso para no retrasar el arranque."""
    def __getattr__(self, name):
        import requests
        return getattr(requests, name)


requests = _LazyRequests()

try:
    from version import VERSION
except ImportError:
    VERSION = "v0.0.0-dev"

GITHUB_REPO  = "aruizciee/docusend"
GITHUB_API   = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
EXE_NAME     = "docusend.exe"

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _get_install_dir():
    """Devuelve %LOCALAPPDATA%\\DocuSend, creandolo si no existe."""
    local_appdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    d = os.path.join(local_appdata, "DocuSend")
    os.makedirs(d, exist_ok=True)
    return d


def check_for_updates(app, manual=False):
    """Lanza la comprobacion en un hilo daemon para no bloquear la UI.

    manual=True (boton "Buscar actualizaciones") informa tambien cuando no hay
    version nueva o cuando la comprobacion falla; la comprobacion automatica
    del arranque es silenciosa en esos casos.
    """
    threading.Thread(target=_check, args=(app, manual), daemon=True).start()


def _check(app, manual=False):
    try:
        resp = requests.get(GITHUB_API, timeout=10)
        resp.raise_for_status()
        data   = resp.json()
        latest = data.get("tag_name", "")
        if not latest or _parse_version(latest) <= _parse_version(VERSION):
            if manual:
                _notify(app, messagebox.showinfo, "Actualizaciones",
                        f"Ya tienes la ultima version ({VERSION}).")
            return

        asset = _find_asset(data)
        if not asset:
            if manual:
                _notify(app, messagebox.showwarning, "Actualizaciones",
                        f"Hay una version nueva ({latest}), pero todavia no tiene "
                        "el ejecutable disponible. Intentalo de nuevo en unos minutos.")
            return

        app.after(0, lambda: _prompt_update(app, latest, asset))

    except Exception as e:
        if manual:
            _notify(app, messagebox.showwarning, "Actualizaciones",
                    "No se pudo comprobar si hay actualizaciones.\n\n"
                    f"Revisa tu conexion a internet.\n({e})")


def _notify(app, func, title, text):
    """Muestra un mensaje desde cualquier hilo (Tkinter no es thread-safe)."""
    try:
        app.after(0, lambda: func(title, text))
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
    """Devuelve {'url', 'sha256', 'size'} del ejecutable de la release, o None.

    'sha256' sale del campo 'digest' que publica GitHub para cada archivo; es ""
    si la release no lo trae (en ese caso solo se comprueba el tamano).
    """
    for asset in release_data.get("assets", []):
        if asset.get("name", "").lower() == EXE_NAME:
            digest = asset.get("digest") or ""
            sha256 = digest.split(":", 1)[1].lower() if digest.lower().startswith("sha256:") else ""
            return {
                "url":    asset["browser_download_url"],
                "sha256": sha256,
                "size":   int(asset.get("size") or 0),
            }
    return None


def _verify_download(path, sha256="", size=0):
    """Lanza ValueError si el archivo descargado no coincide con lo publicado."""
    actual_size = os.path.getsize(path)
    if size and actual_size != size:
        raise ValueError(f"La descarga esta incompleta ({actual_size} de {size} bytes).")
    if actual_size < 1024 * 1024:
        raise ValueError("El archivo descargado no parece un ejecutable valido.")
    if sha256:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        if h.hexdigest().lower() != sha256.lower():
            raise ValueError("El archivo descargado no coincide con el publicado (SHA-256 distinto).")


def _clean_environment():
    """Entorno para el proceso que reinicia la app.

    Una app PyInstaller de un solo archivo deja en el entorno variables que
    indican a sus procesos hijos que reutilicen su carpeta temporal (_MEIxxxx).
    Si la app reiniciada las hereda, busca python3xx.dll en la carpeta de la
    version anterior, que ya se ha borrado, y falla con "Failed to load Python
    DLL". Se eliminan y se pide expresamente un arranque limpio.
    """
    env = {k: v for k, v in os.environ.items()
           if not (k.upper().startswith("_PYI_") or k.upper().startswith("_MEIPASS"))}
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    return env


def _build_updater_script(pid, new_exe, original_exe):
    """Contenido del .bat que sustituye el ejecutable y reinicia la app."""
    error_msg = ("No se pudo actualizar DocuSend: el archivo de la aplicacion esta bloqueado "
                 "(posiblemente por OneDrive o un antivirus). Intentalo de nuevo mas tarde.")
    return (
        "@echo off\n"
        # UTF-8 para que las rutas con acentos o enes se lean bien
        "chcp 65001 >nul\n"
        "setlocal enabledelayedexpansion\n"
        # Arranque limpio de la nueva version (ver _clean_environment)
        "set PYINSTALLER_RESET_ENVIRONMENT=1\n"
        "set _PYI_ARCHIVE_FILE=\n"
        "set _PYI_APPLICATION_HOME_DIR=\n"
        "set _PYI_PARENT_PROCESS_LEVEL=\n"
        "set _PYI_SPLASH_IPC=\n"
        "set _MEIPASS2=\n"
        ":wait_process\n"
        f'tasklist /fi "PID eq {pid}" 2>nul | find /i "{pid}" >nul\n'
        "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait_process)\n"
        # Sin pausa fija: si el archivo sigue bloqueado, lo resuelven los reintentos
        "set MAX_RETRIES=30\n"
        "set RETRY_COUNT=0\n"
        ":retry_update\n"
        # move /y sustituye el archivo sin borrarlo antes: si falla, la version
        # anterior sigue intacta.
        f'move /y "{new_exe}" "{original_exe}" >nul 2>&1\n'
        f'if exist "{new_exe}" (\n'
        "    set /a RETRY_COUNT+=1\n"
        "    if !RETRY_COUNT! geq !MAX_RETRIES! goto on_error\n"
        "    timeout /t 1 /nobreak >nul\n"
        "    goto retry_update\n"
        ")\n"
        f'start "" "{original_exe}"\n'
        "goto end\n"
        ":on_error\n"
        f'del /f /q "{new_exe}" >nul 2>&1\n'
        'powershell -NoProfile -WindowStyle Hidden -Command "Add-Type -AssemblyName PresentationFramework; '
        f"[System.Windows.MessageBox]::Show('{error_msg}','DocuSend','OK','Error')\" >nul 2>&1\n"
        f'start "" "{original_exe}"\n'
        ":end\n"
        'del "%~f0"\n'
    )


def _prompt_update(app, latest, asset):
    if getattr(app, "is_running", False):
        # No interrumpir un lote en curso; se volvera a ofrecer en el proximo arranque
        return
    answer = messagebox.askyesno(
        "Actualizacion disponible",
        f"Hay una nueva version disponible: {latest}\n"
        f"Version actual: {VERSION}\n\n"
        "Deseas actualizar ahora?\n"
        "(La aplicacion se cerrara y se reiniciara automaticamente.)"
    )
    if answer:
        threading.Thread(target=_download_and_restart, args=(app, asset), daemon=True).start()


def _download_and_restart(app, asset):
    """
    Descarga el nuevo .exe, lo verifica y lanza un script batch que reemplaza
    el ejecutable original y lo reinicia.
    """
    if not getattr(sys, "frozen", False):
        app.after(0, lambda: messagebox.showinfo(
            "Dev mode", "Auto-update desactivado en modo desarrollo."))
        return

    original_exe = sys.executable
    temp_dir = _get_install_dir()
    new_exe  = os.path.join(temp_dir, EXE_NAME + ".new")
    bat_path = os.path.join(temp_dir, "_updater.bat")

    try:
        app.after(0, lambda: _show_downloading(app))

        resp = requests.get(asset["url"], stream=True, timeout=120)
        resp.raise_for_status()
        total = asset.get("size") or int(resp.headers.get("content-length") or 0)
        done = 0
        last_shown = -1
        with open(new_exe, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)
                done += len(chunk)
                if total:
                    pct = int(done * 100 / total)
                    if pct != last_shown:
                        last_shown = pct
                        app.after(0, lambda p=pct: _set_progress(app, p))

        app.after(0, lambda: _set_progress(app, 100, "Verificando la descarga..."))
        _verify_download(new_exe, asset.get("sha256", ""), asset.get("size", 0))

        with open(bat_path, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(_build_updater_script(os.getpid(), new_exe, original_exe))

        subprocess.Popen(["cmd", "/c", bat_path], env=_clean_environment(),
                         creationflags=_NO_WINDOW, close_fds=True)
        app.after(0, lambda: _close_for_update(app))

    except Exception as e:
        for path in (new_exe, bat_path):
            try:
                os.remove(path)
            except OSError:
                pass
        app.after(0, lambda: _close_downloading(app))
        app.after(0, lambda: messagebox.showerror(
            "Error de actualizacion",
            f"No se pudo descargar la actualizacion:\n{e}\n\n"
            "La aplicacion sigue funcionando con la version actual."
        ))


def _close_for_update(app):
    """Cierra la app guardando antes la configuracion, si la app lo permite."""
    _close_downloading(app)
    finish = getattr(app, "_finish_close", None)
    if callable(finish):
        finish()
    else:
        app.destroy()


def _show_downloading(app):
    import customtkinter as ctk
    win = ctk.CTkToplevel(app)
    win.title("Actualizando")
    win.geometry("340x110")
    win.resizable(False, False)
    win.transient(app)
    win.protocol("WM_DELETE_WINDOW", lambda: None)   # no cerrar a mitad de descarga
    win.grab_set()
    label = ctk.CTkLabel(win, text="Descargando actualizacion...")
    label.pack(padx=20, pady=(18, 8))
    bar = ctk.CTkProgressBar(win, width=280)
    bar.set(0)
    bar.pack(padx=20, pady=(0, 18))
    app._update_window = (win, label, bar)


def _set_progress(app, percent, text=None):
    widgets = getattr(app, "_update_window", None)
    if not widgets:
        return
    win, label, bar = widgets
    try:
        bar.set(max(0, min(percent, 100)) / 100)
        label.configure(text=text or f"Descargando actualizacion... {percent}%")
    except Exception:
        pass


def _close_downloading(app):
    widgets = getattr(app, "_update_window", None)
    app._update_window = None
    if not widgets:
        return
    try:
        widgets[0].grab_release()
        widgets[0].destroy()
    except Exception:
        pass
