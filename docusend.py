import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from docxtpl import DocxTemplate
import win32com.client as win32
import os
import re
import json
import html
from datetime import datetime
import threading
import subprocess
import csv
import shutil
import time
from io import BytesIO
import sys
from auto_updater import check_for_updates
from version import VERSION
import ctypes

try:
    myappid = f"ciee.docusend.app.{VERSION}"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

try:
    from pyhanko.sign import signers
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    PYHANKO_AVAILABLE = True
except ImportError:
    PYHANKO_AVAILABLE = False

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".docusend_config.json")

def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(data):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

_PLACEHOLDER_RE = re.compile(r'\{\{\s*(.*?)\s*\}\}')

def substitute_variables(text, context, escape_html=False):
    """Sustituye {{Columna}} (con o sin espacios dentro de las llaves) en una
    sola pasada. Los campos desconocidos se dejan tal cual."""
    def repl(m):
        key = m.group(1)
        if key not in context and escape_html:
            key = html.unescape(key)
        if key not in context:
            return m.group(0)
        val = str(context[key])
        return html.escape(val) if escape_html else val
    return _PLACEHOLDER_RE.sub(repl, text or "")

def find_placeholders(text):
    """Nombres de los campos {{...}} presentes en un texto."""
    return {m.group(1) for m in _PLACEHOLDER_RE.finditer(text or "")}

_EMAIL_RE = re.compile(r'^[^@\s<>()\[\],;:"]+@[^@\s<>()\[\],;:"]+\.[^@\s<>()\[\],;:".]{2,}$')

def invalid_emails(text):
    """Direcciones con formato incorrecto en un texto (separadas por ; o ,)."""
    parts = [p.strip() for p in re.split(r'[;,]', text or "") if p.strip()]
    return [p for p in parts if not _EMAIL_RE.match(p)]

AUTOFIRMA_EXE = r'C:\Program Files\Autofirma\Autofirma\AutofirmaCommandLine.exe'
REPORT_DIR = "_DocuSend_informes"

_JINJA_RESERVED = {
    "true", "false", "none", "and", "or", "not", "if", "else", "elif", "for", "in",
    "is", "loop", "self", "super", "varargs", "kwargs", "caller", "range", "dict",
    "lipsum", "cycler", "joiner", "namespace", "block", "set", "with", "import",
}

def column_aliases(columns):
    """La plantilla Word se procesa con Jinja2, que no admite nombres con
    espacios, guiones, etc. Devuelve {columna: alias válido} para esas columnas."""
    aliases = {}
    for i, name in enumerate(columns):
        if not name.isidentifier() or name.lower() in _JINJA_RESERVED:
            aliases[name] = f"campo_{i}_"
    return aliases

class ColumnTemplate(DocxTemplate):
    """DocxTemplate que acepta {{Nombre completo}} y otros nombres de columna
    que no son identificadores válidos, traduciéndolos a un alias interno."""
    def __init__(self, template_file, aliases=None):
        super().__init__(template_file)
        self._col_aliases = dict(aliases or {})

    def patch_xml(self, src_xml):
        xml = super().patch_xml(src_xml)
        for name, alias in self._col_aliases.items():
            for variant in {name, html.escape(name, quote=False), html.escape(name)}:
                xml = re.sub(r'\{\{\s*' + re.escape(variant) + r'\s*\}\}',
                             lambda m, a=alias: '{{ ' + a + ' }}', xml)
        return xml

def format_cell(value):
    """Convierte un valor de Excel en texto legible para plantillas y correos."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, (pd.Timestamp, datetime)):
        if (value.hour, value.minute, value.second) == (0, 0, 0):
            return value.strftime("%d/%m/%Y")
        return value.strftime("%d/%m/%Y %H:%M")
    if isinstance(value, bool):
        return "Sí" if value else "No"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

def unique_base_path(folder, base_name, extensions):
    """Devuelve una ruta base (sin extensión) que no colisione con archivos
    existentes ni con nombres ya usados en este lote."""
    candidate = base_name
    n = 2
    while any(os.path.exists(os.path.join(folder, candidate + ext)) for ext in extensions):
        candidate = f"{base_name}_{n}"
        n += 1
    return os.path.join(folder, candidate)

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ── Sistema visual ─────────────────────────────────────────────────────────────
# Colores como (tema claro, tema oscuro). Usar siempre estos tokens en lugar de
# colores sueltos para que la app se vea coherente en ambos temas.
FONT_FAMILY = "Segoe UI"
C = {
    "header":        ("#1f538d", "#14375e"),
    "header_muted":  ("#c9d6e8", "#9fb3cc"),
    "primary":       ("#1f6aa5", "#2b7bc0"),
    "primary_hover": ("#174f7c", "#1f5f96"),
    "success":       ("#2e7d4f", "#3fb37a"),
    "success_hover": ("#24633f", "#329363"),
    "warning":       ("#b26a00", "#f0a030"),
    "danger":        ("#c62828", "#e5534b"),
    "danger_hover":  ("#a31f1f", "#c94440"),
    "text":          ("gray10", "gray92"),
    "muted":         ("gray40", "gray62"),
    "border":        ("gray82", "gray28"),
    "neutral":       ("gray55", "gray35"),
    "neutral_hover": ("gray45", "gray28"),
    "card":          ("gray96", "gray16"),
}
F = {
    "h1":      (FONT_FAMILY, 16, "bold"),
    "title":   (FONT_FAMILY, 14, "bold"),
    "section": (FONT_FAMILY, 12, "bold"),
    "body":    (FONT_FAMILY, 12),
    "small":   (FONT_FAMILY, 11),
    "mono":    ("Consolas", 11),
}
# Fuente por defecto para los widgets que no indican una explícitamente
try:
    ctk.ThemeManager.theme["CTkFont"]["family"] = FONT_FAMILY
    ctk.ThemeManager.theme["CTkFont"]["size"] = 12
except Exception:
    pass
# Mismo azul primario en todos los controles
try:
    _t = ctk.ThemeManager.theme
    _t["CTkButton"]["fg_color"] = list(C["primary"])
    _t["CTkButton"]["hover_color"] = list(C["primary_hover"])
    for _w in ("CTkCheckBox", "CTkRadioButton"):
        _t[_w]["fg_color"] = list(C["primary"])
        _t[_w]["hover_color"] = list(C["primary_hover"])
    _t["CTkProgressBar"]["progress_color"] = list(C["primary"])
except Exception:
    pass

STEPS = ["Archivos", "Configuración", "Correo", "Generar"]

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DocuSend")
        
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        _ico = os.path.join(base_path, "assets", "icon.ico")
        if os.path.exists(_ico):
            # Intento normal
            try:
                self.iconbitmap(_ico)
            except Exception:
                pass
            
            # Intento forzado a traves de Pillow (corrige el icono en la barra de tareas en CustomTkinter)
            try:
                from PIL import Image, ImageTk
                img = ImageTk.PhotoImage(Image.open(_ico))
                self.wm_iconphoto(True, img)
                self.iconphoto(True, img)
            except Exception:
                pass
                
        self.geometry("720x700")
        self.minsize(700, 640)

        self.after(3000, lambda: check_for_updates(self))

        # --- Variables de datos ---
        self.word_template_path   = ""
        self.excel_data_path      = ""
        self.outlook_template_path = ""
        self.output_folder        = ""
        self.excel_columns        = []
        self.outlook_accounts     = []
        self.extra_attachments    = []
        self.current_step         = 0

        # --- Layout raíz ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── Cabecera ──────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color=C["header"], corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(hdr, text="DocuSend", text_color="white",
                     font=F["h1"]).grid(row=0, column=0, pady=12, padx=20, sticky="w")
        ctk.CTkLabel(hdr, text=VERSION, text_color=C["header_muted"],
                     font=F["small"]).grid(row=0, column=1, pady=12, padx=20, sticky="e")

        # ── Contenido central ─────────────────────────────────────────────────
        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=1, column=0, sticky="nsew", padx=24, pady=(16, 0))
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(1, weight=1)

        # Indicador de pasos
        self._build_step_indicator()

        # Frames de cada paso (apilados en la misma celda, se muestran/ocultan)
        self.step_frames = []
        self._build_step1()
        self._build_step2()
        self._build_step3()
        self._build_step4()
        for f in self.step_frames:
            f.grid(row=1, column=0, sticky="nsew", pady=(12, 0))

        # ── Pie de navegación ─────────────────────────────────────────────────
        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.grid(row=2, column=0, sticky="ew", padx=24, pady=12)
        nav.grid_columnconfigure(1, weight=1)

        self.btn_prev = ctk.CTkButton(nav, text="← Anterior", width=120,
                                      fg_color=C["neutral"], hover_color=C["neutral_hover"],
                                      command=self._prev_step)
        self.btn_prev.grid(row=0, column=0, padx=(0, 8))

        ctk.CTkLabel(nav, text="", font=F["small"]).grid(row=0, column=1)  # spacer

        self.btn_next = ctk.CTkButton(nav, text="Siguiente →", width=120,
                                      command=self._next_step)
        self.btn_next.grid(row=0, column=2, padx=(8, 0))

        # Versión + actualización + configuración (pie derecho)
        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 8))
        foot.grid_columnconfigure(1, weight=1)
        
        self.combo_theme = ctk.CTkOptionMenu(foot, values=["Sistema", "Claro", "Oscuro"],
                                             command=self._change_appearance_mode,
                                             width=80, height=24, font=F["small"], fg_color=C["border"], button_color=C["border"], button_hover_color=C["neutral"], text_color=C["text"])
        self.combo_theme.grid(row=0, column=0, sticky="w")
        self.combo_theme.set("Sistema")

        ctk.CTkButton(foot, text="Buscar actualizaciones", width=160, height=24,
                      font=F["small"], fg_color="transparent", border_width=1,
                      text_color=C["muted"], border_color=C["border"],
                      command=lambda: check_for_updates(self)).grid(row=0, column=2, sticky="e")

        self._ready = False          # no guardar config hasta terminar de restaurarla
        self._saved_account = ""
        self._restore_config()
        self._ready = True
        self._show_step(0)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(1500, self._load_outlook_accounts)

    # ═══════════════════════════════════════════════════════════════════════════
    # WIZARD — indicador de pasos
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_step_indicator(self):
        self.ind_frame = ctk.CTkFrame(self.content, fg_color="transparent")
        self.ind_frame.grid(row=0, column=0, sticky="ew")
        self.ind_labels = []
        self.ind_dots   = []
        cols = len(STEPS) * 2 - 1
        for i in range(cols):
            self.ind_frame.grid_columnconfigure(i, weight=1 if i % 2 == 1 else 0)

        for i, name in enumerate(STEPS):
            col = i * 2
            dot = ctk.CTkLabel(self.ind_frame, text="●", font=(FONT_FAMILY, 18),
                               text_color=C["muted"], width=28)
            dot.grid(row=0, column=col)
            lbl = ctk.CTkLabel(self.ind_frame, text=name, font=F["small"],
                               text_color=C["muted"])
            lbl.grid(row=1, column=col)
            self.ind_dots.append(dot)
            self.ind_labels.append(lbl)
            if i < len(STEPS) - 1:
                ctk.CTkLabel(self.ind_frame, text="────", text_color=C["border"],
                             font=F["small"]).grid(row=0, column=col + 1, sticky="ew")

    def _update_step_indicator(self, step):
        for i, (dot, lbl) in enumerate(zip(self.ind_dots, self.ind_labels)):
            if i < step:
                dot.configure(text="✓", text_color=C["success"])
                lbl.configure(text_color=C["muted"], font=F["small"])
            elif i == step:
                dot.configure(text="●", text_color=C["primary"])
                lbl.configure(text_color=C["text"],
                               font=(FONT_FAMILY, 11, "bold"))
            else:
                dot.configure(text="●", text_color=C["muted"])
                lbl.configure(text_color=C["muted"],
                               font=F["small"])

    def _show_step(self, step):
        self.current_step = step
        for i, f in enumerate(self.step_frames):
            if i == step:
                f.grid()
                f.tkraise()
            else:
                f.grid_remove()
        self._update_step_indicator(step)
        self.btn_prev.configure(state="normal" if step > 0 else "disabled")
        if step == len(STEPS) - 1:
            self._refresh_summary()
            if self.send_mode.get() == "send":
                self.btn_next.configure(text="✉ Generar y enviar", width=150,
                                        fg_color=C["danger"], hover_color=C["danger_hover"],
                                        command=self.start_generation)
            else:
                self.btn_next.configure(text="✓ Generar", width=150,
                                        fg_color=C["success"], hover_color=C["success_hover"],
                                        command=self.start_generation)
        else:
            self.btn_next.configure(text="Siguiente →", width=120,
                                    fg_color=C["primary"], hover_color=C["primary_hover"],
                                    command=self._next_step)

    def _next_step(self):
        if self._validate_step(self.current_step):
            self._save_config()
            self._show_step(self.current_step + 1)

    def _prev_step(self):
        self._save_config()
        self._show_step(self.current_step - 1)

    def _on_close(self):
        if getattr(self, "_closing", False):
            return
        if self.is_running:
            if not messagebox.askyesno(
                    "Proceso en curso",
                    "Hay un proceso en curso.\n\nSe detendrá al terminar el documento actual, "
                    "se guardará el informe y la aplicación se cerrará.\n\n¿Detener y salir?",
                    icon="warning", default="no"):
                return
            self._closing = True
            self.is_running = False
            self._set_stop_enabled(False)
            self.log("Cerrando: esperando a que termine el documento actual...")
            self._close_deadline = time.monotonic() + 150
            self._wait_and_close()
            return
        self._finish_close()

    def _wait_and_close(self):
        worker = getattr(self, "_worker", None)
        if worker is not None and worker.is_alive() and time.monotonic() < self._close_deadline:
            self.after(200, self._wait_and_close)
            return
        self._finish_close()

    def _finish_close(self):
        try:
            self._save_config()
        finally:
            self.destroy()

    def _validate_step(self, step):
        if step == 0:
            if not self.word_template_path:
                messagebox.showwarning("Faltan datos", "Selecciona la plantilla Word.")
                return False
            if not self.excel_data_path:
                messagebox.showwarning("Faltan datos", "Selecciona el archivo Excel.")
                return False
            if not self.output_folder:
                messagebox.showwarning("Faltan datos", "Selecciona la carpeta de salida.")
                return False
            for label, path, check in [
                ("la plantilla Word", self.word_template_path, os.path.isfile),
                ("el archivo Excel",  self.excel_data_path,    os.path.isfile),
                ("la carpeta de salida", self.output_folder,   os.path.isdir),
            ]:
                if not check(path):
                    messagebox.showwarning(
                        "No encontrado",
                        f"No se encuentra {label}:\n{path}\n\n"
                        "Comprueba que la unidad está conectada o selecciónalo de nuevo.")
                    return False
        elif step == 1:
            if not self.entry_filename_pattern.get().strip():
                messagebox.showwarning("Faltan datos", "Indica el patrón de nombre de archivo.")
                return False
        elif step == 2:
            if self.send_mode.get() != "none":
                col = self.entry_email_col.get().strip()
                if not col:
                    messagebox.showwarning("Faltan datos", "Indica la columna de email.")
                    return False
                if self.excel_columns and col not in self.excel_columns:
                    messagebox.showwarning("Columna no encontrada",
                                           f"La columna '{col}' no existe en el Excel.")
                    return False
            if self.send_mode.get() != "none" and \
               self.email_mode.get() == "template" and not self.outlook_template_path:
                messagebox.showwarning("Faltan datos", "Selecciona la plantilla .oft de Outlook.")
                return False
            if self.send_mode.get() != "none" and self.email_mode.get() == "template" \
               and not os.path.isfile(self.outlook_template_path):
                messagebox.showwarning("No encontrado",
                                       f"No se encuentra la plantilla .oft:\n{self.outlook_template_path}")
                return False
        return True

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 1 — Archivos
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_step1(self):
        f = ctk.CTkFrame(self.content)
        self.step_frames.append(f)
        f.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(f, text="Selecciona los archivos necesarios",
                     font=F["title"]).grid(
            row=0, column=0, columnspan=3, pady=(16, 20), padx=16, sticky="w")

        rows = [
            ("Plantilla Word (.docx)", "lbl_word",   self.select_word,   "📄"),
            ("Datos Excel (.xlsx)",    "lbl_excel",  self.select_excel,  "📊"),
            ("Carpeta de salida",      "lbl_output", self.select_output, "📁"),
        ]
        for i, (label, attr, cmd, icon) in enumerate(rows, start=1):
            ctk.CTkLabel(f, text=f"{icon}  {label}",
                         font=F["body"]).grid(row=i*2-1, column=0, columnspan=3,
                                                    padx=16, pady=(10, 2), sticky="w")
            lbl = ctk.CTkLabel(f, text="Sin seleccionar", text_color=C["muted"],
                               font=F["small"], anchor="w")
            lbl.grid(row=i*2, column=0, columnspan=2, padx=20, pady=(0, 4), sticky="ew")
            setattr(self, attr, lbl)
            ctk.CTkButton(f, text="Seleccionar", width=110, command=cmd).grid(
                row=i*2, column=2, padx=(0, 16), pady=(0, 4))

        ctk.CTkFrame(f, height=1, fg_color=C["border"]).grid(
            row=7, column=0, columnspan=3, sticky="ew", padx=16, pady=(12, 8))
            
        ctk.CTkLabel(f, text="📎 Adjuntos estáticos adicionales", font=F["body"]).grid(
            row=8, column=0, columnspan=3, padx=16, pady=(0, 2), sticky="w")
            
        self.lbl_extra_att = ctk.CTkLabel(f, text="Ningún archivo seleccionado", text_color=C["muted"],
                                          font=F["small"], anchor="w", justify="left")
        self.lbl_extra_att.grid(row=9, column=0, columnspan=2, padx=20, pady=(0, 4), sticky="w")
        
        btn_frame = ctk.CTkFrame(f, fg_color="transparent")
        btn_frame.grid(row=9, column=2, padx=(0, 16), pady=(0, 4), sticky="e")
        ctk.CTkButton(btn_frame, text="Añadir", width=55, command=self.add_extra_attachments).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="Borrar", width=55, fg_color=C["neutral"], hover_color=C["neutral_hover"], command=self.clear_extra_attachments).pack(side="left")

    def _change_appearance_mode(self, new_mode: str):
        if new_mode == "Sistema":
            ctk.set_appearance_mode("System")
        elif new_mode == "Claro":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")
        if hasattr(self, "log_box"):
            self._apply_log_colors()
        self._save_config()

    def clear_extra_attachments(self):
        self.extra_attachments = []
        self.lbl_extra_att.configure(text="Ningún archivo seleccionado", text_color=C["muted"])
        self._save_config()

    def add_extra_attachments(self):
        files = filedialog.askopenfilenames(title="Seleccionar adjuntos estáticos")
        if files:
            self.extra_attachments.extend(files)
            count = len(self.extra_attachments)
            names = [os.path.basename(f) for f in self.extra_attachments]
            text_disp = f"{count} archivo(s): " + ", ".join(names)
            if len(text_disp) > 55: text_disp = text_disp[:52] + "..."
            self.lbl_extra_att.configure(text=text_disp, text_color=C["text"])
            self._save_config()

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 2 — Configuración Excel
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_step2(self):
        f = ctk.CTkScrollableFrame(self.content)
        self.step_frames.append(f)

        ctk.CTkLabel(f, text="Configura el nombre de los archivos",
                     font=F["title"]).grid(
            row=0, column=0, columnspan=2, pady=(16, 4), padx=16, sticky="w")
        ctk.CTkLabel(f, text="Usa {{NombreColumna}} para insertar datos del Excel en el nombre del archivo y en la plantilla Word.",
                     text_color=C["muted"], font=F["small"], wraplength=580, justify="left").grid(
            row=1, column=0, columnspan=2, padx=16, pady=(0, 16), sticky="w")

        # Patrón nombre archivo
        ctk.CTkLabel(f, text="Patrón nombre archivo:").grid(row=2, column=0, padx=16, pady=6, sticky="w")
        self.entry_filename_pattern = ctk.CTkEntry(f, placeholder_text="Ej: {{Apellidos}}, {{Nombre}}")
        self.entry_filename_pattern.insert(0, "{{Apellidos}}, {{Nombre}}")
        self.entry_filename_pattern.grid(row=2, column=1, padx=16, pady=6, sticky="ew")

        # Separador
        ctk.CTkFrame(f, height=1, fg_color=C["border"]).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=16, pady=12)

        # Insertar campo
        ctk.CTkLabel(f, text="Insertar campo:", font=F["section"]).grid(
            row=4, column=0, padx=16, pady=(0, 6), sticky="w")
        ctk.CTkLabel(f, text="Selecciona un campo y cópialo para pegarlo en la plantilla Word o en el patrón.",
                     text_color=C["muted"], font=F["small"]).grid(
            row=5, column=0, columnspan=2, padx=16, pady=(0, 8), sticky="w")

        frame_insert = ctk.CTkFrame(f, fg_color="transparent")
        frame_insert.grid(row=6, column=0, columnspan=2, padx=16, sticky="ew")
        frame_insert.grid_columnconfigure(0, weight=1)

        self.combo_fields = ctk.CTkComboBox(frame_insert,
                                            values=["(carga un Excel primero)"],
                                            state="readonly")
        self.combo_fields.set("(carga un Excel primero)")
        self.combo_fields.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        btn_frame = ctk.CTkFrame(frame_insert, fg_color="transparent")
        btn_frame.grid(row=0, column=1)
        ctk.CTkButton(btn_frame, text="📋 Copiar", width=90, fg_color=C["neutral"], hover_color=C["neutral_hover"],
                      command=self._copy_field_to_clipboard).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="→ Patrón", width=90,
                      command=lambda: self._insert_field(self.entry_filename_pattern)).pack(side="left", padx=2)

        # ── Firma Digital ──────────────────────────────────────────────────────
        ctk.CTkFrame(f, height=1, fg_color=C["border"]).grid(
            row=7, column=0, columnspan=2, sticky="ew", padx=16, pady=(16, 8))

        self.enable_signature = ctk.BooleanVar(value=False)
        self.chk_sign = ctk.CTkCheckBox(f, text="Firmar digitalmente los PDFs", font=F["section"],
                                        variable=self.enable_signature, command=self.toggle_signature_ui)
        self.chk_sign.grid(row=8, column=0, columnspan=2, padx=16, pady=(0, 6), sticky="w")

        self.sign_options_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.sign_options_frame.grid(row=9, column=0, columnspan=2, sticky="ew", padx=32)
        self.sign_options_frame.grid_remove()  # Hidden by default

        self.sign_mode = ctk.StringVar(value="pfx")
        rad_frame = ctk.CTkFrame(self.sign_options_frame, fg_color="transparent")
        rad_frame.pack(fill="x", pady=4)
        ctk.CTkRadioButton(rad_frame, text="Certificado PFX (.p12)", variable=self.sign_mode, value="pfx",
                           command=self.toggle_sign_mode_ui).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(rad_frame, text="AutoFirma (Certificado de Windows)", variable=self.sign_mode, value="autofirma",
                           command=self.toggle_sign_mode_ui).pack(side="left")

        self.pfx_frame = ctk.CTkFrame(self.sign_options_frame, fg_color="transparent")
        self.pfx_frame.pack(fill="x", pady=4)

        self.lbl_pfx_path = ctk.CTkLabel(self.pfx_frame, text="Ningún certificado seleccionado", text_color=C["muted"], font=F["small"])
        self.lbl_pfx_path.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

        ctk.CTkButton(self.pfx_frame, text="Seleccionar PFX", width=120, command=self.select_pfx).grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.entry_pfx_pass = ctk.CTkEntry(self.pfx_frame, placeholder_text="Contraseña PFX", show="*", width=140)
        self.entry_pfx_pass.grid(row=1, column=1, sticky="w")
        self.pfx_path = ""

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 3 — Correo
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_step3(self):
        f = ctk.CTkScrollableFrame(self.content)
        self.step_frames.append(f)
        f.grid_columnconfigure(0, weight=0)
        f.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(f, text="Configura el correo electrónico",
                     font=F["title"]).grid(
            row=0, column=0, columnspan=3, pady=(12, 6), padx=16, sticky="w")

        # Modo email
        self.email_mode = ctk.StringVar(value="manual")
        mode_frame = ctk.CTkFrame(f, fg_color="transparent")
        mode_frame.grid(row=1, column=0, columnspan=3, padx=16, pady=(0, 4), sticky="w")
        ctk.CTkRadioButton(mode_frame, text="Escribir asunto y cuerpo",
                           variable=self.email_mode, value="manual",
                           command=self.toggle_email_mode).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(mode_frame, text="Usar plantilla Outlook (.oft)",
                           variable=self.email_mode, value="template",
                           command=self.toggle_email_mode).pack(side="left")

        # Contenedor intercambiable (manual / template)
        self.email_inner = ctk.CTkFrame(f, fg_color="transparent")
        self.email_inner.grid(row=2, column=0, columnspan=3, sticky="ew", padx=8)
        self.email_inner.grid_columnconfigure(1, weight=1)

        # Manual
        self.lbl_subj = ctk.CTkLabel(self.email_inner, text="Asunto:")
        self.entry_subject = ctk.CTkEntry(self.email_inner,
                                          placeholder_text="Ej: Contrato adjunto — {{Nombre}}")
        self.lbl_body = ctk.CTkLabel(self.email_inner, text="Cuerpo:")
        self.txt_body = ctk.CTkTextbox(self.email_inner, height=75)

        # Template
        self.lbl_oft      = ctk.CTkLabel(self.email_inner, text="Archivo .oft:")
        self.lbl_oft_path = ctk.CTkLabel(self.email_inner, text="Sin seleccionar",
                                          text_color=C["muted"])
        self.btn_oft      = ctk.CTkButton(self.email_inner, text="Seleccionar .oft",
                                           command=self.select_oft)
        self.toggle_email_mode()

        # ── Destinatarios ──────────────────────────────────────────────────────
        ctk.CTkFrame(f, height=1, fg_color=C["border"]).grid(
            row=3, column=0, columnspan=3, sticky="ew", padx=16, pady=(6, 4))

        # Para (columna Excel) + adicional
        ctk.CTkLabel(f, text="Para (columna):").grid(row=4, column=0, padx=16, pady=3, sticky="w")
        self.entry_email_col = ctk.CTkComboBox(f, values=["Email"], state="normal")
        self.entry_email_col.set("Email")
        self.entry_email_col.grid(row=4, column=1, columnspan=2, padx=16, pady=3, sticky="ew")

        ctk.CTkLabel(f, text="Para (adicional):").grid(row=5, column=0, padx=16, pady=3, sticky="w")
        self.entry_to_extra = ctk.CTkEntry(f, placeholder_text="extra@ejemplo.com; otro@ejemplo.com")
        self.entry_to_extra.grid(row=5, column=1, columnspan=2, padx=16, pady=3, sticky="ew")

        ctk.CTkLabel(f, text="CC:").grid(row=6, column=0, padx=16, pady=3, sticky="w")
        self.entry_cc = ctk.CTkEntry(f, placeholder_text="copia@ejemplo.com  o  {{ColumnaCC}}")
        self.entry_cc.grid(row=6, column=1, columnspan=2, padx=16, pady=3, sticky="ew")

        ctk.CTkLabel(f, text="CCO:").grid(row=7, column=0, padx=16, pady=3, sticky="w")
        self.entry_bcc = ctk.CTkEntry(f, placeholder_text="oculto@ejemplo.com  o  {{ColumnaCCO}}")
        self.entry_bcc.grid(row=7, column=1, columnspan=2, padx=16, pady=3, sticky="ew")

        ctk.CTkFrame(f, height=1, fg_color=C["border"]).grid(
            row=8, column=0, columnspan=3, sticky="ew", padx=16, pady=(4, 2))

        # Insertar campo — justo debajo del asunto/cuerpo
        ctk.CTkLabel(f, text="Insertar campo:").grid(row=9, column=0, padx=16, pady=(6, 4), sticky="w")
        field_row = ctk.CTkFrame(f, fg_color="transparent")
        field_row.grid(row=9, column=1, columnspan=2, sticky="ew", padx=16, pady=(6, 4))
        field_row.grid_columnconfigure(0, weight=1)
        self.combo_fields_email = ctk.CTkComboBox(field_row, values=["(carga un Excel primero)"], state="readonly")
        self.combo_fields_email.set("(carga un Excel primero)")
        self.combo_fields_email.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        btn_row = ctk.CTkFrame(field_row, fg_color="transparent")
        btn_row.grid(row=0, column=1)
        ctk.CTkButton(btn_row, text="📋 Copiar", width=85, fg_color=C["neutral"], hover_color=C["neutral_hover"],
                      command=self._copy_field_email).pack(side="left", padx=2)
        ctk.CTkButton(btn_row, text="→ Asunto", width=85,
                      command=lambda: self._insert_field_email(self.entry_subject)).pack(side="left", padx=2)
        ctk.CTkButton(btn_row, text="→ Cuerpo", width=85,
                      command=lambda: self._insert_field_email(self.txt_body, is_textbox=True)).pack(side="left", padx=2)

        # Separador
        ctk.CTkFrame(f, height=1, fg_color=C["border"]).grid(
            row=10, column=0, columnspan=3, sticky="ew", padx=16, pady=(6, 4))

        # Cuenta + formato + modo envío
        ctk.CTkLabel(f, text="Cuenta de envío:").grid(row=11, column=0, padx=16, pady=3, sticky="w")
        acc_row = ctk.CTkFrame(f, fg_color="transparent")
        acc_row.grid(row=11, column=1, columnspan=2, sticky="ew", padx=16, pady=3)
        acc_row.grid_columnconfigure(0, weight=1)
        self.combo_account = ctk.CTkComboBox(acc_row, values=["(cargando…)"], state="readonly")
        self.combo_account.set("(cargando…)")
        self.combo_account.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(acc_row, text="↺", width=32,
                      command=self._load_outlook_accounts).grid(row=0, column=1, padx=(6, 0))

        ctk.CTkLabel(f, text="Formato archivo:").grid(row=12, column=0, padx=16, pady=3, sticky="w")
        self.output_format = ctk.StringVar(value="pdf")
        fmt_row = ctk.CTkFrame(f, fg_color="transparent")
        fmt_row.grid(row=12, column=1, columnspan=2, sticky="w", padx=16, pady=3)
        ctk.CTkRadioButton(fmt_row, text="PDF", variable=self.output_format, value="pdf").pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(fmt_row, text="Word (.docx)", variable=self.output_format, value="docx").pack(side="left")

        ctk.CTkLabel(f, text="Modo envío:").grid(row=13, column=0, padx=16, pady=3, sticky="w")
        self.send_mode = ctk.StringVar(value="draft")
        snd_row = ctk.CTkFrame(f, fg_color="transparent")
        snd_row.grid(row=13, column=1, columnspan=2, sticky="w", padx=16, pady=3)
        ctk.CTkRadioButton(snd_row, text="Guardar en Borradores", variable=self.send_mode, value="draft").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(snd_row, text="Enviar directamente", variable=self.send_mode, value="send").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(snd_row, text="Solo generar archivos", variable=self.send_mode, value="none").pack(side="left")

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 4 — Generar
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_step4(self):
        f = ctk.CTkFrame(self.content)
        self.step_frames.append(f)
        f.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(f, text="Revisa y genera",
                     font=F["title"]).grid(row=0, column=0, pady=(14, 8), padx=16, sticky="w")

        # Resumen de lo que va a ocurrir
        self.summary_card = ctk.CTkFrame(f, fg_color=C["card"], border_width=1,
                                         border_color=C["border"], corner_radius=8)
        self.summary_card.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
        self.summary_card.grid_columnconfigure(1, weight=1)
        self.summary_rows = {}
        for i, key in enumerate(["Datos", "Documento", "Salida", "Correo", "Destinatarios"]):
            ctk.CTkLabel(self.summary_card, text=key, font=F["small"], text_color=C["muted"],
                         anchor="w", width=96).grid(row=i, column=0, padx=(14, 6),
                                                     pady=(10 if i == 0 else 1, 12 if i == 4 else 1), sticky="nw")
            val = ctk.CTkLabel(self.summary_card, text="", font=F["body"], text_color=C["text"],
                               anchor="w", justify="left", wraplength=500)
            val.grid(row=i, column=1, padx=(0, 14), pady=(10 if i == 0 else 1, 12 if i == 4 else 1), sticky="w")
            self.summary_rows[key] = val
        self.lbl_send_warning = ctk.CTkLabel(
            self.summary_card, text="⚠  Los correos se enviarán directamente. No se pueden deshacer.",
            font=F["section"], text_color=C["danger"], anchor="w")
        self.lbl_send_warning.grid(row=5, column=0, columnspan=2, padx=14, pady=(0, 12), sticky="w")

        # Barra de progreso + contadores
        prog_frame = ctk.CTkFrame(f, fg_color="transparent")
        prog_frame.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 6))
        prog_frame.grid_columnconfigure(0, weight=1)
        self.progress_bar = ctk.CTkProgressBar(prog_frame, progress_color=C["primary"])
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.progress_bar.set(0)

        status_row = ctk.CTkFrame(prog_frame, fg_color="transparent")
        status_row.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.lbl_progress = ctk.CTkLabel(status_row, text="", text_color=C["muted"],
                                         font=F["small"])
        self.lbl_progress.pack(side="left", padx=(0, 14))
        self.lbl_count_ok   = ctk.CTkLabel(status_row, text="", font=F["small"], text_color=C["success"])
        self.lbl_count_warn = ctk.CTkLabel(status_row, text="", font=F["small"], text_color=C["warning"])
        self.lbl_count_err  = ctk.CTkLabel(status_row, text="", font=F["small"], text_color=C["danger"])
        for w in (self.lbl_count_ok, self.lbl_count_warn, self.lbl_count_err):
            w.pack(side="left", padx=(0, 12))
        self._counts = {"ok": 0, "warn": 0, "err": 0}

        self.is_running = False
        actions = ctk.CTkFrame(prog_frame, fg_color="transparent")
        actions.grid(row=1, column=1, sticky="e", pady=(4, 0))
        ghost = dict(height=26, fg_color="transparent", border_width=1, border_color=C["border"],
                     text_color=C["text"], hover_color=C["border"], font=F["small"])
        self.btn_open_folder = ctk.CTkButton(actions, text="Abrir carpeta", width=100,
                                             command=lambda: self._open_path(self.output_folder), **ghost)
        self.btn_open_folder.pack(side="left", padx=(0, 6))
        self.btn_open_report = ctk.CTkButton(actions, text="Ver informe", width=90, state="disabled",
                                             command=lambda: self._open_path(self._last_report), **ghost)
        self.btn_open_report.pack(side="left", padx=(0, 6))
        self._last_report = ""
        self._run_log = []
        self.btn_stop = ctk.CTkButton(actions, text="■ Detener", width=90, height=26,
                                      fg_color=C["neutral"], hover_color=C["danger_hover"],
                                      command=self._stop_generation, state="disabled")
        self.btn_stop.pack(side="left")

        # Log
        self.log_box = ctk.CTkTextbox(f, state="disabled", font=F["mono"],
                                      border_width=1, border_color=C["border"])
        self.log_box.grid(row=3, column=0, sticky="nsew", padx=16, pady=(0, 16))
        f.grid_rowconfigure(3, weight=1)
        self._apply_log_colors()

    # ═══════════════════════════════════════════════════════════════════════════
    # Resumen y contadores
    # ═══════════════════════════════════════════════════════════════════════════
    def _excel_row_count(self):
        path = self.excel_data_path
        try:
            key = (path, os.path.getmtime(path))
        except OSError:
            return None
        if getattr(self, "_row_count_key", None) != key:
            try:
                df = pd.read_excel(path, dtype=object)
                df.columns = [str(c) for c in df.columns]
                self._excel_df, self._row_count = df, len(df)
            except Exception:
                self._excel_df, self._row_count = None, None
            self._row_count_key = key
        return self._row_count

    def _email_issues(self):
        """(filas sin email, [(fila, valor no válido)]) según la columna elegida."""
        self._excel_row_count()
        df  = getattr(self, "_excel_df", None)
        col = self.entry_email_col.get().strip()
        if df is None or col not in df.columns:
            return [], []
        empty, invalid = [], []
        for n, value in enumerate(df[col].tolist(), start=1):
            text = format_cell(value)
            if not text:
                empty.append(n)
            else:
                bad = invalid_emails(text)
                if bad:
                    invalid.append((n, bad[0]))
        return empty, invalid

    def _refresh_summary(self):
        n = self._excel_row_count()
        rows_txt = f"{n} fila{'s' if n != 1 else ''}" if n is not None else "filas desconocidas"
        self.summary_rows["Datos"].configure(
            text=f"{os.path.basename(self.excel_data_path) or '—'}  ·  {rows_txt}")

        fmt = "PDF" if self.output_format.get() == "pdf" else "Word (.docx)"
        doc_txt = f"{os.path.basename(self.word_template_path) or '—'}  →  {fmt}"
        if self.enable_signature.get() and self.output_format.get() == "pdf":
            doc_txt += "  ·  firmado con " + ("PFX" if self.sign_mode.get() == "pfx" else "AutoFirma")
        self.summary_rows["Documento"].configure(text=doc_txt)
        self.summary_rows["Salida"].configure(text=self.output_folder or "—")

        s_mode  = self.send_mode.get()
        account = self.combo_account.get()
        if account.startswith("("):
            account = "la cuenta predeterminada de Outlook"
        n_att = len(getattr(self, "extra_attachments", []))
        att_txt = f"  ·  +{n_att} adjunto{'s' if n_att != 1 else ''} fijo{'s' if n_att != 1 else ''}" if n_att else ""
        if s_mode == "none":
            mail_txt, color = "No se envía ningún correo (solo se generan archivos)", C["muted"]
        elif s_mode == "draft":
            mail_txt, color = f"Se guardan como borradores en {account}{att_txt}", C["text"]
        else:
            mail_txt, color = f"Se ENVÍAN directamente desde {account}{att_txt}", C["danger"]
        self.summary_rows["Correo"].configure(text=mail_txt, text_color=color)

        if s_mode == "none":
            dest_txt = "—"
        else:
            parts = [f"columna «{self.entry_email_col.get().strip()}»"]
            for label, entry in (("Para", self.entry_to_extra), ("CC", self.entry_cc), ("CCO", self.entry_bcc)):
                v = entry.get().strip()
                if v:
                    parts.append(f"{label}: {v}")
            empty, invalid = self._email_issues()
            if empty:
                parts.append(f"⚠ {len(empty)} sin email")
            if invalid:
                parts.append(f"⚠ {len(invalid)} no válido{'s' if len(invalid) != 1 else ''}")
            dest_txt = "  ·  ".join(parts)
            dest_color = C["warning"] if (empty or invalid) else C["text"]
        self.summary_rows["Destinatarios"].configure(
            text=dest_txt, text_color=C["muted"] if s_mode == "none" else dest_color)

        if s_mode == "send":
            self.summary_card.configure(border_color=C["danger"], border_width=2)
            self.lbl_send_warning.grid()
        else:
            self.summary_card.configure(border_color=C["border"], border_width=1)
            self.lbl_send_warning.grid_remove()

    def _abort_with_problems(self, title, problems):
        """Registra los problemas en el log y avisa al usuario (seguro entre hilos)."""
        self.log(f"ERROR: {title}:")
        for pr in problems:
            self.log(f"   · {pr}", level="error")
        self.log("No se ha generado ni enviado nada. Corrige los campos y vuelve a intentarlo.",
                 level="error")
        msg = title + ":\n\n" + "\n".join("• " + pr for pr in problems) + \
            "\n\nNo se ha generado ni enviado nada."
        self.after(0, lambda: messagebox.showerror("Campos no reconocidos", msg))

    def _open_path(self, path):
        if not path or not os.path.exists(path):
            messagebox.showwarning("No encontrado", f"No se encuentra:\n{path or '(sin ruta)'}")
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showwarning("No se pudo abrir", str(e))

    def _write_report(self, records, stamp):
        """Guarda un CSV con el resultado de cada fila y el log de la ejecución.
        Devuelve la ruta del CSV, o "" si no se pudo escribir."""
        try:
            folder = os.path.join(self.output_folder, REPORT_DIR)
            os.makedirs(folder, exist_ok=True)
            csv_path = os.path.join(folder, f"informe_{stamp}.csv")
            # utf-8-sig y ';' para que Excel en español lo abra directamente
            with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.DictWriter(fh, fieldnames=["Fila", "Estado", "Archivo", "Destinatario", "Detalle"],
                                   delimiter=";")
                w.writeheader()
                w.writerows(records)
            with open(os.path.join(folder, f"log_{stamp}.txt"), "w", encoding="utf-8") as fh:
                fh.writelines(self._run_log)
            return csv_path
        except Exception as e:
            self.log(f"Aviso: no se pudo guardar el informe: {e}")
            return ""

    def _report_ready(self, path):
        self._last_report = path
        self.btn_open_report.configure(state="normal" if path else "disabled")

    def _reset_counts(self):
        self._counts = {"ok": 0, "warn": 0, "err": 0}
        self._update_counts()

    def _update_counts(self):
        c = self._counts
        pl = lambda n, one, many: f"{n} {one if n == 1 else many}"
        self.lbl_count_ok.configure(
            text=f"✓ {pl(c['ok'], 'correcto', 'correctos')}" if any(c.values()) else "")
        self.lbl_count_warn.configure(
            text=f"⚠ {pl(c['warn'], 'con aviso', 'con avisos')}" if c["warn"] else "")
        self.lbl_count_err.configure(
            text=f"✗ {pl(c['err'], 'error', 'errores')}" if c["err"] else "")

    def _set_stop_enabled(self, enabled):
        if enabled:
            self.btn_stop.configure(state="normal", fg_color=C["danger"],
                                    hover_color=C["danger_hover"])
        else:
            self.btn_stop.configure(state="disabled", fg_color=C["neutral"])

    def _stop_generation(self):
        if self.is_running:
            self.is_running = False
            self.log("Solicitando detención del proceso... (se detendrá tras el contrato actual)")
            self._set_stop_enabled(False)




    # ═══════════════════════════════════════════════════════════════════════════
    # Outlook accounts
    # ═══════════════════════════════════════════════════════════════════════════
    def _load_outlook_accounts(self):
        def fetch():
            try:
                import pythoncom
                pythoncom.CoInitialize()
                ol = win32.Dispatch("Outlook.Application")
                accounts = ol.Session.Accounts
                result = [(accounts.Item(i).DisplayName, accounts.Item(i).SmtpAddress)
                          for i in range(1, accounts.Count + 1)]
                self.outlook_accounts = result
                names = [f"{name} ({smtp})" for name, smtp in result]
                self.after(0, lambda: self._set_account_combo(names))
            except Exception:
                self.after(0, lambda: self.combo_account.configure(
                    values=["(no se pudieron cargar las cuentas)"]))
        threading.Thread(target=fetch, daemon=True).start()

    def _set_account_combo(self, names):
        if names:
            self.combo_account.configure(values=names, state="readonly")
            current = self.combo_account.get()
            saved = getattr(self, "_saved_account", "")
            # Mantener la cuenta ya elegida o la guardada; si no, la primera
            self.combo_account.set(current if current in names
                                   else saved if saved in names else names[0])
        else:
            self.combo_account.configure(values=["(sin cuentas)"])
            self.combo_account.set("(sin cuentas)")

    # ═══════════════════════════════════════════════════════════════════════════
    # Columnas Excel
    # ═══════════════════════════════════════════════════════════════════════════
    def _update_columns(self):
        try:
            cols = list(pd.read_excel(self.excel_data_path, nrows=0).columns.astype(str))
        except Exception:
            return
        self.excel_columns = cols
        self.entry_email_col.configure(values=cols)
        self.combo_fields.configure(values=cols, state="readonly")
        self.combo_fields_email.configure(values=cols, state="readonly")
        if cols:
            self.combo_fields.set(cols[0])
            self.combo_fields_email.set(cols[0])
            if self.entry_email_col.get() not in cols:
                self.entry_email_col.set(cols[0])
            self.log(f"Excel cargado: {len(cols)} columnas — {', '.join(cols)}"
                     if hasattr(self, 'log_box') else None)

    def _copy_field_to_clipboard(self):
        col = self.combo_fields.get()
        if not col or col == "(carga un Excel primero)":
            return
        self.clipboard_clear()
        self.clipboard_append(f"{{{{{col}}}}}")

    def _insert_field(self, widget, is_textbox=False):
        col = self.combo_fields.get()
        if not col or col == "(carga un Excel primero)":
            return
        tag = f"{{{{{col}}}}}"
        if is_textbox:
            widget.insert("insert", tag)
        else:
            widget.insert(widget.index("insert"), tag)
        widget.focus_set()

    def _copy_field_email(self):
        col = self.combo_fields_email.get()
        if not col or col == "(carga un Excel primero)":
            return
        self.clipboard_clear()
        self.clipboard_append(f"{{{{{col}}}}}")

    def _insert_field_email(self, widget, is_textbox=False):
        col = self.combo_fields_email.get()
        if not col or col == "(carga un Excel primero)":
            return
        tag = f"{{{{{col}}}}}"
        if is_textbox:
            widget.insert("insert", tag)
        else:
            widget.insert(widget.index("insert"), tag)
        widget.focus_set()

    # ═══════════════════════════════════════════════════════════════════════════
    # Configuración persistente
    # ═══════════════════════════════════════════════════════════════════════════
    def _restore_config(self):
        cfg = load_config()
        if not cfg:
            return
        for attr, lbl, key in [
            ("word_template_path",    self.lbl_word,    "word_template_path"),
            ("excel_data_path",       self.lbl_excel,   "excel_data_path"),
            ("output_folder",         self.lbl_output,  "output_folder"),
            ("outlook_template_path", self.lbl_oft_path,"outlook_template_path"),
        ]:
            path = cfg.get(key, "")
            if not path:
                continue
            # Conservar la ruta aunque ahora no exista (p. ej. unidad desconectada)
            setattr(self, attr, path)
            if os.path.exists(path):
                lbl.configure(text=os.path.basename(path) if os.path.isfile(path) else path,
                              text_color=C["text"])
            else:
                shown = path if key == "output_folder" else os.path.basename(path)
                lbl.configure(text=f"{shown}  (no encontrado)", text_color=C["warning"])
        if self.excel_data_path and os.path.isfile(self.excel_data_path):
            self._update_columns()
            if cfg.get("email_col"):
                self.combo_fields_email.set(cfg["email_col"])
        if cfg.get("email_col"):
            self.entry_email_col.set(cfg["email_col"])
        for entry, key in [(self.entry_to_extra, "to_extra"),
                           (self.entry_cc,       "cc"),
                           (self.entry_bcc,      "bcc")]:
            if cfg.get(key):
                entry.delete(0, "end")
                entry.insert(0, cfg[key])
        if cfg.get("filename_pattern"):
            self.entry_filename_pattern.delete(0, "end")
            self.entry_filename_pattern.insert(0, cfg["filename_pattern"])
        if cfg.get("email_subject"):
            self.entry_subject.delete(0, "end")
            self.entry_subject.insert(0, cfg["email_subject"])
        if cfg.get("email_body"):
            self.txt_body.delete("1.0", "end")
            self.txt_body.insert("1.0", cfg["email_body"])
        if cfg.get("email_mode"):
            self.email_mode.set(cfg["email_mode"])
            self.toggle_email_mode()
        if cfg.get("output_format"):
            self.output_format.set(cfg["output_format"])
        if cfg.get("send_mode"):
            self.send_mode.set(cfg["send_mode"])
        if cfg.get("appearance_mode"):
            self.combo_theme.set(cfg["appearance_mode"])
            self._change_appearance_mode(cfg["appearance_mode"])
        if cfg.get("extra_attachments"):
            self.extra_attachments = cfg["extra_attachments"]
            if self.extra_attachments:
                count = len(self.extra_attachments)
                names = [os.path.basename(f) for f in self.extra_attachments]
                text_disp = f"{count} archivo(s): " + ", ".join(names)
                if len(text_disp) > 55: text_disp = text_disp[:52] + "..."
                self.lbl_extra_att.configure(text=text_disp, text_color=C["text"])
        # Firma (la contraseña del certificado nunca se guarda)
        if cfg.get("sign_mode") in ("pfx", "autofirma"):
            self.sign_mode.set(cfg["sign_mode"])
        if cfg.get("pfx_path"):
            self.pfx_path = cfg["pfx_path"]
            if os.path.isfile(self.pfx_path):
                self.lbl_pfx_path.configure(text=os.path.basename(self.pfx_path), text_color=C["text"])
            else:
                self.lbl_pfx_path.configure(
                    text=f"{os.path.basename(self.pfx_path)}  (no encontrado)", text_color=C["warning"])
        if cfg.get("enable_signature"):
            self.enable_signature.set(True)
        self.toggle_signature_ui()
        self.toggle_sign_mode_ui()
        self._saved_account = cfg.get("outlook_account", "")

    def _save_config(self):
        if not getattr(self, "_ready", False):
            return
        account = self.combo_account.get()
        if account.startswith("("):          # "(cargando…)", "(sin cuentas)"…
            account = self._saved_account
        else:
            self._saved_account = account
        save_config({
            "enable_signature":      bool(self.enable_signature.get()),
            "sign_mode":             self.sign_mode.get(),
            "pfx_path":              self.pfx_path,
            "outlook_account":       account,
            "word_template_path":    self.word_template_path,
            "excel_data_path":       self.excel_data_path,
            "output_folder":         self.output_folder,
            "outlook_template_path": self.outlook_template_path,
            "email_col":             self.entry_email_col.get(),
            "to_extra":              self.entry_to_extra.get(),
            "cc":                    self.entry_cc.get(),
            "bcc":                   self.entry_bcc.get(),
            "filename_pattern":      self.entry_filename_pattern.get(),
            "email_subject":         self.entry_subject.get(),
            "email_body":            self.txt_body.get("1.0", "end-1c"),
            "email_mode":            self.email_mode.get(),
            "output_format":         self.output_format.get(),
            "send_mode":             self.send_mode.get(),
            "appearance_mode":       self.combo_theme.get(),
            "extra_attachments":     getattr(self, "extra_attachments", []),
        })

    # ═══════════════════════════════════════════════════════════════════════════
    # Log
    # ═══════════════════════════════════════════════════════════════════════════
    @staticmethod
    def _log_level(text):
        t = text.lower()
        if text.startswith("==="):
            return "head"
        if "error" in t:
            return "error"
        if t.startswith("aviso") or "sin email" in t or "sin firmar" in t:
            return "warn"
        if ": ok" in t:
            return "ok"
        return None

    def _apply_log_colors(self):
        dark = ctk.get_appearance_mode() == "Dark"
        idx = 1 if dark else 0
        for tag, key in (("ok", "success"), ("warn", "warning"), ("error", "danger"), ("head", "primary")):
            try:
                self.log_box.tag_config(tag, foreground=C[key][idx])
            except Exception:
                pass
        try:
            self.log_box.tag_config("time", foreground=C["muted"][idx])
        except Exception:
            pass

    def log(self, text, level=None):
        text = str(text)
        stamp = f"[{datetime.now().strftime('%H:%M:%S')}] "
        if hasattr(self, "_run_log"):
            self._run_log.append(stamp + text + "\n")
        level = level or self._log_level(text)
        if threading.current_thread() is threading.main_thread():
            self._append_log(stamp, text, level)
        else:
            # Tkinter no es thread-safe: delegar en el hilo principal
            self.after(0, self._append_log, stamp, text, level)

    def _append_log(self, stamp, text, level=None):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", stamp, "time")
        if level:
            self.log_box.insert("end", text + "\n", level)
        else:
            self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ═══════════════════════════════════════════════════════════════════════════
    # Toggle email mode
    # ═══════════════════════════════════════════════════════════════════════════
    def toggle_email_mode(self):
        if self.email_mode.get() == "manual":
            self.lbl_oft.grid_forget()
            self.lbl_oft_path.grid_forget()
            self.btn_oft.grid_forget()
            self.lbl_subj.grid(row=0, column=0, padx=8, pady=(10, 4), sticky="w")
            self.entry_subject.grid(row=0, column=1, padx=8, pady=(10, 4), sticky="ew")
            self.lbl_body.grid(row=1, column=0, padx=8, pady=4, sticky="nw")
            self.txt_body.grid(row=1, column=1, padx=8, pady=4, sticky="ew")
        else:
            self.lbl_subj.grid_forget()
            self.entry_subject.grid_forget()
            self.lbl_body.grid_forget()
            self.txt_body.grid_forget()
            self.lbl_oft.grid(row=0, column=0, padx=8, pady=(10, 4), sticky="w")
            self.lbl_oft_path.grid(row=0, column=1, padx=8, pady=(10, 4), sticky="w")
            self.btn_oft.grid(row=1, column=0, columnspan=2, padx=8, pady=4, sticky="w")

    # ═══════════════════════════════════════════════════════════════════════════
    # Selectores de archivo
    # ═══════════════════════════════════════════════════════════════════════════
    def select_word(self):
        f = filedialog.askopenfilename(title="Plantilla Word", filetypes=[("Word", "*.docx")])
        if f:
            self.word_template_path = f
            self.lbl_word.configure(text=os.path.basename(f), text_color=C["text"])
            self._save_config()

    def select_excel(self):
        f = filedialog.askopenfilename(title="Datos Excel", filetypes=[("Excel", "*.xlsx")])
        if f:
            self.excel_data_path = f
            self.lbl_excel.configure(text=os.path.basename(f), text_color=C["text"])
            self._update_columns()
            self._save_config()

    def select_output(self):
        d = filedialog.askdirectory(title="Carpeta de salida")
        if d:
            self.output_folder = d
            self.lbl_output.configure(text=d, text_color=C["text"])
            self._save_config()

    def select_oft(self):
        f = filedialog.askopenfilename(title="Plantilla Outlook", filetypes=[("Outlook Template", "*.oft")])
        if f:
            self.outlook_template_path = f
            self.lbl_oft_path.configure(text=os.path.basename(f), text_color=C["text"])
            self._save_config()

    def select_pfx(self):
        f = filedialog.askopenfilename(title="Certificado PFX/P12", filetypes=[("Certificados", "*.pfx *.p12")])
        if f:
            self.pfx_path = f
            self.lbl_pfx_path.configure(text=os.path.basename(f), text_color=C["text"])
            self._save_config()
            
    def toggle_signature_ui(self):
        if self.enable_signature.get():
            self.sign_options_frame.grid()
        else:
            self.sign_options_frame.grid_remove()

    def toggle_sign_mode_ui(self):
        if self.sign_mode.get() == "pfx":
            self.pfx_frame.pack(fill="x", pady=4)
        else:
            self.pfx_frame.pack_forget()

    def _load_signer(self, password):
        """Carga el certificado PFX. Devuelve el firmante o None si no se puede abrir."""
        encodings = ("utf-8", "latin1") if password else (None,)
        for enc in encodings:
            try:
                pwd = password.encode(enc) if enc else None
                signer = signers.SimpleSigner.load_pkcs12(self.pfx_path, passphrase=pwd)
                if signer is not None:
                    return signer
            except Exception:
                continue
        return None

    def _preflight_signing(self):
        """Comprueba la firma una sola vez antes de empezar. False = no continuar."""
        self._signer = None
        if not self.enable_signature.get():
            return True
        if self.output_format.get() != "pdf":
            return messagebox.askyesno(
                "Firma no aplicable",
                "La firma digital solo se aplica a PDF, y el formato elegido es Word.\n\n"
                "¿Continuar sin firmar?", icon="warning", default="no")
        if self.sign_mode.get() == "pfx":
            if not PYHANKO_AVAILABLE:
                messagebox.showerror("Firma", "La librería de firma (pyhanko) no está disponible.")
                return False
            if not self.pfx_path or not os.path.isfile(self.pfx_path):
                messagebox.showerror("Firma", "Selecciona el certificado PFX en el paso Configuración.")
                self._show_step(1)
                return False
            self._signer = self._load_signer(self.entry_pfx_pass.get())
            if self._signer is None:
                messagebox.showerror(
                    "Firma", "No se pudo abrir el certificado.\n\n"
                    "Comprueba la contraseña y que el archivo PFX sea válido.")
                self._show_step(1)
                return False
        else:
            if not os.path.isfile(AUTOFIRMA_EXE) and not shutil.which("AutofirmaCommandLine"):
                messagebox.showerror(
                    "Firma", "No se encuentra AutoFirma en este equipo.\n\n"
                    "Instálalo o elige firmar con certificado PFX.")
                return False
        return True

    def _sign_pdf(self, input_pdf, output_pdf):
        mode = self._run["sign_mode"]
        if mode == "pfx":
            try:
                signer = getattr(self, "_signer", None) or self._load_signer(self._run["pfx_pass"])
                if signer is None:
                    self.log("ERROR: no se pudo abrir el certificado PFX.")
                    return False

                with open(input_pdf, 'rb') as doc_b, open(output_pdf, 'wb') as out_f:
                    pdf_w = IncrementalPdfFileWriter(doc_b, strict=False)
                    signers.sign_pdf(
                        pdf_w, signers.PdfSignatureMetadata(field_name='Signature1'),
                        signer=signer,
                        output=out_f
                    )
                return True
            except Exception as e:
                self.log(f"ERROR firmando con PFX: {e}")
                return False
        elif mode == "autofirma":
            try:
                # Buscar AutofirmaCommandLine.exe en Program Files
                af_cmd = AUTOFIRMA_EXE
                if not os.path.isfile(af_cmd):
                    # Fallback: intentar desde PATH
                    af_cmd = 'AutofirmaCommandLine'
                cmd = [
                    af_cmd, "sign",
                    "-i", os.path.abspath(input_pdf),
                    "-o", os.path.abspath(output_pdf),
                    "-format", "pades",
                    "-certgui"
                ]
                self.log("Abriendo AutoFirma (selecciona tu certificado en la ventana que aparece)...")
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=120,
                                       creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0 and os.path.isfile(output_pdf):
                    return True
                else:
                    # Filtrar las advertencias de Java Preferences y mostrar solo el error real
                    err_lines = [l for l in (result.stderr or result.stdout or "").splitlines()
                                 if "WindowsPreferences" not in l and "ADVERTENCIA" not in l
                                 and "prefs root node" not in l and l.strip()]
                    err_msg = "\n".join(err_lines[-5:]) if err_lines else "Sin detalle de error"
                    self.log(f"AutoFirma error: {err_msg}")
                    return False
            except subprocess.TimeoutExpired:
                self.log("ERROR: AutoFirma excedió el tiempo de espera (2 min).")
                return False
            except FileNotFoundError:
                self.log("ERROR: No se encontró AutofirmaCommandLine.exe. Verifica que AutoFirma esté instalado.")
                return False
            except Exception as e:
                self.log(f"ERROR ejecutando AutoFirma: {e}")
                return False
        return False

    # ═══════════════════════════════════════════════════════════════════════════
    # Generación
    # ═══════════════════════════════════════════════════════════════════════════
    def start_generation(self):
        for step in range(3):
            if not self._validate_step(step):
                self._show_step(step)
                return

        if self.is_running:
            return
        if not self._preflight_signing():
            return

        s_mode = self.send_mode.get()
        if s_mode != "none":
            empty, invalid = self._email_issues()
            notes = []
            if empty:
                rows = ", ".join(map(str, empty[:8])) + ("…" if len(empty) > 8 else "")
                notes.append(f"• {len(empty)} fila(s) sin email (filas {rows})")
            if invalid:
                ex = "; ".join(f"fila {n}: {v}" for n, v in invalid[:4]) + ("…" if len(invalid) > 4 else "")
                notes.append(f"• {len(invalid)} fila(s) con email no válido ({ex})")
            n = self._excel_row_count()
            skipped = len(empty) + len(invalid)
            if s_mode == "send":
                qty = f"{max(n - skipped, 0)} correos" if n is not None else "los correos"
                msg = f"Se van a ENVIAR {qty} directamente desde Outlook.\n\n"
                if notes:
                    msg += "No se enviará nada a estas filas (el documento sí se genera):\n" \
                           + "\n".join(notes) + "\n\n"
                msg += "Esta acción no se puede deshacer. ¿Continuar?"
                if not messagebox.askyesno("Confirmar envío", msg, icon="warning", default="no"):
                    return
            elif notes:
                if not messagebox.askyesno(
                        "Revisa los destinatarios",
                        "No se creará borrador para estas filas (el documento sí se genera):\n"
                        + "\n".join(notes) + "\n\n¿Continuar?", icon="warning"):
                    return

        self._save_config()
        # Leer todos los valores de la interfaz aquí (hilo principal):
        # el hilo de trabajo no debe tocar widgets de Tkinter.
        self._run = {
            "email_col":        self.entry_email_col.get().strip(),
            "email_mode":       self.email_mode.get(),
            "send_mode":        self.send_mode.get(),
            "output_format":    self.output_format.get(),
            "account_label":    self.combo_account.get(),
            "filename_pattern": self.entry_filename_pattern.get().strip(),
            "subject":          self.entry_subject.get(),
            "body":             self.txt_body.get("1.0", "end-1c"),
            "to_extra":         self.entry_to_extra.get().strip(),
            "cc":               self.entry_cc.get().strip(),
            "bcc":              self.entry_bcc.get().strip(),
            "sign":             bool(self.enable_signature.get()),
            "sign_mode":        self.sign_mode.get(),
            "pfx_pass":         self.entry_pfx_pass.get(),
            "extra_attachments": list(getattr(self, "extra_attachments", [])),
        }
        self._show_step(3)
        self.progress_bar.set(0)
        self.lbl_progress.configure(text="")
        self._reset_counts()
        self._run_log = []
        self._report_ready("")
        self.log("=== INICIANDO PROCESO ===")
        self.is_running = True
        self._set_stop_enabled(True)
        self.btn_next.configure(state="disabled")
        self.btn_prev.configure(state="disabled")
        self._worker = threading.Thread(target=self.process_data, daemon=True)
        self._worker.start()

    def process_data(self):
        word_app = None
        com_initialized = False
        cfg = self._run
        records = []                      # una entrada por fila, para el informe
        run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            # COM debe inicializarse en este hilo siempre (Word y/o Outlook)
            try:
                import pythoncom
                pythoncom.CoInitialize()
                com_initialized = True
            except Exception as e:
                self.log(f"Aviso: no se pudo inicializar COM: {e}")

            email_col = cfg["email_col"]
            mode      = cfg["email_mode"]
            s_mode    = cfg["send_mode"]

            self.log("Leyendo Excel...")
            # dtype=object evita que pandas convierta enteros en float (123.0)
            df = pd.read_excel(self.excel_data_path, dtype=object)
            df.columns = [str(c) for c in df.columns]

            if s_mode != "none" and email_col not in df.columns:
                self.log(f"ERROR: Columna '{email_col}' no encontrada.")
                self.log(f"Columnas disponibles: {', '.join(df.columns)}")
                return

            # ── Comprobación previa de campos {{...}} ─────────────────────────
            columns = list(df.columns)
            aliases = column_aliases(columns)
            problems = []
            text_fields = [("Patrón de nombre de archivo", cfg["filename_pattern"])]
            if s_mode != "none":
                text_fields += [("Para (adicional)", cfg["to_extra"]),
                                ("CC", cfg["cc"]), ("CCO", cfg["bcc"])]
                if mode == "manual":
                    text_fields += [("Asunto", cfg["subject"]), ("Cuerpo", cfg["body"])]
            for label, text in text_fields:
                unknown = sorted(find_placeholders(text) - set(columns))
                if unknown:
                    problems.append(f"{label}: " + ", ".join("{{" + u + "}}" for u in unknown))
            try:
                tpl_vars = ColumnTemplate(self.word_template_path, aliases) \
                    .get_undeclared_template_variables()
                unknown = sorted(tpl_vars - set(columns) - set(aliases.values()))
                if unknown:
                    problems.append("Plantilla Word: " + ", ".join("{{" + u + "}}" for u in unknown))
            except Exception as e:
                problems.append(f"Plantilla Word: no se pudo analizar ({e})")
            if problems:
                self._abort_with_problems(
                    "Hay campos que no coinciden con ninguna columna del Excel", problems)
                return

            outlook      = None
            send_account = None
            if s_mode != "none":
                try:
                    outlook = win32.Dispatch("Outlook.Application")
                    selected_label = cfg["account_label"]
                    for name, smtp in self.outlook_accounts:
                        if selected_label == f"{name} ({smtp})":
                            accounts = outlook.Session.Accounts
                            for i in range(1, accounts.Count + 1):
                                if accounts.Item(i).SmtpAddress == smtp:
                                    send_account = accounts.Item(i)
                                    break
                            break
                    if send_account:
                        self.log(f"Cuenta: {selected_label}")
                except Exception as e:
                    self.log(f"ERROR iniciando Outlook: {e}")
                    return

            if s_mode != "none" and mode != "manual":
                try:
                    probe = outlook.CreateItemFromTemplate(self.outlook_template_path)
                    found = find_placeholders(probe.Subject or "")
                    # En el HTML se ignoran coincidencias partidas por etiquetas
                    found |= {html.unescape(n) for n in find_placeholders(probe.HTMLBody or "")
                              if "<" not in n and len(n) < 80}
                    try: probe.Close(1)  # olDiscard
                    except Exception: pass
                    unknown = sorted(found - set(columns))
                except Exception as e:
                    self.log(f"Aviso: no se pudo revisar la plantilla .oft: {e}")
                    unknown = []
                if unknown:
                    self._abort_with_problems(
                        "Hay campos que no coinciden con ninguna columna del Excel",
                        ["Plantilla .oft: " + ", ".join("{{" + u + "}}" for u in unknown)])
                    return

            if cfg["output_format"] == "pdf":
                try:
                    word_app = win32.Dispatch("Word.Application")
                    word_app.Visible = False
                    try:
                        word_app.DisplayAlerts = 0   # sin diálogos que bloqueen el proceso oculto
                    except Exception:
                        pass
                except Exception as e:
                    self.log(f"ERROR iniciando Word: {e}")
                    return

            rows_total = len(df)
            self.log(f"Registros a procesar: {rows_total}")
            self.after(0, lambda: self.lbl_progress.configure(text=f"0 / {rows_total}"))

            used_names = set()
            for row_num, (_, row) in enumerate(df.iterrows(), start=1):
                if not getattr(self, "is_running", True):
                    self.log("=== PROCESO DETENIDO POR EL USUARIO ===")
                    break

                row_status = "ok"
                rec = {"Fila": row_num, "Estado": "", "Archivo": "", "Destinatario": "", "Detalle": ""}
                held = []
                try:
                    context = {col: format_cell(row[col]) for col in df.columns}

                    render_ctx = dict(context)
                    render_ctx.update({alias: context[col] for col, alias in aliases.items()})
                    doc = ColumnTemplate(self.word_template_path, aliases)
                    # autoescape: sin él, caracteres como & o < desaparecen del documento
                    doc.render(render_ctx, autoescape=True)

                    name_for_file = substitute_variables(cfg["filename_pattern"], context)
                    name_for_file = re.sub(r'\{\{.*?\}\}', '', name_for_file)
                    safe_name = "".join(c for c in name_for_file
                                        if c.isalpha() or c.isdigit() or c in ' ,-_').strip()
                    if not safe_name:
                        safe_name = f"Documento_{row_num}"

                    # Evitar sobrescribir archivos existentes o de filas anteriores
                    base = safe_name
                    n = 2
                    while base.lower() in used_names:
                        base = f"{safe_name}_{n}"
                        n += 1
                    base_path = unique_base_path(
                        self.output_folder, base, [".docx", ".pdf", "_Firmado.pdf"])
                    safe_name = os.path.basename(base_path)
                    used_names.add(safe_name.lower())

                    out_docx = base_path + ".docx"
                    out_pdf  = base_path + ".pdf"
                    doc.save(out_docx)

                    if cfg["output_format"] == "pdf":
                        try:
                            self.log(f"Convirtiendo {safe_name} a PDF…")
                            wd = word_app.Documents.Open(os.path.abspath(out_docx))
                            wd.SaveAs(os.path.abspath(out_pdf), FileFormat=17)
                            wd.Close()
                            try: os.remove(out_docx)
                            except OSError: pass
                            final_path = out_pdf
                        except Exception as pdf_err:
                            self.log(f"Error PDF {safe_name}: {pdf_err}. Usando Word.")
                            row_status = "warn"
                            held.append("no se pudo convertir a PDF")
                            final_path = out_docx
                    else:
                        final_path = out_docx

                    if cfg["sign"] and final_path.endswith(".pdf"):
                        self.log(f"Firmando documento {safe_name}...")
                        signed_path = base_path + "_Firmado.pdf"
                        success = self._sign_pdf(final_path, signed_path)
                        if success:
                            try: os.remove(final_path)
                            except OSError: pass
                            final_path = signed_path
                        else:
                            self.log(f"Aviso: no se pudo firmar {safe_name}; queda sin firmar.")
                            row_status = "warn"
                            held.append("no se pudo firmar")

                    rec["Archivo"] = os.path.basename(final_path)
                    if s_mode == "none":
                        rec["Estado"] = "Generado"
                        self.log(f"Fila {row_num}: OK — {os.path.basename(final_path)}")
                        continue

                    dest_email = context.get(email_col, "").strip()
                    rec["Destinatario"] = dest_email
                    if not dest_email:
                        rec["Estado"] = "Sin email"
                        self.log(f"Fila {row_num}: sin email — archivo: {os.path.basename(final_path)}")
                        row_status = "warn"
                        continue
                    bad = invalid_emails(dest_email)
                    if bad:
                        rec["Estado"] = "Email no válido"
                        self.log(f"Fila {row_num}: email no válido «{bad[0]}» — no se crea correo; "
                                 f"archivo: {os.path.basename(final_path)}", level="warn")
                        row_status = "warn"
                        continue

                    if mode == "manual":
                        mail = outlook.CreateItem(0)
                        mail.Subject  = substitute_variables(cfg["subject"], context)
                        body_esc      = html.escape(substitute_variables(
                            cfg["body"], context)).replace("\n", "<br>")
                        mail.HTMLBody = f"<html><body>{body_esc}</body></html>"
                    else:
                        mail = outlook.CreateItemFromTemplate(self.outlook_template_path)
                        try:
                            # Escapar los valores para no romper el HTML de la plantilla
                            mail.HTMLBody = substitute_variables(
                                mail.HTMLBody, context, escape_html=True)
                        except Exception:
                            mail.Body = substitute_variables(mail.Body, context)
                        mail.Subject = substitute_variables(mail.Subject or "", context)

                    to_extra = substitute_variables(cfg["to_extra"], context)
                    cc_val   = substitute_variables(cfg["cc"], context)
                    bcc_val  = substitute_variables(cfg["bcc"], context)

                    mail.To  = "; ".join(filter(None, [dest_email, to_extra]))
                    rec["Destinatario"] = mail.To
                    if cc_val:  mail.CC  = cc_val
                    if bcc_val: mail.BCC = bcc_val
                    if send_account:
                        mail.SendUsingAccount = send_account
                    mail.Attachments.Add(os.path.abspath(final_path))

                    for att in cfg["extra_attachments"]:
                        if os.path.isfile(att):
                            mail.Attachments.Add(os.path.abspath(att))
                        else:
                            self.log(f"Aviso: adjunto no encontrado: {att}")
                            row_status = "warn"
                            held.append(f"falta el adjunto {os.path.basename(att)}")

                    if s_mode == "send" and held:
                        # Nunca enviar un correo incompleto: se deja en Borradores para revisarlo
                        mail.Save()
                        rec["Estado"] = "No enviado (borrador)"
                        self.log(f"Fila {row_num}: NO ENVIADO ({'; '.join(held)}) — "
                                 f"guardado como borrador para {mail.To}", level="warn")
                        continue

                    if s_mode == "send":
                        mail.Send()
                        action = "Enviado a"
                        rec["Estado"] = "Enviado"
                    else:
                        mail.Save()
                        action = "Borrador para"
                        rec["Estado"] = "Borrador"

                    extras = "".join([
                        f" | CC: {cc_val}"  if cc_val  else "",
                        f" | CCO: {bcc_val}" if bcc_val else "",
                    ])
                    self.log(f"Fila {row_num}: OK — {os.path.basename(final_path)} | {action} {mail.To}{extras}")

                except Exception as e:
                    self.log(f"Error fila {row_num}: {e}")
                    row_status = "err"
                    rec["Estado"] = "Error"
                    held.append(str(e))
                finally:
                    rec["Detalle"] = "; ".join(held)
                    records.append(rec)
                    self._counts[row_status] += 1
                    self.after(0, self._update_counts)
                    p = row_num / rows_total
                    self.after(0, lambda p=p, n=row_num, t=rows_total: (
                        self.progress_bar.set(p),
                        self.lbl_progress.configure(text=f"{n} / {t}")
                    ))

            # Filas que no llegaron a procesarse (proceso detenido)
            for n in range(len(records) + 1, rows_total + 1):
                records.append({"Fila": n, "Estado": "No procesado", "Archivo": "",
                                "Destinatario": "", "Detalle": "proceso detenido"})

            completed = self.is_running
            if completed:
                self.log("=== PROCESO COMPLETADO ===")
            report = self._write_report(records, run_stamp)
            records = []                  # ya guardado
            if report:
                self.log(f"Informe guardado: {os.path.join(REPORT_DIR, os.path.basename(report))}")
            self.after(0, self._report_ready, report)
            if not getattr(self, "_closing", False):
                c = dict(self._counts)
                title = "Completado" if completed else "Proceso detenido"
                msg = (f"{'El proceso ha finalizado' if completed else 'El proceso se detuvo'}.\n\n"
                       f"✓ {c['ok']} correctos   ⚠ {c['warn']} con avisos   ✗ {c['err']} errores")
                if report:
                    msg += "\n\nHay un informe detallado en la carpeta de salida (botón «Ver informe»)."
                show = messagebox.showinfo if (completed and not c["err"]) else messagebox.showwarning
                self.after(0, lambda: show(title, msg))

        except Exception as e:
            self.log(f"Error general: {e}")
            if records:                   # guardar lo procesado hasta el fallo
                self.after(0, self._report_ready, self._write_report(records, run_stamp))
        finally:
            self.is_running = False
            self.after(0, lambda: self._set_stop_enabled(False))
            self.after(0, lambda: self.btn_next.configure(state="normal"))
            self.after(0, lambda: self.btn_prev.configure(state="normal"))
            try:
                if word_app is not None:
                    word_app.Quit()
            except Exception:
                pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

if __name__ == "__main__":
    app = App()
    app.mainloop()
