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
from auto_updater import check_for_updates
from version import VERSION

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".generador_contratos_config.json")

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

def substitute_variables(text, context):
    for k, v in context.items():
        text = text.replace("{{" + k + "}}", str(v))
    return text

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

STEPS = ["Archivos", "Configuración", "Correo", "Generar"]

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Contratos y Correos")
        self.geometry("700x620")
        self.resizable(False, False)

        self.after(3000, lambda: check_for_updates(self))

        # --- Variables de datos ---
        self.word_template_path   = ""
        self.excel_data_path      = ""
        self.outlook_template_path = ""
        self.output_folder        = ""
        self.excel_columns        = []
        self.outlook_accounts     = []
        self.current_step         = 0

        # --- Layout raíz ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── Cabecera ──────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color=("gray85", "gray20"), corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(hdr, text="Generador de Contratos y Correos",
                     font=("System", 16, "bold")).grid(row=0, column=0, pady=12, padx=20, sticky="w")
        ctk.CTkLabel(hdr, text=f"v{VERSION}", text_color="gray",
                     font=("System", 11)).grid(row=0, column=1, pady=12, padx=20, sticky="e")

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
                                      fg_color="gray50", hover_color="gray40",
                                      command=self._prev_step)
        self.btn_prev.grid(row=0, column=0, padx=(0, 8))

        ctk.CTkLabel(nav, text="", font=("System", 11)).grid(row=0, column=1)  # spacer

        self.btn_next = ctk.CTkButton(nav, text="Siguiente →", width=120,
                                      command=self._next_step)
        self.btn_next.grid(row=0, column=2, padx=(8, 0))

        # Versión + actualización (pie derecho)
        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 8))
        foot.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(foot, text="Buscar actualizaciones", width=160, height=24,
                      font=("System", 11), fg_color="transparent", border_width=1,
                      text_color=("gray40", "gray60"),
                      command=lambda: check_for_updates(self)).grid(row=0, column=1, sticky="e")

        self._restore_config()
        self._show_step(0)
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
            dot = ctk.CTkLabel(self.ind_frame, text="●", font=("System", 18),
                               text_color="gray60", width=28)
            dot.grid(row=0, column=col)
            lbl = ctk.CTkLabel(self.ind_frame, text=name, font=("System", 11),
                               text_color="gray60")
            lbl.grid(row=1, column=col)
            self.ind_dots.append(dot)
            self.ind_labels.append(lbl)
            if i < len(STEPS) - 1:
                ctk.CTkLabel(self.ind_frame, text="────", text_color="gray50",
                             font=("System", 11)).grid(row=0, column=col + 1, sticky="ew")

    def _update_step_indicator(self, step):
        for i, (dot, lbl) in enumerate(zip(self.ind_dots, self.ind_labels)):
            if i < step:
                dot.configure(text="✓", text_color=("green3", "green2"))
                lbl.configure(text_color=("gray50", "gray50"))
            elif i == step:
                dot.configure(text="●", text_color=("dodger blue", "dodger blue"))
                lbl.configure(text_color=("black", "white"),
                               font=("System", 11, "bold"))
            else:
                dot.configure(text="●", text_color="gray60")
                lbl.configure(text_color="gray60",
                               font=("System", 11))

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
            self.btn_next.configure(text="✓ Generar", fg_color=("green4", "green3"),
                                    hover_color=("green3", "green2"),
                                    command=self.start_generation)
        else:
            self.btn_next.configure(text="Siguiente →", fg_color=("#1f6aa5", "#1f538a"),
                                    hover_color=("#144870", "#144870"),
                                    command=self._next_step)

    def _next_step(self):
        if self._validate_step(self.current_step):
            self._show_step(self.current_step + 1)

    def _prev_step(self):
        self._show_step(self.current_step - 1)

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
        elif step == 1:
            if not self.entry_email_col.get().strip():
                messagebox.showwarning("Faltan datos", "Indica la columna de email.")
                return False
            if not self.entry_filename_pattern.get().strip():
                messagebox.showwarning("Faltan datos", "Indica el patrón de nombre de archivo.")
                return False
        elif step == 2:
            if self.email_mode.get() == "template" and not self.outlook_template_path:
                messagebox.showwarning("Faltan datos", "Selecciona la plantilla .oft de Outlook.")
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
                     font=("System", 13, "bold")).grid(
            row=0, column=0, columnspan=3, pady=(16, 20), padx=16, sticky="w")

        rows = [
            ("Plantilla Word (.docx)", "lbl_word",   self.select_word,   "📄"),
            ("Datos Excel (.xlsx)",    "lbl_excel",  self.select_excel,  "📊"),
            ("Carpeta de salida",      "lbl_output", self.select_output, "📁"),
        ]
        for i, (label, attr, cmd, icon) in enumerate(rows, start=1):
            ctk.CTkLabel(f, text=f"{icon}  {label}",
                         font=("System", 12)).grid(row=i*2-1, column=0, columnspan=3,
                                                    padx=16, pady=(10, 2), sticky="w")
            lbl = ctk.CTkLabel(f, text="Sin seleccionar", text_color="gray",
                               font=("System", 11), anchor="w")
            lbl.grid(row=i*2, column=0, columnspan=2, padx=20, pady=(0, 4), sticky="ew")
            setattr(self, attr, lbl)
            ctk.CTkButton(f, text="Seleccionar", width=110, command=cmd).grid(
                row=i*2, column=2, padx=(0, 16), pady=(0, 4))

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 2 — Configuración Excel
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_step2(self):
        f = ctk.CTkFrame(self.content)
        self.step_frames.append(f)
        f.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(f, text="Configura el nombre de los archivos",
                     font=("System", 13, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(16, 4), padx=16, sticky="w")
        ctk.CTkLabel(f, text="Usa {{NombreColumna}} para insertar datos del Excel en el nombre del archivo y en la plantilla Word.",
                     text_color="gray", font=("System", 11), wraplength=580, justify="left").grid(
            row=1, column=0, columnspan=2, padx=16, pady=(0, 16), sticky="w")

        # Patrón nombre archivo
        ctk.CTkLabel(f, text="Patrón nombre archivo:").grid(row=2, column=0, padx=16, pady=6, sticky="w")
        self.entry_filename_pattern = ctk.CTkEntry(f, placeholder_text="Ej: {{Apellidos}}, {{Nombre}}")
        self.entry_filename_pattern.insert(0, "{{Apellidos}}, {{Nombre}}")
        self.entry_filename_pattern.grid(row=2, column=1, padx=16, pady=6, sticky="ew")

        # Separador
        ctk.CTkFrame(f, height=1, fg_color="gray70").grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=16, pady=12)

        # Insertar campo
        ctk.CTkLabel(f, text="Insertar campo:", font=("System", 12, "bold")).grid(
            row=4, column=0, padx=16, pady=(0, 6), sticky="w")
        ctk.CTkLabel(f, text="Selecciona un campo y cópialo para pegarlo en la plantilla Word o en el patrón.",
                     text_color="gray", font=("System", 11)).grid(
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
        ctk.CTkButton(btn_frame, text="📋 Copiar", width=90, fg_color="gray40",
                      command=self._copy_field_to_clipboard).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="→ Patrón", width=90,
                      command=lambda: self._insert_field(self.entry_filename_pattern)).pack(side="left", padx=2)

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 3 — Correo
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_step3(self):
        f = ctk.CTkScrollableFrame(self.content)
        self.step_frames.append(f)
        # f.grid_columnconfigure(1, weight=1)  # INCORRECTO: La columna 1 es para la scrollbar en CTKScrollableFrame.

        ctk.CTkLabel(f, text="Configura el correo electrónico",
                     font=("System", 13, "bold")).grid(
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
                                          text_color="gray")
        self.btn_oft      = ctk.CTkButton(self.email_inner, text="Seleccionar .oft",
                                           command=self.select_oft)
        self.toggle_email_mode()

        # ── Destinatarios ──────────────────────────────────────────────────────
        ctk.CTkFrame(f, height=1, fg_color="gray70").grid(
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

        ctk.CTkFrame(f, height=1, fg_color="gray70").grid(
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
        ctk.CTkButton(btn_row, text="📋 Copiar", width=85, fg_color="gray40",
                      command=self._copy_field_email).pack(side="left", padx=2)
        ctk.CTkButton(btn_row, text="→ Asunto", width=85,
                      command=lambda: self._insert_field_email(self.entry_subject)).pack(side="left", padx=2)
        ctk.CTkButton(btn_row, text="→ Cuerpo", width=85,
                      command=lambda: self._insert_field_email(self.txt_body, is_textbox=True)).pack(side="left", padx=2)

        # Separador
        ctk.CTkFrame(f, height=1, fg_color="gray70").grid(
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
        f.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(f, text="Todo listo — pulsa Generar para empezar",
                     font=("System", 13, "bold")).grid(
            row=0, column=0, pady=(16, 4), padx=16, sticky="w")
        ctk.CTkLabel(f, text="El progreso aparecerá en el log de abajo.",
                     text_color="gray", font=("System", 11)).grid(
            row=1, column=0, padx=16, pady=(0, 10), sticky="w")

        # Barra de progreso
        prog_frame = ctk.CTkFrame(f, fg_color="transparent")
        prog_frame.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 6))
        prog_frame.grid_columnconfigure(0, weight=1)
        self.progress_bar = ctk.CTkProgressBar(prog_frame)
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_bar.set(0)
        self.lbl_progress = ctk.CTkLabel(prog_frame, text="", text_color="gray",
                                          font=("System", 11))
        self.lbl_progress.grid(row=1, column=0, sticky="w", pady=(2, 0))

        # Log
        self.log_box = ctk.CTkTextbox(f, state="disabled")
        self.log_box.grid(row=3, column=0, sticky="nsew", padx=16, pady=(0, 16))
        f.grid_rowconfigure(3, weight=1)


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
            self.combo_account.set(names[0])
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
            if path and os.path.exists(path):
                setattr(self, attr, path)
                lbl.configure(text=os.path.basename(path) if os.path.isfile(path) else path,
                               text_color="black")
        if self.excel_data_path:
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

    def _save_config(self):
        save_config({
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
        })

    # ═══════════════════════════════════════════════════════════════════════════
    # Log
    # ═══════════════════════════════════════════════════════════════════════════
    def log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
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
            self.lbl_word.configure(text=os.path.basename(f), text_color="black")
            self._save_config()

    def select_excel(self):
        f = filedialog.askopenfilename(title="Datos Excel", filetypes=[("Excel", "*.xlsx")])
        if f:
            self.excel_data_path = f
            self.lbl_excel.configure(text=os.path.basename(f), text_color="black")
            self._update_columns()
            self._save_config()

    def select_output(self):
        d = filedialog.askdirectory(title="Carpeta de salida")
        if d:
            self.output_folder = d
            self.lbl_output.configure(text=d, text_color="black")
            self._save_config()

    def select_oft(self):
        f = filedialog.askopenfilename(title="Plantilla Outlook", filetypes=[("Outlook Template", "*.oft")])
        if f:
            self.outlook_template_path = f
            self.lbl_oft_path.configure(text=os.path.basename(f), text_color="black")
            self._save_config()

    # ═══════════════════════════════════════════════════════════════════════════
    # Generación
    # ═══════════════════════════════════════════════════════════════════════════
    def start_generation(self):
        for step in range(3):
            if not self._validate_step(step):
                self._show_step(step)
                return

        self._save_config()
        self._show_step(3)
        self.progress_bar.set(0)
        self.lbl_progress.configure(text="")
        self.log("=== INICIANDO PROCESO ===")
        threading.Thread(target=self.process_data, daemon=True).start()

    def process_data(self):
        try:
            email_col = self.entry_email_col.get().strip()
            mode      = self.email_mode.get()
            s_mode    = self.send_mode.get()

            self.log("Leyendo Excel...")
            df = pd.read_excel(self.excel_data_path)

            if email_col not in df.columns:
                self.log(f"ERROR: Columna '{email_col}' no encontrada.")
                self.log(f"Columnas disponibles: {', '.join(df.columns)}")
                return

            outlook      = None
            send_account = None
            if s_mode != "none":
                try:
                    import pythoncom
                    pythoncom.CoInitialize()
                    outlook = win32.Dispatch("Outlook.Application")
                    selected_label = self.combo_account.get()
                    for name, smtp in self.outlook_accounts:
                        if selected_label.startswith(name):
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

            try:
                word_app = None
                if self.output_format.get() == "pdf":
                    word_app = win32.Dispatch("Word.Application")
                    word_app.Visible = False
            except Exception as e:
                self.log(f"ERROR iniciando Word: {e}")
                return

            rows_total = len(df)
            self.log(f"Registros a procesar: {rows_total}")
            self.after(0, lambda: self.lbl_progress.configure(text=f"0 / {rows_total}"))

            for index, row in df.iterrows():
                row_num = index + 1
                try:
                    context = {str(col): ("" if pd.isna(row[col]) else str(row[col]))
                               for col in df.columns}

                    doc = DocxTemplate(self.word_template_path)
                    doc.render(context)

                    pattern      = self.entry_filename_pattern.get().strip()
                    name_for_file = substitute_variables(pattern, context)
                    name_for_file = re.sub(r'\{\{.*?\}\}', '', name_for_file)
                    if not name_for_file.strip():
                        name_for_file = f"Contrato_{row_num}"
                    safe_name = "".join(c for c in name_for_file
                                        if c.isalpha() or c.isdigit() or c in ' ,-_').strip()

                    out_docx = os.path.join(self.output_folder, f"{safe_name}.docx")
                    out_pdf  = os.path.join(self.output_folder, f"{safe_name}.pdf")
                    doc.save(out_docx)

                    if self.output_format.get() == "pdf":
                        try:
                            self.log(f"Convirtiendo {safe_name} a PDF…")
                            wd = word_app.Documents.Open(os.path.abspath(out_docx))
                            wd.SaveAs(os.path.abspath(out_pdf), FileFormat=17)
                            wd.Close()
                            try: os.remove(out_docx)
                            except: pass
                            final_path = out_pdf
                        except Exception as pdf_err:
                            self.log(f"Error PDF {safe_name}: {pdf_err}. Usando Word.")
                            final_path = out_docx
                    else:
                        final_path = out_docx

                    if s_mode == "none":
                        self.log(f"Fila {row_num}: OK — {os.path.basename(final_path)}")
                        continue

                    dest_email = context.get(email_col, "").strip()
                    if not dest_email:
                        self.log(f"Fila {row_num}: sin email — archivo: {os.path.basename(final_path)}")
                        continue

                    if mode == "manual":
                        mail = outlook.CreateItem(0)
                        mail.Subject  = substitute_variables(self.entry_subject.get(), context)
                        body_esc      = html.escape(substitute_variables(
                            self.txt_body.get("1.0", "end-1c"), context)).replace("\n", "<br>")
                        mail.HTMLBody = f"<html><body>{body_esc}</body></html>"
                    else:
                        mail = outlook.CreateItemFromTemplate(self.outlook_template_path)
                        try:
                            mail.HTMLBody = substitute_variables(mail.HTMLBody, context)
                        except:
                            mail.Body = substitute_variables(mail.Body, context)
                        mail.Subject = substitute_variables(mail.Subject or "", context)

                    to_extra = substitute_variables(self.entry_to_extra.get().strip(), context)
                    cc_val   = substitute_variables(self.entry_cc.get().strip(), context)
                    bcc_val  = substitute_variables(self.entry_bcc.get().strip(), context)

                    mail.To  = "; ".join(filter(None, [dest_email, to_extra]))
                    if cc_val:  mail.CC  = cc_val
                    if bcc_val: mail.BCC = bcc_val
                    if send_account:
                        mail.SendUsingAccount = send_account
                    mail.Attachments.Add(os.path.abspath(final_path))

                    if s_mode == "send":
                        mail.Send()
                        action = "Enviado a"
                    else:
                        mail.Save()
                        action = "Borrador para"

                    extras = "".join([
                        f" | CC: {cc_val}"  if cc_val  else "",
                        f" | CCO: {bcc_val}" if bcc_val else "",
                    ])
                    self.log(f"Fila {row_num}: OK — {os.path.basename(final_path)} | {action} {mail.To}{extras}")

                except Exception as e:
                    self.log(f"Error fila {row_num}: {e}")
                finally:
                    p = row_num / rows_total
                    self.after(0, lambda p=p, n=row_num, t=rows_total: (
                        self.progress_bar.set(p),
                        self.lbl_progress.configure(text=f"{n} / {t}")
                    ))

            self.log("=== PROCESO COMPLETADO ===")
            messagebox.showinfo("Completado", "El proceso ha finalizado.")

        except Exception as e:
            self.log(f"Error general: {e}")
        finally:
            try:
                if word_app is not None:
                    word_app.Quit()
            except:
                pass

if __name__ == "__main__":
    app = App()
    app.mainloop()
