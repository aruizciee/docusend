"""Tests de extremo a extremo con la aplicación real y un Outlook de mentira.

Se genera siempre en formato Word, que no necesita Microsoft Word instalado.
Se omiten automáticamente si no hay entorno gráfico.
"""
import csv
import json
import os
import time

import docx
import pandas as pd
import pytest

import docusend


def make_template(folder, text, name="plantilla.docx"):
    document = docx.Document()
    document.add_paragraph(text)
    path = folder / name
    document.save(path)
    return str(path)


def set_entry(entry, value):
    entry.delete(0, "end")
    entry.insert(0, value)


def run(app, test=False, timeout=30):
    """Lanza la generación y ejecuta el bucle de la interfaz hasta que termina.

    El hilo de trabajo se comunica con la ventana mediante after(), que solo
    funciona si el hilo principal está dentro de mainloop().
    """
    deadline = time.time() + timeout

    def poll():
        if app.is_running and time.time() < deadline:
            app.after(20, poll)
        else:
            app.after(150, app.quit)   # margen para los after() pendientes

    def kick():
        # Se lanza ya dentro de mainloop() para que el hilo de trabajo nunca
        # encuentre la ventana fuera de su bucle de eventos.
        app.start_generation(test=test)
        app.after(20, poll)

    app.after(0, kick)
    app.mainloop()
    assert not app.is_running, "la generación no terminó a tiempo"


def log_text(app):
    return app.log_box.get("1.0", "end")


def read_report(app):
    with open(app._last_report, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh, delimiter=";"))


@pytest.fixture
def project(app, tmp_path):
    """App configurada con una plantilla, un Excel de 5 filas y carpeta de salida."""
    out = tmp_path / "salida"
    out.mkdir()
    excel = tmp_path / "datos.xlsx"
    with pd.ExcelWriter(excel) as writer:
        pd.DataFrame({
            "Nombre": ["Ana", "Ana", "Luis", "Marta", "Eva"],
            "Empresa": ["Pérez & Hijos <SL>", "ACME", "ACME", "ACME", "ACME"],
            "Fecha": [pd.Timestamp("2026-03-01")] * 5,
            "CP": ["08001", "28001", "41001", "46001", "50001"],
            "Email": ["ana@ej.com", None, "luis@ej", "marta@ej.com", "eva@ej.com"],
        }).to_excel(writer, sheet_name="Datos", index=False)
        pd.DataFrame({"Otra": [1]}).to_excel(writer, sheet_name="Notas", index=False)
    app.word_template_path = make_template(tmp_path, "Hola {{Nombre}} de {{Empresa}}, {{Fecha}}, CP {{CP}}")
    app.excel_data_path = str(excel)
    app.output_folder = str(out)
    app._load_sheets("Datos")
    app._update_columns()
    set_entry(app.entry_filename_pattern, "{{Nombre}}")
    set_entry(app.entry_subject, "Contrato {{ Nombre }}")
    app.txt_body.insert("1.0", "Empresa: {{Empresa}}")
    app.entry_email_col.set("Email")
    app.output_format.set("docx")
    app.send_mode.set("none")
    app.out = out
    return app


def test_generate_only(project):
    run(project)
    files = sorted(p for p in os.listdir(project.out) if p.endswith(".docx"))
    # El nombre repetido no se sobrescribe
    assert files == ["Ana.docx", "Ana_2.docx", "Eva.docx", "Luis.docx", "Marta.docx"]
    text = docx.Document(project.out / "Ana.docx").paragraphs[0].text
    assert text == "Hola Ana de Pérez & Hijos <SL>, 01/03/2026, CP 08001"
    assert project._counts == {"ok": 5, "warn": 0, "err": 0}
    assert [r["Estado"] for r in read_report(project)] == ["Generado"] * 5


def test_send_skips_empty_and_invalid_emails(project, outlook):
    project.send_mode.set("send")
    run(project)
    assert [m.To for m in outlook.log["sent"]] == ["ana@ej.com", "marta@ej.com", "eva@ej.com"]
    assert outlook.log["sent"][0].Subject == "Contrato Ana"
    assert "Pérez &amp; Hijos &lt;SL&gt;" in outlook.log["sent"][0].HTMLBody
    estados = [r["Estado"] for r in read_report(project)]
    assert estados == ["Enviado", "Sin email", "Email no válido", "Enviado", "Enviado"]
    # La confirmación de envío avisa de las filas que se saltan
    confirm = [a for kind, a in project.dialogs.shown if kind == "askyesno"][-1][1]
    assert "3 correos" in confirm and "sin email" in confirm and "no válido" in confirm


def test_send_declined_does_nothing(project, outlook):
    project.send_mode.set("send")
    project.dialogs.answers["yes"] = False
    project.start_generation()
    assert not project.is_running
    assert outlook.log["sent"] == [] and os.listdir(project.out) == []


def test_missing_attachment_is_not_sent(project, outlook):
    project.send_mode.set("send")
    project.extra_attachments = [str(project.out / "no_existe.pdf")]
    run(project)
    assert outlook.log["sent"] == []
    assert len(outlook.log["saved"]) == 3
    assert "No enviado (borrador)" in [r["Estado"] for r in read_report(project)]


def test_unknown_field_in_subject_aborts(project, outlook):
    project.send_mode.set("send")
    set_entry(project.entry_subject, "Hola {{Nombr}}")
    run(project)
    assert outlook.log["sent"] == [] and outlook.log["saved"] == []
    assert os.listdir(project.out) == []
    assert "{{Nombr}}" in log_text(project)


def test_unknown_field_in_word_template_aborts(project, tmp_path):
    project.word_template_path = make_template(tmp_path, "Hola {{Nombr}}", "mala.docx")
    run(project)
    assert os.listdir(project.out) == []
    assert "Plantilla Word" in log_text(project)


def test_row_range(project, outlook):
    project.send_mode.set("draft")
    set_entry(project.entry_row_from, "4")
    set_entry(project.entry_row_to, "5")
    run(project)
    assert [m.To for m in outlook.log["saved"]] == ["marta@ej.com", "eva@ej.com"]
    assert [r["Fila"] for r in read_report(project)] == ["4", "5"]


def test_invalid_row_range_is_rejected(project):
    set_entry(project.entry_row_from, "5")
    set_entry(project.entry_row_to, "2")
    assert project._validate_step(0) is False


def test_sheet_selection(project):
    assert project._sheet_names == ["Datos", "Notas"]
    project.combo_sheet.set("Notas")
    project._on_sheet_change()
    assert project.excel_columns == ["Otra"]
    assert project._excel_row_count() == 1


def test_trial_run_sends_nothing(project, outlook):
    project.send_mode.set("send")
    run(project, test=True)
    assert [m.To for m in outlook.log["shown"]] == ["ana@ej.com"]
    assert outlook.log["sent"] == [] and outlook.log["saved"] == []
    assert os.listdir(project.out) == [docusend.TEST_DIR]
    assert os.listdir(project.out / docusend.TEST_DIR) == ["Ana.docx"]


def test_trial_run_opens_document(project):
    run(project, test=True)
    assert [os.path.basename(p) for p in project.opened] == ["Ana.docx"]


def test_exact_account_match(project, outlook):
    project.outlook_accounts = [("Ana", "ana@x.org"), ("Ana Maria", "anamaria@x.org")]
    project._set_account_combo(["Ana (ana@x.org)", "Ana Maria (anamaria@x.org)"])
    project.combo_account.set("Ana Maria (anamaria@x.org)")
    project.send_mode.set("send")
    run(project)
    assert outlook.log["sent"][0].SendUsingAccount.SmtpAddress == "anamaria@x.org"


def test_config_survives_open_and_close(project, monkeypatch):
    project.extra_attachments = ["adjunto.pdf"]
    project.enable_signature.set(True)
    project.sign_mode.set("autofirma")
    project.word_template_path = "Z:/unidad_desconectada/plantilla.docx"
    project._save_config()
    with open(docusend.CONFIG_PATH, encoding="utf-8") as fh:
        saved = json.load(fh)
    assert saved["extra_attachments"] == ["adjunto.pdf"]
    assert saved["excel_sheet"] == "Datos"

    second = docusend.App()
    try:
        second.update()
        # Abrir y cerrar no debe perder nada, ni siquiera rutas que ahora no existen
        assert second.extra_attachments == ["adjunto.pdf"]
        assert second.word_template_path == "Z:/unidad_desconectada/plantilla.docx"
        assert "(no encontrado)" in second.lbl_word.cget("text")
        assert second.enable_signature.get() and second.sign_mode.get() == "autofirma"
        second._save_config()
    finally:
        second.destroy()
    with open(docusend.CONFIG_PATH, encoding="utf-8") as fh:
        assert json.load(fh) == saved


def test_saved_excel_is_restored_after_window_shows(project):
    """El Excel guardado se lee tras mostrar la ventana, no durante el arranque."""
    project.combo_sheet.set("Datos")
    project._save_config()

    second = docusend.App()
    try:
        # Recién creada: todavía no ha leído el Excel, pero ya sabe qué hoja usar
        assert second.excel_columns == []
        assert second._sheet() == "Datos"
        second.after(0, second._warm_up)
        deadline = time.time() + 30

        def poll():
            if second._excel_restore_pending and time.time() < deadline:
                second.after(20, poll)
            else:
                second.after(50, second.quit)
        second.after(10, poll)
        second.mainloop()
        assert second.excel_columns == ["Nombre", "Empresa", "Fecha", "CP", "Email"]
        assert second.combo_sheet.get() == "Datos"
        assert second.combo_fields_email.get() == "Email"
    finally:
        second.destroy()
