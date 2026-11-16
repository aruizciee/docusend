"""Tests de las funciones sin interfaz."""
from datetime import datetime

import docx
import pandas as pd
import pytest

import docusend
from auto_updater import _parse_version


class TestFormatCell:
    @pytest.mark.parametrize("value,expected", [
        (None, ""),
        (float("nan"), ""),
        (pd.NaT, ""),
        (123.0, "123"),
        (1.5, "1.5"),
        (7, "7"),
        ("08001", "08001"),
        ("  texto  ", "texto"),
        (True, "Sí"),
        (pd.Timestamp("2026-03-01"), "01/03/2026"),
        (datetime(2026, 3, 1, 9, 30), "01/03/2026 09:30"),
    ])
    def test_values(self, value, expected):
        assert docusend.format_cell(value) == expected


class TestSubstituteVariables:
    def test_basic_and_spaces_inside_braces(self):
        ctx = {"Nombre": "Ana", "Nombre completo": "Ana Pérez"}
        assert docusend.substitute_variables("Hola {{Nombre}} / {{ Nombre completo }}", ctx) \
            == "Hola Ana / Ana Pérez"

    def test_unknown_placeholder_is_kept(self):
        assert docusend.substitute_variables("{{Nombr}}", {"Nombre": "Ana"}) == "{{Nombr}}"

    def test_value_is_not_substituted_again(self):
        ctx = {"A": "{{B}}", "B": "secreto"}
        assert docusend.substitute_variables("{{A}}", ctx) == "{{B}}"

    def test_html_escape(self):
        out = docusend.substitute_variables("<p>{{E}}</p>", {"E": "R&D <x>"}, escape_html=True)
        assert out == "<p>R&amp;D &lt;x&gt;</p>"

    def test_find_placeholders(self):
        assert docusend.find_placeholders("{{A}} y {{ B c }} y nada") == {"A", "B c"}


class TestEmails:
    @pytest.mark.parametrize("text", ["a@ej.com", "a.b@sub.ej.co.uk", "a@ej.com; b@ej.es", ""])
    def test_valid(self, text):
        assert docusend.invalid_emails(text) == []

    @pytest.mark.parametrize("text,bad", [
        ("a@ej", ["a@ej"]),
        ("sin arroba", ["sin arroba"]),
        ("a@ej.com; malo@", ["malo@"]),
        ("a@@ej.com", ["a@@ej.com"]),
    ])
    def test_invalid(self, text, bad):
        assert docusend.invalid_emails(text) == bad


class TestRowsAndNames:
    @pytest.mark.parametrize("total,rng,expected", [
        (5, (None, None), [1, 2, 3, 4, 5]),
        (5, (2, 3), [2, 3]),
        (5, (4, None), [4, 5]),
        (5, (None, 2), [1, 2]),
        (5, (4, 99), [4, 5]),
        (5, (9, None), []),
    ])
    def test_rows_in_range(self, total, rng, expected):
        assert docusend.App._rows_in_range(total, rng) == expected

    def test_unique_base_path(self, tmp_path):
        (tmp_path / "Juan.pdf").write_text("x")
        (tmp_path / "Juan_2.docx").write_text("x")
        path = docusend.unique_base_path(str(tmp_path), "Juan", [".docx", ".pdf"])
        assert path == str(tmp_path / "Juan_3")


class TestWordTemplate:
    def _template(self, tmp_path, text):
        document = docx.Document()
        document.add_paragraph(text)
        path = tmp_path / "plantilla.docx"
        document.save(path)
        return str(path)

    def _render(self, tmp_path, text, context):
        aliases = docusend.column_aliases(list(context))
        tpl = docusend.ColumnTemplate(self._template(tmp_path, text), aliases)
        ctx = dict(context)
        ctx.update({alias: context[col] for col, alias in aliases.items()})
        tpl.render(ctx, autoescape=True)
        out = tmp_path / "salida.docx"
        tpl.save(out)
        return "\n".join(p.text for p in docx.Document(out).paragraphs)

    def test_aliases_only_for_invalid_names(self):
        aliases = docusend.column_aliases(["Nombre", "Dirección", "Nombre completo", "Fecha-inicio", "for"])
        assert set(aliases) == {"Nombre completo", "Fecha-inicio", "for"}

    def test_columns_with_spaces_hyphens_and_accents(self, tmp_path):
        text = self._render(
            tmp_path, "{{Nombre completo}} / {{ Fecha-inicio }} / {{Dirección}}",
            {"Nombre completo": "Ana Pérez", "Fecha-inicio": "01/03/2026", "Dirección": "C/ Ñu 1"})
        assert text == "Ana Pérez / 01/03/2026 / C/ Ñu 1"

    def test_special_characters_are_preserved(self, tmp_path):
        text = self._render(tmp_path, "Empresa: {{Empresa}}", {"Empresa": "Pérez & Hijos <SL>"})
        assert text == "Empresa: Pérez & Hijos <SL>"

    def test_undeclared_variables_detect_typos(self, tmp_path):
        columns = ["Nombre", "Nombre completo"]
        aliases = docusend.column_aliases(columns)
        tpl = docusend.ColumnTemplate(
            self._template(tmp_path, "{{Nombre}} {{Nombre completo}} {{Nombr}}"), aliases)
        unknown = tpl.get_undeclared_template_variables() - set(columns) - set(aliases.values())
        assert unknown == {"Nombr"}


class TestUpdaterVersion:
    def test_parse(self):
        assert _parse_version("v2026.03.23-4") == (2026, 3, 23, 4)

    def test_order(self):
        assert _parse_version("v2026.11.16-10") > _parse_version("v2026.11.16-9")
        assert _parse_version("v2026.11.16-1") > _parse_version("v0.0.0-dev")

    def test_garbage_does_not_crash(self):
        assert _parse_version("no-es-una-version") == (0,)
