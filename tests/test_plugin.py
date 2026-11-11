import asyncio
import json
import os
import sys
import zipfile
from pathlib import Path

import pytest
from astra_plugin_sdk import BadArguments, Timeout
from astra_plugin_sdk.testing import Harness, WireHarness, fuzz_configs
from docx import Document
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation

from src import excel
from src.plugin import OfficeFiles
from src.registry import TOOLS
from src.runtime import Runner
from src.safety import package, preservation


@pytest.fixture
def h():
    with Harness(OfficeFiles(), timeout=150).start() as harness:
        yield harness


def call(h, tool_name, **args):
    result = h.call_tool(
        tool_name, **{k: str(v) if isinstance(v, Path) else v for k, v in args.items()}
    )
    assert result.success, (result.code, result.error)
    return result.json


def image(tmp_path):
    p = tmp_path / "logo.png"
    Image.new("RGB", (80, 40), "#3478ab").save(p)
    return str(p)


def word_spec(tmp_path):
    return {
        "header": "Отчёт • Office Files",
        "footer": "Конфиденциально",
        "page": {"margin_left_mm": 20},
        "styles": {"Accent": {"size": 14, "color": "225599"}},
        "blocks": [
            {"type": "heading", "text": "Квартальный отчёт", "level": 1},
            {
                "type": "paragraph",
                "runs": [
                    {"text": "При", "font": {"bold": True}},
                    {"text": "вет мир", "font": {"italic": True}},
                ],
            },
            {"type": "table", "rows": [["Показатель", "Сумма"], ["Выручка", 42]]},
            {"type": "bullet_list", "items": ["Первый", "Второй"]},
            {"type": "image", "path": image(tmp_path), "width_mm": 30},
        ],
    }


def excel_spec(tmp_path):
    return {
        "sheets": [
            {
                "name": "Data",
                "rows": [
                    ["Месяц", "Выручка"],
                    ["Январь", 42],
                    ["Февраль", 58],
                    ["Итого", {"formula": "SUM(B2:B3)", "number_format": "0.00"}],
                ],
                "column_widths": {"A": 22, "B": 16},
                "freeze_panes": "B2",
                "autofilter": "A1:B3",
                "tables": [{"name": "Sales", "range": "A1:B3"}],
                "validations": [{"range": "C2:C4", "formula": '"Да,Нет"'}],
                "conditional_formats": [
                    {
                        "range": "B2:B3",
                        "type": "cell_is",
                        "operator": "greaterThan",
                        "formulas": ["50"],
                    }
                ],
                "charts": [
                    {
                        "type": "bar",
                        "data_range": "B1:B3",
                        "categories_range": "A2:A3",
                        "anchor": "E2",
                    }
                ],
                "images": [{"path": image(tmp_path), "anchor": "E18"}],
            }
        ],
        "defined_names": {"Revenue": "'Data'!$B$2:$B$3"},
    }


def ppt_spec(tmp_path):
    return {
        "slides": [
            {
                "layout": 6,
                "notes": "Итоги квартала",
                "background": "FFFFFF",
                "objects": [
                    {
                        "type": "text",
                        "text": "Отчёт • 2026",
                        "x": 0.6,
                        "y": 0.4,
                        "width": 10,
                        "height": 1,
                        "font": {"size": 30, "color": "225599"},
                    },
                    {
                        "type": "image",
                        "path": image(tmp_path),
                        "x": 10,
                        "y": 0.4,
                        "width": 2,
                        "height": 1,
                    },
                    {
                        "type": "table",
                        "rows": [["Месяц", "Сумма"], ["Январь", 42]],
                        "x": 0.6,
                        "y": 2,
                        "width": 4,
                        "height": 1.5,
                    },
                    {
                        "type": "chart",
                        "x": 5,
                        "y": 2,
                        "width": 7,
                        "height": 4,
                        "chart": {
                            "type": "bar",
                            "categories": ["Январь", "Февраль"],
                            "series": [{"name": "Выручка", "values": [42, 58]}],
                        },
                    },
                ],
            }
        ]
    }


def test_registry_help_and_validation(h):
    assert set(h.tool_names()) == set(TOOLS)
    for name, tool in TOOLS.items():
        from jsonschema import Draft202012Validator

        Draft202012Validator.check_schema(tool["schema"])
        result = call(h, "office_help", tool=name, language="ru")
        assert result["schema"]["type"] == "object"
        assert len(json.dumps(result).encode()) < 28 * 1024
    assert not h.call_tool(
        "word_create", output_path="x.docx", spec={"bogus": 1}
    ).success
    assert not h.call_tool("nonexistent").success


def test_help_teaches_the_workflow_in_both_languages(h):
    for language in ("en", "ru"):
        result = call(h, "office_help", language=language)
        flows = result["workflows"]
        assert set(flows) == {"presentation", "document", "spreadsheet", "slide_design"}
        # The presentation recipe ends in LOOKING at the slides, not in a schema.
        joined = " ".join(flows["presentation"])
        assert "office_render" in joined and "read_image" in joined
        assert len(json.dumps(result, ensure_ascii=False).encode()) < 28 * 1024
    assert call(h, "office_help", language="ru")["workflows"]["presentation"][
        0
    ].startswith("Сначала")


def test_a_person_facing_file_says_how_to_check_it(h, tmp_path):
    h.set_config({"libreoffice_path": str(tmp_path / "missing.exe")})
    deck = call(
        h, "powerpoint_create", output_path=tmp_path / "a.pptx", spec=ppt_spec(tmp_path)
    )
    assert "LibreOffice" in deck["next_step"]
    assert "geometry" in deck["next_step"]
    doc = call(h, "word_create", output_path=tmp_path / "a.docx", spec={})
    assert "next_step" in doc
    book = call(
        h, "excel_create", output_path=tmp_path / "a.xlsx", spec=excel_spec(tmp_path)
    )
    assert "next_step" not in book, "a workbook is read, not looked at"


def test_config_fuzz(h):
    for config in fuzz_configs():
        h.set_config(config)


def test_word_roundtrip_runs_table_headers(h, tmp_path):
    source = tmp_path / "отчёт.docx"
    created = call(h, "word_create", output_path=source, spec=word_spec(tmp_path))
    info = call(h, "office_inspect", input_path=source)
    assert created["revision"] == info["revision"]
    output = tmp_path / "правка.docx"
    call(
        h,
        "word_edit",
        input_path=source,
        output_path=output,
        revision=info["revision"],
        operations=[
            {"op": "replace_text", "find": "Привет", "replace": "Здравствуйте"},
            {"op": "set_cell", "id": "body/2", "row": 1, "column": 1, "text": "100"},
        ],
    )
    doc = Document(output)
    assert doc.paragraphs[1].text == "Здравствуйте мир"
    assert doc.paragraphs[1].runs[0].bold
    assert doc.paragraphs[1].runs[1].italic
    assert doc.tables[0].cell(1, 1).text == "100"
    assert doc.sections[0].header.paragraphs[0].text.startswith("Отчёт")
    assert len(doc.inline_shapes) == 1
    assert source.read_bytes() != output.read_bytes()
    package(output, {})


def test_word_template_and_order(h, tmp_path):
    template = tmp_path / "template.docx"
    d = Document()
    d.add_paragraph("Имя: {{person}}")
    d.save(template)
    out = tmp_path / "filled.docx"
    call(
        h,
        "word_create",
        output_path=out,
        spec={"template_path": str(template), "values": {"person": "Ирина"}},
    )
    assert Document(out).paragraphs[0].text == "Имя: Ирина"
    dest = tmp_path / "inserted.docx"
    rev = call(h, "office_inspect", input_path=out)["revision"]
    call(
        h,
        "word_edit",
        input_path=out,
        output_path=dest,
        revision=rev,
        operations=[
            {
                "op": "insert_block",
                "before_id": "body/0",
                "block": {"type": "paragraph", "text": "Начало"},
            },
            {"op": "set_paragraph", "id": "body/0", "text": "Конец"},
        ],
    )
    assert [p.text for p in Document(dest).paragraphs] == ["Начало", "Конец"]


def test_excel_roundtrip(h, tmp_path):
    source = tmp_path / "report.xlsx"
    call(h, "excel_create", output_path=source, spec=excel_spec(tmp_path))
    info = call(h, "office_inspect", input_path=source, sheet="Data", range="A1:C4")
    formula = next(i for i in info["items"] if i["id"] == "B4")
    assert formula["formula"] == "=SUM(B2:B3)" and formula["cached_value"] is None
    output = tmp_path / "edited.xlsx"
    call(
        h,
        "excel_edit",
        input_path=source,
        output_path=output,
        revision=info["revision"],
        operations=[
            {"op": "set_cells", "sheet": "Data", "range": "B2:B2", "rows": [[100]]}
        ],
    )
    wb = load_workbook(output)
    assert (
        wb["Data"]["B2"].value == 100
        and len(wb["Data"]._charts) == 1
        and len(wb["Data"]._images) == 1
    )
    assert "Sales" in wb["Data"].tables and "Revenue" in wb.defined_names
    assert wb["Data"].freeze_panes == "B2" and wb.calculation.fullCalcOnLoad
    wb.close()


@pytest.mark.parametrize("kind", ["bar", "line", "pie", "scatter"])
def test_excel_charts(kind, tmp_path):
    path = tmp_path / (kind + ".xlsx")
    excel.create(
        {
            "sheets": [
                {
                    "name": "S",
                    "rows": [["X", "Y"], [1, 2], [2, 3]],
                    "charts": [{"type": kind, "data_range": "A1:B3"}],
                }
            ]
        },
        path,
    )
    package(path, {})
    assert len(load_workbook(path)["S"]._charts) == 1


def test_csv_roundtrip_no_formula_injection(h, tmp_path):
    path = tmp_path / "x.xlsx"
    csv = tmp_path / "in.csv"
    csv.write_text("Текст;Число\n=1+1;42\n", encoding="utf-8")
    call(h, "excel_create", output_path=path, spec={"sheets": [{"name": "S"}]})
    rev = call(h, "office_inspect", input_path=path)["revision"]
    edited = tmp_path / "edited.xlsx"
    call(
        h,
        "excel_edit",
        input_path=path,
        output_path=edited,
        revision=rev,
        operations=[
            {
                "op": "import_csv",
                "path": str(csv),
                "sheet": "S",
                "encoding": "utf-8",
                "delimiter": ";",
            }
        ],
    )
    wb = load_workbook(edited)
    assert wb["S"]["A2"].value == "=1+1" and wb["S"]["A2"].data_type == "s"
    out = tmp_path / "out.csv"
    call(
        h,
        "office_convert",
        input_path=edited,
        output_path=out,
        sheet="S",
        encoding="utf-8",
        delimiter=";",
    )
    assert "Текст;Число" in out.read_text(encoding="utf-8")


def test_powerpoint_duplicate_independent_chart(h, tmp_path):
    path = tmp_path / "deck.pptx"
    call(h, "powerpoint_create", output_path=path, spec=ppt_spec(tmp_path))
    info = call(h, "office_inspect", input_path=path)
    output = tmp_path / "edited.pptx"
    call(
        h,
        "powerpoint_edit",
        input_path=path,
        output_path=output,
        revision=info["revision"],
        operations=[
            {"op": "duplicate_slide", "slide": 0},
            {
                "op": "update_chart",
                "slide": 1,
                "shape_id": 5,
                "chart": {
                    "type": "bar",
                    "categories": ["A", "B"],
                    "series": [{"name": "Новые", "values": [100, 200]}],
                },
            },
            {"op": "set_notes", "slide": 1, "text": "Новая заметка"},
        ],
    )
    prs = Presentation(output)
    assert len(prs.slides) == 2
    charts = [next(s.chart for s in sl.shapes if s.has_chart) for sl in prs.slides]
    assert list(charts[0].series[0].values) == [42, 58]
    assert list(charts[1].series[0].values) == [100, 200]
    assert prs.slides[1].notes_slide.notes_text_frame.text == "Новая заметка"
    package(output, {})


def test_powerpoint_delete_and_move(h, tmp_path):
    path = tmp_path / "deck.pptx"
    call(
        h,
        "powerpoint_create",
        output_path=path,
        spec={"slides": [{"title": "A"}, {"title": "B"}, {"title": "C"}]},
    )
    out = tmp_path / "out.pptx"
    rev = call(h, "office_inspect", input_path=path)["revision"]
    call(
        h,
        "powerpoint_edit",
        input_path=path,
        output_path=out,
        revision=rev,
        operations=[
            {"op": "move_slide", "slide": 2, "to": 0},
            {"op": "delete_slide", "slide": 1},
        ],
    )
    texts = [s.shapes[0].text for s in Presentation(out).slides]
    assert texts == ["C", "B"]


def test_powerpoint_template_placeholder(h, tmp_path):
    template = tmp_path / "tpl.pptx"
    p = Presentation()
    slide = p.slides.add_slide(p.slide_layouts[0])
    slide.shapes.title.text = "{{title}}"
    p.save(template)
    out = tmp_path / "filled.pptx"
    call(
        h,
        "powerpoint_create",
        output_path=out,
        spec={
            "template_path": str(template),
            "slides": [
                {
                    "layout": 0,
                    "objects": [
                        {"type": "placeholder", "index": 0, "text": "Дополнение"}
                    ],
                }
            ],
            "values": {"title": "Итоги"},
        },
    )
    assert Presentation(out).slides[0].shapes.title.text == "Итоги"


def test_overwrite_backup_and_stale_revision(h, tmp_path):
    path = tmp_path / "a.docx"
    created = call(
        h,
        "word_create",
        output_path=path,
        spec={"blocks": [{"type": "paragraph", "text": "A"}]},
    )
    original = path.read_bytes()
    denied = h.call_tool(
        "word_edit",
        input_path=str(path),
        output_path=str(path),
        revision=created["revision"],
        operations=[],
    )
    assert not denied.success
    result = call(
        h,
        "word_edit",
        input_path=path,
        output_path=path,
        revision=created["revision"],
        overwrite=True,
        operations=[{"op": "replace_text", "find": "A", "replace": "B"}],
    )
    assert Path(result["backup_path"]).read_bytes() == original
    out = tmp_path / "never.docx"
    denied = h.call_tool(
        "word_edit",
        input_path=str(path),
        output_path=str(out),
        revision=created["revision"],
        operations=[],
    )
    assert not denied.success and not out.exists()


def test_atomic_batch_failure(h, tmp_path):
    path = tmp_path / "a.docx"
    created = call(
        h,
        "word_create",
        output_path=path,
        spec={"blocks": [{"type": "paragraph", "text": "A"}]},
    )
    before = path.read_bytes()
    out = tmp_path / "never.docx"
    failure = h.call_tool(
        "word_edit",
        input_path=str(path),
        output_path=str(out),
        revision=created["revision"],
        operations=[
            {"op": "replace_text", "find": "A", "replace": "B"},
            {"op": "delete", "id": "missing"},
        ],
    )
    assert not failure.success and path.read_bytes() == before and not out.exists()


def test_roots_and_wrong_extensions(h, tmp_path):
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    h.set_config({"allowed_roots": str(allowed)})
    assert not h.call_tool(
        "word_create", output_path=str(tmp_path / "out.docx"), spec={}
    ).success
    assert not h.call_tool(
        "word_create", output_path=str(allowed / "out.xlsx"), spec={}
    ).success
    assert not (allowed / "out.xlsx").exists()


def test_pagination_large_text(h, tmp_path):
    path = tmp_path / "a.docx"
    call(
        h,
        "word_create",
        output_path=path,
        spec={
            "blocks": [
                {"type": "paragraph", "text": "Я" * 50000},
                {"type": "paragraph", "text": "needle"},
            ]
        },
    )
    h.set_config({"max_result_kb": 4})
    info = call(h, "office_inspect", input_path=path, limit=1)
    assert info["next_offset"] == 1 and info["items"][0]["truncated"]
    info = call(h, "office_inspect", input_path=path, offset=1)
    assert info["items"][0]["text"] == "needle"
    assert len(json.dumps(info).encode()) < 4096
    found = call(h, "office_inspect", input_path=path, query="needle")
    assert found["total_items"] == 1


def test_unavailable_libreoffice(h, tmp_path):
    path = tmp_path / "a.docx"
    call(h, "word_create", output_path=path, spec={})
    h.set_config({"libreoffice_path": str(tmp_path / "missing.exe")})
    out = tmp_path / "a.pdf"
    failure = h.call_tool(
        "office_export_pdf", input_path=str(path), output_path=str(out)
    )
    assert not failure.success and failure.code == "UNAVAILABLE" and not out.exists()


def test_bad_zip_limits_and_missing_relationship(tmp_path):
    p = tmp_path / "bad.docx"
    p.write_bytes(b"not a zip")
    with pytest.raises(BadArguments):
        package(p, {})
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("../word/document.xml", "x")
    with pytest.raises(BadArguments):
        package(p, {})
    d = Document()
    d.add_paragraph("X")
    d.save(p)
    blobs = package(p, {})
    blobs["_rels/.rels"] = blobs["_rels/.rels"].replace(
        b"word/document.xml", b"word/missing.xml"
    )
    with zipfile.ZipFile(p, "w") as z:
        for key, data in blobs.items():
            z.writestr(key, data)
    with pytest.raises(BadArguments):
        package(p, {})


def test_unsupported_excel_shape_not_silently_lost(h, tmp_path):
    path = tmp_path / "x.xlsx"
    excel.create({"sheets": [{"name": "S"}]}, path)
    blobs = package(path, {})
    blobs["xl/worksheets/sheet1.xml"] = blobs["xl/worksheets/sheet1.xml"].replace(
        b"</worksheet>", b"<extLst/></worksheet>"
    )
    with zipfile.ZipFile(path, "w") as z:
        for key, data in blobs.items():
            z.writestr(key, data)
    info = call(h, "office_inspect", input_path=path)
    assert info["limitations"]
    out = tmp_path / "never.xlsx"
    failure = h.call_tool(
        "excel_edit",
        input_path=str(path),
        output_path=str(out),
        revision=info["revision"],
        operations=[],
    )
    assert not failure.success and not out.exists()


def test_preservation_guard():
    with pytest.raises(BadArguments):
        preservation({"image.png": b"A"}, {"image.png": b"B"})
    with pytest.raises(BadArguments):
        preservation({"opaque.xml": b"A"}, {})


def test_wire_registration_authentication_config(tmp_path):
    with WireHarness(
        OfficeFiles(),
        plugin_id="office-files",
        capabilities=["tools"],
        config={"allowed_roots": str(tmp_path)},
    ) as w:
        assert len(w.list_tools()) == len(TOOLS) == 13
        assert w.health().healthy
        reply = w.call_tool(
            "word_create",
            output_path=str(tmp_path / "wire.docx"),
            spec={"blocks": [{"type": "paragraph", "text": "Через gRPC"}]},
        )
        assert reply.success, reply.error
        assert Document(tmp_path / "wire.docx").paragraphs[0].text == "Через gRPC"
        assert not w.daemon.unauthenticated_calls


# Review regressions: each test covers an observed failure, not implementation syntax.
def test_shared_strings_from_excel_origin(h, tmp_path):
    import xlsxwriter

    path = tmp_path / "external.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    ws.write("A1", "Привет")
    ws.write("B1", 1)
    wb.close()
    assert "xl/sharedStrings.xml" in package(path, {})
    out = tmp_path / "edited.xlsx"
    rev = call(h, "office_inspect", input_path=path)["revision"]
    call(
        h,
        "excel_edit",
        input_path=path,
        output_path=out,
        revision=rev,
        operations=[
            {"op": "set_cells", "sheet": "Data", "range": "B1:B1", "rows": [[2]]}
        ],
    )
    wb = load_workbook(out)
    assert wb["Data"]["A1"].value == "Привет" and wb["Data"]["B1"].value == 2
    wb.close()


def test_word_image_insertion_structure(h, tmp_path):
    path = tmp_path / "a.docx"
    c = call(
        h,
        "word_create",
        output_path=path,
        spec={"blocks": [{"type": "paragraph", "text": "Текст"}]},
    )
    out = tmp_path / "out.docx"
    call(
        h,
        "word_edit",
        input_path=path,
        output_path=out,
        revision=c["revision"],
        operations=[
            {
                "op": "insert_block",
                "before_id": "body/0",
                "block": {"type": "image", "path": image(tmp_path)},
            }
        ],
    )
    doc = Document(out)
    assert doc._element.body[0].tag.endswith("}p")
    assert len(doc.inline_shapes) == 1


def test_free_key_maps_are_not_resource_paths(h, tmp_path):
    template = tmp_path / "a.docx"
    d = Document()
    d.add_paragraph("{{path}}")
    d.save(template)
    out = tmp_path / "out.docx"
    call(
        h,
        "word_create",
        output_path=out,
        spec={
            "template_path": str(template),
            "styles": {"path": {"bold": True}},
            "values": {"path": "Displayed text"},
        },
    )
    assert Document(out).paragraphs[0].text == "Displayed text"
    call(
        h,
        "excel_create",
        output_path=tmp_path / "a.xlsx",
        spec={"sheets": [{"name": "Data"}], "defined_names": {"path": "'Data'!$A$1"}},
    )


def test_excel_copy_preserves_metadata_and_underline(h, tmp_path):
    path = tmp_path / "a.xlsx"
    c = call(
        h,
        "excel_create",
        output_path=path,
        spec={
            "sheets": [
                {
                    "name": "S",
                    "rows": [[{"value": "A", "font": {"underline": True}}], [1]],
                    "freeze_panes": "A2",
                    "autofilter": "A1:A2",
                    "validations": [{"range": "B2", "formula": '"X,Y"'}],
                    "conditional_formats": [
                        {"range": "A2", "type": "cell_is", "formulas": ["0"]}
                    ],
                }
            ]
        },
    )
    out = tmp_path / "out.xlsx"
    call(
        h,
        "excel_edit",
        input_path=path,
        output_path=out,
        revision=c["revision"],
        operations=[{"op": "copy_sheet", "sheet": "S", "name": "Copy"}],
    )
    wb = load_workbook(out)
    copied = wb["Copy"]
    assert copied["A1"].font.underline == "single"
    assert copied.freeze_panes == "A2" and copied.auto_filter.ref == "A1:A2"
    assert (
        copied.data_validations.count == 1 and len(copied.conditional_formatting) == 1
    )
    wb.close()


def test_sparse_overview_and_large_summary(h, tmp_path):
    from openpyxl import Workbook
    from openpyxl.workbook.defined_name import DefinedName

    path = tmp_path / "sparse.xlsx"
    wb = Workbook()
    wb.active["CV2000"] = "end"
    for i in range(500):
        wb.defined_names.add(
            DefinedName("Field_" + str(i) + "_" + "x" * 85, attr_text="'Sheet'!$A$1")
        )
    wb.save(path)
    info = call(h, "office_inspect", input_path=path, mode="overview")
    assert len(json.dumps(info).encode()) < 24 * 1024
    assert info["summary"]["truncated"]


def test_worker_timeout_never_publishes(tmp_path, monkeypatch):
    from src import runtime

    original = runtime.asyncio.create_subprocess_exec

    async def stalled(*argv, **kwargs):
        return await original(
            sys.executable, "-c", "import time; time.sleep(20)", **kwargs
        )

    monkeypatch.setattr(runtime.asyncio, "create_subprocess_exec", stalled)
    runner = Runner()
    dest = tmp_path / "never.docx"

    async def check():
        with pytest.raises(Timeout):
            await runner.run(
                "word_create",
                {"output_path": str(dest), "spec": {}},
                {"timeout_secs": 1},
            )
        assert not runner.processes
        await runner.close()

    asyncio.run(check())
    assert not dest.exists()


def test_link_race_never_clobbers(tmp_path, monkeypatch):
    from src import runtime

    original = runtime.os.link
    target = tmp_path / "out.docx"

    def race(source, dest):
        Path(dest).write_bytes(b"someone else's file")
        return original(source, dest)

    monkeypatch.setattr(runtime.os, "link", race)
    runner = Runner()

    async def check():
        with pytest.raises(FileExistsError):
            await runner.run(
                "word_create", {"output_path": str(target), "spec": {}}, {}
            )
        await runner.close()

    asyncio.run(check())
    assert target.read_bytes() == b"someone else's file"


@pytest.fixture
def lo_settings():
    path = os.environ.get("OFFICE_TEST_LIBREOFFICE")
    if not path or not Path(path).is_file():
        pytest.skip("Set OFFICE_TEST_LIBREOFFICE to test real LibreOffice")
    return {"libreoffice_path": path}


@pytest.mark.parametrize(
    "prefix,spec_fn,ext",
    [
        ("word", word_spec, "docx"),
        ("excel", excel_spec, "xlsx"),
        ("powerpoint", ppt_spec, "pptx"),
    ],
)
def test_real_pdf_and_png(h, tmp_path, lo_settings, prefix, spec_fn, ext):
    h.set_config(lo_settings)
    path = tmp_path / ("report." + ext)
    call(h, prefix + "_create", output_path=path, spec=spec_fn(tmp_path))
    pdf = tmp_path / "report.pdf"
    result = call(h, "office_export_pdf", input_path=path, output_path=pdf)
    assert result["pages"] >= 1 and pdf.read_bytes().startswith(b"%PDF")
    rendered = call(
        h, "office_render", input_path=path, output_dir=tmp_path / "preview", limit=1
    )
    png = Path(rendered["images"][0])
    assert png.is_file() and str(tmp_path / "preview") in str(png)
    assert Image.open(png).size[0] > 500


def test_real_recalculation_empty_string_and_numeric(h, tmp_path, lo_settings):
    h.set_config(lo_settings)
    path = tmp_path / "formulas.xlsx"
    call(
        h,
        "excel_create",
        output_path=path,
        spec={
            "sheets": [
                {
                    "name": "Data",
                    "rows": [
                        [
                            42,
                            58,
                            {"formula": "SUM(A1:B1)"},
                            {"formula": 'IF(TRUE,"","x")'},
                            {"formula": "1/0"},
                        ]
                    ],
                }
            ]
        },
    )
    out = tmp_path / "calculated.xlsx"
    result = call(h, "excel_recalculate", input_path=path, output_path=out)
    assert result["recalculated_formulas"] == 3
    wb = load_workbook(out, data_only=True)
    assert wb["Data"]["C1"].value == 100 and wb["Data"]["D1"].value in (None, "")
    assert wb["Data"]["E1"].value == "#DIV/0!" and result["formula_errors"]
    wb.close()
    formulas = load_workbook(out, data_only=False)
    assert formulas["Data"]["C1"].value == "=SUM(A1:B1)"
    formulas.close()


def test_arrays_refused_before_recalculation(h, tmp_path, lo_settings):
    import xlsxwriter

    path = tmp_path / "array.xlsx"
    wb = xlsxwriter.Workbook(path)
    s = wb.add_worksheet("S")
    s.write_array_formula("A1:A2", "=ROW(A1:A2)", None, 1)
    wb.close()
    h.set_config(lo_settings)
    result = h.call_tool(
        "excel_recalculate",
        input_path=str(path),
        output_path=str(tmp_path / "never.xlsx"),
    )
    assert not result.success and "array" in result.error


def test_mixed_word_run_refused_without_image_loss(h, tmp_path):
    path, out = tmp_path / "mixed.docx", tmp_path / "out.docx"
    doc = Document()
    run = doc.add_paragraph().add_run("hello ")
    run.add_picture(image(tmp_path))
    run.add_text("world")
    doc.save(path)
    rev = call(h, "office_inspect", input_path=path)["revision"]
    result = h.call_tool(
        "word_edit",
        input_path=str(path),
        output_path=str(out),
        revision=rev,
        operations=[{"op": "replace_text", "find": "hello", "replace": "good"}],
    )
    assert not result.success and "drawing/field" in result.error
    assert not out.exists() and len(Document(path).inline_shapes) == 1


def test_large_ppt_shape_name_bounded(h, tmp_path):
    path = tmp_path / "big.pptx"
    call(
        h,
        "powerpoint_create",
        output_path=path,
        spec={
            "slides": [
                {
                    "layout": 6,
                    "objects": [
                        {
                            "type": "text",
                            "text": "small",
                            "name": "Я" * 10000,
                            "x": 0,
                            "y": 0,
                            "width": 4,
                            "height": 1,
                        }
                    ],
                }
            ]
        },
    )
    h.set_config({"max_result_kb": 4})
    result = call(h, "office_inspect", input_path=path, offset=1, limit=1)
    assert len(json.dumps(result, ensure_ascii=False).encode()) <= 4096


def test_word_special_break_refused(h, tmp_path):
    from docx.enum.text import WD_BREAK

    path = tmp_path / "page.docx"
    doc = Document()
    run = doc.add_paragraph().add_run("before")
    run.add_break(WD_BREAK.PAGE)
    doc.save(path)
    rev = call(h, "office_inspect", input_path=path)["revision"]
    out = tmp_path / "out.docx"
    result = h.call_tool(
        "word_edit",
        input_path=str(path),
        output_path=str(out),
        revision=rev,
        operations=[{"op": "replace_text", "find": "before", "replace": "after"}],
    )
    assert not result.success and not out.exists()
    assert 'w:type="page"' in Document(path).paragraphs[0]._p.xml


@pytest.mark.skipif(os.name != "nt", reason="Windows exclusive lock")
def test_locked_destination_keeps_source(h, tmp_path):
    import ctypes
    from ctypes import wintypes

    source, dest = tmp_path / "source.docx", tmp_path / "locked.docx"
    call(
        h,
        "word_create",
        output_path=source,
        spec={"blocks": [{"type": "paragraph", "text": "before"}]},
    )
    dest.write_bytes(source.read_bytes())
    original = dest.read_bytes()
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateFileW.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    kernel.CreateFileW.restype = wintypes.HANDLE
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel.CreateFileW(str(dest), 0x80000000, 0, None, 3, 0, None)
    assert handle != ctypes.c_void_p(-1).value
    try:
        result = h.call_tool(
            "word_edit",
            input_path=str(source),
            output_path=str(dest),
            overwrite=True,
            revision=call(h, "office_inspect", input_path=source)["revision"],
            operations=[{"op": "replace_text", "find": "before", "replace": "after"}],
        )
        assert not result.success
    finally:
        kernel.CloseHandle(handle)
    assert dest.read_bytes() == original and source.read_bytes() == original


@pytest.mark.parametrize(
    "prefix,ext,legacy,filter_name",
    [
        ("word", "docx", "doc", "MS Word 97"),
        ("excel", "xlsx", "xls", "MS Excel 97"),
        ("powerpoint", "pptx", "ppt", "MS PowerPoint 97"),
    ],
)
def test_legacy_conversion(h, tmp_path, lo_settings, prefix, ext, legacy, filter_name):
    import subprocess

    h.set_config(lo_settings)
    specs = {
        "word": {"blocks": [{"type": "paragraph", "text": "Кириллица legacy"}]},
        "excel": {"sheets": [{"name": "S", "rows": [["Кириллица legacy", 42]]}]},
        "powerpoint": {"slides": [{"layout": 6, "title": "Кириллица legacy"}]},
    }
    source = tmp_path / ("sample." + ext)
    call(h, prefix + "_create", output_path=source, spec=specs[prefix])
    result = subprocess.run(
        [
            lo_settings["libreoffice_path"],
            "-env:UserInstallation=" + (tmp_path / "profile").as_uri(),
            "--headless",
            "--convert-to",
            legacy + ":" + filter_name,
            "--outdir",
            str(tmp_path),
            str(source),
        ],
        timeout=60,
        capture_output=True,
        check=False,
        creationflags=0x08000000 if os.name == "nt" else 0,
    )
    assert result.returncode == 0 and (tmp_path / ("sample." + legacy)).is_file()
    output = tmp_path / ("converted." + ext)
    call(
        h,
        "office_convert",
        input_path=tmp_path / ("sample." + legacy),
        output_path=output,
    )
    assert call(h, "office_inspect", input_path=output)["format"] == ext


def text_pdf(path, texts):
    """A PDF with a real text layer, built through PDFium alone."""
    import ctypes

    import pypdfium2 as pdfium
    from pypdfium2 import raw

    doc = pdfium.PdfDocument.new()
    for text in texts:
        page = doc.new_page(595, 842)
        obj = raw.FPDFPageObj_NewTextObj(doc.raw, b"Helvetica", ctypes.c_float(12))
        buffer = ctypes.create_string_buffer((text + "\0").encode("utf-16-le"))
        raw.FPDFText_SetText(obj, ctypes.cast(buffer, ctypes.POINTER(raw.FPDF_WCHAR)))
        raw.FPDFPageObj_Transform(obj, 1, 0, 0, 1, 50, 700)
        raw.FPDFPage_InsertObject(page.raw, obj)
        raw.FPDFPage_GenerateContent(page.raw)
        page.close()
    doc.save(str(path))
    doc.close()
    return path


def pdf_pages(path):
    import pypdfium2 as pdfium

    with pdfium.PdfDocument(str(path)) as doc:
        result = []
        for page in doc:
            textpage = page.get_textpage()
            result.append((textpage.get_text_range().strip(), page.get_rotation()))
            textpage.close()
            page.close()
        return result


def test_chat_shaped_arguments_are_repaired(h, tmp_path):
    # Verbatim shapes from the 2026-10-26 Astra test chat: a nested spec sent
    # as a JSON string, a boolean as "true", and `path` for `input_path`.
    docx = tmp_path / "test.docx"
    result = call(
        h,
        "word_create",
        output_path=docx,
        overwrite="true",
        spec=json.dumps(
            {
                "blocks": [
                    {"type": "heading", "text": "Тестовый отчёт", "level": 1},
                    {"type": "paragraph", "text": "Абзац"},
                    {"type": "bullet_list", "items": ["Пункт один", "Пункт два"]},
                    {"type": "table", "rows": [["Имя", "Значение"], ["A", 1]]},
                ]
            },
            ensure_ascii=False,
        ),
    )
    assert result["validation"] == "reopened_and_package_checked"
    pptx = tmp_path / "test.pptx"
    call(
        h,
        "powerpoint_create",
        output_path=pptx,
        overwrite="true",
        spec=json.dumps(
            {
                "slide_size": "16:9",
                "slides": [
                    {
                        "title": "Тест",
                        "objects": [
                            {
                                "type": "table",
                                "x": 0.8,
                                "y": 2.6,
                                "width": 11.7,
                                "height": 2,
                                "rows": [["X", 10]],
                            }
                        ],
                    }
                ],
            }
        ),
    )
    inspected = call(h, "office_inspect", path=docx)
    texts = [item.get("text") for item in inspected["items"]]
    assert "Тестовый отчёт" in texts and "Пункт два" in texts
    assert call(h, "office_inspect", path=pptx, limit="50")["summary"]["slides"] == 1
    # A string that merely looks like JSON stays text where text is allowed.
    xlsx = tmp_path / "t.xlsx"
    call(
        h,
        "excel_create",
        output_path=xlsx,
        spec={"sheets": [{"name": "S", "rows": [['{"a": 1}', "true"]]}]},
    )
    values = [
        item["value"] for item in call(h, "office_inspect", input_path=xlsx)["items"]
    ]
    assert values == ['{"a": 1}', "true"]
    # Unparseable strings are still refused, naming the field.
    bad = h.call_tool("word_create", output_path=str(docx), spec="{not json")
    assert not bad.success and "/spec" in bad.error


def test_pdf_inspect_edit_and_render(h, tmp_path):
    a = text_pdf(tmp_path / "a.pdf", ["alpha one", "alpha two", "alpha three"])
    b = text_pdf(tmp_path / "b.pdf", ["beta one", "beta two"])
    inspected = call(h, "office_inspect", input_path=a)
    assert inspected["format"] == "pdf"
    assert inspected["summary"]["pages"] == 3
    assert [i["text"].strip() for i in inspected["items"]] == [
        "alpha one",
        "alpha two",
        "alpha three",
    ]
    found = call(h, "office_inspect", input_path=a, query="TWO")
    assert [i["page"] for i in found["items"]] == [1]
    paged = call(h, "office_inspect", input_path=a, offset=1, limit=1)
    assert [i["page"] for i in paged["items"]] == [1] and paged["next_offset"] == 2

    merged = tmp_path / "merged.pdf"
    result = call(
        h,
        "pdf_edit",
        input_path=a,
        output_path=merged,
        revision=inspected["revision"],
        operations=[
            {"op": "delete_pages", "pages": [1]},
            {"op": "insert_pdf", "path": str(b), "pages": [1], "at": 1},
            {"op": "insert_pdf", "path": str(b)},
            {"op": "rotate_pages", "pages": [0], "angle": 90},
            {"op": "move_page", "page": 4, "to": 0},
        ],
    )
    assert result["pages"] == 5 and result["operations_applied"] == 5
    assert pdf_pages(merged) == [
        ("beta two", 0),
        ("alpha one", 90),
        ("beta two", 0),
        ("alpha three", 0),
        ("beta one", 0),
    ]
    assert pdf_pages(a)[1] == ("alpha two", 0)  # the source is untouched

    split = tmp_path / "split.pdf"
    call(
        h,
        "pdf_edit",
        input_path=merged,
        output_path=split,
        revision=result["revision"],
        operations=[{"op": "keep_pages", "pages": [3, 1]}],
    )
    assert [t for t, _ in pdf_pages(split)] == ["alpha three", "alpha one"]

    refused = h.call_tool(
        "pdf_edit",
        input_path=str(a),
        output_path=str(tmp_path / "x.pdf"),
        revision=inspected["revision"],
        operations=[{"op": "delete_pages", "pages": [7]}],
    )
    assert not refused.success and "outside 0..2" in refused.error
    assert not (tmp_path / "x.pdf").exists()
    everything = h.call_tool(
        "pdf_edit",
        input_path=str(a),
        output_path=str(tmp_path / "x.pdf"),
        revision=inspected["revision"],
        operations=[{"op": "delete_pages", "pages": [0, 1, 2]}],
    )
    assert not everything.success

    # Rendering a PDF needs no LibreOffice, even when it is misconfigured.
    h.set_config({"libreoffice_path": str(tmp_path / "missing-soffice.exe")})
    rendered = call(h, "office_render", input_path=a, output_dir=tmp_path / "png")
    assert rendered["pages"] == 3 and len(rendered["images"]) == 3
    assert Image.open(rendered["images"][0]).size[0] > 0


def test_pdf_scan_and_garbage(h, tmp_path):
    scan = tmp_path / "scan.pdf"
    Image.new("RGB", (300, 400), "white").save(scan)
    inspected = call(h, "office_inspect", input_path=scan)
    assert inspected["summary"]["pages_without_text"] == [0]
    assert "office_render" in inspected["summary"]["hint"]
    fake = tmp_path / "fake.pdf"
    fake.write_bytes(b"not a pdf at all")
    result = h.call_tool("office_inspect", input_path=str(fake))
    assert not result.success and "PDF" in result.error
