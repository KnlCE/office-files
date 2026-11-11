"""One registry owns dispatch, validation, model schemas and bilingual help."""

import json
import math

from astra_plugin_sdk import BadArguments, NotFound
from jsonschema import Draft202012Validator
from jsonschema.exceptions import best_match


def obj(properties, required=()):
    return {
        "type": "object",
        "properties": properties,
        "required": list(required),
        "additionalProperties": False,
    }


def arr(items, maximum=1000):
    return {"type": "array", "items": items, "maxItems": maximum}


S = {"type": "string", "maxLength": 100000}
PATH = {"type": "string", "minLength": 1, "maxLength": 4096}
N = {"type": "number"}
B = {"type": "boolean"}
I = {"type": "integer", "minimum": 0}
COLOR = {"type": "string", "pattern": "^[0-9A-Fa-f]{6}$"}
VALUE = {"type": ["string", "number", "boolean", "null"], "maxLength": 100000}
FONT = obj(
    {
        "name": S,
        "size": {"type": "number", "minimum": 1, "maximum": 400},
        "bold": B,
        "italic": B,
        "underline": B,
        "color": COLOR,
    }
)
RUN = obj({"text": S, "font": FONT}, ["text"])
TEXT = {
    "text": S,
    "runs": arr(RUN),
    "style": S,
    "alignment": {"enum": ["left", "center", "right", "justify"]},
    "font": FONT,
}
ROWS = arr(arr(VALUE, 1000), 10000)
BLOCKS = []


def variant(op, properties=None, required=(), tag="op"):
    return obj({tag: {"const": op}, **(properties or {})}, [tag, *required])


WORD_BLOCK = {
    "oneOf": [
        variant("paragraph", TEXT, tag="type"),
        variant(
            "heading",
            {**TEXT, "level": {"type": "integer", "minimum": 1, "maximum": 9}},
            tag="type",
        ),
        variant("bullet_list", {"items": arr(S)}, ["items"], tag="type"),
        variant("numbered_list", {"items": arr(S)}, ["items"], tag="type"),
        variant("table", {"rows": ROWS, "style": S}, ["rows"], tag="type"),
        variant(
            "image",
            {"path": PATH, "width_mm": {"type": "number", "exclusiveMinimum": 0}},
            ["path"],
            tag="type",
        ),
        variant("page_break", tag="type"),
    ]
}
CHART = obj(
    {
        "type": {"enum": ["bar", "line", "pie", "scatter"]},
        "title": S,
        "categories": arr(VALUE),
        "series": arr(
            obj({"name": S, "values": arr(N), "x_values": arr(N)}, ["values"])
        ),
        "data_range": S,
        "categories_range": S,
        "anchor": S,
        "width": N,
        "height": N,
    },
    ["type"],
)
CELL = obj(
    {
        "value": VALUE,
        "formula": S,
        "type": {"enum": ["date", "datetime"]},
        "number_format": S,
        "font": FONT,
        "fill": COLOR,
        "horizontal": {"enum": ["left", "center", "right"]},
        "wrap": B,
    }
)
CELL_VALUE = {"oneOf": [VALUE, CELL]}
VALIDATION = obj(
    {"range": S, "formula": S, "type": {"enum": ["list", "whole", "decimal", "date"]}},
    ["range", "formula"],
)
CONDITIONAL = obj(
    {
        "range": S,
        "type": {"enum": ["cell_is", "color_scale"]},
        "operator": {
            "enum": ["equal", "notEqual", "lessThan", "greaterThan", "between"]
        },
        "formulas": arr(S, 2),
        "fill": COLOR,
        "start_color": COLOR,
        "end_color": COLOR,
    },
    ["range", "type"],
)
SHEET = obj(
    {
        "name": S,
        "rows": arr(arr(CELL_VALUE, 1000), 10000),
        "cells": {
            "type": "object",
            "propertyNames": {"pattern": "^[A-Z]{1,3}[1-9][0-9]*$"},
            "additionalProperties": CELL_VALUE,
        },
        "column_widths": {
            "type": "object",
            "additionalProperties": {"type": "number", "exclusiveMinimum": 0},
        },
        "row_heights": {
            "type": "object",
            "additionalProperties": {"type": "number", "exclusiveMinimum": 0},
        },
        "merges": arr(S),
        "freeze_panes": S,
        "autofilter": S,
        "tables": arr(obj({"name": S, "range": S, "style": S}, ["name", "range"])),
        "validations": arr(VALIDATION),
        "conditional_formats": arr(CONDITIONAL),
        "charts": arr(CHART),
        "images": arr(
            obj({"path": PATH, "anchor": S, "width": N, "height": N}, ["path"])
        ),
    },
    ["name"],
)
GEOMETRY = {
    "x": {
        **N,
        "description": "Left position in INCHES, not pixels. A 16:9 slide is 13.333 inches wide; e.g. x=0.8.",
    },
    "y": {
        **N,
        "description": "Top position in INCHES, not pixels. Slide height is 7.5 inches; e.g. y=1.2.",
    },
    "width": {
        "type": "number",
        "exclusiveMinimum": 0,
        "description": "Object width in INCHES, not pixels; e.g. width=11.7 for a title.",
    },
    "height": {
        "type": "number",
        "exclusiveMinimum": 0,
        "description": "Object height in INCHES, not pixels; e.g. height=0.8 for a title.",
    },
    "allow_outside_slide": {
        **B,
        "description": "Default false: reject geometry beyond the slide. Use true only for deliberate cropping/off-slide assets; results still warn.",
    },
    "name": S,
}
PPT_OBJECT = {
    "oneOf": [
        variant("text", {**GEOMETRY, **TEXT}, tag="type"),
        variant(
            "shape",
            {
                **GEOMETRY,
                **TEXT,
                "shape": {
                    "enum": [
                        "rectangle",
                        "rounded_rectangle",
                        "ellipse",
                        "arrow",
                        "line",
                    ]
                },
                "fill": COLOR,
            },
            tag="type",
        ),
        variant("image", {**GEOMETRY, "path": PATH}, ["path"], tag="type"),
        variant("table", {**GEOMETRY, "rows": ROWS}, ["rows"], tag="type"),
        variant("chart", {**GEOMETRY, "chart": CHART}, ["chart"], tag="type"),
        variant(
            "placeholder",
            {
                **TEXT,
                "index": I,
                "allow_outside_slide": GEOMETRY["allow_outside_slide"],
            },
            ["index"],
            tag="type",
        ),
    ]
}
SLIDE = obj(
    {
        "layout": I,
        "title": S,
        "background": COLOR,
        "notes": S,
        "objects": arr(PPT_OBJECT, 200),
    }
)
WORD_SPEC = obj(
    {
        "template_path": PATH,
        "blocks": arr(WORD_BLOCK),
        "page": obj(
            {
                k: {"type": "number", "exclusiveMinimum": 0}
                for k in [
                    "width_mm",
                    "height_mm",
                    "margin_top_mm",
                    "margin_bottom_mm",
                    "margin_left_mm",
                    "margin_right_mm",
                ]
            }
        ),
        "header": S,
        "footer": S,
        "styles": {"type": "object", "additionalProperties": FONT},
        "values": {"type": "object", "additionalProperties": S},
    }
)
EXCEL_SPEC = obj(
    {
        "template_path": PATH,
        "sheets": arr(SHEET, 100),
        "defined_names": {"type": "object", "additionalProperties": S},
    },
    ["sheets"],
)
PPT_SPEC = obj(
    {
        "template_path": PATH,
        "slide_size": {"enum": ["16:9", "4:3"]},
        "slides": arr(SLIDE, 1000),
        "values": {"type": "object", "additionalProperties": S},
    },
    ["slides"],
)
PPT_SPEC["description"] = (
    "PowerPoint geometry uses INCHES (fonts: points). 16:9 = 13.333 x 7.5 inches, 4:3 = 10 x 7.5. Put visible content in slide.objects, use x=0.8, y=1.2, width=11.7, height=0.8. Pixel coordinates such as x=60/width=840 are invalid."
)
TARGET = {"id": S}
WORD_OPS = [
    variant(
        "replace_text", {"find": S, "replace": S, "replace_all": B}, ["find", "replace"]
    ),
    variant(
        "fill_template",
        {"values": {"type": "object", "additionalProperties": S}},
        ["values"],
    ),
    variant("insert_block", {"before_id": S, "block": WORD_BLOCK}, ["block"]),
    variant("delete", TARGET, ["id"]),
    variant("set_paragraph", {**TARGET, **TEXT}, ["id"]),
    variant(
        "set_cell", {**TARGET, "row": I, "column": I, **TEXT}, ["id", "row", "column"]
    ),
    variant("set_header_footer", {"header": S, "footer": S}),
    variant("set_style", {"name": S, "font": FONT}, ["name", "font"]),
]
EXCEL_OPS = [
    variant(
        "set_cells",
        {"sheet": S, "range": S, "rows": arr(arr(CELL_VALUE, 1000), 10000)},
        ["sheet", "range", "rows"],
    ),
    variant("configure_sheet", {"sheet": S, "settings": SHEET}, ["sheet", "settings"]),
    variant("add_sheet", {"spec": SHEET}, ["spec"]),
    variant("copy_sheet", {"sheet": S, "name": S}, ["sheet", "name"]),
    variant("rename_sheet", {"sheet": S, "name": S}, ["sheet", "name"]),
    variant("delete_sheet", {"sheet": S}, ["sheet"]),
    variant("define_name", {"name": S, "reference": S}, ["name", "reference"]),
    variant(
        "import_csv",
        {
            "path": PATH,
            "sheet": S,
            "encoding": S,
            "delimiter": {"type": "string", "minLength": 1, "maxLength": 1},
        },
        ["path", "sheet", "encoding", "delimiter"],
    ),
]
PPT_OPS = [
    variant(
        "replace_text", {"find": S, "replace": S, "replace_all": B}, ["find", "replace"]
    ),
    variant(
        "fill_template",
        {"values": {"type": "object", "additionalProperties": S}},
        ["values"],
    ),
    variant("add_slide", {"spec": SLIDE}, ["spec"]),
    variant("delete_slide", {"slide": I}, ["slide"]),
    variant("move_slide", {"slide": I, "to": I}, ["slide", "to"]),
    variant("duplicate_slide", {"slide": I}, ["slide"]),
    variant("add_object", {"slide": I, "spec": PPT_OBJECT}, ["slide", "spec"]),
    variant("delete_object", {"slide": I, "shape_id": I}, ["slide", "shape_id"]),
    variant("set_text", {"slide": I, "shape_id": I, **TEXT}, ["slide", "shape_id"]),
    variant(
        "set_cell",
        {"slide": I, "shape_id": I, "row": I, "column": I, **TEXT},
        ["slide", "shape_id", "row", "column"],
    ),
    variant(
        "set_geometry", {"slide": I, "shape_id": I, **GEOMETRY}, ["slide", "shape_id"]
    ),
    variant("set_background", {"slide": I, "color": COLOR}, ["slide", "color"]),
    variant("set_notes", {"slide": I, "text": S}, ["slide", "text"]),
    variant(
        "update_chart",
        {"slide": I, "shape_id": I, "chart": CHART},
        ["slide", "shape_id", "chart"],
    ),
]
PAGES = {**arr(I, 10000), "description": "Zero-based page indices."}
PDF_OPS = [
    variant("delete_pages", {"pages": PAGES}, ["pages"]),
    variant(
        "keep_pages",
        {
            "pages": {
                **PAGES,
                "description": "Zero-based pages to keep, in this order (extract/split/reorder).",
            }
        },
        ["pages"],
    ),
    variant(
        "rotate_pages",
        {
            "pages": {**PAGES, "description": "Zero-based pages; omit for all."},
            "angle": {
                "enum": [90, 180, 270],
                "description": "Clockwise degrees.",
            },
        },
        ["angle"],
    ),
    variant("move_page", {"page": I, "to": I}, ["page", "to"]),
    variant(
        "insert_pdf",
        {
            "path": PATH,
            "pages": {
                **PAGES,
                "description": "Pages of the inserted PDF; omit for all.",
            },
            "at": {
                **I,
                "description": "Insert before this page of the result; omit to append (merge).",
            },
        },
        ["path"],
    ),
]
TOOLS = {}


def register(name, description, properties, required, help_ru, example):
    TOOLS[name] = {
        "description": description,
        "schema": obj(properties, required),
        "help_ru": help_ru,
        "example": example,
    }


register(
    "office_help",
    "Office Word DOCX Excel XLSX PowerPoint PPTX and PDF tools: get schemas and examples before creation/editing. Select tool and optionally operation for concise help.",
    {"tool": S, "operation": S, "language": {"enum": ["en", "ru"]}},
    [],
    "Выберите tool и при необходимости operation; все индексы с нуля, геометрия слайдов в дюймах.",
    {"tool": "word_create", "language": "ru"},
)
register(
    "office_inspect",
    "Read/search DOCX XLSX PPTX and PDF files: text, tables, sheets, formulas and cached results, PDF text per page with outline and metadata. Returns revision for editing, stable object IDs and paginated items (PDF: one item per page). Excel optionally select sheet and A1 range.",
    {
        "input_path": PATH,
        "sheet": S,
        "range": S,
        "query": S,
        "offset": I,
        "limit": {"type": "integer", "minimum": 1, "maximum": 200},
        "mode": {"enum": ["overview", "content"]},
    },
    ["input_path"],
    "Сначала прочитайте файл. Ревизию передайте в edit. Для Excel выберите sheet/range.",
    {"input_path": "report.xlsx", "sheet": "Data", "range": "A1:C20"},
)
for prefix, spec, operations, example in [
    (
        "word",
        WORD_SPEC,
        WORD_OPS,
        {
            "blocks": [
                {"type": "heading", "text": "Отчёт", "level": 1},
                {"type": "paragraph", "text": "Итоги"},
            ]
        },
    ),
    (
        "excel",
        EXCEL_SPEC,
        EXCEL_OPS,
        {
            "sheets": [
                {
                    "name": "Data",
                    "rows": [
                        ["Товар", "Сумма"],
                        ["A", 42],
                        ["Итого", {"formula": "SUM(B2:B2)"}],
                    ],
                }
            ]
        },
    ),
    (
        "powerpoint",
        PPT_SPEC,
        PPT_OPS,
        {
            "slides": [
                {
                    "layout": 6,
                    "objects": [
                        {
                            "type": "text",
                            "text": "Итоги",
                            "x": 1,
                            "y": 1,
                            "width": 8,
                            "height": 1,
                        }
                    ],
                }
            ]
        },
    ),
]:
    ext = {"word": "docx", "excel": "xlsx", "powerpoint": "pptx"}[prefix]
    units = (
        " Geometry x/y/width/height is in INCHES, not pixels: 16:9 is 13.333 x 7.5; fonts in points. Out-of-slide objects are rejected unless explicitly allowed. Inspect the output; package validation is not visual verification."
        if prefix == "powerpoint"
        else ""
    )
    register(
        prefix + "_create",
        f"Create {prefix} {ext.upper()} Office file from typed JSON spec, optionally a template. Styles, tables, images and charts where supported. Call office_help for spec details, and office_help with no tool for how to make a good one (workflows). The result's next_step says how to check it."
        + units,
        {"output_path": PATH, "spec": spec, "overwrite": B},
        ["output_path", "spec"],
        "Создание из JSON. Сначала получите схему через office_help. output_path обязателен."
        + (
            " Координаты и размеры в дюймах, не пикселях; слайд 16:9 имеет размер 13,333×7,5. Проверьте результат инспекцией и, при наличии LibreOffice, рендерингом."
            if prefix == "powerpoint"
            else ""
        ),
        {"output_path": "report." + ext, "spec": example},
    )
    register(
        prefix + "_edit",
        f"Edit {prefix} {ext.upper()} Office file with an atomic batch of typed operations. Get revision via office_inspect. Save to a separate output_path; explicit overwrite required for existing files. IDs refer to inspected source; slide indices are zero-based."
        + units,
        {
            "input_path": PATH,
            "output_path": PATH,
            "revision": {"type": "string", "pattern": "^[a-f0-9]{64}$"},
            "operations": arr({"oneOf": operations}),
            "overwrite": B,
        },
        ["input_path", "output_path", "revision", "operations"],
        "Пакетная правка: исходник проверяется по SHA-256. Индексы относятся к текущему состоянию после предыдущих операций, Word IDs — к инспекции исходника.",
        {
            "input_path": "report." + ext,
            "output_path": "edited." + ext,
            "revision": "<office_inspect.revision>",
            "operations": [],
        },
    )
register(
    "pdf_edit",
    "Edit PDF pages with an atomic batch: delete, keep (extract/split/reorder), rotate, move, insert pages of another PDF (merge). Get revision via office_inspect. Pages are zero-based and refer to the result of the previous operation. Save to a separate output_path; explicit overwrite required for existing files. Text inside a PDF is not editable: create a DOCX and office_export_pdf instead.",
    {
        "input_path": PATH,
        "output_path": PATH,
        "revision": {"type": "string", "pattern": "^[a-f0-9]{64}$"},
        "operations": arr({"oneOf": PDF_OPS}),
        "overwrite": B,
    },
    ["input_path", "output_path", "revision", "operations"],
    "Пакетная правка страниц PDF: удалить, оставить (извлечь/разделить/переупорядочить), повернуть, переместить, вставить страницы другого PDF (объединить). Индексы с нуля и относятся к результату предыдущей операции. Текст внутри PDF не редактируется: создайте DOCX и экспортируйте через office_export_pdf.",
    {
        "input_path": "a.pdf",
        "output_path": "merged.pdf",
        "revision": "<office_inspect.revision>",
        "operations": [
            {"op": "insert_pdf", "path": "b.pdf"},
            {"op": "rotate_pages", "pages": [0], "angle": 90},
        ],
    },
)
for name, desc in [
    ("office_export_pdf", "Export Office file as PDF using installed LibreOffice."),
    (
        "excel_recalculate",
        "Recalculate XLSX formulas into a separate workbook using installed LibreOffice; reports formula errors.",
    ),
    (
        "office_convert",
        "Convert legacy DOC XLS PPT to DOCX XLSX PPTX or export one XLSX sheet to CSV. Requires LibreOffice for legacy conversion; CSV does not.",
    ),
]:
    register(
        name,
        desc + " Explicit output_path; no overwrite by default.",
        {
            "input_path": PATH,
            "output_path": PATH,
            "overwrite": B,
            "sheet": S,
            "encoding": S,
            "delimiter": {"type": "string", "minLength": 1, "maxLength": 1},
        },
        ["input_path", "output_path"],
        "Экспорт/пересчёт выполняется в копии. CSV требует sheet, encoding и delimiter.",
        {"input_path": "report.docx", "output_path": "report.pdf"},
    )
register(
    "office_render",
    "Render Office document/slides or PDF pages as PNG images (Office files need LibreOffice; PDF does not). Read returned PNG paths with read_image, e.g. to see a scanned PDF. Output directory must not exist.",
    {
        "input_path": PATH,
        "output_dir": PATH,
        "offset": I,
        "limit": {"type": "integer", "minimum": 1, "maximum": 20},
        "scale": {"type": "number", "minimum": 0.5, "maximum": 2},
    },
    ["input_path", "output_dir"],
    "Создаёт новую папку с PDF и PNG. Изображения читайте через read_image.",
    {"input_path": "deck.pptx", "output_dir": "preview", "limit": 5},
)


# Models address the one file a tool reads by whatever name they last saw.
PATH_ALIASES = ("path", "file_path", "file", "input")


def _types(schema):
    kind = schema.get("type")
    if kind is None and "properties" in schema:
        kind = "object"
    return set(kind) if isinstance(kind, list) else {kind}


def coerce(value, schema):
    """Undo the shapes models put on the wire for a typed argument.

    Seen in Astra's own chat logs: a nested `spec` object sent as a JSON
    string, and `overwrite: "true"`. Each repair only fires when the value
    has the wrong JSON type for the schema AND parses into the right one,
    so a well-formed call is untouched and a real string field stays a string.
    """
    if "oneOf" in schema:
        for variant in schema["oneOf"]:
            candidate = coerce(value, variant)
            if Draft202012Validator(variant).is_valid(candidate):
                return candidate
        return value
    types = _types(schema)
    if isinstance(value, str) and "string" not in types:
        text = value.strip()
        if types & {"object", "array"} and text[:1] in "{[":
            try:
                parsed = json.loads(text)
            except ValueError:
                parsed = None
            wanted = {dict: "object", list: "array"}.get(type(parsed))
            if wanted in types:
                value = parsed
        elif "boolean" in types and text.lower() in ("true", "false"):
            return text.lower() == "true"
        elif types & {"integer", "number"}:
            try:
                number = float(text)
            except ValueError:
                return value
            if "integer" in types and number.is_integer():
                return int(number)
            if "number" in types and math.isfinite(number):
                return number
            return value
    if isinstance(value, dict) and "object" in types:
        properties = schema.get("properties", {})
        extra = schema.get("additionalProperties")
        return {
            k: coerce(v, properties[k])
            if k in properties
            else coerce(v, extra)
            if isinstance(extra, dict)
            else v
            for k, v in value.items()
        }
    if isinstance(value, list) and "array" in types:
        items = schema.get("items", {})
        return [coerce(v, items) for v in value]
    return value


def normalize(name, args):
    """Validated, coerced arguments for `name`; raises BadArguments otherwise."""
    if name not in TOOLS:
        raise NotFound("Unknown tool: " + name)
    if not isinstance(args, dict):
        raise BadArguments("Expected a JSON object")
    schema = TOOLS[name]["schema"]
    properties = schema["properties"]
    args = dict(args)
    if "input_path" in properties and "input_path" not in args:
        for alias in PATH_ALIASES:
            if alias not in properties and isinstance(args.get(alias), str):
                args["input_path"] = args.pop(alias)
                break
    args = coerce(args, schema)
    errors = list(Draft202012Validator(schema).iter_errors(args))
    if errors:
        # best_match descends into oneOf, so a bad block names its own field.
        error = best_match(errors)
        raise BadArguments(
            "/" + "/".join(map(str, error.absolute_path)) + ": " + error.message[:1000]
        )
    return args


# How a person-facing file is MADE, not just which calls exist. A model that only
# knows the schema writes a deck that validates and reads badly; the order below is
# what turns it into one a person would present: plan, build in one call, look at it
# as an image, fix, look again. Kept here, next to the tools it names, so the two
# change together.
WORKFLOWS = {
    "en": {
        "presentation": [
            "Plan first: one idea per slide, a title that states it, 3-6 short bullets or one visual. Match the request's slide count, audience and language.",
            "Build the whole deck in ONE powerpoint_create call from a full spec (with template_path if the user gave a template). Geometry is in inches: 16:9 is 13.333 x 7.5.",
            "Look at it: office_render to PNG, then read_image each slide. Check for text overflowing or cut off, overlapping objects, empty slides, weak contrast, body text under ~18 pt, titles that differ in place or style.",
            "Fix with powerpoint_edit (revision from the last result), render again, look again, until nothing is wrong.",
            "Report the path, the slide count and what you checked by eye - or, if rendering was unavailable, that only the geometry was checked.",
        ],
        "document": [
            "Plan the outline first: headings that tell the story, then the body under each.",
            "Build it in ONE word_create call: real heading styles (not bold paragraphs), lists as lists, tables for tabular data.",
            "Look at it: office_render the pages and read_image them. Check for broken layout, tables running off the page, empty pages, inconsistent styles.",
            "Fix with word_edit (revision from the last result) and look again.",
            "Report the path and what you checked - or that only the structure was checked when rendering was unavailable.",
        ],
        "spreadsheet": [
            "Build with excel_create: values and formulas, then formatting.",
            "openpyxl does not calculate formulas: run excel_recalculate (LibreOffice) and office_inspect the result before you quote any computed number; otherwise say the values are not computed yet.",
            "Charts are seen, not read: office_render and read_image them.",
        ],
        "slide_design": [
            "One message per slide; the title states the takeaway, not the topic.",
            "Body text 18 pt or larger, titles 32-44 pt; at most ~6 bullets of ~12 words.",
            "Margins of at least 0.5 in, aligned edges, one font family, 2-3 colours with strong contrast.",
            "A chart, table or picture beats a wall of text; put detail in speaker notes.",
        ],
    },
    "ru": {
        "presentation": [
            "Сначала план: одна мысль на слайд, заголовок её формулирует, 3–6 коротких пунктов или один визуальный элемент. Число слайдов, аудитория и язык — как в просьбе.",
            "Соберите всю презентацию ОДНИМ вызовом powerpoint_create из полной спецификации (с template_path, если дан шаблон). Геометрия в дюймах: 16:9 — это 13,333 × 7,5.",
            "Посмотрите на неё: office_render в PNG, затем read_image каждого слайда. Ищите вылезающий или обрезанный текст, наложения, пустые слайды, слабый контраст, текст мельче ~18 пт, заголовки в разных местах и стилях.",
            "Исправьте через powerpoint_edit (revision из последнего результата), отрендерите снова и посмотрите снова — пока не останется ошибок.",
            "В отчёте: путь, число слайдов и что проверено глазами — или, если рендер недоступен, что проверена только геометрия.",
        ],
        "document": [
            "Сначала структура: заголовки, по которым понятна суть, затем текст под каждым.",
            "Соберите ОДНИМ вызовом word_create: настоящие стили заголовков (а не жирные абзацы), списки — списками, табличные данные — таблицами.",
            "Посмотрите: office_render страниц и read_image. Ищите поехавшую вёрстку, таблицы за краем страницы, пустые страницы, разнобой стилей.",
            "Исправьте через word_edit (revision из последнего результата) и посмотрите снова.",
            "В отчёте: путь и что проверено — или что проверена только структура, если рендер недоступен.",
        ],
        "spreadsheet": [
            "Соберите через excel_create: значения и формулы, затем оформление.",
            "openpyxl не вычисляет формулы: выполните excel_recalculate (LibreOffice) и office_inspect результата, прежде чем называть вычисленные числа; иначе скажите, что значения ещё не посчитаны.",
            "Диаграммы проверяют глазами: office_render и read_image.",
        ],
        "slide_design": [
            "Одна мысль на слайд; заголовок — вывод, а не тема.",
            "Основной текст от 18 пт, заголовки 32–44 пт; не больше ~6 пунктов по ~12 слов.",
            "Поля не меньше 0,5 дюйма, выровненные края, одно семейство шрифтов, 2–3 цвета с сильным контрастом.",
            "Диаграмма, таблица или картинка лучше стены текста; подробности — в заметки докладчика.",
        ],
    },
}


def help_result(args):
    name = args.get("tool")
    if not name:
        language = "ru" if args.get("language") == "ru" else "en"
        return {
            "tools": {k: t["description"] for k, t in TOOLS.items()},
            "workflow": "office_help → office_inspect → create/edit → render → read_image → fix → render again",
            "workflows": WORKFLOWS[language],
            "units": "slides: inches; Word page/image: mm; fonts: pt; indices: zero-based",
        }
    if name not in TOOLS:
        raise NotFound(name)
    t = TOOLS[name]
    schema = t["schema"]
    if args.get("operation"):
        variants = (
            schema.get("properties", {})
            .get("operations", {})
            .get("items", {})
            .get("oneOf", [])
        )
        schema = next(
            (
                v
                for v in variants
                if v["properties"]["op"]["const"] == args["operation"]
            ),
            None,
        )
        if schema is None:
            raise BadArguments("Unknown operation for " + name)
    # Full edit unions exceed tool result budgets; discover operations individually.
    elif name.endswith("_edit"):
        schema = {
            **schema,
            "properties": {
                **schema["properties"],
                "operations": {
                    "type": "array",
                    "description": "Request operation schema using office_help(tool, operation).",
                    "items": obj(
                        {
                            "op": {
                                "enum": [
                                    v["properties"]["op"]["const"]
                                    for v in schema["properties"]["operations"][
                                        "items"
                                    ]["oneOf"]
                                ]
                            }
                        }
                    ),
                },
            },
        }
    return {
        "tool": name,
        "help": t["help_ru"] if args.get("language") == "ru" else t["description"],
        "schema": schema,
        "example": t["example"],
    }
