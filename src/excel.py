"""Excel data, styles, native charts, validation and CSV interop."""

import csv
import datetime

from astra_plugin_sdk import BadArguments
from openpyxl import Workbook, load_workbook
from openpyxl.chart import (
    BarChart,
    LineChart,
    PieChart,
    Reference,
    ScatterChart,
    Series,
)
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.cell import range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


def load(path):
    return load_workbook(path, rich_text=True, keep_links=True)


def cell(target, data):
    if not isinstance(data, dict):
        # A scalar '=...' is literal text; formulas must be explicit.
        target.value = data
        if isinstance(data, str) and data.startswith("="):
            target.data_type = "s"
        return
    if "formula" in data:
        target.value = "=" + data["formula"].lstrip("=")
    elif "value" in data:
        value = data["value"]
        if data.get("type") in ["date", "datetime"]:
            value = (
                datetime.date.fromisoformat(value)
                if data["type"] == "date"
                else datetime.datetime.fromisoformat(value)
            )
        target.value = value
        if isinstance(value, str) and value.startswith("="):
            target.data_type = "s"
    if "number_format" in data:
        target.number_format = data["number_format"]
    if "font" in data:
        values = dict(data["font"])
        if "size" in values:
            values["sz"] = values.pop("size")
        if "underline" in values:
            values["underline"] = "single" if values["underline"] else None
        target.font = Font(**values)
    if "fill" in data:
        target.fill = PatternFill("solid", fgColor=data["fill"])
    if "horizontal" in data or "wrap" in data:
        target.alignment = Alignment(
            horizontal=data.get("horizontal"), wrap_text=data.get("wrap")
        )


def bounds(text):
    try:
        b = range_boundaries(text)
        if (
            any(v is None for v in b)
            or b[0] < 1
            or b[1] < 1
            or b[2] > 16384
            or b[3] > 1048576
            or b[2] < b[0]
            or b[3] < b[1]
        ):
            raise ValueError(text)
        return b
    except ValueError as exc:
        raise BadArguments("Expected bounded A1 range: " + text) from exc


def write_rows(ws, text, rows):
    minc, minr, maxc, maxr = bounds(text)
    if len(rows) > maxr - minr + 1 or any(len(r) > maxc - minc + 1 for r in rows):
        raise BadArguments("rows do not fit the target range")
    for r, row in enumerate(rows, minr):
        for c, value in enumerate(row, minc):
            cell(ws.cell(r, c), value)


def chart(ws, spec):
    cls = {
        "bar": BarChart,
        "line": LineChart,
        "pie": PieChart,
        "scatter": ScatterChart,
    }[spec["type"]]
    chart = cls()
    if "data_range" not in spec:
        raise BadArguments(
            "Excel charts require data_range (A1 range, first row contains series names)"
        )
    minc, minr, maxc, maxr = bounds(spec["data_range"])
    data = Reference(ws, min_col=minc, min_row=minr, max_col=maxc, max_row=maxr)
    if spec["type"] == "scatter":
        if maxc <= minc or maxr <= minr:
            raise BadArguments(
                "Scatter data_range needs X column and Y column with header"
            )
        x = Reference(ws, min_col=minc, min_row=minr + 1, max_row=maxr)
        for c in range(minc + 1, maxc + 1):
            chart.series.append(
                Series(
                    Reference(ws, min_col=c, min_row=minr, max_row=maxr),
                    x,
                    title_from_data=True,
                )
            )
    else:
        chart.add_data(data, titles_from_data=True)
        if "categories_range" in spec:
            a, b, c, d = bounds(spec["categories_range"])
            chart.set_categories(
                Reference(ws, min_col=a, min_row=b, max_col=c, max_row=d)
            )
    chart.title = spec.get("title")
    chart.width, chart.height = spec.get("width", 16), spec.get("height", 9)
    ws.add_chart(chart, spec.get("anchor", "E2"))


def configure(ws, spec):
    for r, row in enumerate(spec.get("rows", []), 1):
        for c, value in enumerate(row, 1):
            cell(ws.cell(r, c), value)
    for coordinate, value in spec.get("cells", {}).items():
        bounds(coordinate)
        cell(ws[coordinate], value)
    for key, value in spec.get("column_widths", {}).items():
        ws.column_dimensions[key].width = value
    for key, value in spec.get("row_heights", {}).items():
        ws.row_dimensions[int(key)].height = value
    for merged in spec.get("merges", []):
        bounds(merged)
        ws.merge_cells(merged)
    if "freeze_panes" in spec:
        ws.freeze_panes = spec["freeze_panes"]
    if "autofilter" in spec:
        bounds(spec["autofilter"])
        ws.auto_filter.ref = spec["autofilter"]
    for t in spec.get("tables", []):
        bounds(t["range"])
        table = Table(displayName=t["name"], ref=t["range"])
        table.tableStyleInfo = TableStyleInfo(
            name=t.get("style", "TableStyleMedium2"), showRowStripes=True
        )
        ws.add_table(table)
    for v in spec.get("validations", []):
        bounds(v["range"])
        validation = DataValidation(type=v.get("type", "list"), formula1=v["formula"])
        ws.add_data_validation(validation)
        validation.add(v["range"])
    for f in spec.get("conditional_formats", []):
        bounds(f["range"])
        if f["type"] == "color_scale":
            rule = ColorScaleRule(
                start_type="min",
                start_color=f.get("start_color", "FFFFFF"),
                end_type="max",
                end_color=f.get("end_color", "63BE7B"),
            )
        else:
            rule = CellIsRule(
                operator=f.get("operator", "greaterThan"),
                formula=f.get("formulas", ["0"]),
                fill=PatternFill("solid", fgColor=f.get("fill", "FFCCCC")),
            )
        ws.conditional_formatting.add(f["range"], rule)
    for c in spec.get("charts", []):
        chart(ws, c)
    for image in spec.get("images", []):
        from openpyxl.drawing.image import Image

        pic = Image(image["path"])
        pic.width, pic.height = (
            image.get("width", pic.width),
            image.get("height", pic.height),
        )
        ws.add_image(pic, image.get("anchor", "A1"))


def create(spec, output):
    wb = load(spec["template_path"]) if "template_path" in spec else Workbook()
    if "template_path" not in spec:
        wb.remove(wb.active)
    for s in spec["sheets"]:
        if s["name"] in wb:
            raise BadArguments("Duplicate sheet name")
        configure(wb.create_sheet(s["name"]), s)
    for name, ref in spec.get("defined_names", {}).items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))
    wb.calculation = CalcProperties(fullCalcOnLoad=True, forceFullCalc=True)
    wb.save(output)
    wb.close()


def edit(path, ops, output):
    wb = load(path)
    for op in ops:
        kind = op["op"]
        ws = wb[op["sheet"]] if "sheet" in op else None
        if kind == "set_cells":
            write_rows(ws, op["range"], op["rows"])
        elif kind == "configure_sheet":
            configure(ws, op["settings"])
        elif kind == "add_sheet":
            if op["spec"]["name"] in wb:
                raise BadArguments("Duplicate sheet name")
            configure(wb.create_sheet(op["spec"]["name"]), op["spec"])
        elif kind in {"copy_sheet", "rename_sheet"}:
            if op["name"] in wb:
                raise BadArguments("Duplicate sheet name")
            if kind == "copy_sheet":
                if ws._charts or ws._images or ws.tables:
                    raise BadArguments(
                        "copy_sheet refuses charts/images/tables that openpyxl does not copy"
                    )
                import copy

                copied = wb.copy_worksheet(ws)
                copied.title = op["name"]
                copied.data_validations = copy.deepcopy(ws.data_validations)
                copied.conditional_formatting = copy.deepcopy(ws.conditional_formatting)
                copied.freeze_panes = ws.freeze_panes
                copied.auto_filter = copy.deepcopy(ws.auto_filter)
            else:
                if (
                    any(c.data_type == "f" for s in wb for row in s for c in row)
                    or wb.defined_names
                    or any(s._charts for s in wb)
                ):
                    raise BadArguments(
                        "rename_sheet refuses formula/name/chart dependencies; update references explicitly in a new workbook"
                    )
                ws.title = op["name"]
        elif kind == "delete_sheet":
            if len(wb.worksheets) == 1:
                raise BadArguments("Workbook must retain a sheet")
            if (
                any(c.data_type == "f" for s in wb for row in s for c in row)
                or wb.defined_names
                or ws._charts
                or ws._images
            ):
                raise BadArguments(
                    "delete_sheet refuses formula/name/chart/image dependencies"
                )
            wb.remove(ws)
        elif kind == "define_name":
            wb.defined_names.add(DefinedName(op["name"], attr_text=op["reference"]))
        elif kind == "import_csv":
            with open(op["path"], encoding=op["encoding"], newline="") as stream:
                for r, row in enumerate(
                    csv.reader(stream, delimiter=op["delimiter"]), 1
                ):
                    if r > 10000 or len(row) > 1000:
                        raise BadArguments("CSV exceeds 10000 rows / 1000 columns")
                    for c, value in enumerate(row, 1):
                        cell(ws.cell(r, c), value)
    wb.calculation = CalcProperties(fullCalcOnLoad=True, forceFullCalc=True)
    wb.save(output)
    wb.close()
    return []


def inspect(path, args):
    # Read-only avoids allocating a giant rectangular sparse worksheet.
    wb = load_workbook(path, read_only=True, data_only=False)
    cached = load_workbook(path, read_only=True, data_only=True)
    summary = {
        "sheets": [
            {"name": s.title, "rows": s.max_row, "columns": s.max_column} for s in wb
        ],
        "defined_names": {k: v.attr_text for k, v in wb.defined_names.items()},
        "formulas": "cached_value is last saved result, not freshly calculated",
    }
    if args.get("mode") == "overview":
        wb.close()
        cached.close()
        return {"summary": summary, "items": []}
    ws = wb[args.get("sheet", wb.sheetnames[0])]
    if "range" in args:
        a, b, c, d = bounds(args["range"])
    else:
        a, b, c, d = 1, 1, min(ws.max_column or 1, 100), min(ws.max_row or 1, 10000)
    if (c - a + 1) * (d - b + 1) > 100000:
        raise BadArguments("Read at most 100000 cells; select a smaller range")
    items = []
    if args.get("mode") != "overview":
        for formula_row, value_row in zip(
            ws.iter_rows(min_row=b, max_row=d, min_col=a, max_col=c),
            cached[ws.title].iter_rows(min_row=b, max_row=d, min_col=a, max_col=c),
        ):
            for formula_cell, value_cell in zip(formula_row, value_row):
                if formula_cell.value is not None:
                    item = {
                        "sheet": ws.title,
                        "id": formula_cell.coordinate,
                        "value": formula_cell.value,
                        "number_format": formula_cell.number_format,
                    }
                    if formula_cell.data_type == "f":
                        item.update(
                            formula=formula_cell.value, cached_value=value_cell.value
                        )
                    items.append(item)
    wb.close()
    cached.close()
    return {"summary": summary, "items": items}


def export_csv(path, output, args):
    if not all(k in args for k in ["sheet", "encoding", "delimiter"]):
        raise BadArguments("CSV requires sheet, encoding and delimiter")
    wb = load_workbook(path, read_only=True, data_only=False)
    ws = wb[args["sheet"]]
    if (ws.max_row or 1) > 10000 or (ws.max_column or 1) > 1000:
        raise BadArguments("CSV export exceeds 10000 rows / 1000 columns")
    with open(output, "w", encoding=args["encoding"], newline="") as stream:
        writer = csv.writer(stream, delimiter=args["delimiter"])
        for row in ws.iter_rows(values_only=True):
            writer.writerow(row)
    wb.close()


def verify_shared_string_transform(before, after, operations):
    """Validate every untouched shared-string cell before allowing inlineStr save."""
    import posixpath

    from defusedxml import ElementTree as ET

    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

    def strings(blobs):
        shared = [
            "".join(n.text or "" for n in si.iter() if n.tag.endswith("}t"))
            for si in ET.fromstring(
                blobs.get(
                    "xl/sharedStrings.xml",
                    b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
                )
            )
        ]
        rels = {
            r.attrib["Id"]: r.attrib["Target"]
            for r in ET.fromstring(blobs["xl/_rels/workbook.xml.rels"])
        }
        values = {}
        for sheet in ET.fromstring(blobs["xl/workbook.xml"]).find("s:sheets", ns):
            rid = sheet.attrib[
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            ]
            path = rels[rid]
            path = (
                path.lstrip("/")
                if path.startswith("/")
                else posixpath.normpath("xl/" + path)
            )
            for c in ET.fromstring(blobs[path]).findall(".//s:c", ns):
                if c.attrib.get("t") == "s":
                    value = shared[int(c.find("s:v", ns).text)]
                elif c.attrib.get("t") == "inlineStr":
                    value = "".join(
                        n.text or "" for n in c.iter() if n.tag.endswith("}t")
                    )
                else:
                    continue
                values[sheet.attrib["name"], c.attrib["r"]] = value
        return values

    changed = set()
    renamed = {}
    deleted = set()
    for op in operations:
        if op["op"] == "rename_sheet":
            renamed[op["sheet"]] = op["name"]
        elif op["op"] == "delete_sheet":
            deleted.add(op["sheet"])
        elif op["op"] == "set_cells":
            a, b, _, _ = bounds(op["range"])
            from openpyxl.utils import get_column_letter

            for r, row in enumerate(op["rows"], b):
                for c in range(a, a + len(row)):
                    changed.add((op["sheet"], f"{get_column_letter(c)}{r}"))
        elif op["op"] in {"configure_sheet", "import_csv"}:
            # These intentionally replace cells; compare all other worksheets.
            deleted.add(op["sheet"])
    after_values = strings(after)
    for (sheet, coordinate), value in strings(before).items():
        if sheet in deleted or (sheet, coordinate) in changed:
            continue
        if after_values.get((renamed.get(sheet, sheet), coordinate)) != value:
            raise BadArguments(
                "Shared string would be lost: " + sheet + "!" + coordinate
            )
