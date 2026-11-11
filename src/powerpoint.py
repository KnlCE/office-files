"""Native PowerPoint slides, object edits and relationship-safe duplication."""

import copy
import math
import re

from astra_plugin_sdk import BadArguments
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.opc.packuri import PackURI
from pptx.util import Inches, Pt

from .safety import text_replace


def check_geometry(slide, shape, spec):
    """Validate against the actual template canvas, not a guessed pixel size."""
    canvas = slide.part.package.presentation_part.presentation
    x, y, width, height = shape.left, shape.top, shape.width, shape.height
    if (
        not all(math.isfinite(v) for v in (x, y, width, height))
        or width < 0
        or height < 0
    ):
        raise BadArguments(
            "PowerPoint geometry must be finite with nonnegative dimensions (inches)"
        )
    tolerance = 2  # two EMUs account for rounding at an exact slide edge
    outside = (
        x < -tolerance
        or y < -tolerance
        or x + width > canvas.slide_width + tolerance
        or y + height > canvas.slide_height + tolerance
    )
    if outside and not spec.get("allow_outside_slide", False):
        raise BadArguments(
            f"Object geometry is outside the slide ({canvas.slide_width / 914400:.3f} x {canvas.slide_height / 914400:.3f} INCHES). "
            f"Received x={x / 914400:g}, y={y / 914400:g}, width={width / 914400:g}, height={height / 914400:g}. "
            "Use INCHES, not pixels; e.g. x=0.8, y=1.2, width=11.7, height=0.8. "
            "Set allow_outside_slide=true only for intentional cropping. No output was saved."
        )


def layout_report(prs):
    """Geometric diagnostics work without LibreOffice; never claim visual QA."""
    count, outside, empty = 0, [], []
    for i, slide in enumerate(prs.slides):
        content = False
        for shape in slide.shapes:
            count += 1
            content |= (
                bool(shape.text.strip())
                if shape.has_text_frame
                and shape.shape_type
                in {MSO_SHAPE_TYPE.TEXT_BOX, MSO_SHAPE_TYPE.PLACEHOLDER}
                else True
            )
            if (
                shape.left < -2
                or shape.top < -2
                or shape.left + shape.width > prs.slide_width + 2
                or shape.top + shape.height > prs.slide_height + 2
            ):
                outside.append({"slide": i, "shape_id": shape.shape_id})
        if not content:
            empty.append(i)
    return {
        "units": "inches",
        "slides": len(prs.slides),
        "objects": count,
        "outside_slide_count": len(outside),
        "outside_slide": outside[:20],
        "empty_slide_count": len(empty),
        "empty_slides": empty[:20],
        "visual_verification": "not_performed",
    }


def layout_warnings(report):
    warnings = []
    if report["outside_slide_count"]:
        warnings.append(
            f"{report['outside_slide_count']} object(s) extend outside the slide; content may be invisible/cropped. Geometry is in inches, not pixels."
        )
    if report["empty_slide_count"]:
        warnings.append(
            f"{report['empty_slide_count']} slide(s) have no visible text or content objects."
        )
    return warnings


def font(target, spec):
    for key in ["name", "bold", "italic", "underline"]:
        if key in spec:
            setattr(target, key, spec[key])
    if "size" in spec:
        target.size = Pt(spec["size"])
    if "color" in spec:
        target.color.rgb = RGBColor.from_string(spec["color"])


def text(frame, spec):
    if "text" in spec or "runs" in spec:
        frame.clear()
        lines = spec.get("text", "").split("\n")
        if "runs" in spec:
            lines = [""]
        for i, line in enumerate(lines):
            p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            for run_spec in spec.get("runs", [{"text": line}]):
                r = p.add_run()
                r.text = run_spec["text"]
                font(r.font, {**spec.get("font", {}), **run_spec.get("font", {})})
    else:
        for p in frame.paragraphs:
            for r in p.runs:
                font(r.font, spec.get("font", {}))
    if "alignment" in spec:
        for p in frame.paragraphs:
            p.alignment = {
                "left": PP_ALIGN.LEFT,
                "right": PP_ALIGN.RIGHT,
                "center": PP_ALIGN.CENTER,
                "justify": PP_ALIGN.JUSTIFY,
            }[spec["alignment"]]


def paragraphs(prs):
    def walk(shapes):
        for s in shapes:
            if s.has_text_frame:
                yield from s.text_frame.paragraphs
            if s.has_table:
                for row in s.table.rows:
                    for c in row.cells:
                        yield from c.text_frame.paragraphs
            if hasattr(s, "shapes"):
                yield from walk(s.shapes)

    for slide in prs.slides:
        yield from walk(slide.shapes)
        if slide.has_notes_slide:
            yield from slide.notes_slide.notes_text_frame.paragraphs


def chart_data(spec):
    if spec["type"] == "scatter":
        data = XyChartData()
        for s in spec.get("series", []):
            x = s.get("x_values", [])
            if len(x) != len(s["values"]):
                raise BadArguments("Scatter x_values and values must have equal length")
            series = data.add_series(s.get("name", "Series"))
            for a, b in zip(x, s["values"]):
                series.add_data_point(a, b)
    else:
        data = CategoryChartData()
        data.categories = spec.get("categories", [])
        for s in spec.get("series", []):
            if len(s["values"]) != len(spec.get("categories", [])):
                raise BadArguments("Chart categories and series must have equal length")
            data.add_series(s.get("name", "Series"), s["values"])
    return data


def add_object(slide, spec):
    kind = spec["type"]
    canvas = slide.part.package.presentation_part.presentation
    slide_width, slide_height = (
        canvas.slide_width / 914400,
        canvas.slide_height / 914400,
    )
    left, top = min(1, slide_width / 10), min(1, slide_height / 10)
    geometry = [
        Inches(spec.get(k, default))
        for k, default in [
            ("x", left),
            ("y", top),
            ("width", min(5, slide_width - 2 * left)),
            ("height", min(1, slide_height - 2 * top)),
        ]
    ]
    if kind == "placeholder":
        try:
            s = slide.placeholders[spec["index"]]
        except KeyError as exc:
            raise BadArguments("Placeholder index absent in selected layout") from exc
    elif kind == "text":
        s = slide.shapes.add_textbox(*geometry)
    elif kind == "shape":
        shape = spec.get("shape", "rectangle")
        if shape == "line":
            s = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                geometry[0],
                geometry[1],
                geometry[0] + geometry[2],
                geometry[1] + geometry[3],
            )
        else:
            s = slide.shapes.add_shape(
                {
                    "rectangle": MSO_SHAPE.RECTANGLE,
                    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
                    "ellipse": MSO_SHAPE.OVAL,
                    "arrow": MSO_SHAPE.RIGHT_ARROW,
                }[shape],
                *geometry,
            )
        if "fill" in spec and shape != "line":
            s.fill.solid()
            s.fill.fore_color.rgb = RGBColor.from_string(spec["fill"])
    elif kind == "image":
        s = slide.shapes.add_picture(spec["path"], *geometry)
    elif kind == "table":
        rows = spec["rows"]
        if not rows or not rows[0] or any(len(r) != len(rows[0]) for r in rows):
            raise BadArguments("Table requires nonempty rectangular rows")
        s = slide.shapes.add_table(len(rows), len(rows[0]), *geometry)
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                s.table.cell(r, c).text = "" if value is None else str(value)
    elif kind == "chart":
        cs = spec["chart"]
        s = slide.shapes.add_chart(
            {
                "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
                "line": XL_CHART_TYPE.LINE,
                "pie": XL_CHART_TYPE.PIE,
                "scatter": XL_CHART_TYPE.XY_SCATTER,
            }[cs["type"]],
            *geometry,
            chart_data(cs),
        )
        if "title" in cs:
            s.chart.has_title = True
            s.chart.chart_title.text_frame.text = cs["title"]
    else:
        raise BadArguments("Unknown object type")
    if "name" in spec:
        s.name = spec["name"]
    check_geometry(slide, s, spec)
    if s.has_text_frame:
        text(s.text_frame, spec)
    return s


def add_slide(prs, spec):
    layout = spec.get("layout", 6)
    if layout >= len(prs.slide_layouts):
        raise BadArguments("Layout index outside template")
    slide = prs.slides.add_slide(prs.slide_layouts[layout])
    if "title" in spec:
        if slide.shapes.title is not None:
            slide.shapes.title.text = spec["title"]
        else:
            width, height = prs.slide_width / 914400, prs.slide_height / 914400
            left, top = min(0.6, width / 10), min(0.3, height / 10)
            add_object(
                slide,
                {
                    "type": "text",
                    "text": spec["title"],
                    "x": left,
                    "y": top,
                    "width": min(8, width - 2 * left),
                    "height": min(0.8, height - 2 * top),
                    "font": {"size": 28},
                },
            )
    if "background" in spec:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(spec["background"])
    if "notes" in spec:
        slide.notes_slide.notes_text_frame.text = spec["notes"]
    for s in spec.get("objects", []):
        add_object(slide, s)
    return slide


def fill(prs, values):
    for key, value in values.items():
        text_replace(paragraphs(prs), "{{" + key + "}}", value, True)
    if any(re.search(r"\{\{[^{}]+\}\}", p.text) for p in paragraphs(prs)):
        raise BadArguments("Unfilled template placeholders")


def create(spec, output):
    prs = Presentation(spec.get("template_path"))
    if "slide_size" in spec:
        prs.slide_width = Inches(13.333333 if spec["slide_size"] == "16:9" else 10)
        prs.slide_height = Inches(7.5)
    elif "template_path" not in spec:
        prs.slide_width, prs.slide_height = Inches(13.333333), Inches(7.5)
    for s in spec["slides"]:
        add_slide(prs, s)
    if "values" in spec:
        fill(prs, spec["values"])
    prs.save(output)


def duplicate(prs, source):
    dest = prs.slides.add_slide(source.slide_layout)
    for shape in list(dest.shapes):
        shape._element.getparent().remove(shape._element)
    package = prs.part.package
    used = {str(p.partname) for p in package.iter_parts()}
    clones = {}

    def clone_part(part):
        if part in clones:
            return clones[part]
        old = str(part.partname)
        pattern = re.sub(r"\d+(?=\.[^.]+$)", "%d", old)
        if "%d" not in pattern:
            stem, dot, suffix = old.rpartition(".")
            pattern = stem + "-copy%d" + dot + suffix
        i = 1
        while pattern % i in used:
            i += 1
        partname = PackURI(pattern % i)
        used.add(str(partname))
        new = type(part).load(partname, part.content_type, package, part.blob)
        clones[part] = new
        mapping = {}
        for rel in part.rels.values():
            target = rel.target_ref if rel.is_external else clone_part(rel.target_part)
            mapping[rel.rId] = new.relate_to(target, rel.reltype, rel.is_external)
        if hasattr(new, "_element"):
            remap(new._element, mapping)
        return new

    mapping = {}
    for rel in source.part.rels.values():
        if rel.reltype.endswith("/notesSlide") or rel.reltype.endswith("/slideLayout"):
            continue
        target = rel.target_ref if rel.is_external else rel.target_part
        if not rel.is_external and (
            "/chart" in rel.reltype or "/package" in rel.reltype
        ):
            target = clone_part(target)
        mapping[rel.rId] = dest.part.relate_to(target, rel.reltype, rel.is_external)
    for shape in source.shapes:
        element = copy.deepcopy(shape._element)
        remap(element, mapping)
        dest.shapes._spTree.insert_element_before(element, "p:extLst")
    if source.has_notes_slide:
        dest.notes_slide.notes_text_frame.text = (
            source.notes_slide.notes_text_frame.text
        )
    if (
        source._element.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}cSld/{http://schemas.openxmlformats.org/presentationml/2006/main}bg"
        )
        is not None
    ):
        bg = copy.deepcopy(source._element.cSld.bg)
        remap(bg, mapping)
        dest._element.cSld.insert(0, bg)
    return dest


def remap(element, mapping):
    namespace = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    for node in element.iter():
        for attr, value in list(node.attrib.items()):
            if attr.startswith(namespace) and value in mapping:
                node.set(attr, mapping[value])


def removable_parts(part, seen=None):
    seen = set() if seen is None else seen
    if str(part.partname) in seen:
        return seen
    seen.add(str(part.partname))
    for rel in part.rels.values():
        if not rel.is_external and not rel.reltype.endswith(
            ("/slideLayout", "/notesMaster", "/slide")
        ):
            removable_parts(rel.target_part, seen)
    return seen


def edit(path, ops, output):
    prs = Presentation(path)
    deleted, changed = set(), set()
    for op in ops:
        kind = op["op"]
        if "slide" in op and op["slide"] >= len(prs.slides):
            raise BadArguments("Slide index out of bounds")
        slide = prs.slides[op["slide"]] if "slide" in op else None
        shape = None
        if "shape_id" in op:
            shape = next(
                (s for s in slide.shapes if s.shape_id == op["shape_id"]), None
            )
            if shape is None:
                raise BadArguments("Unknown shape_id")
        if kind == "replace_text":
            text_replace(
                paragraphs(prs), op["find"], op["replace"], op.get("replace_all", False)
            )
        elif kind == "fill_template":
            fill(prs, op["values"])
        elif kind == "add_slide":
            add_slide(prs, op["spec"])
        elif kind == "duplicate_slide":
            duplicate(prs, slide)
        elif kind == "delete_slide":
            deleted.update(removable_parts(slide.part))
            sid = prs.slides._sldIdLst[op["slide"]]
            prs.part.drop_rel(sid.rId)
            prs.slides._sldIdLst.remove(sid)
        elif kind == "move_slide":
            if op["to"] >= len(prs.slides):
                raise BadArguments("Destination slide index out of bounds")
            sid = prs.slides._sldIdLst[op["slide"]]
            prs.slides._sldIdLst.remove(sid)
            prs.slides._sldIdLst.insert(op["to"], sid)
        elif kind == "add_object":
            add_object(slide, op["spec"])
        elif kind == "delete_object":
            shape._element.getparent().remove(shape._element)
        elif kind == "set_text":
            if not shape.has_text_frame:
                raise BadArguments("Shape has no text frame")
            text(shape.text_frame, op)
        elif kind == "set_cell":
            if (
                not shape.has_table
                or op["row"] >= len(shape.table.rows)
                or op["column"] >= len(shape.table.columns)
            ):
                raise BadArguments("Invalid table cell")
            text(shape.table.cell(op["row"], op["column"]).text_frame, op)
        elif kind == "set_geometry":
            for key, attr in [
                ("x", "left"),
                ("y", "top"),
                ("width", "width"),
                ("height", "height"),
            ]:
                if key in op:
                    setattr(shape, attr, Inches(op[key]))
            if "name" in op:
                shape.name = op["name"]
            check_geometry(slide, shape, op)
        elif kind == "set_background":
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor.from_string(op["color"])
        elif kind == "set_notes":
            slide.notes_slide.notes_text_frame.text = op["text"]
        elif kind == "update_chart":
            if not shape.has_chart:
                raise BadArguments("Shape has no chart")
            changed.add(
                str(shape.chart.part.chart_workbook.xlsx_part.partname).lstrip("/")
            )
            shape.chart.replace_data(chart_data(op["chart"]))
    prs.save(output)
    return {"deleted": deleted, "changed": changed}


def inspect(path, args):
    prs = Presentation(path)
    items = []
    report = layout_report(prs)
    for i, slide in enumerate(prs.slides):
        items.append(
            {
                "slide": i,
                "type": "slide",
                "notes": slide.notes_slide.notes_text_frame.text
                if slide.has_notes_slide
                else "",
            }
        )
        for s in slide.shapes:
            item = {
                "slide": i,
                "shape_id": s.shape_id,
                "name": s.name,
                "x": s.left / 914400,
                "y": s.top / 914400,
                "width": s.width / 914400,
                "height": s.height / 914400,
            }
            if s.has_text_frame:
                item["text"] = s.text
            if s.has_table:
                item["rows"] = [[c.text for c in r.cells] for r in s.table.rows]
            if s.has_chart:
                item["chart"] = {
                    "type": str(s.chart.chart_type),
                    "series": [
                        {"name": se.name, "values": list(se.values)}
                        for se in s.chart.series
                    ],
                }
            if s.is_placeholder:
                item["placeholder_index"] = s.placeholder_format.idx
            items.append(item)
    return {
        "summary": {
            "slides": len(prs.slides),
            "width_inches": prs.slide_width / 914400,
            "height_inches": prs.slide_height / 914400,
            "layouts": [l.name for l in prs.slide_layouts],
            "layout_validation": report,
        },
        "warnings": layout_warnings(report),
        "items": items,
    }
