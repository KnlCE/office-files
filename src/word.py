"""Word document creation, structural inspection and atomic batch editing."""

from astra_plugin_sdk import BadArguments
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

from .safety import text_replace


def font(target, spec):
    for key in ["name", "bold", "italic", "underline"]:
        if key in spec:
            setattr(target, key, spec[key])
    if "size" in spec:
        target.size = Pt(spec["size"])
    if "color" in spec:
        target.color.rgb = RGBColor.from_string(spec["color"])


def paragraph(p, spec):
    if "style" in spec:
        p.style = spec["style"]
    if "text" in spec or "runs" in spec:
        p.clear()
        for r in spec.get("runs", [{"text": spec.get("text", "")}]):
            run = p.add_run(r["text"])
            font(run.font, {**spec.get("font", {}), **r.get("font", {})})
    elif "font" in spec:
        for run in p.runs:
            font(run.font, spec["font"])
    if "alignment" in spec:
        p.alignment = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }[spec["alignment"]]


def objects(doc):
    """IDs bind to original XML objects, retained across operations in one batch."""
    result = {}
    seen = set()

    def walk(container, prefix):
        for i, block in enumerate(container.iter_inner_content()):
            key = f"{prefix}/{i}"
            result[key] = block
            if isinstance(block, Table):
                for r, row in enumerate(block.rows):
                    for c, cell in enumerate(row.cells):
                        if cell._tc not in seen:
                            seen.add(cell._tc)
                            walk(cell, f"{key}/cell/{r}/{c}")

    walk(doc, "body")
    for i, section in enumerate(doc.sections):
        for kind in [
            "header",
            "footer",
            "first_page_header",
            "first_page_footer",
            "even_page_header",
            "even_page_footer",
        ]:
            part = getattr(section, kind)
            if part.is_linked_to_previous:
                continue
            walk(part, f"section/{i}/{kind}")
    return result


def paragraphs(doc):
    return [p for p in objects(doc).values() if isinstance(p, Paragraph)]


def fill(doc, values):
    for key, value in values.items():
        text_replace(paragraphs(doc), "{{" + key + "}}", value, True)
    import re

    unfilled = sorted(
        set(re.findall(r"\{\{[^{}]+\}\}", "\n".join(p.text for p in paragraphs(doc))))
    )
    if unfilled:
        raise BadArguments("Unfilled template placeholders: " + ", ".join(unfilled))


def block(doc, spec):
    kind = spec["type"]
    if kind in ["paragraph", "heading"]:
        p = (
            doc.add_heading(level=spec.get("level", 1))
            if kind == "heading"
            else doc.add_paragraph()
        )
        paragraph(p, spec)
        return [p._p]
    if kind in ["bullet_list", "numbered_list"]:
        return [
            doc.add_paragraph(
                text, "List Bullet" if kind == "bullet_list" else "List Number"
            )._p
            for text in spec["items"]
        ]
    if kind == "table":
        rows = spec["rows"]
        if not rows or not rows[0] or any(len(r) != len(rows[0]) for r in rows):
            raise BadArguments("Table requires a nonempty rectangular rows array")
        table = doc.add_table(rows=len(rows), cols=len(rows[0]))
        table.style = spec.get("style", "Table Grid")
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                table.cell(r, c).text = "" if value is None else str(value)
        return [table._tbl]
    if kind == "image":
        shape = doc.add_picture(spec["path"], width=Mm(spec.get("width_mm", 120)))
        return [shape._inline.getparent().getparent().getparent()]
    if kind == "page_break":
        return [doc.add_page_break()._p]
    raise BadArguments("Unknown block type")


def header_footer(doc, spec):
    for kind in ["header", "footer"]:
        if kind in spec:
            for section in doc.sections:
                container = getattr(section, kind)
                paragraph(container.paragraphs[0], {"text": spec[kind]})


def style(doc, name, spec):
    s = (
        doc.styles[name]
        if name in doc.styles
        else doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    )
    font(s.font, spec)


def create(spec, output):
    doc = Document(spec.get("template_path"))
    for section in doc.sections:
        for key, value in spec.get("page", {}).items():
            attr = {"width_mm": "page_width", "height_mm": "page_height"}.get(
                key, key.removesuffix("_mm")
            )
            # Schema uses margin_top; python-docx uses top_margin.
            if attr.startswith("margin_"):
                attr = attr[7:] + "_margin"
            setattr(section, attr, Mm(value))
    for name, values in spec.get("styles", {}).items():
        style(doc, name, values)
    header_footer(doc, spec)
    for b in spec.get("blocks", []):
        block(doc, b)
    if "values" in spec:
        fill(doc, spec["values"])
    doc.save(output)


def edit(path, ops, output):
    doc = Document(path)
    ids = objects(doc)
    for op in ops:
        kind = op["op"]
        if kind == "replace_text":
            text_replace(
                paragraphs(doc), op["find"], op["replace"], op.get("replace_all", False)
            )
        elif kind == "fill_template":
            fill(doc, op["values"])
        elif kind == "insert_block":
            elements = block(doc, op["block"])
            if "before_id" in op:
                target = ids.get(op["before_id"])
                if not target or target._element.getparent() is None:
                    raise BadArguments("Unknown/deleted before_id")
                for element in elements:
                    target._element.addprevious(element)
        elif kind in {"delete", "set_paragraph", "set_cell"}:
            target = ids.get(op["id"])
            if target is None or target._element.getparent() is None:
                raise BadArguments("Unknown/deleted object id: " + op["id"])
            if kind == "delete":
                target._element.getparent().remove(target._element)
            elif kind == "set_paragraph":
                if not isinstance(target, Paragraph):
                    raise BadArguments("id must identify a paragraph")
                paragraph(target, op)
            else:
                if not isinstance(target, Table):
                    raise BadArguments("id must identify a table")
                if op["row"] >= len(target.rows) or op["column"] >= len(target.columns):
                    raise BadArguments("Cell index outside table")
                paragraph(target.cell(op["row"], op["column"]).paragraphs[0], op)
        elif kind == "set_header_footer":
            header_footer(doc, op)
        elif kind == "set_style":
            style(doc, op["name"], op["font"])
    doc.save(output)
    return []


def inspect(path, args):
    doc = Document(path)
    mapping = objects(doc)
    items = []
    for key, item in mapping.items():
        if isinstance(item, Paragraph):
            items.append(
                {
                    "id": key,
                    "type": "paragraph",
                    "text": item.text,
                    "style": item.style.name,
                    "runs": [
                        {"text": r.text, "bold": r.bold, "italic": r.italic}
                        for r in item.runs
                    ],
                }
            )
        else:
            items.append(
                {
                    "id": key,
                    "type": "table",
                    "rows": len(item.rows),
                    "columns": len(item.columns),
                }
            )
    return {
        "summary": {
            "sections": len(doc.sections),
            "paragraphs": sum(isinstance(v, Paragraph) for v in mapping.values()),
            "tables": sum(isinstance(v, Table) for v in mapping.values()),
            "images": len(doc.inline_shapes),
        },
        "items": items,
    }
