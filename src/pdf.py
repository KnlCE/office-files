"""PDF through PDFium: per-page text, outline, page operations. No LibreOffice."""

from contextlib import ExitStack

import pypdfium2 as pdfium
from astra_plugin_sdk import BadArguments

# One page of a dense document is ~3-5k characters; anything above this is a
# pathological text layer and is cut rather than starving the result budget.
PAGE_TEXT_MAX = 8000
MAX_OUTPUT_PAGES = 10000
TOC_MAX = 100


def open_pdf(path):
    try:
        return pdfium.PdfDocument(str(path))
    except pdfium.PdfiumError as exc:
        message = str(exc)
        if "password" in message.lower():
            raise BadArguments(
                "PDF is password-protected; encrypted PDFs are not supported"
            ) from exc
        raise BadArguments("Not a readable PDF: " + message[:500]) from exc


def page_text(page):
    textpage = page.get_textpage()
    try:
        return textpage.get_text_range().replace("\r\n", "\n").replace("\r", "\n")
    finally:
        textpage.close()


def inspect(path, args):
    with open_pdf(path) as doc:
        overview = args.get("mode") == "overview"
        items, empty = [], []
        for index in range(len(doc)):
            page = doc[index]
            try:
                width, height = page.get_size()
                item = {
                    "page": index,
                    "type": "page",
                    "width_pt": round(width, 1),
                    "height_pt": round(height, 1),
                    "rotation": page.get_rotation(),
                }
                label = doc.get_page_label(index)
                if label and label != str(index + 1):
                    item["label"] = label
                text = page_text(page)
            finally:
                page.close()
            if not text.strip():
                empty.append(index)
            if not overview:
                item["text"] = text[:PAGE_TEXT_MAX]
                if len(text) > PAGE_TEXT_MAX:
                    item.update(truncated=True, text_chars=len(text))
            items.append(item)
        toc = []
        for entry in doc.get_toc():
            if len(toc) >= TOC_MAX:
                break
            toc.append(
                {
                    "title": entry.get_title(),
                    "level": entry.level,
                    "page": entry.get_dest().get_index() if entry.get_dest() else None,
                }
            )
        summary = {
            "pages": len(doc),
            "pdf_version": doc.get_version(),
            "metadata": doc.get_metadata_dict(skip_empty=True),
            "outline": toc,
            "pages_without_text": empty[:200],
            "units": "page indices zero-based; sizes in points (1/72 inch)",
        }
        if empty:
            summary["hint"] = (
                "Pages without a text layer are usually scans: render them with "
                "office_render and read the PNG files with read_image."
            )
        return {"summary": summary, "items": items}


def check(where, indices, size):
    for index in indices:
        if not 0 <= index < size:
            raise BadArguments(f"{where}: page {index} is outside 0..{size - 1}")


def edit(source, operations, output):
    """Rebuild the document from an ordered list of (document, page, rotation)."""
    with ExitStack() as stack:
        opened = {}

        def load(path):
            if path not in opened:
                opened[path] = stack.enter_context(open_pdf(path))
            return opened[path]

        pages = [(str(source), i, 0) for i in range(len(load(str(source))))]
        for n, op in enumerate(operations):
            kind = op["op"]
            where = f"operations[{n}] {kind}"
            if kind == "delete_pages":
                check(where, op["pages"], len(pages))
                drop = set(op["pages"])
                pages = [p for i, p in enumerate(pages) if i not in drop]
                if not pages:
                    raise BadArguments(where + ": cannot delete every page")
            elif kind == "keep_pages":
                check(where, op["pages"], len(pages))
                if not op["pages"]:
                    raise BadArguments(where + ": keep at least one page")
                pages = [pages[i] for i in op["pages"]]
            elif kind == "rotate_pages":
                targets = op.get("pages", range(len(pages)))
                check(where, targets, len(pages))
                for i in set(targets):
                    path, index, rotation = pages[i]
                    pages[i] = (path, index, (rotation + op["angle"]) % 360)
            elif kind == "move_page":
                check(where, [op["page"], op["to"]], len(pages))
                pages.insert(op["to"], pages.pop(op["page"]))
            elif kind == "insert_pdf":
                other = load(op["path"])
                selected = op.get("pages", range(len(other)))
                check(where, selected, len(other))
                at = op.get("at", len(pages))
                check(where, [at], len(pages) + 1)
                pages[at:at] = [(op["path"], i, 0) for i in selected]
            if len(pages) > MAX_OUTPUT_PAGES:
                raise BadArguments(f"Result exceeds {MAX_OUTPUT_PAGES} pages")

        result = stack.enter_context(pdfium.PdfDocument.new())
        # Import consecutive pages of one source together so shared fonts and
        # images are copied once, not once per page.
        start = 0
        while start < len(pages):
            end = start
            while end + 1 < len(pages) and pages[end + 1][0] == pages[start][0]:
                end += 1
            result.import_pages(
                opened[pages[start][0]], [p[1] for p in pages[start : end + 1]]
            )
            start = end + 1
        for i, (_, _, rotation) in enumerate(pages):
            if rotation:
                page = result[i]
                try:
                    page.set_rotation((page.get_rotation() + rotation) % 360)
                finally:
                    page.close()
        result.save(str(output))
    return {
        "pages": len(pages),
        "warnings": [
            (
                "Pages are copied into a new PDF: bookmarks, document metadata and "
                "interactive form fields of the source are not carried over"
            )
        ],
    }
