"""Shared path, package, revision and preservation checks."""

import hashlib
import posixpath
import zipfile
from pathlib import Path

from astra_plugin_sdk import BadArguments, NotFound, Unauthorized
from defusedxml import ElementTree as ET

EXTENSIONS = {".docx": "word", ".xlsx": "excel", ".pptx": "powerpoint"}


def revision(path):
    h = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def checked_path(raw, settings, exists=False):
    p = Path(raw).expanduser().resolve()
    roots = settings.get("allowed_roots", "")
    roots = roots.splitlines() if isinstance(roots, str) else roots
    if roots and not any(
        p.is_relative_to(Path(r).expanduser().resolve()) for r in roots if r.strip()
    ):
        raise Unauthorized("Path outside allowed_roots: " + str(p))
    if exists and not p.is_file():
        raise NotFound("File not found: " + str(p))
    return p


def check_size(path, settings):
    if path.stat().st_size > int(settings.get("max_file_mb", 50)) * 1024 * 1024:
        raise BadArguments("File exceeds max_file_mb")


def package(path, settings):
    check_size(path, settings)
    try:
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            names = [i.filename for i in infos]
            if len(infos) > 20000 or len(names) != len(set(names)):
                raise BadArguments("Too many ZIP entries or duplicate entries")
            if (
                sum(i.file_size for i in infos)
                > int(settings.get("max_unpacked_mb", 256)) * 1024 * 1024
            ):
                raise BadArguments("ZIP exceeds max_unpacked_mb")
            for i in infos:
                if (
                    i.flag_bits & 1
                    or i.filename.startswith(("/", "\\"))
                    or ".." in i.filename.replace("\\", "/").split("/")
                ):
                    raise BadArguments("Encrypted or unsafe ZIP package")
            blobs = {i.filename: z.read(i) for i in infos}
    except (zipfile.BadZipFile, RuntimeError, OSError) as exc:
        raise BadArguments("Invalid Office ZIP package: " + str(exc)) from exc
    required = {
        ".docx": "word/document.xml",
        ".xlsx": "xl/workbook.xml",
        ".pptx": "ppt/presentation.xml",
    }.get(path.suffix.lower())
    if not required or required not in blobs or "[Content_Types].xml" not in blobs:
        raise BadArguments("Extension does not match Office package")
    for name, data in blobs.items():
        if name.endswith((".xml", ".rels")):
            try:
                root = ET.fromstring(data)
            except Exception as exc:
                raise BadArguments("Unsafe or malformed XML in " + name) from exc
            if name.endswith(".rels"):
                base = posixpath.dirname(posixpath.dirname(name))
                for rel in root:
                    target = rel.attrib.get("Target", "")
                    if rel.attrib.get("TargetMode") == "External" or target.startswith(
                        "#"
                    ):
                        continue
                    target = target.split("#")[0]
                    from urllib.parse import unquote

                    resolved = (
                        posixpath.normpath(posixpath.join(base, unquote(target)))
                        if not target.startswith("/")
                        else target[1:]
                    )
                    if resolved not in blobs:
                        raise BadArguments(
                            "Missing relationship target: " + name + " → " + target
                        )
    return blobs


def limitations(blobs):
    issues = set()
    for name, data in blobs.items():
        lower = name.lower()
        if "vbaproject" in lower:
            issues.add("macros")
        # PPT chart workbooks are supported; other embedded objects are not.
        if ("activex" in lower or "/embeddings/" in lower) and not (
            lower.startswith("ppt/embeddings/") and lower.endswith(".xlsx")
        ):
            issues.add("embedded_or_activex_objects")
        if "signatures" in lower:
            issues.add("digital_signatures")
        if "diagrams/" in lower:
            issues.add("SmartArt")
        if lower.endswith(".xml"):
            root = ET.fromstring(data)
            tags = {e.tag.split("}")[-1] for e in root.iter()}
            if lower.startswith("word/") and tags & {
                "ins",
                "del",
                "moveFrom",
                "moveTo",
                "pPrChange",
                "rPrChange",
                "tblPrChange",
            }:
                issues.add("tracked_changes")
            if lower.startswith("ppt/slides/") and tags & {"timing", "transition"}:
                issues.add("slide_animation_or_transition")
            if lower.startswith("xl/") and tags & {
                "extLst",
                "AlternateContent",
                "sp",
                "contentPart",
            }:
                issues.add("unsupported_excel_extensions_or_shapes")
            if (
                lower.startswith("xl/")
                and "graphicFrame" in tags
                and "chart" not in tags
            ):
                issues.add("unsupported_excel_extensions_or_shapes")
            if lower.startswith("xl/") and any(
                e.tag.split("}")[-1] == "f"
                and e.attrib.get("t") in {"array", "dataTable"}
                for e in root.iter()
            ):
                issues.add("array_or_data_table_formulas")
            if lower.startswith("xl/") and (
                "pivot" in lower or "slicer" in lower or "externalLink" in name
            ):
                issues.add("pivot_slicer_or_external_workbook")
    return sorted(issues)


def writable_package(path, settings):
    blobs = package(path, settings)
    issues = limitations(blobs)
    if issues:
        raise BadArguments(
            "Editing is refused to preserve unsupported content: " + ", ".join(issues)
        )
    return blobs


def preservation(before, after, deleted_parts=()):
    """No silent part loss. Untouched opaque resources must retain exact bytes."""
    deleted = set(deleted_parts)
    missing = set(before) - set(after) - deleted
    if missing:
        raise BadArguments(
            "Save would remove package parts: " + ", ".join(sorted(missing)[:10])
        )
    for name, data in before.items():
        if name in deleted:
            continue
        if not name.endswith((".xml", ".rels")) and after.get(name) != data:
            raise BadArguments("Save would change an opaque resource: " + name)


def replace_runs(paragraph, old, new, replace_all=False):
    if not old:
        raise BadArguments("find must not be empty")
    runs = list(paragraph.runs)
    text = "".join(r.text for r in runs)
    starts = []
    at = 0
    while (at := text.find(old, at)) >= 0:
        starts.append(at)
        at += len(old)
    if not replace_all and len(starts) > 1:
        raise BadArguments("Ambiguous text match; use replace_all=true")
    for start in reversed(starts):
        end = start + len(old)
        cursor = 0
        inserted = False
        for run in runs:
            length = len(run.text)
            lo, hi = max(0, start - cursor), min(length, end - cursor)
            if lo < hi:
                # run.text clears all XML children. Refuse mixed field/drawing runs
                # rather than silently losing a picture or a dynamic Word field.
                if any(
                    child.tag.split("}")[-1] not in {"rPr", "t", "tab", "br", "cr"}
                    or (child.tag.split("}")[-1] in {"br", "cr"} and child.attrib)
                    for child in run._r
                ):
                    raise BadArguments(
                        "Text replacement touches a run with drawing/field content; split the run in Word first"
                    )
                run.text = run.text[:lo] + (new if not inserted else "") + run.text[hi:]
                inserted = True
            cursor += length
    return len(starts)


def text_replace(paragraphs, old, new, all_matches=False):
    paragraphs = list(paragraphs)
    if not old:
        raise BadArguments("find must not be empty")
    total = sum("".join(r.text for r in p.runs).count(old) for p in paragraphs)
    if not total:
        raise BadArguments("Text not found: " + old)
    if total != 1 and not all_matches:
        raise BadArguments("Ambiguous text match; use replace_all=true")
    for p in paragraphs:
        replace_runs(p, old, new, True)
    return total
