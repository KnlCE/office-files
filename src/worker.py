"""Private process protocol. Only writes to private job outputs, never user paths."""

import datetime
import json
import sys
import zipfile
from pathlib import Path

from astra_plugin_sdk import BadArguments, PluginError, Unavailable

from . import conversion, excel, pdf, powerpoint, word
from .safety import (
    EXTENSIONS,
    limitations,
    package,
    preservation,
    revision,
    writable_package,
)

FORMATS = {"word": word, "excel": excel, "powerpoint": powerpoint}

# Files a person LOOKS at. Package checks and geometry say a file opens; only a
# rendered page says it reads well — and the model reads every tool result, so the
# result is where the owed check is spelled out.
LOOKED_AT = {".pptx": "slide", ".docx": "page"}


def next_step(output, tool, settings):
    """What checking a freshly written deck or document still needs, or None."""
    unit = LOOKED_AT.get(output.suffix)
    if unit is None or not tool.endswith(("_create", "_edit")):
        return None
    try:
        conversion.engine(settings)
    except Unavailable:
        return (
            "Not seen yet, and it cannot be: rendering needs LibreOffice, which is not installed "
            "or not configured (libreoffice_path). Only the package and the geometry were "
            "checked - say so when you report this file."
        )
    edit = tool.rsplit("_", 1)[0] + "_edit"
    return (
        f"Not seen yet. Render it with office_render, look at every {unit} with read_image, "
        f"fix what looks wrong with {edit} (revision above), and report it done only after "
        f"that."
    )


def serialize(value):
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.isoformat()
    return str(value)


def inspect_file(path, args, settings):
    if path.suffix == ".pdf":
        result, issues = pdf.inspect(path, args), []
    else:
        blobs = package(path, settings)
        result = FORMATS[EXTENSIONS[path.suffix]].inspect(path, args)
        issues = limitations(blobs)
    result.update(revision=revision(path), format=path.suffix[1:], limitations=issues)
    budget = int(settings.get("max_result_kb", 24)) * 1024
    summary = result["summary"]
    if (
        len(json.dumps(summary, ensure_ascii=False, default=serialize).encode("utf-8"))
        > budget // 2
    ):
        # Preserve counts, advertise excerpt instead of blowing the model budget.
        result["summary"] = {
            "truncated": True,
            "excerpt": json.dumps(summary, ensure_ascii=False, default=serialize)[
                : budget // 10
            ],
            "hint": "Select a sheet/range for content; workbook metadata exceeded the result budget",
        }
    query = args.get("query", "").casefold()
    items = [
        item
        for item in result["items"]
        if not query
        or query in json.dumps(item, ensure_ascii=False, default=serialize).casefold()
    ]
    start, limit = args.get("offset", 0), args.get("limit", 50)
    result["total_items"] = len(items)
    result["items"] = []
    for item in items[start : start + limit]:
        tentative = {**result, "items": result["items"] + [item]}
        if (
            len(
                json.dumps(tentative, ensure_ascii=False, default=serialize).encode(
                    "utf-8"
                )
            )
            > budget - 512
        ):
            if not result["items"]:
                # Keep ID and slice text; omit massive table/chart payloads.
                compact = {
                    k: v
                    for k, v in item.items()
                    if k in {"id", "page", "slide", "shape_id", "sheet", "type", "name"}
                }
                compact = {
                    k: v[:64] if isinstance(v, str) else v for k, v in compact.items()
                }
                content = json.dumps(item, ensure_ascii=False, default=serialize)
                compact.update(
                    content_excerpt=content[: min(2000, budget // 8)], truncated=True
                )
                result["items"].append(compact)
            break
        result["items"].append(item)
    result["next_offset"] = (
        start + len(result["items"])
        if start + len(result["items"]) < len(items)
        else None
    )
    if args.get("mode") == "overview":
        result["items"] = []
        result["next_offset"] = None
    # UTF-8, rather than character counts, is the actual transport budget.
    while (
        len(json.dumps(result, ensure_ascii=False, default=serialize).encode("utf-8"))
        > budget
    ):
        shortened = False
        for container, key in [
            (result["summary"], "excerpt"),
            *[(item, "content_excerpt") for item in result["items"]],
        ]:
            if container.get(key):
                container[key] = container[key][: len(container[key]) // 2]
                shortened = True
        if not shortened:
            result["summary"] = {"truncated": True}
            if (
                len(
                    json.dumps(result, ensure_ascii=False, default=serialize).encode(
                        "utf-8"
                    )
                )
                > budget
            ):
                raise BadArguments(
                    "Inspection metadata exceeds response budget; select a narrower range"
                )
    return result


def run(name, args, settings):
    # Parent validated user arguments; paths have since been snapshotted.
    if name == "office_inspect":
        path = Path(args["input_path"])
        if path.suffix not in EXTENSIONS and path.suffix != ".pdf":
            raise BadArguments(
                "Inspect supports DOCX/XLSX/PPTX/PDF; convert legacy formats first"
            )
        return inspect_file(path, args, settings)
    if name == "office_render":
        return conversion.render(args["input_path"], args["output_dir"], args, settings)
    output = Path(args["output_path"])
    before, deleted, details = {}, set(), {}
    if name == "pdf_edit":
        if Path(args["input_path"]).suffix != ".pdf" or output.suffix != ".pdf":
            raise BadArguments("pdf_edit requires PDF input and output")
        details = pdf.edit(args["input_path"], args["operations"], output)
        details["operations_applied"] = len(args["operations"])
    elif name.endswith(("_create", "_edit")):
        prefix, operation = name.rsplit("_", 1)
        expected = {"word": ".docx", "excel": ".xlsx", "powerpoint": ".pptx"}[prefix]
        if output.suffix != expected:
            raise BadArguments("Output extension must be " + expected)
        module = FORMATS[prefix]
        if operation == "create":
            template = args["spec"].get("template_path")
            if template:
                if Path(template).suffix != expected:
                    raise BadArguments("Template format does not match tool")
                before = writable_package(Path(template), settings)
            module.create(args["spec"], output)
        else:
            source = Path(args["input_path"])
            if source.suffix != expected:
                raise BadArguments("Input extension must be " + expected)
            before = writable_package(source, settings)
            changes = module.edit(source, args["operations"], output)
            deleted = changes["deleted"] if isinstance(changes, dict) else changes
            if isinstance(changes, dict):
                before = {
                    k: v for k, v in before.items() if k not in changes["changed"]
                }
            details["operations_applied"] = len(args["operations"])
        if prefix == "excel":
            details["warnings"] = [
                "Formulas are not evaluated by openpyxl; recalculate with LibreOffice or open in Excel"
            ]
    elif name == "office_export_pdf":
        if output.suffix != ".pdf":
            raise BadArguments("PDF output must end in .pdf")
        conversion.convert(args["input_path"], output, settings)
    elif name == "excel_recalculate":
        if Path(args["input_path"]).suffix != ".xlsx" or output.suffix != ".xlsx":
            raise BadArguments("Recalculation requires XLSX input and output")
        before = writable_package(Path(args["input_path"]), settings)
        details = conversion.recalculate(args["input_path"], output, settings)
    elif name == "office_convert":
        ext = Path(args["input_path"]).suffix
        if ext == ".xlsx" and output.suffix == ".csv":
            package(Path(args["input_path"]), settings)
            excel.export_csv(args["input_path"], output, args)
        else:
            mapping = {".doc": ".docx", ".xls": ".xlsx", ".ppt": ".pptx"}
            if mapping.get(ext) != output.suffix:
                raise BadArguments(
                    "Convert supports DOC→DOCX, XLS→XLSX, PPT→PPTX, XLSX→CSV"
                )
            conversion.convert(args["input_path"], output, settings)
            details["warnings"] = [
                "Converted copy; layout/features can differ from original"
            ]
    if output.suffix in EXTENSIONS:
        after = package(output, settings)
        if before:
            allowed = {p.lstrip("/") for p in deleted}
            if output.suffix == ".xlsx":
                if (
                    "xl/sharedStrings.xml" in before
                    and "xl/sharedStrings.xml" not in after
                ):
                    excel.verify_shared_string_transform(
                        before, after, args.get("operations", [])
                    )
                    allowed.add("xl/sharedStrings.xml")
                # Calculation chains are an optimization, not formula/value data.
                allowed.add("xl/calcChain.xml")
            # Relationship files follow explicitly deleted PPT parts.
            for p in list(allowed):
                directory, file = p.rsplit("/", 1)
                allowed.add(directory + "/_rels/" + file + ".rels")
            preservation(before, after, allowed)
        # Reopen with the native library, independent of ZIP health checking.
        module = FORMATS[EXTENSIONS[output.suffix]]
        opened = (
            module.Document(output)
            if output.suffix == ".docx"
            else module.load(output)
            if output.suffix == ".xlsx"
            else module.Presentation(output)
        )
        if output.suffix == ".pptx":
            report = powerpoint.layout_report(opened)
            details["layout_validation"] = report
            details.setdefault("warnings", []).extend(
                powerpoint.layout_warnings(report)
            )
        if hasattr(opened, "close"):
            opened.close()
    elif output.suffix == ".pdf":
        with pdf.open_pdf(output) as document:
            details["pages"] = len(document)
    step = next_step(output, name, settings)
    if step:
        details["next_step"] = step
    return {
        "output_path": str(output),
        "revision": revision(output),
        "validation": "reopened_and_package_checked"
        if output.suffix in EXTENSIONS
        else "output_checked",
        **details,
    }


def main():
    job = Path(sys.argv[1])
    try:
        request = json.loads(sys.stdin.buffer.read().decode("utf-8"))
        result = run(request["name"], request["args"], request["settings"])
    except PluginError as exc:
        result = {"error": {"type": type(exc).__name__, "message": str(exc)}}
    except (ValueError, KeyError, IndexError, TypeError, zipfile.BadZipFile) as exc:
        result = {"error": {"type": "BadArguments", "message": str(exc)[:2000]}}
    except OSError as exc:
        result = {
            "error": {
                "type": "Unavailable",
                "message": "File unavailable/locked: " + str(exc)[:2000],
            }
        }
    except Exception as exc:  # noqa: BLE001 -- isolate unexpected library failures at the worker boundary
        result = {
            "error": {
                "type": "InternalError",
                "message": type(exc).__name__ + ": " + str(exc)[:2000],
            }
        }
    (job / "response.json").write_text(
        json.dumps(result, ensure_ascii=False, default=serialize), encoding="utf-8"
    )


if __name__ == "__main__":
    main()
