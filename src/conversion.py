"""Optional LibreOffice engine: isolated profile, no macros/link updates."""

import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from astra_plugin_sdk import BadArguments, Timeout, Unavailable
from lxml import etree

from .safety import limitations, package


def engine(settings):
    configured = settings.get("libreoffice_path", "")
    candidates = (
        [configured]
        if configured
        else [shutil.which("soffice"), shutil.which("libreoffice")]
    )
    if not configured and os.name == "nt":
        for root in [
            os.environ.get("ProgramFiles", ""),
            os.environ.get("ProgramFiles(x86)", ""),
        ]:
            if root:
                candidates.append(str(Path(root) / "LibreOffice/program/soffice.exe"))
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(Path(candidate).resolve())
    raise Unavailable(
        "LibreOffice unavailable. Install LibreOffice or configure libreoffice_path; DOCX/XLSX/PPTX tools remain available."
    )


PROFILE = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>1</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>2</value></prop></item>
</oor:items>"""


def convert(input_path, output_path, settings):
    input_path, output_path = Path(input_path), Path(output_path)
    if input_path.suffix in [".docx", ".xlsx", ".pptx"]:
        blobs = package(input_path, settings)
        bad = limitations(blobs)
        if any(x in bad for x in ["macros", "embedded_or_activex_objects"]):
            raise BadArguments(
                "Rendering/conversion refuses macros and active embedded content"
            )
        for name, data in blobs.items():
            if name.endswith(".xml") and (
                b"WEBSERVICE(" in data.upper() or b"DDE(" in data.upper()
            ):
                raise BadArguments(
                    "Conversion refuses formulas that can access external services"
                )
    exe = engine(settings)
    with tempfile.TemporaryDirectory(prefix="office-engine-") as folder:
        root = Path(folder)
        profile = root / "profile"
        (profile / "user").mkdir(parents=True)
        (profile / "user/registrymodifications.xcu").write_text(
            PROFILE, encoding="utf-8"
        )
        source = root / ("source" + input_path.suffix)
        shutil.copyfile(input_path, source)
        out = root / "converted"
        out.mkdir()
        fmt = {
            ".pdf": "pdf",
            ".docx": "docx:Office Open XML Text",
            ".xlsx": "xlsx:Calc MS Excel 2007 XML",
            ".pptx": "pptx:Impress MS PowerPoint 2007 XML",
        }[output_path.suffix]
        try:
            result = subprocess.run(
                [
                    exe,
                    "-env:UserInstallation=" + profile.as_uri(),
                    "--headless",
                    "--nologo",
                    "--nodefault",
                    "--nolockcheck",
                    "--nofirststartwizard",
                    "--convert-to",
                    fmt,
                    "--outdir",
                    str(out),
                    str(source),
                ],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=min(110, int(settings.get("timeout_secs", 120))),
                creationflags=0x08000000 if os.name == "nt" else 0,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise Timeout("LibreOffice conversion timed out") from exc
        converted = out / (source.stem + output_path.suffix)
        if (
            result.returncode
            or not converted.is_file()
            or converted.stat().st_size == 0
        ):
            raise Unavailable(
                "LibreOffice could not convert this file; it may be encrypted, unsupported or damaged"
            )
        shutil.copyfile(converted, output_path)


def render(input_path, output_dir, args, settings):

    out = Path(output_dir)
    out.mkdir()
    pdf = out / "preview.pdf"
    if Path(input_path).suffix == ".pdf":
        # Already a PDF: PDFium renders it directly, LibreOffice is not needed.
        shutil.copyfile(input_path, pdf)
    else:
        convert(input_path, pdf, settings)
    images = []
    from .pdf import open_pdf

    with open_pdf(pdf) as doc:
        offset = args.get("offset", 0)
        if offset >= len(doc):
            raise BadArguments("Page offset outside document")
        for i in range(offset, min(len(doc), offset + args.get("limit", 5))):
            page = doc[i]
            try:
                scale = args.get("scale", 1.5)
                if page.get_width() * page.get_height() * scale * scale > 20_000_000:
                    raise BadArguments("Rendered page exceeds 20 million pixels")
                bitmap = page.render(scale=scale)
                dest = out / f"page-{i + 1:04}.png"
                try:
                    bitmap.to_pil().save(dest)
                finally:
                    bitmap.close()
                images.append(str(dest))
            finally:
                page.close()
        total = len(doc)
    return {
        "output_dir": str(out),
        "pdf_path": str(pdf),
        "images": images,
        "pages": total,
        "next_offset": offset + len(images) if offset + len(images) < total else None,
        "warnings": []
        if Path(input_path).suffix == ".pdf"
        else ["LibreOffice layout can differ from Microsoft Office"],
    }


def recalculate(input_path, output_path, settings):
    """Only transplant cached formula values; never replace layout with LO's save."""
    from defusedxml import ElementTree as ET

    blobs = package(Path(input_path), settings)
    if limitations(blobs):
        raise BadArguments("Recalculation refuses unsupported workbook content")
    with tempfile.TemporaryDirectory() as tmp:
        calculated = Path(tmp) / "calculated.xlsx"
        convert(input_path, calculated, settings)
        computed = package(calculated, settings)
        ns = {
            "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }
        workbook = ET.fromstring(blobs["xl/workbook.xml"])
        rels = {
            r.attrib["Id"]: r.attrib["Target"]
            for r in ET.fromstring(blobs["xl/_rels/workbook.xml.rels"])
        }
        computed_wb = ET.fromstring(computed["xl/workbook.xml"])
        computed_rels = {
            r.attrib["Id"]: r.attrib["Target"]
            for r in ET.fromstring(computed["xl/_rels/workbook.xml.rels"])
        }
        import posixpath

        def sheet_path(target):
            return (
                target.lstrip("/")
                if target.startswith("/")
                else posixpath.normpath("xl/" + target)
            )

        computed_sheets = {
            s.attrib["name"]: sheet_path(computed_rels[s.attrib["{" + ns["r"] + "}id"]])
            for s in computed_wb.find("s:sheets", ns)
        }
        strings = []
        if "xl/sharedStrings.xml" in computed:
            strings = [
                "".join(si.itertext())
                for si in ET.fromstring(computed["xl/sharedStrings.xml"])
            ]
        errors, count = [], 0
        for sheet in workbook.find("s:sheets", ns):
            target = rels[sheet.attrib["{" + ns["r"] + "}id"]]
            name = sheet_path(target)
            xml = etree.fromstring(blobs[name])
            calculated_xml = ET.fromstring(
                computed[computed_sheets[sheet.attrib["name"]]]
            )
            values = {c.attrib["r"]: c for c in calculated_xml.findall(".//s:c", ns)}
            for node in xml.findall(".//s:c", ns):
                if node.find("s:f", ns) is None:
                    continue
                cell = values.get(node.attrib["r"])
                if cell is None:
                    raise Unavailable(
                        "LibreOffice returned no cached cell for " + node.attrib["r"]
                    )
                cached_v = cell.find("s:v", ns)
                kind = cell.attrib.get("t", "n")
                value = cached_v.text if cached_v is not None else None
                if value is None and kind not in {"str", "inlineStr"}:
                    raise Unavailable(
                        "LibreOffice returned no cached value for " + node.attrib["r"]
                    )
                if kind == "s":
                    value, kind = strings[int(value)], "str"
                elif kind == "inlineStr":
                    value, kind = "".join(cell.find("s:is", ns).itertext()), "str"
                value = value or ""
                v = node.find("s:v", ns)
                if v is None:
                    v = etree.SubElement(node, "{" + ns["s"] + "}v")
                if kind == "e":
                    node.set("t", "e")
                    if len(errors) < 100:
                        errors.append(
                            {
                                "sheet": sheet.attrib["name"],
                                "cell": node.attrib["r"],
                                "error": value,
                            }
                        )
                elif kind in {"b", "str"}:
                    node.set("t", kind)
                else:
                    node.attrib.pop("t", None)
                v.text = str(value)
                count += 1
            blobs[name] = etree.tostring(
                xml, xml_declaration=True, encoding="UTF-8", standalone=True
            )
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out:
            for name, data in blobs.items():
                out.writestr(name, data)
    return {
        "recalculated_formulas": count,
        "formula_errors": errors,
        "warnings": [
            "Results calculated by LibreOffice; unsupported Excel functions may return errors"
        ],
    }
