"""Run from the plugin root: python -m examples.generate [--libreoffice PATH]."""

import argparse
import json
from pathlib import Path

from astra_plugin_sdk.testing import Harness
from PIL import Image, ImageDraw

from src.plugin import OfficeFiles


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--libreoffice")
    parser.add_argument("--format", choices=["word", "excel", "powerpoint"])
    args = parser.parse_args()
    root = Path(__file__).resolve().parent
    asset = root / "mark.png"
    img = Image.new("RGB", (240, 120), "#225599")
    draw = ImageDraw.Draw(img)
    draw.rectangle((25, 25, 65, 95), fill="#FFFFFF")
    draw.rectangle((85, 50, 125, 95), fill="#61CDBB")
    draw.rectangle((145, 15, 185, 95), fill="#F4C66B")
    img.save(asset)
    specs = {
        "word": {
            "header": "Office Files • квартальный отчёт",
            "footer": "Демонстрационный документ",
            "page": {"margin_left_mm": 22, "margin_right_mm": 22},
            "blocks": [
                {"type": "heading", "text": "Результаты квартала", "level": 1},
                {
                    "type": "paragraph",
                    "runs": [
                        {"text": "Выручка выросла на ", "font": {"size": 12}},
                        {"text": "38%", "font": {"bold": True, "color": "225599"}},
                        {"text": ". План следующего квартала — сохранить темп."},
                    ],
                },
                {"type": "image", "path": str(asset), "width_mm": 42},
                {
                    "type": "table",
                    "style": "Light Shading Accent 1",
                    "rows": [
                        ["Месяц", "Выручка, тыс. ₽"],
                        ["Январь", 42],
                        ["Февраль", 58],
                        ["Март", 65],
                    ],
                },
                {"type": "heading", "text": "Следующие шаги", "level": 2},
                {
                    "type": "bullet_list",
                    "items": [
                        "Запустить новый продукт",
                        "Сократить сроки поставки",
                        "Проверить результат в конце квартала",
                    ],
                },
            ],
        },
        "excel": {
            "sheets": [
                {
                    "name": "Продажи",
                    "rows": [
                        ["Месяц", "Выручка, тыс. ₽"],
                        ["Январь", 42],
                        ["Февраль", 58],
                        ["Март", 65],
                        ["Итого", {"formula": "SUM(B2:B4)", "number_format": "0.00"}],
                    ],
                    "column_widths": {"A": 22, "B": 24},
                    "freeze_panes": "B2",
                    "cells": {
                        "A1": {
                            "value": "Месяц",
                            "font": {"bold": True, "color": "FFFFFF"},
                        },
                        "B1": {
                            "value": "Выручка, тыс. ₽",
                            "font": {"bold": True, "color": "FFFFFF"},
                        },
                    },
                    "tables": [{"name": "Sales", "range": "A1:B4"}],
                    "charts": [
                        {
                            "type": "bar",
                            "title": "Выручка по месяцам",
                            "data_range": "B1:B4",
                            "categories_range": "A2:A4",
                            "anchor": "A8",
                            "width": 12.5,
                            "height": 8,
                        }
                    ],
                    "images": [
                        {"path": str(asset), "anchor": "D1", "width": 140, "height": 70}
                    ],
                }
            ],
            "defined_names": {"Revenue": "'Продажи'!$B$2:$B$4"},
        },
        "powerpoint": {
            "slide_size": "16:9",
            "slides": [
                {
                    "layout": 6,
                    "background": "FFFFFF",
                    "notes": "Объяснить рост выручки и план следующего квартала.",
                    "objects": [
                        {
                            "type": "text",
                            "text": "Результаты квартала",
                            "font": {"size": 32, "bold": True, "color": "225599"},
                            "x": 0.65,
                            "y": 0.4,
                            "width": 10,
                            "height": 0.8,
                        },
                        {
                            "type": "text",
                            "text": "Выручка, тыс. ₽ • январь — март",
                            "font": {"size": 16, "color": "555555"},
                            "x": 0.65,
                            "y": 1.3,
                            "width": 10,
                            "height": 0.5,
                        },
                        {
                            "type": "image",
                            "path": str(asset),
                            "x": 10.8,
                            "y": 0.4,
                            "width": 1.8,
                            "height": 0.9,
                        },
                        {
                            "type": "table",
                            "rows": [
                                ["Месяц", "Выручка"],
                                ["Январь", 42],
                                ["Февраль", 58],
                                ["Март", 65],
                            ],
                            "x": 0.65,
                            "y": 2.2,
                            "width": 4.1,
                            "height": 2.5,
                        },
                        {
                            "type": "chart",
                            "chart": {
                                "type": "bar",
                                "categories": ["Январь", "Февраль", "Март"],
                                "series": [{"name": "Выручка", "values": [42, 58, 65]}],
                            },
                            "x": 5.1,
                            "y": 2,
                            "width": 7.5,
                            "height": 4.5,
                        },
                        {
                            "type": "text",
                            "text": "Следующий шаг: запуск нового продукта",
                            "font": {"size": 16, "color": "225599"},
                            "x": 0.65,
                            "y": 6.6,
                            "width": 11,
                            "height": 0.5,
                        },
                    ],
                }
            ],
        },
    }
    with Harness(OfficeFiles(), timeout=150) as h:
        if args.libreoffice:
            h.set_config({"libreoffice_path": args.libreoffice})
        for prefix, spec in specs.items():
            if args.format and prefix != args.format:
                continue
            ext = {"word": "docx", "excel": "xlsx", "powerpoint": "pptx"}[prefix]
            target = root / ("quarterly." + ext)
            result = h.call_tool(
                prefix + "_create", output_path=str(target), spec=spec, overwrite=True
            )
            if not result.success:
                raise RuntimeError(result.error)
            # Portable specifications: resolve this relative asset before tool calls.
            portable = json.dumps(spec, ensure_ascii=False).replace(
                str(asset).replace("\\", "\\\\"), "mark.png"
            )
            (root / (prefix + ".json")).write_text(portable, encoding="utf-8")
            if args.libreoffice:
                out = root / (prefix + "-preview")
                result = h.call_tool(
                    "office_render",
                    input_path=str(target),
                    output_dir=str(out),
                    limit=2,
                )
                if not result.success:
                    raise RuntimeError(result.error)
            print(target.name)


if __name__ == "__main__":
    main()
