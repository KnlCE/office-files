"""Regression for the real Astra call that supplied a 960x540 pixel canvas."""

import json

import pytest
from pptx import Presentation
from pptx.util import Inches
from test_plugin import call

from src.registry import TOOLS

pytest_plugins = ["test_plugin"]


def spec(x=0.8, y=1.2, width=11.7, height=0.8, **extra):
    return {
        "slide_size": "16:9",
        "slides": [
            {
                "background": "1B1F3B",
                "objects": [
                    {
                        "type": "text",
                        "text": "Астра",
                        "x": x,
                        "y": y,
                        "width": width,
                        "height": height,
                        "font": {"color": "FFFFFF", "size": 36},
                        **extra,
                    }
                ],
            }
        ],
    }


def test_real_pixel_coordinates_refused_before_output(h, tmp_path):
    output = tmp_path / "blank.pptx"
    result = h.call_tool(
        "powerpoint_create", output_path=str(output), spec=spec(60, 170, 840, 110)
    )
    assert not result.success and result.code == "BAD_ARGUMENTS"
    assert "INCHES, not pixels" in result.error and "13.333 x 7.500" in result.error
    assert not output.exists()


def test_inch_positions_reported_without_libreoffice(h, tmp_path):
    output = tmp_path / "good.pptx"
    made = call(h, "powerpoint_create", output_path=output, spec=spec())
    report = made["layout_validation"]
    assert report["objects"] == 1 and report["outside_slide_count"] == 0
    assert report["visual_verification"] == "not_performed"
    inspected = call(h, "office_inspect", input_path=output)
    assert inspected["summary"]["layout_validation"] == report
    shape = Presentation(output).slides[0].shapes[0]
    assert shape.text == "Астра" and shape.left == Inches(0.8)
    assert shape.left + shape.width <= Inches(13.333333)


@pytest.mark.parametrize(
    "x,y,width,height",
    [(-0.1, 0, 1, 1), (0, -0.1, 1, 1), (13, 1, 1, 1), (0.8, 7, 1, 1), (0, 0, 840, 110)],
)
def test_partial_overflow_refused(h, tmp_path, x, y, width, height):
    result = h.call_tool(
        "powerpoint_create",
        output_path=str(tmp_path / "out.pptx"),
        spec=spec(x, y, width, height),
    )
    assert not result.success


def test_intentional_crop_is_explicit_and_warned(h, tmp_path):
    path = tmp_path / "crop.pptx"
    made = call(
        h,
        "powerpoint_create",
        output_path=path,
        spec=spec(-0.2, 1, 1, 1, allow_outside_slide=True),
    )
    assert made["layout_validation"]["outside_slide_count"] == 1
    assert "outside" in made["warnings"][0]
    assert call(h, "office_inspect", input_path=path)["warnings"]


def test_template_actual_dimensions(h, tmp_path):
    template = tmp_path / "small.pptx"
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(5), Inches(3)
    prs.save(template)
    args = {"template_path": str(template), "slides": spec(4, 1, 2, 1)["slides"]}
    bad = h.call_tool(
        "powerpoint_create", output_path=str(tmp_path / "bad.pptx"), spec=args
    )
    assert not bad.success and "5.000 x 3.000" in bad.error
    args["slides"][0]["objects"][0]["width"] = 1
    good = call(h, "powerpoint_create", output_path=tmp_path / "good.pptx", spec=args)
    assert good["layout_validation"]["outside_slide_count"] == 0


def test_invalid_edit_batch_keeps_original_and_existing_output(h, tmp_path):
    path, out = tmp_path / "source.pptx", tmp_path / "out.pptx"
    created = call(h, "powerpoint_create", output_path=path, spec=spec())
    original = path.read_bytes()
    out.write_bytes(original)
    shape_id = Presentation(path).slides[0].shapes[0].shape_id
    result = h.call_tool(
        "powerpoint_edit",
        input_path=str(path),
        output_path=str(out),
        overwrite=True,
        revision=created["revision"],
        operations=[
            {"op": "set_text", "slide": 0, "shape_id": shape_id, "text": "New"},
            {"op": "set_geometry", "slide": 0, "shape_id": shape_id, "x": 60},
        ],
    )
    assert not result.success and "INCHES" in result.error
    assert path.read_bytes() == original and out.read_bytes() == original
    assert not list(tmp_path.glob("*.backup-*"))


def test_existing_bad_deck_inspected_and_repaired(h, tmp_path):
    bad = tmp_path / "old.pptx"
    # Model old successful code before geometry validation, not a malformed ZIP.
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(Inches(60), Inches(170), Inches(840), Inches(110))
    shape.text = "Астра"
    prs.save(bad)
    info = call(h, "office_inspect", input_path=bad)
    assert (
        info["summary"]["layout_validation"]["outside_slide_count"] == 1
        and info["warnings"]
    )
    repaired = call(
        h,
        "powerpoint_edit",
        input_path=bad,
        output_path=tmp_path / "fixed.pptx",
        revision=info["revision"],
        operations=[
            {
                "op": "set_geometry",
                "slide": 0,
                "shape_id": shape.shape_id,
                "x": 0.8,
                "y": 1.2,
                "width": 11.7,
                "height": 0.8,
            }
        ],
    )
    assert repaired["layout_validation"]["outside_slide_count"] == 0


def test_geometry_units_in_described_schema_and_help(h):
    definition = TOOLS["powerpoint_create"]
    assert "INCHES" in definition["description"]
    data = call(h, "office_help", tool="powerpoint_create", language="ru")
    assert "дюйм" in data["help"]
    objects = definition["schema"]["properties"]["spec"]["properties"]["slides"][
        "items"
    ]["properties"]["objects"]["items"]["oneOf"]
    for variant in objects:
        for key in ["x", "y", "width", "height"]:
            if key in variant["properties"]:
                assert "INCHES" in variant["properties"][key]["description"]
    assert len(json.dumps(data).encode()) < 28 * 1024


def test_empty_slide_warning(h, tmp_path):
    result = call(
        h,
        "powerpoint_create",
        output_path=tmp_path / "empty.pptx",
        spec={"slides": [{"background": "1B1F3B"}]},
    )
    assert result["layout_validation"]["empty_slide_count"] == 1 and result["warnings"]


def test_small_template_auto_title_and_default_objects(h, tmp_path):
    template = tmp_path / "small.pptx"
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(5), Inches(3)
    prs.save(template)
    result = call(
        h,
        "powerpoint_create",
        output_path=tmp_path / "result.pptx",
        spec={
            "template_path": str(template),
            "slides": [
                {"title": "Title", "objects": [{"type": "text", "text": "Body"}]}
            ],
        },
    )
    assert result["layout_validation"]["outside_slide_count"] == 0
    saved = Presentation(tmp_path / "result.pptx")
    assert [s.text for s in saved.slides[0].shapes] == ["Title", "Body"]


def test_filled_shape_is_content(h, tmp_path):
    result = call(
        h,
        "powerpoint_create",
        output_path=tmp_path / "shape.pptx",
        spec={
            "slides": [
                {
                    "objects": [
                        {
                            "type": "shape",
                            "fill": "FF0000",
                            "x": 1,
                            "y": 1,
                            "width": 5,
                            "height": 3,
                        }
                    ]
                }
            ]
        },
    )
    assert (
        result["layout_validation"]["empty_slide_count"] == 0 and not result["warnings"]
    )
