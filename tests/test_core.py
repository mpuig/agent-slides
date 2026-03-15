from __future__ import annotations

import base64
from pathlib import Path

from slides_cli import OperationBatch, Presentation
from slides_cli.agentic import (
    DeckPlan,
    DesignProfile,
    SlidePlan,
    compile_plan_to_operations,
    lint_design,
)


def test_create_add_text_and_chart_and_save(tmp_path: Path) -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)

    pres.add_text(
        slide_index=slide.index,
        text="Quarterly Results",
        left=0.5,
        top=0.3,
        width=8.5,
        height=0.8,
        font_size=26,
        bold=True,
    )
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3", "Q4"],
        series=[("Revenue", [10.0, 12.0, 14.0, 16.0])],
        left=0.8,
        top=1.4,
        width=8.0,
        height=4.0,
    )

    out = tmp_path / "deck.pptx"
    pres.save(out)

    assert out.exists()
    assert out.stat().st_size > 0


def test_deterministic_save_same_bytes(tmp_path: Path) -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_text(
        slide_index=slide.index,
        text="Deterministic",
        left=1,
        top=1,
        width=4,
        height=1,
    )

    out1 = tmp_path / "a.pptx"
    out2 = tmp_path / "b.pptx"
    pres.save(out1, deterministic=True)
    pres.save(out2, deterministic=True)

    assert out1.read_bytes() == out2.read_bytes()


def test_operations_dry_run() -> None:
    pres = Presentation.create()

    batch = OperationBatch.model_validate(
        {
            "operations": [
                {"op": "add_slide", "layout_index": 6},
                {
                    "op": "add_text",
                    "slide_index": 0,
                    "text": "hello",
                    "left": 0.5,
                    "top": 0.5,
                    "width": 3,
                    "height": 1,
                },
            ]
        }
    )

    report = pres.apply_operations(batch, dry_run=True)
    assert report.ok
    assert pres.slide_count == 0
    assert all(evt.status == "planned" for evt in report.events)
    assert report.applied_count == 0
    assert report.failed_index is None


def test_transaction_rolls_back_on_failure() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)

    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "add_text",
                    "slide_index": 0,
                    "text": "before",
                    "left": 0.5,
                    "top": 0.5,
                    "width": 2,
                    "height": 1,
                },
                {
                    "op": "add_text",
                    "slide_index": 999,
                    "text": "bad",
                    "left": 0.5,
                    "top": 0.5,
                    "width": 2,
                    "height": 1,
                },
            ]
        }
    )

    baseline = pres.to_bytes(deterministic=True)
    report = pres.apply_operations(batch, transactional=True)

    assert not report.ok
    assert report.applied_count == 1
    assert report.failed_index == 1
    assert pres.to_bytes(deterministic=True) == baseline


def test_add_table_and_move_slide_and_notes(tmp_path: Path) -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    pres.add_slide(layout_index=6)

    pres.add_table(
        slide_index=0,
        rows=[["KPI", "Value"], ["Revenue", "$10M"], ["Margin", "22%"]],
        left=0.5,
        top=1.0,
        width=5.0,
        height=2.0,
    )
    pres.add_notes(slide_index=0, text="Narration for slide 1")
    pres.move_slide(from_index=0, to_index=1)

    out = tmp_path / "ops.pptx"
    pres.save(out)
    assert out.exists()


def test_semantic_replace_text_accepts_consistent_slide_id_and_uid() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    pres.add_text(
        slide_index=0,
        text="legacy roadmap",
        left=0.5,
        top=0.5,
        width=4,
        height=1,
    )
    inspect = pres.inspect()
    slide_uid = str(inspect["slides"][0]["slide_uid"])

    result = pres.semantic_replace_text(
        query="legacy",
        replacement="target",
        slide_id="slide-1",
        slide_uid=slide_uid,
    )

    assert result["replaced_paragraphs"] >= 1


def test_add_table_with_table_xml_preserves_style() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    table_xml = (
        '<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:tblPr firstRow="1" bandRow="1">'
        "<a:tableStyleId>{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}</a:tableStyleId>"
        "</a:tblPr>"
        '<a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>'
        '<a:tr h="457200">'
        "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"
        "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"
        "</a:tr>"
        "</a:tbl>"
    )
    pres.add_table(
        slide_index=slide.index,
        rows=[["A", "B"]],
        left=0.5,
        top=0.5,
        width=4.0,
        height=1.0,
        table_xml=table_xml,
    )
    shape = next(s for s in pres._prs.slides[0].shapes if getattr(s, "has_table", False))
    assert "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}" in shape._element.graphic.graphicData.tbl.xml


def test_repair_removes_template_tokens() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_text(
        slide_index=slide.index,
        text="Hello {{name}}",
        left=0.5,
        top=0.5,
        width=5,
        height=1,
    )
    before = pres.validate()
    assert any(i.code == "UNRESOLVED_TEMPLATE_TOKEN" for i in before.issues)
    after = pres.repair()
    assert after.ok
    assert all(i.code != "UNRESOLVED_TEMPLATE_TOKEN" for i in after.issues)


def test_set_core_properties_and_summary() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_text(
        slide_index=slide.index,
        text="Deck title",
        left=0.5,
        top=0.5,
        width=3.0,
        height=1.0,
    )
    pres.set_core_properties(title="T", subject="S", author="A", keywords="k1,k2")

    summary = pres.summarize()
    assert summary.slide_count == 1
    assert summary.shape_count >= 1
    assert summary.text_shape_count >= 1


def test_operation_set_core_properties() -> None:
    pres = Presentation.create()
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_core_properties",
                    "title": "Board Update",
                    "author": "Agent",
                }
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok
    assert report.applied_count == 1


def test_deterministic_across_fresh_presentations() -> None:
    ops_payload = {
        "operations": [
            {"op": "add_slide", "layout_index": 6},
            {
                "op": "add_text",
                "slide_index": 0,
                "text": "Same content",
                "left": 0.5,
                "top": 0.5,
                "width": 3,
                "height": 1,
            },
            {
                "op": "set_core_properties",
                "title": "Deck",
                "author": "Agent",
            },
        ]
    }
    batch = OperationBatch.model_validate(ops_payload)

    p1 = Presentation.create()
    p2 = Presentation.create()
    assert p1.apply_operations(batch).ok
    assert p2.apply_operations(batch).ok
    assert p1.to_bytes(deterministic=True) == p2.to_bytes(deterministic=True)
    assert p1.fingerprint() == p2.fingerprint()


def test_update_table_cell_operation() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_table(
        slide_index=slide.index,
        rows=[["A", "B"], ["1", "2"]],
        left=0.5,
        top=0.5,
        width=4.0,
        height=1.5,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "update_table_cell",
                    "slide_index": 0,
                    "table_index": 0,
                    "row": 1,
                    "col": 1,
                    "text": "99",
                }
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_update_chart_data_operation() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2"],
        series=[("Revenue", [10.0, 20.0])],
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "update_chart_data",
                    "slide_index": 0,
                    "chart_index": 0,
                    "categories": ["Q1", "Q2", "Q3"],
                    "series": [["Revenue", [10.0, 20.0, 30.0]]],
                }
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_chart_ops_allow_nullable_values_and_category_fallback() -> None:
    pres = Presentation.create()
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {"op": "add_slide", "layout_index": 6},
                {
                    "op": "add_bar_chart",
                    "slide_index": 0,
                    "categories": [],
                    "series": [["Revenue", [10.0, None, 30.0]]],
                    "style": "clustered",
                    "left": 0.5,
                    "top": 0.5,
                    "width": 4.0,
                    "height": 2.0,
                },
                {
                    "op": "update_chart_data",
                    "slide_index": 0,
                    "chart_index": 0,
                    "categories": [],
                    "series": [["Revenue", [None, 22.0, None, 28.0]]],
                },
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok


def test_add_bar_chart_horizontal_orientation() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["A", "B"],
        series=[("S1", [1.0, 2.0])],
        style="clustered",
        orientation="bar",
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    chart_shape = next(s for s in pres._prs.slides[0].shapes if getattr(s, "has_chart", False))
    assert str(chart_shape.chart.chart_type).startswith("BAR_CLUSTERED")


def test_layout_name_and_background_operation() -> None:
    pres = Presentation.create()
    first_layout_name = pres._prs.slide_layouts[0].name
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {"op": "add_slide", "layout_name": first_layout_name},
                {"op": "set_slide_background", "slide_index": 0, "color_hex": "00AA55"},
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok
    assert pres.slide_count == 1


def test_add_slide_falls_back_to_layout_index_when_name_missing() -> None:
    pres = Presentation.create()
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {"op": "add_slide", "layout_name": "__missing__", "layout_index": 6},
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok
    assert pres.slide_count == 1


def test_add_slide_hidden_flag() -> None:
    pres = Presentation.create()
    batch = OperationBatch.model_validate(
        {"operations": [{"op": "add_slide", "layout_index": 6, "hidden": True}]}
    )
    report = pres.apply_operations(batch)
    assert report.ok
    assert pres._prs.slides[0]._element.get("show") == "0"


def test_add_raw_shape_xml_operation() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_text(
        slide_index=slide.index,
        text="seed",
        left=1.0,
        top=1.0,
        width=2.0,
        height=0.5,
    )
    shape_xml = pres._prs.slides[0].shapes[0]._element.xml
    batch = OperationBatch.model_validate(
        {"operations": [{"op": "add_raw_shape_xml", "slide_index": 0, "shape_xml": shape_xml}]}
    )
    report = pres.apply_operations(batch)
    assert report.ok
    assert len(pres._prs.slides[0].shapes) >= 2


def test_add_raw_shape_xml_with_rel_images(tmp_path: Path) -> None:
    one_px_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO0N4p0AAAAASUVORK5CYII="
    )
    image_path = tmp_path / "pixel.png"
    image_path.write_bytes(one_px_png)

    pres = Presentation.create()
    s0 = pres.add_slide(layout_index=6)
    pres.add_image(
        slide_index=s0.index,
        path=image_path,
        left=1.0,
        top=1.0,
        width=1.0,
        height=1.0,
    )
    shape_xml = pres._prs.slides[0].shapes[0]._element.xml
    s1 = pres.add_slide(layout_index=6)
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "add_raw_shape_xml",
                    "slide_index": s1.index,
                    "shape_xml": shape_xml,
                    "rel_images": [["rId2", str(image_path)]],
                }
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok
    assert len(pres._prs.slides[s1.index].shapes) >= 1


def test_partname_template_generation() -> None:
    assert Presentation._partname_template("/ppt/diagrams/data11.xml") == (
        "/ppt/diagrams/data%d.xml"
    )
    assert Presentation._partname_template(
        "/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"
    ) == ("/ppt/embeddings/Microsoft_Excel_Worksheet%d.xlsx")


def test_add_raw_shape_xml_model_accepts_rel_parts() -> None:
    shape_xml = (
        "<p:sp xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' "
        "xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>"
        "<p:nvSpPr><p:cNvPr id='2' name='x'/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        "<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "add_raw_shape_xml",
                    "slide_index": 0,
                    "shape_xml": shape_xml,
                    "rel_parts": [
                        [
                            "rId2",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
                            "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
                            "/ppt/diagrams/data11.xml",
                            "/tmp/missing.xml",
                        ]
                    ],
                    "rel_external": [
                        [
                            "rId9",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                            "https://example.com",
                        ]
                    ],
                }
            ]
        }
    )
    assert batch.operations[0].op == "add_raw_shape_xml"


def test_clear_slides() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    pres.add_slide(layout_index=6)
    assert pres.slide_count == 2
    pres.clear_slides()
    assert pres.slide_count == 0


def test_set_placeholder_text_and_title_subtitle_ops() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=0)
    placeholders = pres.list_placeholders(slide_index=0)
    assert placeholders
    idxs = [p["idx"] for p in placeholders if p["has_text_frame"]]
    assert idxs

    batch = OperationBatch.model_validate(
        {
            "operations": [
                {"op": "set_title_subtitle", "slide_index": 0, "title": "Main Title"},
                {
                    "op": "set_placeholder_text",
                    "slide_index": 0,
                    "placeholder_idx": int(idxs[0]),
                    "text": "Overridden by idx",
                },
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok


def test_set_semantic_text_body() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=1)  # title and content layout
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_semantic_text",
                    "slide_index": 0,
                    "role": "title",
                    "text": "Semantic title",
                },
                {
                    "op": "set_semantic_text",
                    "slide_index": 0,
                    "role": "body",
                    "text": "Semantic body",
                },
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok


def test_set_chart_legend_operation() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2"],
        series=[("Revenue", [10.0, 12.0])],
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_chart_legend",
                    "slide_index": 0,
                    "chart_index": 0,
                    "visible": True,
                    "position": "bottom",
                    "include_in_layout": False,
                }
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_set_image_crop_operation(tmp_path: Path) -> None:
    one_px_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO0N4p0AAAAASUVORK5CYII="
    )
    image_path = tmp_path / "pixel.png"
    image_path.write_bytes(one_px_png)

    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_image(
        slide_index=slide.index,
        path=image_path,
        left=0.5,
        top=0.5,
        width=1.0,
        height=1.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_image_crop",
                    "slide_index": 0,
                    "image_index": 0,
                    "crop_left": 0.1,
                    "crop_right": 0.1,
                    "crop_top": 0.0,
                    "crop_bottom": 0.0,
                }
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_set_chart_title_axis_labels_ops() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2"],
        series=[("Revenue", [10.0, 12.0])],
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_chart_title",
                    "slide_index": 0,
                    "chart_index": 0,
                    "text": "Revenue Trend",
                },
                {
                    "op": "set_chart_axis_titles",
                    "slide_index": 0,
                    "chart_index": 0,
                    "category_title": "Quarter",
                    "value_title": "M EUR",
                },
                {
                    "op": "set_chart_data_labels",
                    "slide_index": 0,
                    "chart_index": 0,
                    "enabled": True,
                    "show_value": True,
                    "number_format": "#,##0.00",
                },
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_set_image_crop_invalid_value_fails(tmp_path: Path) -> None:
    one_px_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO0N4p0AAAAASUVORK5CYII="
    )
    image_path = tmp_path / "pixel.png"
    image_path.write_bytes(one_px_png)
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_image(
        slide_index=slide.index,
        path=image_path,
        left=0.5,
        top=0.5,
        width=1.0,
        height=1.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_image_crop",
                    "slide_index": 0,
                    "image_index": 0,
                    "crop_left": 3.0,
                }
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert not report.ok
    assert report.failed_index == 0


def test_set_placeholder_image_errors(tmp_path: Path) -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=0)
    missing = tmp_path / "missing.png"
    report = pres.apply_operations(
        OperationBatch.model_validate(
            {
                "operations": [
                    {
                        "op": "set_placeholder_image",
                        "slide_index": 0,
                        "placeholder_idx": 999,
                        "path": str(missing),
                    }
                ]
            }
        )
    )
    assert not report.ok


def test_placeholder_ops_fallback_to_geometry(tmp_path: Path) -> None:
    one_px_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO0N4p0AAAAASUVORK5CYII="
    )
    image_path = tmp_path / "pixel.png"
    image_path.write_bytes(one_px_png)

    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_placeholder_text",
                    "slide_index": 0,
                    "placeholder_idx": 999,
                    "text": "fallback",
                    "left": 1.0,
                    "top": 1.0,
                    "width": 2.0,
                    "height": 0.5,
                },
                {
                    "op": "set_placeholder_image",
                    "slide_index": 0,
                    "placeholder_idx": 999,
                    "path": str(image_path),
                    "left": 1.0,
                    "top": 2.0,
                    "width": 1.0,
                    "height": 1.0,
                    "crop_left": 0.1,
                },
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok


def test_placeholder_text_geometry_matches_existing_placeholder() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=0)
    title = pres._prs.slides[0].shapes.title
    assert title is not None
    before = len(pres._prs.slides[0].shapes)
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_placeholder_text",
                    "slide_index": 0,
                    "placeholder_idx": 999,
                    "text": "Geometry match title",
                    "left": title.left / 914400,
                    "top": title.top / 914400,
                    "width": title.width / 914400,
                    "height": title.height / 914400,
                }
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert report.ok
    assert len(pres._prs.slides[0].shapes) == before
    assert "Geometry match title" in (title.text or "")


def test_add_line_pie_area_doughnut_scatter_radar_and_bubble_chart_ops() -> None:
    pres = Presentation.create()
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {"op": "add_slide", "layout_index": 6},
                {
                    "op": "add_line_chart",
                    "slide_index": 0,
                    "categories": ["Q1", "Q2", "Q3"],
                    "series": [["Run-rate", [8.0, 9.5, 10.2]]],
                    "style": "line_markers",
                    "left": 0.5,
                    "top": 0.5,
                    "width": 4.2,
                    "height": 2.2,
                },
                {
                    "op": "add_pie_chart",
                    "slide_index": 0,
                    "categories": ["A", "B", "C"],
                    "series": [["Share", [40.0, 35.0, 25.0]]],
                    "style": "exploded",
                    "left": 5.0,
                    "top": 0.5,
                    "width": 4.0,
                    "height": 2.2,
                },
                {
                    "op": "add_area_chart",
                    "slide_index": 0,
                    "categories": ["Q1", "Q2", "Q3"],
                    "series": [["Area", [5.0, 7.0, 6.0]]],
                    "style": "stacked",
                    "left": 0.5,
                    "top": 3.0,
                    "width": 4.2,
                    "height": 2.2,
                },
                {
                    "op": "add_doughnut_chart",
                    "slide_index": 0,
                    "categories": ["X", "Y", "Z"],
                    "series": [["Mix", [20.0, 30.0, 50.0]]],
                    "style": "exploded",
                    "left": 5.0,
                    "top": 3.0,
                    "width": 4.0,
                    "height": 2.2,
                },
                {
                    "op": "add_scatter_chart",
                    "slide_index": 0,
                    "series": [["S", [[1.0, 1.5], [2.0, 2.5], [3.0, 3.2]]]],
                    "style": "smooth_no_markers",
                    "left": 0.5,
                    "top": 5.5,
                    "width": 4.2,
                    "height": 2.0,
                },
                {
                    "op": "add_radar_chart",
                    "slide_index": 0,
                    "categories": ["North", "South", "East", "West"],
                    "series": [["Coverage", [80.0, 65.0, 72.0, 90.0]]],
                    "style": "markers",
                    "left": 5.0,
                    "top": 5.5,
                    "width": 4.0,
                    "height": 2.0,
                },
                {
                    "op": "add_bubble_chart",
                    "slide_index": 0,
                    "series": [["Portfolio", [[1.0, 2.0, 5.0], [2.0, 2.5, 7.0], [3.0, 3.0, 9.0]]]],
                    "style": "bubble_3d",
                    "left": 0.5,
                    "top": 7.7,
                    "width": 4.2,
                    "height": 2.0,
                },
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_invalid_chart_style_fails() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    try:
        pres.add_line_chart(
            slide_index=slide.index,
            categories=["Q1", "Q2"],
            series=[("S", [1.0, 2.0])],
            style="not_a_style",
            left=0.5,
            top=0.5,
            width=4.0,
            height=2.0,
        )
        raise AssertionError("Expected style validation error")
    except Exception as exc:  # noqa: BLE001
        assert getattr(exc, "code", "") == "INVALID_LINE_STYLE"


def test_chart_axis_scale_plot_and_series_style_ops() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3"],
        series=[("Revenue", [10.0, 12.0, 14.0])],
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_chart_axis_scale",
                    "slide_index": 0,
                    "chart_index": 0,
                    "minimum": 0.0,
                    "maximum": 20.0,
                    "major_unit": 5.0,
                    "number_format": "#,##0",
                },
                {
                    "op": "set_chart_plot_style",
                    "slide_index": 0,
                    "chart_index": 0,
                    "vary_by_categories": False,
                    "gap_width": 120,
                    "overlap": 0,
                },
                {
                    "op": "set_chart_series_style",
                    "slide_index": 0,
                    "chart_index": 0,
                    "series_index": 0,
                    "fill_color_hex": "0A4280",
                    "invert_if_negative": True,
                },
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_add_media_op(tmp_path: Path) -> None:
    media_path = tmp_path / "clip.mp4"
    media_path.write_bytes(b"fake-video-bytes")
    poster_path = tmp_path / "poster.png"
    poster_path.write_bytes(
        base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO0N4p0AAAAASUVORK5CYII="
        )
    )

    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "add_media",
                    "slide_index": 0,
                    "path": str(media_path),
                    "left": 1.0,
                    "top": 1.0,
                    "width": 2.0,
                    "height": 2.0,
                    "mime_type": "video/mp4",
                    "poster_path": str(poster_path),
                }
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_combo_overlay_and_secondary_axis_view_ops() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "add_combo_chart_overlay",
                    "slide_index": 0,
                    "categories": ["Q1", "Q2", "Q3"],
                    "bar_series": [["Revenue", [10.0, 12.0, 14.0]]],
                    "line_series": [["Margin", [20.0, 22.0, 21.0]]],
                    "left": 0.5,
                    "top": 0.5,
                    "width": 8.0,
                    "height": 3.0,
                },
                {
                    "op": "set_chart_secondary_axis",
                    "slide_index": 0,
                    "chart_index": 1,
                    "enable": True,
                },
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_line_marker_trendline_errorbars_ops() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_line_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3"],
        series=[("Run-rate", [8.0, 9.0, 10.0])],
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_line_series_marker",
                    "slide_index": 0,
                    "chart_index": 0,
                    "series_index": 0,
                    "style": "diamond",
                    "size": 8,
                },
                {
                    "op": "set_chart_series_trendline",
                    "slide_index": 0,
                    "chart_index": 0,
                    "series_index": 0,
                    "trend_type": "linear",
                },
                {
                    "op": "set_chart_series_error_bars",
                    "slide_index": 0,
                    "chart_index": 0,
                    "series_index": 0,
                    "value": 0.2,
                    "direction": "y",
                    "bar_type": "both",
                },
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_secondary_axis_series_mapping_not_supported() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2"],
        series=[("Revenue", [1.0, 2.0])],
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_chart_secondary_axis",
                    "slide_index": 0,
                    "chart_index": 0,
                    "enable": True,
                    "series_indices": [0],
                }
            ]
        }
    )
    report = pres.apply_operations(batch)
    assert not report.ok


def test_chart_combo_secondary_mapping_op() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3"],
        series=[("Revenue", [10.0, 12.0, 14.0]), ("Margin", [20.0, 22.0, 21.0])],
        left=0.5,
        top=0.5,
        width=8.0,
        height=3.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_chart_combo_secondary_mapping",
                    "slide_index": 0,
                    "chart_index": 0,
                    "series_indices": [1],
                }
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_chart_combo_secondary_mapping_from_line_base() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_line_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3"],
        series=[("S1", [1.0, 2.0, 3.0]), ("S2", [2.0, 3.0, 4.0])],
        left=0.5,
        top=0.5,
        width=8.0,
        height=3.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {
                    "op": "set_chart_combo_secondary_mapping",
                    "slide_index": 0,
                    "chart_index": 0,
                    "series_indices": [1],
                }
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_chart_style_axis_options_and_line_style_ops() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_line_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3"],
        series=[("S1", [1.0, 2.0, 3.0])],
        left=0.5,
        top=0.5,
        width=6.0,
        height=3.0,
    )
    batch = OperationBatch.model_validate(
        {
            "operations": [
                {"op": "set_chart_style", "slide_index": 0, "chart_index": 0, "style_id": 4},
                {
                    "op": "set_chart_axis_options",
                    "slide_index": 0,
                    "chart_index": 0,
                    "axis": "value",
                    "reverse_order": False,
                    "major_tick_mark": "outside",
                    "minor_tick_mark": "none",
                    "tick_label_position": "next_to_axis",
                    "crosses": "minimum",
                },
                {
                    "op": "set_chart_data_labels_style",
                    "slide_index": 0,
                    "chart_index": 0,
                    "position": "above",
                    "show_legend_key": False,
                    "font_size": 10,
                },
                {
                    "op": "set_chart_series_line_style",
                    "slide_index": 0,
                    "chart_index": 0,
                    "series_index": 0,
                    "line_color_hex": "0055AA",
                    "line_width_pt": 1.5,
                },
                {
                    "op": "set_line_series_smooth",
                    "slide_index": 0,
                    "chart_index": 0,
                    "series_index": 0,
                    "smooth": True,
                },
            ]
        }
    )
    assert pres.apply_operations(batch).ok


def test_lint_skips_structural_slide_visual_rules() -> None:
    plan = DeckPlan(
        deck_title="Board Update",
        brief="Board Update",
        slides=[
            SlidePlan(
                slide_number=1,
                story_role="title",
                archetype_id="title_slide",
                action_title="Board update confirms progress",
                key_points=["Prepared for quarterly review"],
            ),
            SlidePlan(
                slide_number=2,
                story_role="closing",
                archetype_id="end_slide",
                action_title="Next steps are agreed",
            ),
        ],
    )
    pres = Presentation.create()
    report = pres.apply_operations(compile_plan_to_operations(plan))
    assert report.ok

    lint_report = lint_design(deck_index=pres.inspect(), profile=DesignProfile())
    issue_codes = {issue["code"] for issue in lint_report["issues"]}
    assert "MISSING_VISUAL_ELEMENT" not in issue_codes
    assert "SHAPE_OUT_OF_BOUNDS" not in issue_codes
    assert "SHAPE_OVERLAP" not in issue_codes
