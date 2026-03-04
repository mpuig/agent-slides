from __future__ import annotations

import base64
import io
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from slides_cli import Presentation

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def _read_zip_xml(data: bytes, name: str) -> ET.Element:
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        xml_bytes = zf.read(name)
    return ET.fromstring(xml_bytes)


def test_golden_presentation_parts_exist() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=0)
    data = pres.to_bytes(deterministic=True)

    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        names = set(zf.namelist())

    assert "[Content_Types].xml" in names
    assert "_rels/.rels" in names
    assert "ppt/presentation.xml" in names
    assert "ppt/slides/slide1.xml" in names


def test_golden_slide_contains_text_and_background() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_text(
        slide_index=slide.index,
        text="Golden text",
        left=0.5,
        top=0.5,
        width=4,
        height=1,
    )
    pres.set_slide_background(slide_index=slide.index, color_hex="F4F7F9")

    slide_xml = _read_zip_xml(pres.to_bytes(deterministic=True), "ppt/slides/slide1.xml")
    assert slide_xml.find(".//a:t", NS) is not None
    assert slide_xml.find(".//p:bg", NS) is not None


def test_golden_chart_xml_part_present() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_bar_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3"],
        series=[("Revenue", [1.0, 2.0, 3.0])],
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    data = pres.to_bytes(deterministic=True)

    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        names = set(zf.namelist())
    assert any(name.startswith("ppt/charts/chart") and name.endswith(".xml") for name in names)


def test_golden_chart_legend_xml_present() -> None:
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
    pres.set_chart_legend(
        slide_index=slide.index,
        chart_index=0,
        visible=True,
        position="bottom",
        include_in_layout=False,
    )

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_name = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )[0]
    chart_xml = _read_zip_xml(data, chart_name)
    assert chart_xml.find(".//c:legend", NS) is not None


def test_golden_image_crop_xml_present(tmp_path: Path) -> None:
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
    pres.set_image_crop(
        slide_index=slide.index,
        image_index=0,
        crop_left=0.1,
        crop_right=0.1,
        crop_top=0.0,
        crop_bottom=0.0,
    )

    slide_xml = _read_zip_xml(pres.to_bytes(deterministic=True), "ppt/slides/slide1.xml")
    blip_fill = slide_xml.find(".//p:pic/p:blipFill", NS)
    assert blip_fill is not None
    src_rect = blip_fill.find("a:srcRect", NS)
    assert src_rect is not None


def test_golden_chart_title_axis_and_labels_xml_present() -> None:
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
    pres.set_chart_title(slide_index=slide.index, chart_index=0, text="Revenue Trend")
    pres.set_chart_axis_titles(
        slide_index=slide.index,
        chart_index=0,
        category_title="Quarter",
        value_title="M EUR",
    )
    pres.set_chart_data_labels(
        slide_index=slide.index,
        chart_index=0,
        enabled=True,
        show_value=True,
        number_format="#,##0.00",
    )

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_name = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )[0]
    chart_xml = _read_zip_xml(data, chart_name)
    assert chart_xml.find(".//c:title", NS) is not None
    assert chart_xml.find(".//c:catAx/c:title", NS) is not None
    assert chart_xml.find(".//c:valAx/c:title", NS) is not None
    assert chart_xml.find(".//c:dLbls", NS) is not None


def test_golden_line_and_pie_chart_xml_present() -> None:
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
    pres.add_pie_chart(
        slide_index=slide.index,
        categories=["A", "B", "C"],
        series=[("Share", [40.0, 35.0, 25.0])],
        left=5.0,
        top=0.5,
        width=4.0,
        height=2.0,
    )

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_parts = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )
    assert len(chart_parts) >= 2
    chart_xml_1 = _read_zip_xml(data, chart_parts[0])
    chart_xml_2 = _read_zip_xml(data, chart_parts[1])
    tags = {chart_xml_1.tag, chart_xml_2.tag}
    assert any(xml.find(".//c:lineChart", NS) is not None for xml in (chart_xml_1, chart_xml_2))
    assert any(xml.find(".//c:pieChart", NS) is not None for xml in (chart_xml_1, chart_xml_2))
    assert tags


def test_golden_media_part_present(tmp_path: Path) -> None:
    media_path = tmp_path / "clip.mp4"
    media_path.write_bytes(b"fake-video-bytes")
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_media(
        slide_index=slide.index,
        path=media_path,
        left=1.0,
        top=1.0,
        width=2.0,
        height=2.0,
        mime_type="video/mp4",
    )

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        names = set(zf.namelist())
    assert any(name.startswith("ppt/media/") for name in names)


def test_golden_combo_overlay_and_secondary_axis_xml() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_combo_chart_overlay(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3"],
        bar_series=[("Revenue", [10.0, 12.0, 14.0])],
        line_series=[("Margin", [20.0, 22.0, 21.0])],
        left=0.5,
        top=0.5,
        width=8.0,
        height=3.0,
    )
    pres.set_chart_secondary_axis(slide_index=slide.index, chart_index=1, enable=True)
    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_parts = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )
    assert len(chart_parts) >= 2
    second = _read_zip_xml(data, chart_parts[1])
    assert second.find(".//c:valAx/c:axPos[@val='r']", NS) is not None


def test_golden_trendline_and_errorbars_xml_present() -> None:
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
    pres.set_line_series_marker(
        slide_index=slide.index, chart_index=0, series_index=0, style="diamond", size=8
    )
    pres.set_chart_series_trendline(
        slide_index=slide.index,
        chart_index=0,
        series_index=0,
        trend_type="linear",
    )
    pres.set_chart_series_error_bars(
        slide_index=slide.index,
        chart_index=0,
        series_index=0,
        value=0.2,
        direction="y",
        bar_type="both",
    )

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_name = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )[0]
    chart_xml = _read_zip_xml(data, chart_name)
    assert chart_xml.find(".//c:ser/c:marker", NS) is not None
    assert chart_xml.find(".//c:ser/c:trendline", NS) is not None
    assert chart_xml.find(".//c:ser/c:errBars", NS) is not None


def test_golden_true_combo_mapping_xml_present() -> None:
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
    pres.set_chart_combo_secondary_mapping(
        slide_index=slide.index,
        chart_index=0,
        series_indices=[1],
    )

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_name = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )[0]
    chart_xml = _read_zip_xml(data, chart_name)
    assert chart_xml.find(".//c:barChart", NS) is not None
    assert chart_xml.find(".//c:lineChart", NS) is not None
    assert chart_xml.find(".//c:lineChart/c:ser/c:idx[@val='1']", NS) is not None
    assert chart_xml.find(".//c:valAx/c:axPos[@val='r']", NS) is not None


def test_golden_true_combo_mapping_from_line_base_xml_present() -> None:
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
    pres.set_chart_combo_secondary_mapping(
        slide_index=slide.index,
        chart_index=0,
        series_indices=[1],
    )

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_name = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )[0]
    chart_xml = _read_zip_xml(data, chart_name)
    assert chart_xml.find(".//c:lineChart", NS) is not None
    assert chart_xml.find(".//c:barChart", NS) is not None
    assert chart_xml.find(".//c:barChart/c:ser/c:idx[@val='1']", NS) is not None


def test_golden_chart_style_and_smooth_xml_present() -> None:
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
    pres.set_chart_style(slide_index=slide.index, chart_index=0, style_id=4)
    pres.set_line_series_smooth(slide_index=slide.index, chart_index=0, series_index=0, smooth=True)

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_name = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )[0]
    chart_xml = _read_zip_xml(data, chart_name)
    assert chart_xml.find(".//c:style[@val='4']", NS) is not None
    assert chart_xml.find(".//c:ser/c:smooth", NS) is not None


def test_golden_area_doughnut_scatter_radar_bubble_xml_present() -> None:
    pres = Presentation.create()
    slide = pres.add_slide(layout_index=6)
    pres.add_area_chart(
        slide_index=slide.index,
        categories=["Q1", "Q2", "Q3"],
        series=[("Area", [1.0, 2.0, 3.0])],
        style="percent_stacked",
        left=0.5,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    pres.add_doughnut_chart(
        slide_index=slide.index,
        categories=["A", "B", "C"],
        series=[("Mix", [30.0, 40.0, 30.0])],
        style="exploded",
        left=5.0,
        top=0.5,
        width=4.0,
        height=2.0,
    )
    pres.add_scatter_chart(
        slide_index=slide.index,
        series=[("S1", [(1.0, 2.0), (2.0, 3.0), (3.0, 3.5)])],
        style="smooth",
        left=0.5,
        top=3.0,
        width=4.0,
        height=2.0,
    )
    pres.add_radar_chart(
        slide_index=slide.index,
        categories=["A", "B", "C", "D"],
        series=[("R", [4.0, 3.0, 5.0, 2.0])],
        style="filled",
        left=5.0,
        top=3.0,
        width=4.0,
        height=2.0,
    )
    pres.add_bubble_chart(
        slide_index=slide.index,
        series=[("B", [(1.0, 2.0, 3.0), (2.0, 2.5, 4.0)])],
        style="bubble",
        left=0.5,
        top=5.4,
        width=4.0,
        height=2.0,
    )

    data = pres.to_bytes(deterministic=True)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        chart_names = sorted(
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        )

    chart_xmls = [_read_zip_xml(data, name) for name in chart_names]
    assert any(xml.find(".//c:areaChart", NS) is not None for xml in chart_xmls)
    assert any(xml.find(".//c:doughnutChart", NS) is not None for xml in chart_xmls)
    assert any(xml.find(".//c:scatterChart", NS) is not None for xml in chart_xmls)
    assert any(xml.find(".//c:radarChart", NS) is not None for xml in chart_xmls)
    assert any(xml.find(".//c:bubbleChart", NS) is not None for xml in chart_xmls)
