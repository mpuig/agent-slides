from __future__ import annotations

import io
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from slides_cli import Presentation
from slides_cli.agentic import DesignProfile, verify_assets
from slides_cli.validator import validate_package_bytes


def test_validation_ok_on_new_presentation() -> None:
    pres = Presentation.create()
    report = pres.validate()
    assert report.ok
    assert all(issue.severity != "error" for issue in report.issues)
    assert any(issue.code == "NO_SLIDES" for issue in report.issues)


def test_deep_validation_on_valid_package() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    report = pres.validate(deep=True)
    assert report.ok
    assert all(issue.code != "MISSING_PACKAGE_PART" for issue in report.issues)


def _rewrite_zip_entry(data: bytes, name: str, new_bytes: bytes) -> bytes:
    src = io.BytesIO(data)
    out = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            payload = new_bytes if item.filename == name else zin.read(item.filename)
            zout.writestr(item, payload)
    return out.getvalue()


def _append_zip_entry(data: bytes, name: str, payload: bytes) -> bytes:
    src = io.BytesIO(data)
    out = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr(name, payload)
    return out.getvalue()


def test_deep_validation_detects_broken_relationship_target() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    data = pres.to_bytes(deterministic=True)

    rels_name = "ppt/slides/_rels/slide1.xml.rels"
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        original = zf.read(rels_name)
    needle = b"../slideLayouts/"
    pivot = original.find(needle)
    assert pivot != -1
    end = original.find(b".xml", pivot)
    assert end != -1
    broken = original[:pivot] + b"../slideLayouts/does-not-exist" + original[end:]
    modified = _rewrite_zip_entry(data, rels_name, broken)

    report = validate_package_bytes(modified)
    assert any(issue.code == "BROKEN_REL_TARGET" for issue in report.issues)


def test_deep_validation_detects_invalid_relationship_xml() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    data = pres.to_bytes(deterministic=True)
    rels_name = "ppt/slides/_rels/slide1.xml.rels"
    modified = _rewrite_zip_entry(data, rels_name, b"<Relationships>")

    report = validate_package_bytes(modified)
    assert any(issue.code == "INVALID_RELS_XML" for issue in report.issues)


def test_deep_validation_detects_missing_relationship_type() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    data = pres.to_bytes(deterministic=True)
    rels_name = "ppt/slides/_rels/slide1.xml.rels"
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        root = ET.fromstring(zf.read(rels_name))

    rel_node = root.find(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    )
    assert rel_node is not None
    rel_node.attrib.pop("Type", None)
    modified = _rewrite_zip_entry(data, rels_name, ET.tostring(root, encoding="utf-8"))
    report = validate_package_bytes(modified)
    assert any(issue.code == "MISSING_REL_TYPE" for issue in report.issues)


def test_deep_validation_detects_missing_content_type_override() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    data = pres.to_bytes(deterministic=True)
    ct_name = "[Content_Types].xml"
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        root = ET.fromstring(zf.read(ct_name))

    target = "/ppt/presentation.xml"
    overrides = root.findall(
        "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
    )
    for node in overrides:
        if node.attrib.get("PartName") == target:
            root.remove(node)
            break
    modified = _rewrite_zip_entry(data, ct_name, ET.tostring(root, encoding="utf-8"))
    report = validate_package_bytes(modified)
    assert any(issue.code == "MISSING_CONTENT_TYPE_OVERRIDE" for issue in report.issues)


def test_deep_validation_require_xsd_without_dir_fails() -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    report = pres.validate(deep=True, require_xsd=True)
    assert not report.ok
    assert any(issue.code == "XSD_DIR_NOT_CONFIGURED" for issue in report.issues)


def test_xsd_validation_error_with_custom_schema(tmp_path: Path) -> None:
    xsd = """<?xml version="1.0" encoding="UTF-8"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"
    targetNamespace="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    elementFormDefault="qualified">
  <xs:element name="notPresentation" type="xs:string"/>
</xs:schema>
"""
    schema_dir = tmp_path / "xsd"
    schema_dir.mkdir(parents=True, exist_ok=True)
    (schema_dir / "pml-presentation.xsd").write_text(xsd, encoding="utf-8")

    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    report = pres.validate(deep=True, xsd_dir=schema_dir, require_xsd=True)
    assert not report.ok
    assert any(issue.code == "XSD_VALIDATION_ERROR" for issue in report.issues)
    msg = next(i.message for i in report.issues if i.code == "XSD_VALIDATION_ERROR")
    assert "line=" in msg and "column=" in msg and "type=" in msg


def test_xsd_validation_reports_unrouted_xml_part(tmp_path: Path) -> None:
    pres = Presentation.create()
    pres.add_slide(layout_index=6)
    data = pres.to_bytes(deterministic=True)
    data = _append_zip_entry(data, "ppt/custom/customData.xml", b"<x/>")
    schema_dir = tmp_path / "xsd"
    schema_dir.mkdir(parents=True, exist_ok=True)

    report = validate_package_bytes(data, xsd_dir=schema_dir, require_xsd=False)
    assert any(issue.code == "XSD_PART_UNROUTED" for issue in report.issues)


def test_verify_assets_does_not_treat_pptx_inputs_as_images(tmp_path: Path) -> None:
    input_path = tmp_path / "deck.pptx"
    template_path = tmp_path / "template.pptx"
    input_path.write_bytes(b"deck")
    template_path.write_bytes(b"template")

    report = verify_assets(
        profile=DesignProfile(asset_roots=[str(tmp_path)]),
        input_path=input_path,
        template_path=template_path,
    )
    codes = {issue["code"] for issue in report["issues"]}
    assert "ASSET_EXTENSION_BLOCKED" not in codes
