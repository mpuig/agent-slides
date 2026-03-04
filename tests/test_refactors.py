"""Tests for refactored code: dispatch dict, _resolve_chart_type, ConfigDict, pagination."""

from __future__ import annotations

import json

import pytest
from pydantic import ValidationError

from slides_cli.api import Presentation
from slides_cli.cli import _emit_paginated_results
from slides_cli.errors import SlidesError
from slides_cli.model import (
    AddNotesOp,
    AddSlideOp,
    AddTextOp,
    DeleteSlideOp,
    OperationBatch,
    OperationEvent,
    OperationReport,
    SetCorePropertiesOp,
)

# ---------------------------------------------------------------------------
# ConfigDict(extra="forbid") on OperationEvent and OperationReport
# ---------------------------------------------------------------------------


class TestConfigDictEnforcement:
    def test_operation_event_rejects_extra_fields(self) -> None:
        with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
            OperationEvent(index=0, op="add_slide", status="applied", extra_field="bad")  # type: ignore[unknown-argument]

    def test_operation_report_rejects_extra_fields(self) -> None:
        with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
            OperationReport(ok=True, dry_run=False, events=[], sneaky="nope")  # type: ignore[unknown-argument]

    def test_operation_event_valid(self) -> None:
        evt = OperationEvent(index=0, op="add_slide", status="applied")
        assert evt.index == 0
        assert evt.status == "applied"
        assert evt.detail is None
        assert evt.duration_ms is None

    def test_operation_report_valid(self) -> None:
        evt = OperationEvent(index=0, op="add_text", status="applied")
        rpt = OperationReport(ok=True, dry_run=False, events=[evt])
        assert rpt.ok is True
        assert rpt.applied_count == 0
        assert rpt.failed_index is None


# ---------------------------------------------------------------------------
# _resolve_chart_type static method
# ---------------------------------------------------------------------------


class TestResolveChartType:
    def test_valid_style_returns_chart_type(self) -> None:
        result = Presentation._resolve_chart_type(
            "clustered", {"clustered": 99, "stacked": 100}, "bar"
        )
        assert result == 99

    def test_invalid_style_raises(self) -> None:
        with pytest.raises(SlidesError) as exc_info:
            Presentation._resolve_chart_type(
                "invalid", {"clustered": 99}, "bar"
            )
        assert exc_info.value.code == "INVALID_BAR_STYLE"
        assert "clustered" in str(exc_info.value.suggested_fix)

    def test_error_code_reflects_chart_kind(self) -> None:
        with pytest.raises(SlidesError) as exc_info:
            Presentation._resolve_chart_type(
                "bad", {"good": 1}, "scatter"
            )
        assert exc_info.value.code == "INVALID_SCATTER_STYLE"
        assert exc_info.value.path is not None
        assert "add_scatter_chart.style" in exc_info.value.path


# ---------------------------------------------------------------------------
# _OP_DISPATCH table completeness and _apply_op dispatch
# ---------------------------------------------------------------------------


class TestOpDispatch:
    def test_dispatch_table_covers_all_op_types(self) -> None:
        """Every Op type in the OperationBatch discriminator should be in _OP_DISPATCH."""
        from slides_cli import model

        # Collect all Op classes that have a Literal["..."] op field
        op_types = []
        for name in dir(model):
            cls = getattr(model, name)
            if (
                isinstance(cls, type)
                and issubclass(cls, model.BaseModel)
                and "op" in getattr(cls, "model_fields", {})
                and name.endswith("Op")
            ):
                op_types.append(cls)

        dispatch_keys = set(Presentation._OP_DISPATCH.keys())
        for op_cls in op_types:
            assert op_cls in dispatch_keys, (
                f"{op_cls.__name__} is missing from _OP_DISPATCH"
            )

    def test_dispatch_methods_exist(self) -> None:
        """Every method name in the dispatch table must exist on Presentation."""
        for op_type, method_name in Presentation._OP_DISPATCH.items():
            assert hasattr(Presentation, method_name), (
                f"_OP_DISPATCH[{op_type.__name__}] references missing method '{method_name}'"
            )

    def test_apply_op_add_slide(self) -> None:
        pres = Presentation.create()
        op = AddSlideOp(op="add_slide", layout_index=6)
        pres._apply_op(op)
        assert len(pres._prs.slides) == 1

    def test_apply_op_add_text(self) -> None:
        pres = Presentation.create()
        pres.add_slide(layout_index=6)
        op = AddTextOp(
            op="add_text", slide_index=0, text="Hello", left=1.0, top=1.0,
            width=5.0, height=1.0, font_size=14,
        )
        pres._apply_op(op)
        shapes = list(pres._prs.slides[0].shapes)
        assert any("Hello" in (getattr(s, "text", "") or "") for s in shapes)

    def test_apply_op_unknown_type_raises(self) -> None:
        pres = Presentation.create()
        # Create a mock object that isn't a recognized op type
        from pydantic import BaseModel

        class FakeOp(BaseModel):
            op: str = "fake"

        with pytest.raises(SlidesError, match="UNKNOWN_OPERATION"):
            pres._apply_op(FakeOp())  # type: ignore[invalid-argument-type]

    def test_apply_operations_uses_dispatch(self) -> None:
        """End-to-end: apply_operations should use the dispatch table."""
        pres = Presentation.create()
        batch = OperationBatch(
            operations=[
                AddSlideOp(op="add_slide", layout_index=6),
                AddTextOp(
                    op="add_text", slide_index=0, text="Test", left=1.0, top=1.0,
                    width=5.0, height=1.0,
                ),
                AddNotesOp(op="add_notes", slide_index=0, text="Speaker notes"),
            ]
        )
        report = pres.apply_operations(batch)
        assert report.ok is True
        assert report.applied_count == 3

    def test_apply_op_set_core_properties(self) -> None:
        pres = Presentation.create()
        op = SetCorePropertiesOp(op="set_core_properties", title="My Deck", author="Test")
        pres._apply_op(op)
        assert pres._prs.core_properties.title == "My Deck"

    def test_apply_op_delete_slide(self) -> None:
        pres = Presentation.create()
        pres.add_slide(layout_index=6)
        pres.add_slide(layout_index=6)
        assert len(pres._prs.slides) == 2
        op = DeleteSlideOp(op="delete_slide", slide_index=0)
        pres._apply_op(op)
        assert len(pres._prs.slides) == 1


# ---------------------------------------------------------------------------
# Chart method deduplication — ensure chart methods still work
# ---------------------------------------------------------------------------


class TestChartMethodsAfterRefactor:
    @pytest.fixture()
    def pres_with_slide(self) -> Presentation:
        pres = Presentation.create()
        pres.add_slide(layout_index=6)
        return pres

    def test_add_bar_chart_clustered(self, pres_with_slide: Presentation) -> None:
        pres_with_slide.add_bar_chart(
            slide_index=0, categories=["A", "B"],
            series=[("S1", [1.0, 2.0])],
            left=1.0, top=1.0, width=5.0, height=3.0,
        )
        charts = [
            s for s in pres_with_slide._prs.slides[0].shapes
            if getattr(s, "has_chart", False)
        ]
        assert len(charts) == 1

    def test_add_bar_chart_horizontal(self, pres_with_slide: Presentation) -> None:
        pres_with_slide.add_bar_chart(
            slide_index=0, categories=["A", "B"],
            series=[("S1", [1.0, 2.0])],
            orientation="bar",
            left=1.0, top=1.0, width=5.0, height=3.0,
        )
        charts = [
            s for s in pres_with_slide._prs.slides[0].shapes
            if getattr(s, "has_chart", False)
        ]
        assert len(charts) == 1

    def test_add_bar_chart_invalid_orientation(self, pres_with_slide: Presentation) -> None:
        with pytest.raises(SlidesError, match="INVALID_BAR_ORIENTATION"):
            pres_with_slide.add_bar_chart(
                slide_index=0, categories=["A"],
                series=[("S1", [1.0])],
                orientation="vertical",
                left=1.0, top=1.0, width=5.0, height=3.0,
            )

    def test_add_bar_chart_invalid_style(self, pres_with_slide: Presentation) -> None:
        with pytest.raises(SlidesError, match="INVALID_BAR_STYLE"):
            pres_with_slide.add_bar_chart(
                slide_index=0, categories=["A"],
                series=[("S1", [1.0])],
                style="invalid",
                left=1.0, top=1.0, width=5.0, height=3.0,
            )

    def test_add_line_chart(self, pres_with_slide: Presentation) -> None:
        pres_with_slide.add_line_chart(
            slide_index=0, categories=["Q1", "Q2"],
            series=[("Sales", [10.0, 20.0])],
            left=1.0, top=1.0, width=5.0, height=3.0,
        )
        charts = [
            s for s in pres_with_slide._prs.slides[0].shapes
            if getattr(s, "has_chart", False)
        ]
        assert len(charts) == 1

    def test_add_pie_chart(self, pres_with_slide: Presentation) -> None:
        pres_with_slide.add_pie_chart(
            slide_index=0, categories=["A", "B", "C"],
            series=[("Share", [30.0, 50.0, 20.0])],
            left=1.0, top=1.0, width=5.0, height=3.0,
        )
        charts = [
            s for s in pres_with_slide._prs.slides[0].shapes
            if getattr(s, "has_chart", False)
        ]
        assert len(charts) == 1

    def test_add_line_chart_invalid_style(self, pres_with_slide: Presentation) -> None:
        with pytest.raises(SlidesError, match="INVALID_LINE_STYLE"):
            pres_with_slide.add_line_chart(
                slide_index=0, categories=["A"],
                series=[("S1", [1.0])],
                style="bad",
                left=1.0, top=1.0, width=5.0, height=3.0,
            )


# ---------------------------------------------------------------------------
# _emit_paginated_results deduplication
# ---------------------------------------------------------------------------


class TestEmitPaginatedResults:
    @pytest.fixture()
    def items(self) -> list[dict]:
        return [{"id": i, "value": f"item-{i}"} for i in range(5)]

    def test_single_page_no_pagination(
        self, items: list[dict], capsys: pytest.CaptureFixture,
    ) -> None:
        _emit_paginated_results(
            items,
            envelope_key="results",
            summary_fields={"query": "test"},
            payload_name="find",
            out_path=None,
            page_all=False,
            page_size=None,
            page_token=None,
            fields=None,
            ndjson=False,
            quiet=False,
            compact=False,
        )
        output = json.loads(capsys.readouterr().out)
        assert output["query"] == "test"
        assert len(output["results"]) == 5
        assert output["next_page_token"] is None

    def test_single_page_with_size(self, items: list[dict], capsys: pytest.CaptureFixture) -> None:
        _emit_paginated_results(
            items,
            envelope_key="results",
            summary_fields={"query": "test"},
            payload_name="find",
            out_path=None,
            page_all=False,
            page_size=3,
            page_token=None,
            fields=None,
            ndjson=False,
            quiet=False,
            compact=False,
        )
        output = json.loads(capsys.readouterr().out)
        assert len(output["results"]) == 3
        assert output["next_page_token"] == "3"

    def test_page_all_ndjson(self, items: list[dict], capsys: pytest.CaptureFixture) -> None:
        _emit_paginated_results(
            items,
            envelope_key="results",
            summary_fields={"query": "test"},
            payload_name="find",
            out_path=None,
            page_all=True,
            page_size=2,
            page_token=None,
            fields=None,
            ndjson=True,
            quiet=False,
            compact=False,
        )
        lines = capsys.readouterr().out.strip().split("\n")
        assert len(lines) == 3  # 5 items / 2 per page = 3 pages
        first_page = json.loads(lines[0])
        assert len(first_page["results"]) == 2
        assert first_page["query"] == "test"
        last_page = json.loads(lines[-1])
        assert last_page["next_page_token"] is None

    def test_page_all_wraps_in_pages_key(
        self, items: list[dict], capsys: pytest.CaptureFixture,
    ) -> None:
        _emit_paginated_results(
            items,
            envelope_key="slides",
            summary_fields={"summary": {"count": 5}},
            payload_name="inspect",
            out_path=None,
            page_all=True,
            page_size=3,
            page_token=None,
            fields=None,
            ndjson=False,
            quiet=False,
            compact=False,
        )
        output = json.loads(capsys.readouterr().out)
        assert "pages" in output
        assert len(output["pages"]) == 2
        assert output["pages"][0]["summary"] == {"count": 5}

    def test_writes_to_file(
        self, items: list[dict], tmp_path: object,
    ) -> None:
        from pathlib import Path

        out = Path(str(tmp_path)) / "out.json"
        _emit_paginated_results(
            items,
            envelope_key="results",
            summary_fields={},
            payload_name="test",
            out_path=out,
            page_all=False,
            page_size=None,
            page_token=None,
            fields=None,
            ndjson=False,
            quiet=True,
            compact=False,
        )
        assert out.exists()
        data = json.loads(out.read_text())
        assert len(data["results"]) == 5
