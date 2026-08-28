from io import BytesIO
from decimal import Decimal
from pathlib import Path
import re

import pytest
from openpyxl import load_workbook
from pydantic import ValidationError

from app.main import METER_KEYS, ReadingPayload, build_workbook, month_bounds, validate_settings


def sample_record(day: int, factor: int = 1) -> dict:
    return {
        "date": f"2026-08-{day:02d}",
        "time": "08:00",
        "reader": "测试人员",
        "values": {key: Decimal(index * factor) for index, key in enumerate(METER_KEYS, 1)},
    }


def test_payload_requires_all_nonnegative_meter_values():
    payload = ReadingPayload(reader=" 测试人员 ", values=sample_record(1)["values"])
    assert payload.reader == "测试人员"
    assert len(payload.values) == 18
    with pytest.raises(ValidationError):
        ReadingPayload(reader="测试人员", values={"boka": Decimal("1")})
    invalid = sample_record(1)["values"]
    invalid["boka"] = Decimal("-1")
    with pytest.raises(ValidationError):
        ReadingPayload(reader="测试人员", values=invalid)


def test_settings_require_distinct_secrets(monkeypatch):
    monkeypatch.setenv("DATABASE_URL", "postgresql://example")
    monkeypatch.setenv("SUBMISSION_CODE", "same-secret")
    monkeypatch.setenv("ADMIN_PASSWORD", "same-secret")
    with pytest.raises(RuntimeError, match="must be different"):
        validate_settings()


def test_frontend_contains_every_backend_meter():
    html = (Path(__file__).parents[1] / "public" / "index.html").read_text(encoding="utf-8")
    frontend_keys = tuple(re.findall(r'\{ id:"([^"]+)", group:', html))
    assert frontend_keys == METER_KEYS


def test_month_bounds_handles_december():
    assert month_bounds("2026-08") == ("2026-08-01", "2026-09-01")
    assert month_bounds("2026-12") == ("2026-12-01", "2027-01-01")


def test_export_matches_template_and_keeps_formula_results():
    content = build_workbook([sample_record(1), sample_record(2, 2)])
    formulas = load_workbook(BytesIO(content), data_only=False)["每日抄表"]
    values = load_workbook(BytesIO(content), data_only=True)["每日抄表"]

    assert formulas.max_column == 20
    assert formulas["A1"].value == "每日水电用量抄表记录"
    assert formulas["D3"].value == "一号并网柜"
    assert formulas["B26"].value == "=SUM(B4:B25)*3000"
    assert formulas["D26"].value == "=SUM(D4:D25)*6000"
    assert formulas["K26"].value == "=SUM(K4:K25)"
    assert values["B26"].value == 9000
    assert values["D26"].value == 54000
    assert values["K26"].value == 30
    assert formulas["A1"].font.name == "Noto IKEA Simplified Chinese"
    assert {str(item) for item in formulas.merged_cells.ranges} == {
        "A1:S1", "A2:A3", "B2:C2", "D2:G2", "H2:J2", "K2:P2", "Q2:S2", "T2:T3"
    }
