from __future__ import annotations

import hmac
import io
import os
import re
from contextlib import asynccontextmanager
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Annotated
from urllib.parse import quote, urlsplit
from zoneinfo import ZoneInfo

import psycopg
import qrcode
import xlsxwriter
from fastapi import Depends, FastAPI, Form, Header, HTTPException, Query, Request, Response, status
from fastapi.staticfiles import StaticFiles
from psycopg.rows import dict_row
from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

SITE = "STO401"
TIME_ZONE = ZoneInfo("Asia/Shanghai")
MONTH_PATTERN = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
ROOT = Path(__file__).resolve().parent.parent

METERS = (
    ("boka", "博卡线", 3000),
    ("huaman", "花漫线", 3000),
    ("solar1", "一号并网柜", 6000),
    ("solar2", "二号并网柜", 6000),
    ("solar3", "三号并网柜", 6000),
    ("solar4", "四号并网柜", 6000),
    ("charger1", "1AA9-7", 1),
    ("charger2", "3AA4-6", 1),
    ("charger3", "4AA4-5", 1),
    ("domestic_water", "生活水表", None),
    ("domestic_water_electronic", "生活水表  \n（电子表）", None),
    ("fire_east", "消防东表", None),
    ("fire_east_electronic", "消防东表\n   （电子表）", None),
    ("fire_west", "消防西表", None),
    ("fire_west_electronic", "消防西表 \n （电子表）", None),
    ("reclaimed_total", "总水", None),
    ("reclaimed_in", "进水", None),
    ("reclaimed_out", "出水", None),
)
METER_KEYS = tuple(item[0] for item in METERS)
MAX_READING = Decimal("1000000000000000")


class ReadingPayload(BaseModel):
    model_config = ConfigDict(extra="forbid")

    reader: str = Field(min_length=1, max_length=40)
    values: dict[str, Decimal]

    @field_validator("reader")
    @classmethod
    def clean_reader(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("抄表人不能为空")
        return value

    @model_validator(mode="after")
    def validate_values(self) -> "ReadingPayload":
        supplied = set(self.values)
        expected = set(METER_KEYS)
        if supplied != expected:
            missing = sorted(expected - supplied)
            extra = sorted(supplied - expected)
            raise ValueError(f"点位不完整；缺少 {missing}，多余 {extra}")
        for key, value in self.values.items():
            if not value.is_finite() or value < 0 or value > MAX_READING:
                raise ValueError(f"{key} 的读数无效")
        return self


def get_database_url() -> str:
    value = os.getenv("DATABASE_URL", "").strip()
    if not value:
        raise RuntimeError("DATABASE_URL is required")
    return value


def validate_settings() -> None:
    get_database_url()
    submission_code = os.getenv("SUBMISSION_CODE", "")
    admin_password = os.getenv("ADMIN_PASSWORD", "")
    if len(submission_code) < 8:
        raise RuntimeError("SUBMISSION_CODE must contain at least 8 characters")
    if len(admin_password) < 8:
        raise RuntimeError("ADMIN_PASSWORD must contain at least 8 characters")
    if hmac.compare_digest(submission_code.encode(), admin_password.encode()):
        raise RuntimeError("SUBMISSION_CODE and ADMIN_PASSWORD must be different")


def verify_secret(actual: str | None, variable: str) -> None:
    expected = os.getenv(variable, "")
    if len(expected) < 8:
        raise RuntimeError(f"{variable} must contain at least 8 characters")
    if actual is None or not hmac.compare_digest(actual.encode(), expected.encode()):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="验证码或密码不正确")


def require_admin(x_admin_password: Annotated[str | None, Header()] = None) -> None:
    verify_secret(x_admin_password, "ADMIN_PASSWORD")


def month_bounds(month: str) -> tuple[str, str]:
    if not MONTH_PATTERN.fullmatch(month):
        raise HTTPException(status_code=422, detail="月份格式必须为 YYYY-MM")
    year, value = map(int, month.split("-"))
    next_month = f"{year + (value == 12):04d}-{1 if value == 12 else value + 1:02d}-01"
    return f"{month}-01", next_month


def connect() -> psycopg.Connection:
    return psycopg.connect(get_database_url(), connect_timeout=10, row_factory=dict_row)


def initialize_database() -> None:
    columns = ",\n".join(
        f"{key} NUMERIC(22, 6) NOT NULL CHECK ({key} >= 0 AND {key} <= {MAX_READING})"
        for key in METER_KEYS
    )
    sql = f"""
        CREATE TABLE IF NOT EXISTS meter_readings (
            id BIGSERIAL PRIMARY KEY,
            site VARCHAR(20) NOT NULL,
            reading_date DATE NOT NULL,
            recorded_at TIMESTAMPTZ NOT NULL,
            reader VARCHAR(40) NOT NULL,
            {columns},
            created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            UNIQUE (site, reading_date)
        );
        CREATE INDEX IF NOT EXISTS meter_readings_month_idx
        ON meter_readings (site, reading_date);
    """
    with connect() as connection:
        connection.execute(sql)


def fetch_month(month: str) -> list[dict]:
    start, end = month_bounds(month)
    columns = ", ".join(METER_KEYS)
    with connect() as connection:
        rows = connection.execute(
            f"""
            SELECT reading_date, recorded_at, reader, {columns}
            FROM meter_readings
            WHERE site = %s AND reading_date >= %s AND reading_date < %s
            ORDER BY reading_date
            """,
            (SITE, start, end),
        ).fetchall()
    result = []
    for row in rows:
        values = {key: row[key] for key in METER_KEYS}
        result.append(
            {
                "date": row["reading_date"].isoformat(),
                "time": row["recorded_at"].astimezone(TIME_ZONE).strftime("%H:%M"),
                "reader": row["reader"],
                "values": values,
            }
        )
    return result


def build_workbook(records: list[dict]) -> bytes:
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    workbook.set_properties({"title": "每日水电用量抄表记录", "author": "杭州商场"})
    sheet = workbook.add_worksheet("每日抄表")
    sheet.hide_gridlines(2)
    sheet.freeze_panes(21, 0)
    sheet.set_landscape()
    sheet.set_paper(9)
    sheet.fit_to_pages(1, 0)
    sheet.set_column("A:A", 9.08203125)
    sheet.set_column("B:S", 14.4140625)
    sheet.set_column("T:T", 15.75)

    title = workbook.add_format(
        {"font_name": "Noto IKEA Simplified Chinese", "font_size": 11, "bold": True, "align": "center", "valign": "vcenter"}
    )
    header = workbook.add_format(
        {"font_name": "Noto IKEA Simplified Chinese", "font_size": 10, "bold": True, "align": "center", "valign": "vcenter", "text_wrap": True}
    )
    data = workbook.add_format(
        {"font_name": "Noto IKEA Simplified Chinese", "font_size": 11, "align": "center", "valign": "vcenter"}
    )
    total = workbook.add_format(
        {"font_name": "Noto IKEA Simplified Chinese", "font_size": 11, "align": "center", "valign": "vcenter"}
    )

    sheet.set_row(0, 33.5)
    sheet.set_row(1, 26.25)
    sheet.set_row(2, 36)
    sheet.merge_range("A1:S1", "每日水电用量抄表记录", title)
    sheet.merge_range("A2:A3", "日期", header)
    sheet.merge_range("B2:C2", "高压进线", header)
    sheet.merge_range("D2:G2", "光伏抄表", header)
    sheet.merge_range("H2:J2", "电动车充电桩", header)
    sheet.merge_range("K2:P2", "用水量(t)", header)
    sheet.merge_range("Q2:S2", "中水", header)
    sheet.merge_range("T2:T3", "抄表人", header)
    for column, (_, label, _) in enumerate(METERS, 1):
        sheet.write(2, column, label, header)

    for row_index, record in enumerate(records, 3):
        sheet.set_row(row_index, 27.75)
        sheet.write(row_index, 0, record["date"], data)
        for column, (key, _, multiplier) in enumerate(METERS, 1):
            reading = Decimal(str(record["values"][key]))
            if multiplier is not None:
                reading *= multiplier
            sheet.write_number(row_index, column, float(reading), data)
        sheet.write(row_index, 19, record["reader"], data)

    data_end_row = max(25, len(records) + 3)
    total_row = data_end_row
    sheet.set_row(total_row, 27.75)
    sheet.write(total_row, 0, "单项合计", total)
    for column, (key, _, multiplier) in enumerate(METERS, 1):
        excel_column = xlsxwriter.utility.xl_col_to_name(column)
        formula = f"SUM({excel_column}4:{excel_column}{data_end_row})"
        value = sum(Decimal(str(record["values"][key])) for record in records)
        if multiplier is not None:
            value *= multiplier
        sheet.write_formula(total_row, column, f"={formula}", total, float(value))
    sheet.print_area(0, 0, total_row, 19)
    workbook.close()
    return output.getvalue()


@asynccontextmanager
async def lifespan(_: FastAPI):
    validate_settings()
    initialize_database()
    yield


app = FastAPI(title="STO401 每日抄表", docs_url=None, redoc_url=None, lifespan=lifespan)


@app.middleware("http")
async def security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "same-origin"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    if request.url.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/api/health")
def health() -> dict:
    with connect() as connection:
        connection.execute("SELECT 1")
    return {"status": "ok"}


@app.get("/api/config")
def config() -> dict:
    now = datetime.now(TIME_ZONE)
    return {"site": SITE, "serverTime": now.isoformat(timespec="minutes")}


@app.post("/api/readings")
def save_reading(payload: ReadingPayload, x_submission_code: Annotated[str | None, Header()] = None) -> dict:
    verify_secret(x_submission_code, "SUBMISSION_CODE")
    now = datetime.now(TIME_ZONE)
    columns = ", ".join(METER_KEYS)
    placeholders = ", ".join(["%s"] * len(METER_KEYS))
    updates = ", ".join(f"{key} = EXCLUDED.{key}" for key in METER_KEYS)
    values = [payload.values[key] for key in METER_KEYS]
    with connect() as connection:
        connection.execute(
            f"""
            INSERT INTO meter_readings
                (site, reading_date, recorded_at, reader, {columns})
            VALUES (%s, %s, %s, %s, {placeholders})
            ON CONFLICT (site, reading_date) DO UPDATE SET
                recorded_at = EXCLUDED.recorded_at,
                reader = EXCLUDED.reader,
                {updates},
                updated_at = NOW()
            """,
            (SITE, now.date(), now, payload.reader, *values),
        )
    return {"saved": True, "date": now.date().isoformat(), "time": now.strftime("%H:%M")}


@app.get("/api/readings", dependencies=[Depends(require_admin)])
def list_readings(month: Annotated[str, Query()]) -> dict:
    return {"month": month, "records": fetch_month(month)}


@app.delete("/api/readings", dependencies=[Depends(require_admin)])
def delete_readings(month: Annotated[str, Query()]) -> dict:
    start, end = month_bounds(month)
    with connect() as connection:
        cursor = connection.execute(
            "DELETE FROM meter_readings WHERE site = %s AND reading_date >= %s AND reading_date < %s",
            (SITE, start, end),
        )
        deleted = cursor.rowcount
    return {"deleted": deleted}


def build_export_response(month: str) -> Response:
    records = fetch_month(month)
    if not records:
        raise HTTPException(status_code=404, detail="所选月份没有记录")
    content = build_workbook(records)
    filename = f"杭州商场每日水电用量抄表_{month}.xlsx"
    headers = {
        "Content-Disposition": (
            f'attachment; filename="meter-reading-{month}.xlsx"; '
            f"filename*=UTF-8''{quote(filename)}"
        ),
        "Content-Length": str(len(content)),
    }
    return Response(
        content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/export", dependencies=[Depends(require_admin)])
def export_readings(month: Annotated[str, Query()]) -> Response:
    return build_export_response(month)


@app.post("/api/export")
def export_readings_from_form(
    month: Annotated[str, Form()],
    admin_password: Annotated[str, Form()],
) -> Response:
    verify_secret(admin_password, "ADMIN_PASSWORD")
    return build_export_response(month)


def resolve_public_url(request: Request) -> str:
    forwarded_host = request.headers.get("x-forwarded-host", "").split(",")[0].strip()
    host = forwarded_host or request.headers.get("host", "").strip()
    if host and re.fullmatch(r"[A-Za-z0-9.:[\]-]+", host):
        hostname = urlsplit(f"//{host}").hostname or ""
        is_local = hostname in {"localhost", "127.0.0.1", "::1"} or hostname.endswith(".local")
        scheme = request.url.scheme if is_local else "https"
        return f"{scheme}://{host}"

    configured = os.getenv("PUBLIC_URL", "").strip().rstrip("/")
    return configured or str(request.base_url).rstrip("/")


@app.get("/api/qr.png")
def qr_code(request: Request) -> Response:
    url = resolve_public_url(request)
    qr = qrcode.QRCode(
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    image = qr.make_image(fill_color="black", back_color="white")
    output = io.BytesIO()
    image.save(output, format="PNG")
    return Response(output.getvalue(), media_type="image/png", headers={"Content-Disposition": "inline; filename=meter-reading-qr.png"})


app.mount("/", StaticFiles(directory=ROOT / "public", html=True), name="static")
