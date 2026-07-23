"""
FastAPI-приложение: загрузка PDF net rates -> предпросмотр -> заполненный .xlsx.

Этап 1 (MVP):
  * загрузка PDF (текстового);
  * фиксированный пустой Excel-шаблон (template/);
  * извлечение net rates, предпросмотр найденных комнат/дат/цен;
  * генерация заполненного .xlsx;
  * отчёт: что найдено, что не сопоставилось.
"""
from __future__ import annotations

import base64
import tempfile
import os
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.middleware.cors import CORSMiddleware

from .parser import parse_pdf
from .filler import fill_template
from .mapping import PERIOD_COLUMNS

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "template" / "Hyatt Ziva Cap Cana.xlsx"
# Фронт статичный (docs/ для GitHub Pages); бэкенд тоже отдаёт его для локальной разработки.
INDEX_HTML = BASE_DIR / "docs" / "index.html"

app = FastAPI(title="Net Rates -> Excel", version="0.1.0")

# CORS: фронт на GitHub Pages обращается к бэку на Railway (разные origin).
# По умолчанию разрешаем только Pages-origin владельца; переопределяется env ALLOWED_ORIGINS
# (список через запятую, либо "*" чтобы открыть всем). Локальный фронт ходит same-origin — CORS не нужен.
_origins_env = os.environ.get("ALLOWED_ORIGINS", "https://alexeykos02.github.io").strip()
_allow_origins = ["*"] if _origins_env == "*" else [o.strip() for o in _origins_env.split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_allow_origins,
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return INDEX_HTML.read_text(encoding="utf-8")


@app.get("/config.js")
def config_js() -> Response:
    # Локально фронт всё равно ходит на тот же origin; отдаём файл, чтобы не было 404.
    path = BASE_DIR / "docs" / "config.js"
    body = path.read_text(encoding="utf-8") if path.exists() else "window.API_BASE='';"
    return Response(content=body, media_type="application/javascript")


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok", "template_exists": TEMPLATE_PATH.exists()}


@app.post("/api/process")
async def process(file: UploadFile = File(...)) -> JSONResponse:
    """
    Принять PDF, вернуть предпросмотр + отчёт + base64 заполненного .xlsx.
    """
    if not TEMPLATE_PATH.exists():
        raise HTTPException(500, "Шаблон не найден на сервере.")
    if not (file.filename or "").lower().endswith(".pdf"):
        raise HTTPException(400, "Ожидается PDF-файл.")

    raw = await file.read()
    if not raw:
        raise HTTPException(400, "Пустой файл.")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(raw)
            tmp_path = tmp.name

        try:
            categories, _lines = parse_pdf(tmp_path)
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(422, f"Не удалось прочитать PDF: {exc}")

        xlsx_bytes, report = fill_template(str(TEMPLATE_PATH), categories)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)

    period_cols = PERIOD_COLUMNS  # для сопоставления заголовков в UI
    rooms_payload = []
    for rr in report.rooms:
        rooms_payload.append({
            "code": rr.code,
            "name": rr.name,
            "matched": rr.matched,
            "note": rr.note,
            "written": rr.written,  # field -> список из 6 значений
        })

    # Берём даты периодов из первой найденной категории (для шапки таблицы предпросмотра).
    periods = []
    for cat in categories:
        if cat.period_starts and cat.period_ends:
            periods = [{"from": s, "to": e}
                       for s, e in zip(cat.period_starts, cat.period_ends)]
            break

    out_name = _output_name(file.filename)
    return JSONResponse({
        "summary": {
            "matched": report.matched_count,
            "total_template_rooms": len(report.rooms),
            "pdf_categories": len(categories),
            "empty_template_codes": report.empty_template_codes,
            "unmatched_pdf_codes": report.unmatched_pdf_codes,
            "warnings": report.warnings,
        },
        "periods": periods,
        "rooms": rooms_payload,
        "xlsx_base64": base64.b64encode(xlsx_bytes).decode("ascii"),
        "output_filename": out_name,
    })


def _output_name(pdf_name: str | None) -> str:
    stem = Path(pdf_name or "result").stem
    return f"{stem} — filled.xlsx"
