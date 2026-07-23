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

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.middleware.cors import CORSMiddleware

from . import parser as parser_mod
from .filler import fill_template
from .profiles import PROFILES, detect_profile, get_profile

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = BASE_DIR / "template"
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
    return {"status": "ok", "profiles": list(PROFILES)}


@app.get("/api/profiles")
def profiles_list() -> dict:
    """Список доступных форматов (для выпадашки на фронте)."""
    return {"profiles": [
        {"id": p.id, "name": p.name,
         "template_exists": (TEMPLATE_DIR / p.template).exists()}
        for p in PROFILES.values()
    ]}


@app.post("/api/process")
async def process(file: UploadFile = File(...),
                  profile: str | None = Form(None)) -> JSONResponse:
    """
    Принять PDF (+ опц. id формата), вернуть предпросмотр + отчёт + base64 .xlsx.
    Если формат не указан — определяется автоматически по содержимому PDF.
    """
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
            lines = parser_mod.pdf_lines(tmp_path)
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(422, f"Не удалось прочитать PDF: {exc}")

        # Выбор формата: явный из формы, иначе авто-определение.
        detected = detect_profile(lines)
        profile_id = profile or detected
        prof = get_profile(profile_id) if profile_id else None
        if prof is None:
            raise HTTPException(
                422,
                "Не удалось определить формат PDF. Выберите формат вручную."
                if not profile else f"Неизвестный формат: {profile}",
            )

        template_path = TEMPLATE_DIR / prof.template
        if not template_path.exists():
            raise HTTPException(500, f"Шаблон не найден: {prof.template}")

        categories, _ = prof.parse(tmp_path)
        xlsx_bytes, report = fill_template(str(template_path), categories, prof)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)

    rooms_payload = [{
        "code": rr.code, "name": rr.name, "matched": rr.matched, "note": rr.note,
        "written": rr.written, "confidence": rr.confidence,
        "level": rr.level, "flags": rr.flags,
    } for rr in report.rooms]

    # Периоды для шапки предпросмотра: from/to (labeled_blocks) или диапазон (rate_matrix).
    periods = []
    for cat in categories:
        if cat.period_starts:
            ends = cat.period_ends or [""] * len(cat.period_starts)
            periods = [{"from": s, "to": e} for s, e in zip(cat.period_starts, ends)]
            break

    return JSONResponse({
        "profile": {"id": prof.id, "name": prof.name, "auto_detected": profile is None},
        "summary": {
            "matched": report.matched_count,
            "total_template_rooms": len(report.rooms),
            "pdf_categories": len(categories),
            "empty_template_codes": report.empty_template_codes,
            "unmatched_pdf_codes": report.unmatched_pdf_codes,
            "warnings": report.warnings,
            "confidence_counts": {
                "high": sum(1 for r in report.rooms if r.matched and r.level == "high"),
                "medium": sum(1 for r in report.rooms if r.matched and r.level == "medium"),
                "low": sum(1 for r in report.rooms if r.matched and r.level == "low"),
            },
        },
        "periods": periods,
        "rooms": rooms_payload,
        "xlsx_base64": base64.b64encode(xlsx_bytes).decode("ascii"),
        "output_filename": _output_name(file.filename),
    })


def _output_name(pdf_name: str | None) -> str:
    stem = Path(pdf_name or "result").stem
    return f"{stem} — filled.xlsx"
