"""
Заполнение фиксированного Excel-шаблона извлечёнными ставками + отчёт.

Пишем ТОЛЬКО во входные ячейки "Bed prices". Формулы комбинаций в шаблоне не трогаем —
openpyxl сохраняет их (в т.ч. массивные) при повторной записи, если ячейки не изменялись.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import openpyxl

from .mapping import (
    PERIOD_COLUMNS,
    OFFSET_ADULT_MB,
    OFFSET_ADULT_EXB,
    OFFSET_CHILD312_MB,
    OFFSET_CHILD312_EXB,
    OFFSET_PERIOD_START,
)
from .parser import Category, CODE_RE


@dataclass
class RoomReport:
    code: str
    name: str
    matched: bool
    written: dict[str, list[float]] = field(default_factory=dict)  # field -> 6 значений
    note: str = ""


@dataclass
class FillReport:
    rooms: list[RoomReport] = field(default_factory=list)
    unmatched_pdf_codes: list[str] = field(default_factory=list)  # есть в PDF, нет в шаблоне
    empty_template_codes: list[str] = field(default_factory=list)  # есть в шаблоне, нет в PDF
    warnings: list[str] = field(default_factory=list)

    @property
    def matched_count(self) -> int:
        return sum(1 for r in self.rooms if r.matched)


def _template_room_blocks(ws) -> list[tuple[int, str, str]]:
    """Список (row, full_name, code) для всех блоков комнат шаблона."""
    blocks = []
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "Room":
            name = str(ws.cell(r, 2).value or "")
            m = CODE_RE.search(name)
            code = m.group(1) if m else ""
            blocks.append((r, name, code))
    return blocks


def _write_row(ws, row: int, values: list[float]) -> None:
    for col, val in zip(PERIOD_COLUMNS, values):
        ws.cell(row=row, column=col, value=val)


def fill_template(template_path: str, categories: list[Category]) -> tuple[bytes, FillReport]:
    """
    Заполнить шаблон. Возвращает (xlsx_bytes, report).
    """
    wb = openpyxl.load_workbook(template_path, data_only=False)
    ws = wb.active

    blocks = _template_room_blocks(ws)
    code_to_block = {code: (row, name) for row, name, code in blocks if code}

    # code -> Category (King и Double делят ставки одной категории)
    code_to_cat: dict[str, Category] = {}
    all_pdf_codes: list[str] = []
    for cat in categories:
        for code in cat.codes:
            all_pdf_codes.append(code)
            code_to_cat[code] = cat

    report = FillReport()
    matched_codes: set[str] = set()

    for code, (row, name) in sorted(code_to_block.items(), key=lambda x: x[1][0]):
        cat = code_to_cat.get(code)
        if cat is None:
            report.rooms.append(RoomReport(code=code, name=name, matched=False,
                                           note="нет в PDF (не сопоставлено)"))
            report.empty_template_codes.append(code)
            continue

        matched_codes.add(code)
        rr = RoomReport(code=code, name=name, matched=True)

        # Проверка соответствия периодов по датам (по порядку) — только предупреждение.
        _check_periods(ws, row, cat, code, report)

        adult_mb = cat.rates.get("adult_mb")
        adult_exb = cat.rates.get("adult_exb")
        child_exb = cat.rates.get("child312_exb")

        if adult_mb:
            _write_row(ws, row + OFFSET_ADULT_MB, adult_mb)
            rr.written["Adult MB"] = adult_mb
        if adult_exb:
            _write_row(ws, row + OFFSET_ADULT_EXB, adult_exb)
            rr.written["Adult ExB"] = adult_exb
        # Child 3-12 заполняем только если PDF даёт эту строку.
        if child_exb:
            _write_row(ws, row + OFFSET_CHILD312_EXB, child_exb)
            rr.written["Child 3-12 ExB"] = child_exb
            if adult_mb:  # Child 3-12 MB = Adult MB (правило подтверждено на эталоне)
                _write_row(ws, row + OFFSET_CHILD312_MB, adult_mb)
                rr.written["Child 3-12 MB"] = adult_mb

        if not rr.written:
            rr.note = "категория найдена, но ставки не извлеклись"
        report.rooms.append(rr)

    # Коды из PDF, которых нет в шаблоне.
    seen = set()
    for code in all_pdf_codes:
        if code not in code_to_block and code not in seen:
            seen.add(code)
            report.unmatched_pdf_codes.append(code)

    # openpyxl не сохраняет кэш значений формул -> просим Excel пересчитать при открытии,
    # иначе комбинации покажут 0 до ручного пересчёта.
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), report


def _check_periods(ws, row: int, cat: Category, code: str, report: FillReport) -> None:
    """Мягкая сверка: количество периодов и первая дата start."""
    if len(cat.period_starts) != len(PERIOD_COLUMNS):
        report.warnings.append(
            f"{code}: в PDF {len(cat.period_starts)} периодов вместо {len(PERIOD_COLUMNS)}"
        )
