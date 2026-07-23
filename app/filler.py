"""
Заполнение фиксированного Excel-шаблона извлечёнными ставками + отчёт.

Пишем ТОЛЬКО во входные ячейки "Bed prices". Формулы комбинаций в шаблоне не трогаем —
openpyxl сохраняет их (в т.ч. массивные) при повторной записи, если ячейки не изменялись.
"""
from __future__ import annotations

import io
from datetime import date, datetime
from dataclasses import dataclass, field

import openpyxl

from .mapping import (
    PERIOD_COLUMNS,
    N_PERIODS,
    OFFSET_ADULT_MB,
    OFFSET_ADULT_EXB,
    OFFSET_CHILD312_MB,
    OFFSET_CHILD312_EXB,
    OFFSET_SINGLE_1A,
    OFFSET_PERIOD_START,
)
from .parser import Category, CODE_RE

# Множитель одиночного размещения (1A) в формуле шаблона: Single = Adult MB * 1.7.
# Используем для кросс-проверки: parsed "Single net" должен совпасть -> высокая достоверность.
SINGLE_MULTIPLIER = 1.7


@dataclass
class RoomReport:
    code: str
    name: str
    matched: bool
    written: dict[str, list[float]] = field(default_factory=dict)  # field -> 6 значений
    note: str = ""
    confidence: float = 0.0          # 0..1
    level: str = "—"                 # high | medium | low | —
    flags: list[str] = field(default_factory=list)  # причины снижения достоверности


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

        adult_mb = cat.rates.get("adult_mb")
        adult_exb = cat.rates.get("adult_exb")
        child_exb = cat.rates.get("child312_exb")
        single = cat.rates.get("single")

        if adult_mb:
            _write_row(ws, row + OFFSET_ADULT_MB, adult_mb)
            rr.written["Adult MB"] = adult_mb
        # 1A (Single): формула шаблона не учитывает надбавку за одноместное размещение —
        # пишем "Single net" из PDF поверх формулы.
        if single:
            _write_row(ws, row + OFFSET_SINGLE_1A, single)
            rr.written["1A (Single)"] = single
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

        _score_room(ws, row, cat, rr)
        if rr.level in ("low", "medium"):
            report.warnings.append(f"{code}: достоверность {rr.level} — {'; '.join(rr.flags)}")
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


def _parse_any_date(value) -> date | None:
    """Привести дату из PDF ('3-Jan-27') или шаблона ('03/01/2027', datetime) к date."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ("%d-%b-%y", "%d/%m/%Y", "%d/%m/%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _score_room(ws, row: int, cat: Category, rr: RoomReport) -> None:
    """
    Оценить достоверность извлечения комнаты. Заполняет rr.confidence/level/flags.

    Сигналы (каждый штраф снижает балл и добавляет флаг):
      * ровно 6 периодов в PDF;
      * обязательные ставки Adult MB / Adult ExB присутствуют и содержат 6 значений;
      * даты периодов из PDF совпадают с датами шаблона (по порядку);
      * кросс-проверка: parsed "Single net" ≈ Adult MB * 1.7 (как в формуле шаблона);
      * все значения положительны.
    """
    score = 1.0
    flags: list[str] = []

    # 1) число периодов
    n = len(cat.period_starts)
    if n != N_PERIODS:
        score -= 0.4
        flags.append(f"{n} периодов вместо {N_PERIODS}")

    # 2) полнота обязательных строк
    for key, label in (("adult_mb", "Adult MB"), ("adult_exb", "Adult ExB")):
        vals = cat.rates.get(key)
        if not vals:
            score -= 0.3
            flags.append(f"нет строки «{label}»")
        elif len(vals) != N_PERIODS:
            score -= 0.2
            flags.append(f"«{label}»: {len(vals)} значений вместо {N_PERIODS}")

    # 3) сверка дат периодов с шаблоном
    tmpl_dates = [_parse_any_date(ws.cell(row + OFFSET_PERIOD_START, c).value)
                  for c in PERIOD_COLUMNS]
    pdf_dates = [_parse_any_date(d) for d in cat.period_starts[:N_PERIODS]]
    mism = sum(1 for a, b in zip(tmpl_dates, pdf_dates)
               if a and b and a != b)
    if mism:
        score -= min(0.3, 0.1 * mism)
        flags.append(f"даты периодов расходятся с шаблоном ({mism})")

    # 4) кросс-проверка Single ≈ Adult MB * 1.7
    single = cat.rates.get("single")
    adult_mb = cat.rates.get("adult_mb")
    if single and adult_mb and len(single) == len(adult_mb):
        bad = sum(1 for s, a in zip(single, adult_mb)
                  if abs(s - a * SINGLE_MULTIPLIER) > 1.0)
        if bad:
            score -= min(0.3, 0.1 * bad)
            flags.append(f"Single ≠ AdultMB×1.7 в {bad} периодах")

    # 5) положительность значений
    for key, vals in cat.rates.items():
        if any((v is None or v <= 0) for v in vals):
            score -= 0.1
            flags.append(f"нечисловые/неположительные значения в «{key}»")
            break

    score = max(0.0, min(1.0, score))
    rr.confidence = round(score, 2)
    rr.flags = flags
    rr.level = "high" if score >= 0.9 else ("medium" if score >= 0.6 else "low")
