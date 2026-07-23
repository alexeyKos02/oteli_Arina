"""
Извлечение net-rate из текстовых PDF (Этап 1 — стратегия «чистый текст»).

Разбираем PDF построчно в конечном автомате. Каждая «категория» комнаты в PDF —
это заголовок из 1-2 строк, где присутствуют коды комнат в скобках (King и Double
делят одни ставки), строки дат `from` / `to` (по 6 периодов) и строки ставок.

Результат — список Category с кодами, периодами и ставками. Robustness к «кривым»
страницам не гарантируется: то, что не распозналось, просто не попадёт в результат
и будет честно перечислено в отчёте на этапе сопоставления.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber

from .mapping import PDF_LABEL_MAP, N_PERIODS

# Код комнаты: 3-5 заглавных букв/цифр в скобках, напр. (SKNG), (1VWK), (2V1B)
CODE_RE = re.compile(r"\(([A-Z0-9]{3,5})\)")
# Дата вида 3-Jan-27
DATE_RE = re.compile(r"\d{1,2}-[A-Za-z]{3}-\d{2}")
# Денежная сумма: две цифры после точки, с опциональными разделителями тысяч
MONEY_RE = re.compile(r"\d[\d,]*\.\d{2}")


def _norm_label(text: str) -> str:
    """Нормализация начала строки-метки: lower-case, схлопнутые пробелы, без хвоста цифр."""
    # Отрезаем числовую часть (ставки/даты) — берём только текст до первой суммы.
    return re.sub(r"\s+", " ", text).strip().lower()


def _parse_money(text: str) -> list[float]:
    """Все денежные значения в строке (в порядке следования)."""
    return [float(m.replace(",", "")) for m in MONEY_RE.findall(text)]


def _match_label(line_lower: str) -> Optional[str]:
    """Сопоставить строку с известной меткой ставки. Возвращает ключ или None."""
    for label, key in PDF_LABEL_MAP.items():
        if line_lower.startswith(label):
            return key
    return None


@dataclass
class Category:
    """Одна ценовая категория PDF (набор кодов, делящих ставки)."""
    codes: list[str] = field(default_factory=list)
    period_starts: list[str] = field(default_factory=list)
    period_ends: list[str] = field(default_factory=list)
    rates: dict[str, list[float]] = field(default_factory=dict)  # key -> 6 значений

    def is_complete(self) -> bool:
        return bool(self.codes) and "adult_mb" in self.rates


def _reconstruct_lines(page, y_tol: float = 3.0, space_gap: float = 1.5) -> list[str]:
    """
    Собрать строки текста из координат символов (устойчиво к «кривой» вёрстке).

    Штатный extract_text() угадывает порядок чтения и на нестандартно свёрстанных
    таблицах (напр. стр. 2 Appendix A) рассыпает числа по вертикали. Здесь мы:
      1) кластеризуем символы в строки по вертикали (top с допуском y_tol);
      2) внутри строки сортируем по x и склеиваем, вставляя пробел там, где между
         символами есть заметный горизонтальный разрыв (> space_gap).
    Внутри одного числа разрывы малы -> оно остаётся цельным для regex.
    Работает и для нормальных страниц, поэтому применяем ко всем.
    """
    chars = [c for c in page.chars if c["text"].strip()]
    if not chars:
        return []
    chars.sort(key=lambda c: (round(c["top"], 1), c["x0"]))

    rows: list[list[dict]] = []
    cur: list[dict] = []
    cur_top: float | None = None
    for c in chars:
        if cur_top is None or abs(c["top"] - cur_top) <= y_tol:
            cur.append(c)
            cur_top = c["top"] if cur_top is None else cur_top
        else:
            rows.append(cur)
            cur = [c]
            cur_top = c["top"]
    if cur:
        rows.append(cur)

    lines: list[str] = []
    for row in rows:
        row.sort(key=lambda c: c["x0"])
        out = []
        prev_x1: float | None = None
        for c in row:
            if prev_x1 is not None and c["x0"] - prev_x1 > space_gap:
                out.append(" ")
            out.append(c["text"])
            prev_x1 = c["x1"]
        text = "".join(out).strip()
        if text:
            lines.append(text)
    return lines


def parse_pdf(path: str) -> tuple[list[Category], list[str]]:
    """
    Разобрать PDF. Возвращает (categories, raw_lines).
    raw_lines — плоский список строк текста (для диагностики/отчёта).
    """
    lines: list[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            lines.extend(_reconstruct_lines(page))

    categories: list[Category] = []
    current: Optional[Category] = None

    def flush():
        nonlocal current
        if current and current.is_complete():
            categories.append(current)
        current = None

    i = 0
    while i < len(lines):
        line = lines[i]
        low = line.lower()
        dates = DATE_RE.findall(line)
        codes = CODE_RE.findall(line)

        # Строка дат содержит >=6 дат. Слово-маркер отличает start («from») от end («to»).
        is_dateline = len(dates) >= N_PERIODS
        is_from = is_dateline and "from" in low
        is_to = is_dateline and not is_from  # King-строка = from, Double-строка = to

        if is_from:
            # Начало новой категории — сбросим предыдущую.
            flush()
            current = Category()
            current.period_starts = dates[:N_PERIODS]
            current.codes.extend(codes)
            # Заголовок кода может быть на этой же строке ИЛИ на соседних (до/после).
            # Смотрим строку выше — вдруг код там (случай "from" на отдельной строке).
            if i > 0:
                current.codes.extend(CODE_RE.findall(lines[i - 1]))
            i += 1
            continue

        if is_to and current is not None:
            current.period_ends = dates[:N_PERIODS]
            current.codes.extend(codes)
            # Код может быть на строке между from и to (случай 2VST).
            if i > 0:
                current.codes.extend(CODE_RE.findall(lines[i - 1]))
            i += 1
            continue

        if current is not None:
            key = _match_label(low)
            if key:
                money = _parse_money(line)
                if len(money) >= N_PERIODS:
                    current.rates[key] = money[:N_PERIODS]
                i += 1
                continue

        # Строка вне контекста категории — если встретили новый заголовок с кодом
        # без from/to рядом, игнорируем (страница 2 «кривая»).
        i += 1

    flush()

    # Убираем дубликаты кодов внутри категории, сохраняя порядок.
    for cat in categories:
        seen = set()
        uniq = []
        for c in cat.codes:
            if c not in seen:
                seen.add(c)
                uniq.append(c)
        cat.codes = uniq

    return categories, lines
