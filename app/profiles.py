"""
Профили форматов PDF. Формат определяется структурой PDF (разные сети/поставщики
присылают по-разному), а НЕ конкретным отелем. Каждый профиль связывает:
парсер PDF + целевой Excel-шаблон + колонки периодов + правила заполнения/оценки.

Добавить новый формат = добавить сюда Profile (+ при необходимости парсер в
parser.py и функцию заполнения в filler.py).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from . import parser as P
from . import filler as F


@dataclass
class Profile:
    id: str                       # идентификатор формата (не отеля)
    name: str                     # человекочитаемое имя (с указанием шаблона)
    template: str                 # файл шаблона в template/
    period_columns: list[int]     # колонки периодов в шаблоне (1-indexed)
    n_periods: int
    required_keys: list[str]      # ключи ставок, обязательные для высокой достоверности
    detect_keywords: list[str]    # маркеры в тексте PDF для авто-определения формата
    parse: Callable               # (path) -> (list[Category], list[str])
    fill_room: Callable           # (ws,row,cat,cols,room_rows) -> dict written
    single_multiplier: float | None = None  # для кросс-проверки Single (если есть)


# Колонки периодов
COLS_6 = [11, 13, 15, 17, 19, 21]                     # K..U
COLS_8 = [11, 13, 15, 17, 19, 21, 23, 25]             # K..Y


PROFILES: dict[str, Profile] = {
    "labeled_blocks": Profile(
        id="labeled_blocks",
        name="Labeled blocks (шаблон Hyatt Ziva Cap Cana)",
        template="Hyatt Ziva Cap Cana.xlsx",
        period_columns=COLS_6,
        n_periods=6,
        required_keys=["adult_mb", "adult_exb"],
        detect_keywords=["Double PP PN net", "NET RATE AGREEMENT"],
        parse=P.parse_labeled_blocks,
        fill_room=F.fill_labeled_blocks,
        single_multiplier=1.7,
    ),
    "rate_matrix": Profile(
        id="rate_matrix",
        name="Rate matrix (шаблон Riu Ventura)",
        template="Riu Ventura.xlsx",
        period_columns=COLS_8,
        n_periods=8,
        required_keys=["adult_mb"],
        detect_keywords=["CONTRACT RATES", "per person / day"],
        parse=P.parse_rate_matrix,
        fill_room=F.fill_rate_matrix,
    ),
}


def detect_profile(lines: list[str]) -> str | None:
    """Определить формат по маркерам в тексте PDF. Возвращает id профиля или None."""
    text = "\n".join(lines).lower()
    for prof in PROFILES.values():
        if any(kw.lower() in text for kw in prof.detect_keywords):
            return prof.id
    return None


def get_profile(profile_id: str) -> Profile | None:
    return PROFILES.get(profile_id)
