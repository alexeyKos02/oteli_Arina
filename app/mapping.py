"""
Правила сопоставления net-rate PDF -> фиксированный Excel-шаблон.

Логика (проверена на эталонном заполненном файле Hyatt Ziva Cap Cana 27.xlsx):

Каждый блок комнаты в шаблоне начинается со строки, где A='Room', B='<Название (КОД)>'.
Относительно строки `Room` (br) расположены входные ячейки "Bed prices":

    br + 3  -> Adult    MB   (main bed)
    br + 4  -> Adult    ExB  (extra bed)
    br + 6  -> Child3-12 MB
    br + 7  -> Child3-12 ExB
    br + 9  -> Child0-2  MB
    br + 10 -> Child0-2  ExB
    br + 13 -> строка дат "from" (start) периодов   (уже заполнена в шаблоне)
    br + 14 -> строка дат "to"   (end)   периодов   (уже заполнена в шаблоне)

Цены за 6 периодов лежат в колонках K, M, O, Q, S, U.
Промежуточные колонки (L, N, ...) хранят "min stay" / "type" и не трогаются.

Все комбинации (1A, 2A, 1A+1C12, ...) в шаблоне — это массивные формулы,
которые сами считаются от "Bed prices". Мы их НЕ трогаем. В частности строку 1A
(single) не переписываем: формула даёт AdultMB*1.7, что совпадает с "Single net" из PDF.
"""

# 6 периодов -> колонки (1-indexed) K, M, O, Q, S, U
PERIOD_COLUMNS = [11, 13, 15, 17, 19, 21]

# Смещения строк входов "Bed prices" относительно строки Room
OFFSET_ADULT_MB = 3
OFFSET_ADULT_EXB = 4
OFFSET_CHILD312_MB = 6
OFFSET_CHILD312_EXB = 7
OFFSET_CHILD02_MB = 9
OFFSET_CHILD02_EXB = 10
OFFSET_PERIOD_START = 13  # строка дат "from"
OFFSET_PERIOD_END = 14    # строка дат "to"

# Метки строк в PDF -> ключ извлечённой ставки.
# Ключи нормализуются: lower-case, схлопнутые пробелы.
PDF_LABEL_MAP = {
    "double pp pn net": "adult_mb",
    "3rd & 4th pax per person net": "adult_exb",
    "child 3-12 net": "child312_exb",
    "single net": "single",  # не пишется, но извлекается для отчёта/сверки
}

# Количество периодов, которое ожидаем в каждом блоке
N_PERIODS = 6
