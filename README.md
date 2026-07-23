# Net Rates → Excel (MVP, Этап 1)

Веб-приложение: загружает текстовый PDF с net-тарифами отеля и возвращает
заполненный Excel-шаблон Hyatt Ziva Cap Cana + отчёт о сопоставлении.

## Что делает

1. Принимает один текстовый PDF (`Appendix A`).
2. Извлекает net rates по комнатам и 6 сезонным периодам.
3. Пишет базовые ставки «Bed prices» в фиксированный шаблон
   `template/Hyatt Ziva Cap Cana.xlsx`. Комбинации (1A, 2A, 1A+1C12 …) в шаблоне —
   это готовые массивные формулы, которые считаются автоматически; мы их не трогаем.
4. Показывает предпросмотр и отчёт: что сопоставилось, что нет.

## Правила сопоставления PDF → шаблон

| Строка PDF | Ячейка «Bed prices» |
|---|---|
| `Double PP PN net` | Adult **MB** |
| `3rd & 4th pax per person net` | Adult **ExB** |
| `Child 3-12 net` | Child 3-12 **ExB** |
| *(= Adult MB, если есть строка Child 3-12)* | Child 3-12 **MB** |
| `Single net` | не пишется — формула строки 1A даёт то же (AdultMB×1.7) |

Каждый заголовок PDF содержит два кода (King + Double), делящих ставки — заполняются
оба блока шаблона. Комнаты сопоставляются по коду в скобках, напр. `(SKNG)`.

## Запуск

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8077
```

Открыть http://127.0.0.1:8077 , загрузить PDF, нажать «Обработать», скачать `.xlsx`.

## Деплой (фронт → GitHub Pages, бэк → Railway)

Фронт статичный (`docs/`), бэкенд — FastAPI на Railway. Один репозиторий.

**Бэкенд на Railway:**
1. Установить CLI: `npm i -g @railway/cli` (или `brew install railway`).
2. `railway login` (откроется браузер — это ваш аккаунт).
3. Из корня репозитория: `railway init` → `railway up`.
   Билд по `railway.json` / `Procfile` (`uvicorn app.main:app --host 0.0.0.0 --port $PORT`).
4. `railway domain` — получить публичный URL, напр. `https://<app>.up.railway.app`.
5. (Опц.) переменная `ALLOWED_ORIGINS=https://<user>.github.io` — сузить CORS.

**Фронт на GitHub Pages:**
1. В `docs/config.js` вписать URL Railway в `window.API_BASE`.
2. Закоммитить, запушить в `main`.
3. Settings → Pages → Source: `Deploy from a branch`, ветка `main`, папка `/docs`
   (или `gh api` — см. ниже). Сайт: `https://<user>.github.io/oteli_Arina/`.

Локально фронт всегда ходит на тот же origin, `config.js` игнорируется.

## Структура

```
app/
  main.py       FastAPI: /  и  /api/process
  parser.py     PDF -> категории комнат/периоды/ставки (стратегия «чистый текст»)
  filler.py     заполнение шаблона + отчёт
  mapping.py    правила: смещения строк, колонки периодов, метки PDF
  templates/index.html   веб-морда (vanilla JS)
template/       фиксированный пустой Excel-шаблон
samples/        пример PDF и эталонный заполненный xlsx
```

## Ограничения MVP (задел на Этап 2)

- **«Кривые» PDF.** На тестовом Appendix A страница 2 (комнаты `2VCK`, `2V1B`,
  `3VST`, `SWST`, `PRES`) свёрстана вертикально: извлечение текста рассыпает таблицу,
  поэтому эти комнаты попадают в отчёт как «не сопоставлено». Стратегия по координатам
  слов — задача Этапа 2.
- **Child 3-12 MB** заполняется по правилу «= Adult MB по каждому периоду» (так у 13 из
  15 комнат эталона). В ручном эталоне у первой категории `SKNG`/`SDBL` эта строка
  заполнена плоским `402.75` по всем периодам — вероятно, ручная протяжка; наш вывод
  внутренне консистентнее.
- Шаблон фиксированный (один отель/год).
```
