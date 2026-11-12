# Парсер prosemena.ru → Excel

Собирает карточки товаров с prosemena.ru по культурам, сохраняет по одному
JSON на культуру и генерирует Excel (`.xlsx` / `.xls` / `.csv`).

## 1. Подготовка виртуального окружения

**Windows PowerShell:**

```powershell
cd C:\Users\lion_\learn\ggit\parsing_prosemena.ru\parsing_semena
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 2. Парсинг ссылок

main.py — листает категории из CATALOG_URLS, открывает карточки,
раскладывает по культурам и пишет json_data/<culture>.json.

```powershell
python main.py
Результат: json_data/amarant.json, json_data/gorokh.json, …
```

Настройки в шапке main.py: REQUEST_DELAY, CATALOG_DELAY,
MAX_PAGES_PER_CATALOG.

## 3. Генерация Excel

make_xlsx.py — берёт все json_data/\*.json и пишет
xlsx_data/<имя>.xlsx. Одна строка = одна культура.

```powershell
python make_xlsx.py # авто: xlsx → xls → csv
python make_xlsx.py --xlsx # современный Excel (2007+)
python make_xlsx.py --xls # старый Excel (2003)
python make_xlsx.py --csv # универсальный CSV
```
