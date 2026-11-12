# Парсер статей блога prosemena.ru → PDF

Собирает ссылки на статьи со страниц блога и сохраняет каждую статью в PDF
(как «Ctrl+P → Сохранить как PDF» в браузере) — с картинками, таблицами,
цитатами, но без меню, шапки, подвала и блока «Еще статьи на эту тему».

## 1. Подготовка виртуального окружения

**Windows PowerShell:**

```powershell
cd C:\Users\lion_\learn\ggit\parsing_prosemena.ru\parsing_semena
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 2. Сбор ссылок на статьи

main_articles_links.py — листает /blog/?page=1 … ?page=6,
вытаскивает ссылки на карточки статей и сохраняет в articles_links.txt
(один URL на строку).

```powershell
python main_articles_links.py
```

Настройки в шапке main.py: REQUEST_DELAY, CATALOG_DELAY,
MAX_PAGES_PER_CATALOG.

## 3. Генерация PDF

main_articles.py — читает articles_links.txt, открывает каждую статью
в Chromium через Playwright, скрывает лишнее через CSS и печатает в PDF.
Результат — articles_pdf/<slug>.pdf.

```powershell
python main_articles.py
```
