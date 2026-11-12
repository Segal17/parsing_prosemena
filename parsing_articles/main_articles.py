#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсит статьи блога prosemena.ru → PDF через Playwright.

PDF = как браузер «Ctrl+P → Сохранить как PDF»:
  - картинки, таблицы, цитаты, стили — как на сайте;
  - без шапки, меню, подвала, сайдбаров, «Еще статьи на эту тему»;
  - с заголовком статьи, без автора и даты.

ARTICLES — явный список URL.
"""

import sys
import time
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup

VERSION = "v4-articles-pdf-only"

BASE = "https://prosemena.ru"

# --- список статей: явно, как CATALOG_URLS ---
# ARTICLES = [
#     f"{BASE}/blog/kak-vyrastit/kak-vyrastit-bazilik-iz-semyan-tekhnologicheskaya-karta/",
#     # сюда добавляйте остальные URL статей
# ]
# --- список статей: читаем из articles_links.txt ---
LINKS_FILE = Path("articles_links.txt")
if LINKS_FILE.exists():
    ARTICLES = [
        line.strip()
        for line in LINKS_FILE.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
else:
    ARTICLES = [
        f"{BASE}/blog/kak-vyrastit/kak-vyrastit-bazilik-iz-semyan-tekhnologicheskaya-karta/",
    ]
    print(f"[!] {LINKS_FILE} не найден — использую один тестовый URL", file=sys.stderr)

REQUEST_DELAY = 1.2
OUT_PDF = Path("articles_pdf")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.5",
}

session = requests.Session()
session.headers.update(HEADERS)

# CSS: прячем всё лишнее, но ОСТАВЛЯЕМ заголовок статьи.
HIDE_CSS = """
/* --- скрываем всё лишнее --- */
header, footer,
.header, .header-fixed, .header-desktop, .header-bottom,
.c-sidebar-wrapper, .c-sidebar-nav, .c-sidebar-widget, .c-sidebar-subscribe,
.c-advantages,
.b-posts-wrapper,             /* блок «Еще статьи на эту тему» */
.accept-cookies, .fly-widgets, .c-backtotop, .c-toasts,
.c-search, .c-icons, .c-auth,
.c-logo, .c-logo-desc, .c-address, .c-phones,
.breadcrumbs, .global-header,
.picture-lb, script, style, iframe, .b-share, .b-post-actions {
    display: none !important;
}

/* --- дубль заголовка перед «Еще статьи» (по обёртке, не по h1) --- */
div.border-top.b-post-name.display-4 {
    display: none !important;
}

/* --- шапку статьи НЕ скрываем, но прячем автора/дату/аватар --- */
.b-post-header .avatar,
.b-post-header .b-post-author,
.b-post-header time,
.b-post-header .row.py-5 {
    display: none !important;
}

/* --- причесать контент --- */
body { background: #fff !important; }
main { margin: 0 !important; padding: 0 !important; }
.b-content-wrapper, .b-content {
    max-width: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
}
.b-content { font-size: 11pt; line-height: 1.5; color: #222; }
.b-content h2 { color: #28a745; }
.b-content img { max-width: 100%; height: auto; }
.b-content blockquote {
    border-left: 3pt solid #28a745;
    background: #f6f6f6;
    padding: 8pt 12pt;
    margin: 10pt 0;
}
.b-content table { border-collapse: collapse; width: 100%; }
.b-content td, .b-content th { border: 1px solid #ccc; padding: 4pt 6pt; font-size: 10pt; }
"""


def slug_from_url(url):
    path = urlparse(url).path.rstrip("/")
    return path.split("/")[-1] or "article"


def get_title(url):
    """Быстрый GET только чтобы вытащить title для лога."""
    try:
        r = session.get(url, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        h1 = (soup.select_one("h1.b-post-name span[itemprop='name']")
              or soup.select_one("h1"))
        return h1.get_text(" ", strip=True) if h1 else ""
    except Exception:
        return ""


def write_pdf_playwright(url, path):
    """
    Открывает URL статьи в Chromium, инжектит HIDE_CSS,
    печатает в PDF как браузер (Ctrl+P → Сохранить как PDF).
    """
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url, wait_until="networkidle", timeout=60000)
        page.add_style_tag(content=HIDE_CSS)
        page.emulate_media(media="print")
        page.pdf(
            path=str(path),
            format="A4",
            margin={"top": "18mm", "right": "16mm",
                    "bottom": "18mm", "left": "16mm"},
            print_background=True,
        )
        browser.close()


def main():
    print(f"main_articles.py {VERSION}")
    OUT_PDF.mkdir(exist_ok=True)

    if not ARTICLES:
        print("[!] Список ARTICLES пуст — добавьте URL статей", file=sys.stderr)
        sys.exit(1)

    print(f"[i] Статей: {len(ARTICLES)}\n")
    ok, fail = 0, 0

    for i, url in enumerate(ARTICLES, 1):
        slug = slug_from_url(url)
        title = get_title(url)
        pdf_path = OUT_PDF / f"{slug}.pdf"

        try:
            write_pdf_playwright(url, pdf_path)
        except Exception as e:
            print(f"  [{i}/{len(ARTICLES)}] FAIL {slug}: {e}", file=sys.stderr)
            fail += 1
            time.sleep(REQUEST_DELAY)
            continue

        print(f"  [{i}/{len(ARTICLES)}] OK  {(title or slug)[:60]:60}  → {pdf_path.name}")
        ok += 1
        time.sleep(REQUEST_DELAY)

    print(f"\n[i] Готово: {ok} статей, {fail} ошибок")
    print(f"    PDF: {OUT_PDF}/")


if __name__ == "__main__":
    main()