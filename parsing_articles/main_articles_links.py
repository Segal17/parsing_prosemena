#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Собирает ссылки на статьи блога prosemena.ru со страниц 1–6.
Сохраняет только список URL в articles_links.txt (по одному в строке).
"""

import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

VERSION = "v2-blog-links-txt"
BASE = "https://prosemena.ru"
BLOG_URL = f"{BASE}/blog/"
PAGES = range(1, 7)
REQUEST_DELAY = 1.2
OUT_TXT = Path("articles_links.txt")

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


def get_soup(url, timeout=25, retries=3):
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            r = session.get(url, timeout=timeout)
            r.raise_for_status()
            return BeautifulSoup(r.text, "html.parser")
        except Exception as e:
            last_exc = e
            time.sleep(2 * attempt)
    raise last_exc


def extract_article_links(soup):
    links = set()
    for card in soup.select(".b-posts-list .b-post-wrapper"):
        a = card.select_one(".b-post-name a") or card.select_one("a.card-img-top")
        if not a or not a.get("href"):
            continue
        url = urljoin(BASE, a["href"])
        path = urlparse(url).path.rstrip("/")
        if not path.startswith("/blog/") or path == "/blog":
            continue
        if len(path.strip("/").split("/")) < 3:
            continue
        links.add(url)
    return links


def main():
    print(f"main_articles_links.py {VERSION}")
    print(f"[i] Страницы: {list(PAGES)}\n")

    all_links = set()
    for page in PAGES:
        url = BLOG_URL if page == 1 else f"{BLOG_URL}?page={page}"
        try:
            soup = get_soup(url)
        except Exception as e:
            print(f"  [!] {url}: {e}", file=sys.stderr)
            continue
        found = extract_article_links(soup)
        new = found - all_links
        all_links.update(found)
        print(f"  [+] page={page}: найдено {len(found)} (новых {len(new)}, всего {len(all_links)})")
        time.sleep(REQUEST_DELAY)

    if not all_links:
        print("[!] Ни одной ссылки не собрано.", file=sys.stderr)
        sys.exit(1)

    with OUT_TXT.open("w", encoding="utf-8") as f:
        for u in sorted(all_links):
            f.write(u + "\n")

    print(f"\n[i] Всего статей: {len(all_links)}")
    print(f"    {OUT_TXT}")


if __name__ == "__main__":
    main()