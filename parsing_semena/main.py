#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер prosemena.ru → по одному JSON на культуру.

CATALOG_URLS — явный список категорий (ваш).
Культура товара = slug его категории (последний сегмент URL
/semena/<slug>/, например /semena/amarant/ → amarant).

Файлы: <culture>.json в текущей папке.
"""

import copy
import json
import re
import sys
import time
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

VERSION = "v5-explicit-catalogs"

BASE = "https://prosemena.ru"

CATALOG_URLS = [
    # f"{BASE}/semena/amarant/",
    # f"{BASE}/semena/anis/",
    # f"{BASE}/semena/bazilik/",



    # f"{BASE}/semena/borago-ogurechnaya-trava/",
    # f"{BASE}/semena/brokkoli/",
    # f"{BASE}/semena/gorokh/",
    # f"{BASE}/semena/gorchitsa/",
    # f"{BASE}/semena/dushitsa/",
    # f"{BASE}/semena/zveroboi/",
    # f"{BASE}/semena/issop/",


    # f"{BASE}/semena/kapusta/",
    # f"{BASE}/semena/kervel/",
    # f"{BASE}/semena/koriandr-kinza/",
    # f"{BASE}/semena/klever/",
    # f"{BASE}/semena/kolrabi/",
    # f"{BASE}/semena/kress-salat/",
    # f"{BASE}/semena/kukuruza/",
    # f"{BASE}/semena/lebeda/",
    # f"{BASE}/semena/luk/",
    # f"{BASE}/semena/lyubistok/",


    f"{BASE}/semena/lyupin/",
    f"{BASE}/semena/mangold/",
    f"{BASE}/semena/mayoran/",
    f"{BASE}/semena/melissa/",
    f"{BASE}/semena/morkov/",
    f"{BASE}/semena/myata/",
    f"{BASE}/semena/pazhitnik/",
    f"{BASE}/semena/paprika/",
    f"{BASE}/semena/pasternak/",
    f"{BASE}/semena/perilla/",
    f"{BASE}/semena/petrushka/",
    f"{BASE}/semena/podsolnechnik/",
    f"{BASE}/semena/portulak/",
    f"{BASE}/semena/rastoropsha/",
    f"{BASE}/semena/redis/",
    f"{BASE}/semena/redka/",
    f"{BASE}/semena/repa/",
    f"{BASE}/semena/rozmarin/",
    f"{BASE}/semena/rukola-rukkola/",
    f"{BASE}/semena/salat/",
    f"{BASE}/semena/svekla/",
    f"{BASE}/semena/selderey/",
    f"{BASE}/semena/spilantes/",
    f"{BASE}/semena/tarkhun-estragon/",
    f"{BASE}/semena/timyan/",
    f"{BASE}/semena/tmin/",
    f"{BASE}/semena/ukrop/",
    f"{BASE}/semena/ustrichnyy-list-mertenziya/",
    f"{BASE}/semena/fenkhel/",
    f"{BASE}/semena/chaber-perechnaya-trava/",
    f"{BASE}/semena/shalfey/",
    f"{BASE}/semena/shpinat/",
    f"{BASE}/semena/shchavel/",
    f"{BASE}/semena/endivii/",
]

REQUEST_DELAY = 1.2
CATALOG_DELAY = 1.5
MAX_PAGES_PER_CATALOG = 30

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.5",
}

session = requests.Session()
session.headers.update(HEADERS)


# --------------------------------------------------------------------------
# Утилита: из URL категории достать культуру (slug)
# --------------------------------------------------------------------------
def culture_from_catalog_url(cat_url):
    """
    https://prosemena.ru/semena/amarant/  → 'amarant'
    https://prosemena.ru/semena/borago-ogurechnaya-trava/ → 'borago-ogurechnaya-trava'
    """
    path = urlparse(cat_url).path.rstrip("/")
    slug = path.split("/")[-1]
    return slug or "unknown"


# --------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------
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


# --------------------------------------------------------------------------
# Листинг / пагинация
# --------------------------------------------------------------------------
def extract_product_links(soup):
    urls = set()
    for a in soup.select(
        ".s-product-wrapper a.s-product-header, "
        ".s-product-wrapper a.s-image-wrapper"
    ):
        href = a.get("href")
        if href:
            urls.add(urljoin(BASE, href))
    return urls


def has_next_page(soup, current_page):
    if soup.select_one("a[rel='next']"):
        return True
    nxt = str(current_page + 1)
    for a in soup.select("a[href]"):
        if re.search(rf"[?&]page={nxt}(\D|$)", a.get("href", "")):
            return True
    return False


def collect_product_urls(cat_url):
    all_urls = set()
    page = 1
    while page <= MAX_PAGES_PER_CATALOG:
        url = cat_url if page == 1 else f"{cat_url}?page={page}"
        try:
            soup = get_soup(url)
        except Exception as e:
            print(f"    [!] {url}: {e}", file=sys.stderr)
            break
        links = extract_product_links(soup)
        all_urls.update(links)
        print(f"    [+] {url}: +{len(links)} (всего {len(all_urls)})")
        if not links or not has_next_page(soup, page):
            break
        page += 1
        time.sleep(CATALOG_DELAY)
    return all_urls


# --------------------------------------------------------------------------
# Карточка (те же парсеры, что и в v4)
# --------------------------------------------------------------------------
def _info_root(soup):
    return (
        soup.select_one("#overview-tab")
        or soup.select_one("#s-product-desc")
        or soup.select_one("#s-product-info")
    )


def parse_ldjson(soup):
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "")
        except (json.JSONDecodeError, TypeError):
            continue
        if isinstance(data, dict) and data.get("@type") == "Product":
            return data
    return None


def parse_preview(soup):
    el = soup.select_one("p.s-summary")
    return el.get_text(" ", strip=True) if el else None


def parse_full_description(soup):
    root = _info_root(soup)
    if not root:
        return None
    root = copy.copy(root)
    for tag in root.select("table, h2, h3, blockquote"):
        tag.decompose()
    paragraphs = []
    for p in root.find_all("p"):
        if p.find_parent(class_="table-responsive"):
            continue
        paragraphs.append(p.get_text(" ", strip=True))
    text = "\n\n".join(p for p in paragraphs if p)
    return text or None


def parse_tables(soup):
    root = _info_root(soup)
    if not root:
        return []
    tables = []
    current_title = None
    for el in root.find_all(["h2", "h3", "table"]):
        if el.name in ("h2", "h3"):
            txt = el.get_text(" ", strip=True)
            if txt.lower().startswith("описан"):
                current_title = None
                continue
            current_title = txt
        elif el.name == "table":
            rows, headers = [], []
            for tr in el.find_all("tr"):
                cells = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
                if not cells:
                    continue
                if not rows and not headers and tr.find("th"):
                    headers = cells
                else:
                    rows.append(cells)
            if not headers and rows:
                headers = rows.pop(0)
            tables.append({"title": current_title, "headers": headers, "rows": rows})
    return tables


def parse_useful_articles(soup):
    root = _info_root(soup)
    if not root:
        return None
    for bq in root.select("blockquote"):
        h = bq.find(["h3", "h2"])
        title = h.get_text(" ", strip=True) if h else None
        links = bq.select("ol li a[href]") or bq.select("ul li a[href]")
        if not links:
            continue
        return {
            "title": title,
            "articles": [
                {"title": a.get_text(" ", strip=True), "url": urljoin(BASE, a.get("href"))}
                for a in links
            ],
        }
    return None


def parse_skus(soup):
    skus = []
    for inp in soup.select("input[name='skus']"):
        skus.append({
            "sku_id": inp.get("value"),
            "title": inp.get("data-title"),
            "article": inp.get("data-article"),
            "price": inp.get("data-price"),
            "compare_price": inp.get("data-compare-price"),
            "product_id": inp.get("data-product"),
            "url": urljoin(BASE, inp.get("data-href") or ""),
        })
    return skus


def parse_categories(soup):
    block = soup.select_one("#s-product-categories")
    if not block:
        return []
    return [a.get_text(strip=True) for a in block.select("a")]


def parse_product(url):
    soup = get_soup(url)
    ld = parse_ldjson(soup) or {}
    offers = ld.get("offers") or {}
    if isinstance(offers, list):
        offers = offers[0] if offers else {}
    images = ld.get("image") or []
    if isinstance(images, dict):
        images = [images]

    return {
        "url": ld.get("@id") or url,
        "name": ld.get("name"),
        "sku": ld.get("sku"),
        "brand": (ld.get("brand") or {}).get("name"),
        "price": offers.get("price"),
        "currency": offers.get("priceCurrency"),
        "availability": offers.get("availability"),
        "categories": parse_categories(soup),
        "preview": parse_preview(soup),
        "description": parse_full_description(soup),
        "tables": parse_tables(soup),
        "useful_articles": parse_useful_articles(soup),
        "images": [img.get("url") for img in images if isinstance(img, dict)],
        "variants": parse_skus(soup),
    }


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def main():
    print(f"main.py {VERSION}")
    print(f"[i] Категорий в CATALOG_URLS: {len(CATALOG_URLS)}")

    # 1. Собираем ссылки по каждой категории, привязывая к культуре
    #    culture -> set(url)
    culture_links = {}
    for cat_url in CATALOG_URLS:
        culture = culture_from_catalog_url(cat_url)
        print(f"\n[1/3] Культура '{culture}' ({cat_url})")
        links = collect_product_urls(cat_url)
        if links:
            culture_links.setdefault(culture, set()).update(links)
        time.sleep(CATALOG_DELAY)

    if not culture_links:
        print("[!] Ни одной ссылки не собрано.", file=sys.stderr)
        return

    # 2. Уникальные карточки — не парсим один URL дважды
    all_unique = set()
    for links in culture_links.values():
        all_unique.update(links)
    print(f"\n[2/3] Уникальных карточек: {len(all_unique)}")

    parsed = {}  # url -> product dict
    for i, u in enumerate(sorted(all_unique), 1):
        try:
            parsed[u] = parse_product(u)
            n_t = len(parsed[u]["tables"])
            n_a = len((parsed[u]["useful_articles"] or {}).get("articles") or [])
            print(f"  [{i}/{len(all_unique)}] OK  {parsed[u]['name']}  "
                  f"(таблиц={n_t}, статей={n_a})")
        except Exception as e:
            print(f"  [{i}/{len(all_unique)}] FAIL {u}: {e}", file=sys.stderr)
        time.sleep(REQUEST_DELAY)

    # 3. Раскладываем по культурам и пишем <culture>.json
    print(f"\n[3/3] Пишем JSON по культурам…")
    total = 0
    for culture in sorted(culture_links):
        links = culture_links[culture]
        seen = set()
        items = []
        for u in links:
            if u in parsed and u not in seen:
                items.append(parsed[u])
                seen.add(u)
        if not items:
            continue
        path = f"{culture}.json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
        print(f"  [+] {path}: {len(items)} товаров")
        total += len(items)

    print(f"\n[i] Всего товаров (с учётом дублей по культурам): {total}")
    print(f"[i] Уникальных карточек: {len(parsed)}")


if __name__ == "__main__":
    main()