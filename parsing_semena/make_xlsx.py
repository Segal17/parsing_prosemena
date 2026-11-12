#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генерирует Excel-файлы по всем JSON из папки json_data.

Для каждого json_data/<name>.json создаёт xlsx_data/<name>.xlsx
(или .xls / .csv — в зависимости от флага).

Порядок выбора формата (без флагов):
  1) .xlsx через openpyxl   (Excel 2007+, LibreOffice, Google Sheets)
  2) .xls  через xlwt        (Excel 2003 и старше)
  3) .csv через встроенный csv (откроется вообще везде)

Одна строка = одна культура (товар). Поля sku и brand не выводятся.

Флаги:
  --xls    принудительно .xls (если стоит xlwt)
  --csv    принудительно .csv
  --xlsx   принудительно .xlsx
  (без флагов — авто: xlsx → xls → csv)

Примеры:
  python make_xlsx.py            # все .json → .xlsx (fallback xls/csv)
  python make_xlsx.py --xls      # все .json → .xls (для Excel 2003)
  python make_xlsx.py --csv      # все .json → .csv
"""

import csv
import json
import sys
from pathlib import Path

VERSION = "v3-batch"
INPUT_DIR = Path("json_data")
OUTPUT_DIR = Path("xlsx_data")
SHEET_NAME = "products"
MAX_TABLES = 5
MAX_COL_WIDTH = 60
MIN_COL_WIDTH = 10


# ---------------------------------------------------------------- flatten
def flat_variants(variants):
    if not variants:
        return ""
    parts = []
    for v in variants:
        title = v.get("title") or ""
        price = v.get("price") or ""
        if title and price:
            parts.append(f"{title}={price}")
        elif title:
            parts.append(title)
    return "; ".join(parts)


def flat_variants_urls(variants):
    return "\n".join(v.get("url") or "" for v in (variants or []))


def flat_tables(tables, max_tables=MAX_TABLES):
    out = {}
    for i in range(1, max_tables + 1):
        out[f"table_{i}_title"] = ""
        out[f"table_{i}_rows"] = ""
    for i, t in enumerate((tables or [])[:max_tables], start=1):
        title = t.get("title") or ""
        all_rows = [t.get("headers") or []] + (t.get("rows") or [])
        rendered = []
        for row in all_rows:
            rendered.append(": ".join("" if c is None else str(c) for c in row))
        out[f"table_{i}_title"] = title
        out[f"table_{i}_rows"] = " / ".join(rendered)
    return out


def flat_articles(ua):
    if not ua:
        return {"articles_title": "", "articles_titles": "", "articles_urls": ""}
    articles = ua.get("articles") or []
    return {
        "articles_title": ua.get("title") or "",
        "articles_titles": " | ".join(a.get("title") or "" for a in articles),
        "articles_urls": "\n".join(a.get("url") or "" for a in articles),
    }


def flatten_product(p):
    tables_cols = flat_tables(p.get("tables"))
    articles_cols = flat_articles(p.get("useful_articles"))
    return {
        "name": p.get("name") or "",
        "url": p.get("url") or "",
        "price": p.get("price") or "",
        "currency": p.get("currency") or "",
        "availability": p.get("availability") or "",
        "categories": "; ".join(p.get("categories") or []),
        "preview": p.get("preview") or "",
        "description": p.get("description") or "",
        "images": "\n".join(p.get("images") or []),
        "variants": flat_variants(p.get("variants")),
        "variants_urls": flat_variants_urls(p.get("variants")),
        **tables_cols,
        **articles_cols,
    }


# ---------------------------------------------------------------- xlsx
def write_xlsx(rows, fields, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    ws.append(fields)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r in rows:
        ws.append([r.get(fld) if r.get(fld) is not None else "" for fld in fields])

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for cell in ws["A"][1:]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(fields))}{ws.max_row}"

    for col_idx, fld in enumerate(fields, start=1):
        letter = get_column_letter(col_idx)
        max_len = len(fld)
        for cell in ws[letter][1:]:
            v = cell.value
            if v is None:
                continue
            for line in str(v).split("\n"):
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[letter].width = max(MIN_COL_WIDTH, min(MAX_COL_WIDTH, max_len + 2))

    wb.save(path)


# ---------------------------------------------------------------- xls
def write_xls(rows, fields, path):
    import xlwt

    if len(rows) + 1 > 65536:
        raise ValueError(f"слишком много строк для .xls: {len(rows)} > 65535")
    if len(fields) > 256:
        raise ValueError(f"слишком много колонок для .xls: {len(fields)} > 256")

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(SHEET_NAME)

    header_style = xlwt.easyxf("font: bold on; align: wrap on, horiz center, vert center")
    body_style = xlwt.easyxf("align: wrap on, vert top")
    name_style = xlwt.easyxf("font: bold on; align: wrap on, vert top")

    for c, fld in enumerate(fields):
        ws.write(0, c, fld, header_style)

    col_widths = [len(fld) for fld in fields]
    for r_idx, r in enumerate(rows, start=1):
        for c_idx, fld in enumerate(fields):
            v = r.get(fld) if r.get(fld) is not None else ""
            style = name_style if c_idx == 0 else body_style
            ws.write(r_idx, c_idx, str(v), style)
            for line in str(v).split("\n"):
                if len(line) > col_widths[c_idx]:
                    col_widths[c_idx] = len(line)

    for c_idx, w in enumerate(col_widths):
        w = max(MIN_COL_WIDTH, min(MAX_COL_WIDTH, w + 2))
        ws.col(c_idx).width = w * 256

    ws.panes_frozen = True
    ws.horz_split_pos = 1

    wb.save(path)


# ---------------------------------------------------------------- csv
def write_csv(rows, fields, path):
    with Path(path).open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


# ---------------------------------------------------------------- one file
def process_one(json_path, forced=None):
    """Обрабатывает один JSON. Возвращает (fmt, out_path) или бросает исключение."""
    try:
        with json_path.open(encoding="utf-8") as f:
            products = json.load(f)
    except Exception as e:
        raise RuntimeError(f"не удалось прочитать {json_path}: {e}")

    if not isinstance(products, list) or not products:
        raise RuntimeError(f"{json_path}: пусто или не список")

    rows = [flatten_product(p) for p in products]
    fields = list(rows[0].keys())

    order = [forced] if forced else ["xlsx", "xls", "csv"]
    last_exc = None
    for fmt in order:
        out_path = OUTPUT_DIR / f"{json_path.stem}.{fmt}"
        try:
            if fmt == "xlsx":
                write_xlsx(rows, fields, out_path)
            elif fmt == "xls":
                write_xls(rows, fields, out_path)
            elif fmt == "csv":
                write_csv(rows, fields, out_path)
            return fmt, out_path, len(rows), len(fields)
        except ImportError as e:
            last_exc = e
        except Exception as e:
            last_exc = e
    raise RuntimeError(f"ни один формат не сработал для {json_path}: {last_exc}")


# ---------------------------------------------------------------- main
def main():
    print(f"make_xlsx.py {VERSION}")

    # флаги
    args = set(sys.argv[1:])
    forced = None
    if "--xls" in args:
        forced = "xls"
    elif "--csv" in args:
        forced = "csv"
    elif "--xlsx" in args:
        forced = "xlsx"

    if not INPUT_DIR.exists():
        print(f"[!] Нет папки {INPUT_DIR}/ — создайте её и положите *.json", file=sys.stderr)
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    json_files = sorted(INPUT_DIR.glob("*.json"))
    if not json_files:
        print(f"[!] В {INPUT_DIR}/ нет *.json", file=sys.stderr)
        sys.exit(1)

    print(f"[i] Найдено JSON: {len(json_files)}")
    print(f"[i] Формат: {forced or 'авто (xlsx → xls → csv)'}")
    print(f"[i] Вывод: {OUTPUT_DIR}/\n")

    ok, fail = 0, 0
    for jp in json_files:
        try:
            fmt, out_path, n_rows, n_cols = process_one(jp, forced)
            print(f"  [+] {jp.name} → {out_path.name}  "
                  f"({fmt.upper()}, {n_rows} строк, {n_cols} колонок)")
            ok += 1
        except Exception as e:
            print(f"  [!] {jp.name}: {e}", file=sys.stderr)
            fail += 1

    print(f"\n[i] Готово: {ok} успешно, {fail} с ошибками")
    if fail:
        sys.exit(1)


if __name__ == "__main__":
    main()