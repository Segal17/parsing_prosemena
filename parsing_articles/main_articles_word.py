#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсит статьи блога prosemena.ru → .docx по шаблону templates_new.docx.

Times New Roman 14pt — основной текст, переносы (кроме заголовков), без картинок.
Заголовки — Arial, по центру, без переносов, без красной строки.
Основной текст — по ширине, красная строка 1.25 см, межабзацный 3pt.
Курсив (blockquote) — Arial 11pt, с красной строкой, без сдвига влево.
Таблицы — без красной строки, TNR 12pt, рамки через XML.
Ctrl+C — аккуратное завершение.
"""

import copy
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup

VERSION = "v14-articles-word-headings-arial"

BASE = "https://prosemena.ru"
TEMPLATE = Path("templates_new.docx")
LINKS_FILE = Path("articles_links.txt")
OUT_DOCX = Path("articles_docx")
REQUEST_DELAY = 1.2

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


# ---------------------------------------------------------------- HTTP
def get_soup(url, timeout=25, retries=3):
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            r = session.get(url, timeout=timeout)
            r.raise_for_status()
            return BeautifulSoup(r.text, "html.parser")
        except requests.exceptions.RequestException as e:
            last_exc = e
            time.sleep(2 * attempt)
    raise last_exc


def slug_from_url(url):
    path = urlparse(url).path.rstrip("/")
    return path.split("/")[-1] or "article"


def is_real_docx(path):
    try:
        with open(path, "rb") as f:
            return f.read(2) == b"PK"
    except Exception:
        return False


def parse_article(url):
    soup = get_soup(url)
    h1 = (soup.select_one("h1.b-post-name span[itemprop='name']")
          or soup.select_one("h1"))
    title = h1.get_text(" ", strip=True) if h1 else ""

    body = soup.select_one(".b-content")
    if not body:
        raise RuntimeError("не найден .b-content (статья)")

    for sel in ("img", "picture", "figure", "figcaption",
                "script", "style", "noscript",
                ".picture-lb", ".b-share", ".b-post-actions", "iframe"):
        for tag in body.select(sel):
            tag.decompose()

    return {"url": url, "title": title, "body": body}


# ---------------------------------------------------------------- docx helpers
def _apply_font(style, name="Times New Roman"):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    style.font.name = name
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)


def _no_hyphenation(target):
    """Универсально для Style и Paragraph (w:suppressAutoHyphens)."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph

    if isinstance(target, Paragraph):
        p = target._p
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)
    else:
        pPr = target.element.get_or_add_pPr()

    el = pPr.find(qn("w:suppressAutoHyphens"))
    if el is None:
        el = OxmlElement("w:suppressAutoHyphens")
        pPr.append(el)
    el.set(qn("w:val"), "true")


def _get_list_style(doc, kind):
    candidates = {
        "bullet": ["List Bullet", "Список маркированный", "Маркированный список"],
        "number": ["List Number", "Список нумерованный", "Нумерованный список"],
    }[kind]
    existing = {s.name for s in doc.styles}
    for name in candidates:
        if name in existing:
            return name
    return None


def enable_hyphenation(doc, lang="ru-RU"):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    settings = doc.settings.element

    def set_bool(tag, value="true"):
        el = settings.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            settings.append(el)
        el.set(qn("w:val"), value)

    def set_int(tag, value):
        el = settings.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            settings.append(el)
        el.set(qn("w:val"), str(value))

    set_bool("w:autoHyphenation", "true")
    set_int("w:hyphenationZone", 360)
    set_int("w:consecutiveHyphenLimit", 2)
    set_bool("w:doNotHyphenateCaps", "true")


def setup_styles(doc):
    """TNR 14pt для Normal; Arial для заголовков (по центру, без переносов).
    Основной текст — с красной строкой и межабзацным 3pt."""
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    try:
        normal = doc.styles["Normal"]
        normal.font.size = Pt(14)
        _apply_font(normal, "Times New Roman")
        normal.paragraph_format.first_line_indent = Cm(1.25)
        normal.paragraph_format.left_indent = Cm(0)
        normal.paragraph_format.right_indent = Cm(0)
        normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.space_after = Pt(3)
    except Exception:
        pass

    heading_sizes = {
        "Heading 1": 18,
        "Heading 2": 15,
        "Heading 3": 13,
        "Title": 22,
    }
    for style_name, size in heading_sizes.items():
        try:
            st = doc.styles[style_name]
            _apply_font(st, "Arial")           # ← Arial для заголовков
            st.font.size = Pt(size)
            st.font.bold = True
            st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            st.paragraph_format.first_line_indent = Cm(0)
            st.paragraph_format.left_indent = Cm(0)
            st.paragraph_format.right_indent = Cm(0)
            st.paragraph_format.keep_with_next = True
            st.paragraph_format.keep_together = True
            st.paragraph_format.space_before = Pt(6)
            st.paragraph_format.space_after = Pt(3)
            _no_hyphenation(st)
        except KeyError:
            pass


def add_paragraph(doc, text, style=None, italic=False, indent_pt=None,
                  prefix=None, align=None, no_first_line_indent=False):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt

    if not text and not prefix:
        return None

    use_style = None
    if style:
        existing = {s.name for s in doc.styles}
        if style in existing:
            use_style = style

    p = doc.add_paragraph(style=use_style) if use_style else doc.add_paragraph()

    font_name = "Arial" if italic else "Times New Roman"
    font_size = Pt(11) if italic else None

    if prefix:
        run = p.add_run(prefix)
        run.italic = italic
        run.font.name = font_name
        if font_size:
            run.font.size = font_size
    if text:
        run = p.add_run(text)
        run.italic = italic
        run.font.name = font_name
        if font_size:
            run.font.size = font_size

    if indent_pt is not None:
        p.paragraph_format.left_indent = indent_pt
    if no_first_line_indent:
        p.paragraph_format.first_line_indent = Cm(0)
    if align is not None:
        p.paragraph_format.alignment = align

    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)

    return p


def _set_table_borders(table):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is not None:
        tblPr.remove(borders)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)


def add_table(doc, table_el):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    rows = []
    for tr in table_el.find_all("tr"):
        cells = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
        if cells:
            rows.append(cells)
    if not rows:
        return

    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.autofit = True
    _set_table_borders(table)

    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "autofit")
    tblPr.append(layout)

    for ri, row in enumerate(rows):
        for ci in range(cols):
            cell = table.cell(ri, ci)
            cell.text = row[ci] if ci < len(row) else ""
            for p in cell.paragraphs:
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.left_indent = Cm(0)
                p.paragraph_format.right_indent = Cm(0)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(12)

    doc.add_paragraph()


def write_body_elements(doc, body):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    bullet_style = _get_list_style(doc, "bullet")
    number_style = _get_list_style(doc, "number")

    for el in body.find_all(
        ["h2", "h3", "p", "ul", "ol", "blockquote", "table"],
        recursive=True,
    ):
        if el.name in ("p", "li") and el.find_parent(["p", "li", "blockquote"]):
            continue
        if el.name == "table" and el.find_parent("table"):
            continue
        if el.name in ("ul", "ol") and el.find_parent(["ul", "ol"]):
            continue

        if el.name == "h2":
            doc.add_heading(el.get_text(" ", strip=True), level=2)
        elif el.name == "h3":
            doc.add_heading(el.get_text(" ", strip=True), level=3)
        elif el.name == "p":
            add_paragraph(doc, el.get_text(" ", strip=True))
        elif el.name == "blockquote":
            add_paragraph(doc, el.get_text(" ", strip=True), italic=True)
        elif el.name in ("ul", "ol"):
            items = el.find_all("li", recursive=False)
            for idx, li in enumerate(items, start=1):
                t = li.get_text(" ", strip=True)
                if not t:
                    continue
                if el.name == "ul":
                    if bullet_style:
                        add_paragraph(doc, t, style=bullet_style,
                                      no_first_line_indent=True)
                    else:
                        add_paragraph(doc, t, prefix="• ", indent_pt=Pt(18),
                                      no_first_line_indent=True)
                else:
                    if number_style:
                        add_paragraph(doc, t, style=number_style,
                                      no_first_line_indent=True)
                    else:
                        add_paragraph(doc, t, prefix=f"{idx}. ",
                                      indent_pt=Pt(18),
                                      no_first_line_indent=True)
        elif el.name == "table":
            add_table(doc, el)


def collect_body_items(body):
    items = []
    for el in body.find_all(
        ["h2", "h3", "p", "ul", "ol", "blockquote", "table"],
        recursive=True,
    ):
        if el.name in ("p", "li") and el.find_parent(["p", "li", "blockquote"]):
            continue
        if el.name == "table" and el.find_parent("table"):
            continue
        if el.name in ("ul", "ol") and el.find_parent(["ul", "ol"]):
            continue

        if el.name in ("h2", "h3", "p", "blockquote"):
            t = el.get_text(" ", strip=True)
            if t:
                items.append((el.name, t))
        elif el.name in ("ul", "ol"):
            for li in el.find_all("li", recursive=False):
                t = li.get_text(" ", strip=True)
                if t:
                    items.append(("li-" + el.name, t))
        elif el.name == "table":
            rows = []
            for tr in el.find_all("tr"):
                cells = [td.get_text(" ", strip=True)
                         for td in tr.find_all(["td", "th"])]
                if cells:
                    rows.append(cells)
            if rows:
                items.append(("table", rows))
    return items


def rows_to_html_table(rows):
    html = "<table>" + "".join(
        "<tr>" + "".join(f"<td>{c}</td>" for c in row) + "</tr>" for row in rows
    ) + "</table>"
    return BeautifulSoup(html, "html.parser").find("table")


def replace_in_paragraph(paragraph, old, new):
    if old not in paragraph.text:
        return False
    full = paragraph.text.replace(old, new)
    for run in list(paragraph.runs):
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = full
    else:
        paragraph.add_run(full)
    return True


def replace_placeholder_everywhere(doc, old, new):
    replaced = False
    for p in doc.paragraphs:
        if replace_in_paragraph(p, old, new):
            replaced = True
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if replace_in_paragraph(p, old, new):
                        replaced = True
    return replaced


def find_paragraph_by_text(doc, needle):
    for p in doc.paragraphs:
        if needle in p.text:
            return p
    return None


def write_docx_from_template(article, path):
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document(str(TEMPLATE))

    setup_styles(doc)
    enable_hyphenation(doc, "ru-RU")

    bullet_style = _get_list_style(doc, "bullet")
    number_style = _get_list_style(doc, "number")

    title = article["title"] or ""
    title_done = replace_placeholder_everywhere(doc, "{{TITLE}}", title)

    body = article["body"]
    body_par = find_paragraph_by_text(doc, "{{BODY}}")

    if body_par is not None:
        for r in list(body_par.runs):
            r.text = ""

        items = collect_body_items(body)
        cur = body_par
        num_idx = 0

        for kind, val in items:
            new_par = cur.insert_paragraph_before()
            cur._p.addnext(new_par._p)

            new_par.paragraph_format.first_line_indent = Cm(1.25)
            new_par.paragraph_format.left_indent = Cm(0)
            new_par.paragraph_format.right_indent = Cm(0)
            new_par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            new_par.paragraph_format.space_before = Pt(0)
            new_par.paragraph_format.space_after = Pt(3)

            if kind == "h2":
                new_par.style = doc.styles["Heading 2"]
                new_par.add_run(val)
                new_par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                new_par.paragraph_format.first_line_indent = Cm(0)
                new_par.paragraph_format.space_before = Pt(6)
                new_par.paragraph_format.space_after = Pt(3)
                _no_hyphenation(new_par)
                for r in new_par.runs:
                    r.font.name = "Arial"           # ← Arial
            elif kind == "h3":
                new_par.style = doc.styles["Heading 3"]
                new_par.add_run(val)
                new_par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                new_par.paragraph_format.first_line_indent = Cm(0)
                new_par.paragraph_format.space_before = Pt(6)
                new_par.paragraph_format.space_after = Pt(3)
                _no_hyphenation(new_par)
                for r in new_par.runs:
                    r.font.name = "Arial"           # ← Arial
            elif kind == "p":
                new_par.add_run(val)
            elif kind == "blockquote":
                run = new_par.add_run(val)
                run.italic = True
                run.font.name = "Arial"
                run.font.size = Pt(11)
                new_par.paragraph_format.first_line_indent = Cm(1.25)
            elif kind == "li-ul":
                new_par.paragraph_format.first_line_indent = Cm(0)
                if bullet_style:
                    new_par.style = doc.styles[bullet_style]
                    new_par.add_run(val)
                else:
                    new_par.add_run("• ")
                    new_par.add_run(val)
                    new_par.paragraph_format.left_indent = Pt(18)
            elif kind == "li-ol":
                num_idx += 1
                new_par.paragraph_format.first_line_indent = Cm(0)
                if number_style:
                    new_par.style = doc.styles[number_style]
                    new_par.add_run(val)
                else:
                    new_par.add_run(f"{num_idx}. ")
                    new_par.add_run(val)
                    new_par.paragraph_format.left_indent = Pt(18)
            elif kind == "table":
                add_table(doc, rows_to_html_table(val))

            if kind not in ("blockquote", "h2", "h3"):
                for r in new_par.runs:
                    r.font.name = "Times New Roman"

            cur = new_par
    else:
        write_body_elements(doc, body)

    if not title_done and title:
        first = doc.paragraphs[0] if doc.paragraphs else None
        if first is not None:
            new_p = first.insert_paragraph_before()
            run = new_p.add_run(title)
            run.bold = True
            run.font.size = Pt(18)
            run.font.name = "Arial"                # ← Arial
            new_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_p.paragraph_format.first_line_indent = Cm(0)
            new_p.paragraph_format.space_before = Pt(0)
            new_p.paragraph_format.space_after = Pt(6)
            _no_hyphenation(new_p)
        else:
            p = doc.add_paragraph()
            run = p.add_run(title)
            run.bold = True
            run.font.size = Pt(18)
            run.font.name = "Arial"                # ← Arial
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            _no_hyphenation(p)

    doc.save(str(path))


# ---------------------------------------------------------------- main
def main():
    print(f"main_articles_word.py {VERSION}")
    OUT_DOCX.mkdir(exist_ok=True)

    if not TEMPLATE.exists():
        print(f"[!] Нет {TEMPLATE} — положите шаблон рядом со скриптом", file=sys.stderr)
        sys.exit(1)
    if not is_real_docx(TEMPLATE):
        print(f"[!] {TEMPLATE.name} не является .docx (не ZIP).", file=sys.stderr)
        print("    Сконвертируйте в .docx (Word: Файл → Сохранить как → .docx)")
        sys.exit(1)

    if not ARTICLES:
        print("[!] Список ARTICLES пуст", file=sys.stderr)
        sys.exit(1)

    print(f"[i] Статей: {len(ARTICLES)}")
    print(f"[i] Шаблон: {TEMPLATE.name}")
    print(f"[i] Шрифт: Times New Roman 14pt (текст), Arial (заголовки, курсив)")
    print(f"[i] Заголовки: Arial, по центру, без переносов")
    print(f"[i] Основной текст: красная строка 1.25 см, по ширине, межабзацный 3pt")
    print(f"[i] Курсив: Arial 11pt, с красной строкой, без сдвига влево")
    print(f"[i] Таблицы: без красной строки, TNR 12pt, рамки, без интервалов")
    print(f"[i] Переносы: включены (кроме заголовков)")
    print(f"[i] Картинки: не сохраняются")
    print(f"[i] Вывод: {OUT_DOCX}/\n")

    ok, fail = 0, 0
    try:
        for i, url in enumerate(ARTICLES, 1):
            slug = slug_from_url(url)
            try:
                art = parse_article(url)
            except Exception as e:
                print(f"  [{i}/{len(ARTICLES)}] FAIL {url}: {e}", file=sys.stderr)
                fail += 1
                time.sleep(REQUEST_DELAY)
                continue

            docx_path = OUT_DOCX / f"{slug}.docx"
            try:
                write_docx_from_template(art, docx_path)
            except Exception as e:
                print(f"  [{i}/{len(ARTICLES)}] DOCX FAIL {slug}: {e}", file=sys.stderr)
                fail += 1
                time.sleep(REQUEST_DELAY)
                continue

            title = art["title"] or slug
            print(f"  [{i}/{len(ARTICLES)}] OK  {title[:60]:60}  → {docx_path.name}")
            ok += 1
            time.sleep(REQUEST_DELAY)
    except KeyboardInterrupt:
        print(f"\n[!] Прервано пользователем (Ctrl+C).", file=sys.stderr)
    finally:
        print(f"\n[i] Готово: {ok} статей, {fail} ошибок")
        print(f"    DOCX: {OUT_DOCX}/")


if __name__ == "__main__":
    main()