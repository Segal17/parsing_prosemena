#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор порядка страниц для печати книгой (booklet).

Ряд: (N,1), (N-2,3), …, (2,N-1), (4,N-3), …
Нечётное N добивается до кратного 4 одной и той же заглушкой.

  python booklet.py 12
  python booklet.py 10               # заглушка = 11
  python booklet.py 10 --blank 10    # заглушка = 10
"""

import math
import sys


def booklet_order(n_pages, blank_page=None):
    n = math.ceil(n_pages / 4) * 4
    blank = blank_page if blank_page is not None else n_pages + 1
    half = n // 2
    order = []
    for i in range(half // 2):
        order += [n - 2 * i, 1 + 2 * i]
    for i in range(half // 2):
        order += [2 + 2 * i, n - 1 - 2 * i]
    return n, [blank if p > n_pages else p for p in order]


def main():
    args = sys.argv[1:]
    blank_page = None
    if "--blank" in args:
        i = args.index("--blank")
        blank_page = int(args[i + 1])
        args = args[:i] + args[i + 2:]

    n_pages = int(args[0]) if args else int(input("Сколько страниц? "))
    if n_pages < 4:
        print("[!] Минимум 4 страницы", file=sys.stderr)
        sys.exit(1)

    n, order = booklet_order(n_pages, blank_page)
    padded = n - n_pages
    blank = blank_page if blank_page is not None else n_pages + 1
    sheets = n // 4

    # --- шапка ---
    print(f"=== Печать книгой ===")
    print(f"Реальных страниц:  {n_pages}")
    print(f"Добиваем до:       {n}  (пустых: {padded})")
    print(f"Листов:            {sheets}")
    print(f"Сторон на печать:  {sheets * 2}")
    if padded:
        print(f"Заглушка (одна и та же): страница {blank}")

    # --- стороны листов ---
    print("\nСтороны листов (левая | правая):")
    for i in range(0, len(order), 2):
        side = i // 2 + 1
        a, b = order[i], order[i + 1]
        print(f"  Сторона {side:>2}:  {a:>3}  |  {b:>3}")

    # --- листы ---
    print("\nЛисты (складываем лист пополам):")
    for s in range(sheets):
        a1, a2 = order[s * 4], order[s * 4 + 1]
        b1, b2 = order[s * 4 + 2], order[s * 4 + 3]
        print(f"  Лист {s + 1:>2}:  лицо) {a1:>3} | {a2:<3}    оборот) {b1:>3} | {b2:<3}")

    if padded:
        print(f"\n[!] Позиции, где было {n_pages + 1}…{n}, заменены на страницу {blank}.")
    print("\nПечать: landscape, 2 страницы на лист, двусторонняя, «переплёт по короткому краю».")

    # --- финальный ряд ---
    print("\nРяд для печати (через запятую):")
    print(", ".join(map(str, order)))


if __name__ == "__main__":
    main()