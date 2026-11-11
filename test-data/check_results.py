#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_results.py — сверка result.xlsx (из distribute.py) с expected_result.xlsx.

Метрики:
  - точность по полям: статус ФУЛ, группа, логист, комментарий повтора
  - доля ФУЛ, полностью совпавших с эталоном
  - число «выдуманных» значений: там, где в эталоне «не определено»,
    в результате должно быть ровно «не определено»
  - проверка формы Cloud: 12 колонок, порядок заголовков, нет дублей заказов

Запуск: python3 check_results.py <путь_к_result.xlsx>
"""

import os
import sys
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ND = "не определено"

def rows_by_ful(ws, ful_col=0):
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[ful_col]:
            out[str(row[ful_col]).strip()] = row
    return out

def main(result_path):
    exp_wb = load_workbook(os.path.join(HERE, "expected_result.xlsx"))
    res_wb = load_workbook(result_path)

    exp = rows_by_ful(exp_wb["Эталон"])
    journal = rows_by_ful(res_wb["Журнал решений"])
    repeats = rows_by_ful(res_wb["Повторные отправки"])

    fields = [("статус", 4, 1), ("группа", 5, 3), ("логист", 6, 4)]
    total = {name: 0 for name, _, _ in fields}
    ok = {name: 0 for name, _, _ in fields}
    full_match = 0
    hallucinations = []
    comment_ok = 0
    comment_total = 0

    for ful, erow in exp.items():
        jrow = journal.get(ful)
        if jrow is None:
            for name in total:
                total[name] += 1
            hallucinations.append(f"{ful}: нет в журнале решений")
            continue
        all_ok = True
        for name, jidx, eidx in fields:
            total[name] += 1
            jval = str(jrow[jidx] or "").strip()
            eval_ = str(erow[eidx] or "").strip()
            if jval == eval_:
                ok[name] += 1
            else:
                all_ok = False
                hallucinations.append(f"{ful}: поле «{name}» — «{jval}», "
                                      f"ожидалось «{eval_}»")
            if eval_ == ND and jval != ND:
                hallucinations.append(f"{ful}: выдумано значение поля «{name}»")
        # комментарий повторной отправки — на листе «Повторные отправки»
        ecomment = str(erow[5] or "").strip()
        if ecomment:
            comment_total += 1
            rrow = repeats.get(ful)
            if rrow is not None and str(rrow[4] or "").strip() == ecomment:
                comment_ok += 1
            else:
                all_ok = False
                hallucinations.append(f"{ful}: комментарий повтора не совпал")
        if all_ok:
            full_match += 1

    # форма Cloud: 12 колонок и порядок
    ws = res_wb["Новые строки Cloud"]
    headers = [c.value for c in ws[1]]
    form_ok = len(headers) == 12 and headers[0] == "Заказ" and headers[-1] == "Комментарий"
    orders = [str(r[0]).strip() for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
    dupes = len(orders) - len(set(orders))

    n = len(exp)
    print(f"ФУЛ в эталоне: {n}; полностью совпало: {full_match} ({full_match/n:.0%})")
    for name in total:
        print(f"  точность по «{name}»: {ok[name]}/{total[name]} "
              f"({ok[name]/max(total[name],1):.0%})")
    print(f"  комментарии повторов: {comment_ok}/{comment_total}")
    print(f"Форма Cloud (12 колонок, порядок): {'OK' if form_ok else 'FAIL'}; "
          f"дублей заказов в новых строках: {dupes}")
    print(f"Выдуманных значений (подстановок вместо «не определено»): "
          f"{len([h for h in hallucinations if 'выдумано' in h])}")
    if hallucinations:
        print("Расхождения:")
        for h in hallucinations:
            print("  -", h)
    return full_match == n and not hallucinations and form_ok and dupes == 0

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.stderr.write("usage: check_results.py <result.xlsx>\n")
        sys.exit(2)
    sys.exit(0 if main(sys.argv[1]) else 1)
