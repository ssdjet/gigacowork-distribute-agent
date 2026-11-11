#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
distribute.py — детерминированное распределение ФУЛ между логистами.

Все решения (группа, логист, статус ФУЛ) принимает этот скрипт, а не модель.
Модель только запускает скрипт и пересказывает его вывод.

Вход: три Excel-файла (.xlsx) — выгрузка ФУЛ из 1С, реестр Cloud, список сотрудников.
Колонки ищутся по заголовкам (вхождение нормализованного заголовка в список
синонимов), поэтому допустимы: другой порядок колонок, лишние колонки,
другое название листа. Первая строка, где найдено >= МИН_КОЛОНОК заголовков,
считается строкой заголовков.

Выход: result.xlsx в папке --out с листами:
  1. «Новые строки Cloud»  — 12 колонок строго в порядке реестра Cloud
  2. «Повторные отправки»  — заказы, уже есть в Cloud (+дубли внутри пачки)
  3. «Возможные пропуски»  — ФУЛ раньше последней записи Cloud, но заказа нет
  4. «Требует внимания»    — заказы, требующие решения человека
  5. «Журнал решений»      — основание по каждому ФУЛ

Также в stdout печатается JSON-итог для модели (краткий ответ в чате).

Зависимости: Python 3, openpyxl. Сеть и внешние сервисы не используются.
"""

import argparse
import io
import json
import re
import sys
import unicodedata
import zipfile
from datetime import datetime

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    sys.stderr.write("ERROR: требуется openpyxl\n")
    sys.exit(2)

NOT_DEFINED = "не определено"
REPEAT_PREFIX = "Повторная отправка ФУЛа"

# Фиксированные метаданные — чтобы result.xlsx был байт-в-байт
# воспроизводим между прогонами (требование «Воспроизводимость»).
FIXED_TS = datetime(2026, 1, 1, 0, 0, 0)

# ---------------------------------------------------------------------------
# Нормализация строк
# ---------------------------------------------------------------------------

def norm(s):
    """Нижний регистр, без пробелов по краям, схлопнутые пробелы, без кавычек."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("«", "").replace("»", "").replace('"', "'").replace("'", "")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def norm_key(s):
    """Ключ для сопоставления имён: без пробелов и пунктуации вообще."""
    return re.sub(r"[^0-9a-zа-яё]", "", norm(s))

def as_date(v):
    """Excel-значение -> datetime.date или None."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if hasattr(v, "date") and not isinstance(v, str):
        try:
            return v.date()
        except Exception:
            return None
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        try:
            return datetime.strptime(m.group(0), "%Y-%m-%d").date()
        except ValueError:
            return None
    return None

def fmt_date(d):
    return d.strftime("%d.%m.%Y") if d else NOT_DEFINED

# ---------------------------------------------------------------------------
# Сопоставление заголовков колонок
# ---------------------------------------------------------------------------

FUL_COLS = {
    "date":      ["дата", "дата и время фул", "дата фул"],
    "cutoff":    ["cutoff", "cut-off", "кутофф", "контрольная дата"],
    "number":    ["номер", "номер фул", "номер упаковочного листа", "№"],
    "status":    ["статус"],
    "comment":   ["комментарий", "комментарии"],
    "delivery":  ["способ доставки", "доставка", "способ перевозки"],
    "brand":     ["бизнес единица", "бренд", "бизнес-единица"],
    "logist":    ["логист"],
    "supplier":  ["контрагент", "поставщик"],
    "country":   ["страна", "страна отправления"],
}

CLOUD_COLS = {
    "order":      ["заказ", "номер заказа"],
    "supplier":   ["поставщик", "контрагент"],
    "brand":      ["бренд", "бизнес единица"],
    "logist":     ["логист"],
    "status":     ["статус"],
    "conf_date":  ["дата подтверждения"],
    "cutoff":     ["cut-off", "cutoff", "кутофф"],
    "ulp_date":   ["дата получения ул", "дата получения фул", "дата получения упаковочного листа"],
    "country":    ["страна отправления", "страна"],
    "delivery":   ["способ доставки"],
    "docs_date":  ["документы получены", "дата получения документов"],
    "comment":    ["комментарий", "комментарии"],
}

STAFF_COLS = {
    "name":    ["наименование", "фио", "сотрудник"],
    "state":   ["текущее состояние", "состояние", "статус"],
    "position":["должность"],
    "code":    ["код", "таб.no", "таб. номер", "табельный номер"],
}

MIN_COLS = {"ful": 5, "cloud": 8, "staff": 3}

def match_headers(cell_values, spec):
    """Вернуть {ключ: индекс_колонки} по списку заголовков строки."""
    mapping = {}
    used = set()
    norm_cells = [norm(c) for c in cell_values]
    for key, aliases in spec.items():
        for alias in aliases:
            na = norm(alias)
            for idx, nc in enumerate(norm_cells):
                if idx in used or not nc:
                    continue
                if nc == na or na in nc or nc in na:
                    # «дата» не должна перехватить «дата подтверждения»
                    if key == "date" and nc != na:
                        continue
                    mapping[key] = idx
                    used.add(idx)
                    break
            if key in mapping:
                break
    return mapping

def read_table(path, spec, min_cols, kind_label):
    """Найти лист и строку заголовков, вернуть (mapping, список dict-строк)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    best = None
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=min(15, ws.max_row or 15),
                                max_col=ws.max_column or 30, values_only=True):
            mapping = match_headers(row, spec)
            if len(mapping) >= min_cols:
                best = (ws.title, row[0] and ws.max_row or ws.max_row, mapping)
                break
        if best:
            break
    if not best:
        wb.close()
        raise ValueError(f"В файле '{path}' не найден лист с колонками {kind_label}")
    ws = wb[best[0]]
    header_row_idx = None
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=15,
                                             max_col=ws.max_column or 30,
                                             values_only=True), start=1):
        mapping = match_headers(row, spec)
        if len(mapping) >= min_cols:
            header_row_idx = r_idx
            break
    rows = []
    for row in ws.iter_rows(min_row=header_row_idx + 1,
                            max_col=ws.max_column or 30, values_only=True):
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        rec = {}
        for key, idx in mapping.items():
            rec[key] = row[idx] if idx < len(row) else None
        rows.append(rec)
    wb.close()
    return rows

# ---------------------------------------------------------------------------
# Извлечение номера заказа из комментария ФУЛ: «Номер:S0001234 от ДД.ММ.ГГГГ»
# ---------------------------------------------------------------------------

ORDER_RE = re.compile(r"н\w*м\w*р\w*[:：]?\s*(s\d{5,})")

def extract_order(comment):
    if comment is None:
        return None
    m = ORDER_RE.search(norm(comment).replace(" ", ""))
    if m:
        return m.group(1).upper()
    m = re.search(r"S\d{5,}", str(comment), re.IGNORECASE)
    return m.group(0).upper() if m else None

# ---------------------------------------------------------------------------
# Группы логистов
# ---------------------------------------------------------------------------

GROUPS = {
    "avia": "Авиаперевозки",
    "ved":  "ВЭД",
    "multi":"Мультимодальные перевозки",
}

CHINA = {"китай", "china", "кнр", "гонконг", "hongkong", "hong kong", "hk",
         "гонконг (сар)", "макао", "macau", "macao"}

def group_of(delivery, country):
    """Вернуть (ключ_группы или None, причина)."""
    d = norm(delivery)
    c = norm(country)
    if not d:
        return None, "нет способа доставки — группа не определена"
    if "air" in d or "авиа" in d:
        return "avia", "air → авиаперевозки"
    if not c:
        return None, "нет страны отправления — группа не определена"
    if c in CHINA or norm_key(c) in {norm_key(x) for x in CHINA}:
        return "ved", "не air + Китай/Гонконг → ВЭД"
    return "multi", "не air + не Китай → мультимодальные перевозки"

# ---------------------------------------------------------------------------
# Сотрудники
# ---------------------------------------------------------------------------

POSITION_GROUP = [
    ("avia",  ["авиаперевозк"]),
    ("ved",   ["вэд"]),
    ("multi", ["мультимодальн"]),
]

def position_group(position):
    p = norm(position)
    for gkey, marks in POSITION_GROUP:
        if any(m in p for m in marks):
            return gkey
    return None

def is_distributor(position):
    """Ассистенты и руководители заказы не распределяют."""
    p = norm(position)
    return any(w in p for w in ("ассистент", "руководитель", "начальник", "менеджер группы"))

def is_active(state):
    """Доступен только явно работающий сотрудник; прочие статусы — исключение."""
    s = norm(state)
    return s.startswith("работает") or s in ("работаю", "active", "в работе")

# ---------------------------------------------------------------------------
# Сопоставление имён: короткие имена Cloud («Иванов», «Иванов И.И.») с ФИО 1С
# ---------------------------------------------------------------------------

def name_tokens(fio):
    return [t for t in re.split(r"\s+", norm(fio)) if t]

def cloud_name_matches(name_cloud, fio):
    """Сопоставить полное ФИО, фамилию или фамилию с верными инициалами."""
    nc = norm(name_cloud)
    fio_n = norm(fio)
    if not nc or not fio_n:
        return False
    if nc == fio_n:
        return True

    fio_toks = name_tokens(fio)
    cloud_toks = name_tokens(name_cloud)
    if len(fio_toks) < 2 or not cloud_toks or cloud_toks[0] != fio_toks[0]:
        return False
    # Только фамилия. Неоднозначность однофамильцев разрешает match_logistician.
    if len(cloud_toks) == 1:
        return True
    # «Иванов Иван» или «Иванов Иван Иванович», но не просто «И».
    if fio_n.startswith(nc):
        return True

    # «Иванов И.И.» / «иванов и и»: инициалы должны совпасть полностью.
    cloud_initials = re.sub(r"[^0-9a-zа-яё]", "", "".join(cloud_toks[1:]))
    expected_initials = "".join(t[0] for t in fio_toks[1:] if t)
    return bool(cloud_initials) and cloud_initials == expected_initials

def match_logistician(cloud_name, staff):
    """Cloud-имя -> единственный подходящий сотрудник или None."""
    matches = [emp for emp in staff
               if cloud_name_matches(cloud_name, emp["name"])]
    if len(matches) == 1:
        return matches[0]
    # Полное совпадение ФИО имеет приоритет над совпадением только по фамилии.
    exact = [emp for emp in matches if norm(cloud_name) == norm(emp["name"])]
    return exact[0] if len(exact) == 1 else None

# ---------------------------------------------------------------------------
# Основной алгоритм
# ---------------------------------------------------------------------------

def distribute(ful_rows, cloud_rows, staff_rows):
    staff = []
    for r in staff_rows:
        staff.append({
            "name": str(r.get("name") or "").strip(),
            "state": str(r.get("state") or "").strip(),
            "position": str(r.get("position") or "").strip(),
            "code": str(r.get("code") or "").strip(),
            "gkey": position_group(r.get("position")),
            "distributor": is_distributor(r.get("position")),
            "active": is_active(r.get("state")),
        })

    def eligible(gkey):
        return [e for e in staff
                if e["gkey"] == gkey and e["active"] and not e["distributor"]]

    # История: число заказов по поставщику на каждого логиста (по всем строкам
    # Cloud, включая недоступных — фильтр доступности применяется позже).
    history = {}   # sup_key -> {emp_key: count}
    work_load = {} # emp_key -> число заказов «в работе»
    cloud_orders = set()
    last_cloud_date = None
    for r in cloud_rows:
        o = norm_key(r.get("order"))
        if o:
            cloud_orders.add(o)
        sup = norm_key(r.get("supplier"))
        emp = match_logistician(r.get("logist"), staff)
        if sup and emp:
            history.setdefault(sup, {})
            k = emp["name"]
            history[sup][k] = history[sup].get(k, 0) + 1
        if norm(r.get("status")) == "в работе" and emp:
            work_load[emp["name"]] = work_load.get(emp["name"], 0) + 1
        d = as_date(r.get("ulp_date")) or as_date(r.get("conf_date"))
        if d and (last_cloud_date is None or d > last_cloud_date):
            last_cloud_date = d

    def least_loaded(cands, extra_load):
        """Наименее загруженный: «в работе» в Cloud + уже назначено в пачке.
        При равенстве — по алфавиту ФИО (детерминированно)."""
        return sorted(cands,
                      key=lambda e: (work_load.get(e["name"], 0) + extra_load.get(e["name"], 0),
                                     e["name"]))[0]

    results = []       # по каждому ФУЛ
    seen_in_batch = {} # order_key -> индекс первого ФУЛ в пачке
    supplier_pick = {} # (sup_key, gkey) -> выбранный сотрудник
    extra_load = {}    # emp_key -> назначено в этой пачке

    for i, r in enumerate(ful_rows):
        ful_no = str(r.get("number") or f"ФУЛ-{i+1}").strip()
        ful_date = as_date(r.get("date"))
        order = extract_order(r.get("comment"))
        order_key = norm_key(order)
        sup_raw = str(r.get("supplier") or "").strip()
        sup_key = norm_key(sup_raw)
        brand = str(r.get("brand") or "").strip()
        delivery = str(r.get("delivery") or "").strip()
        country = str(r.get("country") or "").strip()
        cutoff = as_date(r.get("cutoff"))

        rec = {
            "ful": ful_no, "date": ful_date, "order": order, "supplier": sup_raw,
            "brand": brand, "delivery": delivery, "country": country,
            "cutoff": cutoff, "status": None, "group": None, "logist": None,
            "comment": "", "attention": "", "reason": "",
        }

        # --- повторная отправка / дубль ---
        if order_key and order_key in cloud_orders:
            rec["status"] = "повторная отправка"
            rec["comment"] = f"{REPEAT_PREFIX}, {fmt_date(ful_date)}"
            rec["reason"] = (f"заказ {order} уже есть в Cloud — новая строка "
                             f"не создаётся, в комментарий дописывается «{rec['comment']}»")
            results.append(rec)
            continue
        if order_key and order_key in seen_in_batch:
            first = results[seen_in_batch[order_key]]
            rec["status"] = "дубль в пачке"
            rec["group"] = first["group"]
            rec["logist"] = first["logist"]
            rec["comment"] = f"{REPEAT_PREFIX}, {fmt_date(ful_date)}"
            rec["reason"] = (f"заказ {order} уже встречался в этой пачке "
                             f"({first['ful']}) — дубль, строка не создаётся")
            results.append(rec)
            continue
        if order_key:
            seen_in_batch[order_key] = len(results)

        # --- возможный пропуск ---
        miss = bool(ful_date and last_cloud_date and ful_date < last_cloud_date)

        # --- группа ---
        gkey, why_group = group_of(delivery, country)
        rec["group"] = GROUPS.get(gkey, NOT_DEFINED)
        rec["status"] = "новый"

        # --- логист ---
        if gkey is None:
            rec["logist"] = NOT_DEFINED
            rec["attention"] = why_group
            rec["reason"] = f"{why_group}; логист не назначен"
            rec["miss"] = miss
            results.append(rec)
            continue

        cands = eligible(gkey)
        if not cands:
            rec["logist"] = NOT_DEFINED
            rec["attention"] = (f"в группе «{rec['group']}» нет доступных "
                                f"логистов (отпуска/увольнения/не распределяют)")
            rec["reason"] = rec["attention"]
            rec["miss"] = miss
            results.append(rec)
            continue

        pick_key = (sup_key, gkey)
        if pick_key in supplier_pick:
            pick = supplier_pick[pick_key]
            rec["logist"] = pick["name"]
            rec["reason"] = (f"один поставщик — один логист: закреплён {pick['name']} "
                             f"(выбран для {sup_raw or 'поставщика'} ранее в пачке)")
        else:
            hist = history.get(sup_key, {})
            ranked = sorted(hist.items(), key=lambda kv: (-kv[1], kv[0]))
            available = {e["name"] for e in cands}
            chosen = None
            skipped = []
            for name, cnt in ranked:
                if name in available:
                    chosen = name
                    rec["reason"] = (f"история Cloud: {name} чаще всех ведёт "
                                     f"поставщика ({cnt} зак.)")
                    break
                skipped.append(f"{name} ({cnt} зак., недоступен)")
            if chosen is None:
                emp = least_loaded(cands, extra_load)
                chosen = emp["name"]
                if ranked:
                    rec["reason"] = (f"предпочтительные по истории недоступны "
                                     f"({'; '.join(skipped)}) → наименее загруженный "
                                     f"в группе по «в работе»")
                else:
                    rec["reason"] = ("новый поставщик (нет истории в Cloud) → "
                                     "наименее загруженный в группе по «в работе»")
            pick = next(e for e in staff if e["name"] == chosen)
            supplier_pick[pick_key] = pick
            rec["logist"] = chosen
        extra_load[chosen] = extra_load.get(chosen, 0) + 1
        if not order:
            rec["attention"] = "в ФУЛ нет номера заказа — заполнить вручную"
            rec["reason"] += "; номер заказа не определён"
        rec["miss"] = miss
        results.append(rec)

    return results, last_cloud_date

# ---------------------------------------------------------------------------
# Выгрузка result.xlsx
# ---------------------------------------------------------------------------

CLOUD_HEADERS = ["Заказ", "Поставщик", "Бренд", "Логист", "Статус",
                 "Дата подтверждения", "Cut-off", "Дата получения УЛ",
                 "Страна отправления", "Способ доставки",
                 "Документы получены", "Комментарий"]

def write_result(results, out_path):
    wb = Workbook()
    wb.properties.creator = "distribute-agent"
    wb.properties.created = FIXED_TS
    wb.properties.modified = FIXED_TS

    ws = wb.active
    ws.title = "Новые строки Cloud"
    ws.append(CLOUD_HEADERS)
    for r in results:
        if r["status"] != "новый":
            continue
        ws.append([
            r["order"] or NOT_DEFINED, r["supplier"] or NOT_DEFINED,
            r["brand"] or NOT_DEFINED, r["logist"] or NOT_DEFINED,
            "в работе", NOT_DEFINED,
            r["cutoff"].strftime("%d.%m.%Y") if r["cutoff"] else NOT_DEFINED,
            r["date"].strftime("%d.%m.%Y") if r["date"] else NOT_DEFINED,
            r["country"] or NOT_DEFINED, r["delivery"] or NOT_DEFINED,
            NOT_DEFINED, "",
        ])

    ws = wb.create_sheet("Повторные отправки")
    ws.append(["Номер ФУЛ", "Дата ФУЛ", "Заказ", "Поставщик", "Комментарий для Cloud"])
    for r in results:
        if r["status"] in ("повторная отправка", "дубль в пачке"):
            ws.append([r["ful"], fmt_date(r["date"]), r["order"] or NOT_DEFINED,
                       r["supplier"] or NOT_DEFINED, r["comment"]])

    ws = wb.create_sheet("Возможные пропуски")
    ws.append(["Номер ФУЛ", "Дата ФУЛ", "Заказ", "Поставщик", "Пояснение"])
    for r in results:
        if r.get("miss"):
            ws.append([r["ful"], fmt_date(r["date"]), r["order"] or NOT_DEFINED,
                       r["supplier"] or NOT_DEFINED,
                       "ФУЛ раньше последней записи Cloud, но заказа нет в реестре — "
                       "проверить вручную, не потерян ли он"])

    ws = wb.create_sheet("Требует внимания")
    ws.append(["Номер ФУЛ", "Заказ", "Проблема", "Что сделал агент"])
    for r in results:
        if r["attention"]:
            ws.append([r["ful"], r["order"] or NOT_DEFINED, r["attention"],
                       "строка создана без спорных данных, логист: "
                       + (r["logist"] or NOT_DEFINED)])

    ws = wb.create_sheet("Журнал решений")
    ws.append(["Номер ФУЛ", "Дата ФУЛ", "Заказ", "Поставщик", "Статус ФУЛ",
               "Группа", "Логист", "Основание решения"])
    for r in results:
        ws.append([r["ful"], fmt_date(r["date"]), r["order"] or NOT_DEFINED,
                   r["supplier"] or NOT_DEFINED, r["status"],
                   r["group"] or "", r["logist"] or "",
                   r["reason"]])

    # Сохраняем через буфер и переписываем zip с фиксированными метками
    # времени — иначе байты файла зависят от времени запуска и хэш
    # результата различается между прогонами (требование воспроизводимости).
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    with zipfile.ZipFile(buf) as zin, \
         zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                # openpyxl ставит modified=now при сохранении — фиксируем
                data = re.sub(
                    rb"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)",
                    rb"\g<1>2026-01-01T00:00:00Z\g<2>", data)
            ni = zipfile.ZipInfo(item.filename, date_time=(2026, 1, 1, 0, 0, 0))
            ni.compress_type = zipfile.ZIP_DEFLATED
            ni.external_attr = item.external_attr
            zout.writestr(ni, data)

def summary(results):
    new = [r for r in results if r["status"] == "новый"]
    rep = [r for r in results if r["status"] == "повторная отправка"]
    dup = [r for r in results if r["status"] == "дубль в пачке"]
    miss = [r for r in results if r.get("miss")]
    att = [r for r in results if r["attention"]]
    no_log = [r for r in new if not r["logist"] or r["logist"] == NOT_DEFINED]
    return {
        "всего_фул": len(results),
        "новых_строк_cloud": len(new),
        "повторных_отправок": len(rep),
        "дублей_в_пачке": len(dup),
        "возможных_пропусков": len(miss),
        "требует_внимания": len(att),
        "без_логиста": len(no_log),
    }

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def detect_role(path):
    """Определить тип файла по найденным заголовкам (для устойчивости
    к порядку приложенных файлов)."""
    scores = {}
    for role, spec, mn in (("ful", FUL_COLS, MIN_COLS["ful"]),
                           ("cloud", CLOUD_COLS, MIN_COLS["cloud"]),
                           ("staff", STAFF_COLS, MIN_COLS["staff"])):
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            best = 0
            for ws in wb.worksheets:
                for row in ws.iter_rows(min_row=1, max_row=10,
                                        max_col=ws.max_column or 30,
                                        values_only=True):
                    best = max(best, len(match_headers(row, spec)))
            wb.close()
            if best >= mn:
                scores[role] = best
        except Exception:
            continue
    if not scores:
        return None
    return max(scores, key=scores.get)

def main():
    ap = argparse.ArgumentParser(description="Распределение ФУЛ между логистами")
    ap.add_argument("--ful", help="путь к выгрузке ФУЛ из 1С (.xlsx)")
    ap.add_argument("--cloud", help="путь к реестру Cloud (.xlsx)")
    ap.add_argument("--staff", help="путь к списку сотрудников из 1С (.xlsx)")
    ap.add_argument("--files", nargs="*", default=[],
                    help="три файла без ролей — роли определятся по заголовкам")
    ap.add_argument("--out", required=True, help="папка для result.xlsx")
    args = ap.parse_args()

    roles = {"ful": args.ful, "cloud": args.cloud, "staff": args.staff}
    for p in args.files:
        role = detect_role(p)
        if role and roles[role] is None:
            roles[role] = p
    missing = [k for k, v in roles.items() if not v]
    if missing:
        sys.stderr.write("ERROR: не переданы файлы: %s\n" % ", ".join(missing))
        sys.exit(2)

    ful_rows = read_table(roles["ful"], FUL_COLS, MIN_COLS["ful"], "ФУЛ")
    cloud_rows = read_table(roles["cloud"], CLOUD_COLS, MIN_COLS["cloud"], "Cloud")
    staff_rows = read_table(roles["staff"], STAFF_COLS, MIN_COLS["staff"], "сотрудников")

    results, last_cloud = distribute(ful_rows, cloud_rows, staff_rows)

    import os
    os.makedirs(args.out, exist_ok=True)
    out_path = os.path.join(args.out, "result.xlsx")
    write_result(results, out_path)

    s = summary(results)
    s["файл"] = out_path
    s["последняя_запись_cloud"] = fmt_date(last_cloud)
    print(json.dumps(s, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
