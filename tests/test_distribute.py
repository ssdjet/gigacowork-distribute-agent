# -*- coding: utf-8 -*-
"""Модульные тесты детерминированных правил distribute.py."""

import importlib.util
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "skill" / "scripts" / "distribute.py"
SPEC = importlib.util.spec_from_file_location("distribute_orders", SCRIPT)
d = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(d)


POSITIONS = {
    "avia": "Специалист по авиаперевозкам",
    "ved": "Специалист ВЭД",
    "multi": "Специалист по мультимодальным перевозкам",
}


def employee(name, group="ved", state="Работает", position=None, code="1"):
    return {
        "name": name,
        "state": state,
        "position": position or POSITIONS[group],
        "code": code,
    }


def ful(number="ФУЛ-1", order="S9000001", supplier="ООО Восток",
        delivery="sea", country="Китай", ful_date=None, brand="Бренд"):
    return {
        "number": number,
        "date": ful_date or datetime(2026, 10, 23, 12, 0),
        "cutoff": datetime(2026, 10, 30),
        "status": "новый",
        "comment": f"Номер:{order} от 23.10.2026" if order else "Без номера",
        "delivery": delivery,
        "brand": brand,
        "supplier": supplier,
        "country": country,
    }


def cloud(order="S8000001", supplier="ООО Восток", logist="Иванов",
          status="confirm", ulp_date=None):
    return {
        "order": order,
        "supplier": supplier,
        "brand": "Бренд",
        "logist": logist,
        "status": status,
        "conf_date": None,
        "cutoff": None,
        "ulp_date": ulp_date or datetime(2026, 10, 22),
        "country": "Китай",
        "delivery": "sea",
        "docs_date": None,
        "comment": "",
    }


def test_norm_removes_quotes_and_normalizes_whitespace():
    assert d.norm('  ООО «Восток»  "Север"  ') == "ооо восток север"
    assert d.norm(None) == ""


def test_norm_key_keeps_only_letters_and_digits():
    assert d.norm_key(' ООО "Восток-Север" ') == "ооовостоксевер"


@pytest.mark.parametrize("raw,expected", [
    ("22.10.2026", date(2026, 10, 22)),
    ("2026-10-22", date(2026, 10, 22)),
    ("22/10/2026", date(2026, 10, 22)),
    ("22.10.26", date(2026, 10, 22)),
    ("2026-10-22 15:30:00", date(2026, 10, 22)),
    (datetime(2026, 10, 22, 15, 30), date(2026, 10, 22)),
    ("не дата", None),
    ("", None),
    (None, None),
])
def test_as_date(raw, expected):
    assert d.as_date(raw) == expected


def test_match_headers_finds_reordered_columns_and_ignores_extra():
    cells = [
        "Лишняя колонка", " Страна ", "Дата и время ФУЛ", "Номер",
        "Комментарий", "Способ доставки", "Контрагент", "Бизнес единица",
        "Cutoff date", "Статус",
    ]
    mapping = d.match_headers(cells, d.FUL_COLS)
    assert mapping["country"] == 1
    assert mapping["date"] == 2
    assert mapping["number"] == 3
    assert mapping["comment"] == 4
    assert mapping["delivery"] == 5
    assert mapping["supplier"] == 6


def test_match_headers_does_not_use_confirmation_date_as_ful_date():
    mapping = d.match_headers(
        ["Заказ", "Дата подтверждения", "Поставщик"], d.CLOUD_COLS)
    assert mapping["conf_date"] == 1
    assert "date" not in mapping


def test_read_table_finds_headers_outside_first_row(tmp_path):
    path = tmp_path / "employees.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Произвольное имя"
    ws.append(["Выгрузка сотрудников"])
    ws.append(["Должность", "Код", "Наименование", "Текущее состояние", "Лишнее"])
    ws.append([POSITIONS["multi"], "001", "Иванов Иван Иванович", "Работает", 0])
    wb.save(path)

    rows = d.read_table(path, d.STAFF_COLS, d.MIN_COLS["staff"], "сотрудников")
    assert rows == [{
        "position": POSITIONS["multi"],
        "code": "001",
        "name": "Иванов Иван Иванович",
        "state": "Работает",
    }]


@pytest.mark.parametrize("comment,expected", [
    ("Номер:S0001234 от 01.10.2026", "S0001234"),
    ("номер：s9876543", "S9876543"),
    ("Произвольный комментарий S12345", "S12345"),
    ("Без номера", None),
    ("S123", None),
    (None, None),
])
def test_extract_order(comment, expected):
    assert d.extract_order(comment) == expected


@pytest.mark.parametrize("delivery,country,expected", [
    ("air", "Китай", "avia"),
    ("Авиа", None, "avia"),
    ("sea", "Китай", "ved"),
    ("truck", "Гонконг", "ved"),
    ("railway", "Макао", "ved"),
    ("sea", "Hong Kong", "ved"),
    ("sea", "Германия", "multi"),
    ("direct truck", "Турция", "multi"),
    (None, "Китай", None),
    ("sea", None, None),
])
def test_group_of(delivery, country, expected):
    actual, reason = d.group_of(delivery, country)
    assert actual == expected
    assert reason


@pytest.mark.parametrize("position,expected", [
    ("Специалист по авиаперевозкам", "avia"),
    ("Специалист ВЭД", "ved"),
    ("Специалист по мультимодальным перевозкам", "multi"),
    ("Менеджер", None),
])
def test_position_group(position, expected):
    assert d.position_group(position) == expected


@pytest.mark.parametrize("position,excluded", [
    ("Ассистент по авиаперевозкам", True),
    ("Руководитель группы ВЭД", True),
    ("Начальник отдела", True),
    ("Менеджер группы мультимодальных перевозок", True),
    ("Специалист ВЭД", False),
])
def test_is_distributor_flags_excluded_positions(position, excluded):
    assert d.is_distributor(position) is excluded


@pytest.mark.parametrize("state,expected", [
    ("Работает", True),
    ("работаю", True),
    ("active", True),
    ("Отпуск ежегодный", False),
    ("Уволен", False),
    ("Больничный", False),
    ("Командировка", False),
    ("Декрет", False),
    ("", False),
])
def test_is_active_accepts_only_explicitly_available_states(state, expected):
    assert d.is_active(state) is expected


@pytest.mark.parametrize("short_name,expected", [
    ("Иванов Иван Иванович", True),
    ("Иванов", True),
    ("Иванов И.И.", True),
    ("иванов и и", True),
    ("Иванов Иван", True),
    ("Иванов П.П.", False),
    ("И", False),
    ("Ивановский", False),
    ("Петров", False),
    (None, False),
])
def test_cloud_name_matches(short_name, expected):
    assert d.cloud_name_matches(short_name, "Иванов Иван Иванович") is expected


def test_match_logistician_rejects_ambiguous_surname():
    staff = [
        employee("Иванов Иван Иванович", code="1"),
        employee("Иванов Пётр Петрович", code="2"),
    ]
    assert d.match_logistician("Иванов", staff) is None
    assert d.match_logistician("Иванов И.И.", staff) == staff[0]
    assert d.match_logistician("Иванов Пётр Петрович", staff) == staff[1]


def test_history_uses_next_available_logistician():
    staff = [
        employee("Иванов Иван Иванович"),
        employee("Петров Пётр Петрович"),
        employee("Сидоров Сергей Сергеевич", state="Отпуск"),
    ]
    history = (
        [cloud(order=f"S700000{i}", logist="Сидоров") for i in range(3)]
        + [cloud(order=f"S700010{i}", logist="Петров") for i in range(2)]
        + [cloud(order="S7000200", logist="Иванов")]
    )
    results, _ = d.distribute([ful()], history, staff)
    assert results[0]["logist"] == "Петров Пётр Петрович"
    assert "2 зак." in results[0]["reason"]


def test_least_loaded_includes_current_batch_and_uses_name_tiebreak():
    staff = [
        employee("Алексеев Алексей Алексеевич"),
        employee("Борисов Борис Борисович"),
    ]
    history = [cloud(logist="Алексеев", status="в работе", supplier="Другой")]
    batch = [
        ful(number="ФУЛ-1", order="S9000001", supplier="Новый-1"),
        ful(number="ФУЛ-2", order="S9000002", supplier="Новый-2"),
    ]
    results, _ = d.distribute(batch, history, staff)
    assert [r["logist"] for r in results] == [
        "Борисов Борис Борисович",
        "Алексеев Алексей Алексеевич",
    ]


def test_supplier_normalization_preserves_single_logistician():
    staff = [
        employee("Алексеев Алексей Алексеевич"),
        employee("Борисов Борис Борисович"),
    ]
    batch = [
        ful(number="ФУЛ-1", order="S9000001", supplier='ООО "Восток"'),
        ful(number="ФУЛ-2", order="S9000002", supplier=" ооо  восток "),
    ]
    results, _ = d.distribute(batch, [], staff)
    assert results[0]["logist"] == results[1]["logist"]
    assert "один поставщик — один логист" in results[1]["reason"]


def test_repeat_duplicate_missing_group_and_summary():
    staff = [employee("Иванов Иван Иванович")]
    batch = [
        ful(number="ФУЛ-repeat", order="S8000001"),
        ful(number="ФУЛ-new", order="S9000001"),
        ful(number="ФУЛ-dup", order="S9000001"),
        ful(number="ФУЛ-no-group", order="S9000002", delivery=None),
    ]
    results, _ = d.distribute(batch, [cloud(order="S8000001")], staff)
    assert [r["status"] for r in results] == [
        "повторная отправка", "новый", "дубль в пачке", "новый",
    ]
    assert results[0]["comment"] == "Повторная отправка ФУЛа, 23.10.2026"
    assert results[2]["logist"] == results[1]["logist"]
    assert results[3]["group"] == d.NOT_DEFINED
    assert results[3]["logist"] == d.NOT_DEFINED
    assert d.summary(results) == {
        "всего_фул": 4,
        "новых_строк_cloud": 2,
        "повторных_отправок": 1,
        "дублей_в_пачке": 1,
        "возможных_пропусков": 0,
        "требует_внимания": 1,
        "без_логиста": 1,
    }


def test_possible_miss_uses_latest_cloud_date():
    staff = [employee("Иванов Иван Иванович")]
    batch = [
        ful(number="ФУЛ-old", order="S9000001", ful_date=datetime(2026, 10, 21)),
        ful(number="ФУЛ-new", order="S9000002", ful_date=datetime(2026, 10, 23)),
    ]
    results, last_date = d.distribute(batch, [cloud()], staff)
    assert last_date == date(2026, 10, 22)
    assert [r["miss"] for r in results] == [True, False]


def test_no_available_group_does_not_transfer_order():
    staff = [employee("Иванов Иван Иванович", group="avia", state="Уволен")]
    results, _ = d.distribute(
        [ful(delivery="air", country="Китай")], [], staff)
    assert results[0]["logist"] == d.NOT_DEFINED
    assert "нет доступных" in results[0]["attention"]


def test_missing_order_is_not_invented():
    staff = [employee("Иванов Иван Иванович")]
    results, _ = d.distribute([ful(order=None)], [], staff)
    assert results[0]["order"] is None
    assert "нет номера заказа" in results[0]["attention"]


def test_write_result_keeps_cloud_form_and_excludes_repeats(tmp_path):
    staff = [employee("Иванов Иван Иванович")]
    batch = [
        ful(number="ФУЛ-repeat", order="S8000001"),
        ful(number="ФУЛ-new", order="S9000001"),
    ]
    results, _ = d.distribute(batch, [cloud(order="S8000001")], staff)
    output = tmp_path / "result.xlsx"
    d.write_result(results, output)

    wb = load_workbook(output)
    ws = wb["Новые строки Cloud"]
    assert [cell.value for cell in ws[1]] == d.CLOUD_HEADERS
    assert len(d.CLOUD_HEADERS) == 12
    assert ws.max_row == 2
    assert ws.cell(2, 1).value == "S9000001"
    assert wb["Повторные отправки"].max_row == 2
