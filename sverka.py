#!/usr/bin/env python3
"""
Сверка выручки по портфелю проектов (по направлению) за полугодие:
  Сводный отчёт по портфелю  <->  Паспорта проектов  <->  Выгрузка выручки из бухгалтерии

Запуск:
    python3 sverka.py [--config config.json]

Все параметры (пути к файлам, направление, полугодие, допуски) — в config.json,
рядом со скриптом. Правьте config.json для сверки другого полугодия/направления/года,
сам скрипт трогать не нужно.
"""
import argparse
import json
import os
import re
import unicodedata
import warnings

from xlsx_min import Workbook, row_get, to_float, write_workbook, SheetSpec, col_index_to_letters
from xlsx_min import (
    STYLE_NUMBER, STYLE_NUMBER_OK, STYLE_NUMBER_MISMATCH,
    STYLE_WARN, STYLE_GRAY,
)

warnings.filterwarnings("ignore")


# ---------------------------------------------------------------------------
# Период сверки: h1 (1е полугодие) / h2 (2е полугодие) / year (весь год)
# ---------------------------------------------------------------------------

# Метки блоков в Сводном отчёте (строка над "Прогноз"/"Факт") — четыре
# одинаковых по составу блока колонок идут подряд: H1, H2, весь год, весь
# срок проекта. Нам нужны только первые три.
PERIOD_LABEL_TEMPLATES = {
    "h1": "{year}г - 1е ПОЛУГОДИЕ",
    "h2": "{year}г - 2е ПОЛУГОДИЕ",
    "year": "ПО ВСЕМУ {year}г году",
}

# Метки строк в паспорте (лист "4. Факт БДР") под блоком года. None для "year"
# означает "не искать под-строку, взять саму строку года" — см.
# read_passport_h1_actual().
PERIOD_HALF_LABELS = {
    "h1": "1е полугодие",
    "h2": "2е полугодие",
    "year": None,
}


def period_half_label(config):
    return PERIOD_HALF_LABELS[config["period"]]


def resolve_path(disk_root, path):
    """Абсолютные пути (например, файл лежит рядом со скриптом, а не на
    примонтированном диске Disk-O) используются как есть; относительные —
    считаются от disk_root, как раньше."""
    return path if os.path.isabs(path) else os.path.join(disk_root, path)


# ---------------------------------------------------------------------------
# Нормализация кода проекта
# ---------------------------------------------------------------------------

def normalize_code(raw):
    """
    Приводит код проекта к единому виду для сопоставления между источниками.
    Примеры: 'MFB0-030' / 'MFBO_030 ' / 'mfbo 030' -> 'MFBO_030'
    Цифра 0 -> буква O только в буквенном префиксе (до первого разделителя),
    в числовом "хвосте" код 0 остаётся цифрой.
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("нет кода", "-", "—"):
        return None
    s = s.upper()
    s = re.sub(r"[\s\-]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    if not s:
        return None
    parts = s.split("_", 1)
    prefix = parts[0].replace("0", "O")
    rest = parts[1] if len(parts) > 1 else ""
    return prefix + ("_" + rest if rest else "")


# ---------------------------------------------------------------------------
# 1. Сводный отчёт по портфелю
# ---------------------------------------------------------------------------

def load_summary(config, warnings_out):
    path = resolve_path(config["disk_root"], config["summary_file"])
    wb = Workbook(path)
    rows = wb.read_sheet(config["summary_sheet"])

    header_row_num = None
    header = None
    for row_num, row in rows:
        if any(isinstance(v, str) and v.strip().startswith("Код проекта") for v in row.values()):
            header_row_num = row_num
            header = row
            break
    if header is None:
        raise RuntimeError("Не найдена строка заголовка в Сводном отчёте (ищу ячейку 'Код проекта...')")

    def find_col(predicate, start=0):
        for idx in sorted(header):
            if idx < start:
                continue
            v = header[idx]
            if isinstance(v, str) and predicate(v):
                return idx
        return None

    col_dp = find_col(lambda v: v.strip() == "ДП")
    col_proekt = find_col(lambda v: v.strip() == "Проекты")
    col_napravlenie = find_col(lambda v: v.strip().startswith("Направление"))
    col_kod = find_col(lambda v: v.strip().startswith("Код проекта"))
    col_rp = find_col(lambda v: v.strip() == "РП")

    # Над строкой заголовка ("Прогноз"/"Факт") есть строка с метками периода
    # блока (H1/H2/весь год/весь срок проекта) — четыре одинаковых по составу
    # блока колонок идут подряд, и нужный "Прогноз"/"Факт" ищем именно в блоке
    # с нужной меткой периода, а не просто "первый попавшийся" (это всегда
    # был бы H1).
    period_row = dict(rows).get(header_row_num - 2, {})
    period_label = PERIOD_LABEL_TEMPLATES[config["period"]].format(year=config["year"])
    col_prognoz = None
    for idx in sorted(period_row):
        v = period_row[idx]
        if isinstance(v, str) and v.strip() == period_label:
            col_prognoz = idx
            break
    if col_prognoz is None:
        available = sorted(v for v in period_row.values() if isinstance(v, str) and v.strip())
        raise RuntimeError(
            f"Сводный: не нашёл блок периода '{period_label}' (config['period']='{config['period']}'). "
            f"Найденные метки периодов в файле: {available}"
        )
    col_fact = col_prognoz + 1
    if not (isinstance(header.get(col_fact), str) and header[col_fact].strip() == "Факт"):
        warnings_out.append(
            f"Сводный: колонка сразу после 'Прогноз' в блоке '{period_label}' не называется "
            f"'Факт' — проверьте структуру листа вручную."
        )
    col_status = find_col(lambda v: v.strip() == "Статус проекта")

    for needed, name in [
        (col_proekt, "Проекты"), (col_napravlenie, "Направление"),
        (col_kod, "Код проекта"), (col_rp, "РП"), (col_fact, f"Факт ({period_label})"),
    ]:
        if needed is None:
            raise RuntimeError(f"Сводный: не нашёл колонку '{name}' в заголовке")

    # direction=None/""/"*" в конфиге означает "все направления" — не фильтровать
    # по колонке "Направление" вообще (в т.ч. строки, где там мусор — имя человека
    # вместо направления, пустая ячейка и т.п. — такие строки сейчас незаметно
    # выпадали из ЛЮБОЙ точечной сверки по направлению, хотя это реальные проекты
    # с реальным кодом).
    direction_raw = config.get("direction")
    direction_wanted = direction_raw.strip().lower() if direction_raw and direction_raw != "*" else None

    # code_prefixes=[...] — альтернатива/дополнение к фильтру по "Направление":
    # берём только строки, чей код проекта начинается с одного из префиксов.
    # Нужно там, где один "Направление" в Сводном фактически ведут несколько
    # разных людей/ЦФО (например "VK Cloud" = DVPO+GVCL+MRCA+MRCO у разных РП) —
    # тогда фильтр по коду точнее, чем по тексту направления.
    code_prefixes_raw = config.get("code_prefixes")
    code_prefixes = None
    if code_prefixes_raw and code_prefixes_raw != "*":
        code_prefixes = [p.upper() for p in code_prefixes_raw]

    out = []
    for row_num, row in rows:
        if row_num <= header_row_num:
            continue
        proekt = row.get(col_proekt)
        if not proekt or not isinstance(proekt, str) or not proekt.strip():
            continue
        napravlenie = row.get(col_napravlenie)
        if direction_wanted is not None:
            if not isinstance(napravlenie, str) or napravlenie.strip().lower() != direction_wanted:
                continue
        dp = row.get(col_dp) if col_dp is not None else None
        kod_raw = row.get(col_kod)
        kod_norm = normalize_code(kod_raw)
        if code_prefixes is not None:
            if not kod_norm or not any(kod_norm.startswith(p) for p in code_prefixes):
                continue
        rp = row.get(col_rp)
        status = row.get(col_status) if col_status is not None else None
        fact_h1 = to_float(row.get(col_fact), default=0.0)
        out.append({
            "row_num": row_num,
            "dp": dp.strip() if isinstance(dp, str) else dp,
            "proekt": proekt.strip(),
            "napravlenie": napravlenie.strip() if isinstance(napravlenie, str) else napravlenie,
            "kod_raw": kod_raw,
            "kod_norm": kod_norm,
            "rp": rp.strip() if isinstance(rp, str) else rp,
            "status": status,
            "fact_h1": fact_h1,
        })
    return out


def group_summary(summary_rows, warnings_out):
    """
    Группирует строки Сводного по нормализованному коду для сверки.
    Строки с ДП == 'Вычет' исключаются из факта H1 (это корректировки, не отдельные
    проекты), но их сумма по каждому коду сохраняется отдельно в поле "vychet_sum"
    группы — чтобы показать её в отчёте, а не просто молча выбросить.
    Строки без кода остаются отдельными записями (сверяются только с паспортом).
    Возвращает список групп:
    {kod_norm, proekts:[...], rp_set, fact_h1_sum, vychet_sum, rows:[...]}
    """
    excluded_vychet = [r for r in summary_rows if isinstance(r["dp"], str) and r["dp"].strip() == "Вычет"]
    kept = [r for r in summary_rows if r not in excluded_vychet]
    if excluded_vychet:
        warnings_out.append(
            f"Сводный: исключено {len(excluded_vychet)} строк с ДП='Вычет' "
            f"({', '.join(r['proekt'] for r in excluded_vychet)})"
        )

    vychet_by_code = {}
    for r in excluded_vychet:
        if r["kod_norm"] is not None:
            vychet_by_code[r["kod_norm"]] = vychet_by_code.get(r["kod_norm"], 0.0) + r["fact_h1"]

    groups = {}
    no_code_items = []
    for r in kept:
        if r["kod_norm"] is None:
            no_code_items.append(r)
            continue
        g = groups.setdefault(r["kod_norm"], {
            "kod_norm": r["kod_norm"], "proekts": [], "kod_raws": set(),
            "rp_set": set(), "fact_h1_sum": 0.0, "rows": [], "status_set": set(),
            "napravlenie_set": set(),
            "vychet_sum": vychet_by_code.get(r["kod_norm"], 0.0),
        })
        g["proekts"].append(r["proekt"])
        g["kod_raws"].add(str(r["kod_raw"]))
        if r["rp"]:
            g["rp_set"].add(r["rp"])
        if r["status"]:
            g["status_set"].add(str(r["status"]))
        if r.get("napravlenie"):
            g["napravlenie_set"].add(r["napravlenie"])
        g["fact_h1_sum"] += r["fact_h1"]
        g["rows"].append(r)

    for r in no_code_items:
        groups[f"__NOCODE__{r['row_num']}"] = {
            "kod_norm": None, "proekts": [r["proekt"]], "kod_raws": {str(r["kod_raw"])},
            "rp_set": {r["rp"]} if r["rp"] else set(), "fact_h1_sum": r["fact_h1"],
            "vychet_sum": 0.0,
            "napravlenie_set": {r["napravlenie"]} if r.get("napravlenie") else set(),
            "rows": [r], "status_set": {str(r["status"])} if r["status"] else set(),
        }

    return list(groups.values())


# ---------------------------------------------------------------------------
# 2. Паспорта проектов
# ---------------------------------------------------------------------------

def _find_row_by_first_col(rows, target_text):
    for row_num, row in rows:
        v = row.get(0)
        if isinstance(v, str) and v.strip() == target_text:
            return row_num, row
    return None, None


def read_passport_code(path):
    kod_raw, _ = read_passport_code_and_project(path)
    return kod_raw


def read_passport_code_and_project(path):
    """Возвращает (Код проекта, ПРОЕКТ) из листа "3. Договоры_Этапы_Доходы".
    Название проекта нужно, чтобы отличить настоящие дубликаты одного паспорта
    (одинаковый ПРОЕКТ, разные версии/копии файла) от ситуации, когда один код
    честно делят несколько разных под-проектов (разный ПРОЕКТ у каждого) —
    см. select_passport()."""
    with Workbook(path) as wb:
        if "3. Договоры_Этапы_Доходы" not in wb.sheet_names:
            return None, None
        rows = wb.read_sheet("3. Договоры_Этапы_Доходы")
        _, kod_row = _find_row_by_first_col(rows, "Код проекта")
        _, proekt_row = _find_row_by_first_col(rows, "ПРОЕКТ")
        kod_raw = kod_row.get(1) if kod_row is not None else None
        proekt = proekt_row.get(1) if proekt_row is not None else None
        return kod_raw, proekt


def read_passport_h1_actual(path, year_label, half_label):
    """
    На листе "4. Факт БДР" таблица "Доп.аналитика по полугодиям" повторяется
    для каждого года проекта (2025, 2026, 2027, ...), и метка "1е полугодие"
    в колонке A встречается по разу под каждым годом. Поэтому сначала находим
    строку нужного года, и "1е полугодие" ищем только после неё (до следующей
    строки-года) — иначе можно случайно взять данные не того года.

    half_label=None означает "весь год целиком" — тогда значение берётся
    прямо из самой строки года (col2 у неё — Фактические Доходы за год),
    а не из под-строки полугодия под ней.
    """
    with Workbook(path) as wb:
        if "4. Факт БДР" not in wb.sheet_names:
            return None, "нет листа '4. Факт БДР'"
        rows = wb.read_sheet("4. Факт БДР")
        year_row = None
        year_row_num = None
        for row_num, row in rows:
            v = row.get(0)
            if isinstance(v, str) and v.strip() == str(year_label):
                year_row_num = row_num
                year_row = row
                break
        if year_row_num is None:
            return None, f"нет блока года '{year_label}'"
        if half_label is None:
            # col0 = метка года, col1 = Прогнозные (Доходы), col2 = Фактические (Доходы) — за весь год
            return to_float(year_row.get(2), default=0.0), None
        for row_num, row in rows:
            if row_num <= year_row_num:
                continue
            v = row.get(0)
            if isinstance(v, str) and v.strip() == half_label:
                # col0 = метка, col1 = Прогнозные (Доходы), col2 = Фактические (Доходы)
                return to_float(row.get(2), default=0.0), None
            if isinstance(v, str) and re.fullmatch(r"\d{4}", v.strip()):
                break  # начался блок следующего года, а нужную строку не нашли
        return None, f"нет строки '{half_label}' в блоке года '{year_label}'"


ACTIVE_KEYWORDS = ("действ", "текущ")
COMPLETED_KEYWORDS = ("заверш", "закрыт")


def _folder_kind(name, exclude_kw):
    # macOS/APFS отдаёт имена файлов в NFD (Й = И + отдельный значок), а строки-константы
    # в коде обычно NFC — без нормализации сравнение "действ" со строкой из os.listdir
    # молча не совпадает ни на одной букве с "й".
    n = unicodedata.normalize("NFC", name).lower()
    if any(k in n for k in exclude_kw):
        return "exclude"
    if any(k in n for k in ACTIVE_KEYWORDS):
        return "active"
    if any(k in n for k in COMPLETED_KEYWORDS):
        return "completed"
    return "unknown"


def find_passport_files(config, warnings_out):
    """
    Сканирует папки РП и строит индекс: нормализованный код ->
    {"active": [{"path","mtime","kod_raw","proekt"}, ...], "completed": [...]}.
    Названия папок 2-го уровня (Действующие/Завершенные/Закрытые/Текущие проекты
    и т.п. — у разных РП называются по-разному) распознаются по ключевым словам,
    не по точному совпадению. Потенциальные/Переданные/Архив — исключаются.
    Сопоставление проекту — по содержимому файла ('Код проекта'), не по имени.
    Разделение на active/completed нужно, т.к. один и тот же код проекта может
    встречаться и в текущем паспорте, и в старом архивном файле прошлых лет
    (лежащем в "Закрытые проекты") — брать "просто самый свежий по mtime" файл
    ненадёжно: старые файлы иногда трогают позже, чем обновляют актуальные.
    """
    root = resolve_path(config["disk_root"], config["pm_folders_root"])
    exclude_kw = config["passport_folder_keywords_exclude"]

    candidates = []  # (path, tier)
    try:
        pm_dirs = [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]
    except OSError as e:
        raise RuntimeError(f"Не смог открыть папку РП: {root} ({e})")

    for pm in pm_dirs:
        pm_path = os.path.join(root, pm)
        try:
            sub_dirs = [d for d in os.listdir(pm_path) if os.path.isdir(os.path.join(pm_path, d))]
        except OSError as e:
            warnings_out.append(f"Паспорт: не смог открыть папку РП '{pm_path}': {e}")
            continue
        for sub in sub_dirs:
            tier = _folder_kind(sub, exclude_kw)
            if tier not in ("active", "completed"):
                continue
            sub_path = os.path.join(pm_path, sub)
            for dirpath, dirnames, filenames in os.walk(sub_path):
                dirnames[:] = [
                    d for d in dirnames
                    if _folder_kind(d, exclude_kw) != "exclude"
                ]
                for fn in filenames:
                    if fn.lower().endswith(".xlsx") and not fn.startswith("~$") and not fn.startswith("._"):
                        candidates.append((os.path.join(dirpath, fn), tier))

    index = {}  # kod_norm -> {"active": [...], "completed": [...]}
    for path, tier in candidates:
        try:
            kod_raw, proekt = read_passport_code_and_project(path)
        except Exception as e:
            warnings_out.append(f"Паспорт: не смог прочитать код из '{path}': {e}")
            continue
        kod_norm = normalize_code(kod_raw)
        if kod_norm is None:
            continue
        mtime = os.path.getmtime(path)
        entry = index.setdefault(kod_norm, {"active": [], "completed": []})
        entry[tier].append({"path": path, "mtime": mtime, "kod_raw": kod_raw, "proekt": proekt})

    return index


def _normalize_proekt_name(name):
    if not name:
        return None
    return re.sub(r"\s+", " ", str(name).strip().lower())


def select_passport(entry, kod_norm, year_label, half_label, warnings_out):
    """
    Возвращает суммарный факт H1 по коду проекта из паспортов и список файлов,
    из которых он собран: {"value": float, "paths": [...], "errors": [...]}.
    None, если по коду вообще не нашлось ни одного паспорта.

    Кандидаты берутся из ОБЕИХ групп — и "действующие", и "завершённые" — вместе:
    у одного кода может быть несколько по-настоящему разных под-проектов, и они
    не обязаны все лежать в одной и той же группе (например, у FINO-009 три
    разных паспорта — два в "Действующих" у разных РП и один в "Завершённых",
    и Сводный суммирует все три; искать только в одной группе занижало сумму).

    Кандидаты группируются по названию "ПРОЕКТ" из самого файла (нормализованному).
    Одинаковый ПРОЕКТ = разные версии/копии одного и того же паспорта — из них
    выбирается один: с ненулевым фактом за нужное полугодие (предпочтительно),
    и самый свежий по дате изменения среди таких. Разный ПРОЕКТ при одинаковом
    коде = реально разные под-проекты, которые в Сводном тоже суммируются под
    одним кодом (см. group_summary) — поэтому их выбранные значения складываются,
    а не выбираются "или-или".

    Почему внутри бакета не просто "самый свежий файл": при передаче проекта
    другому РП в его папке иногда создают свежую копию паспорта, которая ещё не
    заполнена (факт = 0), пока старая копия у прежнего РП всё ещё содержит
    реальные данные за прошедшие месяцы — голый mtime в этом случае выбрал бы
    пустышку.

    Это эвристика по имени, не гарантия: если проект в течение года
    переименовали, разные версии его паспорта могут ошибочно попасть в разные
    бакеты и просуммироваться как "два под-проекта". Каждое такое суммирование
    явно попадает в warnings_out — при подозрении на завышенное расхождение
    сначала проверяй эти строки.
    """
    if entry is None:
        return None
    cands = (entry.get("active") or []) + (entry.get("completed") or [])
    if not cands:
        return None

    buckets = {}
    for c in cands:
        key = _normalize_proekt_name(c["proekt"]) or c["path"]  # без имени — считаем уникальным
        buckets.setdefault(key, []).append(c)

    total = 0.0
    paths = []
    errors = []
    for key, group in buckets.items():
        read = []
        for c in group:
            val, err = read_passport_h1_actual(c["path"], year_label, half_label)
            read.append({**c, "value": val, "error": err})
        nonzero = [r for r in read if r["value"] not in (None, 0.0)]
        pool = nonzero or read
        best = max(pool, key=lambda r: r["mtime"])
        for r in read:
            if r["path"] == best["path"]:
                continue
            reason = "нулевое/непрочитанное значение" if nonzero and r not in nonzero else "не самый свежий"
            warnings_out.append(
                f"Паспорт {kod_norm} ('{best.get('proekt') or '?'}'): несколько файлов, "
                f"выбран '{best['path']}' ({reason} у остальных); проигнорирован '{r['path']}'"
            )
        paths.append(best["path"])
        if best["error"]:
            errors.append(f"{best.get('proekt') or kod_norm}: {best['error']}")
        else:
            total += best["value"]

    if len(buckets) > 1:
        proekt_names = sorted({(group[0].get("proekt") or "?") for group in buckets.values()})
        warnings_out.append(
            f"Паспорт {kod_norm}: код делят {len(buckets)} разных под-проекта "
            f"({', '.join(proekt_names)}) — их H1-факты из паспортов суммированы."
        )

    value = None if len(errors) == len(buckets) else total
    return {"value": value, "paths": paths, "errors": errors}


# ---------------------------------------------------------------------------
# 3. Выгрузка выручки из бухгалтерии
# ---------------------------------------------------------------------------

def parse_revenue_extract(config, warnings_out):
    """Диспетчер по формату выгрузки (config['revenue_format'], по умолчанию 'pivot'
    — для обратной совместимости со старыми конфигами без этого поля)."""
    fmt = config.get("revenue_format", "pivot")
    if fmt == "journal":
        return parse_revenue_extract_journal(config, warnings_out)
    if fmt == "pivot":
        return parse_revenue_extract_pivot(config, warnings_out)
    raise RuntimeError(f"Неизвестный revenue_format в конфиге: {fmt!r} (ожидается 'pivot' или 'journal')")


def parse_revenue_extract_journal(config, warnings_out):
    """
    Формат "журнал проводок" — каждая строка уже сама по себе конкретная
    транзакция с готовыми Код проекта / Проект / Группа статей (=вид выручки) /
    Номер и дата договора / Сумма — в отличие от pivot-формата, тут нет
    "схлопнутых" повторов и Итог-строк, каждая колонка заполнена в каждой
    строке. Найден на примере файла бухгалтерии "FinOps 7 мес 2026.xlsx".

    Суммовая колонка — "Сумма без НДС с плюсом (для бизнеса)": проверено, что
    она равна "Сумма без НДС" с обратным знаком (та по умолчанию отрицательная
    для доходных проводок) — то есть это тот же факт выручки, но в удобном для
    отчётности положительном виде; остальные суммовые колонки ("Сумма",
    "Сумма без НДС") не согласованы друг с другом по знаку и не подходят.

    Период не фильтруется скриптом — берётся весь период, покрытый файлом
    (для 7-месячного файла это и есть Январь-Июль, как договорился с
    бухгалтерией; если пришлют файл за другой период, просто укажи его в
    config['revenue_file'] — сравнивать нужно c тем же периодом из Сводного/
    паспортов через config['period']).
    """
    path = resolve_path(config["disk_root"], config["revenue_file"])
    wb = Workbook(path)
    rows = wb.read_sheet(config["revenue_sheet"])

    header_row_num = None
    header = None
    for row_num, row in rows:
        if any(isinstance(v, str) and v.strip() == "Код проекта" for v in row.values()):
            header_row_num = row_num
            header = row
            break
    if header is None:
        raise RuntimeError("Выгрузка выручки (journal): не нашёл заголовок с ячейкой 'Код проекта'")

    def find_col(text):
        for idx, v in header.items():
            if isinstance(v, str) and v.strip() == text:
                return idx
        return None

    col_kod = find_col("Код проекта")
    col_proekt = find_col("Проект")
    # "Вид выручки": в исходном файле сразу два кандидата — "Бюджетный
    # классификатор.Группа статей.Наименование" почти всегда пустая (в файле
    # 'FinOps 7 мес 2026.xlsx' заполнена только в ~18% строк), а колонка с
    # заголовком "Middle" (несмотря на бессмысленное имя — явно техническая
    # метка из выгрузки 1С) заполнена всегда и несёт ту же классификацию
    # (включая деление на "...-отложенная выручка VK Tech"), просто подробнее.
    # Поэтому берём колонку с наименьшим числом пустых значений среди
    # кандидатов, а не первую попавшуюся по имени — так не промахнёмся, если
    # в другой выгрузке этот столбец назовут иначе или заполненность будет
    # другой.
    vid_candidates = [c for c in (find_col("Middle"), find_col("Группа статей"),
                                   find_col("Бюджетный классификатор.Группа статей.Наименование"))
                      if c is not None]
    col_vid = None
    if vid_candidates:
        data_rows = [row for row_num, row in rows if row_num > header_row_num]
        fill_rate = {c: sum(1 for row in data_rows if row.get(c)) for c in vid_candidates}
        col_vid = max(fill_rate, key=fill_rate.get)
        if len(set(fill_rate.values())) > 1:
            warnings_out.append(
                f"Выгрузка выручки (journal): для 'Вид выручки' среди кандидатов "
                f"{[(col_index_to_letters(c), fill_rate[c]) for c in vid_candidates]} "
                f"(колонка, заполненность) выбрана {col_index_to_letters(col_vid)} — самая заполненная."
            )
    col_dogovor = find_col("Номер договора")
    col_data = find_col("Дата договора")
    col_summa = find_col("Сумма без НДС с плюсом (для бизнеса)")

    for needed, name in [
        (col_kod, "Код проекта"), (col_summa, "Сумма без НДС с плюсом (для бизнеса)"),
    ]:
        if needed is None:
            raise RuntimeError(f"Выгрузка выручки (journal): не нашёл колонку '{name}'")

    totals_by_code = {}
    details_by_code = {}
    for row_num, row in rows:
        if row_num <= header_row_num:
            continue
        kod_raw = row.get(col_kod)
        kod_norm = normalize_code(kod_raw)
        if kod_norm is None:
            continue
        summa = to_float(row.get(col_summa), default=0.0)
        totals_by_code[kod_norm] = totals_by_code.get(kod_norm, 0.0) + summa
        details_by_code.setdefault(kod_norm, []).append({
            "kod_raw": kod_raw,
            "proekt": row.get(col_proekt) if col_proekt is not None else None,
            "vid_vyruchki": row.get(col_vid) if col_vid is not None else None,
            "nomer_dogovora": row.get(col_dogovor) if col_dogovor is not None else None,
            "data_dogovora": row.get(col_data) if col_data is not None else None,
            "summa": summa,
        })

    return totals_by_code, details_by_code


def parse_revenue_extract_pivot(config, warnings_out):
    path = resolve_path(config["disk_root"], config["revenue_file"])
    wb = Workbook(path)
    rows = wb.read_sheet(config["revenue_sheet"])

    header_row_num = None
    header = None
    for row_num, row in rows:
        if row.get(0) == "Код проекта":
            header_row_num = row_num
            header = row
            break
    if header is None:
        raise RuntimeError("Выгрузка выручки: не нашёл строку заголовка ('Код проекта' в колонке A)")

    col_proekt = 1
    col_vid = 2
    col_dogovor = 3
    col_data = 4
    col_itog = max(header)  # "Общий итог" — последняя колонка

    totals_by_code = {}
    details_by_code = {}
    current_code_raw = None
    current_code_norm = None
    current_proekt = None
    current_vid = None

    for row_num, row in rows:
        if row_num <= header_row_num:
            continue
        col0 = row.get(0)
        if isinstance(col0, str) and col0.strip():
            text = col0.strip()
            if text.endswith(" Итог"):
                code_raw = text[: -len(" Итог")]
                code_norm = normalize_code(code_raw)
                if code_norm:
                    totals_by_code[code_norm] = to_float(row.get(col_itog), default=0.0)
                current_code_raw = None
                current_code_norm = None
                current_proekt = None
                current_vid = None
                continue
            else:
                current_code_raw = text
                current_code_norm = normalize_code(text)
                current_proekt = row.get(col_proekt)
                current_vid = None  # новая группа кода — вид выручки задаётся заново ниже
        # деталь-строка (может быть как первая строка группы, так и "продолжение"
        # с пустой колонкой A — тогда current_code_norm остаётся от предыдущей строки).
        # "Вид выручки" в этой pivot-таблице — тоже вложенная группа: заполнена только
        # в первой строке своей подгруппы, дальше пустая колонка C значит "тот же вид,
        # что и выше", а не "вид неизвестен" — поэтому наследуем current_vid так же,
        # как код и проект.
        if current_code_norm is None:
            continue
        vid_cell = row.get(col_vid)
        if isinstance(vid_cell, str) and vid_cell.strip():
            current_vid = vid_cell
        dogovor = row.get(col_dogovor)
        data_dog = row.get(col_data)
        summa = to_float(row.get(col_itog), default=0.0)
        if vid_cell is None and dogovor is None and summa == 0:
            continue
        details_by_code.setdefault(current_code_norm, []).append({
            "kod_raw": current_code_raw,
            "proekt": current_proekt,
            "vid_vyruchki": current_vid,
            "nomer_dogovora": dogovor,
            "data_dogovora": data_dog,
            "summa": summa,
        })

    return totals_by_code, details_by_code


def find_orphan_codes(totals_by_code, master_codes, prefixes):
    """prefixes=None/[]/"*" — без фильтра по префиксу (для сверки по всем
    направлениям сразу, где 'осиротевший' — любой код не из портфеля)."""
    no_filter = not prefixes or prefixes == "*"
    orphans = []
    for code_norm, total in totals_by_code.items():
        if code_norm in master_codes:
            continue
        if no_filter or any(code_norm.startswith(p) for p in prefixes):
            orphans.append((code_norm, total))
    return sorted(orphans, key=lambda x: -abs(x[1]))


# ---------------------------------------------------------------------------
# 4. Сверка
# ---------------------------------------------------------------------------

def is_match(a, b, tol_rub, tol_pct):
    if a is None or b is None:
        return None
    diff = abs(a - b)
    tol = max(tol_rub, tol_pct * max(abs(a), abs(b)))
    return diff <= tol


DEFERRED_REVENUE_MARKER = "отложенная выручка"
LICENSE_MARKER = "передачи лицензий"
STP_MARKER = "стандартной технической поддержки"


def categorize_revenue_breakdown(details_by_code, kod_norm):
    """
    Разбивает сумму выручки по коду (из выгрузки бухгалтерии) на 4 корзины —
    по просьбе пользователя, т.к. расхождения Сводный↔Бухгалтерия часто
    объясняются не пропажей денег, а тем, что выручка попала не в ту
    категорию/ЦФО:

    - license: любой вид выручки, где встречается "передачи лицензий"
      (охватывает все 4 варианта: обычная / отложенная / конс.продажи /
      конс.продажи-отложенная) — "выручка от лицензий относится не на то
      ЦФО", подлежит корректировке бухгалтерией.
    - stp: любой вид выручки со "стандартной технической поддержки" (те же
      4 варианта) — пересорт между видами ТП или ТП/сервисные услуги,
      на наших ЦФО СТП в принципе быть не должно.
    - deferred: то, что осталось после исключения license/stp, но содержит
      "отложенная выручка" — "чистая" отложенная выручка (сервисные услуги,
      премиальная ТП, подписочные сервисы, ПАК, "Технологии для бизнеса" и
      т.п. с пометкой "-отложенная выручка"), без license/stp-отложенной,
      которая теперь в своих корзинах.
    - other ("Прочее"): всё остальное — сюда попадает и обычная (не
      отложенная, не license, не stp) выручка вроде премиальной ТП/сервисных
      услуг/подписки/ПАК, и любой вид выручки, которого нет в этом списке
      вообще. Пользователь просил это как диагностическую корзину — "посмотрим
      сколько таких найдётся", не обязательно проблема.

    Категории проверяются по порядку license -> stp -> deferred -> other,
    поэтому license/stp "перехватывают" свои отложенные варианты раньше,
    чем сработает общая проверка на "отложенная выручка". Суммы по всем
    4 корзинам в сумме всегда дают totals_by_code[kod_norm] (проверено на
    примерах пользователя — PLTO-052, PLTO-035, PLTO-050).
    """
    result = {"license": 0.0, "stp": 0.0, "deferred": 0.0, "other": 0.0}
    if kod_norm is None:
        return result
    for d in details_by_code.get(kod_norm, []):
        vid = d.get("vid_vyruchki")
        vid_lower = vid.lower() if isinstance(vid, str) else ""
        summa = d["summa"]
        if LICENSE_MARKER in vid_lower:
            result["license"] += summa
        elif STP_MARKER in vid_lower:
            result["stp"] += summa
        elif DEFERRED_REVENUE_MARKER in vid_lower:
            result["deferred"] += summa
        else:
            result["other"] += summa
    return result


def build_comparison(groups, passport_index, totals_by_code, details_by_code, config, warnings_out):
    tol_rub = config["match_tolerance_rub"]
    tol_pct = config["match_tolerance_pct"]
    year_label = config["year"]
    half_label = period_half_label(config)

    results = []
    for g in groups:
        kod_norm = g["kod_norm"]
        proekt_names = " / ".join(sorted(set(g["proekts"])))
        kod_display = " / ".join(sorted(g["kod_raws"])) if kod_norm else "нет кода"
        rp_display = ", ".join(sorted(g["rp_set"]))
        status_display = ", ".join(sorted(g["status_set"]))
        napravlenie_display = ", ".join(sorted(g.get("napravlenie_set") or ()))
        summary_fact = g["fact_h1_sum"]

        passport_fact = None
        passport_paths = []
        passport_errors = []
        if kod_norm is not None:
            entry = passport_index.get(kod_norm)
            chosen = select_passport(entry, kod_norm, year_label, half_label, warnings_out)
            if chosen is None:
                passport_errors.append("паспорт не найден")
            else:
                passport_paths.extend(chosen["paths"])
                passport_errors.extend(chosen["errors"])
                if chosen["value"] is not None:
                    passport_fact = chosen["value"]
        else:
            passport_errors.append("нет кода — паспорт не ищем по коду")

        buh_fact = totals_by_code.get(kod_norm) if kod_norm else None
        if kod_norm is not None and buh_fact is None:
            passport_errors.append("нет строки в выгрузке бухгалтерии по этому коду")

        match_sp = is_match(summary_fact, passport_fact, tol_rub, tol_pct)
        match_sb = is_match(summary_fact, buh_fact, tol_rub, tol_pct)

        if kod_norm is None:
            status = "НЕТ КОДА"
        elif passport_fact is None and buh_fact is None:
            status = "НЕ НАЙДЕНО НИГДЕ"
        elif passport_fact is None:
            status = "ПАСПОРТ НЕ НАЙДЕН"
        elif buh_fact is None:
            status = "НЕТ В БУХГАЛТЕРИИ"
        elif match_sp and match_sb:
            status = "OK"
        else:
            status = "РАСХОЖДЕНИЕ"

        revenue_breakdown = categorize_revenue_breakdown(details_by_code, kod_norm)

        results.append({
            "proekt": proekt_names,
            "kod": kod_display,
            "kod_norm": kod_norm,
            "rp": rp_display,
            "napravlenie": napravlenie_display,
            "status_proekta": status_display,
            "summary_fact": summary_fact,
            "vychet_sum": g["vychet_sum"],
            "deferred_sum": revenue_breakdown["deferred"],
            "license_sum": revenue_breakdown["license"],
            "stp_sum": revenue_breakdown["stp"],
            "other_sum": revenue_breakdown["other"],
            "passport_fact": passport_fact,
            "buh_fact": buh_fact,
            "delta_sp": None if passport_fact is None else summary_fact - passport_fact,
            "delta_sb": None if buh_fact is None else summary_fact - buh_fact,
            "status": status,
            "comment": "; ".join(passport_errors) if passport_errors else "",
            "passport_path": passport_paths[0] if passport_paths else None,
        })

    return results


# ---------------------------------------------------------------------------
# 5. Отчёт
# ---------------------------------------------------------------------------

def status_style_for_row(row_values, status_col_idx):
    status = row_values[status_col_idx]
    if status == "OK":
        return STYLE_NUMBER_OK
    if status in ("РАСХОЖДЕНИЕ",):
        return STYLE_NUMBER_MISMATCH
    if status in ("ПАСПОРТ НЕ НАЙДЕН", "НЕТ В БУХГАЛТЕРИИ", "НЕ НАЙДЕНО НИГДЕ", "НЕТ КОДА"):
        return STYLE_WARN
    return None


PERIOD_DISPLAY = {"h1": "H1", "h2": "H2", "year": "Год"}


def build_main_header_comments(config):
    year = config["year"]
    period = config["period"]
    period_disp = PERIOD_DISPLAY[period]
    summary_period_label = PERIOD_LABEL_TEMPLATES[period].format(year=year)
    passport_row_desc = (
        f"строка «{period_half_label(config)}»" if period != "year"
        else "сама строка года (не под-строка полугодия)"
    )
    revenue_format = config.get("revenue_format", "pivot")
    revenue_desc = (
        f"строка «‹КОД› Итог», колонка «Общий итог» — сумма всех видов выручки и договоров "
        f"по данному коду проекта за весь период выгрузки"
        if revenue_format == "pivot" else
        f"построчный журнал проводок — сумма колонки «Сумма без НДС с плюсом (для бизнеса)» "
        f"по всем строкам с этим «Код проекта» за период, покрытый файлом выгрузки"
    )
    direction = config["direction"]
    summary_file = config["summary_file"]
    summary_sheet = config["summary_sheet"]
    revenue_file = config["revenue_file"]
    return [
        # Проект
        f"Файл: {summary_file}\nЛист «{summary_sheet}», колонка «Проекты».\n"
        f"Если один код проекта встречается у нескольких строк — имена объединены через « / ».",
        # Направление
        f"Файл: {summary_file}\nЛист «{summary_sheet}», колонка «Направление». "
        f"Если у кода несколько строк-проектов с разными направлениями — объединены через запятую.\n"
        f"Пусто — в Сводном для этой строки направление не указано (мусорное значение вроде "
        f"имени человека вместо направления тоже не попадает сюда, см. `load_summary()`).",
        # Код проекта
        f"Файл: {summary_file}\nЛист «{summary_sheet}», колонка «Код проекта/ ЦФО» — как есть, без изменений.\n"
        f"Для сопоставления между тремя источниками код приводится к единому виду "
        f"(верхний регистр, один разделитель вместо -/_/пробела, цифра 0 → буква O в буквенном префиксе).",
        # РП
        f"Файл: {summary_file}\nЛист «{summary_sheet}», колонка «РП».",
        # Статус проекта (Сводный)
        f"Файл: {summary_file}\nЛист «{summary_sheet}», колонка «Статус проекта».",
        # Факт {period}: Сводный
        f"Файл: {summary_file}\nЛист «{summary_sheet}», колонка «Факт» в блоке «{summary_period_label}» (Доходы).\n"
        f"Строки с ДП='Вычет' в этой сумме НЕ участвуют (их сумма — в соседнем столбце "
        f"«Сумма Вычетов»). Если код встречается у нескольких строк-проектов — суммируются.",
        # Сумма Вычетов (Сводный)
        f"Файл: {summary_file}\nЛист «{summary_sheet}», колонка «Факт» в блоке «{summary_period_label}», "
        f"сумма по строкам с ДП='Вычет' для этого кода проекта (корректировки к соседним строкам, "
        f"не отдельные проекты — см. `group_summary()`). Показывается со знаком минус.",
        # Сумма отложенной выручки (Бухгалтерия)
        f"Файл: {revenue_file}\n«Чистая» отложенная выручка по этому коду: строки, где в «Вид выручки» "
        f"встречается «{DEFERRED_REVENUE_MARKER}» (регистронезависимо), КРОМЕ отложенной выручки по "
        f"лицензиям и СТП — та теперь в соседних столбцах «Выручка от лицензий»/«Выручка от СТП» "
        f"(см. `categorize_revenue_breakdown()`). Реальное значение как есть в выгрузке, без смены знака.",
        # Выручка от лицензий
        f"Файл: {revenue_file}\nСумма по строкам этого кода, где в «Вид выручки» встречается "
        f"«{LICENSE_MARKER}» (все варианты: обычная / отложенная / конс.продажи / конс.продажи-отложенная). "
        f"Диагностика: выручка от лицензий иногда попадает не на тот код/ЦФО — такая выручка подлежит "
        f"корректировке бухгалтерией. Реальное значение как есть, без смены знака.",
        # Выручка от СТП
        f"Файл: {revenue_file}\nСумма по строкам этого кода, где в «Вид выручки» встречается "
        f"«{STP_MARKER}» (все варианты). Диагностика: пересорт между видами ТП или между ТП и "
        f"сервисными услугами — на наших ЦФО стандартной ТП в принципе быть не должно, такая выручка "
        f"подлежит корректировке. Реальное значение как есть, без смены знака.",
        # Прочее
        f"Файл: {revenue_file}\nОстаток по этому коду: строки, не попавшие ни в «Сумма отложенной "
        f"выручки», ни в «Выручка от лицензий», ни в «Выручка от СТП» — обычно это выручка по "
        f"вспомогательным/подписочным статьям без пометки «отложенная». Диагностическая колонка — "
        f"по просьбе пользователя посмотреть, что не покрывается остальными категориями.",
        # Факт {period}: Паспорт
        f"Файл: индивидуальный паспорт проекта (найден по совпадению ячейки «Код проекта» "
        f"на листе «3. Договоры_Этапы_Доходы» с кодом из Сводного, среди файлов в папках "
        f"«Действующие/Текущие» и «Завершенные/Закрытые проекты» каждого РП).\n"
        f"Лист «4. Факт БДР», блок года «{year}», {passport_row_desc}, колонка «Фактические» (Доходы).\n"
        f"Если по коду найдено несколько файлов — берётся файл из группы, соответствующей "
        f"статусу проекта (действующий/завершённый), а внутри группы — самый свежий по дате изменения.\n"
        f"Если у кода несколько проектов-строк в Сводном — суммируются значения из их паспортов.",
        # Факт {period}: Бухгалтерия
        f"Файл: {revenue_file}\n{revenue_desc} (по данным 1С/МСФО).",
        # Δ Сводный-Паспорт
        f"Факт {period_disp}: Сводный минус Факт {period_disp}: Паспорт.",
        # Δ Сводный-Бухгалтерия
        f"Факт {period_disp}: Сводный минус Факт {period_disp}: Бухгалтерия.",
        # Статус сверки
        f"OK — значения из всех доступных источников совпадают в пределах допуска "
        f"(≤ {config['match_tolerance_rub']:g} руб. или ≤ {config['match_tolerance_pct']*100:g}%, что больше).\n"
        f"РАСХОЖДЕНИЕ — расхождение сверх допуска между Сводным и Паспортом и/или Бухгалтерией.\n"
        f"ПАСПОРТ НЕ НАЙДЕН / НЕТ В БУХГАЛТЕРИИ / НЕ НАЙДЕНО НИГДЕ — не удалось найти "
        f"соответствующий источник по коду проекта.\n"
        f"НЕТ КОДА — в Сводном у проекта не указан код, сверка с паспортом/бухгалтерией по коду невозможна.",
        # Комментарий
        "Пояснение к статусу сверки: что именно не найдено и почему (например, "
        "«паспорт не найден», «нет строки в выгрузке бухгалтерии по этому коду»).",
    ]


def build_report(comparison, details_by_code, orphans, config, warnings_out, out_path):
    p = PERIOD_DISPLAY[config["period"]]
    main_headers = [
        "Проект", "Направление", "Код проекта", "РП", "Статус проекта (Сводный)",
        f"Факт {p}: Сводный", "Сумма Вычетов (Сводный)", "Сумма отложенной выручки (Бухгалтерия)",
        "Выручка от лицензий", "Выручка от СТП", "Прочее",
        f"Факт {p}: Паспорт", f"Факт {p}: Бухгалтерия",
        "Δ Сводный-Паспорт", "Δ Сводный-Бухгалтерия", "Статус сверки", "Комментарий",
    ]
    main_rows = []
    for r in sorted(comparison, key=lambda x: (x["status"] != "РАСХОЖДЕНИЕ", x["proekt"])):
        main_rows.append([
            r["proekt"], r["napravlenie"], r["kod"], r["rp"], r["status_proekta"],
            round(r["summary_fact"], 2),
            round(r["vychet_sum"], 2),
            round(r["deferred_sum"], 2),
            round(r["license_sum"], 2),
            round(r["stp_sum"], 2),
            round(r["other_sum"], 2),
            round(r["passport_fact"], 2) if r["passport_fact"] is not None else None,
            round(r["buh_fact"], 2) if r["buh_fact"] is not None else None,
            round(r["delta_sp"], 2) if r["delta_sp"] is not None else None,
            round(r["delta_sb"], 2) if r["delta_sb"] is not None else None,
            r["status"], r["comment"],
        ])
    status_col_idx = main_headers.index("Статус сверки")

    def col_style_factory(col_idx):
        def _f(row):
            return status_style_for_row(row, status_col_idx)
        return _f

    # Числовые колонки определяем по названию заголовка, а не по фиксированному
    # индексу — так добавление/удаление колонок (например "Направление" выше)
    # не рассинхронизирует раскраску.
    numeric_col_names = {
        f"Факт {p}: Сводный", "Сумма Вычетов (Сводный)", "Сумма отложенной выручки (Бухгалтерия)",
        "Выручка от лицензий", "Выручка от СТП", "Прочее",
        f"Факт {p}: Паспорт", f"Факт {p}: Бухгалтерия",
        "Δ Сводный-Паспорт", "Δ Сводный-Бухгалтерия",
    }
    main_col_styles = [
        col_style_factory(i) if h in numeric_col_names else None
        for i, h in enumerate(main_headers)
    ]

    main_sheet = SheetSpec(
        "Сверка", main_headers, main_rows,
        col_styles=main_col_styles,
        col_widths=[42, 20, 16, 20, 24, 16, 18, 24, 16, 16, 16, 16, 16, 16, 18, 16, 40],
        header_comments=build_main_header_comments(config),
    )

    detail_headers = ["Код проекта", "Проект (из выгрузки)", "Вид выручки", "Номер договора", "Дата договора", "Сумма"]
    detail_rows = []
    codes_in_scope = {r["kod_norm"] for r in comparison if r["kod_norm"]}
    for kod_norm in sorted(codes_in_scope):
        for d in details_by_code.get(kod_norm, []):
            detail_rows.append([
                kod_norm, d["proekt"], d["vid_vyruchki"], d["nomer_dogovora"],
                d["data_dogovora"], round(d["summa"], 2),
            ])
    detail_sheet = SheetSpec(
        "Детали выручки", detail_headers, detail_rows,
        col_widths=[16, 40, 45, 22, 14, 16],
    )

    orphan_headers = [
        "Код проекта (норм.)", "Итого по коду", "Проект / Клиент (из выгрузки)",
        "Вид выручки", "Номер договора", "Дата договора", "Сумма по строке",
    ]
    orphan_rows = []
    for code, total in orphans:
        lines = details_by_code.get(code, [])
        if not lines:
            orphan_rows.append([code, round(total, 2), None, None, None, None, None])
            continue
        for d in lines:
            orphan_rows.append([
                code, round(total, 2), d["proekt"], d["vid_vyruchki"],
                d["nomer_dogovora"], d["data_dogovora"], round(d["summa"], 2),
            ])
    orphan_sheet = SheetSpec(
        "Осиротевшие коды", orphan_headers, orphan_rows,
        col_widths=[18, 16, 40, 45, 22, 14, 16],
        header_comments=[
            f"Нормализованный код проекта из выгрузки бухгалтерии ({config['revenue_file']}), "
            + (
                f"начинающийся с одного из префиксов {config['orphan_code_prefixes']}, "
                if config.get("orphan_code_prefixes") and config["orphan_code_prefixes"] != "*"
                else "любой (префикс не фильтруется), "
            )
            + f"которого нет в портфеле "
            + (f"направления «{config['direction']}» " if config.get("direction") and config["direction"] != "*" else "")
            + "в Сводном отчёте.",
            "Сумма «Общий итог» по этому коду целиком (строка «‹КОД› Итог» в выгрузке) — "
            "повторяется на каждой строке кода для удобства фильтрации.",
            "Файл: " + config["revenue_file"] + "\nЛист «Свод в разрезе проектов», колонка «Проект» — "
            "обычно содержит название клиента, но это не отдельное поле «Заказчик», а как есть из выгрузки.",
            "Лист «Свод в разрезе проектов», колонка «Вид выручки».",
            "Лист «Свод в разрезе проектов», колонка «Номер договора».",
            "Лист «Свод в разрезе проектов», колонка «Дата договора».",
            "Лист «Свод в разрезе проектов», сумма по этой строке (колонка «Общий итог» для данной строки, "
            "а не по всему коду).",
        ],
    )

    write_workbook(out_path, [main_sheet, detail_sheet, orphan_sheet])


# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--config", default=os.path.join(os.path.dirname(__file__), "config.json"))
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as f:
        config = json.load(f)

    warnings_out = []

    direction_label = config.get("direction") or "все направления"
    if config.get("code_prefixes") and config["code_prefixes"] != "*":
        direction_label += f" / коды {config['code_prefixes']}"
    print(f"[1/5] Читаю Сводный отчёт ({config['summary_file']}) ...")
    summary_rows = load_summary(config, warnings_out)
    print(f"      найдено строк по направлению '{direction_label}': {len(summary_rows)}")
    groups = group_summary(summary_rows, warnings_out)
    print(f"      после группировки по коду и исключения 'Вычет': {len(groups)} позиций")

    print("[2/5] Ищу паспорта проектов ...")
    passport_index = find_passport_files(config, warnings_out)
    print(f"      найдено и распознано паспортов (по всем направлениям): {len(passport_index)}")

    print(f"[3/5] Читаю выгрузку выручки ({config['revenue_file']}) ...")
    totals_by_code, details_by_code = parse_revenue_extract(config, warnings_out)
    print(f"      найдено итоговых сумм по кодам: {len(totals_by_code)}")

    print("[4/5] Сверяю ...")
    master_codes = {g["kod_norm"] for g in groups if g["kod_norm"]}
    orphans = find_orphan_codes(totals_by_code, master_codes, config["orphan_code_prefixes"])
    comparison = build_comparison(groups, passport_index, totals_by_code, details_by_code, config, warnings_out)

    n_ok = sum(1 for r in comparison if r["status"] == "OK")
    n_mismatch = sum(1 for r in comparison if r["status"] == "РАСХОЖДЕНИЕ")
    n_other = len(comparison) - n_ok - n_mismatch
    print(f"      OK: {n_ok}, расхождений: {n_mismatch}, требуют внимания (не найдено/нет кода): {n_other}")
    print(f"      осиротевших кодов в бухгалтерии: {len(orphans)}")

    out_path = os.path.join(os.path.dirname(args.config) or ".", config["output_file"])
    print(f"[5/5] Пишу отчёт: {out_path}")
    build_report(comparison, details_by_code, orphans, config, warnings_out, out_path)

    if warnings_out:
        print(f"\nПредупреждения ({len(warnings_out)}):")
        for w in warnings_out:
            print(f"  - {w}")

    print("\nГотово.")


if __name__ == "__main__":
    main()
