"""
Парсинг отчёта Talantix "Кандидаты на вакансиях" (.xlsx) в структуру трекера воронки.

Поддерживает несколько РАЗНЫХ аккаунтов Talantix (наш собственный + аккаунт(ы)
другого рекрутёра) — у каждого свой словарь названий этапов, но все они
приводятся к одному общему набору статусов трекера (см. ACCOUNTS ниже).

Логика:
- берём только нужные вакансии конкретного аккаунта
- переименовываем вакансии в человеко-понятные названия
- группируем "родные" этапы Talantix этого аккаунта в общие статусы трекера
- считаем "дней в статусе" от даты перевода на этап до сегодняшнего дня
"""
import hashlib
import io
from datetime import datetime
from openpyxl import load_workbook

REQUIRED_COLUMNS = [
    "Название вакансии в Talantix",
    "ID вакансии в Talantix",
    "Этап",
    "Дата перевода кандидата на этап",
    "ФИО",
    "Источник кандидата",
    "Причины отказов",
    "Дата рождения; возраст",
    "Пол",
    "Город",
    "Email",
    "Телефон",
    "Образование",
    "Общий стаж",
    "Последнее место работы",
    "Должность",
    "Знание языков",
    "Зарплатные ожидания",
    "Комментарий 1",
    "Комментарий 2",
    "Комментарий 3",
]

# ---------------------------------------------------------------------------
# Аккаунт 1 — наш собственный Talantix. Вакансия "Product Lead || Кибершеф" использует
# ДРУГОЙ pipeline-шаблон Talantix (третий по счёту в рамках аккаунта!) — со
# своими "родными" названиями этапов (Связаться/Жду ответ/Отказ от рекрутера
# и т.д.). Смэтчено и согласовано с человеком:
# ---------------------------------------------------------------------------
ACCOUNT_1_VACANCIES = {
    "Product Lead || Кибершеф": "Product Lead",
    "Product Lead (Кибершеф)": "Product Lead",
}
ACCOUNT_1_STATUS_MAP = {
    "Неразобранные": "Пристрелка",
    "Связаться": "Разбор резюме",
    "Жду ответ — 1": "Разбор резюме",
    "Жду ответ — 2": "Разбор резюме",
    "Назначить собеседование с рекрутером": "Разбор резюме",
    "Фриз после общения": "Разбор резюме",
    "Назначено собеседование с рекрутером": "Интервью с HR",
    "Принятие решения после скрининга": "Ожидание ОС (после HR)",
    "Отправлен заказчику": "Ожидание ОС (после HR)",
    "Назначить собеседование с заказчиком": "Интервью с заказчиком 1",
    "Назначено собеседование с заказчиком": "Интервью с заказчиком 1",
    "Ожидание фидбека от заказчика": "Интервью с заказчиком 1",
    "Тестовое задание": "Интервью с заказчиком 2",
    "Назначить финальное собеседование": "Интервью с заказчиком 2",
    "Назначено финальное собеседование": "Интервью с заказчиком 2",
    "Отправлен оффер": "Оффер / Выход",
    "Выход на работу": "Оффер / Выход",
    "Испытательный срок": "Оффер / Выход",
    "Отказ": "Отказ",
    "Отказ от рекрутера": "Отказ",
    "Отказ от кандидата после рекрутера": "Отказ",
    "Отказ от заказчика после профиля": "Отказ",
    "Отказ от заказчика после собеседования": "Отказ",
    "Отказ от кандидата после заказчика": "Отказ",
    "Отказ от заказчика на финале": "Отказ",
    "Отказ от кандидата на финале": "Отказ",
}

# Эта копия не подключает второй Talantix-аккаунт.
ACCOUNT_2_VACANCIES = {}
ACCOUNT_2_STATUS_MAP = {}

# Все аккаунты вместе — main.py перебирает их по очереди и объединяет результат.
ACCOUNTS = [
    {"vacancies": ACCOUNT_1_VACANCIES, "status_map": ACCOUNT_1_STATUS_MAP},
    {"vacancies": ACCOUNT_2_VACANCIES, "status_map": ACCOUNT_2_STATUS_MAP},
]


def _fmt_date(d):
    if not d:
        return ""
    if isinstance(d, datetime):
        return d.strftime("%d.%m.%Y")
    return str(d)


def _normalize_phone(phone: str) -> str:
    """
    Оставляет только цифры и берёт последние 10 (без кода страны) — так что
    '+7 (916) 151-58-14', '8 916 151 58 14' и '79161515814' дают один и тот же
    ключ. Это нужно, чтобы мелкие расхождения в форматировании телефона между
    ежедневными выгрузками не рвали привязку резюме/плюсов к кандидату.
    """
    digits = "".join(ch for ch in (phone or "") if ch.isdigit())
    return digits[-10:] if len(digits) >= 10 else digits


def _normalize_fio(fio: str) -> str:
    """Схлопывает лишние пробелы и приводит к одному регистру."""
    return " ".join((fio or "").split()).lower()


def parse_candidates_xlsx(file_bytes: bytes, vacancies: dict = None, status_map: dict = None) -> list[dict]:
    """
    Принимает содержимое .xlsx (bytes), возвращает список кандидатов
    в формате, который понимает фронтенд трекера.
    vacancies/status_map по умолчанию — аккаунт 1 (наш собственный Talantix),
    для других аккаунтов передавай явно (см. ACCOUNTS).
    Бросает ValueError, если в файле нет ожидаемых колонок.
    """
    vacancies = vacancies if vacancies is not None else ACCOUNT_1_VACANCIES
    status_map = status_map if status_map is not None else ACCOUNT_1_STATUS_MAP

    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    missing = [h for h in REQUIRED_COLUMNS if h not in idx]
    if missing:
        raise ValueError(
            "В файле не хватает ожидаемых колонок Talantix: " + ", ".join(missing)
        )

    today = datetime.utcnow()
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[idx["Название вакансии в Talantix"]]
        if name not in vacancies:
            continue
        stage = row[idx["Этап"]]
        if stage not in status_map:
            # Do not silently lose a Product Lead candidate: log the exact FIO/stage so Railway
            # immediately explains a Talantix-vs-dashboard count mismatch.
            fio_for_log = row[idx["ФИО"]] or "Без имени"
            print(
                f'Product Lead parser: skipped "{fio_for_log}" — unmapped Talantix stage: "{stage}"'
            )
            continue

        d = row[idx["Дата перевода кандидата на этап"]]
        days = (today - d).days if isinstance(d, datetime) else 0
        fio = row[idx["ФИО"]] or "Без имени"
        reason = row[idx["Причины отказов"]] or ""
        phone = row[idx["Телефон"]] or ""
        email = row[idx["Email"]] or ""

        # ВАЖНО: id кандидата должен быть стабильным между обновлениями отчёта —
        # к нему привязаны плюсы и загруженные резюме. Раньше это был просто
        # номер строки в файле (c1, c2, ...), который "плавал" при малейшем
        # изменении порядка строк в новой выгрузке — из-за этого резюме могло
        # потеряться или приклеиться к другому кандидату. Теперь id считается
        # от самого кандидата: вакансия + ФИО + нормализованный телефон/email.
        # Нормализация телефона (только цифры, последние 10) и ФИО (пробелы,
        # регистр) нужна, чтобы мелкие расхождения в форматировании между
        # ежедневными выгрузками (например "+7 (916)..." vs "8 916...") не
        # ломали привязку. Пока ФИО и телефон/email кандидата не меняются
        # в Talantix — id не меняется никогда, порядок строк роли не играет.
        norm_phone = _normalize_phone(phone)
        norm_email = email.strip().lower()

        # Normally phone/email make the candidate identity stable.
        # But Talantix can contain several different people named
        # "Кандидат без контактов" with neither phone nor email.
        # In that case using only vacancy+FIO would incorrectly merge them.
        if norm_phone or norm_email:
            identity_key = (
                f"{vacancies[name]}|{_normalize_fio(fio)}|{norm_phone}|{norm_email}"
            )
        else:
            birth = str(row[idx["Дата рождения; возраст"]] or "").strip().lower()
            position = str(row[idx["Должность"]] or "").strip().lower()
            last_job = str(row[idx["Последнее место работы"]] or "").strip().lower()
            identity_key = (
                f"{vacancies[name]}|{_normalize_fio(fio)}|no-contacts|"
                f"{birth}|{position}|{last_job}"
            )

        cand_id = "c" + hashlib.md5(identity_key.encode("utf-8")).hexdigest()[:10]

        out.append({
            "id": cand_id,
            "fio": fio,
            "vac": vacancies[name],
            "raw_stage": stage,
            "status": status_map[stage],
            "date": _fmt_date(d),
            "days": days,
            "reason": reason,
            "source": row[idx["Источник кандидата"]] or "",
            "birth": row[idx["Дата рождения; возраст"]] or "",
            "gender": row[idx["Пол"]] or "",
            "city": row[idx["Город"]] or "",
            "email": email,
            "phone": phone,
            "education": row[idx["Образование"]] or "",
            "experience": row[idx["Общий стаж"]] or "",
            "last_job": row[idx["Последнее место работы"]] or "",
            "position": row[idx["Должность"]] or "",
            "languages": row[idx["Знание языков"]] or "",
            "salary": row[idx["Зарплатные ожидания"]] or "",
            "comment1": row[idx["Комментарий 1"]] or "",
            "comment2": row[idx["Комментарий 2"]] or "",
            "comment3": row[idx["Комментарий 3"]] or "",
        })

    if not out:
        raise ValueError(
            "В файле не найдено ни одного кандидата по отслеживаемым вакансиям этого аккаунта. "
            "Проверь, что это тот самый отчёт «Кандидаты на вакансиях» из Talantix и что название "
            "вакансии совпадает с настроенным."
        )

    return out


def parse_candidates_xlsx_from_path(path: str, vacancies: dict = None, status_map: dict = None) -> list[dict]:
    """Читает .xlsx с диска и парсит его."""
    with open(path, "rb") as f:
        return parse_candidates_xlsx(f.read(), vacancies, status_map)


# Порядок статусов "по продвинутости в воронке" — используется, когда один и
# тот же реальный человек (совпадают вакансия+ФИО+телефон/email — например,
# откликнулся и через наш Talantix, и через Talantix другого рекрутёра на ту
# же смэтченную вакансию) встречается в объединённом списке дважды. Оставляем
# карточку с самым продвинутым статусом, "Отказ" — в самом низу приоритета:
# если человека отклонили в одном канале, но он ещё активен в другом, важнее
# показать, что он ещё в игре.
STAGE_PRIORITY = {
    "Отказ": 0,
    "Пристрелка": 1,
    "Разбор резюме": 2,
    "Интервью с HR": 3,
    "Ожидание ОС (после HR)": 4,
    "Интервью с заказчиком 1": 5,
    "Интервью с заказчиком 2": 6,
    "Оффер / Выход": 7,
}

# Поля, которые при дедупликации подтягиваются из "проигравшей" карточки,
# если в "победившей" они пустые — чтобы объединение не теряло данные.
_FILL_IF_EMPTY_FIELDS = [
    "source", "birth", "gender", "city", "email", "phone", "education",
    "experience", "last_job", "position", "languages", "salary", "reason",
]


def merge_duplicate_candidates(candidates: list[dict]) -> list[dict]:
    """
    Схлопывает кандидатов с одинаковым id (см. STAGE_PRIORITY) в одну карточку:
    оставляет самый продвинутый статус, добор недостающих полей и объединение
    комментариев — из "проигравшей" копии, чтобы информация не терялась.
    """
    by_id: dict[str, dict] = {}
    for c in candidates:
        cid = c["id"]
        if cid not in by_id:
            by_id[cid] = c
            continue

        existing = by_id[cid]
        print(
            "Product Lead parser: duplicate candidate identity merged: "
            f'{c.get("fio", cid)} | {existing.get("raw_stage")} + {c.get("raw_stage")}'
        )
        existing_rank = STAGE_PRIORITY.get(existing["status"], -1)
        new_rank = STAGE_PRIORITY.get(c["status"], -1)
        keep, other = (c, existing) if new_rank > existing_rank else (existing, c)

        merged = dict(keep)
        for field in _FILL_IF_EMPTY_FIELDS:
            if not merged.get(field) and other.get(field):
                merged[field] = other[field]

        seen = set()
        combined_comments = []
        for cm in [keep.get("comment1"), keep.get("comment2"), keep.get("comment3"),
                   other.get("comment1"), other.get("comment2"), other.get("comment3")]:
            if cm and cm not in seen:
                combined_comments.append(cm)
                seen.add(cm)
        merged["comment1"] = combined_comments[0] if len(combined_comments) > 0 else ""
        merged["comment2"] = combined_comments[1] if len(combined_comments) > 1 else ""
        merged["comment3"] = combined_comments[2] if len(combined_comments) > 2 else ""

        by_id[cid] = merged

    return list(by_id.values())
