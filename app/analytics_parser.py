"""
Парсинг двух аналитических отчётов Talantix:
- "Аналитика воронки по всем вакансиям" -> воронка (сколько кандидатов прошло через каждый этап)
- "Процесс работы над вакансиями" -> причины отказов + среднее время на этапе

Собирают данные для графиков "как у Леры": воронка, топ причин отказа
(компания/кандидат), отказы по этапам, среднее время на этапах.

Поддерживает несколько аккаунтов Talantix с разным словарём "родных" названий
этапов (см. ACCOUNT_1_STAGE_TRANSLATE / ACCOUNT_2_STAGE_TRANSLATE) — все они
приводятся к одному общему набору подписей (см. *_ORDER ниже), чтобы графики
по разным вакансиям из разных аккаунтов можно было честно сравнивать и
суммировать на вкладке "Все вакансии".
"""
import io
from openpyxl import load_workbook

# Порядок этапов для отображения воронки (сверху вниз).
# "Неразобранные" сюда намеренно не включаем: этот этап предшествует "Новые"
# (это ещё не разобранный входящий поток), а не следующий за ним шаг воронки —
# включение его дало бы % от базы "Новые" больше 100%, что вводит в заблуждение.
FUNNEL_STAGE_ORDER = [
    "Новые", "Приглашение на вакансию", "Интервью с HR",
    "Оценка заказчиком", "Интервью с заказчиком", "Финальное интервью",
    "Предложение о работе", "Оффер принят", "Выход на работу",
]

# Этапы, для которых среднее время на этапе информативно (см. фидбек:
# "Предложение о работе / Оффер принят / Выход на работу" — бесполезны,
# кандидат дальше не двигается, среднее по 0-1 наблюдению не показательно).
AVG_TIME_STAGE_ORDER = ["Интервью с HR", "Оценка заказчиком", "Интервью с заказчиком", "Финальное интервью"]

# Этапы, которые показываем в "Отказы по этапам" (исключаем строку самого
# этапа "Отказ", где отказов от отказа не бывает, и служебные нулевые этапы)
REJECTION_STAGE_ORDER = [
    "Неразобранные", "Новые", "Приглашение на вакансию", "Интервью с HR",
    "Оценка заказчиком", "Интервью с заказчиком", "Финальное интервью",
    "Резерв",
]



# Talantix использует двойную "||" в отчёте воронки. В "Процесс работы над
# вакансиями" исторически встречалась одинарная "|", но формат может дрейфовать
# (видели оба варианта в разных выгрузках) — поддерживаем сразу оба написания,
# чтобы не падать при малейшем изменении формата на стороне Talantix/выгрузки.
ACCOUNT_1_FUNNEL_VACANCIES = {
    "Product Lead || Кибершеф": "Product Lead",
    "Product Lead (Кибершеф)": "Product Lead",
}
ACCOUNT_1_PROCESS_VACANCIES = {
    "Product Lead || Кибершеф": "Product Lead",
    "Product Lead (Кибершеф)": "Product Lead",
}

# Перевод "родных" этапов кастомного pipeline-шаблона в унифицированные
# подписи графиков аналитики (см. FUNNEL_STAGE_ORDER/REJECTION_STAGE_ORDER выше).
ACCOUNT_1_STAGE_TRANSLATE = {
    "Неразобранные": "Новые",
    "Связаться": "Приглашение на вакансию",
    "Жду ответ — 1": "Приглашение на вакансию",
    "Жду ответ — 2": "Приглашение на вакансию",
    "Назначить собеседование с рекрутером": "Приглашение на вакансию",
    "Фриз после общения": "Новые",
    "Назначено собеседование с рекрутером": "Интервью с HR",
    "Принятие решения после скрининга": "Оценка заказчиком",
    "Отправлен заказчику": "Оценка заказчиком",
    "Назначить собеседование с заказчиком": "Интервью с заказчиком",
    "Назначено собеседование с заказчиком": "Интервью с заказчиком",
    "Ожидание фидбека от заказчика": "Интервью с заказчиком",
    "Тестовое задание": "Финальное интервью",
    "Назначить финальное собеседование": "Финальное интервью",
    "Назначено финальное собеседование": "Финальное интервью",
    "Отправлен оффер": "Предложение о работе",
    "Выход на работу": "Выход на работу",
    "Испытательный срок": "Выход на работу",
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
ACCOUNT_2_STAGE_TRANSLATE = {}
ACCOUNT_2_FUNNEL_VACANCIES = {}
ACCOUNT_2_PROCESS_VACANCIES = {}


def _load_ws(file_bytes: bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    return wb.active


def parse_funnel_xlsx(file_bytes: bytes, vacancies: dict = None, stage_translate: dict = None) -> dict:
    """Возвращает {vac_name: {unified_stage: count}} для отслеживаемых вакансий этого аккаунта."""
    vacancies = vacancies if vacancies is not None else ACCOUNT_1_FUNNEL_VACANCIES
    stage_translate = stage_translate if stage_translate is not None else ACCOUNT_1_STAGE_TRANSLATE

    ws = _load_ws(file_bytes)
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    required = ["Название вакансии в Talantix", "Этап", "Количество кандидатов, которые прошли через этап"]
    missing = [h for h in required if h not in idx]
    if missing:
        raise ValueError("В отчёте воронки не хватает колонок: " + ", ".join(missing))

    out = {v: {} for v in vacancies.values()}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[idx["Название вакансии в Talantix"]]
        if name not in vacancies:
            continue
        vac = vacancies[name]
        raw_stage = row[idx["Этап"]]
        stage = stage_translate.get(raw_stage)
        if stage is None:
            continue  # неизвестный/неразмеченный этап для этого аккаунта - пропускаем
        count = row[idx["Количество кандидатов, которые прошли через этап"]] or 0
        out[vac][stage] = out[vac].get(stage, 0) + count

    if all(not v for v in out.values()):
        raise ValueError("В отчёте воронки не найдено данных по отслеживаемым вакансиям.")
    return out


def parse_process_xlsx(file_bytes: bytes, vacancies: dict = None, stage_translate: dict = None) -> dict:
    """
    Возвращает {vac_name: {
        "avg_time": {unified_stage: avg_days},
        "rejections_by_stage": {unified_stage: count},
        "reasons_company": {reason: count},
        "reasons_candidate": {reason: count},
    }}
    Колонки с разбивкой причин отказа ("Причины отказов: ...") есть не у всех
    аккаунтов/шаблонов отчёта — если их нет, reasons_company/reasons_candidate
    останутся пустыми (без ошибки), остальные графики строятся как обычно.
    """
    vacancies = vacancies if vacancies is not None else ACCOUNT_1_PROCESS_VACANCIES
    stage_translate = stage_translate if stage_translate is not None else ACCOUNT_1_STAGE_TRANSLATE

    ws = _load_ws(file_bytes)
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    required = [
        "Название вакансии в Talantix", "Этап",
        "Среднее время соискателя на этапе, факт (дней)", "Всего отказов",
    ]
    missing = [h for h in required if h not in idx]
    if missing:
        raise ValueError("В отчёте процесса работы не хватает колонок: " + ", ".join(missing))

    reason_cols = [h for h in headers if h and h.startswith("Причины отказов:")]

    out = {
        v: {"avg_time": {}, "rejections_by_stage": {}, "reasons_company": {}, "reasons_candidate": {}}
        for v in vacancies.values()
    }

    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[idx["Название вакансии в Talantix"]]
        if name not in vacancies:
            continue
        vac = vacancies[name]
        raw_stage = row[idx["Этап"]]
        stage = stage_translate.get(raw_stage)
        if stage is None:
            continue

        avg_days = row[idx["Среднее время соискателя на этапе, факт (дней)"]]
        if avg_days is not None:
            # если на общий этап суммируются несколько "родных" (напр. Неразобранные+Подумать),
            # берём среднее по ним самим (обычно только один реально ненулевой)
            prev = out[vac]["avg_time"].get(stage)
            out[vac]["avg_time"][stage] = avg_days if prev is None else (prev + avg_days) / 2

        rej_count = row[idx["Всего отказов"]] or 0
        if rej_count:
            out[vac]["rejections_by_stage"][stage] = out[vac]["rejections_by_stage"].get(stage, 0) + rej_count

        for col in reason_cols:
            val = row[idx[col]]
            if not val:
                continue
            # "Причины отказов: мы отказали Нерелевантный опыт" -> компания, причина "Нерелевантный опыт"
            # "Причины отказов: нам отказали Не ищет работу" -> кандидат, причина "Не ищет работу"
            label = col.replace("Причины отказов: ", "")
            if label.startswith("мы отказали "):
                reason = label.replace("мы отказали ", "", 1)
                bucket = out[vac]["reasons_company"]
            elif label.startswith("нам отказали "):
                reason = label.replace("нам отказали ", "", 1)
                bucket = out[vac]["reasons_candidate"]
            else:
                continue
            bucket[reason] = bucket.get(reason, 0) + val

    if all(not any(d.values()) for d in out.values()):
        raise ValueError("В отчёте процесса работы не найдено данных по отслеживаемым вакансиям.")
    return out


def merge_funnel_process(funnel_a: dict, process_a: dict, funnel_b: dict, process_b: dict):
    """
    Объединяет данные из двух источников (аккаунтов), СУММИРУЯ по одноимённым
    вакансиям (а не просто перезаписывая) — нужно, когда вакансия из аккаунта B
    сведена в тот же display-бакет, что и вакансия из аккаунта A (например,
    "Руководитель консалтинговых проектов" слит в "Руководитель проектов").
    Среднее время при слиянии считается взвешенным по количеству кандидатов
    на этапе (та же логика, что и при сборке "Все вакансии" в build_analytics).
    """
    merged_funnel: dict = {}
    for fd in (funnel_a, funnel_b):
        for vac, stages in fd.items():
            bucket = merged_funnel.setdefault(vac, {})
            for s, c in stages.items():
                bucket[s] = bucket.get(s, 0) + c

    merged_process: dict = {}
    time_num: dict = {}
    time_den: dict = {}
    for pd, fd in ((process_a, funnel_a), (process_b, funnel_b)):
        for vac, data in pd.items():
            bucket = merged_process.setdefault(
                vac, {"avg_time": {}, "rejections_by_stage": {}, "reasons_company": {}, "reasons_candidate": {}}
            )
            fc = fd.get(vac, {})
            for s, c in data["rejections_by_stage"].items():
                bucket["rejections_by_stage"][s] = bucket["rejections_by_stage"].get(s, 0) + c
            for r, c in data["reasons_company"].items():
                bucket["reasons_company"][r] = bucket["reasons_company"].get(r, 0) + c
            for r, c in data["reasons_candidate"].items():
                bucket["reasons_candidate"][r] = bucket["reasons_candidate"].get(r, 0) + c
            for s, avg in data["avg_time"].items():
                weight = fc.get(s, 1) or 1
                key = (vac, s)
                time_num[key] = time_num.get(key, 0) + avg * weight
                time_den[key] = time_den.get(key, 0) + weight

    for (vac, s), num in time_num.items():
        den = time_den.get((vac, s))
        if den:
            merged_process[vac]["avg_time"][s] = round(num / den, 1)

    return merged_funnel, merged_process


def build_analytics(funnel_data: dict, process_data: dict) -> dict:
    """
    Собирает финальную структуру для фронтенда: по каждой вакансии + "Все вакансии".
    Среднее время при объединении вакансий считается взвешенным по количеству
    кандидатов на этапе (берём вес из funnel_data этого этапа).
    funnel_data/process_data могут быть объединены из нескольких аккаунтов —
    эта функция ничего не знает про аккаунты, работает с уже унифицированными
    названиями этапов.
    """
    vac_names = list(funnel_data.keys())
    result = {}

    def top_reasons(d: dict, n=6):
        items = sorted(d.items(), key=lambda x: -x[1])[:n]
        total = sum(d.values()) or 1
        return [{"reason": r, "count": c, "pct": round(c / total * 100, 1)} for r, c in items]

    # Этапы, которые НЕ участвуют в принудительной вложенности (там обычно
    # 1 кандидат, который просто последовательно идёт дальше — форсировать
    # вложенность не нужно, и так корректно).
    NO_NEST_STAGES = {"Предложение о работе", "Оффер принят", "Выход на работу"}

    def funnel_list(stage_counts: dict):
        base = stage_counts.get("Новые") or max(stage_counts.values(), default=1) or 1
        out = []
        cap = None  # текущий "потолок" — количество на предыдущем вложенном этапе
        for s in FUNNEL_STAGE_ORDER:
            if s not in stage_counts:
                continue
            raw = stage_counts[s]
            if s in NO_NEST_STAGES:
                c = raw
            else:
                c = raw if cap is None else min(raw, cap)
                cap = c
            out.append({"stage": s, "count": c, "pct": round(c / base * 100, 1)})
        return out

    def avg_time_list(avg_time: dict):
        return [{"stage": s, "days": avg_time[s]} for s in AVG_TIME_STAGE_ORDER if s in avg_time]

    def rejections_list(rej: dict):
        total = sum(rej.values()) or 1
        out = []
        for s in REJECTION_STAGE_ORDER:
            if s not in rej:
                continue
            c = rej[s]
            out.append({"stage": s, "count": c, "pct": round(c / total * 100, 1)})
        out.sort(key=lambda x: -x["count"])
        return out

    for vac in vac_names:
        fc = funnel_data.get(vac, {})
        pd = process_data.get(vac, {"avg_time": {}, "rejections_by_stage": {}, "reasons_company": {}, "reasons_candidate": {}})
        result[vac] = {
            "funnel": funnel_list(fc),
            "avg_time": avg_time_list(pd["avg_time"]),
            "rejections_by_stage": rejections_list(pd["rejections_by_stage"]),
            "reasons_company": top_reasons(pd["reasons_company"]),
            "reasons_candidate": top_reasons(pd["reasons_candidate"]),
            "total_rejections": sum(pd["rejections_by_stage"].values()),
        }

    # --- Комбинированная вакансия "Все вакансии" ---
    combined_funnel = {}
    for vac in vac_names:
        for s, c in funnel_data.get(vac, {}).items():
            combined_funnel[s] = combined_funnel.get(s, 0) + c

    combined_rejections = {}
    combined_reasons_company = {}
    combined_reasons_candidate = {}
    # взвешенное среднее время: сумма(avg*count_на_этапе) / сумма(count)
    weighted_time_num = {}
    weighted_time_den = {}

    for vac in vac_names:
        pd = process_data.get(vac, {"avg_time": {}, "rejections_by_stage": {}, "reasons_company": {}, "reasons_candidate": {}})
        fc = funnel_data.get(vac, {})
        for s, c in pd["rejections_by_stage"].items():
            combined_rejections[s] = combined_rejections.get(s, 0) + c
        for r, c in pd["reasons_company"].items():
            combined_reasons_company[r] = combined_reasons_company.get(r, 0) + c
        for r, c in pd["reasons_candidate"].items():
            combined_reasons_candidate[r] = combined_reasons_candidate.get(r, 0) + c
        for s, avg in pd["avg_time"].items():
            weight = fc.get(s, 1) or 1
            weighted_time_num[s] = weighted_time_num.get(s, 0) + avg * weight
            weighted_time_den[s] = weighted_time_den.get(s, 0) + weight

    combined_avg_time = {
        s: round(weighted_time_num[s] / weighted_time_den[s], 1)
        for s in weighted_time_num if weighted_time_den.get(s)
    }

    result["Все вакансии"] = {
        "funnel": funnel_list(combined_funnel),
        "avg_time": avg_time_list(combined_avg_time),
        "rejections_by_stage": rejections_list(combined_rejections),
        "reasons_company": top_reasons(combined_reasons_company),
        "reasons_candidate": top_reasons(combined_reasons_candidate),
        "total_rejections": sum(combined_rejections.values()),
    }

    return result
