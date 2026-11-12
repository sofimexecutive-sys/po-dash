import json
import os
from datetime import datetime, timezone
from urllib.parse import quote

import requests
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from sqlalchemy import desc
from sqlalchemy.orm import Session
from fastapi import Depends

from .db import Base, engine, get_db, run_light_migrations
from .models import Vote, Snapshot, CandidateExtra, AnalyticsSnapshot, CandidateComment, CandidateScore
from .parser import (
    parse_candidates_xlsx, parse_candidates_xlsx_from_path,
    ACCOUNT_1_VACANCIES, ACCOUNT_1_STATUS_MAP,
    ACCOUNT_2_VACANCIES, ACCOUNT_2_STATUS_MAP,
    merge_duplicate_candidates,
)
from .analytics_parser import (
    parse_funnel_xlsx, parse_process_xlsx, build_analytics, merge_funnel_process,
    ACCOUNT_1_FUNNEL_VACANCIES, ACCOUNT_1_PROCESS_VACANCIES, ACCOUNT_1_STAGE_TRANSLATE,
    ACCOUNT_2_FUNNEL_VACANCIES, ACCOUNT_2_PROCESS_VACANCIES, ACCOUNT_2_STAGE_TRANSLATE,
)

Base.metadata.create_all(bind=engine)
run_light_migrations()

app = FastAPI(title="Трекинг воронки API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC_DIR = os.path.join(os.path.dirname(__file__), "..", "static")

# Постоянная ссылка на отчёт (Google Таблицы/Диск), которую человек сам обновляет
# на своей стороне — тем же файлом по тому же ID. Дашборд НЕ даёт грузить файлы
# через браузер, а просто скачивает свежую версию по этой ссылке при обновлении.
# Задаётся переменной окружения REPORT_SOURCE_URL — см. .env.example.
REPORT_SOURCE_URL = os.getenv("REPORT_SOURCE_URL", "")

# Альтернатива/резерв: путь к .xlsx на диске сервера (например, на Railway Volume),
# если решишь когда-нибудь вернуться к файлу на диске вместо ссылки.
REPORT_FILE_PATH = os.getenv("REPORT_FILE_PATH", "")

# Ссылки на два аналитических отчёта — воронка и процесс работы над вакансиями.
# Нужны для раздела "Аналитика" (графики воронки, причин отказа, среднего времени).
FUNNEL_SOURCE_URL = os.getenv("FUNNEL_SOURCE_URL", "")
PROCESS_SOURCE_URL = os.getenv("PROCESS_SOURCE_URL", "")

# Второй аккаунт Talantix (другой рекрутёр, своя вакансия) — те же 3 отчёта,
# но из его собственного Talantix. Опционально: если не заданы, просто не
# подключаются, работает только первый (основной) аккаунт.
REPORT_SOURCE_URL_2 = os.getenv("REPORT_SOURCE_URL_2", "")
FUNNEL_SOURCE_URL_2 = os.getenv("FUNNEL_SOURCE_URL_2", "")
PROCESS_SOURCE_URL_2 = os.getenv("PROCESS_SOURCE_URL_2", "")


def _google_export_url(url: str) -> str:
    """
    Если это обычная ссылка на просмотр/редактирование Google Таблицы
    (docs.google.com/spreadsheets/d/{ID}/edit...), превращает её в прямую
    ссылку на скачивание .xlsx. Прочие ссылки возвращает как есть.
    """
    if "docs.google.com/spreadsheets" in url and "/export" not in url:
        try:
            file_id = url.split("/d/")[1].split("/")[0]
            return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        except IndexError:
            return url
    return url


class VoteSet(BaseModel):
    candidate_id: str
    person: str
    reaction: str | None = None  # 'like' | 'dislike' | None (снять реакцию)


# ---------- статика (фронтенд) ----------

@app.get("/")
def serve_index():
    return FileResponse(
        os.path.join(STATIC_DIR, "index.html"),
        headers={"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"},
    )


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# ---------- кандидаты / отчёт ----------

def _latest_snapshot(db: Session) -> Snapshot | None:
    return db.query(Snapshot).order_by(desc(Snapshot.id)).first()


@app.get("/api/candidates")
def get_candidates(db: Session = Depends(get_db)):
    """
    Prefer fresh Talantix data on every dashboard load.

    The previous implementation returned the latest DB snapshot first and only
    refreshed in the browser afterwards. If that background refresh failed or
    was slow, the user could keep seeing a stale 9-candidate snapshot even
    though the Drive report already contained 10 candidates.

    Now we try to refresh from REPORT_SOURCE_URL before responding. If the
    source is temporarily unavailable, we safely fall back to the latest saved
    snapshot.
    """
    live_error = None
    if REPORT_SOURCE_URL or REPORT_FILE_PATH:
        try:
            candidates = _fetch_candidates_account1()
            source_label = REPORT_SOURCE_URL or f"disk:{REPORT_FILE_PATH}"

            if REPORT_SOURCE_URL_2:
                try:
                    cand2 = _fetch_candidates_account2()
                    if cand2:
                        candidates += cand2
                    source_label += f" + {REPORT_SOURCE_URL_2}"
                except Exception as exc:
                    print(f"Product Lead dashboard live read: account 2 skipped: {exc}")

            raw_candidate_count = len(candidates)
            candidates = merge_duplicate_candidates(candidates)
            print(
                f"Product Lead dashboard live read: source rows={raw_candidate_count}, "
                f"dashboard candidates={len(candidates)}"
            )
            snap = _save_snapshot(db, candidates, source_label)
        except Exception as exc:
            live_error = str(exc)
            print(f"CFO dashboard live read failed; using cached snapshot: {exc}")
            snap = _latest_snapshot(db)
    else:
        snap = _latest_snapshot(db)

    if not snap:
        return {
            "candidates": [],
            "fetched_at": None,
            "source_url": None,
            "live_error": live_error,
        }

    candidates = json.loads(snap.data_json)

    return {
        "candidates": candidates,
        "fetched_at": snap.fetched_at.isoformat() if snap.fetched_at else None,
        "source_url": snap.source_url,
        "live_error": live_error,
    }


def _save_snapshot(db: Session, candidates: list, source_url: str | None):
    snap = Snapshot(source_url=source_url, data_json=json.dumps(candidates, ensure_ascii=False))
    db.add(snap)
    db.commit()
    return snap


@app.get("/api/report-path")
def report_source_info():
    return {
        "source_url": REPORT_SOURCE_URL or None,
        "file_path": REPORT_FILE_PATH or None,
        "file_exists": os.path.isfile(REPORT_FILE_PATH) if REPORT_FILE_PATH else False,
        "funnel_configured": bool(FUNNEL_SOURCE_URL),
        "process_configured": bool(PROCESS_SOURCE_URL),
        "account2_configured": bool(REPORT_SOURCE_URL_2),
    }


def _fetch_xlsx(url: str) -> bytes:
    resp = requests.get(_google_export_url(url), timeout=30)
    resp.raise_for_status()
    return resp.content


def _fetch_candidates_account1() -> list:
    """Аккаунт 1 (наш собственный Talantix): по ссылке или по файлу с диска."""
    if REPORT_SOURCE_URL:
        content = _fetch_xlsx(REPORT_SOURCE_URL)
        return parse_candidates_xlsx(content, ACCOUNT_1_VACANCIES, ACCOUNT_1_STATUS_MAP)
    if REPORT_FILE_PATH:
        if not os.path.isfile(REPORT_FILE_PATH):
            raise HTTPException(404, f"Файл отчёта не найден на сервере по пути {REPORT_FILE_PATH}")
        return parse_candidates_xlsx_from_path(REPORT_FILE_PATH, ACCOUNT_1_VACANCIES, ACCOUNT_1_STATUS_MAP)
    raise HTTPException(
        400,
        "Источник отчёта (аккаунт 1) не настроен. Задайте REPORT_SOURCE_URL (ссылка) "
        "или REPORT_FILE_PATH (файл на диске) в переменных окружения.",
    )


def _fetch_candidates_account2() -> list | None:
    """Аккаунт 2 (Talantix другого рекрутёра) — опционален, если не настроен, просто None."""
    if not REPORT_SOURCE_URL_2:
        return None
    content = _fetch_xlsx(REPORT_SOURCE_URL_2)
    return parse_candidates_xlsx(content, ACCOUNT_2_VACANCIES, ACCOUNT_2_STATUS_MAP)


@app.post("/api/refresh")
def refresh_report(db: Session = Depends(get_db)):
    """
    Обновляет данные трекера из настроенных постоянных источников: аккаунт 1
    (обязателен) + аккаунт 2 (опционально, если настроен REPORT_SOURCE_URL_2).
    Заодно, если настроены, перечитывает аналитические отчёты обоих аккаунтов.
    """
    try:
        candidates = _fetch_candidates_account1()
    except HTTPException:
        raise
    except requests.RequestException as e:
        raise HTTPException(400, f"Не удалось скачать файл по ссылке (аккаунт 1): {e}")
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"Не удалось прочитать .xlsx (аккаунт 1): {e}")
    source_label = REPORT_SOURCE_URL or f"disk:{REPORT_FILE_PATH}"

    account2_error = None
    if REPORT_SOURCE_URL_2:
        try:
            cand2 = _fetch_candidates_account2()
            candidates += cand2
            source_label += f" + {REPORT_SOURCE_URL_2}"
        except requests.RequestException as e:
            account2_error = f"Не удалось скачать файл по ссылке (аккаунт 2): {e}"
        except ValueError as e:
            account2_error = str(e)
        except Exception as e:
            account2_error = f"Не удалось прочитать .xlsx (аккаунт 2): {e}"

    raw_candidate_count = len(candidates)
    candidates = merge_duplicate_candidates(candidates)
    print(
        f"Product Lead dashboard refresh: source rows={raw_candidate_count}, "
        f"dashboard candidates={len(candidates)}"
    )
    snap = _save_snapshot(db, candidates, source_label)

    analytics_error = None
    try:
        _refresh_analytics(db)
    except Exception as e:
        analytics_error = str(e)

    return {
        "ok": True,
        "count": len(candidates),
        "fetched_at": snap.fetched_at.isoformat() if snap.fetched_at else None,
        "source_url": source_label,
        "analytics_error": analytics_error,
        "account2_error": account2_error,
    }


@app.on_event("startup")
def _auto_refresh_on_startup():
    """При старте сервиса сразу подтягиваем данные, если источник настроен."""
    from .db import SessionLocal
    db = SessionLocal()
    try:
        candidates = _fetch_candidates_account1()
        source_label = REPORT_SOURCE_URL or f"disk:{REPORT_FILE_PATH}"
        if REPORT_SOURCE_URL_2:
            try:
                candidates += _fetch_candidates_account2()
                source_label += f" + {REPORT_SOURCE_URL_2}"
            except Exception as e:
                print(f"[startup] не удалось прочитать отчёт аккаунта 2: {e}")
        raw_candidate_count = len(candidates)
        candidates = merge_duplicate_candidates(candidates)
        print(
            f"Product Lead dashboard startup refresh: source rows={raw_candidate_count}, "
            f"dashboard candidates={len(candidates)}"
        )
        _save_snapshot(db, candidates, source_label)
    except Exception as e:
        print(f"[startup] не удалось прочитать отчёт при запуске: {e}")

    try:
        _refresh_analytics(db)
    except Exception as e:
        print(f"[startup] не удалось прочитать аналитические отчёты: {e}")
    finally:
        db.close()


# ---------- аналитика (воронка / отказы / среднее время) ----------

def _latest_analytics_snapshot(db: Session):
    return db.query(AnalyticsSnapshot).order_by(desc(AnalyticsSnapshot.id)).first()


def _refresh_analytics(db: Session):
    if not (FUNNEL_SOURCE_URL and PROCESS_SOURCE_URL):
        raise ValueError("FUNNEL_SOURCE_URL и/или PROCESS_SOURCE_URL не настроены")

    funnel_bytes = _fetch_xlsx(FUNNEL_SOURCE_URL)
    process_bytes = _fetch_xlsx(PROCESS_SOURCE_URL)
    funnel_data = parse_funnel_xlsx(funnel_bytes, ACCOUNT_1_FUNNEL_VACANCIES, ACCOUNT_1_STAGE_TRANSLATE)
    process_data = parse_process_xlsx(process_bytes, ACCOUNT_1_PROCESS_VACANCIES, ACCOUNT_1_STAGE_TRANSLATE)

    if FUNNEL_SOURCE_URL_2 and PROCESS_SOURCE_URL_2:
        try:
            funnel_bytes_2 = _fetch_xlsx(FUNNEL_SOURCE_URL_2)
            process_bytes_2 = _fetch_xlsx(PROCESS_SOURCE_URL_2)
            funnel_data_2 = parse_funnel_xlsx(funnel_bytes_2, ACCOUNT_2_FUNNEL_VACANCIES, ACCOUNT_2_STAGE_TRANSLATE)
            process_data_2 = parse_process_xlsx(process_bytes_2, ACCOUNT_2_PROCESS_VACANCIES, ACCOUNT_2_STAGE_TRANSLATE)
            funnel_data, process_data = merge_funnel_process(funnel_data, process_data, funnel_data_2, process_data_2)
        except Exception as e:
            print(f"[analytics] не удалось прочитать аналитику аккаунта 2: {e}")

    analytics = build_analytics(funnel_data, process_data)

    snap = AnalyticsSnapshot(data_json=json.dumps(analytics, ensure_ascii=False))
    db.add(snap)
    db.commit()
    return snap


@app.get("/api/analytics")
def get_analytics(db: Session = Depends(get_db)):
    snap = _latest_analytics_snapshot(db)
    if not snap:
        return {"data": None, "fetched_at": None}
    return {
        "data": json.loads(snap.data_json),
        "fetched_at": snap.fetched_at.isoformat() if snap.fetched_at else None,
    }


@app.post("/api/analytics/refresh")
def refresh_analytics_endpoint(db: Session = Depends(get_db)):
    try:
        snap = _refresh_analytics(db)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except requests.RequestException as e:
        raise HTTPException(400, f"Не удалось скачать один из отчётов: {e}")
    except Exception as e:
        raise HTTPException(400, f"Не удалось прочитать отчёты: {e}")
    return {"ok": True, "fetched_at": snap.fetched_at.isoformat() if snap.fetched_at else None}


# ---------- плюсы / минусы ----------

@app.get("/api/votes")
def get_votes(db: Session = Depends(get_db)):
    """{candidate_id: {"likes": [...], "dislikes": [...]}}"""
    votes = db.query(Vote).all()
    out: dict[str, dict[str, list[str]]] = {}
    for v in votes:
        bucket = out.setdefault(v.candidate_id, {"likes": [], "dislikes": []})
        key = "likes" if v.reaction == "like" else "dislikes"
        bucket[key].append(v.person)
    return out


@app.post("/api/votes/set")
def set_vote(payload: VoteSet, db: Session = Depends(get_db)):
    """
    Явно выставляет реакцию человека на кандидата (не toggle — фронтенд сам
    решает, какое состояние должно стать активным, и присылает его напрямую):
    reaction='like' | 'dislike' ставит/меняет реакцию, reaction=None снимает её.
    """
    existing = (
        db.query(Vote)
        .filter(Vote.candidate_id == payload.candidate_id, Vote.person == payload.person)
        .first()
    )
    if payload.reaction is None:
        if existing:
            db.delete(existing)
            db.commit()
    elif payload.reaction not in ("like", "dislike"):
        raise HTTPException(400, "reaction должен быть 'like', 'dislike' или null")
    elif existing:
        existing.reaction = payload.reaction
        db.commit()
    else:
        db.add(Vote(candidate_id=payload.candidate_id, person=payload.person, reaction=payload.reaction))
        db.commit()

    votes = db.query(Vote).filter(Vote.candidate_id == payload.candidate_id).all()
    likes = [v.person for v in votes if v.reaction == "like"]
    dislikes = [v.person for v in votes if v.reaction == "dislike"]
    return {"candidate_id": payload.candidate_id, "likes": likes, "dislikes": dislikes}


# ---------- оценки Product Lead / «Пристрелка» ----------

class ScoreSet(BaseModel):
    candidate_id: str
    person: str | None = None
    profile_score: int | None = None
    company_score: int | None = None


def _validate_score(value: int | None, field: str):
    if value is None:
        return
    if value < 1 or value > 10:
        raise HTTPException(400, f"{field} должен быть от 1 до 10")


SCORE_EVALUATORS = {"Дима", "Никита"}


@app.get("/api/scores")
def get_scores(db: Session = Depends(get_db)):
    rows = (
        db.query(CandidateScore)
        .filter(CandidateScore.person.in_(SCORE_EVALUATORS))
        .order_by(CandidateScore.updated_at.asc())
        .all()
    )
    out: dict[str, dict] = {}
    for row in rows:
        out.setdefault(row.candidate_id, {})[row.person] = {
            "person": row.person,
            "profile_score": row.profile_score,
            "company_score": row.company_score,
            "updated_at": row.updated_at.isoformat() if row.updated_at else None,
        }
    return out


@app.get("/api/scores/{candidate_id}")
def get_candidate_score(candidate_id: str, db: Session = Depends(get_db)):
    rows = (
        db.query(CandidateScore)
        .filter(
            CandidateScore.candidate_id == candidate_id,
            CandidateScore.person.in_(SCORE_EVALUATORS),
        )
        .order_by(CandidateScore.updated_at.asc())
        .all()
    )
    out = {}
    for row in rows:
        out[row.person] = {
            "person": row.person,
            "profile_score": row.profile_score,
            "company_score": row.company_score,
            "updated_at": row.updated_at.isoformat() if row.updated_at else None,
        }
    return out


@app.post("/api/scores/set")
def set_score(payload: ScoreSet, db: Session = Depends(get_db)):
    person = (payload.person or "").strip()
    if person not in SCORE_EVALUATORS:
        raise HTTPException(400, "Оценщик должен быть Дима или Никита")

    _validate_score(payload.profile_score, "Попадание в профиль")
    _validate_score(payload.company_score, "Попадание в компанию")

    row = (
        db.query(CandidateScore)
        .filter(
            CandidateScore.candidate_id == payload.candidate_id,
            CandidateScore.person == person,
        )
        .first()
    )
    now = datetime.now(timezone.utc)
    if not row:
        row = CandidateScore(candidate_id=payload.candidate_id, person=person)
        db.add(row)

    row.profile_score = payload.profile_score
    row.company_score = payload.company_score
    row.updated_at = now
    db.commit()
    db.refresh(row)

    return {
        "candidate_id": row.candidate_id,
        "person": row.person,
        "profile_score": row.profile_score,
        "company_score": row.company_score,
        "updated_at": row.updated_at.isoformat() if row.updated_at else None,
    }


# ---------- комментарии заказчиков ----------

class CommentCreate(BaseModel):
    candidate_id: str
    person: str
    text: str


@app.get("/api/candidate-extra/{candidate_id}/comments")
def get_comments(candidate_id: str, db: Session = Depends(get_db)):
    rows = (
        db.query(CandidateComment)
        .filter(CandidateComment.candidate_id == candidate_id)
        .order_by(CandidateComment.created_at.asc())
        .all()
    )
    return [
        {"id": r.id, "person": r.person, "text": r.text, "created_at": r.created_at.isoformat() if r.created_at else None}
        for r in rows
    ]


@app.post("/api/candidate-extra/{candidate_id}/comments")
def add_comment(candidate_id: str, payload: CommentCreate, db: Session = Depends(get_db)):
    text = (payload.text or "").strip()
    if not text:
        raise HTTPException(400, "Комментарий не может быть пустым")
    if not payload.person:
        raise HTTPException(400, "Не указано имя")
    row = CandidateComment(candidate_id=candidate_id, person=payload.person, text=text)
    db.add(row)
    db.commit()
    db.refresh(row)
    return {"id": row.id, "person": row.person, "text": row.text, "created_at": row.created_at.isoformat() if row.created_at else None}


@app.get("/api/health")
def health():
    return {"status": "ok", "time": datetime.utcnow().isoformat()}


# ---------- резюме + ссылка на Talantix (данные, которых нет в отчётах) ----------

LINK_STALE_DAYS = 7
MAX_RESUME_SIZE = 15 * 1024 * 1024  # 15 МБ


def _get_or_create_extra(db: Session, candidate_id: str) -> CandidateExtra:
    row = db.query(CandidateExtra).filter(CandidateExtra.candidate_id == candidate_id).first()
    if not row:
        row = CandidateExtra(candidate_id=candidate_id)
        db.add(row)
        db.commit()
        db.refresh(row)
    return row


def _extra_to_dict(row: CandidateExtra | None):
    if not row:
        return {
            "resume_filename": None, "resume_uploaded_at": None, "resume_uploaded_by": None,
            "resume_mimetype": None,
            "talantix_link": None, "talantix_link_added_at": None, "link_is_stale": False,
        }
    stale = False
    if row.talantix_link_added_at:
        age = datetime.now(timezone.utc) - row.talantix_link_added_at.replace(tzinfo=timezone.utc) if row.talantix_link_added_at.tzinfo is None else datetime.now(timezone.utc) - row.talantix_link_added_at
        stale = age.days >= LINK_STALE_DAYS
    return {
        "resume_filename": row.resume_filename,
        "resume_uploaded_at": row.resume_uploaded_at.isoformat() if row.resume_uploaded_at else None,
        "resume_uploaded_by": row.resume_uploaded_by,
        "resume_mimetype": row.resume_mimetype,
        "talantix_link": row.talantix_link,
        "talantix_link_added_at": row.talantix_link_added_at.isoformat() if row.talantix_link_added_at else None,
        "link_is_stale": stale,
    }


@app.get("/api/candidate-extra/{candidate_id}")
def get_candidate_extra(candidate_id: str, db: Session = Depends(get_db)):
    row = db.query(CandidateExtra).filter(CandidateExtra.candidate_id == candidate_id).first()
    return _extra_to_dict(row)


@app.post("/api/candidate-extra/{candidate_id}/resume")
async def upload_resume(
    candidate_id: str,
    file: UploadFile = File(...),
    uploaded_by: str = Form(""),
    db: Session = Depends(get_db),
):
    content = await file.read()
    if len(content) > MAX_RESUME_SIZE:
        raise HTTPException(400, "Файл больше 15 МБ — загрузите файл поменьше")

    row = _get_or_create_extra(db, candidate_id)
    row.resume_filename = file.filename
    row.resume_mimetype = file.content_type or "application/octet-stream"
    row.resume_data = content
    row.resume_uploaded_at = datetime.utcnow()
    row.resume_uploaded_by = uploaded_by or "не указано"
    db.commit()
    return _extra_to_dict(row)


def _content_disposition(disposition: str, filename: str) -> str:
    """
    HTTP-заголовки поддерживают только ASCII, а имена файлов резюме часто на
    кириллице (и могут содержать тире/эмодзи и т.п.). Отдаём ASCII-заглушку
    в filename= (для старых клиентов) и настоящее имя в filename* по RFC 5987
    (percent-encoded UTF-8) — так его правильно показывают все современные браузеры.
    """
    ascii_fallback = filename.encode("ascii", "ignore").decode("ascii").strip() or "resume"
    return f"{disposition}; filename=\"{ascii_fallback}\"; filename*=UTF-8''{quote(filename)}"


@app.get("/api/candidate-extra/{candidate_id}/resume")
def download_resume(candidate_id: str, db: Session = Depends(get_db)):
    row = db.query(CandidateExtra).filter(CandidateExtra.candidate_id == candidate_id).first()
    if not row or not row.resume_data:
        raise HTTPException(404, "Резюме не загружено")
    headers = {"Content-Disposition": _content_disposition("attachment", row.resume_filename)}
    return Response(content=row.resume_data, media_type=row.resume_mimetype, headers=headers)


@app.get("/api/candidate-extra/{candidate_id}/resume/view")
def view_resume(candidate_id: str, db: Session = Depends(get_db)):
    """Тот же файл, но с inline-заголовком — для встроенного просмотра PDF в браузере без скачивания."""
    row = db.query(CandidateExtra).filter(CandidateExtra.candidate_id == candidate_id).first()
    if not row or not row.resume_data:
        raise HTTPException(404, "Резюме не загружено")
    headers = {"Content-Disposition": _content_disposition("inline", row.resume_filename)}
    return Response(content=row.resume_data, media_type=row.resume_mimetype, headers=headers)


@app.delete("/api/candidate-extra/{candidate_id}/resume")
def delete_resume(candidate_id: str, db: Session = Depends(get_db)):
    row = db.query(CandidateExtra).filter(CandidateExtra.candidate_id == candidate_id).first()
    if row:
        row.resume_filename = None
        row.resume_mimetype = None
        row.resume_data = None
        row.resume_uploaded_at = None
        row.resume_uploaded_by = None
        db.commit()
    return _extra_to_dict(row)


class LinkPayload(BaseModel):
    url: str


@app.post("/api/candidate-extra/{candidate_id}/link")
def set_talantix_link(candidate_id: str, payload: LinkPayload, db: Session = Depends(get_db)):
    row = _get_or_create_extra(db, candidate_id)
    row.talantix_link = payload.url
    row.talantix_link_added_at = datetime.utcnow()
    db.commit()
    return _extra_to_dict(row)
