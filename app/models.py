from sqlalchemy import Column, Integer, String, Text, DateTime, UniqueConstraint, LargeBinary
from sqlalchemy.sql import func
from .db import Base


class Vote(Base):
    """
    Одна реакция от одного человека по одному кандидату: 'like' или 'dislike'.
    До введения дизлайков тут были только лайки — существующие строки
    мигрируются в reaction='like' (см. _run_migrations в db.py/main.py),
    так что старые данные не теряются.
    """
    __tablename__ = "votes"

    id = Column(Integer, primary_key=True)
    candidate_id = Column(String, nullable=False, index=True)
    person = Column(String, nullable=False)
    reaction = Column(String, nullable=False, server_default="like")  # 'like' | 'dislike'
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    __table_args__ = (
        UniqueConstraint("candidate_id", "person", name="uq_candidate_person"),
    )


class CandidateExtra(Base):
    """
    Данные по кандидату, которых нет в отчётах Talantix и которые команда
    добавляет вручную: файл резюме и ссылка на карточку в Talantix.
    Один ряд на кандидата (candidate_id уникален и переживает перезагрузку
    отчёта, пока ID кандидата не меняется между версиями файла).
    """
    __tablename__ = "candidate_extra"

    id = Column(Integer, primary_key=True)
    candidate_id = Column(String, nullable=False, unique=True, index=True)

    resume_filename = Column(String, nullable=True)
    resume_mimetype = Column(String, nullable=True)
    resume_data = Column(LargeBinary, nullable=True)
    resume_uploaded_at = Column(DateTime(timezone=True), nullable=True)
    resume_uploaded_by = Column(String, nullable=True)

    talantix_link = Column(Text, nullable=True)
    talantix_link_added_at = Column(DateTime(timezone=True), nullable=True)


class CandidateComment(Base):
    """
    Комментарий заказчика по кандидату (Иван/Никита/Павел) — отдельно от
    комментариев рекрутёра из Talantix. Привязан к candidate_id, поэтому не
    теряется при смене этапа и переживает обновление отчёта.
    """
    __tablename__ = "candidate_comments"

    id = Column(Integer, primary_key=True)
    candidate_id = Column(String, nullable=False, index=True)
    person = Column(String, nullable=False)
    text = Column(Text, nullable=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now())


class CandidateScore(Base):
    """
    Две 10-балльные оценки заказчика по кандидату на этапе «Пристрелка».
    Два фиксированных оценщика — Дима и Никита. Оценки живут отдельно от комментариев
    и сохраняются между обновлениями Talantix-отчёта.
    """
    __tablename__ = "candidate_scores"

    id = Column(Integer, primary_key=True)
    candidate_id = Column(String, nullable=False, index=True)
    person = Column(String, nullable=False)
    profile_score = Column(Integer, nullable=True)
    company_score = Column(Integer, nullable=True)
    updated_at = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now())

    __table_args__ = (
        UniqueConstraint("candidate_id", "person", name="uq_candidate_score_person"),
    )


class Snapshot(Base):
    """
    Последняя загруженная и распарсенная версия отчёта.
    Храним как JSON-текст, чтобы не городить отдельную таблицу кандидатов —
    список кандидатов целиком заменяется при каждом обновлении отчёта.
    """
    __tablename__ = "snapshots"

    id = Column(Integer, primary_key=True)
    source_url = Column(Text, nullable=True)
    data_json = Column(Text, nullable=False)
    fetched_at = Column(DateTime(timezone=True), server_default=func.now())


class AnalyticsSnapshot(Base):
    """
    Последняя собранная версия аналитики (воронка, причины отказа, среднее
    время на этапе) — из отчётов "Аналитика воронки" + "Процесс работы".
    """
    __tablename__ = "analytics_snapshots"

    id = Column(Integer, primary_key=True)
    data_json = Column(Text, nullable=False)
    fetched_at = Column(DateTime(timezone=True), server_default=func.now())
