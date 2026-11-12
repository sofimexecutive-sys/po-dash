import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

# Railway Postgres задаёт DATABASE_URL автоматически при подключении плагина Postgres.
# Локально (и если плагин не подключён) используем SQLite-файл рядом с кодом.
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./tracker.db")

# Railway/Heroku иногда отдают "postgres://", а SQLAlchemy 2.x требует "postgresql://"
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, connect_args=connect_args)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


def run_light_migrations():
    """
    Мы не используем Alembic — для маленького проекта это лишний вес. Вместо
    этого: Base.metadata.create_all() создаёт таблицы, которых ещё нет, но
    НЕ добавляет новые колонки в уже существующие (например, на проде, где
    таблица votes уже есть с реальными данными). Эта функция добивает такие
    точечные ALTER TABLE вручную — безопасно: если колонка уже есть, просто
    ловим ошибку и идём дальше, ничего не теряя.
    """
    from sqlalchemy import text
    statements = [
        "ALTER TABLE votes ADD COLUMN reaction VARCHAR DEFAULT 'like'",
    ]
    for stmt in statements:
        try:
            with engine.begin() as conn:
                conn.execute(text(stmt))
        except Exception:
            pass  # колонка уже существует (или таблицы ещё нет — её создаст create_all)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
