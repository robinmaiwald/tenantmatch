"""
Database configuration for TenantMatch.

This module is responsible only for connecting the application
to the database and providing SQLAlchemy sessions.

Database models themselves live in database_models.py.
"""

from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, sessionmaker


DATABASE_URL = "sqlite:///./tenantmatch.db"


engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False},
)


class Base(DeclarativeBase):
    pass


SessionLocal = sessionmaker(
    bind=engine,
    autoflush=False,
    autocommit=False,
)


def get_db():
    """
    Provide a database session to the caller.

    The session is automatically closed after the caller is finished.
    FastAPI routes can use this as a dependency.
    """
    db = SessionLocal()

    try:
        yield db
    finally:
        db.close()