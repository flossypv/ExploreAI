
from __future__ import annotations

import os
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, DateTime, Text
from sqlalchemy.orm import declarative_base, sessionmaker

DB_PATH = os.getenv('RAG_DB_PATH', 'rag_report.db')
ENGINE = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=ENGINE, autocommit=False, autoflush=False)
Base = declarative_base()


class Project(Base):
    __tablename__ = 'projects'

    id = Column(Integer, primary_key=True, index=True)

    client_name = Column(String(255), nullable=False)
    project_owner = Column(String(255), nullable=False)
    project_name = Column(String(255), nullable=False)
    budget_po_name = Column(String(255), nullable=True)
    schedule_timeline = Column(String(255), nullable=True)

    rag_budget_cost = Column(String(10), nullable=False, default='Green')
    rag_scope_requirement = Column(String(10), nullable=False, default='Green')
    rag_quality = Column(String(10), nullable=False, default='Green')
    rag_risk = Column(String(10), nullable=False, default='Green')
    rag_resource_team = Column(String(10), nullable=False, default='Green')
    rag_mohi = Column(String(50), nullable=True)  # free text / misc tag

    notes = Column(Text, nullable=True)

    updated_by = Column(String(255), nullable=True)
    updated_at = Column(DateTime, nullable=False, default=datetime.utcnow)


def init_db():
    Base.metadata.create_all(bind=ENGINE)


def get_session():
    return SessionLocal()
