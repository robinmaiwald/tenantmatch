"""
Database models for TenantMatch.

Defines the three main database entities:

- ApplicantDB: information about applicants.
- ListingDB: basic information about apartments.
- CriterionDB: landlord requirements and preferences
  associated with a listing.

These are SQLAlchemy database models. API/Pydantic models
are defined separately in models.py.
"""

from sqlalchemy import Boolean, Float, ForeignKey, Integer, String
from sqlalchemy.orm import Mapped, mapped_column

from backend.app.database import Base


class ApplicantDB(Base):
    __tablename__ = "applicants"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    name: Mapped[str] = mapped_column(String, nullable=False)

    income: Mapped[float | None] = mapped_column(Float, nullable=True)
    employment_contract_type: Mapped[str | None] = mapped_column(
        String,
        nullable=True,
    )
    household_size: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )
    desired_rooms: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )
    rent: Mapped[float | None] = mapped_column(Float, nullable=True)

    mietschuldenfreiheitsbescheinigung: Mapped[bool | None] = mapped_column(
        Boolean,
        nullable=True,
    )
    smoking: Mapped[bool | None] = mapped_column(
        Boolean,
        nullable=True,
    )

    application_text: Mapped[str | None] = mapped_column(
        String,
        nullable=True,
    )


class ListingDB(Base):
    __tablename__ = "listings"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)

    title: Mapped[str] = mapped_column(String, nullable=False)

    rent: Mapped[float] = mapped_column(Float, nullable=False)
    rooms: Mapped[int] = mapped_column(Integer, nullable=False)


class CriterionDB(Base):
    __tablename__ = "criteria"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)

    listing_id: Mapped[int] = mapped_column(
        ForeignKey("listings.id"),
        nullable=False,
    )

    name: Mapped[str] = mapped_column(String, nullable=False)

    value: Mapped[str] = mapped_column(
        String,
        nullable=False,
    )

    importance: Mapped[str] = mapped_column(
        String,
        nullable=False,
    )

    required: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
        nullable=False,
    )

    explanation: Mapped[str | None] = mapped_column(
        String,
        nullable=True,
    )

    source: Mapped[str] = mapped_column(
        String,
        default="landlord",
        nullable=False,
    )