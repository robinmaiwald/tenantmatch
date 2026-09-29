"""
Pydantic models for TenantMatch.

These models define the data exchanged between the application,
API endpoints, and services.

They are separate from the SQLAlchemy database models in
database_models.py and are used for validation and structured
data handling.
"""

from pydantic import BaseModel, Field


class Applicant(BaseModel):
    name: str

    income: float | None = None

    employment_contract_type: str | None = None

    household_size: int | None = Field(
        default=None,
        ge=1,
    )

    desired_rooms: int | None = Field(
        default=None,
        ge=1,
    )

    rent: float | None = Field(
        default=None,
        ge=0,
    )

    mietschuldenfreiheitsbescheinigung: bool | None = None

    smoking: bool | None = None

    application_text: str | None = None


class Criterion(BaseModel):
    name: str
    value: str
    importance: str
    required: bool = False
    explanation: str | None = None
    source: str = "landlord"


class Listing(BaseModel):
    title: str

    rent: float = Field(
        ge=0,
    )

    rooms: int = Field(
        ge=1,
    )

    criteria: list[Criterion] = []


class MatchResult(BaseModel):
    applicant_id: int
    criteria_met: list[str]
    criteria_not_met: list[str]
    missing_information: list[str]




class LandlordInterviewRequest(BaseModel):
    session_id: str
    message: str