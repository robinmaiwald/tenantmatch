"""
FastAPI entry point for TenantMatch.
"""

import uuid

from fastapi import FastAPI

from backend.app.models import LandlordInterviewRequest

from backend.app.services.vermieter_assistent import (
    gespraech_schritt,
    gespraech_starten,
    ki_client,
    bewerber_laden,
    bewerber_bewerten,
)


app = FastAPI(
    title="TenantMatch",
    description="Fair tenant matching for landlords.",
)


# One conversation state per landlord session.
interview_sessions = {}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/landlord/interview/start")
def start_interview():
    session_id = str(uuid.uuid4())

    state = gespraech_starten()

    # Store the conversation state.
    interview_sessions[session_id] = state

    # Let Gemini generate the first question.
    client = ki_client()

    message, state = gespraech_schritt(
        client,
        state,
    )

    interview_sessions[session_id] = state

    return {
        "session_id": session_id,
        "message": message,
        "finished": state["fertig"],
        "criteria": state["anforderungen"],
        "wohnung": state["wohnung"],
    }


@app.post("/landlord/interview")
def landlord_interview(request: LandlordInterviewRequest):
    state = interview_sessions.get(request.session_id)

    if state is None:
        return {
            "error": "Interview session not found."
        }

    client = ki_client()

    message, state = gespraech_schritt(
        client,
        state,
        request.message,
    )

    interview_sessions[request.session_id] = state
    matches = []

    if state["fertig"]:
        applicants = bewerber_laden()

        results = bewerber_bewerten(
            applicants,
            state["wohnung"],
            state["anforderungen"],
        )

        matches = results[:5]

    return {
        "session_id": request.session_id,
        "message": message,
        "finished": state["fertig"],
        "criteria": state["anforderungen"],
        "wohnung": state["wohnung"],
        "matches": matches[:5],
    }