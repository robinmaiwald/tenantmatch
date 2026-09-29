import requests
import streamlit as st


API_URL = "http://127.0.0.1:8000"


st.set_page_config(
    page_title="TenantMatch",
    page_icon="🏠",
)


st.title("🏠 TenantMatch")
st.subheader("Vermieter-Assistent")

st.write(
    "Beschreiben Sie Ihre Wohnung und Ihre Anforderungen. "
    "Ich helfe Ihnen passende Bewerber zu finden."
)


# -------------------------
# Start session
# -------------------------

if "session_id" not in st.session_state:

    response = requests.post(
        f"{API_URL}/landlord/interview/start"
    )

    response.raise_for_status()

    data = response.json()

    st.session_state.session_id = data["session_id"]
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": data["message"],
        }
    ]

    st.session_state.finished = False
    st.session_state.matches = []


# -------------------------
# Show conversation
# -------------------------

for message in st.session_state.messages:

    with st.chat_message(message["role"]):
        st.write(message["content"])


# -------------------------
# Show results
# -------------------------

if st.session_state.finished:

    st.divider()

    st.subheader("🏆 Top Bewerber")

    for match in st.session_state.matches:

        applicant = match["bewerber"]

        with st.container():

            st.markdown(
                f"""
                ### #{match['rang']} 
                {applicant.get('first_name','')} {applicant.get('last_name','')}

                **Punkte:** {match['punkte']}

                **Einkommen:** {applicant.get('monthly_net_income_eur')} €

                **Beschäftigung:** {applicant.get('employment_status')}

                **SCHUFA:** {applicant.get('schufa_score_percent')} %

                **Haushalt:** {applicant.get('household_size')} Personen
                """
            )


# -------------------------
# Chat input
# -------------------------

if not st.session_state.finished:

    landlord_message = st.chat_input(
        "Ihre Antwort..."
    )


    if landlord_message:

        st.session_state.messages.append(
            {
                "role": "user",
                "content": landlord_message,
            }
        )


        response = requests.post(
            f"{API_URL}/landlord/interview",
            json={
                "session_id": st.session_state.session_id,
                "message": landlord_message,
            },
        )


        if not response.ok:

            st.error(
                f"Backend Fehler {response.status_code}"
            )

            st.stop()


        data = response.json()


        if data.get("message"):

            st.session_state.messages.append(
                {
                    "role": "assistant",
                    "content": data["message"],
                }
            )


        if data.get("finished"):

            st.session_state.finished = True

            st.session_state.matches = data.get(
                "matches",
                []
            )


        st.rerun()