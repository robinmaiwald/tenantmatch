import requests
import streamlit as st

API_URL = "http://127.0.0.1:8000"

st.title("Find a tenant")
st.caption("Tell me about the apartment and your requirements.")

# Initialize chat session
if "tenantmatch_session_id" not in st.session_state:
    try:
        response = requests.post(
            f"{API_URL}/landlord/interview/start",
            timeout=120,
        )
        response.raise_for_status()

        data = response.json()

        st.session_state.tenantmatch_session_id = data["session_id"]
        st.session_state.tenantmatch_messages = [
            {
                "role": "assistant",
                "content": data["message"],
            }
        ]
        st.session_state.tenantmatch_finished = data.get(
            "finished", False
        )
        st.session_state.tenantmatch_matches = data.get(
            "matches", []
        )

    except requests.RequestException as e:
        st.error(f"Backend nicht erreichbar: {e}")
        st.stop()


# Display conversation
for message in st.session_state.tenantmatch_messages:
    with st.chat_message(message["role"]):
        st.write(message["content"])


# Show results when interview is finished
if st.session_state.tenantmatch_finished:
    st.success("Die Bewerber wurden ausgewertet.")

    if st.session_state.tenantmatch_matches:
        st.subheader("Top 5 Bewerber")

        for match in st.session_state.tenantmatch_matches[:5]:
            applicant = match.get("bewerber", {})

            with st.container(border=True):
                st.markdown(
                    f"### #{match.get('rang', '?')} "
                    f"{applicant.get('first_name', '')} "
                    f"{applicant.get('last_name', '')}"
                )

                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric(
                        "Punkte",
                        match.get("punkte", "-"),
                    )

                with col2:
                    st.metric(
                        "Einkommen",
                        f"{applicant.get('monthly_net_income_eur', '-')} €",
                    )

                with col3:
                    st.metric(
                        "SCHUFA",
                        f"{applicant.get('schufa_score_percent', '-')} %",
                    )

                st.write(
                    f"**Beschäftigung:** "
                    f"{applicant.get('employment_status', '-')}"
                )

                st.write(
                    f"**Haushalt:** "
                    f"{applicant.get('household_size', '-')}"
                )

        if st.button(
            "View Top 5",
            icon=":material/emoji_events:",
            type="primary",
        ):
            st.switch_page("app_pages/top5.py")

    else:
        st.info("Keine Bewerbergebnisse erhalten.")

else:
    # Chat input
    user_message = st.chat_input("Ihre Antwort...")

    if user_message:
        st.session_state.tenantmatch_messages.append(
            {
                "role": "user",
                "content": user_message,
            }
        )

        try:
            response = requests.post(
                f"{API_URL}/landlord/interview",
                json={
                    "session_id": st.session_state.tenantmatch_session_id,
                    "message": user_message,
                },
                timeout=120,
            )

            response.raise_for_status()
            data = response.json()

        except requests.RequestException as e:
            st.error(f"Backend Fehler: {e}")
            st.stop()

        if data.get("message"):
            st.session_state.tenantmatch_messages.append(
                {
                    "role": "assistant",
                    "content": data["message"],
                }
            )

        if data.get("finished"):
            st.session_state.tenantmatch_finished = True
            st.session_state.tenantmatch_matches = data.get(
                "matches", []
            )

        st.rerun()


st.divider()

if st.button(
    "New search",
    icon=":material/refresh:",
):
    for key in [
        "tenantmatch_session_id",
        "tenantmatch_messages",
        "tenantmatch_finished",
        "tenantmatch_matches",
    ]:
        st.session_state.pop(key, None)

    st.rerun()
