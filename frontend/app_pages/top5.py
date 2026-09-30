import streamlit as st

st.title("Top 5")
st.caption("The five best matching applicants for your apartment.")

matches = st.session_state.get(
    "tenantmatch_matches",
    [],
)

if not matches:
    st.info(
        "No matching results yet. Start a tenant search first.",
        icon=":material/info:",
    )

    if st.button(
        "Find a tenant",
        icon=":material/search:",
        type="primary",
    ):
        st.switch_page("app_pages/find_tenant.py")

else:
    for match in matches[:5]:
        applicant = match.get("bewerber", {})

        rank = match.get("rang", "?")
        first_name = applicant.get("first_name", "")
        last_name = applicant.get("last_name", "")

        with st.container(border=True):
            st.markdown(
                f"## #{rank} — {first_name} {last_name}"
            )

            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric(
                    "Punkte",
                    match.get("punkte", "-"),
                )

            with col2:
                income = applicant.get(
                    "monthly_net_income_eur"
                )
                st.metric(
                    "Net income",
                    f"{income} €" if income is not None else "-",
                )

            with col3:
                schufa = applicant.get(
                    "schufa_score_percent"
                )
                st.metric(
                    "SCHUFA",
                    f"{schufa} %" if schufa is not None else "-",
                )

            st.write(
                f"**Employment:** "
                f"{applicant.get('employment_status', '-')}"
            )

            st.write(
                f"**Household:** "
                f"{applicant.get('household_size', '-')}"
            )

            st.divider()

    if st.button(
        "Back to search",
        icon=":material/arrow_back:",
    ):
        st.switch_page("app_pages/find_tenant.py")