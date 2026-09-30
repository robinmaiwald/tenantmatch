# Seite 2: Rolle wählen.
# Button "Tenant"   -> noch ohne Funktion (zeigt nur "Coming soon")
# Button "Landlord" -> Seite 3 (landlord.py)

import streamlit as st

st.title("Who are you?")
st.space("medium")

mieter, vermieter = st.columns(2)

with mieter:
    if st.button("Tenant", icon=":material/person:", width="stretch"):
        # Mieter-Bereich folgt später – bis dahin nur ein kurzer Hinweis.
        st.toast("Coming soon")

with vermieter:
    if st.button("Landlord", icon=":material/real_estate_agent:", type="primary", width="stretch"):
        st.switch_page("app_pages/landlord.py")

if st.button(
    "Back",
    icon=":material/arrow_back:",
):
    st.switch_page("app_pages/home.py")