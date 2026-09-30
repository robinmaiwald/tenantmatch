# Seite 4: Wohnungen des Vermieters – Platzhalter, Inhalt folgt später.

import streamlit as st

st.title("Apartments")
st.info("Coming soon: your apartments will be listed here.", icon=":material/construction:")

if st.button(
    "Back",
    icon=":material/arrow_back:",
):
    st.switch_page("app_pages/landlord.py")