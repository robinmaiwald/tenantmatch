# Seite 5: Bewerber – Platzhalter, Inhalt folgt später.

import streamlit as st

st.title("Applicants")
st.info("Coming soon: all applications will be listed here.", icon=":material/construction:")

if st.button(
    "Back",
    icon=":material/arrow_back:",
):
    st.switch_page("app_pages/role.py")