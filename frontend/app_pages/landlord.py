# Seite 3: Bereich für Vermieter.
# Button "Apartments"    -> Seite 4 (apartments.py)
# Button "Applicants"    -> Seite 5 (applicants.py)
# Button "Find a tenant" -> Seite 6 (find_tenant.py)

import streamlit as st

st.title("Landlord")
st.space("medium")

wohnungen, bewerber, finden = st.columns(3)

with wohnungen:
    if st.button("Apartments", icon=":material/apartment:", width="stretch"):
        st.switch_page("app_pages/apartments.py")

with bewerber:
    if st.button("Applicants", icon=":material/badge:", width="stretch"):
        st.switch_page("app_pages/applicants.py")

with finden:
    if st.button("Find a tenant", icon=":material/search:", type="primary", width="stretch"):
        st.switch_page("app_pages/find_tenant.py")
