# =============================================================================
#  KiezMove – Frontend (Streamlit)
# =============================================================================
#  Einstiegspunkt der App. Hier werden alle 7 Seiten registriert.
#  Jede Seite liegt als eigene Datei in app_pages/.
#
#  Navigation: Es gibt bewusst KEIN Menü (position="hidden").
#  Man wechselt die Seite nur über die Buttons auf den Seiten
#  (st.switch_page). Die Oberfläche ist auf Englisch.
#
#  Starten:  .venv\Scripts\streamlit run streamlit_app.py
# =============================================================================

import streamlit as st

# Die 7 Seiten in Ablaufreihenfolge. Die erste Seite ist die Startseite.
seiten = [
    st.Page("app_pages/home.py", title="Welcome", icon=":material/home:", default=True),
    st.Page("app_pages/role.py", title="Choose your role", icon=":material/group:"),
    st.Page("app_pages/landlord.py", title="Landlord", icon=":material/real_estate_agent:"),
    st.Page("app_pages/apartments.py", title="Apartments", icon=":material/apartment:"),
    st.Page("app_pages/applicants.py", title="Applicants", icon=":material/badge:"),
    st.Page("app_pages/find_tenant.py", title="Find a tenant", icon=":material/chat:"),
    st.Page("app_pages/top5.py", title="Top 5", icon=":material/emoji_events:"),
]

seite = st.navigation(seiten, position="hidden")
seite.run()
