# Seite 1: Startseite – großes, mittiges "Welcome".
# Button "Get started" -> Seite 2 (role.py)

import streamlit as st

# Etwas Abstand nach oben, damit "Welcome" mittig auf der Seite steht.
st.space("large")
st.space("large")

st.markdown("<h1 style='text-align: center; font-size: 5rem;'>Welcome</h1>",
            unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.3rem;'>KiezMove</p>",
            unsafe_allow_html=True)

st.space("medium")

# Button mittig: drei Spalten, der Button steht in der mittleren.
links, mitte, rechts = st.columns([1, 1, 1])
with mitte:
    if st.button("Get started", type="primary", width="stretch"):
        st.switch_page("app_pages/role.py")
