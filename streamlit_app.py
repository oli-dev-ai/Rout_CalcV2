import streamlit as st

st.logo("logo2.png",size="large", icon_image="logo2.png")

#definiowanie stron
home = st.Page("home.py", title="Home", icon=":material/home:")
dashboard = st.Page("dashboard.py", title="Dashboard", icon=":material/dashboard:")
route_calculator = st.Page("routcalc.py", title="Route Calculator", icon=":material/calculate:")
route_status = st.Page("route_status.py", title="Route Status", icon=":material/garage_check:")
converter = st.Page("converter.py", title="Converter", icon=":material/table_convert:")


pg = st.navigation([home, dashboard, route_calculator, route_status, converter])
pg.run()


