import streamlit as st

st.set_page_config(layout="wide")

st.title("Route Status")

option = st.selectbox(
    "How would you like to be contacted?",
    ("Email", "Home phone", "Mobile phone"),
)

product_data = {
    "Product": [
        "Widget Pro",
        "Smart Device",
        "Diddy",
    ],
    "Category": [":blue[Electronics]", ":green[IoT]", ":violet[Bundle]"],
    "Stock": ["🟢 Full", "🟡 Low", "🔴 Empty"],
    "Units sold": [1247, 892, 654],
    "Stauts": [125000, 89000, 98000],
}
st.table(product_data, border="horizontal")

