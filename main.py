import streamlit as st
from finance import finance_highlighter
from purchase import purchase_highlighter

def main():
    st.set_page_config(
        page_title="Finance and Purchase Highlighter"
    )

    st.title("🤖 Finbot")

    # Sidebar navigation
    page = st.sidebar.selectbox(
        "Select Page",
        ["Home", "Finance", "Purchase"]
    )

    if page == "Home":
        st.subheader("At your service")
        st.write("Please use the sidebar to navigate between Finance and Purchase pages.")
    elif page == "Finance":
        st.title("Finance Page")
        finance_highlighter()
    elif page == "Purchase":
        st.title("Purchase Page")
        purchase_highlighter()

if __name__ == "__main__":
    main()
