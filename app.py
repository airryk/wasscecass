import streamlit as st

# Set page config at the very beginning - this must be the first Streamlit command
# st.set_page_config(page_title="WASSCE Student Data Tools", layout="wide")

# Import other modules after setting the page config
import index
import data_analyzer

def main():
    st.title("WASSCE Student Data Tools")
    
    # Create a sidebar for navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.radio(
        "Select a tool:",
        ["Student Score Generator", "Student Data Analyzer"]
    )
    
    # Display the selected page
    if page == "Student Score Generator":
        # Run the score generator without its set_page_config
        index.run_app()
    else:
        # Run the data analyzer without its set_page_config
        data_analyzer.run_app()

if __name__ == "__main__":
    main()
