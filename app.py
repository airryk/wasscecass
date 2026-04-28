import sys


def _is_running_via_streamlit():
    """Return True when this script is executed by Streamlit."""
    return "streamlit" in sys.modules

def main():
    import streamlit as st
    import index
    import data_analyzer
    import seating_arrangement

    st.title("WASSCE Student Data Tools")
    
    # Create a sidebar for navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.radio(
        "Select a tool:",
        ["Student Score Generator", "Student Data Analyzer", "Seating Arrangement"]
    )
    
    # Display the selected page
    if page == "Student Score Generator":
        # Run the score generator without its set_page_config
        index.run_app()
    elif page == "Seating Arrangement":
        # Run the seating arrangement tool
        seating_arrangement.run_app()

    else:
        # Run the data analyzer without its set_page_config
        data_analyzer.run_app()

if __name__ == "__main__":
    if not _is_running_via_streamlit():
        print("This is a Streamlit app. Run it with: python -m streamlit run app.py")
    else:
        main()
