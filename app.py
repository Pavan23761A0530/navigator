import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Sheet Navigator", layout="wide")

st.title("Excel Sheet Navigator")

# Add some CSS to hide the "Failed to fetch" error if it's a persistent UI glitch, 
# though the config fix should handle the root cause.
st.markdown("""
    <style>
    .stAlert { margin-top: 1rem; }
    </style>
    """, unsafe_allow_html=True)

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    try:
        # Use a single loading mechanism
        file_bytes = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = xl.sheet_names
        
        # Sidebar for sheet selection and search
        st.sidebar.header("Sheets")
        search_term = st.sidebar.text_input("Search Sheets")
        
        filtered_sheets = [s for s in sheet_names if search_term.lower() in s.lower()] if search_term else sheet_names
        
        if not filtered_sheets:
            st.sidebar.warning("No sheets match your search.")
            selected_sheet = None
        else:
            selected_sheet = st.sidebar.selectbox("Select a sheet", filtered_sheets)
        
        if selected_sheet:
            st.header(f"Sheet: {selected_sheet}")
            
            # Read the selected sheet from the bytes buffer
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=selected_sheet)
            
            # Data editing
            st.subheader("Edit Data")
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="editor")
            
            # Export functionality
            st.subheader("Export")
            
            # Create a buffer to save the edited excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                edited_df.to_excel(writer, index=False, sheet_name=selected_sheet)
            
            st.download_button(
                label="Download Edited Sheet as Excel",
                data=output.getvalue(),
                file_name=f"{selected_sheet}_edited.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Additional info
            st.info("Note: You can edit cells directly in the table above and then download the result.")
            
    except Exception as e:
        st.error(f"Error loading file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
