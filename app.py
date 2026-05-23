import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Sheet Navigator", layout="wide")

st.title("Excel Sheet Navigator")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the workbook to get sheet names
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        # Sidebar for sheet selection and search
        st.sidebar.header("Sheets")
        search_term = st.sidebar.text_input("Search Sheets")
        
        filtered_sheets = [s for s in sheet_names if search_term.lower() in s.lower()] if search_term else sheet_names
        
        selected_sheet = st.sidebar.selectbox("Select a sheet", filtered_sheets)
        
        if selected_sheet:
            st.header(f"Sheet: {selected_sheet}")
            
            # Read the selected sheet
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            
            # Data editing
            st.subheader("Edit Data")
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
            
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
