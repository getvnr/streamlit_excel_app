import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Sheet Comparison and Update")

# File upload
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # Read Excel file
    xls = pd.ExcelFile(uploaded_file)
    
    if 'Sheet1' not in xls.sheet_names or 'Results' not in xls.sheet_names:
        st.error("Excel must contain 'Sheet1' and 'Results' sheets")
    else:
        sheet1 = pd.read_excel(xls, sheet_name='Sheet1')
        results = pd.read_excel(xls, sheet_name='Results')
        
        # Ensure columns exist
        if sheet1.shape[1] < 2 or results.shape[1] < 1:
            st.error("Sheet1 must have at least 2 columns and Results must have at least 1 column")
        else:
            # Perform matching
            results['UpdatedValue'] = results['A'].map(sheet1.set_index('A')['B'])
            
            st.success("Comparison completed! 'UpdatedValue' column created in Results.")
            st.dataframe(results)
            
            # Provide download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
                results.to_excel(writer, sheet_name='Results', index=False)
            output.seek(0)
            
            st.download_button(
                label="Download Updated Excel",
                data=output,
                file_name="updated_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
