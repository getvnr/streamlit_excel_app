import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Sheet Comparison and Update (Auto Column Detection)")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read Excel file
        xls = pd.ExcelFile(uploaded_file)
        
        if 'Sheet1' not in xls.sheet_names or 'Results' not in xls.sheet_names:
            st.error("Excel must contain 'Sheet1' and 'Results' sheets")
        else:
            # Try reading sheets with headers
            sheet1 = pd.read_excel(xls, sheet_name='Sheet1')
            results = pd.read_excel(xls, sheet_name='Results')

            # Automatically detect first two columns from Sheet1
            sheet1_col_key = sheet1.columns[0]
            sheet1_col_value = sheet1.columns[1]
            
            results_col_key = results.columns[0]

            st.write("Matching Sheet1 column:", sheet1_col_key)
            st.write("Value Sheet1 column:", sheet1_col_value)
            st.write("Results matching column:", results_col_key)

            # Perform matching
            results['UpdatedValue'] = results[results_col_key].map(
                sheet1.set_index(sheet1_col_key)[sheet1_col_value]
            )

            st.success("Comparison completed! 'UpdatedValue' column created in Results.")
            st.dataframe(results)

            # Provide download button
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

    except Exception as e:
        st.error(f"Error processing file: {e}")
