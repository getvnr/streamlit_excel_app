import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Sheet Comparison: Update Change Numbers")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)

        if 'Sheet1' not in xls.sheet_names or 'Results' not in xls.sheet_names:
            st.error("Excel must contain 'Sheet1' and 'Results' sheets")
        else:
            # Read sheets
            sheet1 = pd.read_excel(xls, sheet_name='Sheet1')
            results = pd.read_excel(xls, sheet_name='Results')

            # Automatically use first two columns from Sheet1
            sheet1_key_col = sheet1.columns[0]  # Server/host
            sheet1_value_col = sheet1.columns[1]  # Change number
            results_key_col = results.columns[0]  # Server/host

            # Show detected columns
            st.write("Sheet1 Server Column:", sheet1_key_col)
            st.write("Sheet1 Change Number Column:", sheet1_value_col)
            st.write("Results Server Column:", results_key_col)

            # Update 'UpdatedValue' in Results
            results['UpdatedValue'] = results[results_key_col].map(
                sheet1.set_index(sheet1_key_col)[sheet1_value_col]
            )

            # Replace unmatched with 'Not Found'
            results['UpdatedValue'] = results['UpdatedValue'].fillna("Not Found")

            # Show number of updates
            num_updated = (results['UpdatedValue'] != "Not Found").sum()
            total_servers = results.shape[0]

            st.metric("Number of servers updated", num_updated)
            st.metric("Total servers", total_servers)

            # Chart: Updated vs Not Found
            st.subheader("Update Summary")
            st.bar_chart(pd.DataFrame({
                "Count": [num_updated, total_servers - num_updated]
            }, index=["Updated", "Not Found"]))

            # Show updated Results
            st.subheader("Results with Updated Change Numbers")
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
