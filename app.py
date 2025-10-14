import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Server Update: Highlight Non-FQDN Servers and Chart by Solution")

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

            # Detect key columns
            sheet1_key_col = sheet1.columns[0]  # Server/host
            sheet1_value_col = sheet1.columns[1]  # Change number
            results_key_col = results.columns[0]  # Server/host
            solution_col = "Solution Name"  # Replace with your actual column name if different

            # Optional normalization for matching
            sheet1['normalized'] = sheet1[sheet1_key_col]
            results['normalized'] = results[results_key_col]

            # Merge Results with Sheet1 to get UpdatedValue
            updated_results = results.merge(
                sheet1[['normalized', sheet1_value_col]],
                left_on='normalized',
                right_on='normalized',
                how='left'
            ).rename(columns={sheet1_value_col: "UpdatedValue"})

            # Fill unmatched
            updated_results['UpdatedValue'] = updated_results['UpdatedValue'].fillna("Not Found")

            st.subheader("Results with UpdatedValue")
            st.dataframe(updated_results)

            # Count number of servers per Solution Name
            if solution_col in updated_results.columns:
                server_count = updated_results.groupby(solution_col)[results_key_col].nunique()
                st.subheader("Server Count per Solution Name")
                st.bar_chart(server_count)

            # Prepare Excel with highlighted non-FQDN servers
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                updated_results.to_excel(writer, sheet_name='Results', index=False)
                sheet1.to_excel(writer, sheet_name='Sheet1', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Results']

                # Yellow format for non-FQDN servers
                yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

                # Highlight only rows where server does NOT contain '.'
                for row_num, server in enumerate(updated_results[results_key_col], start=1):
                    if '.' not in str(server):  # simple check for non-FQDN
                        col_idx = updated_results.columns.get_loc(results_key_col)
                        worksheet.write(row_num, col_idx, server, yellow_format)

            output.seek(0)
            st.download_button(
                label="Download Excel with Highlights",
                data=output,
                file_name="highlighted_servers.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing file: {e}")
