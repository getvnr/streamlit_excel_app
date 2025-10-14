import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Server Update: Matched Servers with Highlights and Chart")

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

            # Columns
            sheet1_key_col = sheet1.columns[0]  # Server/host
            sheet1_value_col = sheet1.columns[1]  # Change number
            results_key_col = results.columns[0]  # Server/host
            solution_col = "Solution Name"  # Replace with your actual column name

            # Normalize columns if needed
            sheet1['normalized'] = sheet1[sheet1_key_col]
            results['normalized'] = results[results_key_col]

            # Merge to get UpdatedValue
            updated_results = results.merge(
                sheet1[['normalized', sheet1_value_col]],
                left_on='normalized',
                right_on='normalized',
                how='left'
            ).rename(columns={sheet1_value_col: "UpdatedValue"})

            # Fill unmatched
            updated_results['UpdatedValue'] = updated_results['UpdatedValue'].fillna("Not Found")

            # Filter only matched servers
            matched_servers = updated_results[updated_results['UpdatedValue'] != "Not Found"]

            st.subheader("Matched Servers with UpdatedValue")
            st.dataframe(matched_servers)

            # Chart: Server Count per Solution Name for matched servers
            if solution_col in matched_servers.columns:
                server_count = matched_servers.groupby(solution_col)[results_key_col].nunique()
                st.subheader("Server Count per Solution Name (Matched Only)")
                st.bar_chart(server_count)

            # Prepare Excel with highlights for non-FQDN servers
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                updated_results.to_excel(writer, sheet_name='Results', index=False)
                sheet1.to_excel(writer, sheet_name='Sheet1', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Results']

                yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

                # Highlight non-FQDN servers (without ".") in the server column
                for row_num, server in enumerate(updated_results[results_key_col], start=1):
                    if '.' not in str(server):
                        col_idx = updated_results.columns.get_loc(results_key_col)
                        worksheet.write(row_num, col_idx, server, yellow_format)

            output.seek(0)
            st.download_button(
                label="Download Excel with Highlights",
                data=output,
                file_name="matched_highlighted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing file: {e}")
