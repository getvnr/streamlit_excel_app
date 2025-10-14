import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Server Update with Highlighted Matches")

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

            # Automatically detect columns
            sheet1_key_col = sheet1.columns[0]
            sheet1_value_col = sheet1.columns[1]
            results_key_col = results.columns[0]

            st.write("Sheet1 Server Column:", sheet1_key_col)
            st.write("Sheet1 Change Number Column:", sheet1_value_col)
            st.write("Results Server Column:", results_key_col)

            # OPTIONAL: normalize servers
            normalize = st.checkbox("Normalize server names (strip .next.loc)", value=False)
            if normalize:
                sheet1['normalized'] = sheet1[sheet1_key_col].str.replace(r'\.next\.loc$', '', regex=True)
                results['normalized'] = results[results_key_col].str.replace(r'\.next\.loc$', '', regex=True)
                key_col_merge = 'normalized'
            else:
                sheet1['normalized'] = sheet1[sheet1_key_col]
                results['normalized'] = results[results_key_col]
                key_col_merge = 'normalized'

            # Merge only matched servers (left join can expand multiple changes)
            updated_results = results.merge(
                sheet1[[key_col_merge, sheet1_value_col]],
                left_on=key_col_merge,
                right_on=key_col_merge,
                how='inner'
            )

            # Rename change number column to UpdatedValue
            updated_results = updated_results.rename(columns={sheet1_value_col: "UpdatedValue"})

            # Show metrics
            st.metric("Number of matched servers", updated_results.shape[0])

            st.subheader("Matched Results")
            st.dataframe(updated_results)

            # Prepare Excel for download with highlights
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Write original Sheet1
                sheet1.to_excel(writer, sheet_name='Sheet1', index=False)

                # Write Results with formatting
                updated_results.to_excel(writer, sheet_name='MatchedResults', index=False)
                
                workbook = writer.book
                worksheet = writer.sheets['MatchedResults']

                # Create yellow format
                yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

                # Highlight 'UpdatedValue' column in yellow
                col_idx = updated_results.columns.get_loc('UpdatedValue')
                for row_num in range(1, updated_results.shape[0]+1):
                    worksheet.write(row_num, col_idx, updated_results.iloc[row_num-1, col_idx], yellow_format)

            output.seek(0)
            st.download_button(
                label="Download Matched Excel",
                data=output,
                file_name="matched_servers.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing file: {e}")
