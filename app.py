import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt

st.title("Excel Sheet Comparison and Update with Analytics")

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

            # Detect columns
            sheet1_key_col = sheet1.columns[0]
            sheet1_value_col = sheet1.columns[1]
            results_key_col = results.columns[0]

            st.write(f"Matching Sheet1 column: {sheet1_key_col}")
            st.write(f"Value Sheet1 column: {sheet1_value_col}")
            st.write(f"Results matching column: {results_key_col}")

            # Update 'UpdatedValue' column
            results['UpdatedValue'] = results[results_key_col].map(
                sheet1.set_index(sheet1_key_col)[sheet1_value_col]
            )

            # Count number of changes
            num_changes = results['UpdatedValue'].notna().sum()
            total_rows = results.shape[0]
            st.metric("Number of updates", num_changes)
            st.metric("Total rows / servers", total_rows)

            # Display updated dataframe
            st.subheader("Updated Results Sheet")
            st.dataframe(results)

            # Chart: Number of changes vs unchanged
            fig, ax = plt.subplots()
            counts = [num_changes, total_rows - num_changes]
            ax.bar(["Updated", "Unchanged"], counts, color=["green", "red"])
            ax.set_ylabel("Number of rows")
            ax.set_title("Changes Summary")
            st.pyplot(fig)

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

    except Exception as e:
        st.error(f"Error processing file: {e}")
