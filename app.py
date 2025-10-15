import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Excel Server Update Dashboard", layout="wide")
st.title("Excel Server Update — Match & Highlight Servers")

# --- File uploader ---
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)

        # Check required sheets
        if 'Sheet1' not in xls.sheet_names or 'Results' not in xls.sheet_names:
            st.error("Excel must contain 'Sheet1' and 'Results' sheets.")
        else:
            # --- Read Sheet1 WITHOUT header ---
            sheet1 = pd.read_excel(xls, sheet_name='Sheet1', header=None)
            results = pd.read_excel(xls, sheet_name='Results')

            # Assign column names manually for Sheet1
            sheet1.columns = ['Server', 'ChangeNumber']
            sheet1 = sheet1.dropna(subset=['Server'])

            # Clean Results column names
            results.columns = results.columns.str.strip().str.replace('\n', ' ')

            # Inspect columns
            st.write("Columns in Results sheet:", results.columns.tolist())

            results_key_col = results.columns[0]  # first column assumed as server/host

            # Optional Solution column
            solution_col = next((c for c in ["Solution Name", "SolutionName", "Solution"] if c in results.columns), None)

            # Optional Start Date column
            start_date_col = next((c for c in ["Start Date", "StartDate"] if c in results.columns), None)
            if start_date_col:
                results[start_date_col] = pd.to_datetime(results[start_date_col], errors='coerce')

            # Normalize for merge
            sheet1['normalized'] = sheet1['Server'].astype(str).str.strip().str.lower()
            results['normalized'] = results[results_key_col].astype(str).str.strip().str.lower()

            # Merge to get UpdatedValue
            updated_results = results.merge(
                sheet1[['normalized', 'ChangeNumber']],
                on='normalized',
                how='left'
            ).rename(columns={'ChangeNumber': 'UpdatedValue'})

            updated_results['UpdatedValue'] = updated_results['UpdatedValue'].fillna("Not Found")

            # Matched / Unmatched
            matched_servers = updated_results[updated_results['UpdatedValue'] != "Not Found"]
            unmatched_servers = updated_results[updated_results['UpdatedValue'] == "Not Found"]

            # --- Metrics ---
            total_servers = len(results)
            matched_count = len(matched_servers)
            unmatched_count = total_servers - matched_count

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Servers", total_servers)
            col2.metric("Matched Servers", matched_count)
            col3.metric("Unmatched Servers", unmatched_count)

            # --- Display matched servers ---
            st.subheader("Matched Servers")
            st.dataframe(matched_servers, use_container_width=True)

            # --- Visualizations ---
            st.subheader("Visualizations")
            colA, colB = st.columns(2)

            # Solution Name vs Server Count
            if solution_col:
                sol_count = matched_servers.groupby(solution_col)[results_key_col].nunique().reset_index()
                sol_count = sol_count.rename(columns={results_key_col: "Server Count"})
                fig1 = px.bar(sol_count, x="Server Count", y=solution_col, orientation='h', title="Server Count per Solution Name")
                colA.plotly_chart(fig1, use_container_width=True)

            # Change Number vs Server Count
            chg_count = matched_servers.groupby("UpdatedValue")[results_key_col].nunique().reset_index()
            chg_count = chg_count.rename(columns={results_key_col: "Server Count"})
            fig2 = px.bar(chg_count, x="Server Count", y="UpdatedValue", orientation='h', title="Server Count per Change Number")
            colB.plotly_chart(fig2, use_container_width=True)

            # --- Unmatched servers ---
            with st.expander("View Unmatched Servers"):
                st.dataframe(unmatched_servers[[results_key_col, "UpdatedValue"]], use_container_width=True)

            # --- Excel download with highlights ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                updated_results.to_excel(writer, sheet_name='Results', index=False)
                sheet1.to_excel(writer, sheet_name='Sheet1', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Results']
                yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

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
