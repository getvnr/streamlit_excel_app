import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# Page title
st.set_page_config(page_title="Excel Server Update Dashboard", layout="wide")
st.title("Excel Server Update — Match, Highlight & Visualize")

# File uploader
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)

        # Validate required sheets
        if 'Sheet1' not in xls.sheet_names or 'Results' not in xls.sheet_names:
            st.error("Excel must contain 'Sheet1' and 'Results' sheets.")
        else:
            # ✅ Read Sheet1 with NO header (first row is data)
            sheet1 = pd.read_excel(xls, sheet_name='Sheet1', header=None)
            results = pd.read_excel(xls, sheet_name='Results')

            # Manually assign column names for Sheet1
            sheet1.columns = ['Server', 'ChangeNumber']
            sheet1 = sheet1.dropna(subset=['Server'])

            # Define key columns
            results_key_col = results.columns[0]
            solution_col = "Solution Name" if "Solution Name" in results.columns else results.columns[1]

            # Normalize text
            sheet1['normalized'] = sheet1['Server'].astype(str).str.strip().str.lower()
            results['normalized'] = results[results_key_col].astype(str).str.strip().str.lower()

            # Merge based on normalized server names
            updated_results = results.merge(
                sheet1[['normalized', 'ChangeNumber']],
                on='normalized',
                how='left'
            ).rename(columns={'ChangeNumber': 'UpdatedValue'})

            updated_results['UpdatedValue'] = updated_results['UpdatedValue'].fillna("Not Found")

            # Filter matched servers only
            matched_servers = updated_results[updated_results['UpdatedValue'] != "Not Found"]

            # --- Summary ---
            total_servers = len(results)
            matched_count = len(matched_servers)
            unmatched_count = total_servers - matched_count

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Servers", total_servers)
            col2.metric("Matched Servers", matched_count)
            col3.metric("Unmatched Servers", unmatched_count)

            # --- Matched Servers Table ---
            st.subheader("Matched Servers")
            st.dataframe(matched_servers, use_container_width=True)

            # --- Visualization ---
            st.subheader("Visualizations")

            colA, colB = st.columns(2)

            # Chart 1: Solution Name vs Server Count (horizontal)
            if solution_col in matched_servers.columns:
                with colA:
                    soln_chart = matched_servers.groupby(solution_col)[results_key_col].nunique().reset_index()
                    soln_chart = soln_chart.rename(columns={results_key_col: "Server Count"})
                    fig1 = px.bar(
                        soln_chart,
                        x="Server Count",
                        y=solution_col,
                        orientation='h',
                        title="Server Count per Solution Name"
                    )
                    st.plotly_chart(fig1, use_container_width=True)

            # Chart 2: Change Number vs Server Count (horizontal)
            with colB:
                change_chart = matched_servers.groupby("UpdatedValue")[results_key_col].nunique().reset_index()
                change_chart = change_chart.rename(columns={results_key_col: "Server Count"})
                fig2 = px.bar(
                    change_chart,
                    x="Server Count",
                    y="UpdatedValue",
                    orientation='h',
                    title="Server Count per Change Number"
                )
                st.plotly_chart(fig2, use_container_width=True)

            # --- Excel Output with Highlights ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                updated_results.to_excel(writer, sheet_name='Results', index=False)
                sheet1.to_excel(writer, sheet_name='Sheet1', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Results']
                yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

                # Highlight non-FQDN servers (no dot in hostname)
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
