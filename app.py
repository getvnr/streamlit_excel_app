import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# =========================================================
# 🧭 Streamlit Page Configuration
# =========================================================
st.set_page_config(page_title="Excel Server Update Dashboard", layout="wide")
st.title("📊 Excel Server Update — Match, Highlight & Visualize")
st.markdown("Upload an Excel file containing **Sheet1** and **Results** to begin analysis.")

# =========================================================
# 📂 File Upload Section
# =========================================================
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read Excel File
        xls = pd.ExcelFile(uploaded_file)

        # =========================================================
        # ✅ Sheet Validation
        # =========================================================
        required_sheets = {'Sheet1', 'Results'}
        if not required_sheets.issubset(xls.sheet_names):
            st.error("❌ The uploaded Excel must contain both **'Sheet1'** and **'Results'** sheets.")
        else:
            # =========================================================
            # 📄 Read and Prepare Data
            # =========================================================
            sheet1 = pd.read_excel(xls, sheet_name='Sheet1', header=None)
            results = pd.read_excel(xls, sheet_name='Results')

            # Assign headers manually for Sheet1
            sheet1.columns = ['Server', 'ChangeNumber']
            sheet1 = sheet1.dropna(subset=['Server'])

            # Identify key columns
            results_key_col = results.columns[0]
            solution_col = "Solution Name" if "Solution Name" in results.columns else results.columns[1]

            # =========================================================
            # 🧹 Normalize Hostnames for Matching
            # =========================================================
            def normalize_hostname(name):
                name = str(name).strip().lower()
                if '.' in name:
                    name = name.split('.')[0]
                return name

            sheet1['normalized'] = sheet1['Server'].apply(normalize_hostname)
            results['normalized'] = results[results_key_col].apply(normalize_hostname)

            # =========================================================
            # 🔄 Merge and Match Data
            # =========================================================
            updated_results = (
                results.merge(
                    sheet1[['normalized', 'ChangeNumber']],
                    on='normalized',
                    how='left'
                )
                .rename(columns={'ChangeNumber': 'UpdatedValue'})
            )

            updated_results['UpdatedValue'] = updated_results['UpdatedValue'].fillna("Not Found")

            # Filter matched results
            matched_servers = updated_results[updated_results['UpdatedValue'] != "Not Found"].copy()

            # =========================================================
            # 🧾 Create 'Matched Servers' Sheet Data
            # =========================================================
            # Define fixed headers
            fixed_headers = [
                "CHG", "System Name", "Start Time", "End Time", "Before", "After", "Status"
            ]

            # Dynamic headers — pick columns from Results in the given order if they exist
            results_headers = [
                "System Name", "Solution Name", "Instance Name", "Delivery Instance Owner", "Business Name",
                "Instance Status", "Instance Service Level", "Instance Environment", "Sub Business Name",
                "DC Name", "DC Country", "KPE Name", "KPE Short Name", "Software/Firmware Version",
                "Environment", "Impact", "OS Class", "OS Version", "Usage", "Detailed Usage",
                "IP Type", "IP Address"
            ]

            available_result_columns = [col for col in results_headers if col in matched_servers.columns]
            all_headers = fixed_headers + available_result_columns

            # Build final DataFrame
            matched_servers_sheet = pd.DataFrame(columns=all_headers)

            # Fill the initial columns
            matched_servers_sheet["CHG"] = matched_servers["UpdatedValue"]
            matched_servers_sheet["System Name"] = matched_servers[results_key_col]

            # Copy available columns from results
            for col in available_result_columns:
                matched_servers_sheet[col] = matched_servers.get(col, "")

            # =========================================================
            # 📊 Summary Metrics
            # =========================================================
            total_servers = len(results)
            matched_count = len(matched_servers)
            unmatched_count = total_servers - matched_count

            st.markdown("---")
            st.subheader("📈 Summary Overview")

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Servers", total_servers)
            col2.metric("Matched Servers", matched_count)
            col3.metric("Unmatched Servers", unmatched_count)

            # =========================================================
            # 🧾 Matched Servers Table
            # =========================================================
            st.markdown("---")
            st.subheader("✅ Matched Servers (Detailed)")
            st.dataframe(matched_servers_sheet, use_container_width=True)

            # =========================================================
            # 📉 Visualization Section
            # =========================================================
            st.markdown("---")
            st.subheader("📊 Visual Insights")

            colA, colB = st.columns(2)

            # --- Chart 1: Solution Name vs Server Count ---
            if solution_col in matched_servers.columns:
                with colA:
                    soln_chart = (
                        matched_servers.groupby(solution_col)[results_key_col]
                        .nunique()
                        .reset_index()
                        .rename(columns={results_key_col: "Server Count"})
                    )

                    fig1 = px.bar(
                        soln_chart,
                        x="Server Count",
                        y=solution_col,
                        orientation='h',
                        title="Server Count per Solution Name",
                        text_auto=True
                    )
                    fig1.update_layout(title_font_size=16)
                    st.plotly_chart(fig1, use_container_width=True)

            # --- Chart 2: Change Number vs Server Count ---
            with colB:
                change_chart = (
                    matched_servers.groupby("UpdatedValue")[results_key_col]
                    .nunique()
                    .reset_index()
                    .rename(columns={results_key_col: "Server Count"})
                )

                fig2 = px.bar(
                    change_chart,
                    x="Server Count",
                    y="UpdatedValue",
                    orientation='h',
                    title="Server Count per Change Number",
                    text_auto=True
                )
                fig2.update_layout(title_font_size=16)
                st.plotly_chart(fig2, use_container_width=True)

            # =========================================================
            # 📘 Export to Excel (with Matched Servers)
            # =========================================================
            st.markdown("---")
            st.subheader("📤 Export Matched Results")

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Write all relevant sheets
                sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
                updated_results.to_excel(writer, sheet_name='Results', index=False)
                matched_servers_sheet.to_excel(writer, sheet_name='Matched Servers', index=False)

                # Highlight non-FQDN servers in Results
                workbook = writer.book
                worksheet = writer.sheets['Results']
                yellow_format = workbook.add_format({'bg_color': '#FFF59D'})  # Soft yellow

                for row_num, server in enumerate(updated_results[results_key_col], start=1):
                    if '.' not in str(server):
                        col_idx = updated_results.columns.get_loc(results_key_col)
                        worksheet.write(row_num, col_idx, server, yellow_format)

            output.seek(0)
            st.download_button(
                label="⬇️ Download Excel with Matched Servers",
                data=output,
                file_name="matched_highlighted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.success("✅ Processing complete! You can download the updated Excel file above.")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")

else:
    st.info("📥 Please upload an Excel file to start the analysis.")
