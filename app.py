import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# --- Page Config ---
st.set_page_config(page_title="Automation Excel Comparator", layout="wide")

# --- Custom CSS (Automation Look) ---
st.markdown("""
    <style>
    body {
        background: linear-gradient(120deg, #0f2027, #203a43, #2c5364);
        color: #fff !important;
    }
    .stApp {
        background: radial-gradient(circle at top left, #203a43, #0f2027);
    }
    h1, h2, h3, h4 {
        color: #00c3ff !important;
        text-shadow: 0px 0px 8px rgba(0,195,255,0.5);
    }
    .block-container {
        padding-top: 2rem;
    }
    div[data-testid="stDataFrame"] {
        border: 2px solid #00c3ff;
        border-radius: 10px;
        box-shadow: 0px 0px 20px rgba(0,195,255,0.2);
    }
    </style>
""", unsafe_allow_html=True)

# --- App Title ---
st.title("🤖 Excel Automation Comparator — Match & Visualize Server Updates")

# --- File Upload ---
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)

        # Validation
        if 'Sheet1' not in xls.sheet_names or 'Results' not in xls.sheet_names:
            st.error("❌ Excel must contain 'Sheet1' and 'Results' sheets.")
        else:
            # ✅ Read both sheets WITHOUT treating the first row as a header for Sheet1
            sheet1 = pd.read_excel(xls, sheet_name='Sheet1', header=None)
            results = pd.read_excel(xls, sheet_name='Results')

            # Assume first two columns are key and change number
            sheet1.columns = ['Server', 'ChangeNumber']  # manual columns
            sheet1 = sheet1.dropna(subset=['Server'])  # remove empty rows

            results_key_col = results.columns[0]
            solution_col = "Solution Name" if "Solution Name" in results.columns else results.columns[1]

            # Normalize
            sheet1['normalized'] = sheet1['Server'].astype(str).str.strip().str.lower()
            results['normalized'] = results[results_key_col].astype(str).str.strip().str.lower()

            # Merge
            updated_results = results.merge(
                sheet1[['normalized', 'ChangeNumber']],
                on='normalized',
                how='left'
            ).rename(columns={'ChangeNumber': 'UpdatedValue'})

            updated_results['UpdatedValue'] = updated_results['UpdatedValue'].fillna("Not Found")
            matched_servers = updated_results[updated_results['UpdatedValue'] != "Not Found"]

            # --- Summary Metrics ---
            total_servers = len(results)
            matched_count = len(matched_servers)
            unmatched_count = total_servers - matched_count

            colA, colB, colC = st.columns(3)
            colA.metric("🖥️ Total Servers", total_servers)
            colB.metric("✅ Matched Servers", matched_count)
            colC.metric("❌ Unmatched Servers", unmatched_count)

            # --- Show Matched Data ---
            st.subheader("📊 Matched Servers with Updated Change Numbers")
            st.dataframe(matched_servers, use_container_width=True)

            # --- Visualization Section ---
            st.subheader("📈 Visualization Dashboard")

            col1, col2 = st.columns(2)

            # Horizontal bar: Server count per Solution Name
            if solution_col in matched_servers.columns:
                with col1:
                    soln_chart = matched_servers.groupby(solution_col)[results_key_col].nunique().reset_index()
                    soln_chart = soln_chart.rename(columns={results_key_col: "Server Count"})
                    fig1 = px.bar(
                        soln_chart,
                        x="Server Count",
                        y=solution_col,
                        orientation='h',
                        color="Server Count",
                        color_continuous_scale="tealgrn",
                        title="Solution-wise Server Distribution"
                    )
                    fig1.update_layout(
                        template="plotly_dark",
                        title_font_color="#00c3ff",
                        plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)"
                    )
                    st.plotly_chart(fig1, use_container_width=True)

            # Horizontal bar: Change Number vs Server Count
            with col2:
                change_chart = matched_servers.groupby("UpdatedValue")[results_key_col].nunique().reset_index()
                change_chart = change_chart.rename(columns={results_key_col: "Server Count"})
                fig2 = px.bar(
                    change_chart,
                    x="Server Count",
                    y="UpdatedValue",
                    orientation='h',
                    color="Server Count",
                    color_continuous_scale="bluered",
                    title="Change Number vs Server Count"
                )
                fig2.update_layout(
                    template="plotly_dark",
                    title_font_color="#00c3ff",
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)"
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

                # Highlight non-FQDN servers
                for row_num, server in enumerate(updated_results[results_key_col], start=1):
                    if '.' not in str(server):
                        col_idx = updated_results.columns.get_loc(results_key_col)
                        worksheet.write(row_num, col_idx, server, yellow_format)

            output.seek(0)
            st.download_button(
                label="⬇️ Download Updated Excel with Highlights",
                data=output,
                file_name="matched_highlighted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
