import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px

# ========== PAGE CONFIG ==========
st.set_page_config(page_title="Automation Dashboard", layout="wide")

# ========== BACKGROUND / LOGO ==========
def set_bg_image(image_url):
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url("{image_url}");
            background-attachment: fixed;
            background-size: cover;
            background-position: center;
            color: #FFFFFF;
        }}
        .stMetric {{
            background-color: rgba(0,0,0,0.5);
            border-radius: 10px;
            padding: 10px;
        }}
        .block-container {{
            background: rgba(0, 0, 0, 0.6);
            padding: 2rem;
            border-radius: 15px;
        }}
        h1, h2, h3, h4, h5, h6, p {{
            color: #FFFFFF !important;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# You can use any automation/tech-themed background image here:
set_bg_image("https://cdn.pixabay.com/photo/2017/06/14/03/00/gear-2402781_1280.jpg")

# ========== TITLE ==========
st.markdown("<h1 style='text-align:center; color:#00FFFF;'>🤖 Automation Server Update Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;'>Monitor, Match & Automate Server Updates</p>", unsafe_allow_html=True)

# ========== FILE UPLOAD ==========
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)

        # Validate sheets
        if 'Sheet1' not in xls.sheet_names or 'Results' not in xls.sheet_names:
            st.error("❌ Excel must contain both 'Sheet1' and 'Results' sheets.")
        else:
            sheet1 = pd.read_excel(xls, sheet_name='Sheet1')
            results = pd.read_excel(xls, sheet_name='Results')

            # Define key columns
            sheet1_key_col = sheet1.columns[0]
            sheet1_value_col = sheet1.columns[1]
            results_key_col = results.columns[0]
            solution_col = "Solution Name"  # Adjust if needed

            # Normalize
            sheet1['normalized'] = sheet1[sheet1_key_col].astype(str).str.strip().str.lower()
            results['normalized'] = results[results_key_col].astype(str).str.strip().str.lower()

            # Merge
            updated_results = results.merge(
                sheet1[['normalized', sheet1_value_col]],
                on='normalized',
                how='left'
            ).rename(columns={sheet1_value_col: "UpdatedValue"})

            updated_results['UpdatedValue'] = updated_results['UpdatedValue'].fillna("Not Found")
            matched_servers = updated_results[updated_results['UpdatedValue'] != "Not Found"]
            unmatched_servers = updated_results[updated_results['UpdatedValue'] == "Not Found"]

            # Summary Counts
            total_servers = results[results_key_col].nunique()
            matched_count = matched_servers[results_key_col].nunique()
            unmatched_count = unmatched_servers[results_key_col].nunique()

            st.markdown("### 📊 System Summary")

            c1, c2, c3 = st.columns(3)
            c1.metric("🖥️ Total Servers", total_servers)
            c2.metric("✅ Matched Servers", matched_count)
            c3.metric("❌ Unmatched Servers", unmatched_count)

            st.markdown("---")

            # Matched Table
            st.subheader("📋 Matched Servers with Change Numbers")
            st.dataframe(matched_servers, use_container_width=True)

            st.markdown("### 📈 Automation Insights")
            col1, col2 = st.columns(2)

            # Chart 1: Solution Name vs Server Count
            with col1:
                if solution_col in matched_servers.columns:
                    sol_count = matched_servers.groupby(solution_col)[results_key_col].nunique().sort_values(ascending=True)
                    fig1 = px.bar(
                        sol_count,
                        x=sol_count.values,
                        y=sol_count.index,
                        orientation='h',
                        title="Server Count per Solution Name",
                        text=sol_count.values,
                        color=sol_count.values,
                        color_continuous_scale='tealgrn'
                    )
                    fig1.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='white'),
                        title_font=dict(size=16, color='#00FFFF'),
                        margin=dict(l=20, r=20, t=50, b=20)
                    )
                    fig1.update_traces(textposition='outside')
                    st.plotly_chart(fig1, use_container_width=True)

            # Chart 2: Change Number vs Server Count
            with col2:
                chg_count = matched_servers.groupby("UpdatedValue")[results_key_col].nunique().sort_values(ascending=True)
                fig2 = px.bar(
                    chg_count,
                    x=chg_count.values,
                    y=chg_count.index,
                    orientation='h',
                    title="Server Count per Change Number",
                    text=chg_count.values,
                    color=chg_count.values,
                    color_continuous_scale='viridis'
                )
                fig2.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='white'),
                    title_font=dict(size=16, color='#00FFFF'),
                    margin=dict(l=20, r=20, t=50, b=20)
                )
                fig2.update_traces(textposition='outside')
                st.plotly_chart(fig2, use_container_width=True)

            st.markdown("---")

            # Unmatched Section
            with st.expander("🔍 View Unmatched Servers"):
                st.dataframe(unmatched_servers[[results_key_col, "UpdatedValue"]], use_container_width=True)

            # Excel Download
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
                label="📥 Download Excel with Highlights",
                data=output,
                file_name="matched_highlighted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
