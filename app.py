import streamlit as st
import pandas as pd
import io
import re
import plotly.express as px
from datetime import datetime

# --- PAGE SETUP ---
st.set_page_config(page_title="Pre-Post Audit Pro", page_icon="📡", layout="wide")

# CUSTOM CSS FOR PROFESSIONAL LOOK (Smaller fonts, responsive margins)
st.markdown("""
    <style>
    /* Reduce Header Sizes */
    h1 { font-size: 24px !important; color: #1c2b39; padding-bottom: 0px; }
    h2 { font-size: 18px !important; color: #2c3e50; }
    h3 { font-size: 16px !important; color: #34495e; }
    
    /* Global Font Size */
    html, body, [class*="st-"] { font-size: 14px !important; }
    
    /* Tighter Layout */
    .block-container { padding-top: 2rem !important; padding-bottom: 0rem !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 5px; border: 1px solid #e1e4e8; }
    
    /* Professional Blue Button */
    .stButton>button { 
        width: 100%; border-radius: 4px; height: 2.5em; 
        background-color: #0056b3; color: white; font-size: 14px; 
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1>📡 Precheck & Postcheck Auditor</h1>", unsafe_allow_html=True)
st.write("---")

# 1. UPLOADER SECTION
st.markdown("### 💡 Instructions")
st.caption("Upload 'Precheck_' files in the left box and 'Postcheck_' files in the right box. Logic will auto-align Column A-G side-by-side.")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 🟢 Step 1: Precheck Files")
    raw_pre = st.file_uploader("Select Precheck (.xlsx)", accept_multiple_files=True, key="pre", label_visibility="collapsed")
    pre_files = [f for f in raw_pre if f.name.lower().startswith('precheck')]
    if len(raw_pre) > len(pre_files):
        st.error(f"⚠️ {len(raw_pre)-len(pre_files)} invalid files removed.")
    if pre_files: st.success(f"Validated: {len(pre_files)} files.")

with col2:
    st.markdown("### 🔵 Step 2: Postcheck Files")
    raw_post = st.file_uploader("Select Postcheck (.xlsx)", accept_multiple_files=True, key="post", label_visibility="collapsed")
    post_files = [f for f in raw_post if f.name.lower().startswith('postcheck')]
    if len(raw_post) > len(post_files):
        st.error(f"⚠️ {len(raw_post)-len(post_files)} invalid files removed.")
    if post_files: st.success(f"Validated: {len(post_files)} files.")

# --- PROCESSING ---
if pre_files or post_files:
    if st.button("🚀 EXECUTE NETWORK AUDIT"):
        try:
            pattern = r"^(?:Precheck|Postcheck)_([^_]+)_"

            def process_files(uploaded_list):
                if not uploaded_list: return pd.DataFrame()
                frames = []
                for f in uploaded_list:
                    df = pd.read_excel(f, sheet_name='Summary')
                    df.columns = [str(c).strip() for c in df.columns]
                    match = re.search(pattern, f.name)
                    site_name = match.group(1) if match else "Unknown"
                    df.insert(0, 'Site_File', site_name)
                    df = df.rename(columns={df.columns[1]: 'Sector', df.columns[2]: 'Name'})
                    frames.append(df)
                return pd.concat(frames, ignore_index=True)

            df_pre = process_files(pre_files)
            df_post = process_files(post_files)

            # 2. JOIN & LOGIC
            final_df = pd.merge(df_pre, df_post, on=['Sector', 'Name'], how='outer', suffixes=('_Pre', '_Post'))

            status_cols = []
            if not df_pre.empty:
                status_cols = [c for c in df_pre.columns if c not in ['Site_File', 'Sector', 'Name']]
            elif not df_post.empty:
                status_cols = [c for c in df_post.columns if c not in ['Site_File', 'Sector', 'Name']]

            def verify(row):
                if not df_post.empty and pd.isna(row.get(f'{status_cols[0]}_Post')): return "MISSING POST"
                if not df_pre.empty and pd.isna(row.get(f'{status_cols[0]}_Pre')): return "MISSING PRE"
                diffs = [c for c in status_cols if str(row[f'{c}_Pre']).strip() != str(row[f'{c}_Post']).strip()]
                return "OK" if not diffs else f"MISMATCH: {', '.join(diffs)}"

            final_df['Audit_Status'] = final_df.apply(verify, axis=1)

            # COLUMN RECONSTRUCTION
            final_df = final_df.rename(columns={'Sector': 'Sector_U', 'Name': 'Name_U'})
            final_df['Sector_Pre'] = final_df['Sector_U'].where(final_df['Site_File_Pre'].notna())
            final_df['Name_Pre'] = final_df['Name_U'].where(final_df['Site_File_Pre'].notna())
            final_df['Sector_Post'] = final_df['Sector_U'].where(final_df['Site_File_Post'].notna())
            final_df['Name_Post'] = final_df['Name_U'].where(final_df['Site_File_Post'].notna())

            pre_cols = ['Site_File_Pre', 'Sector_Pre', 'Name_Pre'] + [f"{c}_Pre" for c in status_cols]
            post_cols = ['Site_File_Post', 'Sector_Post', 'Name_Post'] + [f"{c}_Post" for c in status_cols]
            final_df = final_df[pre_cols + post_cols + ['Audit_Status']]

            # --- DASHBOARD ---
            st.markdown("### 📊 Management Dashboard")
            counts = final_df['Audit_Status'].value_counts()
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Rows", len(final_df))
            m2.metric("Mismatches", counts.filter(like='MISMATCH').sum())
            m3.metric("Missing", counts.filter(like='MISSING').sum())
            
            fig = px.pie(names=counts.index, values=counts.values, hole=0.5, height=300)
            fig.update_layout(margin=dict(t=0, b=0, l=0, r=0))
            st.plotly_chart(fig, use_container_width=True)

            # --- EXCEL GENERATION ---
            output = io.BytesIO()
            df_action = final_df[final_df['Audit_Status'] != "OK"]
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for sheet_name, data in [('Full Audit', final_df), ('Action Required', df_action)]:
                    data.to_excel(writer, index=False, sheet_name=sheet_name)
                    ws, wb = writer.sheets[sheet_name], writer.book
                    fmt_data = wb.add_format({'border': 1, 'align': 'center', 'font_size': 10})
                    fmt_pre = wb.add_format({'bg_color': '#D9EAD3', 'bold': True, 'border': 1, 'align': 'center', 'font_size': 10})
                    fmt_post = wb.add_format({'bg_color': '#CFE2F3', 'bold': True, 'border': 1, 'align': 'center', 'font_size': 10})
                    red_cell = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'font_size': 10})

                    for i, col in enumerate(data.columns):
                        width = max(data[col].astype(str).map(len).max(), len(col)) + 2
                        ws.set_column(i, i, width, fmt_data)
                        if i < len(pre_cols): ws.write(0, i, col, fmt_pre)
                        elif len(pre_cols) <= i < (len(pre_cols) + len(post_cols)): ws.write(0, i, col, fmt_post)

                    for j, col in enumerate(status_cols):
                        pre_idx = data.columns.get_loc(f"{col}_Pre")
                        post_idx = data.columns.get_loc(f"{col}_Post")
                        ws.conditional_format(1, post_idx, len(data), post_idx, {
                            'type': 'formula',
                            'criteria': f'=AND(NOT(ISBLANK(${chr(65+post_idx)}2)), {chr(65+post_idx)}2<>{chr(65+pre_idx)}2)',
                            'format': red_cell
                        })

            st.download_button("📥 DOWNLOAD AUDIT REPORT", data=output.getvalue(), file_name="Network_Audit.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")