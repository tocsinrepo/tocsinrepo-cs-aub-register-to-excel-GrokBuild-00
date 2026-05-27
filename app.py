import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="Bank vs GL Reconciler", layout="wide")
st.title("🔍 Bank Register vs GL Missing Items Tool")

def clean_amount(x):
    if isinstance(x, str):
        x = x.replace('$', '').replace(',', '').strip()
    try:
        return float(x)
    except:
        return 0.0

# File Uploaders
col1, col2 = st.columns(2)
with col1:
    bank_file = st.file_uploader("Upload Bank Register Excel", type=["xlsx", "xls"])
with col2:
    gl_file = st.file_uploader("Upload GL Excel", type=["xlsx", "xls"])

if bank_file and gl_file:
    bank_df = pd.read_excel(bank_file)
    gl_df = pd.read_excel(gl_file)

    # Normalize columns
    bank_df.columns = [col.strip().lower().replace(' ', '_') for col in bank_df.columns]
    gl_df.columns = [col.strip().lower().replace(' ', '_') for col in gl_df.columns]

    st.subheader("Detected Columns")
    st.write("**Bank Columns:**", list(bank_df.columns))
    st.write("**GL Columns:**", list(gl_df.columns))

    # Column selection
    date_col = st.text_input("Date Column Name", value="date").strip().lower()
    amt_col = st.text_input("Amount Column Name", value="amount").strip().lower()
    desc_col = st.text_input("Description Column (optional)", value="").strip().lower() or None

    if st.button("Run Reconciliation", type="primary"):
        with st.spinner("Processing..."):
            # Clean data
            for df in [bank_df, gl_df]:
                if date_col in df.columns:
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                if amt_col in df.columns:
                    df[amt_col] = df[amt_col].apply(clean_amount)

            def create_key(df):
                key = df[date_col].dt.strftime('%Y-%m-%d') + '|' + df[amt_col].round(2).astype(str)
                if desc_col and desc_col in df.columns:
                    key = key + '|' + df[desc_col].astype(str).str.lower().str.strip()
                return key

            bank_df['key'] = create_key(bank_df)
            gl_df['key'] = create_key(gl_df)

            bank_only = bank_df[~bank_df['key'].isin(gl_df['key'])].copy()
            gl_only = gl_df[~gl_df['key'].isin(bank_df['key'])].copy()

            # Add status
            bank_df['Status'] = bank_df['key'].isin(gl_df['key']).map({True: 'Matched', False: 'MISSING_IN_GL'})
            gl_df['Status'] = gl_df['key'].isin(bank_df['key']).map({True: 'Matched', False: 'MISSING_IN_BANK'})

            # Create output
            output_dir = Path("output")
            output_dir.mkdir(exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")

            # Generate files
            highlighted_path = output_dir / f"comparison_highlighted_{timestamp}.xlsx"
            missing_path = output_dir / f"missing_items_only_{timestamp}.xlsx"

            with pd.ExcelWriter(highlighted_path, engine='openpyxl') as writer:
                bank_df.to_excel(writer, sheet_name='Bank_Register', index=False)
                gl_df.to_excel(writer, sheet_name='GL', index=False)

            with pd.ExcelWriter(missing_path, engine='openpyxl') as writer:
                if not bank_only.empty:
                    bank_only.drop(columns=['key'], errors='ignore').to_excel(writer, sheet_name='Missing_in_GL', index=False)
                if not gl_only.empty:
                    gl_only.drop(columns=['key'], errors='ignore').to_excel(writer, sheet_name='Missing_in_Bank', index=False)

            st.success("✅ Reconciliation Complete!")

            # Download buttons
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button("📥 Download Highlighted Comparison", 
                                 highlighted_path.read_bytes(), 
                                 f"comparison_highlighted_{timestamp}.xlsx",
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_dl2:
                st.download_button("📥 Download Missing Items Only", 
                                 missing_path.read_bytes(), 
                                 f"missing_items_only_{timestamp}.xlsx",
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.info(f"**Missing in GL:** {len(bank_only)} items | **Missing in Bank:** {len(gl_only)} items")
