import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="Bank vs GL Reconciler", layout="wide")
st.title("🔍 Bank Register vs GL Missing Items Tool")
st.markdown("Upload your files and match the columns correctly.")

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
    bank_file = st.file_uploader("📤 Upload Bank Register Excel", type=["xlsx", "xls"], key="bank")
with col2:
    gl_file = st.file_uploader("📤 Upload GL Excel", type=["xlsx", "xls"], key="gl")

if bank_file is None or gl_file is None:
    st.info("👆 Please upload both Excel files to begin.")
    st.stop()

# Load data
try:
    bank_df = pd.read_excel(bank_file)
    gl_df = pd.read_excel(gl_file)
except Exception as e:
    st.error(f"Error reading files: {e}")
    st.stop()

# Normalize column names
bank_df.columns = [str(col).strip().lower().replace(' ', '_') for col in bank_df.columns]
gl_df.columns = [str(col).strip().lower().replace(' ', '_') for col in gl_df.columns]

st.subheader("📋 Detected Columns")
col1, col2 = st.columns(2)
with col1:
    st.write("**Bank Register Columns:**")
    st.write(list(bank_df.columns))
with col2:
    st.write("**GL Columns:**")
    st.write(list(gl_df.columns))

# Column selection with dropdowns
st.subheader("🔑 Select Matching Columns")
date_col = st.selectbox("Date Column", options=list(bank_df.columns), index=0 if 'date' in bank_df.columns else 0, key="date_col")
amt_col = st.selectbox("Amount Column", options=list(bank_df.columns), index=0 if 'amount' in bank_df.columns else 0, key="amt_col")

desc_col = st.selectbox("Description Column (Optional)", options=["None"] + list(bank_df.columns), index=0, key="desc_col")
desc_col = None if desc_col == "None" else desc_col

if st.button("🚀 Run Reconciliation", type="primary", use_container_width=True):
    with st.spinner("Processing reconciliation..."):
        try:
            # Clean data
            for df in [bank_df, gl_df]:
                if date_col in df.columns:
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                if amt_col in df.columns:
                    df[amt_col] = df[amt_col].apply(clean_amount)

            # Create matching key
            def create_key(df):
                if date_col not in df.columns or amt_col not in df.columns:
                    raise KeyError(f"Column not found: date='{date_col}', amount='{amt_col}'")
                
                date_str = df[date_col].dt.strftime('%Y-%m-%d')
                amount_str = df[amt_col].round(2).astype(str)
                key = date_str + '|' + amount_str
                
                if desc_col and desc_col in df.columns:
                    desc_str = df[desc_col].astype(str).str.lower().str.strip()
                    key = key + '|' + desc_str
                return key

            bank_df['key'] = create_key(bank_df)
            gl_df['key'] = create_key(gl_df)

            # Find missing
            bank_only = bank_df[~bank_df['key'].isin(gl_df['key'])].copy()
            gl_only = gl_df[~gl_df['key'].isin(bank_df['key'])].copy()

            # Add status for highlighted view
            bank_df['Status'] = bank_df['key'].isin(gl_df['key']).map({True: 'Matched', False: 'MISSING_IN_GL'})
            gl_df['Status'] = gl_df['key'].isin(bank_df['key']).map({True: 'Matched', False: 'MISSING_IN_BANK'})

            # Output
            output_dir = Path("output")
            output_dir.mkdir(exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            highlighted_path = output_dir / f"comparison_highlighted_{timestamp}.xlsx"
            missing_path = output_dir / f"missing_items_only_{timestamp}.xlsx"

            with pd.ExcelWriter(highlighted_path, engine='openpyxl') as writer:
                bank_df.drop(columns=['key'], errors='ignore').to_excel(writer, sheet_name='Bank_Register', index=False)
                gl_df.drop(columns=['key'], errors='ignore').to_excel(writer, sheet_name='GL', index=False)

            with pd.ExcelWriter(missing_path, engine='openpyxl') as writer:
                if not bank_only.empty:
                    bank_only.drop(columns=['key'], errors='ignore').to_excel(writer, sheet_name='Missing_in_GL', index=False)
                if not gl_only.empty:
                    gl_only.drop(columns=['key'], errors='ignore').to_excel(writer, sheet_name='Missing_in_Bank', index=False)

            st.success("✅ Reconciliation completed successfully!")

            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Full Highlighted Report",
                    data=highlighted_path.read_bytes(),
                    file_name=f"comparison_highlighted_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.download_button(
                    label="📥 Download Missing Items Only",
                    data=missing_path.read_bytes(),
                    file_name=f"missing_items_only_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.info(f"**Results:** {len(bank_only)} items missing in GL | {len(gl_only)} items missing in Bank")

        except Exception as e:
            st.error(f"❌ Error during reconciliation: {str(e)}")
            st.info("Tip: Make sure the Date and Amount columns you selected actually exist in both files.")
