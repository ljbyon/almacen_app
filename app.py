import io, os
import streamlit as st
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

st.set_page_config(page_title="SharePoint Excel Viewer", layout="wide")

# ─────────────────────────────────────────────────────────────
# 1. Secrets / env-vars
# ─────────────────────────────────────────────────────────────
try:
    SITE_URL   = os.getenv("SP_SITE_URL") or st.secrets["SP_SITE_URL"]
    FILE_ID    = os.getenv("SP_FILE_ID") or st.secrets["SP_FILE_ID"]
    FILE_NAME  = os.getenv("SP_FILE_NAME") or st.secrets.get("SP_FILE_NAME", "")
    SHEET_NAME = os.getenv("SP_SHEET_NAME") or st.secrets.get("SP_SHEET_NAME", "proveedor_credencial")
    USERNAME   = os.getenv("SP_USERNAME") or st.secrets["SP_USERNAME"]
    PASSWORD   = os.getenv("SP_PASSWORD") or st.secrets["SP_PASSWORD"]
except KeyError as e:
    st.error(f"🔒 Missing required secret: {e}")
    st.stop()

# Display configuration (without sensitive data)
with st.expander("📋 Configuration", expanded=False):
    st.write("**Site URL:**", SITE_URL)
    st.write("**File ID:**", FILE_ID)
    st.write("**File Name:**", FILE_NAME or "Not specified")
    st.write("**Sheet Name:**", SHEET_NAME)
    st.write("**Username:**", USERNAME)
    st.write("**Password:**", "●" * len(PASSWORD) if PASSWORD else "Not set")

# ─────────────────────────────────────────────────────────────
# 2. Cached download & parse
# ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Downloading & parsing workbook…")
def fetch_sheet():
    ctx = ClientContext(SITE_URL).with_credentials(
        UserCredential(USERNAME, PASSWORD)
    )

    buf = io.BytesIO()

    try:                                        # ① try GUID first
        st.info("🔍 Trying to download by File ID...")
        ctx.web.get_file_by_id(FILE_ID).download(buf).execute_query()
    except Exception as e:                      # ② fall back to server-relative path
        st.warning(f"File ID failed: {e}")
        st.info("🔍 Trying server-relative path...")
        rel = f"/personal/{USERNAME.split('@')[0].replace('.', '_')}/Documents/{FILE_NAME}"
        st.write(f"Trying path: `{rel}`")
        ctx.web.get_file_by_server_relative_url(rel).download(buf).execute_query()

    if buf.tell() == 0:
        raise RuntimeError("Downloaded file is empty – check FILE_ID / permissions.")

    st.success(f"✅ Downloaded {buf.tell():,} bytes")
    
    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=SHEET_NAME, header=1, engine="openpyxl")
    return df

# ─────────────────────────────────────────────────────────────
# 3. UI
# ─────────────────────────────────────────────────────────────
st.title("📊 SharePoint workbook viewer")

if st.button("Load workbook", type="primary"):
    try:
        df = fetch_sheet()
    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.stop()

    st.success(f"✅ Loaded **{len(df):,}** rows and **{len(df.columns)}** columns from sheet "{SHEET_NAME}"")
    
    # Display basic info about the dataframe
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Rows", len(df))
    with col2:
        st.metric("Columns", len(df.columns))
    with col3:
        st.metric("Size", f"{df.memory_usage(deep=True).sum() / 1024:.1f} KB")
    
    # Show column names
    st.write("**Column names:**")
    st.write(list(df.columns))
    
    # Display the dataframe
    st.write("**Data preview:**")
    st.dataframe(df, use_container_width=True)
    
    # Show data types and basic stats
    with st.expander("📊 Data Info", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Data Types:**")
            st.write(df.dtypes)
        with col2:
            st.write("**Basic Stats:**")
            st.write(df.describe())

else:
    st.info("👆 Click **Load workbook** to download and explore the data.")