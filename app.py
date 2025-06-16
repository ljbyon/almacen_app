import io, os
import streamlit as st
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

st.set_page_config(page_title="SharePoint Excel Viewer", layout="wide")

# ─────────────────────────────────────────────────────────────
# 1. Secrets / env-vars
# ─────────────────────────────────────────────────────────────
SITE_URL   = os.getenv("SP_SITE_URL",   st.secrets.get("SP_SITE_URL",   ""))
FILE_ID    = os.getenv("SP_FILE_ID",    st.secrets.get("SP_FILE_ID",    ""))
FILE_NAME  = os.getenv("SP_FILE_NAME",  st.secrets.get("SP_FILE_NAME",  ""))
SHEET_NAME = os.getenv("SP_SHEET_NAME", st.secrets.get("SP_SHEET_NAME", "proveedor_credencial"))
USERNAME   = os.getenv("SP_USERNAME",   st.secrets.get("SP_USERNAME",   ""))
PASSWORD   = os.getenv("SP_PASSWORD",   st.secrets.get("SP_PASSWORD",   ""))

required = {
    "SP_SITE_URL": SITE_URL,
    "SP_FILE_ID":  FILE_ID,
    "SP_USERNAME": USERNAME,
    "SP_PASSWORD": PASSWORD,
}

missing = [k for k, v in required.items() if not v]
if missing:
    st.error(f"🔒 The following secrets / env-vars are missing: {', '.join(missing)}")
    st.stop()
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
        ctx.web.get_file_by_id(FILE_ID).download(buf).execute_query()
    except Exception:                           # ② fall back to server-relative path
        rel = f"/personal/{USERNAME.split('@')[0]}/Documents/{FILE_NAME}"
        ctx.web.get_file_by_server_relative_url(rel).download(buf).execute_query()

    if buf.tell() == 0:
        raise RuntimeError("Downloaded file is empty – check FILE_ID / permissions.")

    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=SHEET_NAME, header=1, engine="openpyxl")
    return df

# ─────────────────────────────────────────────────────────────
# 3. UI
# ─────────────────────────────────────────────────────────────
st.title("📊 SharePoint workbook viewer")

if st.button("Load workbook"):
    try:
        df = fetch_sheet()
    except Exception as e:
        st.error(f"❌ {e}")
        st.stop()

    st.success(f"Loaded **{len(df):,}** rows from “{SHEET_NAME}”")
    print(df)

else:
    st.info("Click **Load workbook** to download and explore the data.")
