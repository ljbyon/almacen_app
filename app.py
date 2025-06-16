import io
import os
import streamlit as st
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

st.set_page_config(page_title="Autenticación Excel", layout="wide")

# ─────────────────────────────────────────────────────────────
# 1. Configuration
# ─────────────────────────────────────────────────────────────
try:
    SITE_URL = os.getenv("SP_SITE_URL") or st.secrets["SP_SITE_URL"]
    FILE_ID = os.getenv("SP_FILE_ID") or st.secrets["SP_FILE_ID"]
    USERNAME = os.getenv("SP_USERNAME") or st.secrets["SP_USERNAME"]
    PASSWORD = os.getenv("SP_PASSWORD") or st.secrets["SP_PASSWORD"]
except KeyError as e:
    st.error(f"🔒 Falta configuración: {e}")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 2. Excel Download Function
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)  # Cache for 5 minutes
def download_excel_to_memory():
    """Download Excel file from SharePoint to memory"""
    try:
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Get file
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        # Download to memory
        file_content = io.BytesIO()
        
        # Try multiple download methods
        try:
            file.download(file_content)
            ctx.execute_query()
        except TypeError:
            try:
                response = file.download()
                ctx.execute_query()
                file_content = io.BytesIO(response.content)
            except:
                file.download_session(file_content)
                ctx.execute_query()
        
        file_content.seek(0)
        
        # Load credentials sheet
        credentials_df = pd.read_excel(file_content, sheet_name="proveedor_credencial")
        
        return credentials_df
        
    except Exception as e:
        st.error(f"Error descargando Excel: {str(e)}")
        return None

# ─────────────────────────────────────────────────────────────
# 3. Authentication Function
# ─────────────────────────────────────────────────────────────
def authenticate_user(usuario, password):
    """Authenticate user against Excel data"""
    credentials_df = download_excel_to_memory()
    
    if credentials_df is None:
        return False, "Error al cargar credenciales"
    
    # Check credentials
    user_match = credentials_df[
        (credentials_df['usuario'].astype(str).str.strip() == str(usuario).strip()) & 
        (credentials_df['password'].astype(str).str.strip() == str(password).strip())
    ]
    
    if not user_match.empty:
        return True, "Autenticación exitosa"
    
    return False, "Credenciales incorrectas"

# ─────────────────────────────────────────────────────────────
# 4. Main App
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🔐 Autenticación con Excel")
    
    # Download Excel when app starts
    with st.spinner("Descargando archivo de credenciales..."):
        credentials_df = download_excel_to_memory()
    
    if credentials_df is not None:
        st.success(f"✅ Archivo cargado: {len(credentials_df)} usuarios")
    else:
        st.error("❌ Error al cargar archivo")
        return
    
    # Session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    
    # Authentication form
    if not st.session_state.authenticated:
        st.subheader("Iniciar Sesión")
        
        # Show available users for debugging
        with st.expander("Debug: Usuarios disponibles"):
            st.dataframe(credentials_df[['usuario']])
        
        # Login form
        with st.form("login_form"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar Sesión")
            
            if submitted:
                if usuario and password:
                    with st.spinner("Verificando..."):
                        is_valid, message = authenticate_user(usuario, password)
                    
                    if is_valid:
                        st.session_state.authenticated = True
                        st.session_state.supplier_name = usuario
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                else:
                    st.warning("Complete todos los campos")
    
    # Authenticated view
    else:
        st.success(f"¡Bienvenido, {st.session_state.supplier_name}!")
        
        if st.button("Cerrar Sesión"):
            st.session_state.authenticated = False
            st.session_state.supplier_name = None
            st.rerun()
        
        st.info("Autenticación completada exitosamente")

if __name__ == "__main__":
    main()