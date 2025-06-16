import io
import os
import streamlit as st
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

st.set_page_config(page_title="Sistema de Autenticación", layout="wide")

# ─────────────────────────────────────────────────────────────
# 1. Secrets / env-vars
# ─────────────────────────────────────────────────────────────
try:
    SITE_URL = os.getenv("SP_SITE_URL") or st.secrets["SP_SITE_URL"]
    FILE_ID = os.getenv("SP_FILE_ID") or st.secrets["SP_FILE_ID"]
    FILE_NAME = os.getenv("SP_FILE_NAME") or st.secrets.get("SP_FILE_NAME", "")
    USERNAME = os.getenv("SP_USERNAME") or st.secrets["SP_USERNAME"]
    PASSWORD = os.getenv("SP_PASSWORD") or st.secrets["SP_PASSWORD"]
except KeyError as e:
    st.error(f"🔒 Falta configuración requerida: {e}")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 2. SharePoint Functions
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_credentials_from_sharepoint():
    """Load credentials sheet from SharePoint Excel file"""
    try:
        st.info("🔄 Conectando a SharePoint...")
        
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        st.info("✅ Autenticación a SharePoint exitosa")
        
        # Get file
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        st.info("📁 Archivo encontrado, descargando...")
        
        # Create a BytesIO object to store the file content
        file_content = io.BytesIO()
        
        # Try multiple download methods based on library version
        try:
            # Method 1: Newer API version
            file.download(file_content)
            ctx.execute_query()
            st.info("✅ Método de descarga 1 exitoso")
        except TypeError as e:
            st.warning(f"⚠️ Método 1 falló: {e}")
            try:
                # Method 2: Alternative for different versions
                response = file.download()
                ctx.execute_query()
                file_content = io.BytesIO(response.content)
                st.info("✅ Método de descarga 2 exitoso")
            except Exception as e2:
                st.error(f"❌ Método 2 falló: {e2}")
                try:
                    # Method 3: Using download_session
                    file.download_session(file_content)
                    ctx.execute_query()
                    st.info("✅ Método de descarga 3 exitoso")
                except Exception as e3:
                    st.error(f"❌ Método 3 falló: {e3}")
                    raise e3
        
        # Reset pointer to beginning
        file_content.seek(0)
        
        st.info("📊 Procesando archivo Excel...")
        
        # Load credentials sheet
        credentials_df = pd.read_excel(file_content, sheet_name="proveedor_credencial")
        
        st.success(f"✅ Credenciales cargadas: {len(credentials_df)} usuarios encontrados")
        
        return credentials_df
        
    except Exception as e:
        st.error(f"❌ Error al cargar credenciales de SharePoint: {str(e)}")
        st.info("💡 Verifique que:")
        st.info("   • FILE_ID sea correcto")
        st.info("   • SITE_URL sea válida")
        st.info("   • USERNAME y PASSWORD tengan permisos")
        st.info("   • El archivo Excel tenga la hoja 'proveedor_credencial'")
        return None

# ─────────────────────────────────────────────────────────────
# 3. Authentication Functions
# ─────────────────────────────────────────────────────────────
def authenticate_user(usuario, password):
    """Authenticate user against SharePoint Excel data"""
    credentials_df = load_credentials_from_sharepoint()
    
    if credentials_df is None:
        return False, None, "No se pudieron cargar las credenciales"
    
    # Debug: Show what columns we have
    st.write("**Columnas encontradas en el archivo:**", list(credentials_df.columns))
    
    # Check if required columns exist
    if 'usuario' not in credentials_df.columns or 'password' not in credentials_df.columns:
        return False, None, f"Columnas requeridas no encontradas. Columnas disponibles: {list(credentials_df.columns)}"
    
    # Show sample data (without passwords)
    st.write("**Usuarios disponibles:**")
    sample_df = credentials_df[['usuario']].copy()
    st.dataframe(sample_df)
    
    # Check credentials
    user_match = credentials_df[
        (credentials_df['usuario'].astype(str).str.strip() == str(usuario).strip()) & 
        (credentials_df['password'].astype(str).str.strip() == str(password).strip())
    ]
    
    if not user_match.empty:
        return True, usuario, "Autenticación exitosa"
    
    return False, None, "Credenciales incorrectas"

# ─────────────────────────────────────────────────────────────
# 4. Main Application
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🔐 Sistema de Autenticación - Proveedores")
    st.markdown("---")
    
    # Initialize session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    
    # Show current configuration (without sensitive data)
    with st.expander("🔧 Configuración Actual"):
        st.write(f"**Site URL:** {SITE_URL}")
        st.write(f"**File ID:** {FILE_ID}")
        st.write(f"**File Name:** {FILE_NAME}")
        st.write(f"**SharePoint User:** {USERNAME}")
        st.write("**Password:** [HIDDEN]")
    
    # Authentication Section
    if not st.session_state.authenticated:
        st.subheader("🔐 Iniciar Sesión")
        
        # Test connection button
        if st.button("🧪 Probar Conexión a SharePoint"):
            with st.spinner("Probando conexión..."):
                credentials_df = load_credentials_from_sharepoint()
                if credentials_df is not None:
                    st.success("🎉 ¡Conexión exitosa!")
                else:
                    st.error("💥 Error en la conexión")
        
        st.markdown("---")
        
        with st.form("login_form"):
            col1, col2 = st.columns(2)
            with col1:
                usuario = st.text_input("Usuario", placeholder="Ingrese su usuario")
            with col2:
                password = st.text_input("Contraseña", type="password", placeholder="Ingrese su contraseña")
            
            submitted = st.form_submit_button("Iniciar Sesión", use_container_width=True)
            
            if submitted:
                if usuario and password:
                    with st.spinner("Verificando credenciales..."):
                        is_valid, supplier_name, message = authenticate_user(usuario, password)
                    
                    if is_valid:
                        st.session_state.authenticated = True
                        st.session_state.supplier_name = supplier_name
                        st.success(f"✅ {message}")
                        st.balloons()
                        st.rerun()
                    else:
                        st.error(f"❌ {message}")
                else:
                    st.warning("⚠️ Por favor complete todos los campos")
    
    # Authenticated Section
    else:
        st.success(f"🎉 ¡Bienvenido, {st.session_state.supplier_name}!")
        
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("🚪 Cerrar Sesión"):
                st.session_state.authenticated = False
                st.session_state.supplier_name = None
                st.rerun()
        
        st.markdown("---")
        st.info("🚧 Aquí irá el sistema de reservas...")

if __name__ == "__main__":
    main()