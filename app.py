import os
import streamlit as st
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

st.set_page_config(page_title="Autenticación SharePoint Lists", layout="wide")

# ─────────────────────────────────────────────────────────────
# 1. Configuration
# ─────────────────────────────────────────────────────────────
try:
    SITE_URL = os.getenv("SP_SITE_URL") or st.secrets["SP_SITE_URL"]
    USERNAME = os.getenv("SP_USERNAME") or st.secrets["SP_USERNAME"]
    PASSWORD = os.getenv("SP_PASSWORD") or st.secrets["SP_PASSWORD"]
except KeyError as e:
    st.error(f"🔒 Falta configuración: {e}")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 2. SharePoint Lists Functions
# ─────────────────────────────────────────────────────────────
def get_sharepoint_context():
    """Get authenticated SharePoint context"""
    user_credentials = UserCredential(USERNAME, PASSWORD)
    ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
    return ctx

@st.cache_data(ttl=300)
def load_credentials_from_list():
    """Load credentials from ProveedorCredenciales SharePoint List"""
    try:
        ctx = get_sharepoint_context()
        
        # Get the ProveedorCredenciales list
        credentials_list = ctx.web.lists.get_by_title("ProveedorCredenciales")
        credentials_items = credentials_list.items
        ctx.load(credentials_items)
        ctx.execute_query()
        
        st.write(f"**Total items found:** {len(credentials_items)}")
        
        # Debug: Show what properties are available
        if len(credentials_items) > 0:
            first_item = credentials_items[0]
            st.write("**Available properties in first item:**")
            
            # Try different ways to access properties
            properties = first_item.properties
            st.write("All properties:", properties)
            
            # Try specific property access methods
            st.write("**Testing different access methods:**")
            
            # Method 1: Direct property access
            try:
                usuario1 = first_item.properties.get('usuario')
                password1 = first_item.properties.get('password')
                st.write(f"Method 1 - usuario: {usuario1}, password: {password1}")
            except Exception as e:
                st.write(f"Method 1 failed: {e}")
            
            # Method 2: Title and other common fields
            try:
                title = first_item.properties.get('Title')
                st.write(f"Title field: {title}")
            except Exception as e:
                st.write(f"Title access failed: {e}")
            
            # Method 3: Try with different casing
            try:
                usuario2 = first_item.properties.get('Usuario')
                password2 = first_item.properties.get('Password')
                st.write(f"Method 3 - Usuario: {usuario2}, Password: {password2}")
            except Exception as e:
                st.write(f"Method 3 failed: {e}")
        
        # Convert to list of dictionaries with multiple attempts
        credentials_data = []
        for item in credentials_items:
            # Try different property names and access methods
            usuario = (item.properties.get('usuario') or 
                      item.properties.get('Usuario') or 
                      item.properties.get('Title'))
            
            password = (item.properties.get('password') or 
                       item.properties.get('Password'))
            
            credentials_data.append({
                'usuario': usuario,
                'password': password,
                'all_properties': dict(item.properties)  # For debugging
            })
        
        return credentials_data
        
    except Exception as e:
        st.error(f"Error al cargar credenciales: {str(e)}")
        return None

def authenticate_user(usuario, password):
    """Authenticate user against SharePoint List"""
    credentials_data = load_credentials_from_list()
    
    if credentials_data is None:
        return False, "Error al cargar credenciales"
    
    # Check if user exists
    for credential in credentials_data:
        if (credential['usuario'] == usuario and 
            credential['password'] == password):
            return True, "Autenticación exitosa"
    
    return False, "Credenciales incorrectas"

# ─────────────────────────────────────────────────────────────
# 3. Main App
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🔐 Autenticación con SharePoint Lists")
    
    # Session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    
    # Authentication form
    if not st.session_state.authenticated:
        st.subheader("Iniciar Sesión")
        
        # Test connection
        if st.button("🧪 Probar Conexión"):
            with st.spinner("Probando conexión..."):
                credentials = load_credentials_from_list()
                if credentials:
                    st.success(f"✅ Conexión exitosa. {len(credentials)} usuarios encontrados")
                    st.write("Usuarios disponibles:", [c['usuario'] for c in credentials])
                else:
                    st.error("❌ Error de conexión")
        
        st.markdown("---")
        
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