import io
import os
import streamlit as st
import pandas as pd
import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import json

st.set_page_config(page_title="Sistema de Autenticación - SharePoint Lists", layout="wide")

# ─────────────────────────────────────────────────────────────
# 1. Secrets / env-vars
# ─────────────────────────────────────────────────────────────
try:
    SITE_URL = os.getenv("SP_SITE_URL") or st.secrets["SP_SITE_URL"]
    USERNAME = os.getenv("SP_USERNAME") or st.secrets["SP_USERNAME"]
    PASSWORD = os.getenv("SP_PASSWORD") or st.secrets["SP_PASSWORD"]
    
    # Extract base URL for REST API
    # Example: https://yourtenant.sharepoint.com/sites/yoursite
    BASE_API_URL = SITE_URL.rstrip('/') + "/_api/web/lists"
    
except KeyError as e:
    st.error(f"🔒 Falta configuración requerida: {e}")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 2. SharePoint Lists Functions (Much Faster!)
# ─────────────────────────────────────────────────────────────
def get_sharepoint_context():
    """Get authenticated SharePoint context"""
    user_credentials = UserCredential(USERNAME, PASSWORD)
    ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
    return ctx

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_credentials_from_lists():
    """Load credentials from SharePoint List (much faster than Excel)"""
    try:
        st.info("🔄 Conectando a SharePoint Lists...")
        
        ctx = get_sharepoint_context()
        
        # Get ProveedorCredenciales list
        credentials_list = ctx.web.lists.get_by_title("ProveedorCredenciales")
        credentials_items = credentials_list.items
        ctx.load(credentials_items)
        ctx.execute_query()
        
        st.info(f"✅ Lista de credenciales cargada: {len(credentials_items)} usuarios")
        
        # Convert to DataFrame
        credentials_data = []
        for item in credentials_items:
            credentials_data.append({
                'ID': item.id,
                'usuario': item.get_property('usuario'),
                'password': item.get_property('password')
            })
        
        credentials_df = pd.DataFrame(credentials_data)
        
        st.success(f"✅ Credenciales procesadas: {len(credentials_df)} usuarios encontrados")
        
        return credentials_df
        
    except Exception as e:
        st.error(f"❌ Error al cargar credenciales de SharePoint Lists: {str(e)}")
        st.info("💡 Verifique que:")
        st.info("   • La lista 'ProveedorCredenciales' exista")
        st.info("   • Las columnas 'usuario' y 'password' estén creadas")
        st.info("   • USERNAME y PASSWORD tengan permisos")
        return None

@st.cache_data(ttl=60)  # Cache for 1 minute (shorter for booking data)
def load_reservations_from_lists():
    """Load reservations from SharePoint List"""
    try:
        st.info("🔄 Cargando reservas...")
        
        ctx = get_sharepoint_context()
        
        # Get ProveedorReservas list
        reservas_list = ctx.web.lists.get_by_title("ProveedorReservas")
        reservas_items = reservas_list.items
        ctx.load(reservas_items)
        ctx.execute_query()
        
        st.info(f"✅ Reservas cargadas: {len(reservas_items)} registros")
        
        # Convert to DataFrame
        reservas_data = []
        for item in reservas_items:
            reservas_data.append({
                'ID': item.id,
                'Fecha': item.get_property('Fecha'),
                'Hora': item.get_property('Hora'),
                'Proveedor': item.get_property('Proveedor'),
                'Numero_de_bultos': item.get_property('Numero_de_bultos'),
                'Orden_de_compra': item.get_property('Orden_de_compra')
            })
        
        reservas_df = pd.DataFrame(reservas_data)
        
        return reservas_df
        
    except Exception as e:
        st.warning(f"⚠️ Error al cargar reservas: {str(e)}")
        st.info("💡 La lista 'ProveedorReservas' puede estar vacía o no existir aún")
        return pd.DataFrame()

def save_booking_to_lists(new_booking):
    """Save new booking to SharePoint List (much faster than Excel)"""
    try:
        ctx = get_sharepoint_context()
        
        # Get ProveedorReservas list
        reservas_list = ctx.web.lists.get_by_title("ProveedorReservas")
        
        # Create new item
        item_properties = {
            'Fecha': new_booking['Fecha'],
            'Hora': new_booking['Hora'],
            'Proveedor': new_booking['Proveedor'],
            'Numero_de_bultos': new_booking['Numero_de_bultos'],
            'Orden_de_compra': new_booking['Orden_de_compra']
        }
        
        new_item = reservas_list.add_item(item_properties)
        ctx.execute_query()
        
        # Clear cache to refresh data
        load_reservations_from_lists.clear()
        
        st.success("✅ Reserva guardada exitosamente en SharePoint List")
        return True
        
    except Exception as e:
        st.error(f"❌ Error al guardar reserva: {str(e)}")
        return False

# ─────────────────────────────────────────────────────────────
# 3. Authentication Functions
# ─────────────────────────────────────────────────────────────
def authenticate_user(usuario, password):
    """Authenticate user against SharePoint List data"""
    credentials_df = load_credentials_from_lists()
    
    if credentials_df is None or credentials_df.empty:
        return False, None, "No se pudieron cargar las credenciales"
    
    # Debug: Show what columns we have
    st.write("**Columnas encontradas en la lista:**", list(credentials_df.columns))
    
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
# 4. Test Functions
# ─────────────────────────────────────────────────────────────
def test_lists_setup():
    """Test if SharePoint Lists are properly configured"""
    try:
        ctx = get_sharepoint_context()
        
        # Test both lists exist
        st.info("🧪 Verificando listas de SharePoint...")
        
        # Check ProveedorCredenciales
        try:
            credentials_list = ctx.web.lists.get_by_title("ProveedorCredenciales")
            ctx.load(credentials_list)
            ctx.execute_query()
            st.success("✅ Lista 'ProveedorCredenciales' encontrada")
        except Exception as e:
            st.error(f"❌ Lista 'ProveedorCredenciales' no encontrada: {e}")
            return False
        
        # Check ProveedorReservas
        try:
            reservas_list = ctx.web.lists.get_by_title("ProveedorReservas")
            ctx.load(reservas_list)
            ctx.execute_query()
            st.success("✅ Lista 'ProveedorReservas' encontrada")
        except Exception as e:
            st.error(f"❌ Lista 'ProveedorReservas' no encontrada: {e}")
            return False
        
        return True
        
    except Exception as e:
        st.error(f"❌ Error general en la configuración: {e}")
        return False

# ─────────────────────────────────────────────────────────────
# 5. Main Application
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🚀 Sistema de Autenticación - SharePoint Lists (Rápido)")
    st.markdown("---")
    
    # Initialize session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    
    # Show current configuration
    with st.expander("🔧 Configuración Actual"):
        st.write(f"**Site URL:** {SITE_URL}")
        st.write(f"**API Base URL:** {BASE_API_URL}")
        st.write(f"**SharePoint User:** {USERNAME}")
        st.write("**Password:** [HIDDEN]")
    
    # Performance info
    st.info("⚡ **Nota:** Este sistema usa SharePoint Lists en lugar de Excel para mayor velocidad")
    
    # Authentication Section
    if not st.session_state.authenticated:
        st.subheader("🔐 Iniciar Sesión")
        
        # Test setup button
        if st.button("🧪 Verificar Configuración de Listas"):
            with st.spinner("Verificando listas..."):
                setup_ok = test_lists_setup()
                if setup_ok:
                    st.success("🎉 ¡Configuración de listas correcta!")
                else:
                    st.error("💥 Error en la configuración de listas")
                    st.info("📋 **Pasos para crear las listas:**")
                    st.code("""
1. Ir a SharePoint → New → List
2. Crear 'ProveedorCredenciales' con columnas:
   - usuario (Single line of text)
   - password (Single line of text)
   
3. Crear 'ProveedorReservas' con columnas:
   - Fecha (Date)
   - Hora (Single line of text)
   - Proveedor (Single line of text)
   - Numero_de_bultos (Number)
   - Orden_de_compra (Single line of text)
                    """)
        
        # Test connection button
        if st.button("🔄 Probar Conexión Rápida"):
            with st.spinner("Probando conexión a listas..."):
                credentials_df = load_credentials_from_lists()
                if credentials_df is not None:
                    st.success("🚀 ¡Conexión rápida exitosa!")
                    st.metric("Tiempo estimado", "~1 segundo", "vs ~5-10 seg con Excel")
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
                    
                    if is