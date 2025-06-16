import io
import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

st.set_page_config(page_title="Sistema de Reserva de Entregas", layout="wide")

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
# 2. Excel Download Functions
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
        
        # Load both sheets
        credentials_df = pd.read_excel(file_content, sheet_name="proveedor_credencial")
        reservas_df = pd.read_excel(file_content, sheet_name="proveedor_reservas")
        
        return credentials_df, reservas_df
        
    except Exception as e:
        st.error(f"Error descargando Excel: {str(e)}")
        return None, None

def save_booking_to_excel(new_booking):
    """Save new booking to Excel file"""
    try:
        # Load current data
        credentials_df, reservas_df = download_excel_to_memory()
        
        if reservas_df is None:
            return False
        
        # Add new booking
        new_row = pd.DataFrame([new_booking])
        updated_reservas_df = pd.concat([reservas_df, new_row], ignore_index=True)
        
        # Authenticate and upload
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Create Excel file
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            credentials_df.to_excel(writer, sheet_name="proveedor_credencial", index=False)
            updated_reservas_df.to_excel(writer, sheet_name="proveedor_reservas", index=False)
        
        excel_buffer.seek(0)
        
        # Get the file and try to upload
        file = ctx.web.get_file_by_id(FILE_ID)
        
        # Simple upload method
        file.upload(excel_buffer.getvalue())
        ctx.execute_query()
        
        # Clear cache
        download_excel_to_memory.clear()
        
        return True
        
    except Exception as e:
        st.error(f"Error guardando reserva: {str(e)}")
        return False

# ─────────────────────────────────────────────────────────────
# 3. Time Slot Functions
# ─────────────────────────────────────────────────────────────
def generate_time_slots():
    """Generate available time slots"""
    # Monday-Friday: 9:00-16:00, Saturday: 9:00-12:00
    weekday_slots = []
    saturday_slots = []
    
    # Weekday slots (9:00-16:00)
    start_hour = 9
    end_hour = 16
    for hour in range(start_hour, end_hour):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            end_minute = minute + 30
            end_hour_calc = hour if end_minute < 60 else hour + 1
            end_minute = end_minute if end_minute < 60 else 0
            end_time = f"{end_hour_calc:02d}:{end_minute:02d}"
            weekday_slots.append(f"{start_time}-{end_time}")
    
    # Saturday slots (9:00-12:00)
    for hour in range(9, 12):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            end_minute = minute + 30
            end_hour_calc = hour if end_minute < 60 else hour + 1
            end_minute = end_minute if end_minute < 60 else 0
            end_time = f"{end_hour_calc:02d}:{end_minute:02d}"
            saturday_slots.append(f"{start_time}-{end_time}")
    
    return weekday_slots, saturday_slots

def get_available_slots(selected_date, reservas_df):
    """Get available slots for a date"""
    weekday_slots, saturday_slots = generate_time_slots()
    
    # Sunday = 6, no work
    if selected_date.weekday() == 6:
        return []
    
    # Saturday = 5
    if selected_date.weekday() == 5:
        all_slots = saturday_slots
    else:
        all_slots = weekday_slots
    
    # Filter booked slots
    date_str = selected_date.strftime('%Y-%m-%d')
    booked_slots = reservas_df[reservas_df['Fecha'] == date_str]['Hora'].tolist()
    
    return [slot for slot in all_slots if slot not in booked_slots]

# ─────────────────────────────────────────────────────────────
# 4. Authentication Function
# ─────────────────────────────────────────────────────────────
def authenticate_user(usuario, password):
    """Authenticate user against Excel data"""
    credentials_df, _ = download_excel_to_memory()
    
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
    st.title("🚚 Sistema de Reserva de Entregas")
    
    # Download Excel when app starts
    with st.spinner("Cargando datos..."):
        credentials_df, reservas_df = download_excel_to_memory()
    
    if credentials_df is None:
        st.error("❌ Error al cargar archivo")
        return
    
    st.success(f"✅ Datos cargados: {len(credentials_df)} usuarios, {len(reservas_df)} reservas")
    
    # Session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    
    # Authentication
    if not st.session_state.authenticated:
        st.subheader("🔐 Iniciar Sesión")
        
        with st.form("login_form"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar Sesión")
            
            if submitted:
                if usuario and password:
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
    
    # Booking interface
    else:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader(f"Bienvenido, {st.session_state.supplier_name}")
        with col2:
            if st.button("Cerrar Sesión"):
                st.session_state.authenticated = False
                st.session_state.supplier_name = None
                st.rerun()
        
        st.markdown("---")
        
        # Date selection
        st.subheader("📅 Seleccionar Fecha")
        today = datetime.now().date()
        max_date = today + timedelta(days=30)
        
        selected_date = st.date_input(
            "Fecha de entrega",
            min_value=today,
            max_value=max_date,
            value=today
        )
        
        # Check if Sunday
        if selected_date.weekday() == 6:
            st.warning("⚠️ No trabajamos los domingos")
            return
        
        # Time slot selection
        st.subheader("🕐 Horarios Disponibles")
        
        available_slots = get_available_slots(selected_date, reservas_df)
        
        if not available_slots:
            st.warning("❌ No hay horarios disponibles para esta fecha")
            return
        
        # Display slots in columns
        cols = st.columns(3)
        selected_slot = None
        
        for i, slot in enumerate(available_slots):
            with cols[i % 3]:
                if st.button(slot, key=f"slot_{i}"):
                    selected_slot = slot
        
        # Booking form
        if selected_slot or 'selected_slot' in st.session_state:
            if selected_slot:
                st.session_state.selected_slot = selected_slot
            
            st.markdown("---")
            st.subheader("📦 Información de Entrega")
            
            with st.form("booking_form"):
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"Fecha: {selected_date}")
                    st.info(f"Horario: {st.session_state.selected_slot}")
                
                with col2:
                    numero_bultos = st.number_input("Número de bultos", min_value=1, value=1)
                    orden_compra = st.text_input("Orden de compra", placeholder="Ej: OC-2024-001")
                
                submitted = st.form_submit_button("Confirmar Reserva")
                
                if submitted:
                    if orden_compra.strip():
                        new_booking = {
                            'Fecha': selected_date.strftime('%Y-%m-%d'),
                            'Hora': st.session_state.selected_slot,
                            'Proveedor': st.session_state.supplier_name,
                            'Numero_de_bultos': numero_bultos,
                            'Orden_de_compra': orden_compra.strip()
                        }
                        
                        with st.spinner("Guardando reserva..."):
                            success = save_booking_to_excel(new_booking)
                        
                        if success:
                            st.success("✅ Reserva confirmada!")
                            st.balloons()
                            del st.session_state.selected_slot
                            st.rerun()
                        else:
                            st.error("❌ Error al guardar reserva")
                    else:
                        st.warning("⚠️ Ingrese la orden de compra")

if __name__ == "__main__":
    main()