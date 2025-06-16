import io
import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import calendar

st.set_page_config(page_title="Reserva de Entregas - Proveedores", layout="wide")

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
def load_excel_from_sharepoint():
    """Load Excel file from SharePoint and return both sheets as DataFrames"""
    try:
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Get file
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        # Download file content
        response = file.download()
        ctx.execute_query()
        
        # Read Excel file
        excel_data = io.BytesIO(response.content)
        
        # Load both sheets
        credentials_df = pd.read_excel(excel_data, sheet_name="proveedor_credencial")
        reservas_df = pd.read_excel(excel_data, sheet_name="proveedor_reservas")
        
        return credentials_df, reservas_df
    except Exception as e:
        st.error(f"Error al cargar datos de SharePoint: {str(e)}")
        return None, None

def save_booking_to_sharepoint(new_booking):
    """Save new booking to SharePoint Excel file"""
    try:
        # Load current data
        credentials_df, reservas_df = load_excel_from_sharepoint()
        
        if reservas_df is None:
            return False
        
        # Add new booking
        new_row = pd.DataFrame([new_booking])
        updated_reservas_df = pd.concat([reservas_df, new_row], ignore_index=True)
        
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Create Excel file in memory
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            credentials_df.to_excel(writer, sheet_name="proveedor_credencial", index=False)
            updated_reservas_df.to_excel(writer, sheet_name="proveedor_reservas", index=False)
        
        excel_buffer.seek(0)
        
        # Upload file back to SharePoint
        file = ctx.web.get_file_by_id(FILE_ID)
        file.save_binary(excel_buffer.getvalue())
        ctx.execute_query()
        
        # Clear cache to refresh data
        load_excel_from_sharepoint.clear()
        
        return True
    except Exception as e:
        st.error(f"Error al guardar reserva: {str(e)}")
        return False

# ─────────────────────────────────────────────────────────────
# 3. Authentication Functions
# ─────────────────────────────────────────────────────────────
def authenticate_user(usuario, password):
    """Authenticate user against SharePoint Excel data"""
    credentials_df, _ = load_excel_from_sharepoint()
    
    if credentials_df is None:
        return False, None
    
    # Check credentials
    user_match = credentials_df[
        (credentials_df['usuario'] == usuario) & 
        (credentials_df['password'] == password)
    ]
    
    if not user_match.empty:
        return True, usuario
    
    return False, None

# ─────────────────────────────────────────────────────────────
# 4. Time Slot Functions
# ─────────────────────────────────────────────────────────────
def generate_time_slots():
    """Generate available time slots based on business hours"""
    slots = []
    
    # Monday to Friday: 9:00 AM to 4:00 PM (last slot 3:30-4:00)
    weekday_start = time(9, 0)
    weekday_end = time(16, 0)
    
    # Saturday: 9:00 AM to 12:00 PM (last slot 11:30-12:00)
    saturday_start = time(9, 0)
    saturday_end = time(12, 0)
    
    # Generate 30-minute slots
    current_time = datetime.combine(datetime.today(), weekday_start)
    end_time = datetime.combine(datetime.today(), weekday_end)
    
    weekday_slots = []
    while current_time.time() < weekday_end:
        next_time = current_time + timedelta(minutes=30)
        slot_str = f"{current_time.strftime('%H:%M')}-{next_time.strftime('%H:%M')}"
        weekday_slots.append(slot_str)
        current_time = next_time
    
    # Generate Saturday slots
    current_time = datetime.combine(datetime.today(), saturday_start)
    end_time = datetime.combine(datetime.today(), saturday_end)
    
    saturday_slots = []
    while current_time.time() < saturday_end:
        next_time = current_time + timedelta(minutes=30)
        slot_str = f"{current_time.strftime('%H:%M')}-{next_time.strftime('%H:%M')}"
        saturday_slots.append(slot_str)
        current_time = next_time
    
    return weekday_slots, saturday_slots

def get_available_slots(date_selected, reservas_df):
    """Get available time slots for a specific date"""
    weekday_slots, saturday_slots = generate_time_slots()
    
    # Determine which slots to use based on day of week
    if date_selected.weekday() < 5:  # Monday to Friday (0-4)
        all_slots = weekday_slots
    elif date_selected.weekday() == 5:  # Saturday (5)
        all_slots = saturday_slots
    else:  # Sunday (6)
        return []  # No work on Sundays
    
    # Filter out already booked slots
    if reservas_df is not None and not reservas_df.empty:
        booked_slots = reservas_df[
            reservas_df['Fecha'] == date_selected.strftime('%Y-%m-%d')
        ]['Hora'].tolist()
        
        available_slots = [slot for slot in all_slots if slot not in booked_slots]
    else:
        available_slots = all_slots
    
    return available_slots

# ─────────────────────────────────────────────────────────────
# 5. Main Application
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🚚 Sistema de Reserva de Entregas")
    st.markdown("---")
    
    # Initialize session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    
    # Authentication Section
    if not st.session_state.authenticated:
        st.subheader("🔐 Acceso para Proveedores")
        
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
                        is_valid, supplier_name = authenticate_user(usuario, password)
                    
                    if is_valid:
                        st.session_state.authenticated = True
                        st.session_state.supplier_name = supplier_name
                        st.success("✅ Acceso autorizado")
                        st.rerun()
                    else:
                        st.error("❌ Credenciales incorrectas")
                else:
                    st.warning("⚠️ Por favor complete todos los campos")
    
    # Main Booking Interface
    else:
        # Header with logout option
        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader(f"Bienvenido, {st.session_state.supplier_name}")
        with col2:
            if st.button("Cerrar Sesión"):
                st.session_state.authenticated = False
                st.session_state.supplier_name = None
                st.rerun()
        
        st.markdown("---")
        
        # Load current reservations
        with st.spinner("Cargando disponibilidad..."):
            _, reservas_df = load_excel_from_sharepoint()
        
        if reservas_df is None:
            st.error("No se pudo cargar la información de reservas")
            return
        
        # Date Selection
        st.subheader("📅 Seleccionar Fecha de Entrega")
        
        # Show next 30 days, excluding Sundays
        today = datetime.now().date()
        max_date = today + timedelta(days=30)
        
        date_selected = st.date_input(
            "Fecha de entrega",
            min_value=today,
            max_value=max_date,
            value=today
        )
        
        # Check if selected date is Sunday
        if date_selected.weekday() == 6:  # Sunday
            st.warning("⚠️ No trabajamos los domingos. Por favor seleccione otro día.")
            return
        
        # Time Slot Selection
        st.subheader("🕐 Seleccionar Horario")
        
        available_slots = get_available_slots(date_selected, reservas_df)
        
        if not available_slots:
            st.warning("❌ No hay horarios disponibles para esta fecha.")
            return
        
        # Display available slots in a nice format
        st.write(f"**Horarios disponibles para {date_selected.strftime('%A, %d de %B %Y')}:**")
        
        # Create columns for better layout
        cols = st.columns(3)
        selected_slot = None
        
        for i, slot in enumerate(available_slots):
            col_idx = i % 3
            with cols[col_idx]:
                if st.button(slot, key=f"slot_{i}", use_container_width=True):
                    selected_slot = slot
        
        # Booking Form
        if selected_slot or 'selected_slot' in st.session_state:
            if selected_slot:
                st.session_state.selected_slot = selected_slot
            
            st.markdown("---")
            st.subheader("📦 Información de la Entrega")
            
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"**Fecha:** {date_selected.strftime('%d/%m/%Y')}")
                st.info(f"**Horario:** {st.session_state.selected_slot}")
            
            with st.form("booking_form"):
                col1, col2 = st.columns(2)
                with col1:
                    numero_bultos = st.number_input(
                        "Número de bultos/paquetes",
                        min_value=1,
                        max_value=1000,
                        value=1,
                        help="Cantidad de bultos o paquetes a entregar"
                    )
                with col2:
                    orden_compra = st.text_input(
                        "Orden de compra",
                        placeholder="Ej: OC-2024-001",
                        help="Número de orden de compra asociada"
                    )
                
                submitted = st.form_submit_button("🎯 Confirmar Reserva", use_container_width=True)
                
                if submitted:
                    if orden_compra.strip():
                        new_booking = {
                            'Fecha': date_selected.strftime('%Y-%m-%d'),
                            'Hora': st.session_state.selected_slot,
                            'Proveedor': st.session_state.supplier_name,
                            'Numero_de_bultos': numero_bultos,
                            'Orden_de_compra': orden_compra.strip()
                        }
                        
                        with st.spinner("Guardando reserva..."):
                            success = save_booking_to_sharepoint(new_booking)
                        
                        if success:
                            st.success("✅ ¡Reserva confirmada exitosamente!")
                            st.balloons()
                            # Clear selected slot
                            if 'selected_slot' in st.session_state:
                                del st.session_state.selected_slot
                            st.rerun()
                        else:
                            st.error("❌ Error al confirmar la reserva. Intente nuevamente.")
                    else:
                        st.warning("⚠️ Por favor ingrese el número de orden de compra")

        # Show existing bookings for the supplier
        if not reservas_df.empty:
            supplier_bookings = reservas_df[
                reservas_df['Proveedor'] == st.session_state.supplier_name
            ].copy()
            
            if not supplier_bookings.empty:
                st.markdown("---")
                st.subheader("📋 Sus Reservas Actuales")
                
                # Sort by date and time
                supplier_bookings['Fecha'] = pd.to_datetime(supplier_bookings['Fecha'])
                supplier_bookings = supplier_bookings.sort_values(['Fecha', 'Hora'])
                supplier_bookings['Fecha'] = supplier_bookings['Fecha'].dt.strftime('%d/%m/%Y')
                
                # Display in a nice table
                st.dataframe(
                    supplier_bookings[['Fecha', 'Hora', 'Numero_de_bultos', 'Orden_de_compra']],
                    use_container_width=True,
                    hide_index=True
                )

if __name__ == "__main__":
    main()