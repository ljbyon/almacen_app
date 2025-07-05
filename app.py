import io
import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

st.set_page_config(page_title="Dismac: Reserva de Entrega de Mercadería", layout="wide")

# ─────────────────────────────────────────────────────────────
# 1. Configuration
# ─────────────────────────────────────────────────────────────
try:
    SITE_URL = os.getenv("SP_SITE_URL") or st.secrets["SP_SITE_URL"]
    FILE_ID = os.getenv("SP_FILE_ID") or st.secrets["SP_FILE_ID"]
    USERNAME = os.getenv("SP_USERNAME") or st.secrets["SP_USERNAME"]
    PASSWORD = os.getenv("SP_PASSWORD") or st.secrets["SP_PASSWORD"]
    
    # Email configuration
    EMAIL_HOST = os.getenv("EMAIL_HOST") or st.secrets["EMAIL_HOST"]
    EMAIL_PORT = int(os.getenv("EMAIL_PORT") or st.secrets["EMAIL_PORT"])
    EMAIL_USER = os.getenv("EMAIL_USER") or st.secrets["EMAIL_USER"]
    EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD") or st.secrets["EMAIL_PASSWORD"]
    
except KeyError as e:
    st.error(f"🔒 Falta configuración: {e}")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 2. Excel Download Functions - UPDATED TO INCLUDE GESTION SHEET
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)  # Add show_spinner=False
def download_excel_to_memory():
    """Download Excel file from SharePoint to memory - INCLUDES ALL SHEETS"""
    try:
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Get file
        file = ctx.web.get_file_by_id(FILE_ID)
        if file is None:
            raise Exception("File object is None - FILE_ID may be incorrect")
            
        ctx.load(file)
        ctx.execute_query()
        
        # Download to memory
        file_content = io.BytesIO()
        
        # Try multiple download methods
        try:
            file.download(file_content)
            ctx.execute_query()
        except TypeError as e:
            try:
                response = file.download()
                if response is None:
                    raise Exception("Download response is None")
                ctx.execute_query()
                file_content = io.BytesIO(response.content)
            except Exception as e2:
                try:
                    file.download_session(file_content)
                    ctx.execute_query()
                except Exception as e3:
                    raise Exception(f"All download methods failed: {e}, {e2}, {e3}")
        
        file_content.seek(0)
        
        # Load all sheets - UPDATED
        credentials_df = pd.read_excel(file_content, sheet_name="proveedor_credencial", dtype=str)
        reservas_df = pd.read_excel(file_content, sheet_name="proveedor_reservas")
        
        # Try to load gestion sheet, create empty if doesn't exist - NEW
        try:
            gestion_df = pd.read_excel(file_content, sheet_name="proveedor_gestion")
        except ValueError:
            # Create empty gestion dataframe with required columns if sheet doesn't exist
            gestion_df = pd.DataFrame(columns=[
                'Orden_de_compra', 'Proveedor', 'Numero_de_bultos',
                'Hora_llegada', 'Hora_inicio_atencion', 'Hora_fin_atencion',
                'Tiempo_espera', 'Tiempo_atencion', 'Tiempo_total', 'Tiempo_retraso',
                'numero_de_semana', 'hora_de_reserva'
            ])
        
        return credentials_df, reservas_df, gestion_df
        
    except Exception as e:
        st.error(f"Error descargando Excel: {str(e)}")
        st.error(f"SITE_URL: {SITE_URL}")
        st.error(f"FILE_ID: {FILE_ID}")
        st.error(f"Error type: {type(e).__name__}")
        return None, None, None

def save_booking_to_excel(new_booking):
    """Save new booking to Excel file - PRESERVES ALL SHEETS - SINGLE ROW FOR 1-HOUR SLOTS"""
    try:
        # Load current data
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
        
        if reservas_df is None:
            st.error("❌ No se pudo cargar el archivo Excel")
            return False
        
        # Add new booking as single row
        new_row = pd.DataFrame([new_booking])
        reservas_df = pd.concat([reservas_df, new_row], ignore_index=True)
        
        # Authenticate and upload
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Create Excel file - SAVE ALL SHEETS
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            credentials_df.to_excel(writer, sheet_name="proveedor_credencial", index=False)
            reservas_df.to_excel(writer, sheet_name="proveedor_reservas", index=False)
            gestion_df.to_excel(writer, sheet_name="proveedor_gestion", index=False)
        
        # Get the file info
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        file_name = file.properties['Name']
        server_relative_url = file.properties['ServerRelativeUrl']
        folder_url = server_relative_url.replace('/' + file_name, '')
        
        # Upload the updated file
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        excel_buffer.seek(0)
        folder.files.add(file_name, excel_buffer.getvalue(), True)
        ctx.execute_query()
        
        # Clear cache only after successful save
        download_excel_to_memory.clear()
        
        return True
        
    except Exception as e:
        st.error(f"❌ Error guardando reserva: {str(e)}")
        return False

# ─────────────────────────────────────────────────────────────
# 3. Email Functions
# ─────────────────────────────────────────────────────────────
def download_pdf_attachment():
    """Download PDF attachment from SharePoint"""
    try:
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Target filename and exact path
        target_filename = "GUIA_DEL_SELLER_DISMAC_MARKETPLACE_Rev._1.pdf"
        file_path = f"/personal/ljbyon_dismac_com_bo/Documents/{target_filename}"
        
        try:
            # Try to get the file directly
            pdf_file = ctx.web.get_file_by_server_relative_url(file_path)
            ctx.load(pdf_file)
            ctx.execute_query()
            
        except Exception as e:
            # Fallback: List files in Documents folder
            try:
                folder = ctx.web.get_folder_by_server_relative_url("/personal/ljbyon_dismac_com_bo/Documents")
                files = folder.files
                ctx.load(files)
                ctx.execute_query()
                
                found_files = []
                pdf_file = None
                
                for file in files:
                    filename = file.name
                    found_files.append(filename)
                    
                    # Check if this is our target file
                    if filename == target_filename:
                        pdf_file = file
                        break
                
                # If still not found, try any PDF
                if pdf_file is None:
                    pdf_files = [f for f in found_files if f.lower().endswith('.pdf')]
                    
                    if pdf_files:
                        # Use the first PDF found
                        first_pdf = pdf_files[0]
                        pdf_file_path = f"/personal/ljbyon_dismac_com_bo/Documents/{first_pdf}"
                        pdf_file = ctx.web.get_file_by_server_relative_url(pdf_file_path)
                        ctx.load(pdf_file)
                        ctx.execute_query()
                    else:
                        raise Exception(f"No se encontró {target_filename} ni otros PDFs en Documents")
                        
            except Exception as e2:
                raise Exception(f"No se pudo acceder a Documents: {str(e2)}")
        
        if pdf_file is None:
            raise Exception("No se pudo cargar el archivo PDF")
        
        # Download PDF to memory
        pdf_content = io.BytesIO()
        
        try:
            pdf_file.download(pdf_content)
            ctx.execute_query()
        except TypeError:
            try:
                response = pdf_file.download()
                ctx.execute_query()
                pdf_content = io.BytesIO(response.content)
            except:
                pdf_file.download_session(pdf_content)
                ctx.execute_query()
        
        pdf_content.seek(0)
        pdf_data = pdf_content.getvalue()
        
        # Get filename
        try:
            filename = pdf_file.properties.get('Name', target_filename)
        except:
            filename = target_filename
        
        return pdf_data, filename
        
    except Exception as e:
        # Only show error if PDF download fails
        st.warning(f"No se pudo descargar el archivo adjunto: {str(e)}")
        return None, None

def send_booking_email(supplier_email, supplier_name, booking_details, cc_emails=None):
    """Send booking confirmation email with PDF attachment"""
    try:
        # Use provided CC emails or default
        if cc_emails is None or len(cc_emails) == 0:
            cc_emails = ["marketplace@dismac.com.bo", "ljbyon@dismac.com.bo"]
        else:
            # Add default email to the CC list if not already present
            if "marketplace@dismac.com.bo" not in cc_emails:
                cc_emails = cc_emails + ["marketplace@dismac.com.bo", "ljbyon@dismac.com.bo"]
        
        # Email content
        subject = "Confirmación de Reserva para Entrega de Mercadería"
        
        # Format dates for email display
        display_fecha = booking_details['Fecha'].split(' ')[0]  # Remove time part for display
        
        # Handle combined hora format for 1-hour reservations
        hora_field = booking_details['Hora']
        if ',' in hora_field:
            # Combined slots - show as range
            slots = [slot.strip() for slot in hora_field.split(',')]
            start_time = slots[0].rsplit(':', 1)[0]  # Remove seconds
            end_time_parts = slots[1].split(':')
            end_hour = int(end_time_parts[0])
            end_minute = int(end_time_parts[1])
            # Add 30 minutes to get actual end time
            if end_minute == 30:
                end_hour += 1
                end_minute = 0
            else:
                end_minute = 30
            end_time = f"{end_hour:02d}:{end_minute:02d}"
            display_hora = f"{start_time} - {end_time}"
            duration_info = " (Duración: 1 hora)"
        else:
            # Single slot
            display_hora = hora_field.rsplit(':', 1)[0]  # Remove seconds
            duration_info = ""
        
        body = f"""
        Hola {supplier_name},
        
        Su reserva de entrega ha sido confirmada exitosamente.
        
        DETALLES DE LA RESERVA:
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        📅 Fecha: {display_fecha}
        🕐 Horario: {display_hora}{duration_info}
        📦 Número de bultos: {booking_details['Numero_de_bultos']}
        📋 Orden de compra: {booking_details['Orden_de_compra']}
        
        INSTRUCCIONES:
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        • Respeta el horario reservado para tu entrega.
        • En caso de retraso, podrías tener que esperar hasta el próximo cupo disponible del día o reprogramar tu entrega.
        • Dismac no se responsabiliza por los tiempos de espera ocasionados por llegadas fuera de horario.
        • Además, según el tipo de venta, es importante considerar lo siguiente:
          - Venta al contado: Debes entregar el pedido junto con la factura a nombre del comprador y tres (3) copias de la orden de compra.
          - Venta en minicuotas: Debes entregar el pedido junto con la factura a nombre de Dismatec S.A. y una (1) copia de la orden de compra.
        
        📎 Se adjunta documento con instrucciones adicionales.
        
        REQUISITOS DE SEGURIDAD
        • Pantalón largo, sin rasgados
        • Botines de seguridad
        • Casco de seguridad
        • Chaleco o camisa con reflectivo
        • No está permitido manillas, cadenas, y principalmente masticar coca.

        Gracias por utilizar nuestro sistema de reservas.
        
        Saludos cordiales,
        Equipo de Almacén Dismac
        """
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = supplier_email
        msg['Cc'] = ', '.join(cc_emails)
        msg['Subject'] = subject
        
        # Add body
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Download and attach PDF
        pdf_data, pdf_filename = download_pdf_attachment()
        if pdf_data:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(pdf_data)
            encoders.encode_base64(attachment)
            attachment.add_header(
                'Content-Disposition',
                f'attachment; filename= {pdf_filename}'
            )
            msg.attach(attachment)
        
        # Send email
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        
        # Send to supplier + CC recipients
        all_recipients = [supplier_email] + cc_emails
        text = msg.as_string()
        server.sendmail(EMAIL_USER, all_recipients, text)
        server.quit()
        
        return True, cc_emails
        
    except Exception as e:
        st.error(f"Error enviando email: {str(e)}")
        return False, []

# ─────────────────────────────────────────────────────────────
# 4. Time Slot Functions - UPDATED FOR CONTIGUOUS SLOT LOGIC
# ─────────────────────────────────────────────────────────────
def parse_booked_slots(booked_hours):
    """Parse booked hours that may contain single or combined time slots"""
    all_booked_slots = []
    
    for booked_hora in booked_hours:
        hora_str = str(booked_hora).strip()
        
        # Skip empty or NaN values
        if not hora_str or hora_str.lower() in ['nan', 'none', '']:
            continue
        
        # Check if it contains comma (combined slots)
        if ',' in hora_str:
            # Split by comma and clean each slot
            slots = [slot.strip() for slot in hora_str.split(',')]
            for slot in slots:
                formatted_slot = format_time_slot(slot)
                if formatted_slot:
                    all_booked_slots.append(formatted_slot)
        else:
            # Single slot
            formatted_slot = format_time_slot(hora_str)
            if formatted_slot:
                all_booked_slots.append(formatted_slot)
    
    return all_booked_slots

def format_time_slot(time_str):
    """Format time string to HH:MM format, handling various input formats"""
    try:
        time_str = str(time_str).strip()
        
        # Handle different time formats that might come from Excel
        if ':' in time_str:
            parts = time_str.split(':')
            if len(parts) >= 2:
                hour = int(parts[0])
                minute = int(parts[1])
                return f"{hour:02d}:{minute:02d}"
        
        # Handle time objects or datetime objects
        if hasattr(time_str, 'hour') and hasattr(time_str, 'minute'):
            return f"{time_str.hour:02d}:{time_str.minute:02d}"
        
        return None
        
    except (ValueError, AttributeError, TypeError):
        return None

def generate_all_30min_slots():
    """Generate all possible 30-minute slots"""
    weekday_slots = []
    saturday_slots = []
    
    # Weekday slots (9:00-16:00)
    for hour in range(9, 16):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            weekday_slots.append(start_time)
    
    # Saturday slots (9:00-12:00)
    for hour in range(9, 12):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            saturday_slots.append(start_time)
    
    return weekday_slots, saturday_slots

def get_next_slot(slot_time):
    """Get the next 30-minute slot"""
    hour, minute = map(int, slot_time.split(':'))
    if minute == 0:
        next_slot = f"{hour:02d}:30"
    else:
        next_hour = hour + 1
        next_slot = f"{next_hour:02d}:00"
    return next_slot

def find_contiguous_hour_slots(all_slots, booked_slots):
    """Find available contiguous 1-hour slots from available 30-minute slots"""
    available_hour_slots = []
    
    for i in range(len(all_slots) - 1):
        current_slot = all_slots[i]
        next_slot = get_next_slot(current_slot)
        
        # Check if this is indeed the next slot in our list
        if i + 1 < len(all_slots) and all_slots[i + 1] == next_slot:
            # Both slots are available
            if current_slot not in booked_slots and next_slot not in booked_slots:
                available_hour_slots.append(current_slot)
    
    return available_hour_slots

def get_available_slots(selected_date, reservas_df, numero_bultos):
    """Get available slots for a date based on bultos count"""
    weekday_slots, saturday_slots = generate_all_30min_slots()
    
    # Sunday = 6, no work
    if selected_date.weekday() == 6:
        return []
    
    # Saturday = 5
    if selected_date.weekday() == 5:
        all_30min_slots = saturday_slots
    else:
        all_30min_slots = weekday_slots
    
    # Get booked slots for this date - IMPROVED DATE MATCHING
    target_date = selected_date.strftime('%Y-%m-%d')
    
    # Filter reservations for the selected date (handle different date formats)
    date_mask = reservas_df['Fecha'].astype(str).str.contains(target_date, na=False)
    booked_hours = reservas_df[date_mask]['Hora'].tolist()
    
    # Parse booked slots (handles combined slots)
    booked_slots = parse_booked_slots(booked_hours)
    
    if numero_bultos >= 5:
        # For 5+ bultos, find contiguous 1-hour slots
        return find_contiguous_hour_slots(all_30min_slots, booked_slots)
    else:
        # For 1-4 bultos, return available 30-minute slots
        return [slot for slot in all_30min_slots if slot not in booked_slots]

# ─────────────────────────────────────────────────────────────
# 5. Authentication Function - UPDATED TO USE ALL SHEETS
# ─────────────────────────────────────────────────────────────
def authenticate_user(usuario, password):
    """Authenticate user against Excel data and get email + CC emails"""
    credentials_df, _, _ = download_excel_to_memory()  # UPDATED - Now returns 3 values
    
    if credentials_df is None:
        return False, "Error al cargar credenciales", None, None
    
    # Clean and compare (all data is already strings)
    df_usuarios = credentials_df['usuario'].str.strip()
    
    input_usuario = str(usuario).strip()
    input_password = str(password).strip()
    
    # Find user row
    user_row = credentials_df[df_usuarios == input_usuario]
    if user_row.empty:
        return False, "Usuario no encontrado", None, None
    
    # Get stored password and clean it
    stored_password = str(user_row.iloc[0]['password']).strip()
    
    # Compare passwords
    if stored_password == input_password:
        # Get email
        email = None
        try:
            email = user_row.iloc[0]['Email']
            if str(email) == 'nan' or email is None:
                email = None
        except:
            email = None
        
        # Get CC emails
        cc_emails = []
        try:
            cc_data = user_row.iloc[0]['cc']
            if str(cc_data) != 'nan' and cc_data is not None and str(cc_data).strip():
                # Parse semicolon-separated emails
                cc_emails = [email.strip() for email in str(cc_data).split(';') if email.strip()]
        except Exception as e:
            cc_emails = []
        
        return True, "Autenticación exitosa", email, cc_emails
    
    return False, "Contraseña incorrecta", None, None

# ─────────────────────────────────────────────────────────────
# 6. Fresh slot validation function - UPDATED FOR COMBINED SLOT PARSING
# ─────────────────────────────────────────────────────────────
def check_slot_availability(selected_date, slot_time, numero_bultos):
    """Check if a specific slot is still available with fresh data - handles combined slots"""
    try:
        # Force fresh download
        download_excel_to_memory.clear()
        _, fresh_reservas_df, _ = download_excel_to_memory()
        
        if fresh_reservas_df is None:
            return False, "Error al verificar disponibilidad"
        
        # Get booked slots for this date - IMPROVED DATE MATCHING
        target_date = selected_date.strftime('%Y-%m-%d')
        date_mask = fresh_reservas_df['Fecha'].astype(str).str.contains(target_date, na=False)
        booked_hours = fresh_reservas_df[date_mask]['Hora'].tolist()
        
        # Parse booked slots (handles combined slots)
        booked_slots = parse_booked_slots(booked_hours)
        
        if numero_bultos >= 5:
            # For 5+ bultos, check both current and next slot
            next_slot = get_next_slot(slot_time)
            if slot_time in booked_slots:
                return False, "Otro proveedor acaba de reservar este horario. Por favor, elija otro."
            if next_slot in booked_slots:
                return False, "El horario siguiente necesario para su reserva de 1 hora ya está ocupado. Por favor, elija otro."
        else:
            # For 1-4 bultos, check only current slot
            if slot_time in booked_slots:
                return False, "Otro proveedor acaba de reservar este horario. Por favor, elija otro."
        
        return True, "Horario disponible"
        
    except Exception as e:
        return False, f"Error verificando disponibilidad: {str(e)}"
        
        
# ─────────────────────────────────────────────────────────────
# 7. Main App - UPDATED WORKFLOW: BULTOS FIRST, THEN DATE/TIME
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🚚 Dismac: Reserva de Entrega de Mercadería")
    
    # Download Excel when app starts - ONLY INITIAL LOAD
    with st.spinner("Cargando datos..."):
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
    
    if credentials_df is None:
        st.error("❌ Error al cargar archivo")
        return
    
    # Session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    if 'supplier_email' not in st.session_state:
        st.session_state.supplier_email = None
    if 'supplier_cc_emails' not in st.session_state:
        st.session_state.supplier_cc_emails = []
    if 'slot_error_message' not in st.session_state:
        st.session_state.slot_error_message = None
    if 'orden_compra_list' not in st.session_state:
        st.session_state.orden_compra_list = ['']
    
    # Authentication
    if not st.session_state.authenticated:
        st.subheader("🔐 Iniciar Sesión")
        
        with st.form("login_form"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar Sesión")
            
            if submitted:
                if usuario and password:
                    is_valid, message, email, cc_emails = authenticate_user(usuario, password)
                    
                    if is_valid:
                        st.session_state.authenticated = True
                        st.session_state.supplier_name = usuario
                        st.session_state.supplier_email = email
                        st.session_state.supplier_cc_emails = cc_emails
                        # Clear booking session data
                        st.session_state.orden_compra_list = ['']
                        if 'numero_bultos_input' in st.session_state:
                            del st.session_state.numero_bultos_input
                        if 'selected_slot' in st.session_state:
                            del st.session_state.selected_slot
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                else:
                    st.warning("Complete todos los campos")
    
    # Main interface after authentication
    else:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader(f"Bienvenido, {st.session_state.supplier_name}")
        with col2:
            if st.button("Cerrar Sesión"):
                st.session_state.authenticated = False
                st.session_state.supplier_name = None
                st.session_state.supplier_email = None
                st.session_state.supplier_cc_emails = []
                # Clear booking session data
                st.session_state.orden_compra_list = ['']
                if 'numero_bultos_input' in st.session_state:
                    del st.session_state.numero_bultos_input
                if 'selected_slot' in st.session_state:
                    del st.session_state.selected_slot
                st.rerun()
        
        st.markdown("---")
        
        # STEP 1: Delivery Information (MOVED TO FIRST)
        st.subheader("📦 Información de Entrega")
        st.markdown('<p style="color: red; font-size: 14px; margin-top: -10px;">Este sistema es únicamente para el registro de entregas Marketplace (no incluye entregas locales y corporativos).</p>', unsafe_allow_html=True)
        
        # Show permanent information about time slot durations
        st.info("ℹ️ **La duración del horario de reserva dependerá de la cantidad de bultos:** 1-4 bultos = 30 minutos y 5+ bultos = 1 hora")
        
        # Number of bultos (MANDATORY, NO DEFAULT)
        numero_bultos = st.number_input(
            "📦 Número de bultos *", 
            min_value=0, 
            value=None,
            key="numero_bultos_input",
            help="Cantidad de bultos o paquetes a entregar (obligatorio)",
            placeholder="Ingrese el número de bultos"
        )
        
        # Get value from session state (automatically updated by key)
        if 'numero_bultos_input' in st.session_state and st.session_state.numero_bultos_input:
            numero_bultos = st.session_state.numero_bultos_input
        
        # Multiple Purchase orders section
        st.write("📋 **Órdenes de compra** *")
        
        # Display current orden de compra inputs
        orden_compra_values = []
        for i, orden in enumerate(st.session_state.orden_compra_list):
            if len(st.session_state.orden_compra_list) == 1:
                # Single order - full width
                orden_value = st.text_input(
                    f"Orden {i+1}",
                    value=orden,
                    placeholder=f"Ej: 0000000",
                    key=f"orden_{i}"
                )
                orden_compra_values.append(orden_value)
            else:
                # Multiple orders - use columns for remove button
                col1, col2 = st.columns([5, 1])
                with col1:
                    orden_value = st.text_input(
                        f"Orden {i+1}",
                        value=orden,
                        placeholder=f"Ej: OC-2024-00{i+1}",
                        key=f"orden_{i}"
                    )
                    orden_compra_values.append(orden_value)
                with col2:
                    st.write("")  # Empty space for alignment
                    if st.button("🗑️", key=f"remove_{i}"):
                        st.session_state.orden_compra_list.pop(i)
                        st.rerun()
        
        # Update session state with current values
        st.session_state.orden_compra_list = orden_compra_values
        
        # Add button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("➕ Agregar", use_container_width=True):
                st.session_state.orden_compra_list.append('')
                st.rerun()
        
        # Check if minimum requirements are met to proceed
        valid_orders = [orden.strip() for orden in orden_compra_values if orden.strip()]
        can_proceed = numero_bultos and numero_bultos > 0 and valid_orders
        
        if not can_proceed:
            st.warning("⚠️ Complete el número de bultos y al menos una orden de compra para continuar.")
            return
        
        st.markdown("---")
        
        # STEP 2: Date selection (ONLY SHOWN AFTER BULTOS/ORDERS)
        st.subheader("📅 Seleccionar Fecha")
        st.markdown('<p style="color: red; font-size: 14px; margin-top: -10px;">Le rogamos seleccionar la fecha y el horario con atención, ya que, una vez confirmados, no podrán ser modificados ni cancelados.</p>', unsafe_allow_html=True)
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
        
        # STEP 3: Time slot selection (CONDITIONED ON BULTOS)
        st.subheader("🕐 Horarios Disponibles")
        
        # Show any persistent error message
        if st.session_state.slot_error_message:
            st.error(f"❌ {st.session_state.slot_error_message}")
        
        # Get ALL possible slots and determine availability
        weekday_slots, saturday_slots = generate_all_30min_slots()
        
        if selected_date.weekday() == 5:  # Saturday
            all_30min_slots = saturday_slots
        else:  # Monday-Friday
            all_30min_slots = weekday_slots
        
        # Get booked slots for this date
        target_date = selected_date.strftime('%Y-%m-%d')
        date_mask = reservas_df['Fecha'].astype(str).str.contains(target_date, na=False)
        booked_hours = reservas_df[date_mask]['Hora'].tolist()
        booked_slots = parse_booked_slots(booked_hours)
        
        # Generate display slots based on bultos
        if numero_bultos >= 5:
            # For 5+ bultos, show all possible 1-hour slots with availability
            display_slots = []
            for i in range(len(all_30min_slots) - 1):
                current_slot = all_30min_slots[i]
                next_slot = get_next_slot(current_slot)
                
                # Check if this is indeed the next slot in our list
                if i + 1 < len(all_30min_slots) and all_30min_slots[i + 1] == next_slot:
                    is_available = current_slot not in booked_slots and next_slot not in booked_slots
                    display_slots.append((current_slot, is_available))
        else:
            # For 1-4 bultos, show all 30-minute slots with availability
            display_slots = [(slot, slot not in booked_slots) for slot in all_30min_slots]
        
        # DEBUG: Show what slots are booked (optional - can be removed in production)
        if st.checkbox("🔍 Mostrar información de depuración", value=False):
            st.write(f"**Fecha buscada:** {target_date}")
            st.write(f"**Registros encontrados:** {len(booked_hours)}")
            st.write(f"**Horas en BD:** {booked_hours}")
            st.write(f"**Slots parseados:** {booked_slots}")
            st.write(f"**Slots a mostrar:** {len(display_slots)}")
            
            # Show all reservations for this date
            date_reservations = reservas_df[date_mask]
            if not date_reservations.empty:
                st.write("**Reservas para esta fecha:**")
                st.dataframe(date_reservations[['Fecha', 'Hora', 'Proveedor', 'Numero_de_bultos']])
            else:
                st.write("**No se encontraron reservas para esta fecha**")
        
        if not display_slots:
            st.warning("❌ No hay horarios para esta fecha")
            return
        
        # Display slots (2 per row) - SHOW ALL WITH AVAILABILITY STATUS
        selected_slot = None
        
        for i in range(0, len(display_slots), 2):
            col1, col2 = st.columns(2)
            
            # First slot
            slot1, is_available1 = display_slots[i]
            
            # Button text based on bultos and availability
            if numero_bultos >= 5:
                button_text1 = f"✅ {slot1} (1h)" if is_available1 else f"🚫 {slot1} (Ocupado)"
            else:
                button_text1 = f"✅ {slot1}" if is_available1 else f"🚫 {slot1} (Ocupado)"
            
            with col1:
                if not is_available1:
                    st.button(button_text1, disabled=True, key=f"slot_{i}", use_container_width=True)
                else:
                    if st.button(button_text1, key=f"slot_{i}", use_container_width=True):
                        # FRESH CHECK ON CLICK
                        with st.spinner("Verificando disponibilidad..."):
                            is_available, message = check_slot_availability(selected_date, slot1, numero_bultos)
                        
                        if is_available:
                            selected_slot = slot1
                            st.session_state.slot_error_message = None
                        else:
                            st.session_state.slot_error_message = message
                            st.rerun()
            
            # Second slot (if exists)
            if i + 1 < len(display_slots):
                slot2, is_available2 = display_slots[i + 1]
                
                # Button text based on bultos and availability
                if numero_bultos >= 5:
                    button_text2 = f"✅ {slot2} (1h)" if is_available2 else f"🚫 {slot2} (Ocupado)"
                else:
                    button_text2 = f"✅ {slot2}" if is_available2 else f"🚫 {slot2} (Ocupado)"
                
                with col2:
                    if not is_available2:
                        st.button(button_text2, disabled=True, key=f"slot_{i+1}", use_container_width=True)
                    else:
                        if st.button(button_text2, key=f"slot_{i+1}", use_container_width=True):
                            # FRESH CHECK ON CLICK
                            with st.spinner("Verificando disponibilidad..."):
                                is_available, message = check_slot_availability(selected_date, slot2, numero_bultos)
                            
                            if is_available:
                                selected_slot = slot2
                                st.session_state.slot_error_message = None
                            else:
                                st.session_state.slot_error_message = message
                                st.rerun()
        
        # STEP 4: Confirmation (ONLY AFTER SLOT SELECTION)
        if selected_slot or 'selected_slot' in st.session_state:
            if selected_slot:
                st.session_state.selected_slot = selected_slot
            
            st.markdown("---")
            st.subheader("✅ Confirmar Reserva")
            
            # Show summary
            duration_text = " (1 hora)" if numero_bultos >= 5 else ""
            st.info(f"📅 Fecha: {selected_date}")
            st.info(f"🕐 Horario: {st.session_state.selected_slot}{duration_text}")
            st.info(f"📦 Número de bultos: {numero_bultos}")
            st.info(f"📋 Órdenes de compra: {', '.join(valid_orders)}")
            
            # Confirm button
            if st.button("✅ Confirmar Reserva", use_container_width=True):
                with st.spinner("Verificando disponibilidad final..."):
                    is_still_available, availability_message = check_slot_availability(selected_date, st.session_state.selected_slot, numero_bultos)
                
                if not is_still_available:
                    st.error(f"❌ {availability_message}")
                    # Clear the selected slot to force reselection
                    if 'selected_slot' in st.session_state:
                        del st.session_state.selected_slot
                    st.rerun()
                    return
                
                # Join multiple orders with comma
                orden_compra_combined = ', '.join(valid_orders)
                
                # Create booking - SINGLE ROW WITH COMBINED SLOTS FOR 5+ BULTOS
                if numero_bultos >= 5:
                    # For 1-hour reservation, combine both slots in hora field
                    next_slot = get_next_slot(st.session_state.selected_slot)
                    combined_hora = f"{st.session_state.selected_slot}:00, {next_slot}:00"
                else:
                    # For 30-minute reservation, single slot
                    combined_hora = f"{st.session_state.selected_slot}:00"
                
                booking_to_save = {
                    'Fecha': selected_date.strftime('%Y-%m-%d') + ' 00:00:00',
                    'Hora': combined_hora,
                    'Proveedor': st.session_state.supplier_name,
                    'Numero_de_bultos': numero_bultos,
                    'Orden_de_compra': orden_compra_combined
                }
                
                with st.spinner("Guardando reserva..."):
                    success = save_booking_to_excel(booking_to_save)
                
                if success:
                    st.success("✅ Reserva confirmada!")
                    
                    # Send email if email is available
                    if st.session_state.supplier_email:
                        with st.spinner("Enviando confirmación por email..."):
                            email_sent, actual_cc_emails = send_booking_email(
                                st.session_state.supplier_email,
                                st.session_state.supplier_name,
                                booking_to_save,
                                st.session_state.supplier_cc_emails
                            )
                        if email_sent:
                            st.success(f"📧 Email de confirmación enviado a: {st.session_state.supplier_email}")
                            if actual_cc_emails:
                                st.success(f"📧 CC enviado a: {', '.join(actual_cc_emails)}")
                        else:
                            st.warning("⚠️ Reserva guardada pero error enviando email")
                    else:
                        st.warning("⚠️ No se encontró email para enviar confirmación")
                    
                    st.balloons()
                    
                    # Clear session data and log off user
                    st.session_state.orden_compra_list = ['']
                    if 'numero_bultos_input' in st.session_state:
                        del st.session_state.numero_bultos_input
                    st.info("Cerrando sesión automáticamente...")
                    st.session_state.authenticated = False
                    st.session_state.supplier_name = None
                    st.session_state.supplier_email = None
                    st.session_state.supplier_cc_emails = []
                    if 'selected_slot' in st.session_state:
                        del st.session_state.selected_slot
                    
                    # Wait a moment then rerun
                    import time
                    time.sleep(2)
                    st.rerun()
                else:
                    st.error("❌ Error al guardar reserva")

if __name__ == "__main__":
    main()