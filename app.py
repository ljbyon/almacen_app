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
    """Save new booking to Excel file - PRESERVES ALL SHEETS - UPDATED FOR MULTIPLE SLOTS"""
    try:
        # Load current data
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
        
        if reservas_df is None:
            st.error("❌ No se pudo cargar el archivo Excel")
            return False
        
        # Handle multiple bookings for 1-hour slots
        if isinstance(new_booking, list):
            # Multiple bookings (for 1-hour slots)
            new_rows = pd.DataFrame(new_booking)
            updated_reservas_df = pd.concat([reservas_df, new_rows], ignore_index=True)
        else:
            # Single booking
            new_row = pd.DataFrame([new_booking])
            updated_reservas_df = pd.concat([reservas_df, new_row], ignore_index=True)
        
        # Authenticate and upload
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Create Excel file - SAVE ALL SHEETS
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            credentials_df.to_excel(writer, sheet_name="proveedor_credencial", index=False)
            updated_reservas_df.to_excel(writer, sheet_name="proveedor_reservas", index=False)
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
    """Send booking confirmation email with PDF attachment - UPDATED FOR NEW FLOW"""
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
        
        # Format dates for email display - UPDATED FOR NEW FLOW
        display_fecha = booking_details['Fecha'].split(' ')[0]  # Remove time part for display
        display_hora = booking_details['Hora']  # This now contains the full time range for 1-hour slots
        
        body = f"""
        Hola {supplier_name},
        
        Su reserva de entrega ha sido confirmada exitosamente.
        
        DETALLES DE LA RESERVA:
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        📅 Fecha: {display_fecha}
        🕐 Horario: {display_hora}
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
# 4. Time Slot Functions - UPDATED FOR NEW FLOW
# ─────────────────────────────────────────────────────────────
def generate_time_slots():
    """Generate available time slots - showing start time only"""
    # Monday-Friday: 9:00-16:00, Saturday: 9:00-12:00
    weekday_slots = []
    saturday_slots = []
    
    # Weekday slots (9:00-16:00)
    start_hour = 9
    end_hour = 16
    for hour in range(start_hour, end_hour):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            weekday_slots.append(start_time)
    
    # Saturday slots (9:00-12:00)
    for hour in range(9, 12):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            saturday_slots.append(start_time)
    
    return weekday_slots, saturday_slots

def get_available_slots_by_package_count(selected_date, reservas_df, numero_bultos):
    """Get available slots for a date based on package count - NEW FUNCTION"""
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
    date_str = selected_date.strftime('%Y-%m-%d') + ' 00:00:00'
    booked_reservas = reservas_df[reservas_df['Fecha'] == date_str]['Hora'].tolist()
    
    # Convert booked slots to "09:00" format for comparison
    booked_slots = []
    for booked_hora in booked_reservas:
        if ':' in str(booked_hora):
            parts = str(booked_hora).split(':')
            formatted_slot = f"{int(parts[0]):02d}:{parts[1]}"
            booked_slots.append(formatted_slot)
        else:
            booked_slots.append(str(booked_hora))
    
    # Get available individual slots
    available_slots = [slot for slot in all_slots if slot not in booked_slots]
    
    # Return based on package count
    if numero_bultos <= 4:
        # Return individual 30-minute slots
        return [(slot, slot) for slot in available_slots]  # (display_time, booking_slots)
    else:
        # Return 1-hour contiguous slots
        contiguous_slots = []
        for i in range(len(available_slots) - 1):
            slot1 = available_slots[i]
            slot2 = available_slots[i + 1]
            
            # Check if slots are contiguous (30 minutes apart)
            hour1, min1 = map(int, slot1.split(':'))
            hour2, min2 = map(int, slot2.split(':'))
            
            time1_minutes = hour1 * 60 + min1
            time2_minutes = hour2 * 60 + min2
            
            if time2_minutes - time1_minutes == 30:
                # Create 1-hour slot display
                end_hour = hour2
                end_min = min2 + 30
                if end_min >= 60:
                    end_hour += 1
                    end_min -= 60
                
                display_time = f"{slot1} - {end_hour:02d}:{end_min:02d}"
                contiguous_slots.append((display_time, [slot1, slot2]))
        
        return contiguous_slots

def check_slot_availability_new_flow(selected_date, slot_info, numero_bultos):
    """Check if a specific slot is still available with fresh data - NEW FLOW"""
    try:
        # Force fresh download
        download_excel_to_memory.clear()
        _, fresh_reservas_df, _ = download_excel_to_memory()
        
        if fresh_reservas_df is None:
            return False, "Error al verificar disponibilidad"
        
        # Get booked slots for this date
        date_str = selected_date.strftime('%Y-%m-%d') + ' 00:00:00'
        booked_reservas = fresh_reservas_df[fresh_reservas_df['Fecha'] == date_str]['Hora'].tolist()
        
        # Convert booked slots to "09:00" format for comparison
        booked_slots = []
        for booked_hora in booked_reservas:
            if ':' in str(booked_hora):
                parts = str(booked_hora).split(':')
                formatted_slot = f"{int(parts[0]):02d}:{parts[1]}"
                booked_slots.append(formatted_slot)
        
        # Check availability based on package count
        if numero_bultos <= 4:
            # Single slot check
            slot_time = slot_info[1]  # booking_slots is just the slot time
            if slot_time in booked_slots:
                return False, "Otro proveedor acaba de reservar este horario. Por favor, elija otro."
        else:
            # Check both slots for 1-hour reservation
            slots_to_check = slot_info[1]  # booking_slots is a list of two slots
            for slot in slots_to_check:
                if slot in booked_slots:
                    return False, "Otro proveedor acaba de reservar parte de este horario. Por favor, elija otro."
        
        return True, "Horario disponible"
        
    except Exception as e:
        return False, f"Error verificando disponibilidad: {str(e)}"

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
# 6. Main App - UPDATED WITH NEW FLOW
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
    if 'booking_step' not in st.session_state:
        st.session_state.booking_step = 1
    if 'selected_date' not in st.session_state:
        st.session_state.selected_date = None
    if 'numero_bultos' not in st.session_state:
        st.session_state.numero_bultos = 1
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
                        # Reset booking flow
                        st.session_state.booking_step = 1
                        st.session_state.selected_date = None
                        st.session_state.numero_bultos = 1
                        st.session_state.orden_compra_list = ['']
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                else:
                    st.warning("Complete todos los campos")
    
    # NEW BOOKING FLOW
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
                # Reset booking flow
                st.session_state.booking_step = 1
                st.session_state.selected_date = None
                st.session_state.numero_bultos = 1
                st.session_state.orden_compra_list = ['']
                st.rerun()
        
        st.markdown("---")
        
        # STEP 1: DATE AND PACKAGE COUNT SELECTION
        if st.session_state.booking_step == 1:
            st.subheader("📅 Paso 1: Seleccionar Fecha y Número de Bultos")
            st.markdown('<p style="color: red; font-size: 14px; margin-top: -10px;">Le rogamos seleccionar la fecha con atención, ya que, una vez confirmada, no podrá ser modificada ni cancelada.</p>', unsafe_allow_html=True)
            
            today = datetime.now().date()
            max_date = today + timedelta(days=30)
            
            selected_date = st.date_input(
                "Fecha de entrega",
                min_value=today,
                max_value=max_date,
                value=today,
                key="date_input"
            )
            
            # Check if Sunday
            if selected_date.weekday() == 6:
                st.warning("⚠️ No trabajamos los domingos")
                return
            
            # Number of packages
            numero_bultos = st.number_input(
                "📦 Número de bultos", 
                min_value=1, 
                value=st.session_state.numero_bultos,
                help="Cantidad de bultos o paquetes a entregar"
            )
            
            # Package count info
            if numero_bultos <= 4:
                st.info("💡 Con 1-4 bultos, podrá reservar slots de 30 minutos")
            else:
                st.info("💡 Con 5 o más bultos, podrá reservar slots de 1 hora")
            
            col1, col2, col3 = st.columns([1, 1, 1])
            with col2:
                if st.button("Continuar ➡️", use_container_width=True):
                    st.session_state.selected_date = selected_date
                    st.session_state.numero_bultos = numero_bultos
                    st.session_state.booking_step = 2
                    st.rerun()
        

        # STEP 2: PURCHASE ORDERS
        elif st.session_state.booking_step == 2:
            st.subheader("📋 Paso 2: Órdenes de Compra")
            st.info(f"📅 Fecha seleccionada: {st.session_state.selected_date}")
            st.info(f"📦 Número de bultos: {st.session_state.numero_bultos}")
            
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
                        placeholder=f"Ej: OC-2024-00{i+1}",
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
                if st.button("➕ Agregar otra orden", use_container_width=True):
                    st.session_state.orden_compra_list.append('')
                    st.rerun()
            
            # Navigation buttons
            col1, col2, col3 = st.columns([1, 1, 1])
            with col1:
                if st.button("⬅️ Volver", use_container_width=True):
                    st.session_state.booking_step = 1
                    st.rerun()
            with col3:
                # Validate before continuing
                valid_orders = [orden.strip() for orden in orden_compra_values if orden.strip()]
                if valid_orders:
                    if st.button("Continuar ➡️", use_container_width=True):
                        st.session_state.booking_step = 3
                        st.rerun()
                else:
                    st.button("Continuar ➡️", disabled=True, use_container_width=True)
                    st.error("❌ Al menos una orden de compra es obligatoria")
        
        # STEP 3: TIME SLOT SELECTION
        elif st.session_state.booking_step == 3:
            st.subheader("🕐 Paso 3: Seleccionar Horario")
            st.info(f"📅 Fecha: {st.session_state.selected_date}")
            st.info(f"📦 Bultos: {st.session_state.numero_bultos}")
            
            # Show any persistent error message
            if st.session_state.slot_error_message:
                st.error(f"❌ {st.session_state.slot_error_message}")
            
            # Get available slots based on package count
            available_slot_info = get_available_slots_by_package_count(
                st.session_state.selected_date, 
                reservas_df, 
                st.session_state.numero_bultos
            )
            
            if not available_slot_info:
                st.warning("❌ No hay horarios disponibles para esta fecha")
                col1, col2, col3 = st.columns([1, 1, 1])
                with col1:
                    if st.button("⬅️ Volver", use_container_width=True):
                        st.session_state.booking_step = 2
                        st.rerun()
                return
            
            # Display slots (2 per row)
            selected_slot_info = None
            
            for i in range(0, len(available_slot_info), 2):
                col1, col2 = st.columns(2)
                
                # First slot
                display_time1, booking_slots1 = available_slot_info[i]
                
                with col1:
                    slot_text = f"✅ {display_time1}"
                    if st.session_state.numero_bultos > 4:
                        slot_text += " (1 hora)"
                    
                    if st.button(slot_text, key=f"slot_{i}", use_container_width=True):
                        # FRESH CHECK ON CLICK
                        with st.spinner("Verificando disponibilidad..."):
                            is_available, message = check_slot_availability_new_flow(
                                st.session_state.selected_date, 
                                (display_time1, booking_slots1), 
                                st.session_state.numero_bultos
                            )
                        
                        if is_available:
                            selected_slot_info = (display_time1, booking_slots1)
                            st.session_state.slot_error_message = None
                        else:
                            st.session_state.slot_error_message = message
                            st.rerun()
                
                # Second slot (if exists)
                if i + 1 < len(available_slot_info):
                    display_time2, booking_slots2 = available_slot_info[i + 1]
                    
                    with col2:
                        slot_text = f"✅ {display_time2}"
                        if st.session_state.numero_bultos > 4:
                            slot_text += " (1 hora)"
                        
                        if st.button(slot_text, key=f"slot_{i+1}", use_container_width=True):
                            # FRESH CHECK ON CLICK
                            with st.spinner("Verificando disponibilidad..."):
                                is_available, message = check_slot_availability_new_flow(
                                    st.session_state.selected_date, 
                                    (display_time2, booking_slots2), 
                                    st.session_state.numero_bultos
                                )
                            
                            if is_available:
                                selected_slot_info = (display_time2, booking_slots2)
                                st.session_state.slot_error_message = None
                            else:
                                st.session_state.slot_error_message = message
                                st.rerun()
            
            # Navigation and booking
            col1, col2, col3 = st.columns([1, 1, 1])
            with col1:
                if st.button("⬅️ Volver", use_container_width=True):
                    st.session_state.booking_step = 2
                    st.rerun()
            
            # Final booking confirmation
            if selected_slot_info:
                st.markdown("---")
                st.subheader("✅ Confirmar Reserva")
                
                display_time, booking_slots = selected_slot_info
                
                # Show booking summary
                st.info(f"📅 Fecha: {st.session_state.selected_date}")
                st.info(f"🕐 Horario: {display_time}")
                st.info(f"📦 Bultos: {st.session_state.numero_bultos}")
                
                valid_orders = [orden.strip() for orden in st.session_state.orden_compra_list if orden.strip()]
                st.info(f"📋 Órdenes: {', '.join(valid_orders)}")
                
                if st.button("✅ Confirmar Reserva", use_container_width=True):
                    with st.spinner("Verificando disponibilidad final..."):
                        is_still_available, availability_message = check_slot_availability_new_flow(
                            st.session_state.selected_date, 
                            selected_slot_info, 
                            st.session_state.numero_bultos
                        )
                    
                    if not is_still_available:
                        st.error(f"❌ {availability_message}")
                        st.rerun()
                        return
                    
                    # Prepare booking data
                    orden_compra_combined = ', '.join(valid_orders)
                    
                    # Create booking(s) based on package count
                    if st.session_state.numero_bultos <= 4:
                        # Single 30-minute slot
                        new_booking = {
                            'Fecha': st.session_state.selected_date.strftime('%Y-%m-%d') + ' 00:00:00',
                            'Hora': booking_slots + ':00',
                            'Proveedor': st.session_state.supplier_name,
                            'Numero_de_bultos': st.session_state.numero_bultos,
                            'Orden_de_compra': orden_compra_combined
                        }
                    else:
                        # Two 30-minute slots for 1-hour reservation
                        new_booking = []
                        for slot in booking_slots:
                            booking_entry = {
                                'Fecha': st.session_state.selected_date.strftime('%Y-%m-%d') + ' 00:00:00',
                                'Hora': slot + ':00',
                                'Proveedor': st.session_state.supplier_name,
                                'Numero_de_bultos': st.session_state.numero_bultos,
                                'Orden_de_compra': orden_compra_combined
                            }
                            new_booking.append(booking_entry)
                    
                    with st.spinner("Guardando reserva..."):
                        success = save_booking_to_excel(new_booking)
                    
                    if success:
                        st.success("✅ Reserva confirmada!")
                        
                        # Prepare email data
                        email_booking_data = {
                            'Fecha': st.session_state.selected_date.strftime('%Y-%m-%d') + ' 00:00:00',
                            'Hora': display_time,  # Use the display time for email
                            'Numero_de_bultos': st.session_state.numero_bultos,
                            'Orden_de_compra': orden_compra_combined
                        }
                        
                        # Send email if email is available
                        if st.session_state.supplier_email:
                            with st.spinner("Enviando confirmación por email..."):
                                email_sent, actual_cc_emails = send_booking_email(
                                    st.session_state.supplier_email,
                                    st.session_state.supplier_name,
                                    email_booking_data,
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
                        
                        # Reset and log off user
                        st.info("Cerrando sesión automáticamente...")
                        st.session_state.authenticated = False
                        st.session_state.supplier_name = None
                        st.session_state.supplier_email = None
                        st.session_state.supplier_cc_emails = []
                        st.session_state.booking_step = 1
                        st.session_state.selected_date = None
                        st.session_state.numero_bultos = 1
                        st.session_state.orden_compra_list = ['']
                        
                        # Wait a moment then rerun
                        import time
                        time.sleep(2)
                        st.rerun()
                    else:
                        st.error("❌ Error al guardar reserva")

if __name__ == "__main__":
    main()