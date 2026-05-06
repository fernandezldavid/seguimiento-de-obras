import streamlit as st
import pandas as pd
from datetime import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Seguimiento de Obra - Fundación Masaveu",
    page_icon="🏗️",
    layout="centered"
)

# --- 1. INCORPORAR IMAGEN DEL LOGO ---
# Reemplaza 'logo.png' por la ruta o nombre de tu archivo de imagen en GitHub
try:
    st.image("logo.png", width=250)
except FileNotFoundError:
    st.warning("⚠️ No se encontró el archivo 'logo.png'. Por favor, asegúrate de subirlo a tu repositorio de GitHub.")

st.title("🏗️ Control y Seguimiento de Obra")
st.write("Registra el avance de las tareas y exporta los datos de forma segura.")

# --- INICIALIZACIÓN DEL HISTORIAL DE REGISTROS ---
# Usamos session_state para que los datos no se borren mientras el usuario añade registros en la sesión activa.
if "historico_datos" not in st.session_state:
    st.session_state.historico_datos = []

# --- FORMULARIO DE ENTRADA DE DATOS ---
st.subheader("📝 Nuevo Registro de Avance")

with st.form("formulario_registro", clear_on_submit=True):
    # Campo para el nombre del trabajador
    trabajador = st.text_input("Nombre del Trabajador:", placeholder="Ej. Juan Pérez")
    
    # Campo para la fecha de envío
    fecha_envio = st.date_input("Fecha del Reporte:", value=datetime.today())
    
    # Desplegable de Tareas de la Obra
    tareas_disponibles = [
        "Trazado y marcado de cajas, tubos y cuadros",
        "Ejecución rozas en paredes y techos",
        "Montaje de soportes",
        "Colocación tubos y conductos",
        "Tendido de cables",
        "Identificación y etiquetado",
        "Conexionado de cables en bornes o regletas",
        "Instalación y conexionado de mecanismos",
        "Fijación de carril DIN y mecanismos en cuadro eléctrico",
        "Cableado interno del cuadro eléctrico",
        "Configuración de equipos domóticos y/o automáticos",
        "Conexionado de sensores/actuadores de equipos domóticos/automáticos",
        "Pruebas de continuidad",
        "Pruebas de aislamiento",
        "Verificación de tierras",
        "Programación del automatismo",
        "Pruebas de funcionamiento"
    ]
    tarea_seleccionada = st.selectbox("Seleccione la Tarea:", tareas_disponibles)
    
    # Desplegable del Estado de la Tarea
    estados_disponibles = [
        "Avance de la tarea en torno al 25% aprox.",
        "Avance de la tarea en torno al 50% aprox.",
        "Avance de la tarea en torno al 75% aprox.",
        "OK, finalizado sin errores",
        "Finalizado, pero con errores pendientes de corregir",
        "Finalizado y corregidos los errores"
    ]
    estado_seleccionado = st.selectbox("Estado de la Tarea:", estados_disponibles)
    
    # Botón para añadir el registro a la lista temporal
    boton_guardar = st.form_submit_button("Guardar Registro")

if boton_guardar:
    if trabajador.strip() == "":
        st.error("❌ Por favor, introduce el nombre del trabajador antes de guardar.")
    else:
        # Añadir datos al estado de la sesión
        nuevo_registro = {
            "Fecha": fecha_envio.strftime("%Y-%m-%d"),
            "Trabajador": trabajador,
            "Tarea": tarea_seleccionada,
            "Estado": estado_seleccionado
        }
        st.session_state.historico_datos.append(nuevo_registro)
        st.success(f"✔️ Registro añadido: {tarea_seleccionada} - {estado_seleccionado}")

# --- MOSTRAR TABLA DE DATOS ACTUALES ---
st.write("---")
st.subheader("📊 Registros acumulados en esta sesión")

if st.session_state.historico_datos:
    df = pd.DataFrame(st.session_state.historico_datos)
    st.dataframe(df, use_container_width=True)
    
    # --- GENERAR EXCEL EN MEMORIA (BytesIO) ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Seguimiento_Obra')
    buffer.seek(0)
    
    # --- BOTÓN DE DESCARGA DIRECTA ---
    st.download_button(
        label="📥 Descargar Excel",
        data=buffer,
        file_name=f"seguimiento_obra_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # --- ENVÍO POR CORREO ELECTRÓNICO (SMTP) ---
    st.write("---")
    st.subheader("📧 Enviar Reporte por Correo")
    
    # Obtención segura de credenciales desde Streamlit Secrets
    # (Explicado en el apartado de configuración más abajo)
    try:
        EMAIL_EMISOR = st.secrets["correo_alumno"]          # Tu cuenta de correo
        EMAIL_PASS = st.secrets["contrasena_aplicacion"]    # Tu contraseña de aplicación generada
        EMAIL_RECEPTOR = st.secrets["correo_profesora"]     # Correo de destino de la profesora
    except KeyError:
        EMAIL_EMISOR = None
        EMAIL_PASS = None
        EMAIL_RECEPTOR = None
        st.info("💡 Para habilitar el envío automático por correo electrónico, configura los Secrets en el panel de Streamlit.")

    if EMAIL_EMISOR and EMAIL_PASS and EMAIL_RECEPTOR:
        if st.button("📨 Enviar Excel por Correo"):
            try:
                # Creación del mensaje
                msg = MIMEMultipart()
                msg['From'] = EMAIL_EMISOR
                msg['To'] = EMAIL_RECEPTOR
                msg['Cc'] = EMAIL_EMISOR  # Copia para el alumno
                msg['Subject'] = f"Reporte de Seguimiento de Obra - {datetime.now().strftime('%d/%m/%Y')}"
                
                cuerpo = f"Hola,\n\nSe adjunta el reporte de seguimiento de obra generado por el alumno.\n\nAtentamente,\nApp de Seguimiento"
                msg.attach(MIMEText(cuerpo, 'plain'))
                
                # Adjuntar el archivo Excel
                part = MIMEBase('application', "octet-stream")
                part.set_payload(buffer.getvalue())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="seguimiento_obra.xlsx"')
                msg.attach(part)
                
                # Configurar el servidor SMTP (en este ejemplo usamos Gmail, puerto 587)
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(EMAIL_EMISOR, EMAIL_PASS)
                
                # Destinatarios (Tanto la profesora como el alumno en copia)
                destinatarios = [EMAIL_RECEPTOR, EMAIL_EMISOR]
                server.sendmail(EMAIL_EMISOR, destinatarios, msg.as_string())
                server.quit()
                
                st.success("📧 ¡Correo electrónico enviado con éxito a la profesora y con copia para ti!")
            except Exception as e:
                st.error(f"❌ Error al enviar el correo: {e}")
    else:
        st.warning("⚠️ Configura las credenciales secretas en la plataforma de Streamlit para poder usar la función de envío de correo.")
else:
    st.info("Aún no hay registros en la sesión actual. Rellena el formulario de arriba para empezar.")
