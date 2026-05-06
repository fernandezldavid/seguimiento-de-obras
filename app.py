
import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
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
# Comprobamos de forma segura si la imagen existe en el repositorio
nombre_logo = "logo.png"
if os.path.exists(nombre_logo):
    st.image(nombre_logo, width=220)
else:
    st.info("💡 Consejo: Sube tu archivo 'logo.png' a la raíz de tu repositorio de GitHub para que aparezca aquí el logo de la empresa.")

st.title("🏗️ Control y Seguimiento de Obra")
st.write("Registra el avance de las tareas en tiempo real, genera reportes en Excel y envíalos por correo.")

# --- INICIALIZACIÓN DEL HISTORIAL EN SESIÓN ---
# Evita que se borren los registros mientras el usuario añade datos durante la sesión activa
if "historico_datos" not in st.session_state:
    st.session_state.historico_datos = []

# --- FORMULARIO DE ENTRADA DE DATOS ---
st.subheader("📝 Nuevo Registro de Avance")

with st.form("formulario_registro", clear_on_submit=True):
    # Campo: Nombre del trabajador
    trabajador = st.text_input("Nombre del Trabajador:", placeholder="Ej. Pedro Masaveu")
    
    # Campo: Fecha del reporte
    fecha_envio = st.date_input("Fecha del Reporte:", value=datetime.today())
    
    # Desplegable: Tareas de la Obra
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
    
    # Desplegable: Estado de la Tarea
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
        st.error("❌ Por favor, introduce el nombre del trabajador antes de guardar el registro.")
    else:
        # Añadir datos al estado de la sesión
        nuevo_registro = {
            "Fecha": fecha_envio.strftime("%Y-%m-%d"),
            "Trabajador": trabajador,
            "Tarea": tarea_seleccionada,
            "Estado": estado_seleccionado
        }
        st.session_state.historico_datos.append(nuevo_registro)
        st.success(f"✔️ Registro añadido: {tarea_seleccionada} -> {estado_seleccionado}")

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
    try:
        EMAIL_EMISOR = st.secrets["correo_alumno"]          # Tu cuenta de correo
        EMAIL_PASS = st.secrets["contrasena_aplicacion"]    # Tu contraseña de aplicación generada
        EMAIL_RECEPTOR_1 = st.secrets["correo_profesora"]   # fmo@fundacionmasaveu.com (Profesora 1)
    except KeyError:
        EMAIL_EMISOR = None
        EMAIL_PASS = None
        EMAIL_RECEPTOR_1 = None
        st.info("💡 Para habilitar el envío automático por correo electrónico, configura las variables en la sección de 'Secrets' de tu app en Streamlit Cloud.")

    # Definimos el correo de Ana de manera directa
    EMAIL_RECEPTOR_2 = "ana@fundacionmasaveu.com"

    if EMAIL_EMISOR and EMAIL_PASS and EMAIL_RECEPTOR_1:
        if st.button("📨 Enviar Excel por Correo"):
            try:
                # Creación del mensaje de correo electrónico
                msg = MIMEMultipart()
                msg['From'] = EMAIL_EMISOR
                
                # Definimos la lista de destinatarios principales (las dos profesoras)
                destinatarios_principales = [EMAIL_RECEPTOR_1, EMAIL_RECEPTOR_2]
                msg['To'] = ", ".join(destinatarios_principales)
                msg['Cc'] = EMAIL_EMISOR  # Copia para ti como alumno
                
                msg['Subject'] = f"Reporte de Seguimiento de Obra - {datetime.now().strftime('%d/%m/%Y')}"
                
                cuerpo = f"Hola,\n\nSe adjunta el reporte de seguimiento de obra en formato Excel generado por el alumno.\n\nAtentamente,\nApp de Seguimiento"
                msg.attach(MIMEText(cuerpo, 'plain'))
                
                # Adjuntar el archivo Excel desde la memoria (buffer)
                part = MIMEBase('application', "octet-stream")
                part.set_payload(buffer.getvalue())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="seguimiento_obra.xlsx"')
                msg.attach(part)
                
                # Configuración del servidor SMTP (por defecto para cuentas de Gmail)
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(EMAIL_EMISOR, EMAIL_PASS)
                
                # Lista total de personas a las que el servidor SMTP les enviará el mail (profesoras + alumno)
                todos_los_destinatarios = destinatarios_principales + [EMAIL_EMISOR]
                
                server.sendmail(EMAIL_EMISOR, todos_los_destinatarios, msg.as_string())
                server.quit()
                
                st.success("📧 ¡Correo electrónico enviado con éxito a las profesoras y con copia a tu correo!")
            except Exception as e:
                st.error(f"❌ Error al enviar el correo: {e}")
    else:
        st.warning("⚠️ El envío por correo no está activo todavía. Añade tus credenciales seguras en el panel de control de Streamlit Cloud.")
else:
    st.info("Aún no has añadido registros en esta sesión. Completa el formulario de arriba y haz clic en 'Guardar Registro'.")
