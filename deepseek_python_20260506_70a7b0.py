# app_seguimiento_obra.py
import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import base64
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Seguimiento de Obra Eléctrica",
    page_icon="⚡",
    layout="wide"
)

# ==================== CONSTANTES ====================
# Lista completa de tareas
TAREAS = [
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

# Estados de avance
ESTADOS = [
    "Avance de la tarea en torno al 25% aprox.",
    "Avance de la tarea en torno al 50% aprox.",
    "Avance de la tarea en torno al 75% aprox.",
    "OK, finalizado sin errores",
    "Finalizado, pero con errores pendientes de corregir",
    "Finalizado y corregidos los errores"
]

# Configuración de correo (cambiar según necesidades)
EMAIL_DESTINO = "profesora@email.com"  # Email de la profesora
EMAIL_ALUMNO = "tuemail@gmail.com"  # Email del alumno (creador)
# NOTA: La contraseña debe configurarse en secrets de Streamlit Cloud
# o en variables de entorno. Para pruebas locales, usar st.secrets

# ==================== FUNCIONES ====================
def inicializar_session_state():
    """Inicializa las variables de sesión"""
    if 'registros' not in st.session_state:
        # Intentar cargar datos existentes
        if os.path.exists("registros_obra.csv"):
            st.session_state.registros = pd.read_csv("registros_obra.csv").to_dict('records')
        else:
            st.session_state.registros = []
    if 'timestamp_inicio' not in st.session_state:
        st.session_state.timestamp_inicio = datetime.now()

def guardar_excel(registros):
    """Guarda los registros en un archivo Excel y retorna el archivo para descarga"""
    if registros:
        df = pd.DataFrame(registros)
        # Reordenar columnas
        columnas = ["Fecha_Informe", "Trabajador", "Tarea", "Estado", "Fecha_Registro"]
        df = df[columnas]
        
        # Crear archivo Excel en memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Seguimiento_Obra')
        
        return output.getvalue()
    return None

def enviar_correo(archivo_excel, destinatario):
    """
    Envía el archivo Excel por correo electrónico
    NOTA: Para usar en Streamlit Cloud, configurar secrets:
    st.secrets["email"]["user"] y st.secrets["email"]["password"]
    """
    try:
        # Configuración del correo
        remitente = EMAIL_ALUMNO
        
        # Crear mensaje
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario
        msg['Subject'] = f"Informe de Seguimiento de Obra - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Cuerpo del mensaje
        cuerpo = f"""
        Adjunto se encuentra el informe de seguimiento de obra eléctrica.
        
        Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
        Total de registros: {len(st.session_state.registros)}
        
        Este es un informe automático generado desde la aplicación de seguimiento.
        """
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        # Adjuntar archivo
        if archivo_excel:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(archivo_excel)
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="seguimiento_obra_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'
            )
            msg.attach(part)
        
        # Enviar correo
        # NOTA: Para Gmail, usar puerto 587 con STARTTLS
        # La contraseña debe ser una "Contraseña de aplicación" de Gmail
        password = st.secrets.get("EMAIL_PASSWORD", "")
        if not password:
            # Para pruebas locales, se puede usar st.text_input o variable de entorno
            st.warning("No se encontró contraseña en secrets. Usando entrada manual...")
            password = st.text_input("Ingrese la contraseña de la aplicación de correo:", type="password")
            if not password:
                return False, "No se ingresó contraseña"
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        
        return True, "Correo enviado exitosamente"
    
    except Exception as e:
        return False, f"Error al enviar correo: {str(e)}"

def verificar_tiempo():
    """Verifica si han pasado más de 2 horas desde el inicio y muestra advertencia"""
    tiempo_transcurrido = (datetime.now() - st.session_state.timestamp_inicio).total_seconds() / 3600
    if tiempo_transcurrido > 2:
        st.warning("⚠️ Han pasado más de 2 horas. Los datos se perderán pronto. Por favor, descarga o envía el Excel.")
        return True
    elif tiempo_transcurrido > 1.5:
        st.info(f"⏰ Quedan menos de {2 - tiempo_transcurrido:.1f} horas antes de que los datos se eliminen temporalmente.")
    return False

# ==================== INTERFAZ PRINCIPAL ====================
def main():
    inicializar_session_state()
    
    # Verificar tiempo
    verificar_tiempo()
    
    # Barra lateral con logo y configuración
    with st.sidebar:
        # Logo de la empresa
        st.image("logo_empresa.png", caption="Logo de la Empresa", use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 📋 Configuración del Informe")
        
        # Campo nombre del trabajador
        trabajador = st.text_input("👷 Nombre del trabajador que realiza el informe:", 
                                   key="trabajador_input")
        
        # Campo fecha de envío
        fecha_envio = st.date_input("📅 Fecha del informe:", 
                                    value=datetime.now().date(),
                                    key="fecha_envio")
        
        st.markdown("---")
        st.markdown("### 📊 Resumen")
        st.metric("Total de registros", len(st.session_state.registros))
        
        # Botón para guardar (exportar Excel)
        if st.button("💾 Descargar Excel", use_container_width=True):
            archivo_excel = guardar_excel(st.session_state.registros)
            if archivo_excel:
                b64 = base64.b64encode(archivo_excel).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="seguimiento_obra.xlsx">📥 Haga clic aquí para descargar el Excel</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
                st.sidebar.success("✅ Excel generado correctamente")
            else:
                st.sidebar.warning("No hay registros para exportar")
        
        # Botón para enviar por correo
        st.markdown("---")
        st.markdown("### 📧 Envío por Correo")
        email_destino_input = st.text_input("Destinatario:", value=EMAIL_DESTINO)
        
        if st.button("📧 Enviar Excel por Correo", use_container_width=True):
            if not trabajador:
                st.sidebar.error("Por favor, ingrese el nombre del trabajador")
            elif not st.session_state.registros:
                st.sidebar.warning("No hay registros para enviar")
            else:
                archivo_excel = guardar_excel(st.session_state.registros)
                if archivo_excel:
                    with st.spinner("Enviando correo..."):
                        exito, mensaje = enviar_correo(archivo_excel, email_destino_input)
                        if exito:
                            st.sidebar.success(mensaje)
                        else:
                            st.sidebar.error(mensaje)
                else:
                    st.sidebar.error("Error al generar el archivo Excel")
    
    # Zona principal
    st.title("⚡ Sistema de Seguimiento de Obra Eléctrica")
    st.markdown("---")
    
    # Columnas para el formulario
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📝 Registro de Avance")
        
        # Desplegable de tareas
        tarea = st.selectbox("📌 Seleccionar tarea:", TAREAS)
        
        # Desplegable de estado
        estado = st.selectbox("📊 Estado de la tarea:", ESTADOS)
        
        # Información adicional
        st.info(f"**Trabajador:** {trabajador if trabajador else 'No especificado'}")
        st.info(f"**Fecha de informe:** {fecha_envio}")
    
    with col2:
        st.subheader("📋 Últimos Registros")
        
        # Mostrar últimos registros
        if st.session_state.registros:
            df_temp = pd.DataFrame(st.session_state.registros[-5:])
            st.dataframe(df_temp[["Fecha_Informe", "Trabajador", "Tarea", "Estado"]], use_container_width=True)
        else:
            st.info("No hay registros aún. Complete el formulario y presione 'Registrar Avance'")
    
    # Botón para registrar
    st.markdown("---")
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    
    with col_btn2:
        if st.button("✅ REGISTRAR AVANCE", type="primary", use_container_width=True):
            if not trabajador:
                st.error("❌ Por favor, ingrese el nombre del trabajador en la barra lateral")
            else:
                # Crear nuevo registro
                nuevo_registro = {
                    "Fecha_Informe": fecha_envio.strftime("%d/%m/%Y"),
                    "Trabajador": trabajador,
                    "Tarea": tarea,
                    "Estado": estado,
                    "Fecha_Registro": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                }
                
                # Agregar a la lista de registros
                st.session_state.registros.append(nuevo_registro)
                
                # Guardar en CSV para persistencia temporal
                df = pd.DataFrame(st.session_state.registros)
                df.to_csv("registros_obra.csv", index=False, encoding='utf-8-sig')
                
                st.success(f"✅ Registro guardado correctamente!\n\n**Tarea:** {tarea}\n**Estado:** {estado}\n**Trabajador:** {trabajador}")
                st.balloons()
    
    # Mostrar todos los registros
    st.markdown("---")
    st.subheader("📋 Historial Completo de Registros")
    
    if st.session_state.registros:
        df_completo = pd.DataFrame(st.session_state.registros)
        st.dataframe(df_completo, use_container_width=True, height=400)
        
        # Estadísticas rápidas
        col_est1, col_est2, col_est3, col_est4 = st.columns(4)
        with col_est1:
            tareas_completadas = sum(1 for r in st.session_state.registros if "finalizado sin errores" in r["Estado"])
            st.metric("Tareas Completadas ✅", tareas_completadas)
        with col_est2:
            tareas_pendientes = sum(1 for r in st.session_state.registros if "25%" in r["Estado"] or "50%" in r["Estado"] or "75%" in r["Estado"])
            st.metric("Tareas en Progreso 🔄", tareas_pendientes)
        with col_est3:
            tareas_con_errores = sum(1 for r in st.session_state.registros if "errores pendientes" in r["Estado"])
            st.metric("Tareas con Errores ⚠️", tareas_con_errores)
        with col_est4:
            tareas_totales = len(set(r["Tarea"] for r in st.session_state.registros))
            st.metric("Tareas Registradas 📌", tareas_totales)
    else:
        st.info("No hay registros aún. Complete el formulario arriba para comenzar.")

# ==================== EJECUCIÓN ====================
if __name__ == "__main__":
    main()