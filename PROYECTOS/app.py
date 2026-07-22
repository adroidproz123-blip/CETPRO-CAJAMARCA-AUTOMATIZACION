import streamlit as st
import pandas as pd

# 1. Configuración visual de la página
st.set_page_config(page_title="Convenios - CETPRO Cajamarca", page_icon="🔑", layout="centered")

# 2. Control del estado de la sesión (para saber si ya inició sesión)
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

# --- PANTALLA DE LOGEO (INGRESO) ---
def login_screen():
    st.title("🔑 Sistema de Convenios")
    st.subheader("CETPRO Cajamarca")
    
    # Formulario de acceso
    with st.form("login_form"):
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        submit_button = st.form_submit_button("Ingresar")
        
        if submit_button:
            # Aquí defines el usuario y contraseña para entrar
            if username == "admin" and password == "cajamarca2026":
                st.session_state['logged_in'] = True
                st.rerun() # Recarga la página para entrar al sistema
            else:
                st.error("Usuario o contraseña incorrectos. Inténtalo de nuevo.")

# --- PANTALLA DEL BUSCADOR (SISTEMA PRINCIPAL) ---
def main_system():
    st.title("🔍 Buscador de Convenios")
    st.write("Ingresa los datos para realizar la consulta en la base de datos.")
    
    # Botón para salir en la barra lateral
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state['logged_in'] = False
        st.rerun()
    
    # Intentar leer tu archivo real de Excel
    try:
        # Buscamos el archivo en la misma carpeta donde ejecutes el script
        # Si tiene otro nombre exacto, lo cambias aquí:
        df = pd.read_excel("BD_Convenios_CETPRO_CAJAMARCA.xlsx", sheet_name="BD_CONVENIOS", skiprows=2)
    except Exception:
        # Si no encuentra tu Excel, usa estos datos de prueba para que no se caiga la página
        datos_ejemplo = {
            'DNI_REPRESENTANTE': ['12345678', ''],
            'RUC': ['', '20510752938'],
            'EMPRESA_INSTITUCION': ['Juan Pérez', 'MANTENIMIENTO INDUSTRIAL Y COMERCIAL SAC'],
            'PROGRAMA_FORMATIVO': ['Electricidad', 'Computación'],
            'ESTADO': ['VIGENTE', 'VENCIDO'],
            'FECHA_FIN': ['2027-04-10', '2025-05-10']
        }
        df = pd.DataFrame(datos_ejemplo)

    # El buscador interactivo en la web
    busqueda = st.text_input("Escribe el DNI, RUC o Nombre de la empresa:").strip()

    if busqueda