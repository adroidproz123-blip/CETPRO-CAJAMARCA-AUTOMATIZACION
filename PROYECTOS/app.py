import streamlit as st
import pandas as pd

# 1. Configuración visual de la página
st.set_page_config(page_title="Convenios - CETPRO Cajamarca", page_icon="🔑", layout="centered")

# 2. Control del estado de la sesión
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

# --- PANTALLA DE LOGEO (INGRESO) ---
def login_screen():
    st.title("🔑 Sistema de Convenios")
    st.subheader("CETPRO Cajamarca")
    
    with st.form("login_form"):
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        submit_button = st.form_submit_button("Ingresar")
        
        if submit_button:
            if username == "admin" and password == "cajamarca2026":
                st.session_state['logged_in'] = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos. Inténtalo de nuevo.")

# --- PANTALLA DEL BUSCADOR (SISTEMA PRINCIPAL) ---
def main_system():
    st.title("🔍 Buscador de Convenios")
    st.write("Ingresa los datos para realizar la consulta en la base de datos.")
    
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state['logged_in'] = False
        st.rerun()
    
    # Lectura del archivo de Excel con el nombre correcto de tu repositorio
    try:
        df = pd.read_excel("BD_CONVENIOS.xlsx", sheet_name="BD_CONVENIOS", skiprows=2)
    except Exception:
        datos_ejemplo = {
            'DNI_REPRESENTANTE': ['12345678', ''],
            'RUC': ['', '20510752938'],
            'EMPRESA_INSTITUCION': ['Juan Pérez', 'MANTENIMIENTO INDUSTRIAL Y COMERCIAL SAC'],
            'PROGRAMA_FORMATIVO': ['Electricidad', 'Computación'],
            'ESTADO': ['VIGENTE', 'VENCIDO'],
            'FECHA_FIN': ['2027-04-10', '2025-05-10']
        }
        df = pd.DataFrame(datos_ejemplo)

    busqueda = st.text_input("Escribe el DNI, RUC o Nombre de la empresa:").strip()

    if busqueda:
        resultado = df[
            df['DNI_REPRESENTANTE'].astype(str).str.contains(busqueda, case=False, na=False) |
            df['RUC'].astype(str).str.contains(busqueda, case=False, na=False) |
            df['EMPRESA_INSTITUCION'].astype(str).str.contains(busqueda, case=False, na=False)
        ]
        
        if not resultado.empty:
            st.success(f"Se encontraron {len(resultado)} coincidencia(s):")
            st.dataframe(resultado)
        else:
            st.warning("No se encontraron registros que coincidan con la búsqueda.")

# Control principal de pantallas
if not st.session_state['logged_in']:
    login_screen()
else:
    main_system()
