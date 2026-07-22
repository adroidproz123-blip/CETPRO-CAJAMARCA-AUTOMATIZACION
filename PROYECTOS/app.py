# -*- coding: utf-8 -*-
"""
==================================================================
  SISTEMA CETPRO CAJAMARCA - CONVENIOS Y CONSTANCIAS EFSRT
==================================================================
Aplicación 100% local (no realiza consultas a internet).

- Módulo de Convenios: busca por DNI del representante, RUC,
  empresa, representante o programa formativo.
- Módulo de Constancias: busca por DNI, apellidos, nombres,
  programa o módulo, y muestra el enlace a Google Drive.

Detección automática de archivos:
- No hace falta que el Excel se llame exactamente igual. Basta con
  que el nombre del archivo .xlsx contenga la palabra "convenio" o
  "constancia" para que la app lo detecte solo.
- Cada vez que se sube (o reemplaza) una base de datos actualizada,
  la app la vuelve a leer automáticamente: se controla la fecha de
  modificación del archivo, así que no hace falta limpiar el caché
  ni reiniciar la app manualmente.

Autenticación (modifícala en la sección CONFIGURACIÓN):
    Usuario:    CETPRO CAJAMARCA
    Contraseña: cetprocajamarca#2026
==================================================================
"""

import base64
import os
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st

# ==================================================================
# 0. ANCLAR RUTAS A LA CARPETA DONDE VIVE app.py
# ==================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

# ==================================================================
# 1. CONFIGURACIÓN GENERAL
# ==================================================================
st.set_page_config(
    page_title="CETPRO Cajamarca | Gestión Institucional",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

USUARIO_VALIDO = "CETPRO CAJAMARCA"
CONTRASENA_VALIDA = "cetprocajamarca#2026"

# --- Convenios ---
HOJA_DATOS = "BD_CONVENIOS"
FILA_ENCABEZADO = 3  # fila 4 del Excel (índice 0 = fila 1) contiene los encabezados
PALABRA_CLAVE_CONVENIOS = "convenio"

# --- Constancias EFSRT ---
HOJA_CONSTANCIAS = "BD_CONSTANCIAS"
FILA_ENCABEZADO_CONSTANCIAS = 0  # encabezados en la fila 1
PALABRA_CLAVE_CONSTANCIAS = "constancia"

LOGO_PATH = "logo_cetpro.png"

COLUMNAS_ESPERADAS = [
    "ID_CONVENIO", "NRO_CONVENIO", "AÑO", "EMPRESA_INSTITUCION", "RUC",
    "DIRECCION", "REPRESENTANTE", "DNI_REPRESENTANTE", "TELEFONO",
    "PROGRAMA_FORMATIVO", "FECHA_INICIO", "DURACION_AÑOS", "FECHA_FIN",
    "DIAS_RESTANTES", "ESTADO", "CLAVE_BUSQUEDA", "OBSERVACIONES",
]

COLUMNAS_CONSTANCIAS_ESPERADAS = [
    "ID_CONSTANCIA", "DNI", "APELLIDOS", "NOMBRES", "AÑO", "PROGRAMA_ESTUDIOS",
    "MODULO_CONSTANCIA", "ESTADO", "NOMBRE_ARCHIVO_RECOMENDADO", "ENLACE_DRIVE",
    "OBSERVACION", "VALIDACION",
]

# ==================================================================
# 2. ESTILOS (CSS) - INTERFAZ ELEGANTE / PROFESIONAL
# ==================================================================
st.markdown(
    """
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}

        html, body, [class*="css"]  {
            font-family: 'Segoe UI', 'Trebuchet MS', sans-serif;
        }

        html, body {
            background: #e4ecf5 !important;
            margin: 0;
            padding: 0;
        }

        .stApp {
            background: linear-gradient(180deg, #eef3f9 0%, #e4ecf5 100%);
        }

        [data-testid="stHeader"] { display: none; }
        [data-testid="stToolbar"] { display: none; }
        [data-testid="stAppViewContainer"] { padding-top: 0px; }
        .block-container { padding-top: 2.2rem; padding-bottom: 2rem; }

        .stTextInput input {
            background-color: #ffffff !important;
            color: #16273a !important;
            border: 1px solid #c7d4e2 !important;
        }
        .stTextInput label {
            color: #16273a !important;
            font-weight: 600 !important;
        }

        .cetpro-header {
            background: linear-gradient(135deg, #06213a 0%, #0b3d62 55%, #14649e 100%);
            padding: 28px 36px;
            border-radius: 18px;
            box-shadow: 0 10px 25px rgba(6, 33, 58, 0.35);
            margin-bottom: 28px;
            display: flex;
            align-items: center;
            gap: 20px;
        }
        .cetpro-header h1 {
            color: #ffffff;
            font-size: 26px;
            margin: 0;
            font-weight: 800;
            letter-spacing: 0.4px;
        }
        .cetpro-header p {
            color: #eaf4ff;
            margin: 4px 0 0 0;
            font-size: 14px;
            font-weight: 500;
        }

        .login-box {
            max-width: 430px;
            margin: 30px auto 0 auto;
            background: #ffffff;
            padding: 40px 36px 30px 36px;
            border-radius: 20px;
            box-shadow: 0 15px 35px rgba(6, 33, 58, 0.22);
            border: 2px solid #14649e;
        }
        .login-title {
            text-align: center;
            color: #06213a;
            font-weight: 800;
            font-size: 22px;
            margin-bottom: 2px;
        }
        .login-subtitle {
            text-align: center;
            color: #3d5872;
            font-size: 13px;
            margin-bottom: 26px;
            font-weight: 500;
        }

        .result-card {
            background: #ffffff;
            border-radius: 16px;
            padding: 26px 30px;
            margin-bottom: 22px;
            box-shadow: 0 8px 22px rgba(6, 33, 58, 0.14);
            border-left: 6px solid #14649e;
        }
        .result-title {
            font-size: 19px;
            font-weight: 800;
            color: #06213a;
            margin-bottom: 4px;
        }
        .badge {
            display: inline-block;
            padding: 4px 14px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 800;
            letter-spacing: 0.3px;
            margin-bottom: 16px;
        }
        .badge-vigente { background: #0f8a4b; color: #ffffff; }
        .badge-vencido { background: #c1273a; color: #ffffff; }
        .badge-otro    { background: #445b70; color: #ffffff; }

        .result-line {
            font-size: 15px;
            color: #16273a;
            line-height: 2.1;
            margin: 0;
        }
        .result-line b { color: #06213a; }

        .stat-box {
            background: #ffffff;
            border-radius: 14px;
            padding: 16px 20px;
            box-shadow: 0 6px 16px rgba(6, 33, 58, 0.12);
            text-align: center;
            border-top: 4px solid #14649e;
        }
        .stat-number { font-size: 24px; font-weight: 800; color: #06213a; }
        .stat-label {
            font-size: 12px;
            color: #3d5872;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-weight: 700;
        }

        div.stButton > button {
            background: linear-gradient(135deg, #06213a, #14649e);
            color: #ffffff;
            border: none;
            border-radius: 10px;
            padding: 10px 22px;
            font-weight: 700;
            transition: 0.2s;
        }
        div.stButton > button:hover {
            filter: brightness(1.15);
            box-shadow: 0 6px 14px rgba(6, 33, 58, 0.45);
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==================================================================
# 3. FUNCIONES AUXILIARES
# ==================================================================
def normalizar(texto) -> str:
    if texto is None:
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    texto = " ".join(texto.split())
    return texto


def limpiar_numero(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


def _listar_archivos_recursivo(carpeta_raiz="."):
    rutas = []
    for carpeta_actual, subcarpetas, archivos in os.walk(carpeta_raiz):
        subcarpetas[:] = [d for d in subcarpetas if not d.startswith(".")]
        for f in archivos:
            rutas.append(os.path.join(carpeta_actual, f))
    return rutas


def _mtime(ruta: str) -> float:
    """Fecha de última modificación del archivo. Se usa como parte de la
    clave de caché: si el archivo cambia (se sube una versión nueva),
    esta función devuelve un número distinto y Streamlit vuelve a leer
    el Excel automáticamente, sin necesidad de reiniciar la app."""
    try:
        return os.path.getmtime(ruta)
    except Exception:
        return 0.0


@st.cache_data(show_spinner=False)
def encontrar_logo():
    if os.path.exists(LOGO_PATH):
        return LOGO_PATH
    extensiones_validas = (".png", ".jpg", ".jpeg", ".webp")
    try:
        todas_las_rutas = _listar_archivos_recursivo(".")
    except Exception:
        return None
    for ruta in todas_las_rutas:
        if os.path.basename(ruta).lower() == LOGO_PATH.lower():
            return ruta
    for ruta in todas_las_rutas:
        nombre = os.path.basename(ruta).lower()
        if nombre.endswith(extensiones_validas) and ("logo" in nombre or "cetpro" in nombre):
            return ruta
    return None


@st.cache_data(show_spinner=False)
def logo_como_base64(ruta_logo: str, mtime: float):
    try:
        with open(ruta_logo, "rb") as f:
            datos = f.read()
        extension = ruta_logo.lower().rsplit(".", 1)[-1]
        mime = "image/png" if extension == "png" else "image/jpeg"
        return f"data:{mime};base64,{base64.b64encode(datos).decode('utf-8')}"
    except Exception:
        return None


def encontrar_archivos_por_palabra(palabra_clave):
    """Devuelve TODOS los .xlsx cuyo nombre contenga la palabra clave,
    ordenados para intentarse en orden (por si hay duplicados subidos
    por error, la app prueba con el siguiente si el primero falla)."""
    todos_los_archivos = _listar_archivos_recursivo(".")
    candidatos = [
        r for r in todos_los_archivos
        if os.path.basename(r).lower().endswith(".xlsx")
        and palabra_clave.lower() in os.path.basename(r).lower()
    ]
    candidatos.sort()
    return candidatos


# ==================================================================
# 4. CARGA DE DATOS (con recarga automática al detectar cambios)
# ==================================================================
@st.cache_data(show_spinner=False)
def _leer_convenios(ruta_excel: str, mtime: float):
    df = pd.read_excel(ruta_excel, sheet_name=HOJA_DATOS, header=FILA_ENCABEZADO, engine="openpyxl")

    columnas_presentes = [c for c in COLUMNAS_ESPERADAS if c in df.columns]
    df = df[columnas_presentes].copy()
    if "ID_CONVENIO" in df.columns:
        df = df.dropna(subset=["ID_CONVENIO"]).reset_index(drop=True)

    for col in ["RUC", "DNI_REPRESENTANTE"]:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_numero)

    for col in ["FECHA_INICIO", "FECHA_FIN"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    campos_busqueda = ["RUC", "DNI_REPRESENTANTE", "EMPRESA_INSTITUCION", "REPRESENTANTE", "PROGRAMA_FORMATIVO", "NRO_CONVENIO", "ID_CONVENIO"]
    campos_presentes = [c for c in campos_busqueda if c in df.columns]

    def construir_clave(fila):
        return " | ".join(normalizar(fila[c]) for c in campos_presentes)

    df["_BUSQUEDA_NORM"] = df.apply(construir_clave, axis=1)
    return df


def cargar_datos_convenios():
    candidatos = encontrar_archivos_por_palabra(PALABRA_CLAVE_CONVENIOS)
    if not candidatos:
        return None, "No se encontró ningún archivo Excel de convenios en el repositorio (el nombre debe contener la palabra 'convenio')."

    ultimo_error = None
    for ruta_excel in candidatos:
        try:
            df = _leer_convenios(ruta_excel, _mtime(ruta_excel))
            return df, None
        except Exception as e:
            ultimo_error = f"Error al leer '{ruta_excel}': {e}"
            continue
    return None, ultimo_error or "No se pudo leer ningún archivo de convenios válido."


@st.cache_data(show_spinner=False)
def _leer_constancias(ruta_excel: str, mtime: float):
    libro = pd.ExcelFile(ruta_excel, engine="openpyxl")
    hoja = HOJA_CONSTANCIAS if HOJA_CONSTANCIAS in libro.sheet_names else libro.sheet_names[0]
    df = pd.read_excel(ruta_excel, sheet_name=hoja, header=FILA_ENCABEZADO_CONSTANCIAS, engine="openpyxl")

    columnas_presentes = [c for c in COLUMNAS_CONSTANCIAS_ESPERADAS if c in df.columns]
    if len(columnas_presentes) < 4 or "ID_CONSTANCIA" not in columnas_presentes:
        raise ValueError("el archivo no tiene la estructura de columnas esperada para constancias")

    df = df[columnas_presentes].copy()
    df = df.dropna(subset=["ID_CONSTANCIA"]).reset_index(drop=True)

    if "DNI" in df.columns:
        df["DNI"] = df["DNI"].apply(limpiar_numero)

    campos_busqueda = ["DNI", "APELLIDOS", "NOMBRES", "PROGRAMA_ESTUDIOS", "MODULO_CONSTANCIA", "ID_CONSTANCIA", "AÑO"]
    campos_presentes = [c for c in campos_busqueda if c in df.columns]

    def construir_clave(fila):
        return " | ".join(normalizar(fila[c]) for c in campos_presentes)

    df["_BUSQUEDA_NORM"] = df.apply(construir_clave, axis=1)
    return df


def cargar_datos_constancias():
    candidatos = encontrar_archivos_por_palabra(PALABRA_CLAVE_CONSTANCIAS)
    if not candidatos:
        return None, "No se encontró ningún archivo Excel de constancias en el repositorio (el nombre debe contener la palabra 'constancia')."

    ultimo_error = None
    for ruta_excel in candidatos:
        try:
            df = _leer_constancias(ruta_excel, _mtime(ruta_excel))
            return df, None
        except Exception as e:
            ultimo_error = f"Error al leer '{ruta_excel}': {e}"
            continue
    return None, ultimo_error or "No se pudo leer ningún archivo de constancias válido."


def buscar_en_df(df: pd.DataFrame, consulta: str, tipo: str) -> pd.DataFrame:
    consulta_norm = normalizar(consulta)
    if not consulta_norm:
        return df.iloc[0:0]
    consulta_limpia = consulta.strip()

    if tipo == "convenio":
        if "RUC" in df.columns:
            exacto_ruc = df[df["RUC"] == consulta_limpia]
            if not exacto_ruc.empty:
                return exacto_ruc
        if "DNI_REPRESENTANTE" in df.columns:
            exacto_dni = df[df["DNI_REPRESENTANTE"] == consulta_limpia]
            if not exacto_dni.empty:
                return exacto_dni
    else:
        if "DNI" in df.columns:
            exacto_dni = df[df["DNI"] == consulta_limpia]
            if not exacto_dni.empty:
                return exacto_dni

    return df[df["_BUSQUEDA_NORM"].str.contains(consulta_norm, na=False)]


def badge_estado(estado: str) -> str:
    estado_norm = normalizar(estado)
    if estado_norm in ["VIGENTE", "EMITIDO", "ACTIVO", "OK"]:
        return f'<span class="badge badge-vigente">✅ {estado}</span>'
    if "VENC" in estado_norm or "ANUL" in estado_norm:
        return f'<span class="badge badge-vencido">⛔ {estado}</span>'
    return f'<span class="badge badge-otro">ℹ️ {estado if estado else "REGISTRADO"}</span>'


def formato_fecha(valor) -> str:
    if pd.isna(valor):
        return "No registrada"
    try:
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except Exception:
        return str(valor)


def texto_o_guion(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    texto = str(valor).strip()
    return texto if texto else "—"


# ==================================================================
# 5. ESTADO DE SESIÓN Y LOGIN
# ==================================================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False


def cerrar_sesion():
    st.session_state.autenticado = False


def mostrar_login():
    col_izq, col_centro, col_der = st.columns([1, 1.3, 1])
    with col_centro:
        logo_encontrado = encontrar_logo()
        logo_html = ""
        if logo_encontrado:
            logo_data = logo_como_base64(logo_encontrado, _mtime(logo_encontrado))
            if logo_data:
                logo_html = (
                    f'<img src="{logo_data}" '
                    f'style="width:90px; height:auto; display:block; '
                    f'margin:0 auto 14px auto; border-radius:12px;">'
                )

        st.markdown(
            f"""
            <div class="login-box">
                {logo_html}
                <p class="login-title">🎓 CETPRO CAJAMARCA</p>
                <p class="login-subtitle">Sistema de gestión de convenios y constancias</p>
            """,
            unsafe_allow_html=True,
        )

        with st.form("form_login", clear_on_submit=False):
            usuario = st.text_input("👤 Usuario", placeholder="Ingrese su usuario")
            contrasena = st.text_input("🔒 Contraseña", type="password", placeholder="Ingrese su contraseña")
            enviar = st.form_submit_button("Ingresar al sistema")

            if enviar:
                if usuario.strip() == USUARIO_VALIDO and contrasena == CONTRASENA_VALIDA:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")

        st.markdown("</div>", unsafe_allow_html=True)


# ==================================================================
# 6. PANTALLA PRINCIPAL (CON PESTAÑAS: CONVENIOS / CONSTANCIAS)
# ==================================================================
def mostrar_buscador():
    col_logo, col_titulo, col_salir = st.columns([0.15, 0.7, 0.15])
    with col_logo:
        logo_encontrado = encontrar_logo()
        if logo_encontrado:
            st.image(logo_encontrado, width=70)
    with col_titulo:
        st.markdown(
            """
            <div class="cetpro-header" style="margin-top:0;">
                <div>
                    <h1>🎓 Sistema Institucional - CETPRO Cajamarca</h1>
                    <p>Módulo de consulta centralizada de convenios y constancias EFSRT</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_salir:
        st.write("")
        st.button("🚪 Cerrar sesión", on_click=cerrar_sesion)

    pestana_convenios, pestana_constancias = st.tabs(["📁 Convenios", "📜 Constancias EFSRT"])

    # ---------------- MÓDULO DE CONVENIOS ----------------
    with pestana_convenios:
        df_conv, error_conv = cargar_datos_convenios()
        if error_conv:
            st.error(f"⚠️ {error_conv}")
            st.info("Verifica que el archivo Excel de convenios esté en la misma carpeta que app.py, dentro del repositorio.")
        else:
            total = len(df_conv)
            vigentes = len(df_conv[df_conv["ESTADO"].apply(normalizar) == "VIGENTE"]) if "ESTADO" in df_conv.columns else 0

            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="stat-box"><div class="stat-number">{total}</div><div class="stat-label">📁 Convenios registrados</div></div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="stat-box"><div class="stat-number">{vigentes}</div><div class="stat-label">✅ Convenios vigentes</div></div>', unsafe_allow_html=True)
            with c3:
                hoy = datetime.now().strftime("%d/%m/%Y")
                st.markdown(f'<div class="stat-box"><div class="stat-number">{hoy}</div><div class="stat-label">📅 Fecha de consulta</div></div>', unsafe_allow_html=True)

            st.write("")
            col_b1, col_b2 = st.columns([4, 1])
            with col_b1:
                consulta_conv = st.text_input(
                    "🔎 Buscar por DNI del representante, RUC, empresa, representante o programa",
                    placeholder="Ej: 26641673  |  20510752938  |  MANTENIMIENTO INDUSTRIAL",
                    key="input_conv",
                )
            with col_b2:
                st.write("")
                st.write("")
                btn_conv = st.button("Buscar 🔍", key="btn_conv_click")

            if consulta_conv or btn_conv:
                if not consulta_conv.strip():
                    st.warning("Ingresa un DNI, RUC, nombre o palabra clave para buscar.")
                else:
                    resultados_conv = buscar_en_df(df_conv, consulta_conv, "convenio")
                    if resultados_conv.empty:
                        st.markdown('<div class="result-card" style="border-left-color:#b02a37;"><p class="result-line">🚫 <b>No se encontraron convenios</b> que coincidan con la búsqueda.</p></div>', unsafe_allow_html=True)
                    else:
                        st.success(f"Se encontraron {len(resultados_conv)} convenio(s) coincidentes.")
                        for _, fila in resultados_conv.iterrows():
                            empresa = texto_o_guion(fila.get("EMPRESA_INSTITUCION"))
                            estado = texto_o_guion(fila.get("ESTADO"))
                            dias_restantes = fila.get("DIAS_RESTANTES")
                            dias_texto = (
                                f"{int(dias_restantes)} días"
                                if pd.notna(dias_restantes) and str(dias_restantes).strip() != ""
                                else "No calculado"
                            )

                            html = f"""
                            <div class="result-card">
                                <div class="result-title">🏢 {empresa}</div>
                                {badge_estado(estado)}
                                <p class="result-line">🆔 <b>N° de Convenio:</b> {texto_o_guion(fila.get('NRO_CONVENIO'))} &nbsp;
                                (ID: {texto_o_guion(fila.get('ID_CONVENIO'))})</p>
                                <p class="result-line">📇 <b>RUC:</b> {texto_o_guion(fila.get('RUC'))}</p>
                                <p class="result-line">📍 <b>Dirección:</b> {texto_o_guion(fila.get('DIRECCION'))}</p>
                                <p class="result-line">👤 <b>Representante:</b> {texto_o_guion(fila.get('REPRESENTANTE'))}</p>
                                <p class="result-line">🪪 <b>DNI del representante:</b> {texto_o_guion(fila.get('DNI_REPRESENTANTE'))}</p>
                                <p class="result-line">📞 <b>Teléfono:</b> {texto_o_guion(fila.get('TELEFONO'))}</p>
                                <p class="result-line">📚 <b>Programa formativo:</b> {texto_o_guion(fila.get('PROGRAMA_FORMATIVO'))}</p>
                                <p class="result-line">📅 <b>Fecha de inicio:</b> {formato_fecha(fila.get('FECHA_INICIO'))}</p>
                                <p class="result-line">⏳ <b>Duración:</b> {texto_o_guion(fila.get('DURACION_AÑOS'))} año(s)</p>
                                <p class="result-line">🏁 <b>Fecha de finalización:</b> {formato_fecha(fila.get('FECHA_FIN'))}</p>
                                <p class="result-line">⌛ <b>Días restantes:</b> {dias_texto}</p>
                                <p class="result-line">📝 <b>Observaciones:</b> {texto_o_guion(fila.get('OBSERVACIONES'))}</p>
                            </div>
                            """
                            st.markdown(html, unsafe_allow_html=True)

    # ---------------- MÓDULO DE CONSTANCIAS EFSRT ----------------
    with pestana_constancias:
        df_const, error_const = cargar_datos_constancias()
        if error_const:
            st.error(f"⚠️ {error_const}")
            st.info("Verifica que el archivo Excel de constancias esté en la misma carpeta que app.py, dentro del repositorio.")
        else:
            total_c = len(df_const)
            ok_c = len(df_const[df_const["ESTADO"].apply(normalizar) == "OK"]) if "ESTADO" in df_const.columns else 0
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="stat-box"><div class="stat-number">{total_c}</div><div class="stat-label">📜 Constancias registradas</div></div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="stat-box"><div class="stat-number">{ok_c}</div><div class="stat-label">✅ Constancias OK</div></div>', unsafe_allow_html=True)
            with c3:
                hoy = datetime.now().strftime("%d/%m/%Y")
                st.markdown(f'<div class="stat-box"><div class="stat-number">{hoy}</div><div class="stat-label">📅 Fecha de consulta</div></div>', unsafe_allow_html=True)

            st.write("")
            col_cb1, col_cb2 = st.columns([4, 1])
            with col_cb1:
                consulta_const = st.text_input(
                    "🔎 Buscar por DNI, apellidos, nombres, programa o módulo",
                    placeholder="Ej: 26628709  |  MALCA RUITON  |  GESTION Y ADMINISTRACION",
                    key="input_const",
                )
            with col_cb2:
                st.write("")
                st.write("")
                btn_const = st.button("Buscar 🔍", key="btn_const_click")

            if consulta_const or btn_const:
                if not consulta_const.strip():
                    st.warning("Ingresa un DNI, apellido, nombre o palabra clave para buscar.")
                else:
                    resultados_const = buscar_en_df(df_const, consulta_const, "constancia")
                    if resultados_const.empty:
                        st.markdown('<div class="result-card" style="border-left-color:#b02a37;"><p class="result-line">🚫 <b>No se encontraron constancias</b> que coincidan con la búsqueda.</p></div>', unsafe_allow_html=True)
                    else:
                        st.success(f"Se encontraron {len(resultados_const)} constancia(s) coincidentes.")
                        for _, fila in resultados_const.iterrows():
                            apellidos = texto_o_guion(fila.get("APELLIDOS"))
                            nombres = texto_o_guion(fila.get("NOMBRES"))
                            estado = texto_o_guion(fila.get("ESTADO"))
                            enlace_drive = fila.get("ENLACE_DRIVE")

                            enlace_html = ""
                            if pd.notna(enlace_drive) and str(enlace_drive).strip() != "":
                                enlace_html = (
                                    f'<p class="result-line">🔗 <b>Enlace de Drive:</b> '
                                    f'<a href="{enlace_drive}" target="_blank">Abrir Constancia ↗</a></p>'
                                )

                            html = f"""
                            <div class="result-card">
                                <div class="result-title">👤 {apellidos}, {nombres}</div>
                                {badge_estado(estado)}
                                <p class="result-line">📜 <b>ID / Código:</b> {texto_o_guion(fila.get('ID_CONSTANCIA'))} &nbsp;
                                (Año: {texto_o_guion(fila.get('AÑO'))})</p>
                                <p class="result-line">🪪 <b>DNI:</b> {texto_o_guion(fila.get('DNI'))}</p>
                                <p class="result-line">📚 <b>Programa de Estudios:</b> {texto_o_guion(fila.get('PROGRAMA_ESTUDIOS'))}</p>
                                <p class="result-line">⚙️ <b>Módulo:</b> {texto_o_guion(fila.get('MODULO_CONSTANCIA'))}</p>
                                <p class="result-line">📁 <b>Archivo recomendado:</b> {texto_o_guion(fila.get('NOMBRE_ARCHIVO_RECOMENDADO'))}</p>
                                {enlace_html}
                                <p class="result-line">📝 <b>Observación:</b> {texto_o_guion(fila.get('OBSERVACION'))}</p>
                            </div>
                            """
                            st.markdown(html, unsafe_allow_html=True)


# ==================================================================
# 7. ENRUTAMIENTO PRINCIPAL
# ==================================================================
if not st.session_state.autenticado:
    mostrar_login()
else:
    mostrar_buscador()
