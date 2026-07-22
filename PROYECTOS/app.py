# -*- coding: utf-8 -*-
"""
==================================================================
 BUSCADOR DE CONVENIOS - CETPRO CAJAMARCA
==================================================================
Aplicación 100% local (no realiza consultas a internet).
Busca convenios por DNI del representante, RUC, nombre de empresa
o representante, dentro del archivo Excel de la base de datos.

Estructura de carpeta requerida (todo en el mismo repositorio):
    app.py
    requirements.txt
    logo_cetpro.png                              (opcional, tu logo)
    BD_Convenios_CETPRO_CORREGIDA_FINAL_V1.xlsx   (tu base de datos)

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
# Esto evita errores si app.py está dentro de una subcarpeta del repo
# (ej. /PROYECTOS/app.py) en vez de estar en la raíz. Todos los archivos
# relacionados (excel, logo) deben estar en ESTA MISMA carpeta.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

# ==================================================================
# 1. CONFIGURACIÓN GENERAL
# ==================================================================
st.set_page_config(
    page_title="CETPRO Cajamarca | Buscador de Convenios",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

USUARIO_VALIDO = "CETPRO CAJAMARCA"
CONTRASENA_VALIDA = "cetprocajamarca#2026"

ARCHIVO_EXCEL = "BD_Convenios_CETPRO_CORREGIDA_FINAL_V1.xlsx"
HOJA_DATOS = "BD_CONVENIOS"
FILA_ENCABEZADO = 3  # fila 4 del Excel (índice 0 = fila 1) contiene los encabezados
LOGO_PATH = "logo_cetpro.png"

COLUMNAS_ESPERADAS = [
    "ID_CONVENIO", "NRO_CONVENIO", "AÑO", "EMPRESA_INSTITUCION", "RUC",
    "DIRECCION", "REPRESENTANTE", "DNI_REPRESENTANTE", "TELEFONO",
    "PROGRAMA_FORMATIVO", "FECHA_INICIO", "DURACION_AÑOS", "FECHA_FIN",
    "DIAS_RESTANTES", "ESTADO", "CLAVE_BUSQUEDA", "OBSERVACIONES",
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

        /* Elimina por completo la barra superior nativa de Streamlit
           (display:none la saca del flujo, así no tapa el título) */
        [data-testid="stHeader"] {
            display: none;
        }
        [data-testid="stToolbar"] {
            display: none;
        }
        [data-testid="stAppViewContainer"] {
            padding-top: 0px;
        }
        .block-container {
            padding-top: 2.2rem;
            padding-bottom: 2rem;
        }

        /* Fuerza que los campos de texto (usuario, contraseña, búsqueda)
           siempre se vean con fondo claro, sin importar el tema del navegador */
        .stTextInput input {
            background-color: #ffffff !important;
            color: #16273a !important;
            border: 1px solid #c7d4e2 !important;
        }
        .stTextInput label {
            color: #16273a !important;
            font-weight: 600 !important;
        }

        /* ---------- Encabezado institucional ---------- */
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

        /* ---------- Tarjeta de login ---------- */
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

        /* ---------- Tarjeta de resultado (vertical / párrafo) ---------- */
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
        .result-line b {
            color: #06213a;
        }

        .stat-box {
            background: #ffffff;
            border-radius: 14px;
            padding: 16px 20px;
            box-shadow: 0 6px 16px rgba(6, 33, 58, 0.12);
            text-align: center;
            border-top: 4px solid #14649e;
        }
        .stat-number {
            font-size: 24px;
            font-weight: 800;
            color: #06213a;
        }
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
    """Pasa a mayúsculas, quita tildes y espacios extra para comparar texto."""
    if texto is None:
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    texto = " ".join(texto.split())
    return texto


def limpiar_numero(valor) -> str:
    """Convierte valores tipo 26641673.0 -> '26641673' y quita espacios."""
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


@st.cache_data(show_spinner=False)
def encontrar_logo():
    """Busca el logo de forma flexible: nombre exacto, o cualquier imagen
    cuyo nombre contenga 'logo' o 'cetpro', con extensión .png/.jpg/.jpeg,
    revisando también dentro de subcarpetas."""
    if os.path.exists(LOGO_PATH):
        return LOGO_PATH

    extensiones_validas = (".png", ".jpg", ".jpeg", ".webp")
    try:
        todas_las_rutas = _listar_archivos_recursivo(".")
    except Exception:
        return None

    # 1) Coincidencia exacta ignorando mayúsculas/minúsculas
    for ruta in todas_las_rutas:
        if os.path.basename(ruta).lower() == LOGO_PATH.lower():
            return ruta

    # 2) Cualquier imagen que contenga "logo" o "cetpro" en el nombre
    for ruta in todas_las_rutas:
        nombre = os.path.basename(ruta).lower()
        if nombre.endswith(extensiones_validas) and ("logo" in nombre or "cetpro" in nombre):
            return ruta

    return None


@st.cache_data(show_spinner=False)
def logo_como_base64(ruta_logo: str):
    """Convierte el logo a base64 para poder incrustarlo dentro de la
    tarjeta HTML de login/encabezado (así queda realmente 'dentro' del
    marco blanco y no como una imagen suelta encima)."""
    try:
        with open(ruta_logo, "rb") as f:
            datos = f.read()
        extension = ruta_logo.lower().rsplit(".", 1)[-1]
        mime = "image/png" if extension == "png" else "image/jpeg"
        return f"data:{mime};base64,{base64.b64encode(datos).decode('utf-8')}"
    except Exception:
        return None


def _listar_archivos_recursivo(carpeta_raiz="."):
    """Recorre la carpeta raíz y todas sus subcarpetas (ignorando .git,
    .streamlit, .devcontainer y otras carpetas ocultas) y devuelve
    la lista de rutas de archivo encontradas."""
    rutas = []
    for carpeta_actual, subcarpetas, archivos in os.walk(carpeta_raiz):
        subcarpetas[:] = [d for d in subcarpetas if not d.startswith(".")]
        for f in archivos:
            rutas.append(os.path.join(carpeta_actual, f))
    return rutas


def encontrar_archivo_excel():
    """Busca el Excel de forma tolerante a mayúsculas/minúsculas y a la
    ubicación exacta: revisa la carpeta de app.py y también sus subcarpetas,
    por si el archivo quedó dentro de una carpeta distinta al subirlo a GitHub."""
    if os.path.exists(ARCHIVO_EXCEL):
        return ARCHIVO_EXCEL

    todos_los_archivos = _listar_archivos_recursivo(".")

    # 1) Coincidencia exacta de nombre, ignorando mayúsculas/minúsculas
    for ruta in todos_los_archivos:
        if os.path.basename(ruta).lower() == ARCHIVO_EXCEL.lower():
            return ruta

    # 2) Cualquier archivo .xlsx que contenga "convenio" en el nombre
    for ruta in todos_los_archivos:
        nombre = os.path.basename(ruta).lower()
        if nombre.endswith(".xlsx") and "convenio" in nombre:
            return ruta

    # 3) Como último recurso, el primer .xlsx que exista en cualquier subcarpeta
    xlsx_disponibles = [r for r in todos_los_archivos if r.lower().endswith(".xlsx")]
    if xlsx_disponibles:
        return xlsx_disponibles[0]

    return None


@st.cache_data(show_spinner=False)
def cargar_datos():
    """Carga y limpia la base de datos desde el Excel local (sin internet)."""
    ruta_excel = encontrar_archivo_excel()

    if ruta_excel is None:
        archivos_presentes = ", ".join(_listar_archivos_recursivo(".")) or "(la carpeta está vacía)"
        return None, (
            f"No se encontró ningún archivo Excel llamado '{ARCHIVO_EXCEL}' en el repositorio. "
            f"Archivos presentes en la carpeta: {archivos_presentes}"
        )

    try:
        df = pd.read_excel(
            ruta_excel,
            sheet_name=HOJA_DATOS,
            header=FILA_ENCABEZADO,
            engine="openpyxl",
        )
    except Exception as e:
        return None, f"Error al leer el archivo '{ruta_excel}': {e}"

    # Nos quedamos solo con las columnas esperadas que existan
    columnas_presentes = [c for c in COLUMNAS_ESPERADAS if c in df.columns]
    df = df[columnas_presentes].copy()

    # Eliminar filas vacías (sin ID_CONVENIO)
    if "ID_CONVENIO" in df.columns:
        df = df.dropna(subset=["ID_CONVENIO"]).reset_index(drop=True)

    # Normalizar RUC y DNI a texto limpio
    for col in ["RUC", "DNI_REPRESENTANTE"]:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_numero)

    # Normalizar fechas
    for col in ["FECHA_INICIO", "FECHA_FIN"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Columna auxiliar de búsqueda (normalizada, sin tildes)
    campos_busqueda = [
        "RUC", "DNI_REPRESENTANTE", "EMPRESA_INSTITUCION", "REPRESENTANTE",
        "PROGRAMA_FORMATIVO", "NRO_CONVENIO", "ID_CONVENIO",
    ]
    campos_presentes = [c for c in campos_busqueda if c in df.columns]

    def construir_clave(fila):
        partes = [normalizar(fila[c]) for c in campos_presentes]
        return " | ".join(partes)

    df["_BUSQUEDA_NORM"] = df.apply(construir_clave, axis=1)

    return df, None


def buscar(df: pd.DataFrame, consulta: str) -> pd.DataFrame:
    """Búsqueda exacta por RUC/DNI y, si no hay coincidencia exacta,
    búsqueda parcial por nombre/empresa/programa."""
    consulta_norm = normalizar(consulta)
    if not consulta_norm:
        return df.iloc[0:0]

    consulta_limpia = consulta.strip()

    # 1) Coincidencia exacta por RUC
    if "RUC" in df.columns:
        exacto_ruc = df[df["RUC"] == consulta_limpia]
        if not exacto_ruc.empty:
            return exacto_ruc

    # 2) Coincidencia exacta por DNI del representante
    if "DNI_REPRESENTANTE" in df.columns:
        exacto_dni = df[df["DNI_REPRESENTANTE"] == consulta_limpia]
        if not exacto_dni.empty:
            return exacto_dni

    # 3) Búsqueda parcial (nombre de empresa, representante, programa, N° convenio, etc.)
    contiene = df[df["_BUSQUEDA_NORM"].str.contains(consulta_norm, na=False)]
    return contiene


def badge_estado(estado: str) -> str:
    estado_norm = normalizar(estado)
    if estado_norm == "VIGENTE":
        return f'<span class="badge badge-vigente">✅ {estado}</span>'
    if "VENC" in estado_norm:
        return f'<span class="badge badge-vencido">⛔ {estado}</span>'
    return f'<span class="badge badge-otro">ℹ️ {estado if estado else "SIN ESTADO"}</span>'


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
# 4. ESTADO DE SESIÓN
# ==================================================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False


def cerrar_sesion():
    st.session_state.autenticado = False


# ==================================================================
# 5. PANTALLA DE LOGIN
# ==================================================================
def mostrar_login():
    col_izq, col_centro, col_der = st.columns([1, 1.3, 1])
    with col_centro:
        logo_encontrado = encontrar_logo()
        logo_html = ""
        if logo_encontrado:
            logo_data = logo_como_base64(logo_encontrado)
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
                <p class="login-subtitle">Sistema de búsqueda de convenios institucionales</p>
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
# 6. PANTALLA PRINCIPAL (BUSCADOR)
# ==================================================================
def mostrar_buscador():
    # ---- Encabezado ----
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
                    <h1>🎓 Buscador de Convenios - CETPRO Cajamarca</h1>
                    <p>Consulta institucional por DNI del representante, RUC, empresa o programa formativo</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_salir:
        st.write("")
        st.button("🚪 Cerrar sesión", on_click=cerrar_sesion)

    # ---- Cargar datos ----
    df, error = cargar_datos()
    if error:
        st.error(f"⚠️ {error}")
        st.info("Verifica que el archivo Excel esté en la misma carpeta que app.py, dentro del repositorio.")
        return

    # ---- Estadísticas rápidas ----
    total = len(df)
    vigentes = len(df[df["ESTADO"].apply(normalizar) == "VIGENTE"]) if "ESTADO" in df.columns else 0
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{total}</div>'
            f'<div class="stat-label">📁 Convenios registrados</div></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{vigentes}</div>'
            f'<div class="stat-label">✅ Convenios vigentes</div></div>',
            unsafe_allow_html=True,
        )
    with c3:
        hoy = datetime.now().strftime("%d/%m/%Y")
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{hoy}</div>'
            f'<div class="stat-label">📅 Fecha de consulta</div></div>',
            unsafe_allow_html=True,
        )

    st.write("")

    # ---- Buscador ----
    col_buscar, col_boton = st.columns([4, 1])
    with col_buscar:
        consulta = st.text_input(
            "🔎 Buscar por DNI del representante, RUC, empresa, representante o programa",
            placeholder="Ej: 26641673  |  20510752938  |  MANTENIMIENTO INDUSTRIAL  |  ELECTRICIDAD",
        )
    with col_boton:
        st.write("")
        st.write("")
        buscar_click = st.button("Buscar 🔍")

    if consulta or buscar_click:
        if not consulta.strip():
            st.warning("Ingresa un DNI, RUC, nombre o palabra clave para buscar.")
            return

        resultados = buscar(df, consulta)

        if resultados.empty:
            st.markdown(
                """
                <div class="result-card" style="border-left-color:#b02a37;">
                    <p class="result-line">🚫 <b>No se encontraron convenios</b> que coincidan con la búsqueda.
                    Verifica el DNI, RUC o nombre ingresado.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            return

        st.success(f"Se encontraron {len(resultados)} convenio(s) coincidentes.")

        # ---- Mostrar cada resultado como tarjeta vertical tipo párrafo ----
        for _, fila in resultados.iterrows():
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


# ==================================================================
# 7. ENRUTAMIENTO PRINCIPAL
# ==================================================================
if not st.session_state.autenticado:
    mostrar_login()
else:
    mostrar_buscador()
