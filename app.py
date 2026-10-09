import streamlit as st

st.set_page_config(
    page_title="Grupo SERVYRE - Portal Financiero", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS PERSONALIZADOS (ESTILO CORPORATIVO MODERNO) ---
st.markdown("""
    <style>
        /* Ocultar el menú superior predeterminado de Streamlit y el pie de página */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* Estilo general del fondo de la aplicación */
        .stApp {
            background-color: #F8F9FA;
        }

        /* Estilo de la barra lateral (Sidebar) con tono azul corporativo */
        [data-testid="stSidebar"] {
            background-color: #0078D7;
            color: white;
        }

        /* Cambiar color de textos e iconos en la barra lateral */
        [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stRadio div {
            color: white !important;
        }

        /* Estilizar los botones principales de la app */
        .stButton button {
            background-color: #0078D7;
            color: white;
            border-radius: 6px;
            border: none;
            font-weight: 600;
            padding: 0.5rem 1rem;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .stButton button:hover {
            background-color: #005a9e;
            color: white;
        }

        /* Tarjetas de métricas modernas */
        [data-testid="stMetric"] {
            background-color: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            border: 1px solid #E1DFDD;
        }
    </style>
""", unsafe_allow_html=True)

# --- SISTEMA DE CREDENCIALES Y ROLES ---
USUARIOS = {
    "admin": {"password": "123", "rol": "Administrador"},
    "contador": {"password": "456", "rol": "Contabilidad"},
    "finanzas": {"password": "789", "rol": "Finanzas"}
}

def verificar_login():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False
        st.session_state.usuario = None
        st.session_state.rol = None

    if not st.session_state.autenticado:
        st.sidebar.markdown("<h2 style='color: white;'>🔐 Iniciar Sesión</h2>", unsafe_allow_html=True)
        usuario_ingresado = st.sidebar.text_input("Usuario", key="input_user")
        password_ingresada = st.sidebar.text_input("Contraseña", type="password", key="input_pass")
        
        if st.sidebar.button("Entrar al Sistema", use_container_width=True):
            if usuario_ingresado in USUARIOS and USUARIOS[usuario_ingresado]["password"] == password_ingresada:
                st.session_state.autenticado = True
                st.session_state.usuario = usuario_ingresado
                st.session_state.rol = USUARIOS[usuario_ingresado]["rol"]
                st.rerun()
            else:
                st.sidebar.error("Usuario o contraseña incorrectos")
        return False
    return True

# Control de ejecución del login: si no ha iniciado sesión, detenemos la app aquí
if not verificar_login():
    st.markdown("""
        <div style='text-align: center; margin-top: 100px;'>
            <h1 style='color: #0078D7;'>🏢 Grupo SERVYRE</h1>
            <p style='color: #605E5C; font-size: 18px;'>Por favor, ingresa tus credenciales en el panel izquierdo para acceder al portal corporativo.</p>
        </div>
    """, unsafe_allow_html=True)
    st.stop()

# Si ya inició sesión, mostramos el perfil del usuario en la barra lateral
st.sidebar.markdown(f"<h3 style='color: white; font-size: 16px;'>👤 Sesión Activa</h3>", unsafe_allow_html=True)
st.sidebar.markdown(f"<p style='color: #E1DFDD;'><b>Usuario:</b> {st.session_state.usuario}<br><b>Rol:</b> {st.session_state.rol}</p>", unsafe_allow_html=True)

if st.sidebar.button("Cerrar Sesión", use_container_width=True):
    st.session_state.autenticado = False
    st.session_state.usuario = None
    st.session_state.rol = None
    st.rerun()

st.sidebar.markdown("---")

# --- BOTÓN PARA BORRAR LA CACHÉ GENERAL ---
if st.sidebar.button("🧹 Borrar Caché General", use_container_width=True):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.sidebar.success("¡Caché borrada exitosamente!")
    st.rerun()

st.sidebar.markdown("---")

# --- NAVEGACIÓN MODERNA CON ICONOS NATIVOS ---
paginas = {
    "Módulos del Sistema": [
        st.Page("1_Finanzas.py", title="Finanzas & Cobranza", icon="💰"),
        st.Page("2_Contabilidad.py", title="Módulo Contable", icon="📊"),
        st.Page("Amarres.py", title="Amarres y Bancos", icon="🔗"),
        st.Page("SAT.py", title="Módulo SAT", icon="🏛️"),
    ]
}

navegacion = st.navigation(paginas)
navegacion.run()