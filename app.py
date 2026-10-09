import streamlit as st

st.set_page_config(
    page_title="Facturación - Grupo SERVYRE", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

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
        st.sidebar.title("🔐 Iniciar Sesión")
        usuario_ingresado = st.sidebar.text_input("Usuario")
        password_ingresada = st.sidebar.text_input("Contraseña", type="password")
        
        if st.sidebar.button("Entrar"):
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
    st.warning("⚠️ Por favor, inicia sesión en la barra lateral para acceder al sistema.")
    st.stop()

# Si ya inició sesión, mostramos los datos del usuario en la barra lateral
st.sidebar.markdown(f"👤 **Usuario:** {st.session_state.usuario} ({st.session_state.rol})")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.autenticado = False
    st.session_state.usuario = None
    st.session_state.rol = None
    st.rerun()

st.sidebar.markdown("---")

# --- BOTÓN PARA BORRAR LA CACHÉ DE TODOS LOS MENÚS ---
if st.sidebar.button("🧹 Borrar Caché General"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.sidebar.success("¡Caché borrada exitosamente!")
    st.rerun()

st.sidebar.markdown("---")

# --- NAVEGACIÓN MODERNA CON ICONOS NATIVOS ---
# NOTA: Asegúrate de que los nombres de los archivos correspondan exactamente 
# con los nombres reales de tus scripts dentro de tu proyecto.
paginas = {
    "Menú Principal": [
        st.Page("1_Finanzas.py", title="Finanzas", icon="💰"),
        st.Page("2_Contabilidad.py", title="Contabilidad", icon="📊"),
        st.Page("Amarres.py", title="Amarres", icon="🔗"),
        st.Page("SAT.py", title="SAT", icon="🏛️"),
    ]
}

navegacion = st.navigation(paginas)
navegacion.run()