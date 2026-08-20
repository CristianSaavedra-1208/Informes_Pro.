import streamlit as st
from src.core.security_engine import authenticate_user

def render_login():
    """Renderiza la pantalla corporativa de inicio de sesión de Informes Pro."""
    
    # CSS personalizado para centrar el card de Login
    st.markdown("""
        <style>
        .login-card {
            max-width: 480px;
            margin: 40px auto;
            padding: 30px 35px;
            background-color: #FFFFFF;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.08);
            border: 1px solid #E2E8F0;
        }
        .login-header {
            text-align: center;
            margin-bottom: 25px;
        }
        .login-title {
            font-size: 26px;
            font-weight: 700;
            color: #1F4E78;
            margin-bottom: 6px;
        }
        .login-subtitle {
            font-size: 13px;
            color: #64748B;
        }
        </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.write("")
        st.write("")
        
        st.markdown("""
            <div class="login-header">
                <div style="font-size: 42px; margin-bottom: 8px;">📊</div>
                <div class="login-title">Informes Pro</div>
                <div class="login-subtitle">Sistema de Emisión de Estados Financieros bajo IFRS</div>
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("form_login", clear_on_submit=False):
            st.subheader("🔐 Inicio de Sesión")
            
            username = st.text_input("Usuario:", placeholder="Ej: admin", key="login_username_input").strip()
            password = st.text_input("Contraseña:", type="password", placeholder="••••••••", key="login_password_input")
            
            submit = st.form_submit_button("🔑 Iniciar Sesión", type="primary", use_container_width=True)
            
            if submit:
                if not username or not password:
                    st.error("⚠️ Por favor ingresa tu usuario y contraseña.")
                else:
                    res = authenticate_user(username, password)
                    if res == "INACTIVO":
                        st.error("🚫 Tu cuenta de usuario se encuentra deshabilitada. Contacta al Administrador.")
                    elif isinstance(res, dict):
                        st.session_state['is_authenticated'] = True
                        st.session_state['auth_user'] = res['usuario']
                        st.session_state['auth_role'] = res['rol']
                        st.session_state['auth_name'] = res['nombre_completo']
                        st.success(f"✅ ¡Bienvenido(a), {res['nombre_completo']}!")
                        st.rerun()
                    else:
                        st.error("❌ Usuario o contraseña incorrectos. Verifica tus credenciales.")
                        
        st.info("💡 **Cuentas por defecto:** Admin: `admin` (`admin123`) | Contable: `analista_contable` (`contable123`) | Reportes: `analista_reportes` (`reportes123`) | Auditor: `auditor_lector` (`auditor123`)")
