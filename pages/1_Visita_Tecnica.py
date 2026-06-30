import streamlit as st
import pandas as pd
from drive_utils import SharePointAppClient
from docx_visita import llenar_visita


st.set_page_config(
    page_title="Visita Técnica",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="auto",
)

# ── CSS mobile-friendly (mismo del app principal) ───────────────────────────
st.markdown("""
<style>
.stTextInput input, .stTextArea textarea {
    font-size: 16px !important;
    min-height: 44px !important;
}
.stButton button, .stFormSubmitButton button, .stDownloadButton button {
    min-height: 44px !important;
    font-size: 15px !important;
}
@media (max-width: 768px) {
    .main .block-container {
        padding: 0.75rem 0.5rem !important;
        max-width: 100% !important;
    }
    h1, h2 { font-size: 1.3rem !important; }
    h3 { font-size: 1.1rem !important; }
}
</style>
""", unsafe_allow_html=True)


TEMPLATE_PATH = "FormatoVisitaTecnica/Formato para Visita técnica.docx"


# ── Cliente del Excel (mismo patrón que el app principal) ───────────────────
@st.cache_resource
def _get_client_excel() -> SharePointAppClient:
    return SharePointAppClient(
        client_id     = st.secrets["CLIENT_ID"],
        client_secret = st.secrets["CLIENT_SECRET"],
    )

@st.cache_data(ttl=300)
def _cargar_excel(item_id: str) -> pd.DataFrame:
    return _get_client_excel().leer_excel(
        st.secrets["EXCEL_USER"], item_id, "Nexus"
    )


# ── Encabezado ──────────────────────────────────────────────────────────────
st.markdown(
    """
    <div style="background:#1a2a3a;padding:18px 24px;border-radius:8px;margin-bottom:18px">
        <h2 style="color:white;margin:0;font-family:Segoe UI,sans-serif">
            📋 Visita Técnica
        </h2>
    </div>
    """,
    unsafe_allow_html=True,
)

st.caption("Ingresa el RUT del asegurado y su teléfono. El resto se completa automáticamente desde el Excel.")


# ── Formulario ──────────────────────────────────────────────────────────────
with st.form("form_visita"):
    col1, col2 = st.columns(2)
    rut = col1.text_input("RUT del asegurado", placeholder="ej: 6817145-8")
    telefono = col2.text_input("Teléfono", placeholder="ej: +56 9 1234 5678")
    submit = st.form_submit_button("Generar documento", type="primary", use_container_width=True)


if submit:
    if not rut.strip():
        st.error("Por favor ingresa el RUT.")
    elif not telefono.strip():
        st.error("Por favor ingresa el teléfono.")
    else:
        with st.spinner("Buscando en el Excel y generando documento…"):
            try:
                df = _cargar_excel(st.secrets["EXCEL_ITEM_ID"])
                fila = df[df["Rut"].astype(str).str.strip() == rut.strip()]
                if fila.empty:
                    st.error(f"No se encontró asegurado con RUT: {rut}")
                    st.session_state.pop("visita_buf", None)
                else:
                    datos = fila.iloc[0]
                    cols_requeridas = ["Nro_Carpeta", "Num_Siniestro",
                                       "Dirección Riesgo Asegurado",
                                       "Asegurado", "Rut"]
                    faltan = [c for c in cols_requeridas if c not in datos.index]
                    if faltan:
                        st.warning(
                            f"Faltan columnas en el Excel: {', '.join(faltan)}. "
                            "Recargando cache para el próximo intento."
                        )
                        _cargar_excel.clear()
                    else:
                        doc_buf = llenar_visita(TEMPLATE_PATH, datos, telefono.strip())
                        st.session_state["visita_buf"]    = doc_buf.getvalue()
                        st.session_state["visita_nombre"] = f"Visita Tecnica {datos['Nro_Carpeta']}.docx"
                        st.session_state["visita_asegurado"] = str(datos["Asegurado"])
            except Exception as e:
                st.error(f"Error al generar el documento: {e}")
                st.session_state.pop("visita_buf", None)


if st.session_state.get("visita_buf"):
    st.success(f"Listo: **{st.session_state.get('visita_asegurado', '')}**")
    st.download_button(
        label     = f"⬇ Descargar {st.session_state['visita_nombre']}",
        data      = st.session_state["visita_buf"],
        file_name = st.session_state["visita_nombre"],
        mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )


# ── Opciones ────────────────────────────────────────────────────────────────
with st.expander("⚙️ Opciones"):
    if st.button("🔄 Recargar Excel desde OneDrive",
                 help="Úsalo si cambiaste algo en el Excel y la app sigue mostrando datos viejos."):
        _cargar_excel.clear()
        st.success("Cache del Excel limpiado. La próxima búsqueda traerá datos frescos.")
