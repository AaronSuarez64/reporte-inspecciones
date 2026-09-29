import os
import glob
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


def _encontrar_template() -> str:
    """Busca el .docx de la plantilla dentro de FormatoVisitaTecnica/.
    Es robusto a cualquier variación del nombre (acentos, mayúsculas, etc.)."""
    folder = "FormatoVisitaTecnica"
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"No existe la carpeta '{folder}' en el repo.")
    docs = [p for p in glob.glob(os.path.join(folder, "*.docx"))
            if not os.path.basename(p).startswith("~$")]
    if not docs:
        raise FileNotFoundError(f"No hay archivos .docx en '{folder}'.")
    return docs[0]


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

st.caption("Busca al asegurado por RUT o N° de carpeta e ingresa su teléfono. "
           "El resto se completa automáticamente desde el Excel.")


def _normalizar_rut(valor) -> str:
    """'6.817.145-8' y '6817145-8' quedan iguales (sin puntos/espacios, K mayúscula)."""
    return str(valor).replace(".", "").replace(" ", "").strip().upper()


def _normalizar_carpeta(valor) -> str:
    """Quita espacios y el '.0' que aparece cuando pandas lee la columna como decimal."""
    s = str(valor).strip()
    return s[:-2] if s.endswith(".0") else s


COLS_REQUERIDAS = ["Nro_Carpeta", "Num_Siniestro", "Dirección Riesgo Asegurado",
                   "Asegurado", "Rut"]


def _generar(datos, telefono: str):
    faltan = [c for c in COLS_REQUERIDAS if c not in datos.index]
    if faltan:
        st.warning(
            f"Faltan columnas en el Excel: {', '.join(faltan)}. "
            "Recargando cache para el próximo intento."
        )
        _cargar_excel.clear()
        return
    doc_buf = llenar_visita(_encontrar_template(), datos, telefono)
    st.session_state["visita_buf"]       = doc_buf.getvalue()
    st.session_state["visita_nombre"]    = f"Visita Tecnica {_normalizar_carpeta(datos['Nro_Carpeta'])}.docx"
    st.session_state["visita_asegurado"] = str(datos["Asegurado"])
    st.session_state.pop("visita_matches", None)


# ── Formulario ──────────────────────────────────────────────────────────────
modo = st.radio("Buscar por:", ["RUT", "Número de Carpeta"], horizontal=True, key="visita_modo")

with st.form("form_visita"):
    col1, col2 = st.columns(2)
    busqueda = col1.text_input(
        modo, placeholder="ej: 6817145-8" if modo == "RUT" else "ej: 488308"
    )
    telefono = col2.text_input("Teléfono", placeholder="ej: +56 9 1234 5678")
    submit = st.form_submit_button("Generar documento", type="primary", use_container_width=True)


if submit:
    st.session_state.pop("visita_buf", None)
    st.session_state.pop("visita_matches", None)
    if not busqueda.strip():
        st.error(f"Por favor ingresa el {modo}.")
    elif not telefono.strip():
        st.error("Por favor ingresa el teléfono.")
    else:
        with st.spinner("Buscando en el Excel y generando documento…"):
            try:
                df = _cargar_excel(st.secrets["EXCEL_ITEM_ID"])
                if modo == "RUT":
                    fila = df[df["Rut"].map(_normalizar_rut) == _normalizar_rut(busqueda)]
                else:
                    fila = df[df["Nro_Carpeta"].map(_normalizar_carpeta) == _normalizar_carpeta(busqueda)]

                if fila.empty:
                    st.error(f"No se encontró asegurado con {modo}: {busqueda.strip()}")
                elif len(fila) == 1:
                    _generar(fila.iloc[0], telefono.strip())
                else:
                    # Mismo RUT con varias carpetas: el inspector elige cuál
                    st.session_state["visita_matches"] = fila
                    st.session_state["visita_tel"]     = telefono.strip()
            except Exception as e:
                st.error(f"Error al generar el documento: {e}")


mm = st.session_state.get("visita_matches")
if mm is not None:
    st.warning(f"Se encontraron **{len(mm)}** carpetas para esta búsqueda. Selecciona la correcta:")
    for idx, row in mm.iterrows():
        etiqueta = (
            f"Carpeta {_normalizar_carpeta(row.get('Nro_Carpeta', '—'))} · "
            f"Siniestro {row.get('Num_Siniestro', '—')} · "
            f"{row.get('Dirección Riesgo Asegurado', '—')}"
        )
        if st.button(etiqueta, key=f"visita_pick_{idx}", use_container_width=True):
            try:
                _generar(row, st.session_state.get("visita_tel", ""))
            except Exception as e:
                st.error(f"Error al generar el documento: {e}")
            st.rerun()


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
