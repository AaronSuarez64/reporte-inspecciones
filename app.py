import io
import json
import time
import hashlib
import streamlit as st
import pandas as pd
from streamlit_local_storage import LocalStorage
from drive_utils import OneDriveClient, SharePointAppClient
from docx_utils import generar_documento


def _normalizar_imagen(buf: io.BytesIO) -> io.BytesIO:
    """Re-codifica cualquier imagen a JPEG limpio.
    Elimina EXIF problemáticos y convierte HEIC. Compatible con python-docx."""
    try:
        import pillow_heif
        from PIL import Image
        pillow_heif.register_heif_opener()
        buf.seek(0)
        out = io.BytesIO()
        Image.open(buf).convert("RGB").save(out, format="JPEG", quality=90)
        out.seek(0)
        return out
    except Exception:
        buf.seek(0)
        return buf


st.set_page_config(
    page_title="Generador de Informes de Siniestro",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS: mobile-friendly + textarea wrap ─────────────────────────────────────
st.markdown("""
<style>
/* Inputs y botones — touch targets cómodos, evita zoom en iOS */
.stTextInput input, .stTextArea textarea,
.stSelectbox div[role="combobox"], .stNumberInput input {
    font-size: 16px !important;
    min-height: 44px !important;
}
.stButton button, .stFormSubmitButton button, .stDownloadButton button {
    min-height: 44px !important;
    font-size: 15px !important;
}
/* Textareas: permitir crecer verticalmente, wrap del texto */
.stTextArea textarea {
    resize: vertical;
    overflow-wrap: break-word;
    word-wrap: break-word;
    white-space: pre-wrap;
    line-height: 1.4;
}

/* Observaciones display: el texto no se sale */
.obs-texto {
    word-wrap: break-word;
    overflow-wrap: break-word;
    white-space: pre-wrap;
    max-width: 100%;
    line-height: 1.5;
    padding: 4px 0;
}

/* Mobile: ajustes específicos */
@media (max-width: 768px) {
    .main .block-container {
        padding: 0.75rem 0.5rem !important;
        max-width: 100% !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        flex-wrap: wrap;
        gap: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 10px !important;
        font-size: 13px !important;
    }
    h1, h2 { font-size: 1.3rem !important; }
    h3 { font-size: 1.1rem !important; }
    /* Métricas más compactas */
    [data-testid="stMetricValue"] { font-size: 1.1rem !important; }
    [data-testid="stMetricLabel"] { font-size: 0.85rem !important; }
}
</style>
""", unsafe_allow_html=True)


# ── Persistencia en localStorage del navegador ──────────────────────────────
ls = LocalStorage()

STORAGE_KEY = "informe_siniestro_v1"

PERSIST_KEYS = [
    "datos_excel",
    "dir_editada",
    "detalle_visita",
    "hay_deteccion",
    "detecciones",
    "sin_danos",
    "danos_grupos",
    "observaciones",
    "search_mode",
    "nueva_zona_items",
    "nueva_zona_nombre",
    "nueva_zona_ubic",
]


def _serialize_state() -> str:
    state = {}
    for k in PERSIST_KEYS:
        v = st.session_state.get(k)
        if v is None or v == "" or v == [] or v == {}:
            continue
        if isinstance(v, pd.Series):
            state[k] = {"__type__": "Series",
                        "data": {kk: (None if pd.isna(vv) else vv) for kk, vv in v.to_dict().items()}}
        else:
            state[k] = v
    return json.dumps(state, default=str)


def _apply_state(raw: str):
    if not raw:
        return
    try:
        state = json.loads(raw)
    except Exception:
        return
    for k, v in state.items():
        if isinstance(v, dict) and v.get("__type__") == "Series":
            st.session_state[k] = pd.Series(v["data"])
        else:
            st.session_state[k] = v


def _init_state():
    defaults = {
        "datos_excel":         None,
        "dir_editada":         "",
        "detalle_visita":      "",
        "hay_deteccion":       False,
        "detecciones":         [],
        "edit_insp_idx":       None,
        "sin_danos":           False,
        "danos_grupos":        [],
        "edit_dano_idx":       None,
        "edit_dano_item":      None,
        "observaciones":       [],
        "edit_obs_idx":        None,
        "docx_bytes":          None,
        "nombre_archivo":      None,
        "search_mode":         "RUT",
        "nueva_zona_items":    [],
        "nueva_zona_nombre":   "",
        "nueva_zona_ubic":     "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# Restauración desde localStorage (puede tardar 1–2 reruns por el componente async)
if not st.session_state.get("_restored"):
    try:
        raw = ls.getItem(STORAGE_KEY)
    except Exception:
        raw = None
        st.session_state._restored = True  # si falla, seguimos sin persistencia
    if raw:
        _apply_state(raw)
        st.session_state._restored = True
    elif not st.session_state.get("_restored"):
        attempts = st.session_state.get("_restore_attempts", 0)
        st.session_state._restore_attempts = attempts + 1
        if attempts >= 2:
            st.session_state._restored = True
        else:
            with st.spinner("⏳ Cargando datos guardados…"):
                time.sleep(0.4)
            st.rerun()


# ── Clientes OneDrive cacheados ──────────────────────────────────────────────
@st.cache_resource
def _get_client_fotos() -> OneDriveClient:
    return OneDriveClient(
        client_id     = st.secrets["CLIENT_ID"],
        client_secret = st.secrets["CLIENT_SECRET"],
        refresh_token = st.secrets["REFRESH_TOKEN"],
    )

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


# ── Encabezado ───────────────────────────────────────────────────────────────
st.markdown(
    """
    <div style="background:#1a2a3a;padding:18px 24px;border-radius:8px;margin-bottom:18px">
        <h2 style="color:white;margin:0;font-family:Segoe UI,sans-serif">
            Generador de Informes de Siniestro
        </h2>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Opciones (limpiar formulario / recargar Excel) ──────────────────────────
with st.expander("⚙️ Opciones"):
    st.caption("Tu progreso se guarda automáticamente en este navegador.")
    col_op1, col_op2 = st.columns(2)
    if col_op1.button("🗑️ Limpiar todo el formulario", type="secondary", use_container_width=True):
        for k in PERSIST_KEYS + ["edit_insp_idx", "edit_dano_idx", "edit_dano_item",
                                  "edit_obs_idx", "docx_bytes", "nombre_archivo",
                                  "_last_saved_hash"]:
            if k in st.session_state:
                del st.session_state[k]
        try:
            ls.deleteItem(STORAGE_KEY)
        except Exception:
            try:
                ls.setItem(STORAGE_KEY, "")
            except Exception:
                pass
        _init_state()
        st.success("Formulario limpiado.")
        st.rerun()
    if col_op2.button("🔄 Recargar Excel desde OneDrive", type="secondary", use_container_width=True,
                      help="Úsalo si cambiaste algo en el Excel y la app sigue mostrando datos viejos."):
        _cargar_excel.clear()
        st.success("Cache del Excel limpiado. La próxima búsqueda traerá datos frescos.")


tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["1. Datos", "2. Visita", "3. Detección", "4. Daños", "5. Observaciones", "Generar"]
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 1 – Búsqueda por RUT o Carpeta
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab1:
    st.subheader("Búsqueda de Asegurado")

    modo = st.radio(
        "Buscar por:",
        options=["RUT", "Número de Carpeta"],
        horizontal=True,
        key="search_mode",
    )

    placeholder = "ej: 6817145-8" if modo == "RUT" else "ej: 488308"

    with st.form("form_busqueda", clear_on_submit=False):
        col_in, col_btn = st.columns([3, 1])
        busqueda_val = col_in.text_input(
            modo,
            placeholder=placeholder,
            label_visibility="collapsed",
            key="_search_input",
        )
        buscar = col_btn.form_submit_button("Buscar", use_container_width=True, type="primary")

    if buscar and busqueda_val.strip():
        try:
            df = _cargar_excel(st.secrets["EXCEL_ITEM_ID"])
            valor = busqueda_val.strip()
            if modo == "RUT":
                fila = df[df["Rut"].astype(str).str.strip() == valor]
            else:
                fila = df[df["Nro_Carpeta"].astype(str).str.strip() == valor]

            if fila.empty:
                st.error(f"No se encontró asegurado con {modo}: {valor}")
                st.session_state.datos_excel = None
            else:
                st.session_state.datos_excel = fila.iloc[0]
                st.session_state.dir_editada = str(fila.iloc[0]["Dirección Riesgo Asegurado"])
                st.success(f"Asegurado encontrado: **{st.session_state.datos_excel['Asegurado']}**")
        except Exception as e:
            st.error(f"Error al conectar con OneDrive: {e}")

    if st.session_state.datos_excel is not None:
        d = st.session_state.datos_excel
        cols_requeridas = ["Nro_Carpeta", "Num_Siniestro", "Dirección Riesgo Asegurado",
                           "Comuna", "Asegurado", "Rut"]
        faltan = [c for c in cols_requeridas if c not in d.index]
        if faltan:
            st.warning(
                f"Los datos están desactualizados (faltan columnas: {', '.join(faltan)}). "
                "Limpiando cache del Excel y borrando datos guardados. Por favor vuelve a buscar."
            )
            st.session_state.datos_excel = None
            st.session_state.dir_editada = ""
            _cargar_excel.clear()
        else:
            st.divider()
            st.subheader("Datos del Siniestro")
            col_a, col_b = st.columns(2)
            col_a.metric("Número de Carpeta",   str(d["Nro_Carpeta"]))
            col_b.metric("Número de Siniestro", str(d["Num_Siniestro"]))
            st.text_input("Dirección ✏️ (editable)", key="dir_editada")
            st.text_input("Comuna",           value=str(d["Comuna"]),    disabled=True)
            st.text_input("Nombre Asegurado", value=str(d["Asegurado"]), disabled=True)
            st.text_input("RUT Asegurado",    value=str(d["Rut"]),       disabled=True)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 2 – Detalle de la Visita
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab2:
    st.subheader("Detalle de la Visita")
    st.text_area(
        "Detalle",
        key="detalle_visita",
        height=420,
        placeholder="Describa los detalles de la visita realizada...",
        label_visibility="collapsed",
    )

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 3 – Inspección Técnica
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab3:
    st.checkbox("Se realizó detección", key="hay_deteccion")

    if st.session_state.hay_deteccion:
        st.divider()
        detecciones = st.session_state.detecciones

        if detecciones:
            st.markdown("**Metodologías agregadas:**")
            for i, insp in enumerate(detecciones):
                col_n, col_e, col_d = st.columns([7, 1, 1])
                col_n.write(f"**{i + 1}. {insp['nombre']}**")
                if col_e.button("✎", key=f"edit_insp_{i}", help="Editar"):
                    st.session_state.edit_insp_idx = i
                    st.rerun()
                if col_d.button("✕", key=f"del_insp_{i}", help="Eliminar"):
                    detecciones.pop(i)
                    st.session_state.detecciones = detecciones
                    st.rerun()
        else:
            st.info("No hay detecciones agregadas. Use el botón para agregar.")

        if st.button("+ Agregar metodología"):
            st.session_state.edit_insp_idx = -1
            # Reset buffer de edición
            st.session_state["_edit_dims"] = []
            st.session_state["_edit_dim_idx"] = None
            st.rerun()

        edit_idx = st.session_state.edit_insp_idx
        if edit_idx is not None:
            st.divider()
            existente = detecciones[edit_idx] if edit_idx >= 0 else None
            st.subheader("Nueva metodología" if existente is None else f"Editando: {existente['nombre']}")

            # Inicializar buffer de dimensiones la primera vez que entramos en edición
            if "_edit_dims_owner" not in st.session_state or st.session_state["_edit_dims_owner"] != edit_idx:
                st.session_state["_edit_dims"] = (
                    [d.copy() for d in existente["dimensiones"]] if existente else []
                )
                st.session_state["_edit_dims_owner"] = edit_idx
                st.session_state["_edit_dim_idx"] = None

            # Formulario principal (nombre + texto + acciones)
            with st.form("form_deteccion"):
                nombre = st.text_input(
                    "Nombre de la metodología",
                    value=existente["nombre"] if existente else "",
                    placeholder="ej: Gas trazador",
                )
                texto = st.text_area(
                    "Texto de la metodología",
                    value=existente["texto"] if existente else "",
                    height=200,
                )
                col_g, col_c = st.columns(2)
                guardado  = col_g.form_submit_button("Guardar", use_container_width=True, type="primary")
                cancelado = col_c.form_submit_button("Cancelar", use_container_width=True)

            # ── Dimensiones (fuera del form para add/edit/delete sin Enter) ──
            st.markdown("**Dimensiones** *(opcional)*")
            dims_lista = st.session_state["_edit_dims"]

            for di, dim in enumerate(dims_lista):
                if st.session_state.get("_edit_dim_idx") == di:
                    with st.form(f"form_edit_dim_{di}", clear_on_submit=False):
                        c1, c2 = st.columns([3, 5])
                        nt = c1.text_input("Título", value=dim["titulo"], label_visibility="collapsed")
                        nx = c2.text_input("Texto",  value=dim["texto"],  label_visibility="collapsed")
                        cg, cc = st.columns(2)
                        if cg.form_submit_button("✓ Guardar", use_container_width=True):
                            dims_lista[di] = {"titulo": nt.strip(), "texto": nx.strip()}
                            st.session_state["_edit_dim_idx"] = None
                            st.rerun()
                        if cc.form_submit_button("Cancelar", use_container_width=True):
                            st.session_state["_edit_dim_idx"] = None
                            st.rerun()
                else:
                    col_t, col_x, col_e, col_del = st.columns([3, 5, 1, 1])
                    col_t.markdown(f"**{dim['titulo']}**")
                    col_x.write(dim["texto"])
                    if col_e.button("✎", key=f"ed_dim_{di}", help="Editar"):
                        st.session_state["_edit_dim_idx"] = di
                        st.rerun()
                    if col_del.button("✕", key=f"del_dim_{di}", help="Eliminar"):
                        dims_lista.pop(di)
                        st.rerun()

            with st.form("form_add_dim", clear_on_submit=True):
                c1, c2, c3 = st.columns([3, 5, 1])
                d_tit = c1.text_input("Título", placeholder="Título", label_visibility="collapsed")
                d_txt = c2.text_input("Texto / Medidas", placeholder="Texto / Medidas", label_visibility="collapsed")
                if c3.form_submit_button("+", use_container_width=True):
                    if d_tit.strip():
                        dims_lista.append({"titulo": d_tit.strip(), "texto": d_txt.strip()})
                        st.rerun()

            if guardado:
                if not nombre.strip():
                    st.warning("Ingrese el nombre de la metodología.")
                else:
                    data = {
                        "nombre": nombre.strip(),
                        "texto": texto,
                        "dimensiones": [d.copy() for d in dims_lista],
                    }
                    if edit_idx == -1:
                        detecciones.append(data)
                    else:
                        detecciones[edit_idx] = data
                    st.session_state.detecciones = detecciones
                    st.session_state.edit_insp_idx = None
                    for k in ("_edit_dims", "_edit_dims_owner", "_edit_dim_idx"):
                        if k in st.session_state:
                            del st.session_state[k]
                    st.rerun()

            if cancelado:
                st.session_state.edit_insp_idx = None
                for k in ("_edit_dims", "_edit_dims_owner", "_edit_dim_idx"):
                    if k in st.session_state:
                        del st.session_state[k]
                st.rerun()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 4 – Daños Identificados
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab4:
    st.subheader("Daños Identificados")
    st.checkbox("No se observaron daños en el inmueble", key="sin_danos")

    if st.session_state.sin_danos:
        st.info("Se incluirá en el informe: «No se observan daños en el inmueble.»")
    else:
        grupos = st.session_state.danos_grupos

        for gi, grupo in enumerate(grupos):
            with st.expander(f"**{grupo['zona']}** — {grupo['ubicacion']}", expanded=True):

                # Editar zona/ubicación
                if st.session_state.edit_dano_idx == gi:
                    with st.form(f"form_edit_zona_{gi}"):
                        c1, c2 = st.columns(2)
                        nz = c1.text_input("Zona Afectada", value=grupo["zona"])
                        nu = c2.text_input("Ubicación",     value=grupo["ubicacion"])
                        cg, cc = st.columns(2)
                        if cg.form_submit_button("Guardar", type="primary"):
                            grupos[gi]["zona"]      = nz.strip()
                            grupos[gi]["ubicacion"] = nu.strip()
                            st.session_state.edit_dano_idx = None
                            st.rerun()
                        if cc.form_submit_button("Cancelar"):
                            st.session_state.edit_dano_idx = None
                            st.rerun()

                # Ítems existentes
                for ii, item in enumerate(grupo["items"]):
                    sup_d = f"{item['superficie']} {item['unidad']}" if item["unidad"] != "sin unidad" and item["superficie"] else item["superficie"] or "—"
                    col_t, col_e, col_d = st.columns([8, 1, 1])
                    col_t.write(f"• {item['dano']}  —  {sup_d}")
                    if col_e.button("✎", key=f"edit_item_{gi}_{ii}"):
                        st.session_state.edit_dano_item = [gi, ii]
                        st.rerun()
                    if col_d.button("✕", key=f"del_item_{gi}_{ii}"):
                        grupo["items"].pop(ii)
                        if st.session_state.edit_dano_item == [gi, ii]:
                            st.session_state.edit_dano_item = None
                        st.rerun()

                # Formulario de edición de ítem
                edit_dano_item = st.session_state.edit_dano_item
                if edit_dano_item is not None and edit_dano_item[0] == gi:
                    _, eii = edit_dano_item
                    if eii < len(grupo["items"]):
                        item_ed = grupo["items"][eii]
                        st.divider()
                        with st.form(f"form_edit_item_{gi}_{eii}"):
                            c1, c2, c3 = st.columns([4, 2, 2])
                            nd = c1.text_input("Daño",       value=item_ed["dano"])
                            ns = c2.text_input("Superficie", value=item_ed["superficie"])
                            opciones = ["m²", "m", "sin unidad"]
                            idx_u = opciones.index(item_ed["unidad"]) if item_ed["unidad"] in opciones else 0
                            nu = c3.selectbox("Unidad", opciones, index=idx_u)
                            cg, cc = st.columns(2)
                            if cg.form_submit_button("Guardar", type="primary"):
                                grupo["items"][eii] = {"dano": nd.strip(), "superficie": ns.strip(), "unidad": nu}
                                st.session_state.edit_dano_item = None
                                st.rerun()
                            if cc.form_submit_button("Cancelar"):
                                st.session_state.edit_dano_item = None
                                st.rerun()

                # Agregar ítem inline (form → captura al hacer click, sin Enter)
                with st.form(f"form_item_{gi}", clear_on_submit=True):
                    c1, c2, c3 = st.columns([4, 2, 2])
                    nd = c1.text_input("Daño",       label_visibility="collapsed", placeholder="Daño observado")
                    ns = c2.text_input("Superficie", label_visibility="collapsed", placeholder="Superficie")
                    nu = c3.selectbox("Unidad", ["m²", "m", "sin unidad"], label_visibility="collapsed")
                    if st.form_submit_button("+ Agregar daño"):
                        if nd.strip():
                            grupo["items"].append({"dano": nd.strip(), "superficie": ns.strip(), "unidad": nu})
                            st.rerun()

                cb1, cb2 = st.columns(2)
                if cb1.button("✎ Editar zona/ubicación", key=f"edit_grupo_{gi}"):
                    st.session_state.edit_dano_idx = gi
                    st.rerun()
                if cb2.button("✕ Eliminar zona", key=f"del_grupo_{gi}"):
                    grupos.pop(gi)
                    st.session_state.danos_grupos = grupos
                    st.rerun()

        # ── Agregar nueva zona (sin data_editor, basado en form) ────────────
        st.divider()
        st.markdown("**Agregar nueva zona**")

        # Inputs de zona/ubicación con sincronización directa a session_state
        c1, c2 = st.columns(2)
        c1.text_input(
            "Zona Afectada", placeholder="ej: Muro",
            key="nueva_zona_nombre",
        )
        c2.text_input(
            "Ubicación", placeholder="ej: Baño principal",
            key="nueva_zona_ubic",
        )

        st.caption("Daños de esta nueva zona:")
        for ti, it in enumerate(st.session_state.nueva_zona_items):
            sup_d = f"{it['superficie']} {it['unidad']}" if it["unidad"] != "sin unidad" and it["superficie"] else it["superficie"] or "—"
            col_t, col_d = st.columns([8, 1])
            col_t.write(f"• {it['dano']}  —  {sup_d}")
            if col_d.button("✕", key=f"del_nz_item_{ti}"):
                st.session_state.nueva_zona_items.pop(ti)
                st.rerun()

        with st.form("form_nueva_zona_item", clear_on_submit=True):
            c1, c2, c3 = st.columns([4, 2, 2])
            nd = c1.text_input("Daño",       placeholder="Daño observado", label_visibility="collapsed")
            ns = c2.text_input("Superficie", placeholder="Superficie",     label_visibility="collapsed")
            nu = c3.selectbox("Unidad", ["m²", "m", "sin unidad"], label_visibility="collapsed")
            if st.form_submit_button("+ Agregar daño a esta zona"):
                if nd.strip():
                    st.session_state.nueva_zona_items.append({
                        "dano": nd.strip(),
                        "superficie": ns.strip(),
                        "unidad": nu,
                    })
                    st.rerun()

        def _crear_zona_callback():
            nombre = st.session_state.get("nueva_zona_nombre", "").strip()
            if not nombre:
                return
            ubic = st.session_state.get("nueva_zona_ubic", "").strip()
            grupos_cb = st.session_state.danos_grupos
            grupos_cb.append({
                "zona":      nombre,
                "ubicacion": ubic,
                "items":     [it.copy() for it in st.session_state.nueva_zona_items],
            })
            st.session_state.danos_grupos       = grupos_cb
            st.session_state.nueva_zona_items   = []
            # Dentro de un callback sí podemos reasignar widget keys
            st.session_state.nueva_zona_nombre  = ""
            st.session_state.nueva_zona_ubic    = ""

        zona_lista = st.session_state.get("nueva_zona_nombre", "").strip()
        st.button(
            "✓ Crear zona", type="primary", use_container_width=True,
            disabled=not zona_lista, on_click=_crear_zona_callback,
        )

        if not zona_lista:
            st.caption("ℹ️ Ingresa el nombre de la zona para habilitar la creación.")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 5 – Observaciones
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab5:
    st.subheader("Observaciones")
    obs_lista = st.session_state.observaciones

    if not obs_lista:
        st.caption("No hay observaciones. Agrega una con el formulario de abajo.")

    for i, obs in enumerate(obs_lista):
        col_t, col_btns = st.columns([7, 3])
        col_t.markdown(
            f"<div class='obs-texto'><b>{i + 1}.</b> {obs}</div>",
            unsafe_allow_html=True,
        )
        bc1, bc2, bc3, bc4 = col_btns.columns(4)
        up = bc1.button("↑", key=f"up_obs_{i}", disabled=(i == 0))
        dn = bc2.button("↓", key=f"dn_obs_{i}", disabled=(i == len(obs_lista) - 1))
        ed = bc3.button("✎", key=f"ed_obs_{i}")
        dl = bc4.button("✕", key=f"dl_obs_{i}")

        if up:
            obs_lista[i - 1], obs_lista[i] = obs_lista[i], obs_lista[i - 1]
            st.session_state.observaciones = obs_lista
            st.rerun()
        if dn:
            obs_lista[i + 1], obs_lista[i] = obs_lista[i], obs_lista[i + 1]
            st.session_state.observaciones = obs_lista
            st.rerun()
        if ed:
            st.session_state.edit_obs_idx = i
            st.rerun()
        if dl:
            obs_lista.pop(i)
            st.session_state.observaciones = obs_lista
            if st.session_state.edit_obs_idx == i:
                st.session_state.edit_obs_idx = None
            st.rerun()

    # Edición — text_area para que textos largos crezcan
    edit_obs_idx = st.session_state.edit_obs_idx
    if edit_obs_idx is not None and edit_obs_idx < len(obs_lista):
        st.divider()
        with st.form("form_edit_obs"):
            texto_ed = st.text_area(
                "Editar observación",
                value=obs_lista[edit_obs_idx],
                height=140,
            )
            col_g, col_c = st.columns(2)
            if col_g.form_submit_button("Guardar", type="primary"):
                if texto_ed.strip():
                    obs_lista[edit_obs_idx] = texto_ed.strip()
                    st.session_state.observaciones = obs_lista
                st.session_state.edit_obs_idx = None
                st.rerun()
            if col_c.form_submit_button("Cancelar"):
                st.session_state.edit_obs_idx = None
                st.rerun()

    st.divider()
    # Nueva observación con text_area (multilínea)
    with st.form("form_obs", clear_on_submit=True):
        nueva_obs = st.text_area(
            "Nueva observación",
            label_visibility="collapsed",
            placeholder="Escribe la observación aquí...",
            height=120,
        )
        if st.form_submit_button("+ Agregar observación"):
            if nueva_obs.strip():
                obs_lista.append(nueva_obs.strip())
                st.session_state.observaciones = obs_lista
                st.rerun()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 6 – Generar Reporte
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab6:
    st.subheader("Generar Reporte")

    if st.button("Generar Reporte ▶", type="primary", use_container_width=True):
        if st.session_state.datos_excel is None:
            st.error("Primero busque un asegurado por RUT o Número de Carpeta (pestaña 1).")
        else:
            datos = st.session_state.datos_excel.copy()
            datos["Dirección Riesgo Asegurado"] = st.session_state.dir_editada or datos["Dirección Riesgo Asegurado"]

            danos_list = []
            if not st.session_state.sin_danos:
                for grupo in st.session_state.danos_grupos:
                    items = []
                    for item in grupo["items"]:
                        sup = item["superficie"]
                        if sup and item["unidad"] != "sin unidad":
                            sup = f"{sup} {item['unidad']}"
                        items.append({"dano": item["dano"], "superficie": sup or "—"})
                    if items:
                        danos_list.append({"zona": grupo["zona"], "ubicacion": grupo["ubicacion"], "items": items})

            with st.spinner("Descargando fotos y generando reporte…"):
                try:
                    client = _get_client_fotos()
                    imagenes_meta, imagen_error = client.obtener_imagenes(
                        st.secrets["FOTOS_ITEM_ID"], datos["Nro_Carpeta"]
                    )
                    imagenes_data = []
                    for img in imagenes_meta:
                        buf = client.descargar_por_id(img["item_id"])
                        buf = _normalizar_imagen(buf)
                        imagenes_data.append({"descripcion": img["descripcion"], "bytes": buf})

                    docx_buf = generar_documento(
                        datos          = datos,
                        detalle_visita = st.session_state.detalle_visita,
                        hay_inspeccion = st.session_state.hay_deteccion,
                        inspecciones   = st.session_state.detecciones,
                        sin_danos      = st.session_state.sin_danos,
                        danos          = danos_list,
                        observaciones  = st.session_state.observaciones,
                        imagenes_data  = imagenes_data,
                        imagen_error   = imagen_error,
                    )
                    st.session_state.nombre_archivo = f"Informe Inspección {datos['Nro_Carpeta']}.docx"
                    st.session_state.docx_bytes     = docx_buf.getvalue()
                    st.success("Reporte generado.")
                except Exception as e:
                    import traceback
                    st.error(f"Error al generar el documento: {repr(e)}")
                    st.code(traceback.format_exc())

    if st.session_state.docx_bytes is not None:
        st.download_button(
            label     = f"⬇ Descargar {st.session_state.nombre_archivo}",
            data      = st.session_state.docx_bytes,
            file_name = st.session_state.nombre_archivo,
            mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

# ── Auto-guardado al final de cada render (con dedupe por hash) ─────────────
try:
    state_str = _serialize_state()
    new_hash = hashlib.md5(state_str.encode("utf-8")).hexdigest()
    if new_hash != st.session_state.get("_last_saved_hash"):
        ls.setItem(STORAGE_KEY, state_str)
        st.session_state._last_saved_hash = new_hash
except Exception:
    pass
