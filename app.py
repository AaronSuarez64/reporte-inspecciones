import streamlit as st
import pandas as pd
from drive_utils import OneDriveClient
from docx_utils import generar_documento

st.set_page_config(page_title="Generador de Informes de Siniestro", layout="wide")

# ── Inicializar session state ────────────────────────────────────────────────

def _init_state():
    defaults = {
        "datos_excel":    None,
        "dir_editada":    "",
        "inspecciones":   [],
        "edit_insp_idx":  None,
        "observaciones":  [],
        "edit_obs_idx":   None,
        "docx_bytes":     None,
        "nombre_archivo": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ── Cliente OneDrive cacheado ────────────────────────────────────────────────

@st.cache_resource
def _get_client() -> OneDriveClient:
    return OneDriveClient(
        client_id     = st.secrets["CLIENT_ID"],
        client_secret = st.secrets["CLIENT_SECRET"],
        refresh_token = st.secrets["REFRESH_TOKEN"],
    )

@st.cache_data(ttl=3600)
def _cargar_excel(item_id: str) -> pd.DataFrame:
    client = _get_client()
    return client.leer_excel(item_id, "Nexus")  # cambiar "Nexus" si la hoja tiene otro nombre

# ── Encabezado ───────────────────────────────────────────────────────────────

st.markdown(
    """
    <div style="background:#1a2a3a;padding:18px 24px;border-radius:8px;margin-bottom:24px">
        <h2 style="color:white;margin:0;font-family:Segoe UI,sans-serif">
            Generador de Informes de Siniestro
        </h2>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Pestañas ─────────────────────────────────────────────────────────────────

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["  1. Datos  ", "  2. Visita  ", "  3. Inspección  ", "  4. Daños  ", "  5. Observaciones  ", "  Generar Reporte  "]
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 1 – Búsqueda por RUT
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab1:
    st.subheader("Búsqueda de Asegurado")

    col_rut, col_btn = st.columns([3, 1])
    rut_input = col_rut.text_input(
        "RUT", placeholder="ej: 6817145-8", label_visibility="collapsed"
    )
    buscar = col_btn.button("Buscar", use_container_width=True)

    if buscar and rut_input:
        try:
            df = _cargar_excel(st.secrets["EXCEL_ITEM_ID"])
            fila = df[df["Rut"] == rut_input.strip()]
            if fila.empty:
                st.error(f"No se encontró asegurado con RUT: {rut_input}")
                st.session_state.datos_excel = None
            else:
                st.session_state.datos_excel = fila.iloc[0]
                st.session_state.dir_editada = str(fila.iloc[0]["Dirección Riesgo Asegurado"])
                st.success(f"Asegurado encontrado: **{st.session_state.datos_excel['Asegurado']}**")
        except Exception as e:
            st.error(f"Error al conectar con OneDrive: {e}")

    if st.session_state.datos_excel is not None:
        d = st.session_state.datos_excel
        st.divider()
        st.subheader("Datos del Siniestro")
        col_a, col_b = st.columns(2)
        col_a.metric("Número de Carpeta",   str(d["Nro_Carpeta"]))
        col_b.metric("Número de Siniestro", str(d["Num_Siniestro"]))
        st.text_input("Dirección ✏️ (editable)", key="dir_editada")
        st.text_input("Nombre Asegurado",   value=str(d["Asegurado"]),                 disabled=True)
        st.text_input("RUT Asegurado",      value=str(d["Rut"]),             disabled=True)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 2 – Detalle de la Visita
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab2:
    st.subheader("Detalle de la Visita")
    detalle_visita = st.text_area(
        "Detalle", height=420,
        placeholder="Describa los detalles de la visita realizada...",
        label_visibility="collapsed",
    )

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 3 – Inspección Técnica
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab3:
    hay_inspeccion = st.checkbox("Se realizó inspección técnica")

    if hay_inspeccion:
        st.divider()
        inspecciones = st.session_state.inspecciones

        if inspecciones:
            st.markdown("**Metodologías agregadas:**")
            for i, insp in enumerate(inspecciones):
                col_n, col_e, col_d = st.columns([7, 1, 1])
                col_n.write(f"**{i + 1}. {insp['nombre']}**")
                if col_e.button("✎", key=f"edit_{i}", help="Editar"):
                    st.session_state.edit_insp_idx = i
                    st.rerun()
                if col_d.button("✕", key=f"del_{i}", help="Eliminar"):
                    inspecciones.pop(i)
                    st.session_state.inspecciones = inspecciones
                    st.rerun()
        else:
            st.info("No hay inspecciones agregadas. Use el botón para agregar.")

        if st.button("+ Agregar metodología"):
            st.session_state.edit_insp_idx = -1
            st.rerun()

        edit_idx = st.session_state.edit_insp_idx
        if edit_idx is not None:
            st.divider()
            existente = inspecciones[edit_idx] if edit_idx >= 0 else None
            st.subheader("Nueva metodología" if existente is None else f"Editando: {existente['nombre']}")

            with st.form("form_inspeccion"):
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
                st.markdown("**Dimensiones** *(opcional)*")
                dims_iniciales = pd.DataFrame(
                    existente["dimensiones"] if existente else [],
                    columns=["titulo", "texto"],
                )
                dims_editor = st.data_editor(
                    dims_iniciales,
                    num_rows="dynamic",
                    use_container_width=True,
                    column_config={
                        "titulo": st.column_config.TextColumn("Título",          width="medium"),
                        "texto":  st.column_config.TextColumn("Texto / Medidas", width="large"),
                    },
                )
                col_g, col_c = st.columns(2)
                guardado  = col_g.form_submit_button("Guardar",   use_container_width=True, type="primary")
                cancelado = col_c.form_submit_button("Cancelar",  use_container_width=True)

            if guardado:
                if not nombre.strip():
                    st.warning("Ingrese el nombre del espacio.")
                else:
                    dims = [
                        {"titulo": row["titulo"], "texto": row["texto"] or ""}
                        for _, row in dims_editor.iterrows()
                        if row["titulo"]
                    ]
                    data = {"nombre": nombre.strip(), "texto": texto, "dimensiones": dims}
                    if edit_idx == -1:
                        inspecciones.append(data)
                    else:
                        inspecciones[edit_idx] = data
                    st.session_state.inspecciones = inspecciones
                    st.session_state.edit_insp_idx = None
                    st.rerun()

            if cancelado:
                st.session_state.edit_insp_idx = None
                st.rerun()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 4 – Daños Identificados
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab4:
    st.subheader("Daños Identificados")
    sin_danos = st.checkbox("No se observaron daños en el inmueble")

    if sin_danos:
        st.info("Se incluirá en el informe: «No se observan daños en el inmueble.»")
        danos_df = pd.DataFrame({
            "Zona Afectada":  pd.Series([], dtype=str),
            "Ubicación":      pd.Series([], dtype=str),
            "Daño Observado": pd.Series([], dtype=str),
            "Superficie":     pd.Series([], dtype=str),
            "Unidad":         pd.Series([], dtype=str),
        })
    else:
        st.caption("Una fila por cada daño. La unidad aplica a todos los valores de superficie de esa fila.")
        danos_df = st.data_editor(
            pd.DataFrame({
                "Zona Afectada":  pd.Series([], dtype=str),
                "Ubicación":      pd.Series([], dtype=str),
                "Daño Observado": pd.Series([], dtype=str),
                "Superficie":     pd.Series([], dtype=str),
                "Unidad":         pd.Series(["m²"], dtype=str),
            }),
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Zona Afectada":  st.column_config.TextColumn(width="medium"),
                "Ubicación":      st.column_config.TextColumn(width="medium"),
                "Daño Observado": st.column_config.TextColumn(width="large"),
                "Superficie":     st.column_config.TextColumn("Superficie", width="small"),
                "Unidad":         st.column_config.SelectboxColumn(
                    "Unidad",
                    options=["m²", "m", "sin unidad"],
                    default="m²",
                    width="small",
                ),
            },
            key="danos_editor",
        )

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PESTAÑA 5 – Observaciones
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab5:
    st.subheader("Observaciones")
    obs_lista = st.session_state.observaciones

    if not obs_lista:
        st.caption("No hay observaciones. Agrega una con el formulario de abajo.")

    for i, obs in enumerate(obs_lista):
        col_t, col_up, col_dn, col_e, col_d = st.columns([6, 1, 1, 1, 1])
        col_t.markdown(f"**{i + 1}.** {obs}")

        up = col_up.button("↑", key=f"up_obs_{i}", disabled=(i == 0))
        dn = col_dn.button("↓", key=f"dn_obs_{i}", disabled=(i == len(obs_lista) - 1))
        ed = col_e.button("✎",  key=f"ed_obs_{i}")
        dl = col_d.button("✕",  key=f"dl_obs_{i}")

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

    # Formulario de edición
    edit_obs_idx = st.session_state.edit_obs_idx
    if edit_obs_idx is not None and edit_obs_idx < len(obs_lista):
        st.divider()
        with st.form("form_edit_obs"):
            texto_ed = st.text_input("Editar observación", value=obs_lista[edit_obs_idx])
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
    with st.form("form_obs", clear_on_submit=True):
        nueva_obs = st.text_input("Nueva observación", label_visibility="collapsed",
                                  placeholder="Escribe la observación aquí...")
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
            st.error("Primero busque un asegurado por RUT (pestaña 1).")
        else:
            datos = st.session_state.datos_excel.copy()
            datos["Dirección Riesgo Asegurado"] = st.session_state.dir_editada or datos["Dirección Riesgo Asegurado"]

            danos_list = []
            if not sin_danos:
                filas_validas = danos_df.dropna(subset=["Zona Afectada"]).copy()
                filas_validas = filas_validas[filas_validas["Zona Afectada"].str.strip() != ""]
                for _, row in filas_validas.iterrows():
                    zona   = str(row["Zona Afectada"]).strip()
                    ubic   = str(row["Ubicación"]).strip()
                    unidad = str(row.get("Unidad", "m²")).strip()
                    dano_lines = [l.strip() for l in str(row["Daño Observado"]).split("\n") if l.strip()]
                    sup_lines  = [l.strip() for l in str(row["Superficie"]).split("\n")     if l.strip()]
                    if not dano_lines:
                        continue
                    items = []
                    for idx, d_line in enumerate(dano_lines):
                        s_val = sup_lines[idx] if idx < len(sup_lines) else ""
                        if s_val:
                            s_line = s_val if unidad == "sin unidad" else f"{s_val} {unidad}"
                        else:
                            s_line = "—"
                        items.append({"dano": d_line, "superficie": s_line})
                    danos_list.append({"zona": zona, "ubicacion": ubic, "items": items})

            with st.spinner("Descargando fotos y generando reporte…"):
                try:
                    client = _get_client()
                    imagenes_meta, imagen_error = client.obtener_imagenes(
                        st.secrets["FOTOS_ITEM_ID"], datos["Nro_Carpeta"]
                    )
                    imagenes_data = []
                    for img in imagenes_meta:
                        buf = client.descargar_por_id(img["item_id"])
                        imagenes_data.append({"descripcion": img["descripcion"], "bytes": buf})

                    docx_buf = generar_documento(
                        datos          = datos,
                        detalle_visita = detalle_visita,
                        hay_inspeccion = hay_inspeccion,
                        inspecciones   = st.session_state.inspecciones,
                        sin_danos      = sin_danos,
                        danos          = danos_list,
                        observaciones  = st.session_state.observaciones,
                        imagenes_data  = imagenes_data,
                        imagen_error   = imagen_error,
                    )
                    st.session_state.nombre_archivo = f"Informe Inspección {datos['Nro_Carpeta']}.docx"
                    st.session_state.docx_bytes     = docx_buf.getvalue()
                    st.success("Reporte generado.")
                except Exception as e:
                    st.error(f"Error al generar el documento: {e}")

    if st.session_state.docx_bytes is not None:
        st.download_button(
            label     = f"⬇ Descargar {st.session_state.nombre_archivo}",
            data      = st.session_state.docx_bytes,
            file_name = st.session_state.nombre_archivo,
            mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
