import io
import streamlit as st
import pandas as pd
from drive_utils import OneDriveClient
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

st.set_page_config(page_title="Generador de Informes de Siniestro", layout="wide")

# ── Inicializar session state ────────────────────────────────────────────────

def _init_state():
    defaults = {
        "datos_excel":    None,
        "dir_editada":    "",
        "detecciones":    [],
        "edit_insp_idx":  None,
        "danos_grupos":   [],
        "edit_dano_idx":  None,
        "zona_form_v":    0,
        "observaciones":  [],
        "edit_obs_idx":   None,
        "docx_bytes":     None,
        "nombre_archivo": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ── Clientes OneDrive cacheados ──────────────────────────────────────────────

@st.cache_resource
def _get_client_fotos() -> OneDriveClient:
    """Cliente para el OneDrive de fotos."""
    return OneDriveClient(
        client_id     = st.secrets["CLIENT_ID"],
        client_secret = st.secrets["CLIENT_SECRET"],
        refresh_token = st.secrets["REFRESH_TOKEN"],
    )

@st.cache_resource
def _get_client_excel() -> OneDriveClient:
    """Cliente para el OneDrive/SharePoint del Excel."""
    return OneDriveClient(
        client_id     = st.secrets["CLIENT_ID"],
        client_secret = st.secrets["CLIENT_SECRET"],
        refresh_token = st.secrets["REFRESH_TOKEN_EXCEL"],
    )

@st.cache_data(ttl=3600)
def _cargar_excel(item_id: str) -> pd.DataFrame:
    return _get_client_excel().leer_excel(item_id, "Nexus")

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
    ["  1. Datos  ", "  2. Visita  ", "  3. Detección  ", "  4. Daños  ", "  5. Observaciones  ", "  Generar Reporte  "]
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
        st.text_input("Comuna",             value=str(d["Comuna"]),                    disabled=True)
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
    hay_deteccion = st.checkbox("Se realizó detección")

    if hay_deteccion:
        st.divider()
        detecciones = st.session_state.detecciones

        if detecciones:
            st.markdown("**Metodologías agregadas:**")
            for i, insp in enumerate(detecciones):
                col_n, col_e, col_d = st.columns([7, 1, 1])
                col_n.write(f"**{i + 1}. {insp['nombre']}**")
                if col_e.button("✎", key=f"edit_{i}", help="Editar"):
                    st.session_state.edit_insp_idx = i
                    st.rerun()
                if col_d.button("✕", key=f"del_{i}", help="Eliminar"):
                    detecciones.pop(i)
                    st.session_state.detecciones = detecciones
                    st.rerun()
        else:
            st.info("No hay detecciones agregadas. Use el botón para agregar.")

        if st.button("+ Agregar metodología"):
            st.session_state.edit_insp_idx = -1
            st.rerun()

        edit_idx = st.session_state.edit_insp_idx
        if edit_idx is not None:
            st.divider()
            existente = detecciones[edit_idx] if edit_idx >= 0 else None
            st.subheader("Nueva metodología" if existente is None else f"Editando: {existente['nombre']}")

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
                        {
                            "titulo": str(row["titulo"]).strip(),
                            "texto":  "" if pd.isna(row["texto"]) else str(row["texto"]).strip(),
                        }
                        for _, row in dims_editor.iterrows()
                        if pd.notna(row["titulo"]) and str(row["titulo"]).strip()
                    ]
                    data = {"nombre": nombre.strip(), "texto": texto, "dimensiones": dims}
                    if edit_idx == -1:
                        detecciones.append(data)
                    else:
                        detecciones[edit_idx] = data
                    st.session_state.detecciones = detecciones
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
    else:
        grupos = st.session_state.danos_grupos

        # ── Mostrar grupos existentes ───────────────────────────────────────
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
                    col_t, col_d = st.columns([9, 1])
                    col_t.write(f"• {item['dano']}  —  {sup_d}")
                    if col_d.button("✕", key=f"del_item_{gi}_{ii}"):
                        grupo["items"].pop(ii)
                        st.rerun()

                # Agregar ítem inline
                with st.form(f"form_item_{gi}", clear_on_submit=True):
                    c1, c2, c3 = st.columns([4, 2, 2])
                    nd = c1.text_input("Daño",       label_visibility="collapsed", placeholder="Daño observado")
                    ns = c2.text_input("Superficie",  label_visibility="collapsed", placeholder="Superficie")
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

        # ── Agregar nueva zona (zona + daños en un solo formulario) ─────────
        st.divider()
        st.markdown("**Agregar nueva zona**")
        with st.form("form_nueva_zona", clear_on_submit=True):
            c1, c2 = st.columns(2)
            nueva_zona = c1.text_input("Zona Afectada", placeholder="ej: Muro")
            nueva_ubic = c2.text_input("Ubicación",     placeholder="ej: Baño principal")

            st.caption("Daños de esta zona (agrega las filas que necesites):")
            items_df = st.data_editor(
                pd.DataFrame({
                    "Daño Observado": pd.Series([""], dtype=str),
                    "Superficie":     pd.Series([""], dtype=str),
                    "Unidad":         pd.Series(["m²"], dtype=str),
                }),
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Daño Observado": st.column_config.TextColumn(width="large"),
                    "Superficie":     st.column_config.TextColumn(width="small"),
                    "Unidad":         st.column_config.SelectboxColumn(
                        "Unidad", options=["m²", "m", "sin unidad"], default="m²", width="small"
                    ),
                },
                key=f"items_editor_{st.session_state.zona_form_v}",
            )

            if st.form_submit_button("+ Agregar zona", type="primary"):
                if nueva_zona.strip():
                    items = []
                    for _, row in items_df.iterrows():
                        dano = "" if pd.isna(row["Daño Observado"]) else str(row["Daño Observado"]).strip()
                        if dano and dano != "nan":
                            sup  = "" if pd.isna(row["Superficie"]) else str(row["Superficie"]).strip()
                            unit = "m²" if pd.isna(row["Unidad"]) else str(row["Unidad"])
                            items.append({"dano": dano, "superficie": sup, "unidad": unit})
                    grupos.append({"zona": nueva_zona.strip(), "ubicacion": nueva_ubic.strip(), "items": items})
                    st.session_state.danos_grupos = grupos
                    st.session_state.zona_form_v += 1
                    st.rerun()

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
                        detalle_visita = detalle_visita,
                        hay_inspeccion = hay_deteccion,
                        inspecciones   = st.session_state.detecciones,
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
