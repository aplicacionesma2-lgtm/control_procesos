import datetime
import io
import os
import uuid
import gspread
import pandas as pd
import streamlit as st

# Configuración de página
st.set_page_config(
    page_title="Control de Procesos - Líneas",
    page_icon="📋",
    layout="wide",
)

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1eQ64LwSp8cVm0T9o29KJgYqfF5e6yCLeN2RqmuY_ftc/edit?gid=0#gid=0"
FILE_PATH = "procesos.xlsx"

# --- CABECERAS ESTÁNDAR UNIFICADAS ---
HEADERS = [
    "ID_REGISTRO",
    "PRODUCTO",
    "LÍNEA DE PROCESO",
    "CONDICIONES DEL AREA DE TRABAJO",
    "F.P",
    "LOTE",
    "BATCH",
    "EQUIPO UTILIZADO",
    "HORA INICIO",
    "CONDICIONES DEL EQUIPO",
    "CONDICIONES DE LOS INSUMOS",
    "CARACTERISTICAS DEL PRODUCTO",
    "HORA TÉRMINO",
    "TIEMPO",
    "RESPONSABLE",
    "OBSERVACIÓN",
]

# CREDENCIALES POR DEFECTO DE RESPALDO
CRED_FALLBACK = {
    "type": "service_account",
    "project_id": "control-procesos-510021",
    "private_key_id": "aafa643ccb5578aabb4e3747a39c0fa95cac451a",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKkwggSkAgEAAoIBAQDBBJfKLxL38J+e\n7LS6lFB91EN+vbYs9VhrkjvIA3F7DdEKkW0zpcV0zY3BOQ0gadcDzc+Kf2VhHFYi\nVtub62RIm21bigdql/mMNgVliXnHfOOFQNPeTjdQYGMM6/v8xBemqhVubm2RS822\nOkwYaG41iy0av0l3ssqPPe6Os5UpEIGApDc5ERK8kNAAR7ZXGkNzkkp/KzYw/HCL\n0n9X5PBqjqcw4avCwbiGHgw00iZo7sqNAN2GlyuYqQaaQ8W6/aAwQFwgotscvocF\n7rexVHTIQS3iwnLYGscGf0B/qsz+UAzsym7xj1HDsdXs/Cuy9Eu+2ESsryV+wwM8\nUCTeXpGpAgMBAAECggEAHutPrmdijD2rMC/eVpoOF84CHuIgdey6ZIb5FRX6Hnps\n31rC6bhXHFoWKFrtf6D8vMMCCT9Vm9wIbzlHNh+bwabGOpjujbR5GO0JacW/MIXQ\nw4aKOe0BHtrF2yrNQ6Uc3cmWo8lEO3dvZU7K5EkMSH76M3PrfqVxHcePuKPLU9el\n58moIbxis4QXkrdE6eoSP1SFlLHixwFNfO4+6YA7RQbK/g3tTqu+AM8jN8SyuL1f\nrAI2WkVrS5Fj9TkLvU6ylQvF7W4GDd6oSSmZRovMTtJEU9zDJLFJuffLvLWWI6g2\nlVQdxI7vhY+YMXyQ/NYTOHfsZCcGdZU+jLhscps9dQKBgQDuNysdKrTXGRWiorpA\nmnW18cOr0BuZ5mWHeyHxDI3bzKqhWVzqMxw1s2MOLZnoTQ39phIWl0S+s3RcEmCO\nHf8rGpby6CGCHmu897rdp1qkCaU0udhNVn6ctD/NH6srk8aexGpaV0zYQezBQkUu\nMLIj+Nbh96B/SXxhai8bK7Q05QKBgQDPbZkSj3EjeCduSHUQzYYPu0ap0o9ubQHK\nRG9g5tPlDbgII0XJAJQT3S6c+ma9LkSowRDMLZOOXsRAJf5rbcxd/G7SELQ6NDDa\njmv5r+uQ4JaFBFW9OZDwIFTIVp7nq2eSas6zZL91RSVXKa5mufh2Y8LBjpkWG7pG\nWn2fqj+BdQKBgQCj8uJAa7EUzVXfniGT7vqOo3spF8y3SiOcb/l3Pk2v9heFfsx8\n/3ot122YR3hCsi2r4g1W8PtGSJoP+DHt/eUtlFpJicvuEuPRpao9fT3b4iuKs1GU\nQLBZR5EVqvMSxd0QTlxoGudve0fn5qVYWflw2oWB9fzHPhtVrFAJYjXfpQKBgQC0\ncV/uwG+oblbG3itQQam0t7KB+tShOByNizjksAh2wpdsJNsJPwKRwSBSmJWVTtGV\nh9YH+EHbYN8R+rs3Ux2sSPNStAtEcrBo/+o4G+wtbOIjtqCrao+GBGocmRXE7Nu9\niEJl1mejKVKRX4YCgRb+TkxWuqi7jcVefEu6AI0cHQKBgHP0FKxE6bPgYPceXtNV\nVICr1sWxBEamX72kkAZ930B+F6t6wDwSJncfgamlQEpnuSSd3rE8D/K+hm4q4A2d\nTIMDFdBn8eybpXb9mDC9ADdzyz6yF0mPEvKgG+q/De+oPwjqrDBDI7fRrwSUqaaV\ntTOhmvYbJp7/aJD6BPwwFjmn\n-----END PRIVATE KEY-----\n",
    "client_email": "control-procesos@control-procesos-510021.iam.gserviceaccount.com",
    "client_id": "110338130362201619109",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/control-procesos%40control-procesos-510021.iam.gserviceaccount.com",
}

# --- CONEXIÓN ROBUSTA A GOOGLE SHEETS ---
@st.cache_resource
def conectar_google_sheets():
    """Intenta conectar con st.secrets de forma segura y cae en el diccionario local si no existe."""
    cred_dict = None
    
    # Manejo seguro del acceso a st.secrets para evitar el crash/KeyError de Streamlit
    try:
        if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
            cred_dict = dict(st.secrets["gcp_service_account"])
    except Exception:
        cred_dict = None

    # Fallback si secrets.toml no tiene la clave configurada
    if not cred_dict:
        cred_dict = CRED_FALLBACK

    try:
        client = gspread.service_account_from_dict(cred_dict)
        sheet = client.open_by_url(SPREADSHEET_URL).sheet1
        return sheet
    except Exception as e:
        st.error(f"❌ Error al conectar con Google Sheets: {e}")
        return None


ws = conectar_google_sheets()

if ws:
    try:
        data = ws.get_all_records()
        df_actual = pd.DataFrame(data)
    except Exception as e:
        st.error(f"❌ Error al obtener datos de Google Sheets: {e}")
        df_actual = pd.DataFrame(columns=HEADERS)
else:
    df_actual = pd.DataFrame(columns=HEADERS)

CATALOGO_DEFAULT = {
    "MUFFINS": [
        "MUFFIN DE MANZANA - FRESCO - (UND)",
        "MUFFIN DE NARANJA & CHOCOCHIPS - FRESCO - (UND)",
        "MUFFIN DE BERRIES - FRESCO - (UND)",
        "MUFFIN DE CHOCOLATE - FRESCO - (UND)",
        "MUFFIN DE RED VELVET (SIN DECORAR)",
    ],
    "AMASADO": [
        "GALLETA CHOCOCHIP CON AVELLANA STBX",
        "CINNAMON ROLL - STB",
        "BASE DE TARTALETA CIRCULAR",
        "CROUMBLE BERRIES HORNEADO",
        "BASE CROUMBLE DE BERRIES STBX",
        "CROUMBLE BERRY KG STBX",
    ],
    "DECORADO": [
        "ROLL CINNAMON STARBUCKS",
        "TORTA DE CHOCOLATE CON FUDGE Y MANJAR",
        "PYE DE LIMON - FRESCO - (UND) (STBX)",
        "KEKE RECTANGULAR DE LIMON  STBX",
        "KEKE RECTAGULAR DE ZANAHORIA STBX",
        "KEKE RECTANGULAR GINGER DECORADO - B2B",
        "MUFFIN DE RED VELVET - UND",
        "GALLETA CHOCOCHIP CON AVELLANA STBX",
    ],
    "BATIDOS": [
        "KEKE RECTANGULAR DE LIMON STB",
        "KEKE DE ZANAHORIA INTEGRAL Y PANELA STB",
    ],
    "OTRO": [],
}

EQUIPOS_DEFAULT = [
    "BATIDORA",
    "ROBOTCOUPE",
    "AMASADORA",
    "GALLETERA",
    "DIVISORA",
    "DOSIFICADORA",
    "LAMINADORA",
    "NO APLICA",
    "OTRO",
]


@st.cache_data(ttl=3600)
def cargar_catalogos():
    mapa_linea_productos = CATALOGO_DEFAULT.copy()
    equipos = EQUIPOS_DEFAULT
    if os.path.exists(FILE_PATH):
        try:
            xls = pd.ExcelFile(FILE_PATH)
            for sheet in xls.sheet_names:
                df = pd.read_excel(FILE_PATH, sheet_name=sheet)
                for idx, row in df.iterrows():
                    row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
                    if "PRODUCTO" in row_vals and "LÍNEA DE PROCESO" in row_vals:
                        header_idx = idx
                        df_data = df.iloc[header_idx + 1 :].copy()
                        df_data.columns = [str(c).strip() for c in df.iloc[header_idx].values]
                        temp_map = {}
                        for _, r in df_data.iterrows():
                            p = str(r["PRODUCTO"]).strip() if pd.notna(r.get("PRODUCTO")) else None
                            l = str(r["LÍNEA DE PROCESO"]).strip() if pd.notna(r.get("LÍNEA DE PROCESO")) else None
                            if p and l:
                                if l not in temp_map:
                                    temp_map[l] = []
                                if p not in temp_map[l]:
                                    temp_map[l].append(p)
                        if temp_map:
                            if "OTRO" not in temp_map:
                                temp_map["OTRO"] = []
                            mapa_linea_productos = temp_map
                        if "EQUIPO UTILIZADO" in df_data.columns:
                            eqs = (
                                df_data["EQUIPO UTILIZADO"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .unique()
                                .tolist()
                            )
                            if eqs:
                                equipos = list(dict.fromkeys(eqs + ["OTRO"]))
                        break
        except Exception:
            pass
    return mapa_linea_productos, equipos


mapa_linea_productos, lista_equipos = cargar_catalogos()
lista_lineas = list(mapa_linea_productos.keys())

# --- NAVEGACIÓN PRINCIPAL ---
st.title("📋 FORMATO CONTROL DE PROCESOS LÍNEAS")
tab1, tab2 = st.tabs(["📝 Nuevo Registro", "✏️ Gestionar / Editar / Eliminar Historial"])

# ==========================================
# PESTAÑA 1: NUEVO REGISTRO
# ==========================================
with tab1:
    st.subheader("1. Selección de Línea, Producto y Equipo")
    col_a, col_b, col_c = st.columns(3)

    with col_a:
        linea_sel = st.selectbox("LÍNEA DE PROCESO", options=lista_lineas, key="select_linea_filtro")
        if linea_sel == "OTRO":
            linea_text = st.text_input("Especifique Línea de Proceso", placeholder="Nombre de línea")
            linea_final = linea_text
            productos_disponibles = []
        else:
            linea_final = linea_sel
            productos_disponibles = mapa_linea_productos.get(linea_sel, [])

    with col_b:
        if productos_disponibles:
            opciones_producto = productos_disponibles + ["OTRO"]
            prod_sel = st.selectbox("PRODUCTO", options=opciones_producto, key="select_producto_dinamico")
            if prod_sel == "OTRO":
                prod_text = st.text_input("Especifique Producto", placeholder="Nombre del producto")
                producto_final = prod_text
            else:
                producto_final = prod_sel
        else:
            producto_final = st.text_input("PRODUCTO", placeholder="Escriba el nombre del producto")

    with col_c:
        equipo_sel = st.selectbox("EQUIPO UTILIZADO", options=lista_equipos)
        if equipo_sel == "OTRO":
            equipo_text = st.text_input("Especifique Equipo", placeholder="Nombre del equipo")
            equipo_final = equipo_text
        else:
            equipo_final = equipo_sel

    with st.form("form_control_proceso", clear_on_submit=True):
        st.subheader("2. Información General del Proceso")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            fecha_p = st.date_input("F.P (Fecha de Producción)", value=datetime.date.today())
        with col2:
            lote = st.text_input("LOTE", placeholder="Ej. L-20260901", key="input_lote")
        with col3:
            batch = st.text_input("BATCH", placeholder="Ej. B-01", key="input_batch")
        with col4:
            responsable = st.text_input(
                "RESPONSABLE",
                placeholder="Ingrese el nombre del responsable",
                key="input_responsable",
            )

        st.subheader("3. Condiciones del Entorno e Insumos")
        col8, col9, col10, col11 = st.columns(4)
        with col8:
            cond_area = st.radio("CONDICIONES AREA TRABAJO", ["CONFORME", "NO CONFORME"], horizontal=True)
        with col9:
            cond_equipo = st.radio("CONDICIONES EQUIPO", ["CONFORME", "NO CONFORME"], horizontal=True)
        with col10:
            cond_insumos = st.radio("CONDICIONES INSUMOS", ["CONFORME", "NO CONFORME"], horizontal=True)
        with col11:
            caract_producto = st.radio("CARACTERISTICAS PRODUCTO", ["CONFORME", "NO CONFORME"], horizontal=True)

        st.subheader("4. Horarios y Tiempo de Proceso")
        col12, col13, col14 = st.columns(3)
        with col12:
            hora_inicio = st.time_input("HORA INICIO", value=datetime.time(8, 0))
        with col13:
            hora_termino = st.time_input("HORA TÉRMINO", value=datetime.time(8, 15))

        dt_inicio = datetime.datetime.combine(datetime.date.today(), hora_inicio)
        dt_termino = datetime.datetime.combine(datetime.date.today(), hora_termino)
        if dt_termino < dt_inicio:
            dt_termino += datetime.timedelta(days=1)
        minutos_totales = int((dt_termino - dt_inicio).total_seconds() / 60)
        tiempo_calculado = f"{minutos_totales} min"

        with col14:
            st.text_input("TIEMPO CALCULADO", value=tiempo_calculado, disabled=True)

        st.subheader("5. Observaciones Finales")
        col15, col16 = st.columns([1, 2])
        with col15:
            estado_obs = st.radio("OBSERVACIÓN", options=["CONFORME", "NO CONFORME", "OTRO (Texto Libre)"])
        with col16:
            if estado_obs == "OTRO (Texto Libre)":
                observacion_final = st.text_area("Detalle", placeholder="Escriba observación...")
            else:
                obs_adic = st.text_input("Comentario Adicional (Opcional)", placeholder="Detalles...")
                observacion_final = f"{estado_obs} - {obs_adic}".strip(" -") if obs_adic else estado_obs

        btn_guardar = st.form_submit_button("💾 Guardar Registro en Google Sheets", use_container_width=True)

    if btn_guardar:
        if not ws:
            st.error("❌ No hay conexión activa con Google Sheets.")
        else:
            registro_id = str(uuid.uuid4())[:8]

            fila_nueva = [
                registro_id,
                producto_final,
                linea_final,
                cond_area,
                fecha_p.strftime("%Y-%m-%d"),
                lote,
                batch,
                equipo_final,
                hora_inicio.strftime("%H:%M"),
                cond_equipo,
                cond_insumos,
                caract_producto,
                hora_termino.strftime("%H:%M"),
                tiempo_calculado,
                responsable,
                observacion_final,
            ]
            try:
                datos_existentes = ws.get_all_values()
                if not datos_existentes:
                    ws.append_row(HEADERS)
                else:
                    ws.update([HEADERS], range_name="A1")
                ws.append_row(fila_nueva)
                st.success("✅ ¡Registro guardado exitosamente en Google Sheets!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al guardar: {e}")

# ==========================================
# PESTAÑA 2: GESTIONAR, FILTRAR, EDITAR Y ELIMINAR
# ==========================================
with tab2:
    st.subheader("🔍 Filtrar, Editar y Eliminar Registros por Fecha")

    if not ws:
        st.error("❌ Conexión con Google Sheets no disponible.")
    else:
        try:
            rows_all = ws.get_all_values()
            if len(rows_all) > 1:
                head = rows_all[0]
                body = rows_all[1:]
                df_cloud = pd.DataFrame(body, columns=head)

                if "F.P" in df_cloud.columns:
                    fechas_disponibles = sorted(df_cloud["F.P"].dropna().unique().tolist())
                    if fechas_disponibles:
                        fecha_seleccionada = st.selectbox(
                            "📅 Seleccione la Fecha de Producción (F.P) a gestionar:",
                            options=fechas_disponibles,
                        )

                        df_filtrado = df_cloud[df_cloud["F.P"] == fecha_seleccionada]
                        st.markdown(f"### Registros encontrados para la fecha: `{fecha_seleccionada}`")

                        for local_idx, row in df_filtrado.iterrows():
                            sheet_row_num = local_idx + 2

                            with st.expander(
                                f"📦 Producto: {row.get('PRODUCTO')} | Lote: {row.get('LOTE')} | Resp: {row.get('RESPONSABLE')}"
                            ):
                                with st.form(key=f"form_edit_{sheet_row_num}_{row.get('ID_REGISTRO', local_idx)}"):
                                    st.write(f"Editando registro (Fila en Google Sheets: {sheet_row_num})")

                                    col_e1, col_e2, col_e3 = st.columns(3)
                                    with col_e1:
                                        nuevo_prod = st.text_input("PRODUCTO", value=row.get("PRODUCTO", ""))
                                        nueva_linea = st.text_input("LÍNEA DE PROCESO", value=row.get("LÍNEA DE PROCESO", ""))
                                        nuevo_lote = st.text_input("LOTE", value=row.get("LOTE", ""))
                                    with col_e2:
                                        nuevo_batch = st.text_input("BATCH", value=row.get("BATCH", ""))
                                        nuevo_resp = st.text_input("RESPONSABLE", value=row.get("RESPONSABLE", ""))
                                        nuevo_equipo = st.text_input("EQUIPO UTILIZADO", value=row.get("EQUIPO UTILIZADO", ""))
                                    with col_e3:
                                        nuevo_inicio = st.text_input("HORA INICIO", value=row.get("HORA INICIO", ""))
                                        nuevo_termino = st.text_input("HORA TÉRMINO", value=row.get("HORA TÉRMINO", ""))
                                        nueva_obs = st.text_area("OBSERVACIÓN", value=row.get("OBSERVACIÓN", ""))

                                    btn_actualizar = st.form_submit_button("💾 Guardar Cambios de este Registro")

                                    if btn_actualizar:
                                        try:
                                            fila_actualizada = [
                                                row.get("ID_REGISTRO", str(uuid.uuid4())[:8]),
                                                nuevo_prod,
                                                nueva_linea,
                                                row.get("CONDICIONES DEL AREA DE TRABAJO", "CONFORME"),
                                                row.get("F.P", fecha_seleccionada),
                                                nuevo_lote,
                                                nuevo_batch,
                                                nuevo_equipo,
                                                nuevo_inicio,
                                                row.get("CONDICIONES DEL EQUIPO", "CONFORME"),
                                                row.get("CONDICIONES DE LOS INSUMOS", "CONFORME"),
                                                row.get("CARACTERISTICAS DEL PRODUCTO", "CONFORME"),
                                                nuevo_termino,
                                                row.get("TIEMPO", "15 min"),
                                                nuevo_resp,
                                                nueva_obs,
                                            ]
                                            ws.update(
                                                range_name=f"A{sheet_row_num}:P{sheet_row_num}",
                                                values=[fila_actualizada],
                                            )
                                            st.success("✅ ¡Registro actualizado correctamente en Google Sheets!")
                                            st.rerun()
                                        except Exception as err:
                                            st.error(f"Error al actualizar: {err}")

                                if st.button(
                                    "🗑️ Eliminar este registro permanentemente",
                                    key=f"del_{sheet_row_num}_{row.get('ID_REGISTRO', local_idx)}",
                                ):
                                    try:
                                        ws.delete_rows(sheet_row_num)
                                        st.success("✅ ¡Registro eliminado de Google Sheets!")
                                        st.rerun()
                                    except Exception as err:
                                        st.error(f"Error al eliminar: {err}")
                    else:
                        st.info("No hay fechas registradas en la columna F.P.")
                else:
                    st.warning("No se encontró la columna 'F.P' en la hoja de cálculo.")
            else:
                st.info("Aún no hay registros guardados en Google Sheets.")
        except Exception as e:
            st.error(f"Error al leer registros: {e}")

    st.markdown("---")
    st.subheader("📥 Exportar Historial Completo")
    if ws:
        try:
            registros_totales = ws.get_all_records()
            if registros_totales:
                df_registros = pd.DataFrame(registros_totales)
                col_down1, col_down2 = st.columns(2)

                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
                    df_registros.to_excel(writer, index=False, sheet_name="Control_Procesos")
                data_excel = buffer_excel.getvalue()

                with col_down1:
                    st.download_button(
                        label="📊 Descargar en Excel (.xlsx)",
                        data=data_excel,
                        file_name=f"CONTROL_PROCESOS_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                data_csv = df_registros.to_csv(index=False).encode("utf-8")
                with col_down2:
                    st.download_button(
                        label="📄 Descargar en CSV (.csv)",
                        data=data_csv,
                        file_name=f"CONTROL_PROCESOS_{datetime.date.today()}.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )
        except Exception:
            pass
