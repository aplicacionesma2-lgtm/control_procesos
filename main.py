import datetime
import io
import os
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


# --- CONEXIÓN SEGURA A GOOGLE SHEETS ---
@st.cache_resource
def conectar_google_sheets():
  pk_lines = [
      "-----BEGIN PRIVATE KEY-----",
      "MIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQChFFBYCHDbTJt4",
      "QIJVeEpflh0wWXFspNZg2Y8v9CukuWGq2O+28MzxM5O+Pv6ls1bThUGnOAD85Fxz",
      "DBPUm9BqifGRugMNBcE876Af3YuG5s+tDMxF+kWFyQEVVpvZOCK6/WWHOS74F4LN",
      "WSgcJCWAlMqXlulKGgvvzGLW7QqSW0viJ+DesQd5kNT766qvVO0WyIFzh87N9z70",
      "lwjzNIiBov/7XgKaagawNjaVm5Ob/2AGzU/82VMfB/DZoJhDEpcJ6qmHLgcKVqZT",
      "RHLJZshWO4H8Kqxqfmvou6qziydJdBzTrFiqZ359bsUThHpl23mT5OgE4A0xYNY8",
      "Cxr6Nt+FAgMBAAECggEAHNl4TVgPrHtTQg2dukSd33JRnoH6gligi76TemV7Npi8",
      "QR7zChslPZL9BF8geRl+dMpiWJp7dM/KrhFM8PCKOrajlTPRTZEJC8qoLWTe00W+",
      "DtKiuGrLbltyjfmR1q0K7V4qg7ZObwE4/GHaQPYJYHbltRJCjLLPDf8XecKBOPaH",
      "ulp7EErv7WetyxJC2tqb/cZ2RiSu81ixTMhJkm6H4TnjREPj3FigR3kKjGutbdIS",
      "4BmGX1MoX7d1ehYA06vt7Xj6uiPuLRxg/AfpZje9NJiH3W+JhrarGqalY2Wwnxe7",
      "xaA5NJlRQ/pZUkbItVBaOwZc9cmF7hLXq6DKFnCnOQKBgQDTtIXUOBmuHOcC3LTq",
      "1I6p8voTTofMoT1MSS4OD43hQOFMyiGWs0xpSmqJVmAtckKQmA9P6FdgmV93C8yO",
      "Fza3m5o8QsfiamCYGTsoLFC6THdqqic2eUId3DCS79ln3M6HvY24oEj/aWgD5lkY",
      "4b5azGsGce6YO7ZuKne9VWoiDQKBgQDCyCQl0EDxrkjs+FHuhRxyn2jyh10q8T3z",
      "GGzhugeleazB0Wp7Pa9paVJkGkaKs0KgZUX8cO08t4e+5wkcWgDRE56Ss443qf98",
      "fpOOqzdLfy4l/QF+QeJ1DJZCONhKgBwjjV2h5EzVqxRq3yRY+Enn8jdVc9p7FKbU",
      "bqD3DjftWQKBgBWScYicZtF9FHUQNEcxfZAHuD+7Ys8RJwPc+Rppr1VinRKMDjwi",
      "7QhVkuGHsakv2WSOehD0ZeLr/fRNeXyJFQREkMTPMTr7B/i3qXWAfoFdRVXTHMfK",
      "N1h/lVuDoS2aLFlckVJc0tNj1DuBf1avugvahJVVirBsdTxoi2b5iyUJAoGABymo",
      "+qL/4GNSVzSCfszyUNy/1TtZF70rVAcv6dUXduRUkAQNcF7CVpQC7Z9xvKP+7TsM",
      "Kc5VSwhMu55vXVWJ9iZMjISB1FYyCPf2oSZ2sBYLMmZtaaEunLNLyz+f5I11e3E1",
      "YkCs+qaB57Qw9/yZaygjFMdf32rQ/7rZvHwPXnECgYB2InoDsMB1KT+83SL6r3tF",
      "hY7EePnosi2XWmR4ialkcinN6CsxS5jZ7ZISxPb2fBgYDFQVoNhkq1aZA6au5S4T",
      "UWTDUQLOBnTurhCFQ/+nElZAlX4nF4M6wKUTlUEQsy7J2cICbSA4tEvx9DEsy2m7",
      "40bXZzRQtzP8ciNAuypvAg==",
      "-----END PRIVATE KEY-----",
  ]
  private_key_str = "\n".join(pk_lines)

  cred_dict = {
      "type": "service_account",
      "project_id": "control-procesos-510021",
      "private_key_id": "4ba35503d5be0db7e99c9da752bb173efee0c609",
      "private_key": private_key_str,
      "client_email": (
          "control-procesos@control-procesos-510021.iam.gserviceaccount.com"
      ),
      "client_id": "110338130362201619109",
      "auth_uri": "https://accounts.google.com/o/oauth2/auth",
      "token_uri": "https://oauth2.googleapis.com/token",
      "auth_provider_x509_cert_url": (
          "https://www.googleapis.com/oauth2/v1/certs"
      ),
      "client_x509_cert_url": (
          "https://www.googleapis.com/robot/v1/metadata/x509/control-procesos%40control-procesos-510021.iam.gserviceaccount.com"
      ),
      "universe_domain": "googleapis.com",
  }

  client = gspread.service_account_from_dict(cred_dict)
  sheet = client.open_by_url(SPREADSHEET_URL).sheet1
  return sheet


try:
  ws = conectar_google_sheets()
  data = ws.get_all_records()
  df_actual = pd.DataFrame(data)
except Exception as e:
  st.error(f"❌ Error al conectar con Google Sheets: {e}")
  df_actual = pd.DataFrame()

# --- CABECERAS ESTÁNDAR UNIFICADAS ---
headers = [
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

if df_actual.empty:
  df_actual = pd.DataFrame(columns=headers)

# Catálogo por defecto (Línea -> Productos)
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
              p = (
                  str(r["PRODUCTO"]).strip()
                  if pd.notna(r.get("PRODUCTO"))
                  else None
              )
              l = (
                  str(r["LÍNEA DE PROCESO"]).strip()
                  if pd.notna(r.get("LÍNEA DE PROCESO"))
                  else None
              )
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
tab1, tab2 = st.tabs(
    ["📝 Nuevo Registro", "✏️ Gestionar / Editar / Eliminar Historial"]
)

# ==========================================
# PESTAÑA 1: NUEVO REGISTRO (EN CASCADA)
# ==========================================
with tab1:
  st.subheader("1. Selección de Línea, Producto y Equipo")
  col_a, col_b, col_c = st.columns(3)

  with col_a:
    linea_sel = st.selectbox(
        "LÍNEA DE PROCESO", options=lista_lineas, key="select_linea_filtro"
    )
    if linea_sel == "OTRO":
      linea_text = st.text_input(
          "Especifique Línea de Proceso", placeholder="Nombre de línea"
      )
      linea_final = linea_text
      productos_disponibles = []
    else:
      linea_final = linea_sel
      productos_disponibles = mapa_linea_productos.get(linea_sel, [])

  with col_b:
    if productos_disponibles:
      opciones_producto = productos_disponibles + ["OTRO"]
      prod_sel = st.selectbox(
          "PRODUCTO", options=opciones_producto, key="select_producto_dinamico"
      )
      if prod_sel == "OTRO":
        prod_text = st.text_input(
            "Especifique Producto", placeholder="Nombre del producto"
        )
        producto_final = prod_text
      else:
        producto_final = prod_sel
    else:
      producto_final = st.text_input(
          "PRODUCTO", placeholder="Escriba el nombre del producto"
      )

  with col_c:
    equipo_sel = st.selectbox("EQUIPO UTILIZADO", options=lista_equipos)
    if equipo_sel == "OTRO":
      equipo_text = st.text_input(
          "Especifique Equipo", placeholder="Nombre del equipo"
      )
      equipo_final = equipo_text
    else:
      equipo_final = equipo_sel

  with st.form("form_control_proceso", clear_on_submit=True):
    st.subheader("2. Información General del Proceso")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
      fecha_p = st.date_input(
          "F.P (Fecha de Producción)", value=datetime.date.today()
      )
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
      cond_area = st.radio(
          "CONDICIONES DEL AREA DE TRABAJO",
          options=["CONFORME", "NO CONFORME"],
          horizontal=True,
      )
    with col9:
      cond_equipo = st.radio(
          "CONDICIONES DEL EQUIPO",
          options=["CONFORME", "NO CONFORME"],
          horizontal=True,
      )
    with col10:
      cond_insumos = st.radio(
          "CONDICIONES DE LOS INSUMOS",
          options=["CONFORME", "NO CONFORME"],
          horizontal=True,
      )
    with col11:
      caract_producto = st.radio(
          "CARACTERISTICAS DEL PRODUCTO",
          options=["CONFORME", "NO CONFORME"],
          horizontal=True,
      )

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
      estado_obs = st.radio(
          "OBSERVACIÓN", options=["CONFORME", "NO CONFORME", "OTRO (Texto Libre)"]
      )

    with col16:
      if estado_obs == "OTRO (Texto Libre)":
        obs_detalle = st.text_area(
            "Detalle de Observación",
            placeholder="Escriba aquí la observación personalizada...",
            key="input_obs_detalle",
        )
        observacion_final = obs_detalle
      else:
        obs_adicional = st.text_input(
            "Comentario Adicional (Opcional)",
            placeholder="Escriba detalles si aplica...",
            key="input_obs_adic",
        )
        observacion_final = (
            f"{estado_obs} - {obs_adicional}".strip(" -")
            if obs_adicional
            else estado_obs
        )

    st.markdown("---")
    btn_guardar = st.form_submit_button(
        "💾 Guardar Registro en Google Sheets", use_container_width=True
    )

  if btn_guardar:
    # ORDEN EXACTO QUE COINCIDE CON LAS COLUMNAS DE LA HOJA
    fila_nueva = [
        producto_final,  # A: PRODUCTO
        linea_final,  # B: LÍNEA DE PROCESO
        cond_area,  # C: CONDICIONES DEL AREA DE TRABAJO
        fecha_p.strftime("%Y-%m-%d"),  # D: F.P
        lote,  # E: LOTE
        batch,  # F: BATCH
        equipo_final,  # G: EQUIPO UTILIZADO
        hora_inicio.strftime("%H:%M"),  # H: HORA INICIO
        cond_equipo,  # I: CONDICIONES DEL EQUIPO
        cond_insumos,  # J: CONDICIONES DE LOS INSUMOS
        caract_producto,  # K: CARACTERISTICAS DEL PRODUCTO
        hora_termino.strftime("%H:%M"),  # L: HORA TÉRMINO
        tiempo_calculado,  # M: TIEMPO
        responsable,  # N: RESPONSABLE
        observacion_final,  # O: OBSERVACIÓN
    ]

    try:
      datos_existentes = ws.get_all_values()
      if not datos_existentes:
        ws.append_row(headers)
      else:
        # Aseguramos que la primera fila tenga las cabeceras exactas
        ws.update([headers], range_name="A1")

      ws.append_row(fila_nueva)
      st.success("✅ ¡Registro guardado exitosamente en Google Sheets!")
      st.rerun()
    except Exception as e:
      st.error(f"❌ Error al guardar en Google Sheets: {e}")

# ==========================================
# PESTAÑA 2: EDITAR Y ELIMINAR HISTORIAL
# ==========================================
with tab2:
  st.subheader("📊 Edición, Eliminación y Descarga de Registros")
  st.info(
      "💡 **Instrucciones:** Modifica celdas haciendo doble clic sobre ellas o"
      " elimina filas seleccionándolas y presionando la tecla 'Supr/Delete' o"
      " la papelera. Luego presiona **'Guardar Cambios en Google Sheets'**."
  )

  try:
    registros = ws.get_all_records()
    if registros:
      df_registros = pd.DataFrame(registros)
      df_editado = st.data_editor(
          df_registros,
          num_rows="dynamic",
          use_container_width=True,
          key="editor_historial",
      )

      col_save, _ = st.columns([1, 2])
      with col_save:
        if st.button(
            "💾 Guardar Cambios en Google Sheets", use_container_width=True
        ):
          try:
            ws.clear()
            # Forzar la reescritura de cabeceras estándar + datos limpios
            if not df_editado.empty:
              df_limpio = df_editado.fillna("")
              ws.update(
                  [headers] + df_limpio.values.tolist(), range_name="A1"
              )
            else:
              ws.update([headers], range_name="A1")

            st.success(
                "✅ ¡Google Sheets actualizado y registros eliminados con"
                " éxito!"
            )
            st.rerun()
          except Exception as e:
            st.error(f"❌ Error al actualizar: {e}")

      st.markdown("---")
      st.subheader("📥 Exportar Historial Completo")
      col_down1, col_down2 = st.columns(2)

      buffer_excel = io.BytesIO()
      with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
        df_registros.to_excel(
            writer, index=False, sheet_name="Control_Procesos"
        )
      data_excel = buffer_excel.getvalue()

      with col_down1:
        st.download_button(
            label="📊 Descargar Historial en Excel (.xlsx)",
            data=data_excel,
            file_name=f"CONTROL_PROCESOS_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

      data_csv = df_registros.to_csv(index=False).encode("utf-8")
      with col_down2:
        st.download_button(
            label="📄 Descargar Historial en CSV (.csv)",
            data=data_csv,
            file_name=f"CONTROL_PROCESOS_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    else:
      st.info("Aún no hay registros guardados en Google Sheets.")
  except Exception as e:
    st.error(f"Error al leer desde Google Sheets: {e}")
