from datetime import datetime
import gspread
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Control de Producción y Trazabilidad",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Registro y Control de Producción - Pastelería")

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1eQ64LwSp8cVm0T9o29KJgYqfF5e6yCLeN2RqmuY_ftc/edit?gid=0#gid=0"
FILE_PATH = "procesos.xlsx"


# --- CONEXIÓN A GOOGLE SHEETS (Con tu clave privada interna y segura) ---
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
  st.error(f"Error al conectar con Google Sheets: {e}")
  df_actual = pd.DataFrame()

# --- ENCABEZADOS ESTÁNDAR ---
headers = [
    "PRODUCTO",
    "LÍNEA DE PROCESO",
    "F.P",
    "LOTE",
    "BATCH",
    "RESPONSABLE",
    "CONDICIONES DEL AREA DE TRABAJO",
    "CONDICIONES DEL EQUIPO",
    "CONDICIONES DE LOS INSUMOS",
    "CARACTERISTICAS DEL PRODUCTO",
    "HORA INICIO",
    "HORA TÉRMINO",
    "TIEMPO",
    "EQUIPO UTILIZADO",
    "OBSERVACIÓN",
]

if df_actual.empty:
  df_actual = pd.DataFrame(columns=headers)

# --- FORMULARIO PARA NUEVOS REGISTROS ---
with st.form("form_produccion", clear_on_submit=True):
  st.subheader("Registrar Nuevo Proceso")

  col1, col2, col3 = st.columns(3)

  with col1:
    producto_final = st.selectbox(
        "Producto",
        [
            "MUFFIN DE MANZANA - FRESCO - (UND)",
            "MUFFIN DE NARANJA & CHOCOCHIPS - FRESCO - (UND)",
            "MUFFIN DE BERRIES - FRESCO - (UND)",
            "MUFFIN DE CHOCOLATE - FRESCO - (UND)",
        ],
    )
    linea_final = st.selectbox("Línea de Proceso", ["MUFFINS", "AMASADO"])
    fecha_p = st.date_input("Fecha de Producción (F.P)", datetime.now())

  with col2:
    lote = st.text_input("Lote")
    batch = st.text_input("Batch")
    responsable = st.text_input("Responsable")
    equipo_final = st.text_input("Equipo Utilizado", value="BATIDORA")

  with col3:
    cond_area = st.selectbox(
        "Condiciones Área de Trabajo", ["CONFORME", "NO CONFORME"]
    )
    cond_equipo = st.selectbox("Condiciones del Equipo", ["CONFORME", "NO CONFORME"])
    cond_insumos = st.selectbox("Condiciones de Insumos", ["CONFORME", "NO CONFORME"])
    caract_producto = st.selectbox(
        "Características del Producto", ["CONFORME", "NO CONFORME"]
    )

  col4, col5, col6 = st.columns(3)
  with col4:
    hora_inicio = st.time_input("Hora Inicio", datetime.now().time())
  with col5:
    hora_termino = st.time_input("Hora Término", datetime.now().time())
  with col6:
    tiempo_calculado = st.text_input("Tiempo (ej. 15 min)", value="15 min")

  observacion_final = st.text_area("Observación", value="CONFORME")

  submit_button = st.form_submit_button("Agregar Registro a la Tabla")

  if submit_button:
    nueva_fila = {
        "PRODUCTO": producto_final,
        "LÍNEA DE PROCESO": linea_final,
        "F.P": fecha_p.strftime("%Y-%m-%d"),
        "LOTE": lote,
        "BATCH": batch,
        "RESPONSABLE": responsable,
        "CONDICIONES DEL AREA DE TRABAJO": cond_area,
        "CONDICIONES DEL EQUIPO": cond_equipo,
        "CONDICIONES DE LOS INSUMOS": cond_insumos,
        "CARACTERISTICAS DEL PRODUCTO": caract_producto,
        "HORA INICIO": hora_inicio.strftime("%H:%M"),
        "HORA TÉRMINO": hora_termino.strftime("%H:%M"),
        "TIEMPO": tiempo_calculado,
        "EQUIPO UTILIZADO": equipo_final,
        "OBSERVACIÓN": observacion_final,
    }

    df_actual = pd.concat(
        [df_actual, pd.DataFrame([nueva_fila])], ignore_index=True
    )
    st.success(
        "¡Registro agregado! Haz clic en 'Guardar Cambios' abajo para"
        " sincronizar."
    )

# --- TABLA EDITABLE (PERMITE BORRAR Y EDITAR) ---
st.subheader("Edición y Gestión de Registros (Puedes eliminar filas aquí)")
st.info(
    "💡 Para borrar una fila: selecciónala haciendo clic en su casilla izquierda"
    " y presiona el ícono de la papelera 🗑️ en la esquina superior derecha de"
    " la tabla."
)

edited_df = st.data_editor(
    df_actual,
    num_rows="dynamic",
    key="editor_registros",
    use_container_width=True,
)

# --- BOTÓN PARA SINCRONIZAR CON GOOGLE SHEETS ---
if st.button("💾 Guardar Cambios y Actualizar Google Sheets"):
  try:
    ws.clear()

    if not edited_df.empty:
      df_limpio = edited_df.fillna("")
      data_to_update = [df_limpio.columns.tolist()] + df_limpio.values.tolist()
      ws.update(data_to_update)
    else:
      ws.update([headers])

    st.success(
        "¡Google Sheets actualizado correctamente! Los registros eliminados"
        " fueron borrados de la nube."
    )
    st.rerun()

  except Exception as e:
    st.error(f"Error al sincronizar con Google Sheets: {e}")
