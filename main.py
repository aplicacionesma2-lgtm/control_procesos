from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Control de Producción y Trazabilidad",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Registro y Control de Producción - Pastelería")

# --- CONEXIÓN A GOOGLE SHEETS ---
try:
  # Configura tus credenciales (asegúrate de tener tu secret o archivo json configurado)
  scope = [
      "https://spreadsheets.google.com/feeds",
      "https://www.googleapis.com/auth/drive",
  ]

  # Si usas st.secrets para Streamlit Cloud:
  creds_dict = dict(st.secrets["gcp_service_account"])
  creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

  client = gspread.authorize(creds)

  # Abre tu hoja de cálculo (reemplaza con el nombre exacto de tu archivo)
  spreadsheet = client.open("REGISTRO_PRODUCCION")
  hoja = spreadsheet.sheet1  # O usa spreadsheet.worksheet("NombreDePestaña")

  # Leer datos actuales
  data = hoja.get_all_records()
  df_actual = pd.DataFrame(data)

except Exception as e:
  st.error(
      f"Error al conectar con Google Sheets: {e}. Verifica tus credenciales o"
      " el nombre del archivo."
  )
  df_actual = pd.DataFrame()

# --- DEFECTO DE ENCABEZADOS ESTÁNDAR ---
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

# Si la hoja está totalmente vacía, inicializamos el DataFrame con las columnas correctas
if df_actual.empty:
  df_actual = pd.DataFrame(columns=headers)

# --- FORMULARIO PARA NUEVOS REGISTROS ---
with st.form("form_produccion", clear_on_submit=True):
  st.subheader("Registrar Nuevo Proceso")

  col1, col2, col3 = st.columns(3)

  with col1:
    producto_final = st.selectbox(
        "Producto",
        ["MUFFIN DE MANZANA", "MUFFIN DE NARANJA", "BROWNIE", "ALFAJOR"],
    )
    linea_final = st.selectbox("Línea de Proceso", ["MUFFINS", "GALLETAS"])
    fecha_p = st.date_input("Fecha de Producción (F.P)", datetime.now())

  with col2:
    lote = st.text_input("Lote")
    batch = st.text_input("Batch")
    responsable = st.text_input("Responsable")
    equipo_final = st.text_input("Equipo Utilizado")

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
    tiempo_calculado = st.text_input("Tiempo (ej. 15 min)")

  observacion_final = st.text_area("Observación")

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

    # Añadir al DataFrame actual
    df_actual = pd.concat(
        [df_actual, pd.DataFrame([nueva_fila])], ignore_index=True
    )
    st.success("¡Registro agregado localmente! Haz clic en guardar abajo.")

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
    # Limpiamos la hoja completa para evitar registros fantasma o desorden
    hoja.clear()

    if not edited_df.empty:
      # Rellenamos nulos y armamos la estructura con cabeceras ordenadas
      df_limpio = edited_df.fillna("")
      data_to_update = [df_limpio.columns.tolist()] + df_limpio.values.tolist()
      hoja.update(data_to_update)
    else:
      # Si borraron todo, dejamos al menos las cabeceras limpias
      hoja.update([headers])

    st.success(
        "¡Google Sheets actualizado correctamente! Los registros eliminados"
        " fueron borrados de la nube."
    )
    st.rerun()

  except Exception as e:
    st.error(f"Error al sincronizar con Google Sheets: {e}")
