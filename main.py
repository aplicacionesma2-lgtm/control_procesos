import base64
import datetime
import io
import json
import os
import gspread
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Procesos - Líneas",
    page_icon="📋",
    layout="wide",
)

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1eQ64LwSp8cVm0T9o29KJgYqfF5e6yCLeN2RqmuY_ftc/edit?gid=0#gid=0"
FILE_PATH = "procesos.xlsx"


@st.cache_resource
def conectar_google_sheets():
    # Reconstruimos las credenciales usando un diccionario limpio
    # para evitar que GitHub o Streamlit corrompan los saltos de línea \n
    pk_lines = [
        "-----BEGIN PRIVATE KEY-----",
        "MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKkwggSkAgEAAoIBAQDBBJfKLxL38J+e",
        "7LS6lFB91EN+vbYs9VhrkjvIA3F7DdEKkW0zpcV0zY3BOQ0gadcDzc+Kf2VhHFYi",
        "Vtub62RIm21bigdql/mMNgVliXnHfOOFQNPeTjdQYGMM6/v8xBemqhVubm2RS822",
        "OkwYaG41iy0av0l3ssqPPe6Os5UpEIGApDc5ERK8kNAAR7ZXGkNzkkp/KzYw/HCL",
        "0n9X5PBqjqcw4avCwbiGHgw00iZo7sqNAN2GlyuYqQaaQ8W6/aAwQFwgotscvocF",
        "7rexVHTIQS3iwnLYGscGf0B/qsz+UAzsym7xj1HDsdXs/Cuy9Eu+2ESsryV+wwM8",
        "UCTeXpGpAgMBAAECggEAHutPrmdijD2rMC/eVpoOF84CHuIgdey6ZIb5FRX6Hnps",
        "31rC6bhXHFoWKFrtf6D8vMMCCT9Vm9wIbzlHNh+bwabGOpjujbR5GO0JacW/MIXQ",
        "w4aKOe0BHtrF2yrNQ6Uc3cmWo8lEO3dvZU7K5EkMSH76M3PrfqVxHcePuKPLU9el",
        "58moIbxis4QXkrdE6eoSP1SFlLHixwFNfO4+6YA7RQbK/g3tTqu+AM8jN8SyuL1f",
        "rAI2WkVrS5Fj9TkLvU6ylQvF7W4GDd6oSSmZRovMTtJEU9zDJLFJuffLvLWWI6g2",
        "lVQdxI7vhY+YMXyQ/NYTOHfsZCcGdZU+jLhscps9dQKBgQDuNysdKrTXGRWiorpA",
        "mnW18cOr0BuZ5mWHeyHxDI3bzKqhWVzqMxw1s2MOLZnoTQ39phIWl0S+s3RcEmCO",
        "Hf8rGpby6CGCHmu897rdp1qkCaU0udhNVn6ctD/NH6srk8aexGpaV0zYQezBQkUu",
        "MLIj+Nbh96B/SXxhai8bK7Q05QKBgQDPbZkSj3EjeCduSHUQzYYPu0ap0o9ubQHK",
        "RG9g5tPlDbgII0XJAJQT3S6c+ma9LkSowRDMLZOOXsRAJf5rbcxd/G7SELQ6NDDa",
        "jmv5r+uQ4JaFBFW9OZDwIFTIVp7nq2eSas6zZL91RSVXKa5mufh2Y8LBjpkWG7pG",
        "Wn2fqj+BdQKBgQCj8uJAa7EUzVXfniGT7vqOo3spF8y3SiOcb/l3Pk2v9heFfsx8",
        "/3ot122YR3hCsi2r4g1W8PtGSJoP+DHt/eUtlFpJicvuEuPRpao9fT3b4iuKs1GU",
        "QLBZR5EVqvMSxd0QTlxoGudve0fn5qVYWflw2oWB9fzHPhtVrFAJYjXfpQKBgQC0",
        "cV/uwG+oblbG3itQQam0t7KB+tShOByNizjksAh2wpdsJNsJPwKRwSBSmJWVTtGV",
        "h9YH+EHbYN8R+rs3Ux2sSPNStAtEcrBo/+o4G+wtbOIjtqCrao+GBGocmRXE7Nu9",
        "iEJl1mejKVKRX4YCgRb+TkxWuqi7jcVefEu6AI0cHQKBgHP0FKxE6bPgYPceXtNV",
        "VICr1sWxBEamX72kkAZ930B+F6t6wDwSJncfgamlQEpnuSSd3rE8D/K+hm4q4A2d",
        "TIMDFdBn8eybpXb9mDC9ADdzyz6yF0mPEvKgG+q/De+oPwjqrDBDI7fRrwSUqaaV",
        "tTOhmvYbJp7/aJD6BPwwFjmn",
        "-----END PRIVATE KEY-----",
    ]
    private_key_str = "\n".join(pk_lines)

    cred_dict = {
        "type": "service_account",
        "project_id": "control-procesos-510021",
        "private_key_id": "aafa643ccb5578aabb4e3747a39c0fa95cac451a",
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
    }

    client = gspread.service_account_from_dict(cred_dict)
    sheet = client.open_by_url(SPREADSHEET_URL).sheet1
    return sheet


# Intento de conexión
try:
    ws = conectar_google_sheets()
except Exception as e:
    st.error(f"❌ Error al conectar con Google Sheets: {e}")
    st.stop()

# --- (El resto de tu código de la app sigue abajo exactamente igual) ---
st.title("📋 FORMATO CONTROL DE PROCESOS LÍNEAS")
st.success(
    "✅ Conectado correctamente a Google Sheets omitiendo fallas de archivos externos."
)
