import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta, time
import openpyxl
import requests
from io import BytesIO
import os

st.set_page_config(page_title="Control Geovita", layout="wide")

FILE_ID = "1E9zKZGSaU8RIfiZLe8UXszgsde9V7WEf"
URL_DRIVE = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"

@st.cache_data(ttl=60) # Bajamos el tiempo de cache para ver cambios rápido
def cargar_datos(url):
    try:
        r = requests.get(url)
        return BytesIO(r.content)
    except: return None

def extraer_hm(v):
    if isinstance(v, (time, datetime)): return v.hour, v.minute
    if isinstance(v, str) and ':' in v:
        try:
            p = v.split(':')
            return int(p[0]), int(p[1])
        except: return None, None
    return None, None

st.title("📊 Control Geovita - Diagnóstico")

archivo = cargar_datos(URL_DRIVE)

if archivo:
    try:
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active
        
        # --- DIAGNÓSTICO: ¿Qué hay en la fila 2? ---
        fechas = {}
        for c in range(3, 100, 2):
            val = ws.cell(row=2, column=c).value
            if val:
                f_txt = val.strftime("%d/%m/%Y") if isinstance(val, datetime) else str(val)
                fechas[f_txt] = c
        
        if not fechas:
            st.error("❌ No se encontraron fechas en la fila 2. Revisa tu Excel.")
        else:
            f_sel = st.sidebar.selectbox("Seleccione Fecha:", list(fechas.keys()))
            c_idx = fechas[f_sel]
            
            # --- DIAGNÓSTICO: Ver datos crudos ---
            datos_encontrados = []
            f_ref = datetime(2024, 1, 15)
            
            for r in range(3, ws.max_row + 1):
                indicador = str(ws.cell(row=r, column=1).value).strip().lower()
                if indicador == 'inicio':
                    tarea = str(ws.cell(row=r, column=2).value)
                    h_i, m_i = extraer_hm(ws.cell(row=r, column=c_idx).value)
                    h_f, m_f = extraer_hm(ws.cell(row=r+1, column=c_idx).value)
                    
                    if h_i is not None:
                        # Lógica de horas...
                        d_i = f_ref + timedelta(hours=h_i, minutes=m_i)
                        if h_i < 8: d_i += timedelta(days=1)
                        d_f = f_ref + timedelta(hours=h_f, minutes=m_f)
                        if h_f < 8: d_f += timedelta(days=1)
                        if d_f <= d_i: d_f += timedelta(days=1)
                        
                        datos_encontrados.append({'Tarea': tarea, 'Inicio': d_i, 'Fin': d_f})

            if not datos_encontrados:
                st.warning(f"⚠️ Se encontró la fecha {f_sel}, pero NO hay filas con la palabra 'inicio' en la Columna A que tengan horas válidas.")
                # Mostramos una tabla de lo que hay en la columna A para ayudar al usuario
                st.write("Contenido detectado en Columna A (primeras 10 filas):")
                st.write([str(ws.cell(row=i, column=1).value) for i in range(3, 13)])
            else:
                # SI HAY DATOS, DIBUJAMOS
                fig, ax = plt.subplots(figsize=(16, 8))
                fig.patch.set_facecolor("#0F0F0F")
                ax.set_facecolor("#0F0F0F")
                
                # Logo (opcional, no bloquea)
                if os.path.exists("logo_geovita.png"):
                    try:
                        img = plt.imread("logo_geovita.png")
                        fig.figimage(img, xo=100, yo=100, alpha=0.1)
                    except: pass

                for d in datos_encontrados:
                    ax.barh(0, (d['Fin'] - d['Inicio']), left=d['Inicio'], height=2, color="#00FFFF", edgecolor="white")
                    ax.text(d['Inicio'], 2.5, d['Tarea'], color="white", fontsize=9)

                ax.set_xlim(f_ref + timedelta(hours=7.5), f_ref + timedelta(hours=32.5))
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
                plt.xticks(color="white")
                plt.yticks([])
                st.pyplot(fig)
                st.success(f"✅ Se cargaron {len(datos_encontrados)} tareas correctamente.")

    except Exception as e:
        st.error(f"Hubo un problema al leer el Excel: {e}")
