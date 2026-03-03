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

# URL del Excel
FILE_ID = "1E9zKZGSaU8RIfiZLe8UXszgsde9V7WEf"
URL_DRIVE = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"

@st.cache_data(ttl=60)
def cargar_datos(url):
    try:
        r = requests.get(url, timeout=10)
        return BytesIO(r.content)
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        return None

def extraer_hm(v):
    if isinstance(v, (time, datetime)): return v.hour, v.minute
    if isinstance(v, str) and ':' in v:
        try:
            p = v.split(':')
            return int(p[0]), int(p[1])
        except: return None, None
    return None, None

st.title("📊 Panel de Control Geovita")

archivo = cargar_datos(URL_DRIVE)

if archivo:
    try:
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active
        
        # 1. Detectar fechas en Fila 2
        fechas = {}
        for c in range(3, 100, 2):
            val = ws.cell(row=2, column=c).value
            if val:
                f_txt = val.strftime("%d/%m/%Y") if isinstance(val, datetime) else str(val)
                fechas[f_txt] = c
        
        if not fechas:
            st.error("No se encontraron fechas en la Fila 2 del Excel.")
            st.info("Asegúrate de que las fechas empiecen desde la columna C.")
        else:
            f_sel = st.sidebar.selectbox("Seleccione Fecha:", list(fechas.keys()))
            c_idx = fechas[f_sel]
            
            # 2. Extraer Tareas
            tareas = []
            f_ref = datetime(2024, 1, 15)
            
            for r in range(3, ws.max_row + 1):
                indicador = str(ws.cell(row=r, column=1).value).strip().lower()
                if indicador == 'inicio':
                    nombre = str(ws.cell(row=r, column=2).value)
                    h_i, m_i = extraer_hm(ws.cell(row=r, column=c_idx).value)
                    h_f, m_f = extraer_hm(ws.cell(row=r+1, column=c_idx).value)
                    
                    if h_i is not None:
                        dt_i = f_ref + timedelta(hours=h_i, minutes=m_i)
                        if h_i < 8: dt_i += timedelta(days=1)
                        dt_f = f_ref + timedelta(hours=h_f, minutes=m_f)
                        if h_f < 8: dt_f += timedelta(days=1)
                        if dt_f <= dt_i: dt_f += timedelta(days=1)
                        
                        tareas.append({'Tarea': nombre, 'Inicio': dt_i, 'Fin': dt_f})

            # 3. Mostrar Gráfico o Datos
            if tareas:
                fig, ax = plt.subplots(figsize=(16, 7))
                fig.patch.set_facecolor("#0F0F0F")
                ax.set_facecolor("#0F0F0F")

                for t in tareas:
                    ax.barh(0, (t['Fin'] - t['Inicio']), left=t['Inicio'], height=2, color="#00FFFF", edgecolor="white")
                    ax.text(t['Inicio'], 1.2, t['Tarea'], color="white", fontsize=8)

                ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
                ax.set_xlim(f_ref + timedelta(hours=7.5), f_ref + timedelta(hours=32.5))
                plt.xticks(color="white")
                plt.yticks([])
                st.pyplot(fig)
                
                # Mostrar tabla abajo para verificar
                st.write("### Resumen de Tareas Detectadas")
                st.table(pd.DataFrame(tareas))
            else:
                st.warning(f"No hay tareas con la palabra 'inicio' en la Columna A para la fecha {f_sel}.")
                st.write("Primeras 5 filas de la Columna A encontradas:")
                st.write([str(ws.cell(row=i, column=1).value) for i in range(3, 8)])

    except Exception as e:
        st.error(f"Error al leer Excel: {e}")
else:
    st.error("No se pudo descargar el archivo de Drive. Verifica los permisos de compartir.")
