import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta, time
import openpyxl
import requests
from io import BytesIO

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Control Geovita", layout="wide")

# URL del Excel en Drive
FILE_ID = "1E9zKZGSaU8RIfiZLe8UXszgsde9V7WEf"
URL_DRIVE = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"

@st.cache_data(ttl=300)
def descargar_datos(url):
    try:
        r = requests.get(url, timeout=10)
        return BytesIO(r.content)
    except:
        return None

def obtener_hm(v):
    if isinstance(v, (time, datetime)): return v.hour, v.minute
    if isinstance(v, str) and ':' in v:
        try:
            p = v.split(':')
            return int(p[0]), int(p[1])
        except: return None, None
    return None, None

def color_tarea(t):
    colores = {'tronadura': '#FF0000', 'ventilacion': '#00FFFF', 'marina': '#8B4513', 'otros': '#D3D3D3'}
    t_low = str(t).lower()
    for k, v in colores.items():
        if k in t_low: return v
    return colores['otros']

st.title("📊 Control Operacional Geovita")

# --- PROCESO ---
archivo = descargar_datos(URL_DRIVE)

if archivo:
    try:
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active
        
        # Mapeo de fechas
        fechas = {}
        for c in range(3, 150, 2):
            val = ws.cell(row=2, column=c).value
            if val:
                f_txt = val.strftime("%d/%m/%Y") if isinstance(val, datetime) else str(val)
                fechas[f_txt] = c

        if fechas:
            f_sel = st.sidebar.selectbox("📅 Seleccione Fecha:", list(fechas.keys()))
            c_idx = fechas[f_sel]
            
            tareas = []
            f_ref = datetime(2024, 1, 15)
            
            for r in range(3, ws.max_row + 1):
                if str(ws.cell(row=r, column=1).value).lower() == 'inicio':
                    h_i, m_i = obtener_hm(ws.cell(row=r, column=c_idx).value)
                    h_f, m_f = obtener_hm(ws.cell(row=r+1, column=c_idx).value)
                    
                    if h_i is not None:
                        dt_i = f_ref + timedelta(hours=h_i, minutes=m_i)
                        if h_i < 8: dt_i += timedelta(days=1)
                        dt_f = f_ref + timedelta(hours=h_f, minutes=m_f)
                        if h_f < 8: dt_f += timedelta(days=1)
                        if dt_f <= dt_i: dt_f += timedelta(days=1)
                        
                        tareas.append({
                            'nombre': str(ws.cell(row=r, column=2).value),
                            'ini': dt_i, 'fin': dt_f, 'col': color_tarea(ws.cell(row=r, column=2).value)
                        })

            if tareas:
                fig, ax = plt.subplots(figsize=(16, 8))
                fig.patch.set_facecolor("#0F0F0F")
                ax.set_facecolor("#0F0F0F")

                # Grilla y Línea Base
                ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
                ax.grid(axis='x', color='white', alpha=0.1)
                ax.axhline(0, color="white", linewidth=2)

                # Líneas de Turno (8 AM y 8 PM)
                for h in [8, 20, 32]:
                    ax.axvline(f_ref + timedelta(hours=h), color="red", ls="--", lw=2)

                # Zonas de Colación / Charla
                zonas = [(8, 8.75, "CHARLA"), (12, 15, "ALMUERZO"), (20, 20.75, "CHARLA N."), (24, 27, "CENA")]
                for zi, zf, zl in zonas:
                    ax.axvspan(f_ref+timedelta(hours=zi), f_ref+timedelta(hours=zf), color="white", alpha=0.05)
                    ax.text(f_ref+timedelta(hours=(zi+zf)/2), -9, zl, color="white", alpha=0.3, ha='center', fontsize=7)

                # Dibujar Tareas (Forma simplificada para evitar errores)
                for i, t in enumerate(tareas):
                    # Alternamos altura para que no se traslapen etiquetas
                    h_pos = 4 if (i % 2 == 0) else -4
                    ax.barh(0, (t['fin'] - t['ini']), left=t['ini'], height=2, color=t['col'], edgecolor="white")
                    ax.text(t['ini'], h_pos, t['nombre'], color="white", fontsize=8, fontweight='bold', rotation=30)

                ax.set_xlim(f_ref + timedelta(hours=7.5), f_ref + timedelta(hours=32.5))
                ax.set_ylim(-10, 10)
                plt.xticks(color="white", rotation=45)
                plt.yticks([])
                st.pyplot(fig)
    except Exception as e:
        st.error(f"Error en procesamiento: {e}")
