import streamlit as st
import pandas as pd
import datetime
import requests
import zipfile
import io

def convertir_df_a_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultados')
    processed_data = output.getvalue()
    return processed_data

st.set_page_config(
    page_title="Dashboard de Resultados",
    page_icon="reportetcel/telcel2.ico",  # Puede ser una ruta local o una URL
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# 🎨 ESTILOS CORPORATIVOS
# ---------------------------------------------------------
st.markdown("""
<style>
footer {
    display: none !important;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# FUNCIONES AUXILIARES
# ---------------------------------------------------------
def colorear_semaforo(val):
    if isinstance(val, str):
        return ''
    if val >= 0.80:
        color = '#28a745' # Verde
    elif val >= 0.50:
        color = '#ffc107' # Amarillo/Naranja
    else:
        color = '#dc3545' # Rojo
    return f'color: {color}; font-weight: bold;'

def generar_boton_descarga(df, nombre_archivo, btn_key):
    try:
        df_export = df.copy()
        
        # Formateo solo si existe Avance
        if 'Avance' in df_export.columns:
            df_export['Avance'] = pd.to_numeric(df_export['Avance'], errors='coerce').fillna(0)
            df_export['Avance'] = df_export['Avance'].apply(lambda x: f"{x:.0%}")
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Resultados')
        excel_data = output.getvalue()
        
        st.download_button(
            label=f"📥 Descargar {nombre_archivo}.xlsx",
            data=excel_data,
            file_name=f"{nombre_archivo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=btn_key
        )
        
    except Exception as e:
        csv_data = df_export.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=f"📥 Descargar {nombre_archivo} (CSV)",
            data=csv_data,
            file_name=f"{nombre_archivo}.csv",
            mime="text/csv",
            key=f"{btn_key}_csv"
        )

# ---------------------------------------------------------
# DICCIONARIOS DE CONFIGURACIÓN
# ---------------------------------------------------------
estructura_cac = {
    "CIUDAD VICTORIA": ["2008604 TCC CIU100 MANTE4", "2008604 TCC CIU100 VICTORIA II4", "2008604 TCC CIU100 VICTORIA4"],
    "MATAMOROS-REYNOSA": ["2008604 TCC MAT101 MATAMOROS II4", "2008604 TCC MAT101 MATAMOROS4", "2008604 TCC REY115 REYNOSA II4", "2008604 TCC REY115 REYNOSA III4", "2008604 TCC REY116 REYNOSA I4", "2008604 TCC REY116 REYNOSA IV4"],
    "MONTERREY 1": ["2008604 TCC MON103 COUNTRY4", "2008604 TCC MON103 EXPRESS ESFERA4", "2008604 TCC MON103 EXPRESS NUEVO SUR4", "2008604 TCC MON103 SATELITE4", "2008604 TCC MON103 VALLE ORIENTE4", "2008604 TCC MON104 CUMBRES4", "2008604 TCC MON104 SENDERO LINCOLN4", "2008604 TCC MON104 SERVICIO TECNICO Tlc Y CENTRo", "2008604 TCC MON105 CENTRIKA4", "2008604 TCC MON105 GALERIAS4", "2008604 TCC MON106 CENTRO4", "2008604 TCC MON106 EXPRESS FASHION DRIVE4", "2008604 TCC MON106 EXPRESS HUMBERTO LOBO4", "2008604 TCC MON106 EXPRESS PASEO TEC", "2008604 TCC MON106 EXPRESS VILLAS VALLE4", "2008604 TCC MON106 PUNTO VALLE4", "2008604 TCC MON106 SAN AGUSTIN4"],
    "MONTERREY 2": ["2008604 TCC MON107 EXPOSICION4", "2008604 TCC MON107 GUADALUPE4", "2008604 TCC MON108 ANAHUAC4", "2008604 TCC MON108 EXPRESS PLAZA FIESTA ANAHUAC4", "2008604 TCC MON108 PLAZA BELLA4", "2008604 TCC MON108 SANTA CATARINA4", "2008604 TCC MON109 CITADEL4", "2008604 TCC MON109 LAS AMERICAS4", "2008604 TCC MON110 ESCOBEDO4"],
    "MONTERREY 3": ["2008604 TCC MON111 APODACA4", "2008604 TCC MON111 MTY SUN MALL VIP4", "2008604 TCC MON112 EXPRESS MONTEMORELOS"],
    "NUEVO LAREDO": ["2008604 TCC NUE114 LAREDO I4", "2008604 TCC NUE114 LAREDO II4"],
    "TAMPICO": ["2008604 TCC TAM121 TAMPICO I4", "2008604 TCC TAM122 TAMPICO II4", "2008604 TCC TAM122 TAMPICO III4", "2008604 TCC TAM122 TAMPICO IV4"]
}

estructura_cope = {
    "CIUDAD VICTORIA": ["CT CIUDAD MANTE", "CT CIUDAD VICTORIA"],
    "MATAMOROS-REYNOSA": ["CT CIUDAD MIGUEL ALEMAN", "CT MATAMOROS", "CT REYNOSA", "CT RIO BRAVO", "CT SAN FERNANDO", "CT VALLE HERMOSO"],
    "MONTERREY 1": ["CT COLON (MTY)", "CT GONZALITOS", "CT LINCOLN", "CT REVOLUCION", "CT SAN PEDRO [NL]"],
    "MONTERREY 2": ["CT LA SILLA", "CT PUENTES", "CT SANTA CATARINA", "CT SANTA FE [MTY]", "CT UNIVERSIDAD (NL)"],
    "MONTERREY 3": ["CT APODACA", "CT CADEREYTA", "CT GENERAL ESCOBEDO (BRISAS)", "CT LINARES", "CT MONTEMORELOS"],
    "NUEVO LAREDO": ["CT CIUDAD ANAHUAC", "CT NUEVO LAREDO", "CT SABINAS HIDALGO"],
    "TAMPICO": ["CT TAMPICO HIDALGO", "CT TAMPICO MADERO"]
}

catalogo_asesores = {
    "2008604 TCC CIU100 MANTE4": 17, "2008604 TCC CIU100 VICTORIA II4": 19, "2008604 TCC CIU100 VICTORIA4": 28,
    "2008604 TCC MAT101 MATAMOROS II4": 21, "2008604 TCC MAT101 MATAMOROS4": 19, "2008604 TCC REY115 REYNOSA II4": 15, "2008604 TCC REY115 REYNOSA III4": 20, "2008604 TCC REY116 REYNOSA I4": 14, "2008604 TCC REY116 REYNOSA IV4": 16,
    "2008604 TCC MON103 COUNTRY4": 22, "2008604 TCC MON103 EXPRESS ESFERA4": 10, "2008604 TCC MON103 EXPRESS NUEVO SUR4": 7, "2008604 TCC MON103 SATELITE4": 14, "2008604 TCC MON103 VALLE ORIENTE4": 14, "2008604 TCC MON104 CUMBRES4": 29, "2008604 TCC MON104 SENDERO LINCOLN4": 30, "2008604 TCC MON104 SERVICIO TECNICO Tlc Y CENTRo": 19, "2008604 TCC MON105 CENTRIKA4": 15, "2008604 TCC MON105 GALERIAS4": 25, "2008604 TCC MON106 CENTRO4": 24, "2008604 TCC MON106 EXPRESS FASHION DRIVE4": 6, "2008604 TCC MON106 EXPRESS HUMBERTO LOBO4": 13, "2008604 TCC MON106 EXPRESS PASEO TEC": 7, "2008604 TCC MON106 EXPRESS VILLAS VALLE4": 4, "2008604 TCC MON106 PUNTO VALLE4": 11, "2008604 TCC MON106 SAN AGUSTIN4": 18,
    "2008604 TCC MON107 EXPOSICION4": 20, "2008604 TCC MON107 GUADALUPE4": 28, "2008604 TCC MON108 ANAHUAC4": 20, "2008604 TCC MON108 EXPRESS PLAZA FIESTA ANAHUAC4": 8, "2008604 TCC MON108 PLAZA BELLA4": 27, "2008604 TCC MON108 SANTA CATARINA4": 23, "2008604 TCC MON109 CITADEL4": 25, "2008604 TCC MON109 LAS AMERICAS4": 20, "2008604 TCC MON110 ESCOBEDO4": 20,
    "2008604 TCC MON111 APODACA4": 26, "2008604 TCC MON111 MTY SUN MALL VIP4": 20, "2008604 TCC MON112 EXPRESS MONTEMORELOS": 7,
    "2008604 TCC NUE114 LAREDO I4": 9, "2008604 TCC NUE114 LAREDO II4": 15,
    "2008604 TCC TAM121 TAMPICO I4": 34, "2008604 TCC TAM122 TAMPICO II4": 20, "2008604 TCC TAM122 TAMPICO III4": 23, "2008604 TCC TAM122 TAMPICO IV4": 23
}

nombres_simples = {
    "2008604 TCC CIU100 MANTE4": "MANTE",
    "2008604 TCC CIU100 VICTORIA II4": "VICTORIA II",
    "2008604 TCC CIU100 VICTORIA4": "VICTORIA",
    "2008604 TCC MAT101 MATAMOROS II4": "MATAMOROS II",
    "2008604 TCC MAT101 MATAMOROS4": "MATAMOROS",
    "2008604 TCC REY115 REYNOSA II4": "REYNOSA II",
    "2008604 TCC REY115 REYNOSA III4": "REYNOSA III",
    "2008604 TCC REY116 REYNOSA I4": "REYNOSA I",
    "2008604 TCC REY116 REYNOSA IV4": "REYNOSA IV",
    "2008604 TCC MON103 COUNTRY4": "COUNTRY",
    "2008604 TCC MON103 EXPRESS ESFERA4": "EXPRESS ESFERA",
    "2008604 TCC MON103 EXPRESS NUEVO SUR4": "EXPRESS NUEVO SUR",
    "2008604 TCC MON103 SATEL
