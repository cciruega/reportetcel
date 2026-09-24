import streamlit as st
import pandas as pd
import datetime
import requests
import zipfile
import io
import streamlit as st

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
)

# ---------------------------------------------------------
# 🎨 ESTILOS CORPORATIVOS (OCULTAR ICONOS DE STREAMLIT/GITHUB)
# ---------------------------------------------------------
ocultar_iconos = """
<style>
/* 1. Ocultar el encabezado completo (desaparece Fork, GitHub y Menú) */
[data-testid="stHeader"] {
    display: none !important;
}

/* 2. Ocultar barra de herramientas secundaria por seguridad */
[data-testid="stToolbar"] {
    display: none !important;
}

/* 3. Ocultar el menú de hamburguesa nativo */
#MainMenu {
    display: none !important;
}

/* 4. Ocultar pie de página (marca de agua de Streamlit) */
footer {
    display: none !important;
}

/* 5. Ocultar el espacio en blanco que deja el encabezado al desaparecer */
.stApp > header {
    background-color: transparent !important;
}
</style>
"""
st.markdown(ocultar_iconos, unsafe_allow_html=True)
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

# 1. Diccionarios de configuración
estructura_cac = {
    "CIUDAD VICTORIA": ["2008604 TCC CIU100 MANTE4", "2008604 TCC CIU100 VICTORIA II4", "2008604 TCC CIU100 VICTORIA4"],
    "MATAMOROS-REYNOSA": ["2008604 TCC MAT101 MATAMOROS II4", "2008604 TCC MAT101 MATAMOROS4", "2008604 TCC REY115 REYNOSA II4", "2008604 TCC REY115 REYNOSA III4", "2008604 TCC REY116 REYNOSA I4", "2008604 TCC REY116 REYNOSA IV4"],
    "MONTERREY 1": ["2008604 TCC MON103 COUNTRY4", "2008604 TCC MON103 EXPRESS ESFERA4", "2008604 TCC MON103 EXPRESS NUEVO SUR4", "2008604 TCC MON103 SATELITE4", "2008604 TCC MON103 VALLE ORIENTE4", "2008604 TCC MON104 CUMBRES4", "2008604 TCC MON104 SENDERO LINCOLN4", "2008604 TCC MON104 SERVICIO TECNICO Tlc Y CENTRo", "2008604 TCC MON105 CENTRIKA4", "2008604 TCC MON105 GALERIAS4", "2008604 TCC MON106 CENTRO4", "2008604 TCC MON106 EXPRESS FASHION DRIVE4", "2008604 TCC MON106 EXPRESS HUMBERTO LOBO4", "2008604 TCC MON106 EXPRESS PASEO TEC", "2008604 TCC MON106 EXPRESS VILLAS VALLE4", "2008604 TCC MON106 PUNTO VALLE4", "2008604 TCC MON106 SAN AGUSTIN4"],
    "MONTERREY 2": ["2008604 TCC MON107 EXPOSICION4", "2008604 TCC MON107 GUADALUPE4", "2008604 TCC MON108 ANAHUAC4", "2008604 TCC MON108 EXPRESS PLAZA FIESTA ANAHUAC4", "2008604 TCC MON108 PLAZA BELLA4", "2008604 TCC MON108 SANTA CATARINA4", "2008604 TCC MON109 CITADEL4", "2008604 TCC MON109 LAS AMERICAS4", "2008604 TCC MON110 ESCOBEDO4"],
    "MONTERREY 3": ["2008604 TCC MON111 APODACA4", "2008604 TCC MON111 MTY SUN MALL VIP4", "2008604 TCC MON112 EXPRESS MONTEMORELOS"],
    "NUEVO LAREDO": ["2008604 TCC NUE114 LAREDO I4", "2008604 TCC NUE114 LAREDO II4"],
    "TAMPICO": ["2008604 TCC TAM121 TAMPICO I4", "2008604 TCC TAM122 TAMPICO II4", "2008604 TCC TAM122 TAMPICO III4", "2008604 TCC TAM122 TAMPICO IV4"]
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

# Diccionario para nombres limpios en la vista del portal
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
    "2008604 TCC MON103 SATELITE4": "SATELITE",
    "2008604 TCC MON103 VALLE ORIENTE4": "VALLE ORIENTE",
    "2008604 TCC MON104 CUMBRES4": "CUMBRES",
    "2008604 TCC MON104 SENDERO LINCOLN4": "SENDERO LINCOLN",
    "2008604 TCC MON104 SERVICIO TECNICO Tlc Y CENTRo": "SERVICIO TECNICO Tlc Y CENTRo",
    "2008604 TCC MON105 CENTRIKA4": "CENTRIKA",
    "2008604 TCC MON105 GALERIAS4": "GALERIAS",
    "2008604 TCC MON106 CENTRO4": "CENTRO",
    "2008604 TCC MON106 EXPRESS FASHION DRIVE4": "EXPRESS FASHION DRIVE",
    "2008604 TCC MON106 EXPRESS HUMBERTO LOBO4": "EXPRESS HUMBERTO LOBO",
    "2008604 TCC MON106 EXPRESS PASEO TEC": "EXPRESS PASEO TEC",
    "2008604 TCC MON106 EXPRESS VILLAS VALLE4": "EXPRESS VILLAS VALLE",
    "2008604 TCC MON106 PUNTO VALLE4": "PUNTO VALLE",
    "2008604 TCC MON106 SAN AGUSTIN4": "SAN AGUSTIN",
    "2008604 TCC MON107 EXPOSICION4": "EXPOSICION",
    "2008604 TCC MON107 GUADALUPE4": "GUADALUPE",
    "2008604 TCC MON108 ANAHUAC4": "ANAHUAC",
    "2008604 TCC MON108 EXPRESS PLAZA FIESTA ANAHUAC4": "EXPRESS PLAZA FIESTA ANAHUAC",
    "2008604 TCC MON108 PLAZA BELLA4": "PLAZA BELLA",
    "2008604 TCC MON108 SANTA CATARINA4": "SANTA CATARINA",
    "2008604 TCC MON109 CITADEL4": "CITADEL",
    "2008604 TCC MON109 LAS AMERICAS4": "LAS AMERICAS",
    "2008604 TCC MON110 ESCOBEDO4": "ESCOBEDO",
    "2008604 TCC MON111 APODACA4": "APODACA",
    "2008604 TCC MON111 MTY SUN MALL VIP4": "SUN MALL VIP",
    "2008604 TCC MON112 EXPRESS MONTEMORELOS": "MONTEMORELOS",
    "2008604 TCC NUE114 LAREDO I4": "LAREDO I",
    "2008604 TCC NUE114 LAREDO II4": "LAREDO II",
    "2008604 TCC TAM121 TAMPICO I4": "TAMPICO I",
    "2008604 TCC TAM122 TAMPICO II4": "TAMPICO II",
    "2008604 TCC TAM122 TAMPICO III4": "TAMPICO III",
    "2008604 TCC TAM122 TAMPICO IV4": "TAMPICO IV"
}

st.markdown("<h1 style='text-align: center;'>Telmex-Telcel</h1>", unsafe_allow_html=True)

st.markdown("---")
# Función para cargar el archivo sorteando las primeras filas vacías
def cargar_datos(archivo):
    df_temp = pd.read_excel(archivo, sheet_name="Detalle1", header=2)
    if 'NOM_ESTRATEGIA' in df_temp.columns:
        return df_temp
    else:
        return pd.read_excel(archivo, sheet_name="Detalle1")

# ---------------------------------------------------------
# ☁️ LÓGICA DE DETECCIÓN AUTOMÁTICA (CLARO DRIVE)
# ---------------------------------------------------------
@st.cache_data(ttl=600) # Cacheamos por 10 mins para no saturar ClaroDrive
def obtener_archivo_clarodrive():
    # Truco de ClaroDrive: Agregamos /download a tu liga para bajar la carpeta
    url_carpeta = "https://i0000.clarodrive.com/s/FSXKpraaEE8owPZ"
    url_descarga = url_carpeta.rstrip('/') + '/download'
    
    try:
        respuesta = requests.get(url_descarga, timeout=15)
        if respuesta.status_code == 200:
            # Leemos el archivo ZIP directamente en la memoria del servidor
            with zipfile.ZipFile(io.BytesIO(respuesta.content)) as archivo_zip:
                # Buscamos todos los archivos Excel (ignorando los temporales que empiezan con ~)
                excel_infos = [info for info in archivo_zip.infolist() if info.filename.endswith('.xlsx') and not info.filename.startswith('~')]
                
                if excel_infos:
                    # Si hay varios, tomamos el más reciente por fecha de modificación
                    excel_reciente = max(excel_infos, key=lambda x: x.date_time)
                    
                    # Lo extraemos a la memoria
                    archivo_bytes = io.BytesIO(archivo_zip.read(excel_reciente.filename))
                    
                    # =========================================================
                    # 🕒 AJUSTE DE ZONA HORARIA (UTC A CENTRO DE MÉXICO)
                    # =========================================================
                    fecha_tupla = excel_reciente.date_time 
                    
                    # 1. Convertimos la tupla del ZIP a un formato de fecha manipulable
                    fecha_utc = datetime.datetime(
                        year=fecha_tupla[0], month=fecha_tupla[1], day=fecha_tupla[2],
                        hour=fecha_tupla[3], minute=fecha_tupla[4], second=fecha_tupla[5]
                    )
                    
                    # 2. Le restamos 6 horas (Diferencia de México respecto a UTC)
                    fecha_mexico = fecha_utc - datetime.timedelta(hours=6)
                    
                    # 3. Lo convertimos al texto final
                    fecha_str = fecha_mexico.strftime('%d/%m/%Y %H:%M:%S')
                    # =========================================================
                    
                    return archivo_bytes.getvalue(), excel_reciente.filename, fecha_str
    except Exception:
        pass # Si falla el internet del servidor o la liga, no rompe el programa
    
    return None, None, None

# Obtenemos los bytes en lugar del objeto BytesIO para que Streamlit pueda cachear
bytes_automatico, nombre_corto, fecha_actualizacion = obtener_archivo_clarodrive()
archivo_a_procesar = None

col1, col2 = st.columns([2, 1])
with col1:
    if bytes_automatico:
        # Reconstruimos el BytesIO a partir de los bytes cacheados
        archivo_automatico = io.BytesIO(bytes_automatico)
        st.success(f"☁️ **Base de datos (Claro Drive):** {nombre_corto}  \n⏱️ **Actualizado:** {fecha_actualizacion}")
        archivo_a_procesar = archivo_automatico
    else:
        st.warning("⚠️ No se pudo conectar con Claro Drive o la carpeta está vacía.")

with col2:
    # Si Claro Drive falla, habilitamos la subida manual como "Plan B"
    usar_manual = st.checkbox("Subir archivo manualmente", value=False if bytes_automatico else True)

if usar_manual:
    archivo_a_procesar = st.file_uploader("Arrastra aquí tu archivo de Excel", type=['xlsx', 'xls'])

st.divider()

# ---------------------------------------------------------
# PROCESAMIENTO DEL ARCHIVO
# ---------------------------------------------------------
if archivo_a_procesar:
    try:
        df = cargar_datos(archivo_a_procesar)
        
        st.sidebar.header("Filtros Principales")
        
        if 'MES_CAPTURA' in df.columns:
            df['MES_CAPTURA'] = pd.to_datetime(df['MES_CAPTURA'], errors='coerce')
            meses_disponibles = sorted(df['MES_CAPTURA'].dt.month.dropna().unique().astype(int).tolist())
            mes_actual = datetime.datetime.now().month
            
            default_index = meses_disponibles.index(mes_actual) if mes_actual in meses_disponibles else (len(meses_disponibles)-1 if meses_disponibles else 0)
            
            mes_seleccionado = st.sidebar.selectbox("Mes de Captura (Número)", meses_disponibles, index=default_index)
            
            df_filtrado = df[df['MES_CAPTURA'].dt.month == mes_seleccionado]
        else:
            st.error("La columna 'MES_CAPTURA' no se encontró. Verifica el formato del archivo.")
            df_filtrado = df

        if 'NOM_ESTRATEGIA' in df_filtrado.columns:
            
            resumen_cacs = df_filtrado.groupby('NOM_ESTRATEGIA').size().reset_index(name='Avance Mes')
            
            st.header("Resultados por CAC asociado al Area TMX")
            
            ranking_areas = [] # <--- LISTA INICIADA AQUÍ
            
            for area, cacs in estructura_cac.items():
                st.subheader(area)
                
                datos_area = []
                total_avance_mes = 0
                total_asesores = 0
                total_meta = 0
                
                for cac in cacs:
                    avance_fila = resumen_cacs[resumen_cacs['NOM_ESTRATEGIA'] == cac]
                    avance = avance_fila['Avance Mes'].values[0] if not avance_fila.empty else 0
                    
                    asesores = catalogo_asesores.get(cac, 0)
                    meta = asesores * 2
                    
                    porcentaje = (avance / meta) if meta > 0 else 0
                    
                    nombre_mostrar = nombres_simples.get(cac, cac)
                    
                    datos_area.append({
                        "Area/CAC": nombre_mostrar,
                        "Avance Mes": avance,
                        "Asesores": asesores,
                        "Meta": meta,
                        "Avance": porcentaje
                    })
                    
                    total_avance_mes += avance
                    total_asesores += asesores
                    total_meta += meta
                    
                total_porcentaje = (total_avance_mes / total_meta) if total_meta > 0 else 0
                
                # <--- GUARDADO EN LA LISTA AQUÍ
                ranking_areas.append({
                    "Área": area,
                    "Cumplimiento %": total_porcentaje * 100
                })
                
                datos_area.insert(0, {
                    "Area/CAC": f"[-]{area} (TOTAL)",
                    "Avance Mes": total_avance_mes,
                    "Asesores": total_asesores,
                    "Meta": total_meta,
                    "Avance": total_porcentaje
                })
                    
                df_area = pd.DataFrame(datos_area)
                
                # -----------------------------------------------------
                # CÁLCULO DE ALTURA DINÁMICA
                # (Número de filas + 1 del encabezado) * 35 píxeles
                # -----------------------------------------------------
                altura_tabla = (len(df_area) + 1) * 35 + 3
                
                # Creación ÚNICA de la tabla con todos los formatos visuales (Semáforo y Barra)
                st.dataframe(
                    df_area.style.format({"Avance": "{:.0%}"}).map(colorear_semaforo, subset=['Avance']),
                    width="content",  # <--- ASEGÚRATE DE QUE DIGA "content" ENTRE COMILLAS
                    hide_index=True,
                    height=altura_tabla,
                    column_config={
                        "Avance": st.column_config.ProgressColumn(
                            "Avance",
                            help="Cumplimiento de la meta",
                            format="%.2f", 
                            min_value=0,
                            max_value=1,   
                        )
                    }
                )
                
                # --- NUEVO: BOTÓN DE DESCARGA ---
                # Formateamos el DataFrame antes de descargarlo para que los porcentajes se vean bien en Excel
                df_descarga = df_area.copy()
                df_descarga['Avance'] = df_descarga['Avance'].apply(lambda x: f"{x:.0%}")
                
                excel_data = convertir_df_a_excel(df_descarga)
                st.download_button(
                    label=f"📥 Descargar tabla de {area} (Excel)",
                    data=excel_data,
                    file_name=f"Resultados_{area}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"btn_descarga_{area}" # Es vital ponerle una key única a cada botón
                )
                
                st.markdown("---")
            
            # <--- GRÁFICA FUERA DEL CICLO (MISMA ALINEACIÓN QUE EL FOR)
            st.divider()
            st.subheader("📊 Ranking Global de Cumplimiento")
            df_ranking = pd.DataFrame(ranking_areas)
            df_ranking = df_ranking.sort_values(by="Cumplimiento %", ascending=False)
            st.bar_chart(df_ranking.set_index("Área")["Cumplimiento %"])
                
        else:
            st.error("La columna 'NOM_ESTRATEGIA' no se encontró en la base de datos.")

    except Exception as e:
        st.error(f"Hubo un problema al leer el archivo. Error técnico: {e}")
else:
    st.info("Obteniendo datos de Claro Drive o en espera de subida manual...")
