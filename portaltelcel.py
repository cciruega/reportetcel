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
    # Ahora evaluamos en escala de 0 a 100
    if val >= 80.0:
        color = '#28a745' # Verde
    elif val >= 50.0:
        color = '#ffc107' # Amarillo/Naranja
    else:
        color = '#dc3545' # Rojo
    return f'color: {color}; font-weight: bold;'

def generar_boton_descarga(df, nombre_archivo, btn_key):
    try:
        df_export = df.copy()
        
        # Formateo seguro para Excel: Se aplica a Avance y Efectividad
        columnas_porcentaje = ['Avance', 'Efectividad']
        for col in columnas_porcentaje:
            if col in df_export.columns:
                df_export[col] = pd.to_numeric(df_export[col], errors='coerce').fillna(0)
                # Como ahora los datos vienen de 0 a 100, solo agregamos el símbolo %
                df_export[col] = df_export[col].apply(lambda x: f"{x:.0f}%")
            
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

mapa_estados = {
    "Nuevo León": [
        "MONTERREY 1",
        "MONTERREY 2",
        "MONTERREY 3"
    ],
    "Tamaulipas": [
        "CIUDAD VICTORIA",
        "MATAMOROS-REYNOSA",
        "NUEVO LAREDO",
        "TAMPICO"
    ]
}

st.markdown("<h1 style='text-align: center;'>Telmex-Telcel</h1>", unsafe_allow_html=True)
st.markdown("---")

def cargar_datos(archivo):
    df_temp = pd.read_excel(archivo, sheet_name="Detalle1", header=2)
    if 'NOM_ESTRATEGIA' in df_temp.columns:
        return df_temp
    else:
        return pd.read_excel(archivo, sheet_name="Detalle1")

# ---------------------------------------------------------
# LÓGICA DE DETECCIÓN AUTOMÁTICA (CLARO DRIVE)
# ---------------------------------------------------------
@st.cache_data(ttl=600)
def obtener_archivo_clarodrive():
    url_carpeta = "https://i0000.clarodrive.com/s/FSXKpraaEE8owPZ"
    url_descarga = url_carpeta.rstrip('/') + '/download'
    try:
        respuesta = requests.get(url_descarga, timeout=15)
        if respuesta.status_code == 200:
            with zipfile.ZipFile(io.BytesIO(respuesta.content)) as archivo_zip:
                excel_infos = [info for info in archivo_zip.infolist() if info.filename.endswith('.xlsx') and not info.filename.startswith('~')]
                if excel_infos:
                    excel_reciente = max(excel_infos, key=lambda x: x.date_time)
                    archivo_bytes = io.BytesIO(archivo_zip.read(excel_reciente.filename))
                    
                    fecha_tupla = excel_reciente.date_time 
                    fecha_utc = datetime.datetime(
                        year=fecha_tupla[0], month=fecha_tupla[1], day=fecha_tupla[2],
                        hour=fecha_tupla[3], minute=fecha_tupla[4], second=fecha_tupla[5]
                    )
                    fecha_mexico = fecha_utc - datetime.timedelta(hours=6)
                    fecha_str = fecha_mexico.strftime('%d/%m/%Y %H:%M:%S')
                    
                    return archivo_bytes.getvalue(), excel_reciente.filename, fecha_str
    except Exception:
        pass
    return None, None, None

bytes_automatico, nombre_corto, fecha_actualizacion = obtener_archivo_clarodrive()
archivo_a_procesar = None

col1, col2 = st.columns([2, 1])
with col1:
    if bytes_automatico:
        archivo_automatico = io.BytesIO(bytes_automatico)
        st.success(f"☁️ **Base de datos:** {nombre_corto}  \n⏱️ **Actualizado:** {fecha_actualizacion}")
        archivo_a_procesar = archivo_automatico
    else:
        st.warning("⚠️ No se pudo conectar con Claro Drive o la carpeta está vacía.")

with col2:
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

        # ---------------------------------------------------------
        # TIPO DE VISTA EN DROPDOWN (SELECTBOX)
        # ---------------------------------------------------------
        tipo_vista = st.sidebar.selectbox(
            "Selecciona la Vista:",
            options=["CACs", "COPEs"],
            index=0,
            help="Elige si deseas ver el reporte por CACs o agrupado por COPEs (CT)."
        )

        st.sidebar.markdown("---")

        # 1. FILTRO DE ESTADO
        estados_disponibles = list(mapa_estados.keys())
        estados_seleccionados = st.sidebar.multiselect(
            "Selecciona Estado:",
            options=estados_disponibles,
            default=estados_disponibles
        )
        
        if not estados_seleccionados:
            st.sidebar.warning("Selecciona al menos un Estado.")
            st.stop()
        
        # 2. ÁREAS
        areas_disponibles = []
        for estado in estados_seleccionados:
            areas_disponibles.extend(mapa_estados.get(estado, []))
        areas_disponibles = list(dict.fromkeys(areas_disponibles))
        
        # 3. FILTRO DE ÁREA
        area_seleccionada = st.sidebar.multiselect(
            "Selecciona Área:",
            options=areas_disponibles,
            default=areas_disponibles
        )
        
        if not area_seleccionada:
            st.sidebar.warning("Selecciona al menos un Área.")
            st.stop()
        
        estructura_cac_filtrada = {k: v for k, v in estructura_cac.items() if k in area_seleccionada}
        estructura_cope_filtrada = {k: v for k, v in estructura_cope.items() if k in area_seleccionada}

        # ---------------------------------------------------------
        # FECHAS
        # ---------------------------------------------------------
        if 'FECHA_CAPTURA' in df.columns:
            df['FECHA_CAPTURA'] = pd.to_datetime(df['FECHA_CAPTURA'], errors='coerce')
            fecha_hoy = datetime.datetime.now()
            anio_actual = fecha_hoy.year
            mes_actual = fecha_hoy.month

            anios_disponibles = sorted(df['FECHA_CAPTURA'].dt.year.dropna().unique().astype(int).tolist(), reverse=True)

            if anios_disponibles:
                anio_predeterminado = anio_actual if anio_actual in anios_disponibles else anios_disponibles[0]
                indice_anio = anios_disponibles.index(anio_predeterminado)
                anio_seleccionado = st.sidebar.selectbox("Año de Captura", options=anios_disponibles, index=indice_anio)

                df_temp_anio = df[df['FECHA_CAPTURA'].dt.year == anio_seleccionado]
                meses_disponibles = sorted(df_temp_anio['FECHA_CAPTURA'].dt.month.dropna().unique().astype(int).tolist(), reverse=True)

                if meses_disponibles:
                    mes_predeterminado = mes_actual if (anio_seleccionado == anio_actual and mes_actual in meses_disponibles) else meses_disponibles[0]
                    indice_mes = meses_disponibles.index(mes_predeterminado)
                    mes_seleccionado = st.sidebar.selectbox("Mes de Captura", options=meses_disponibles, index=indice_mes)

                    df_filtrado = df[(df['FECHA_CAPTURA'].dt.year == anio_seleccionado) & (df['FECHA_CAPTURA'].dt.month == mes_seleccionado)]
                else:
                    st.warning(f"No hay meses disponibles para el año {anio_seleccionado}.")
                    df_filtrado = df.copy()
            else:
                st.warning("No se encontraron fechas válidas en FECHA_CAPTURA.")
                df_filtrado = df.copy()

        elif 'MES_CAPTURA' in df.columns:
            df['MES_CAPTURA'] = pd.to_datetime(df['MES_CAPTURA'], errors='coerce')
            meses_disponibles = sorted(df['MES_CAPTURA'].dt.month.dropna().unique().astype(int).tolist(), reverse=True)

            if meses_disponibles:
                mes_actual = datetime.datetime.now().month
                mes_predeterminado = mes_actual if mes_actual in meses_disponibles else meses_disponibles[0]
                indice_mes = meses_disponibles.index(mes_predeterminado)
                mes_seleccionado = st.sidebar.selectbox("Mes de Captura (Número)", options=meses_disponibles, index=indice_mes)

                df_filtrado = df[df['MES_CAPTURA'].dt.month == mes_seleccionado]
            else:
                st.warning("No se encontraron meses válidos en MES_CAPTURA.")
                df_filtrado = df.copy()

        else:
            st.error("No se encontró 'FECHA_CAPTURA' ni 'MES_CAPTURA'. Verifica el archivo.")
            df_filtrado = df.copy()

        # =========================================================
        # 6. VISTA CACs
        # =========================================================
        if tipo_vista == "CACs":
            if 'NOM_ESTRATEGIA' in df_filtrado.columns:
                resumen_cacs = df_filtrado.groupby('NOM_ESTRATEGIA').size().reset_index(name='Avance Mes')

                st.header("Resultados por CAC")
                ranking_areas = []

                for area, cacs in estructura_cac_filtrada.items():
                    st.subheader(area)

                    datos_area = []
                    total_avance_mes = total_asesores = total_meta = total_instaladas = 0

                    for cac in cacs:
                        avance_fila = resumen_cacs[resumen_cacs['NOM_ESTRATEGIA'] == cac]
                        avance = avance_fila['Avance Mes'].values[0] if not avance_fila.empty else 0
                        asesores = catalogo_asesores.get(cac, 0)
                        meta = asesores * 2
                        # Calculamos directamente de 0 a 100
                        porcentaje = (avance / meta * 100 if meta > 0 else 0)
                        
                        df_cac = df_filtrado[df_filtrado['NOM_ESTRATEGIA'] == cac]
                        instaladas = 0
                        if 'FECHA_POSTEO' in df_cac.columns:
                            try:
                                instaladas = len(df_cac[(df_cac['FECHA_POSTEO'].dt.year == anio_seleccionado) & (df_cac['FECHA_POSTEO'].dt.month == mes_seleccionado)])
                            except NameError:
                                instaladas = len(df_cac[df_cac['FECHA_POSTEO'].dt.month == mes_seleccionado])

                        # Calculamos directamente de 0 a 100
                        efectividad = (instaladas / avance * 100) if avance > 0 else 0
                        datos_area.append({
                            "Area/CAC": nombres_simples.get(cac, cac),
                            "Avance Mes": avance, "Asesores": asesores, "Meta": meta,
                            "Avance": porcentaje, "Instaladas": instaladas, "Efectividad": efectividad
                        })

                        total_avance_mes += avance
                        total_asesores += asesores
                        total_meta += meta
                        total_instaladas += instaladas

                    # Totales del área también en escala 0-100
                    total_porcentaje = (total_avance_mes / total_meta * 100) if total_meta > 0 else 0
                    total_efectividad = (total_instaladas / total_avance_mes * 100) if total_avance_mes > 0 else 0 

                    # Aquí ya no multiplicamos por 100 porque ya viene en esa escala
                    ranking_areas.append({"Área": area, "Cumplimiento %": total_porcentaje})

                    datos_area.insert(0, {
                        "Area/CAC": f"[-]{area} (TOTAL)",
                        "Avance Mes": total_avance_mes, "Asesores": total_asesores, "Meta": total_meta,
                        "Avance": total_porcentaje, "Instaladas": total_instaladas, "Efectividad": total_efectividad
                    })

                    df_area = pd.DataFrame(datos_area)
                    st.dataframe(
                        df_area.style.map(colorear_semaforo, subset=['Avance']),
                        width="content", hide_index=True, height=((len(df_area) + 1) * 35 + 3),
                        column_config={
                            # Ajustamos max_value a 100 y format a "%.0f%%"
                            "Avance": st.column_config.ProgressColumn("Avance", format="%.0f%%", min_value=0, max_value=100),
                            "Efectividad": st.column_config.ProgressColumn("Efectividad", format="%.0f%%", min_value=0, max_value=100)
                        }
                    )

                    # Botón de Descarga DETALLE por Área (CACs)
                    df_detalle_area = df_filtrado[df_filtrado['NOM_ESTRATEGIA'].isin(cacs)]
                    generar_boton_descarga(df_detalle_area, f"Detalle_{area.replace(' ', '_')}", f"btn_{area}")
                    st.markdown("---")

                st.divider()
                if ranking_areas:
                    st.subheader("📊 Ranking Global de Cumplimiento")
                    st.bar_chart(pd.DataFrame(ranking_areas).sort_values(by="Cumplimiento %", ascending=False).set_index("Área")["Cumplimiento %"])

            else:
                st.error("La columna 'NOM_ESTRATEGIA' no se encontró.")

        # =========================================================
        # 7. VISTA COPEs
        # =========================================================
        elif tipo_vista == "COPEs":
            if 'CT' in df_filtrado.columns:
                resumen_copes = df_filtrado.groupby('CT').size().reset_index(name='Venta Mes')

                st.header("Resultados por COPEs (CT)")

                for area, copes in estructura_cope_filtrada.items():
                    st.subheader(area)

                    datos_cope_area = []
                    total_venta_area = 0
                    total_instaladas_area = 0

                    for ct in copes:
                        venta_fila = resumen_copes[resumen_copes['CT'] == ct]
                        venta = venta_fila['Venta Mes'].values[0] if not venta_fila.empty else 0

                        df_ct = df_filtrado[df_filtrado['CT'] == ct]
                        instaladas = 0
                        if 'FECHA_POSTEO' in df_ct.columns:
                            try:
                                instaladas = len(df_ct[(df_ct['FECHA_POSTEO'].dt.year == anio_seleccionado) & (df_ct['FECHA_POSTEO'].dt.month == mes_seleccionado)])
                            except NameError:
                                instaladas = len(df_ct[df_ct['FECHA_POSTEO'].dt.month == mes_seleccionado])

                        # Calculamos directamente de 0 a 100
                        efectividad = (instaladas / venta * 100) if venta > 0 else 0

                        datos_cope_area.append({
                            "Etiquetas de fila": ct,
                            "Venta Mes": venta,
                            "Instaladas": instaladas,
                            "Efectividad": efectividad
                        })

                        total_venta_area += venta
                        total_instaladas_area += instaladas

                    if datos_cope_area:
                        total_efectividad_area = (total_instaladas_area / total_venta_area * 100) if total_venta_area > 0 else 0
                        
                        df_copes_area = pd.DataFrame(datos_cope_area)
                        
                        # Fila de Total General al final (como en tu imagen)
                        fila_total = pd.DataFrame([{
                            "Etiquetas de fila": "Total general",
                            "Venta Mes": total_venta_area,
                            "Instaladas": total_instaladas_area,
                            "Efectividad": total_efectividad_area
                        }])
                        df_copes_area = pd.concat([df_copes_area, fila_total], ignore_index=True)

                        st.dataframe(
                            df_copes_area.style.map(colorear_semaforo, subset=['Efectividad']),
                            width="content", hide_index=True, height=((len(df_copes_area) + 1) * 35 + 3),
                            column_config={
                                # Ajustamos max_value a 100 y format a "%.0f%%"
                                "Efectividad": st.column_config.ProgressColumn("Efectividad", format="%.0f%%", min_value=0, max_value=100)
                            }
                        )

                        # Botón de Descarga DETALLE por Área (COPEs folios)
                        df_detalle_cope_area = df_filtrado[df_filtrado['CT'].isin(copes)]
                        generar_boton_descarga(
                            df_detalle_cope_area, 
                            nombre_archivo=f"Detalle_COPE_{area.replace(' ', '_')}", 
                            btn_key=f"btn_cope_{area}"
                        )
                    else:
                        st.info(f"No hay registros de COPEs para {area}.")
                        
                    st.markdown("---")
            else:
                st.error("La columna 'CT' no se encontró en la base de datos.")

    except Exception as e:
        st.error(f"Hubo un problema al procesar el archivo. Error técnico: {e}")

else:
    st.info("Obteniendo datos de Claro Drive o en espera de subida manual...")
