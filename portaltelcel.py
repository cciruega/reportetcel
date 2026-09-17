import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Dashboard de Resultados", layout="wide")

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
    # CIUDAD VICTORIA
    "2008604 TCC CIU100 MANTE4": 17, "2008604 TCC CIU100 VICTORIA II4": 19, "2008604 TCC CIU100 VICTORIA4": 28,
    # MATAMOROS-REYNOSA
    "2008604 TCC MAT101 MATAMOROS II4": 21, "2008604 TCC MAT101 MATAMOROS4": 19, "2008604 TCC REY115 REYNOSA II4": 15, "2008604 TCC REY115 REYNOSA III4": 20, "2008604 TCC REY116 REYNOSA I4": 14, "2008604 TCC REY116 REYNOSA IV4": 16,
    # MONTERREY 1
    "2008604 TCC MON103 COUNTRY4": 22, "2008604 TCC MON103 EXPRESS ESFERA4": 10, "2008604 TCC MON103 EXPRESS NUEVO SUR4": 7, "2008604 TCC MON103 SATELITE4": 14, "2008604 TCC MON103 VALLE ORIENTE4": 14, "2008604 TCC MON104 CUMBRES4": 29, "2008604 TCC MON104 SENDERO LINCOLN4": 30, "2008604 TCC MON104 SERVICIO TECNICO Tlc Y CENTRo": 19, "2008604 TCC MON105 CENTRIKA4": 15, "2008604 TCC MON105 GALERIAS4": 25, "2008604 TCC MON106 CENTRO4": 24, "2008604 TCC MON106 EXPRESS FASHION DRIVE4": 6, "2008604 TCC MON106 EXPRESS HUMBERTO LOBO4": 13, "2008604 TCC MON106 EXPRESS PASEO TEC": 7, "2008604 TCC MON106 EXPRESS VILLAS VALLE4": 4, "2008604 TCC MON106 PUNTO VALLE4": 11, "2008604 TCC MON106 SAN AGUSTIN4": 18,
    # MONTERREY 2
    "2008604 TCC MON107 EXPOSICION4": 20, "2008604 TCC MON107 GUADALUPE4": 28, "2008604 TCC MON108 ANAHUAC4": 20, "2008604 TCC MON108 EXPRESS PLAZA FIESTA ANAHUAC4": 8, "2008604 TCC MON108 PLAZA BELLA4": 27, "2008604 TCC MON108 SANTA CATARINA4": 23, "2008604 TCC MON109 CITADEL4": 25, "2008604 TCC MON109 LAS AMERICAS4": 20, "2008604 TCC MON110 ESCOBEDO4": 20,
    # MONTERREY 3
    "2008604 TCC MON111 APODACA4": 26, "2008604 TCC MON111 MTY SUN MALL VIP4": 20, "2008604 TCC MON112 EXPRESS MONTEMORELOS": 7,
    # NUEVO LAREDO
    "2008604 TCC NUE114 LAREDO I4": 9, "2008604 TCC NUE114 LAREDO II4": 15,
    # TAMPICO
    "2008604 TCC TAM121 TAMPICO I4": 34, "2008604 TCC TAM122 TAMPICO II4": 20, "2008604 TCC TAM122 TAMPICO III4": 23, "2008604 TCC TAM122 TAMPICO IV4": 23
}

st.title("Portal de Resultados Operativos")

# Función para cargar el archivo sorteando las primeras filas vacías
@st.cache_data
def cargar_datos(archivo):
    # Intentamos leer asumiendo que los encabezados están en la fila 3 (index 2)
    df_temp = pd.read_excel(archivo, sheet_name="Base", header=2)
    
    # Verificamos si logramos capturar la columna NOM_ESTRATEGIA
    if 'NOM_ESTRATEGIA' in df_temp.columns:
        return df_temp
    else:
        # Si la base fue extraída pura (sin el título inicial), leemos normal
        return pd.read_excel(archivo, sheet_name="Base")

archivo_subido = st.file_uploader("Sube el archivo Excel con la base de datos (Hoja 'Base')", type=["xlsx", "xls"])

if archivo_subido:
    try:
        df = cargar_datos(archivo_subido)
        
        # 3. Filtros en la barra lateral
        st.sidebar.header("Filtros Principales")
        
        if 'MES_CAPTURA' in df.columns:
            # Asegurar formato de fecha para extraer los meses limpios
            df['MES_CAPTURA'] = pd.to_datetime(df['MES_CAPTURA'], errors='coerce')
            meses_disponibles = sorted(df['MES_CAPTURA'].dt.month.dropna().unique().astype(int).tolist())
            mes_actual = datetime.now().month
            
            # Buscar el mes actual (si está) o seleccionar el último disponible por defecto
            default_index = meses_disponibles.index(mes_actual) if mes_actual in meses_disponibles else (len(meses_disponibles)-1 if meses_disponibles else 0)
            
            mes_seleccionado = st.sidebar.selectbox("Mes de Captura (Número)", meses_disponibles, index=default_index)
            
            # Filtrar DataFrame por el mes seleccionado
            df_filtrado = df[df['MES_CAPTURA'].dt.month == mes_seleccionado]
        else:
            st.error("La columna 'MES_CAPTURA' no se encontró. Verifica el formato del archivo.")
            df_filtrado = df

        if 'NOM_ESTRATEGIA' in df_filtrado.columns:
            
            # CÁLCULO DEL PORTAL: Agrupar el acumulado para contar los registros por CAC
            # Usamos 'size' para contar el total de solicitudes/ventas en ese mes
            resumen_cacs = df_filtrado.groupby('NOM_ESTRATEGIA').size().reset_index(name='Avance Mes')
            
            st.header("Resultados por AREA_PDV")
            
            for area, cacs in estructura_cac.items():
                st.subheader(area)
                
                datos_area = []
                total_avance_mes = 0
                total_asesores = 0
                total_meta = 0
                
                for cac in cacs:
                    # Buscar el avance calculado para el CAC actual
                    avance_fila = resumen_cacs[resumen_cacs['NOM_ESTRATEGIA'] == cac]
                    avance = avance_fila['Avance Mes'].values[0] if not avance_fila.empty else 0
                    
                    asesores = catalogo_asesores.get(cac, 0)
                    meta = asesores * 2
                    
                    porcentaje = (avance / meta) if meta > 0 else 0
                    
                    datos_area.append({
                        "Area/CAC": cac,
                        "Avance Mes": avance,
                        "Asesores": asesores,
                        "Meta": meta,
                        "Avance": porcentaje
                    })
                    
                    total_avance_mes += avance
                    total_asesores += asesores
                    total_meta += meta
                    
                total_porcentaje = (total_avance_mes / total_meta) if total_meta > 0 else 0
                datos_area.insert(0, {
                    "Area/CAC": f"[-]{area} (TOTAL)",
                    "Avance Mes": total_avance_mes,
                    "Asesores": total_asesores,
                    "Meta": total_meta,
                    "Avance": total_porcentaje
                })
                    
                df_area = pd.DataFrame(datos_area)
                
                st.dataframe(
                    df_area.style.format({
                        "Avance": "{:.0%}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                st.markdown("---")
        else:
            st.error("La columna 'NOM_ESTRATEGIA' no se encontró en la base de datos.")

    except Exception as e:
        st.error(f"Hubo un problema al leer el archivo. Error técnico: {e}")
else:
    st.info("Sube el archivo de Excel para visualizar los tableros.")
