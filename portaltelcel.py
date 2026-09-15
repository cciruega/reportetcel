import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Dashboard de Resultados", layout="wide")

# 1. Diccionarios de configuración
estructura_cac = {
    "CIUDAD VICTORIA": [
        "2008604 TCC CIU100 MANTE4",
        "2008604 TCC CIU100 VICTORIA II4",
        "2008604 TCC CIU100 VICTORIA4"
    ],
    "MATAMOROS-REYNOSA": [
        "2008604 TCC MAT101 MATAMOROS II4",
        "2008604 TCC MAT101 MATAMOROS4",
        "2008604 TCC REY115 REYNOSA II4",
        "2008604 TCC REY115 REYNOSA III4",
        "2008604 TCC REY116 REYNOSA I4",
        "2008604 TCC REY116 REYNOSA IV4"
    ],
    "MONTERREY 1": [
        "2008604 TCC MON103 COUNTRY4",
        "2008604 TCC MON103 EXPRESS ESFERA4",
        "2008604 TCC MON103 EXPRESS NUEVO SUR4",
        "2008604 TCC MON103 SATELITE4",
        "2008604 TCC MON103 VALLE ORIENTE4",
        "2008604 TCC MON104 CUMBRES4",
        "2008604 TCC MON104 SENDERO LINCOLN4",
        "2008604 TCC MON104 SERVICIO TECNICO Tlc Y CENTRo",
        "2008604 TCC MON105 CENTRIKA4",
        "2008604 TCC MON105 GALERIAS4",
        "2008604 TCC MON106 CENTRO4",
        "2008604 TCC MON106 EXPRESS FASHION DRIVE4",
        "2008604 TCC MON106 EXPRESS HUMBERTO LOBO4",
        "2008604 TCC MON106 EXPRESS PASEO TEC",
        "2008604 TCC MON106 EXPRESS VILLAS VALLE4",
        "2008604 TCC MON106 PUNTO VALLE4",
        "2008604 TCC MON106 SAN AGUSTIN4"
    ],
    "MONTERREY 2": [
        "2008604 TCC MON107 EXPOSICION4",
        "2008604 TCC MON107 GUADALUPE4",
        "2008604 TCC MON108 ANAHUAC4",
        "2008604 TCC MON108 EXPRESS PLAZA FIESTA ANAHUAC4",
        "2008604 TCC MON108 PLAZA BELLA4",
        "2008604 TCC MON108 SANTA CATARINA4",
        "2008604 TCC MON109 CITADEL4",
        "2008604 TCC MON109 LAS AMERICAS4",
        "2008604 TCC MON110 ESCOBEDO4"
    ],
    "MONTERREY 3": [
        "2008604 TCC MON111 APODACA4",
        "2008604 TCC MON111 MTY SUN MALL VIP4",
        "2008604 TCC MON112 EXPRESS MONTEMORELOS"
    ],
    "NUEVO LAREDO": [
        "2008604 TCC NUE114 LAREDO I4",
        "2008604 TCC NUE114 LAREDO II4"
    ],
    "TAMPICO": [
        "2008604 TCC TAM121 TAMPICO I4",
        "2008604 TCC TAM122 TAMPICO II4",
        "2008604 TCC TAM122 TAMPICO III4",
        "2008604 TCC TAM122 TAMPICO IV4"
    ]
}

# Catálogo completo de Asesores por CAC
catalogo_asesores = {
    # CIUDAD VICTORIA
    "2008604 TCC CIU100 MANTE4": 17,
    "2008604 TCC CIU100 VICTORIA II4": 19,
    "2008604 TCC CIU100 VICTORIA4": 28,
    # MATAMOROS-REYNOSA
    "2008604 TCC MAT101 MATAMOROS II4": 21,
    "2008604 TCC MAT101 MATAMOROS4": 19,
    "2008604 TCC REY115 REYNOSA II4": 15,
    "2008604 TCC REY115 REYNOSA III4": 20,
    "2008604 TCC REY116 REYNOSA I4": 14,
    "2008604 TCC REY116 REYNOSA IV4": 16,
    # MONTERREY 1
    "2008604 TCC MON103 COUNTRY4": 22,
    "2008604 TCC MON103 EXPRESS ESFERA4": 10,
    "2008604 TCC MON103 EXPRESS NUEVO SUR4": 7,
    "2008604 TCC MON103 SATELITE4": 14,
    "2008604 TCC MON103 VALLE ORIENTE4": 14,
    "2008604 TCC MON104 CUMBRES4": 29,
    "2008604 TCC MON104 SENDERO LINCOLN4": 30,
    "2008604 TCC MON104 SERVICIO TECNICO Tlc Y CENTRo": 19,
    "2008604 TCC MON105 CENTRIKA4": 15,
    "2008604 TCC MON105 GALERIAS4": 25,
    "2008604 TCC MON106 CENTRO4": 24,
    "2008604 TCC MON106 EXPRESS FASHION DRIVE4": 6,
    "2008604 TCC MON106 EXPRESS HUMBERTO LOBO4": 13,
    "2008604 TCC MON106 EXPRESS PASEO TEC": 7,
    "2008604 TCC MON106 EXPRESS VILLAS VALLE4": 4,
    "2008604 TCC MON106 PUNTO VALLE4": 11,
    "2008604 TCC MON106 SAN AGUSTIN4": 18,
    # MONTERREY 2
    "2008604 TCC MON107 EXPOSICION4": 20,
    "2008604 TCC MON107 GUADALUPE4": 28,
    "2008604 TCC MON108 ANAHUAC4": 20,
    "2008604 TCC MON108 EXPRESS PLAZA FIESTA ANAHUAC4": 8,
    "2008604 TCC MON108 PLAZA BELLA4": 27,
    "2008604 TCC MON108 SANTA CATARINA4": 23,
    "2008604 TCC MON109 CITADEL4": 25,
    "2008604 TCC MON109 LAS AMERICAS4": 20,
    "2008604 TCC MON110 ESCOBEDO4": 20,
    # MONTERREY 3
    "2008604 TCC MON111 APODACA4": 26,
    "2008604 TCC MON111 MTY SUN MALL VIP4": 20,
    "2008604 TCC MON112 EXPRESS MONTEMORELOS": 7,
    # NUEVO LAREDO
    "2008604 TCC NUE114 LAREDO I4": 9,
    "2008604 TCC NUE114 LAREDO II4": 15,
    # TAMPICO
    "2008604 TCC TAM121 TAMPICO I4": 34,
    "2008604 TCC TAM122 TAMPICO II4": 20,
    "2008604 TCC TAM122 TAMPICO III4": 23,
    "2008604 TCC TAM122 TAMPICO IV4": 23
}

st.title("Portal de Resultados Operativos")

# 2. Carga del archivo
archivo_subido = st.file_uploader("Sube el archivo Excel con la base de datos extraída (Hoja 'Base')", type=["xlsx", "xls"])

if archivo_subido:
    # Leemos la base de datos desde la pestaña guardada
    df = pd.read_excel(archivo_subido, sheet_name="Base") 
    
    # 3. Filtros en la barra lateral
    st.sidebar.header("Filtros Principales")
    
    # Asegurar que MES_CAPTURA se trate de forma correcta
    if 'MES_CAPTURA' in df.columns:
        meses_disponibles = sorted(df['MES_CAPTURA'].dropna().unique().tolist())
        mes_actual = datetime.now().month
        
        # Intentar seleccionar el mes actual por defecto
        default_index = meses_disponibles.index(mes_actual) if mes_actual in meses_disponibles else 0
        mes_seleccionado = st.sidebar.selectbox("MES_CAPTURA", meses_disponibles, index=default_index)
        
        df_filtrado = df[df['MES_CAPTURA'] == mes_seleccionado]
    else:
        st.error("La columna 'MES_CAPTURA' no se encontró en la base de datos.")
        df_filtrado = df
    
    # Filtro de Estrategia
    if 'NOM_ESTRATEGIA' in df.columns:
        estrategias_disponibles = ["Todas"] + sorted(df['NOM_ESTRATEGIA'].dropna().unique().tolist())
        estrategia_seleccionada = st.sidebar.selectbox("NOM_ESTRATEGIA", estrategias_disponibles)
        
        if estrategia_seleccionada != "Todas":
            df_filtrado = df_filtrado[df_filtrado['NOM_ESTRATEGIA'] == estrategia_seleccionada]
    else:
        st.error("La columna 'NOM_ESTRATEGIA' no se encontró en la base de datos.")
        
    # Calcular el "Avance Mes" agrupando por CAC
    # NOTA: Cambia 'CAC' y 'Avance' por los nombres exactos de las columnas en tu Excel extraído
    columna_cac = 'CAC'       # Reemplaza si en tu excel se llama diferente (ej. 'Area/CAC')
    columna_avance = 'Avance' # Reemplaza si en tu excel se llama diferente
    
    if columna_cac in df_filtrado.columns and columna_avance in df_filtrado.columns:
        resumen_cacs = df_filtrado.groupby(columna_cac)[columna_avance].sum().reset_index()
        resumen_cacs.rename(columns={columna_avance: 'Avance Mes'}, inplace=True)
    
        # 4. Generación de las 7 Tablas
        st.header("Resultados por AREA_PDV")
        
        for area, cacs in estructura_cac.items():
            st.subheader(area)
            
            datos_area = []
            total_avance_mes = 0
            total_asesores = 0
            total_meta = 0
            
            for cac in cacs:
                # Obtener avance del DataFrame filtrado (0 si no hay coincidencias)
                avance_serie = resumen_cacs.loc[resumen_cacs[columna_cac] == cac, 'Avance Mes']
                avance = avance_serie.sum() if not avance_serie.empty else 0
                
                # Obtener asesores del diccionario estático
                asesores = catalogo_asesores.get(cac, 0)
                meta = asesores * 2
                
                # Cálculo de cumplimiento
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
                
            # Fila de Totales por Área
            total_porcentaje = (total_avance_mes / total_meta) if total_meta > 0 else 0
            datos_area.insert(0, {
                "Area/CAC": f"[-]{area} (TOTAL)",
                "Avance Mes": total_avance_mes,
                "Asesores": total_asesores,
                "Meta": total_meta,
                "Avance": total_porcentaje
            })
                
            df_area = pd.DataFrame(datos_area)
            
            # Mostrar la tabla formateando el Avance como porcentaje sin decimales
            st.dataframe(
                df_area.style.format({
                    "Avance": "{:.0%}"
                }),
                use_container_width=True,
                hide_index=True
            )
            st.markdown("---")
    else:
        st.error(f"Por favor, asegúrate de que tu Excel contenga las columnas '{columna_cac}' y '{columna_avance}'. Modifica las variables en el código si los nombres son diferentes.")
else:
    st.info("Sube el archivo de Excel extraído para visualizar los tableros.")