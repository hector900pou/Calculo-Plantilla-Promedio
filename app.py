import streamlit as st
import pandas as pd
import os

# --- CONFIGURACIÓN DE LA PÁGINA (Aspecto Profesional) ---
st.set_page_config(
    page_title="Control de Plantilla - Conductores",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo visual minimalista mediante títulos y contenedores
st.title("🚛 Cuadro de Mando de Plantilla Promedio")
st.markdown("Consolida los reportes de conductores por bases, calcula la plantilla y genera informes listos para descargar.")
st.hr()

# --- BARRA LATERAL (Configuración del Período) ---
st.sidebar.header("🗓️ Configuración del Período")

nombres_meses = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# Selector interactivo del mes en la barra lateral
mes_texto = st.sidebar.selectbox("Selecciona el mes a procesar:", list(nombres_meses.values()))
num_mes = [k for k, v in nombres_meses.items() if v == mes_texto][0]

# Regla de negocio de 31 días
meses_31_dias = [1, 3, 5, 7, 8, 10, 12]
es_mes_31 = num_mes in meses_31_dias

# Indicador informativo en la barra lateral
if es_mes_31:
    st.sidebar.info(f"💡 {mes_texto} tiene 31 días. Los conductores con días distintos a 30 se dividirán entre 31.")
else:
    st.sidebar.info(f"💡 {mes_texto} se procesará con base estándar de 30 días.")

# --- CUERPO PRINCIPAL: Carga de Archivos Excel ---
st.subheader("📂 1. Carga tus archivos de bases")
uploaded_files = st.file_uploader(
    "Arrastra o selecciona todos los archivos Excel de tus bases simultáneamente:",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

datos_consolidados = []

if uploaded_files:
    st.subheader("🔄 2. Estado del Procesamiento")
    
    # Procesar archivos cargados
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        try:
            df = pd.read_excel(uploaded_file)
            
            # Identificar dinámicamente las columnas
            nombre_col = [col for col in df.columns if 'nombre' in col.lower() or 'nom' in col.lower()]
            dias_col = [col for col in df.columns if 'dia' in col.lower() or 'días' in col.lower()]
            
            if not nombre_col or not dias_col:
                st.warning(f"⚠️ Columnas requeridas no encontradas en '{filename}'. Archivo omitido.")
                continue
                
            col_nombre = nombre_col[0]
            col_dias = dias_col[0]
            
            # Nombre de la base limpia sin extensión
            nombre_base = os.path.splitext(filename)[0]
            
            # Limpieza de datos en memoria
            df_limpio = df.dropna(subset=[col_nombre, col_dias]).copy()
            df_limpio[col_dias] = pd.to_numeric(df_limpio[col_dias], errors='coerce')
            df_limpio = df_limpio.dropna(subset=[col_dias])
            
            # Aplicar lógica matemática de negocio
            plantillas_calculadas = []
            for dias in df_limpio[col_dias]:
                if dias == 30:
                    divisor = 30
                elif es_mes_31:
                    divisor = 31
                else:
                    divisor = 30
                plantillas_calculadas.append(dias / divisor)
                
            df_limpio['Plantilla Individual'] = plantillas_calculadas
            
            # Calcular agregaciones
            total_conductores = int(df_limpio[col_nombre].count())
            plantilla_promedio_base = float(df_limpio['Plantilla Individual'].sum())
            
            datos_consolidados.append({
                'Base': nombre_base,
                'Número de Conductores': total_conductores,
                'Plantilla Promedio': round(plantilla_promedio_base, 2)
            })
            
        except Exception as e:
            st.error(f"❌ Error crítico procesando '{filename}': {e}")
            
    # --- RESULTADOS Y DESCARGA ---
    if datos_consolidados:
        st.success("¡Todos los archivos se han procesado de forma correcta!")
        
        # Crear DataFrame final ordenado
        resultado_final = pd.DataFrame(datos_consolidados)
        resultado_final = resultado_final.sort_values(by='Base').reset_index(drop=True)
        
        st.subheader("📊 3. Resumen Consolidado")
        
        # Tarjetas de KPI profesionales (Métricas totales)
        kpi1, kpi2 = st.columns(2)
        with kpi1:
            st.metric("Total Conductores (Todas las bases)", int(resultado_final['Número de Conductores'].sum()))
        with kpi2:
            st.metric("Plantilla Promedio Global", round(float(resultado_final['Plantilla Promedio'].sum()), 2))
            
        # Mostrar tabla interactiva en la web
        st.dataframe(resultado_final, use_container_width=True)
        
        # Generar la descarga del Excel sin guardarlo físicamente en el servidor
        import io
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            resultado_final.to_excel(writer, index=False, sheet_name='Resumen Plantilla')
        
        st.markdown("### 💾 4. Descargar Resultados")
        st.download_button(
            label="📥 Descargar Reporte Consolidado en Excel",
            data=buffer.getvalue(),
            file_name=f"resumen_plantilla_{mes_texto.lower()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary" # Destaca el botón en azul profesional
        )
else:
    # Estado inicial cuando no se han subido archivos
    st.info("👋 Por favor, configura el mes en la barra lateral y sube tus documentos Excel para comenzar.")
