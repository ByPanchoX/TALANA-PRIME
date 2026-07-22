import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador de Remuneraciones", layout="wide")
st.title("Automatización de Consolidado de Remuneraciones")

# Parámetros del mes
col1, col2 = st.columns(2)
with col1:
    mes_input = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
with col2:
    anio_input = st.number_input("Año", min_value=2020, max_value=2050, value=2026)

rut_empresa = st.text_input("RUT de la Empresa", value="76.455.680-1")

st.markdown("---")

# UI Separada en 2 partes
st.markdown("### 1. Plantilla de Destino")
st.info("Sube el archivo con el formato final deseado (ej. tu archivo Consolidado_Final de diciembre). El sistema leerá las columnas de aquí.")
archivo_plantilla = st.file_uploader("Sube el archivo de Plantilla (Excel)", type=["xlsx", "xls"], key="plantilla")

st.markdown("### 2. Datos del Mes")
st.info("Sube todos los Libros de Remuneraciones y los Informes de Haberes/Descuentos del mes.")
archivos_datos = st.file_uploader("Sube los datos (Excel)", type=["xlsx", "xls"], accept_multiple_files=True, key="datos")

def procesar_informe_haberes(df):
    data = []
    categoria_actual = None
    for index, row in df.iterrows():
        col0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        rut = str(row.get('Rut', '')).strip() if 'Rut' in row and pd.notna(row['Rut']) else ""
        monto = row.get('Monto', 0)
        
        if col0 and col0.lower() != 'nan' and not col0.lower().startswith('total'):
            categoria_actual = col0
            
        if rut and rut.lower() != 'nan':
            if categoria_actual:
                data.append({'Rut Trabajador': rut.replace('.', ''), 'Categoria': categoria_actual, 'Monto': monto})
                
    if not data: 
        return pd.DataFrame()
    return pd.DataFrame(data).pivot_table(index='Rut Trabajador', columns='Categoria', values='Monto', aggfunc='sum').reset_index()

if st.button("Generar Consolidado"):
    if not archivo_plantilla:
        st.warning("Debes subir la plantilla base en el Paso 1.")
    elif not archivos_datos:
        st.warning("Debes subir los archivos de datos del mes en el Paso 2.")
    else:
        with st.spinner("Procesando archivos y cruzando datos..."):
            try:
                # 1. Extraer columnas dinámicamente de la plantilla subida
                df_plantilla = pd.read_excel(archivo_plantilla)
                columnas_objetivo = df_plantilla.columns.tolist()
                
                df_libros = []
                df_infs = []
                
                # 2. Clasificar y leer los datos del mes
                for archivo in archivos_datos:
                    if "Libro de Remuneraciones" in archivo.name:
                        df_libros.append(pd.read_excel(archivo))
                    elif "Informe Haberes" in archivo.name:
                        df_temp = pd.read_excel(archivo, skiprows=15)
                        df_infs.append(procesar_informe_haberes(df_temp))
                        
                if not df_libros:
                    st.error("Error: No se encontró ningún 'Libro de Remuneraciones' entre los archivos subidos.")
                else:
                    df_libro = pd.concat(df_libros, ignore_index=True)
                    df_inf = pd.concat(df_infs, ignore_index=True).groupby('Rut Trabajador').sum().reset_index() if df_infs else pd.DataFrame()
                    
                    # 3. Cruzar Libros con Informes
                    df_merge = pd.merge(df_libro, df_inf, on='Rut Trabajador', how='left') if not df_inf.empty else df_libro.copy()
                    df_final = pd.DataFrame(columns=columnas_objetivo)
                    
                    df_final['RUT EMPRESA'] = rut_empresa
                    df_final['MES'] = mes_input
                    df_final['AÑO'] = anio_input
                    
                    # 4. Mapeo explícito de variables base
                    map_libro = {
                        'Rut Trabajador': 'Rut Trabajador', 'Apellido Paterno': 'Apellido Paterno', 'Apellido Materno': 'Apellido Materno',
                        'Nombres': 'Nombres', 'Cargo': 'Cargo', 'Tipo de Contrato': 'Tipo de Contrato', 'N° Dias Trabajados': 'N° Dias Trabajados',
                        'N° Dias Ausentes': 'N° Dias Ausentes', 'N° Dias Licencia': 'N° Dias Licencia', 'N° Dias Accidentes de Trabajo': 'N° Dias Accidentes de Trabajo',
                        'N° Cargas Familiares': 'N° Cargas Familiares', 'Sueldo Base': 'Sueldo Base (H)', 'Imponible': 'Total Imponible (H)',
                        'Aporte Accidentes de Trabajo Mutual': 'Aporte Accidentes de Trabajo Mutual (P)', 'Seguro de Cesantía Empleador': 'Seguro de Cesantía Empleador (P)'
                    }
                    
                    for src, dst in map_libro.items():
                        if src in df_merge.columns and dst in df_final.columns: 
                            df_final[dst] = df_merge[src]
                            
                    # 5. Mapeo dinámico de Haberes y Descuentos
                    for col in columnas_objetivo:
                        if df_final[col].notna().any(): continue
                        
                        clean_col = str(col).replace(' (H)', '').replace(' (D)', '').replace(' (P)', '').strip()
                        
                        match = None
                        for m_col in df_merge.columns:
                            if str(m_col).lower() == clean_col.lower():
                                match = m_col
                                break
                        
                        if match:
                            df_final[col] = df_merge[match]
                            
                    # 6. SOLUCIÓN AL TYPEERROR: Filtrar relleno por tipo de columna
                    columnas_texto = ['RUT EMPRESA', 'MES', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 'Tipo de Contrato']
                    
                    for col in df_final.columns:
                        if col in columnas_texto or df_final[col].dtype == 'object':
                            df_final[col] = df_final[col].fillna("") # Rellenar textos con vacío
                        else:
                            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0) # Rellenar números con cero
                    
                    # 7. Generar Excel en memoria
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False)
                    output.seek(0)
                    
                    st.success("¡Consolidado Completo Generado Exitosamente!")
                    st.download_button(
                        label="📥 Descargar Consolidado", 
                        data=output, 
                        file_name=f"Consolidado_Final_{mes_input}_{anio_input}.xlsx"
                    )
                    
            except Exception as e:
                st.error(f"Error crítico durante el procesamiento: {str(e)}")
