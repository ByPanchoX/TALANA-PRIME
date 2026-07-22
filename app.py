import streamlit as st
import pandas as pd
import unicodedata
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

st.markdown("### 1. Plantilla de Destino")
st.info("Sube 'PLANILLA SUBIREEEEE.xlsx' o el Consolidado_Final.")
archivo_plantilla = st.file_uploader("Sube la Plantilla", type=["xlsx", "xls"], key="plantilla")

st.markdown("### 2. Datos del Mes")
st.info("Sube AQUÍ todos los: Libros de Remuneraciones, Costos por Trabajador, e Informes de Haberes.")
archivos_datos = st.file_uploader("Sube los Datos", type=["xlsx", "xls"], accept_multiple_files=True, key="datos")

def normalize_str(s):
    """Normaliza textos quitando asteriscos, sufijos y tildes para hacer cruces exactos."""
    s = str(s).replace(' (H)', '').replace(' (D)', '').replace(' (P)', '').replace(' *', '').strip()
    return unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8').lower()

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
                
    if not data: return pd.DataFrame()
    return pd.DataFrame(data).pivot_table(index='Rut Trabajador', columns='Categoria', values='Monto', aggfunc='sum').reset_index()

if st.button("Generar Consolidado"):
    if not archivo_plantilla or not archivos_datos:
        st.warning("Faltan archivos por subir en los pasos 1 y 2.")
    else:
        with st.spinner("Sincronizando todas las fuentes de datos..."):
            try:
                # 1. Leer Plantilla
                df_plantilla = pd.read_excel(archivo_plantilla)
                if "Campos Obligatorios" in str(df_plantilla.columns[0]):
                    df_plantilla = pd.read_excel(archivo_plantilla, header=2)
                columnas_objetivo = df_plantilla.columns.tolist()
                
                # 2. Leer Libros y Costos (El archivo de Costos tiene las leyes patronales que le faltan al Libro)
                df_bases = []
                df_infs = []
                
                for archivo in archivos_datos:
                    if "Libro" in archivo.name or "Costo" in archivo.name:
                        df_bases.append(pd.read_excel(archivo))
                    elif "Informe Haberes" in archivo.name:
                        df_temp = pd.read_excel(archivo, skiprows=15)
                        df_infs.append(procesar_informe_haberes(df_temp))
                
                if not df_bases:
                    st.error("Error: No se subió ningún Libro de Remuneraciones ni Costo por Trabajador.")
                else:
                    # Agrupar por RUT para unir la info del Libro y del Costo por Trabajador
                    df_base = pd.concat(df_bases, ignore_index=True)
                    df_base = df_base.groupby('Rut Trabajador').first().reset_index()
                    
                    df_inf = pd.concat(df_infs, ignore_index=True).groupby('Rut Trabajador').sum().reset_index() if df_infs else pd.DataFrame()
                    df_merge = pd.merge(df_base, df_inf, on='Rut Trabajador', how='left') if not df_inf.empty else df_base.copy()
                    
                    # 3. Inicializar el DF final con las columnas correctas
                    df_final = pd.DataFrame(index=df_merge.index, columns=columnas_objetivo)
                    
                    # RUT Empresa y Fecha
                    if 'RUT EMPRESA' in df_final.columns: df_final['RUT EMPRESA'] = rut_empresa
                    if 'Rut Razón Social *' in df_final.columns: df_final['Rut Razón Social *'] = rut_empresa
                    if 'MES' in df_final.columns: df_final['MES'] = mes_input
                    if 'AÑO' in df_final.columns: df_final['AÑO'] = anio_input
                    
                    meses_num = {'Enero':'01','Febrero':'02','Marzo':'03','Abril':'04','Mayo':'05','Junio':'06','Julio':'07','Agosto':'08','Septiembre':'09','Octubre':'10','Noviembre':'11','Diciembre':'12'}
                    if 'Año_Mes (aaaamm) *' in df_final.columns: df_final['Año_Mes (aaaamm) *'] = f"{anio_input}{meses_num[mes_input]}"
                    
                    # 4. MAPEO UNIVERSAL (Funciona con la Plantilla Subire y con el Consolidado Final Antiguo)
                    map_universal = {
                        'Rut Trabajador': ['Rut Trabajador', 'Rut *'],
                        'Apellido Paterno': ['Apellido Paterno'],
                        'Apellido Materno': ['Apellido Materno'],
                        'Nombres': ['Nombres'],
                        'Cargo': ['Cargo'],
                        'Sueldo Base': ['Sueldo Base (H)', 'Sueldo Base *'],
                        'N° Dias Trabajados': ['N° Dias Trabajados', 'Días Trabajados del Mes *'],
                        'N° Dias Ausentes': ['N° Dias Ausentes', 'Días de Ausentismo *'],
                        'N° Dias Licencia': ['N° Dias Licencia', 'Días de licencia médica *'],
                        'Imponible': ['Total Imponible (H)', 'Remuneración Imponible *'],
                        'Base Tributable': ['Total Tributable (afecta a impuesto) *'],
                        'Líquido': ['Sueldo Liquido *', 'Alcance Liquido'],
                        'Total Haberes': ['Remuneración Total o Sueldo Bruto *'],
                        'Haberes No Imponibles': ['Remuneración No Imponible *'],
                        'Impuesto Unico': ['Impuesto Unico (D)', 'Impuestos *'],
                        'Salud': ['Cotización Institución de Salud (D)', 'Total Isapre *'],
                        'Prevision': ['Descuento Cotización AFP (D)', 'AFP *'],
                        'Seguro de Cesantía': ['Descuento Seguro Cesantia (D)', 'Seguro Cesantía Trabajador *'],
                        'APV': ['APV', 'APV (Régimen B)'],
                        'Aporte Accidentes de Trabajo Mutual': ['Aporte Accidentes de Trabajo Mutual (P)', 'Seguro Accidente de Trabajo *'],
                        'Seguro de Cesantía Empleador': ['Seguro de Cesantía Empleador (P)', 'Seguro Cesantía Empleador *'],
                        'Seguro de Invalidez y Sobrevivencia Empleador': ['Seguro de Invalidez y Sobrevivencia Empleador (P)', 'Seguro Invalidez y Supervivencia (SIS) *'],
                        'Descuentos CCAF': ['Descuento Crédito Personal CCAF', 'Créditos Personales CCAF (D)'],
                        'Anticipos': ['Descuento Anticipo', 'Anticipos (D)'],
                        'Movilizacion': ['Movilización', 'Movilización (H)'],
                        'Colacion': ['Colación', 'Colación (H)'],
                        'Cargas Familiares': ['Asignación familiar y Maternal', 'Cargas Familiares Normales (H)']
                    }
                    
                    # Aplicar Diccionario Base
                    for src, dst_list in map_universal.items():
                        if src in df_merge.columns:
                            for dst in dst_list:
                                if dst in df_final.columns:
                                    df_final[dst] = df_merge[src]
                                    
                    # 5. Mapeo Dinámico (Para atrapar Bonos, Viáticos, y el resto de la lista)
                    for col in columnas_objetivo:
                        if df_final[col].notna().any(): continue
                        clean_col = normalize_str(col)
                        
                        match = None
                        for m_col in df_merge.columns:
                            if normalize_str(m_col) == clean_col:
                                match = m_col
                                break
                                
                        if match:
                            df_final[col] = df_merge[match]
                            
                    # 6. Limpieza final (Textos vacíos, números en 0)
                    columnas_texto = ['RUT EMPRESA', 'MES', 'Rut Trabajador', 'Rut *', 'Rut Razón Social *', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 'Tipo de Contrato', 'Año_Mes (aaaamm) *']
                    for col in df_final.columns:
                        if col in columnas_texto or df_final[col].dtype == 'object':
                            df_final[col] = df_final[col].fillna("")
                        else:
                            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)
                            
                    # 7. Descargar
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False)
                    output.seek(0)
                    
                    st.success("¡Datos cruzados exitosamente! Todo listo para subir al sistema.")
                    st.download_button(
                        label="📥 Descargar Excel Listo", 
                        data=output, 
                        file_name=f"Subida_Sistema_{mes_input}_{anio_input}.xlsx"
                    )
            except Exception as e:
                st.error(f"Error procesando la información: {str(e)}")
