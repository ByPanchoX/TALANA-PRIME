import streamlit as st
import pandas as pd
import unicodedata
import io

st.set_page_config(page_title="Consolidador de Remuneraciones", layout="wide")
st.title("Automatización de Consolidado de Remuneraciones (Formato Estricto Talana)")

col1, col2 = st.columns(2)
with col1:
    mes_input = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
with col2:
    anio_input = st.number_input("Año", min_value=2020, max_value=2050, value=2026)

rut_empresa = st.text_input("RUT de la Empresa", value="76.455.680-1")
st.markdown("---")

st.markdown("### 1. Plantilla de Destino")
st.info("Sube la plantilla exacta y original descargada desde Talana (ej. PLANILLA SUBIR1.xlsx).")
archivo_plantilla = st.file_uploader("Sube la Plantilla (Excel)", type=["xlsx", "xls"], key="plantilla")

st.markdown("### 2. Datos del Mes")
st.info("Sube AQUÍ todos los archivos del mes: Libros de Remuneraciones, Costos por Trabajador, e Informes de Haberes.")
archivos_datos = st.file_uploader("Sube los Datos Múltiples (Excel)", type=["xlsx", "xls"], accept_multiple_files=True, key="datos")

def normalize_str(s):
    if not isinstance(s, str):
        return ""
    s = s.replace(' (H)', '').replace(' (D)', '').replace(' (P)', '').replace(' *', '').strip()
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

if st.button("Generar Consolidado para Talana"):
    if not archivo_plantilla or not archivos_datos:
        st.warning("Faltan archivos por subir en los pasos 1 y 2.")
    else:
        with st.spinner("Procesando y cuadrando datos para Talana..."):
            try:
                # 1. Leer Plantilla Estricta
                df_plantilla_raw = pd.read_excel(archivo_plantilla, header=None)
                
                # Identificar la fila de encabezados (la que tiene los nombres de las columnas)
                header_row_idx = 2 if "Campos Obligatorios" in str(df_plantilla_raw.iloc[0, 0]) else 0
                cols_in_sheet = df_plantilla_raw.iloc[header_row_idx].tolist()
                
                # Guardar las dos primeras filas decorativas de Talana para pegarlas al final
                encabezados_talana = df_plantilla_raw.iloc[:header_row_idx].copy()
                
                # 2. Leer Libros y Haberes
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
                    st.stop()
                    
                df_base = pd.concat(df_bases, ignore_index=True)
                df_base = df_base.groupby('Rut Trabajador').first().reset_index()
                
                df_inf = pd.concat(df_infs, ignore_index=True).groupby('Rut Trabajador').sum().reset_index() if df_infs else pd.DataFrame()
                if not df_inf.empty:
                    cols_duplicadas = [c for c in df_inf.columns if c in df_base.columns and c != 'Rut Trabajador']
                    df_inf = df_inf.drop(columns=cols_duplicadas)
                    
                df_merge = pd.merge(df_base, df_inf, on='Rut Trabajador', how='left') if not df_inf.empty else df_base.copy()
                
                # 3. Crear DataFrame Final basado estrictamente en las columnas de la plantilla
                df_final = pd.DataFrame(index=df_merge.index, columns=cols_in_sheet)
                
                # Variables Base
                meses_num = {'Enero':'01','Febrero':'02','Marzo':'03','Abril':'04','Mayo':'05','Junio':'06','Julio':'07','Agosto':'08','Septiembre':'09','Octubre':'10','Noviembre':'11','Diciembre':'12'}
                
                if 'RUT EMPRESA' in df_final.columns: df_final['RUT EMPRESA'] = rut_empresa
                if 'Rut Razón Social *' in df_final.columns: df_final['Rut Razón Social *'] = rut_empresa
                if 'Año_Mes (aaaamm) *' in df_final.columns: df_final['Año_Mes (aaaamm) *'] = f"{anio_input}{meses_num[mes_input]}"
                
                # 4. Mapeo Universal de Conceptos
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
                    'Anticipos': ['Descuento Anticipo', 'Anticipos (D)', 'Anticipos'],
                    'Movilizacion': ['Movilización', 'Movilización (H)'],
                    'Colacion': ['Colación', 'Colación (H)'],
                    'Cargas Familiares': ['Asignación familiar y Maternal', 'Cargas Familiares Normales (H)']
                }
                
                # Realizar el cruce de datos
                for src, dst_list in map_universal.items():
                    if src in df_merge.columns:
                        for dst in dst_list:
                            if dst in df_final.columns:
                                df_final[dst] = df_merge[src]
                                
                # Mapeo dinámico para todos los demás bonos
                for col in cols_in_sheet:
                    if pd.isna(col) or df_final[col].notna().any(): continue
                    clean_col = normalize_str(col)
                    
                    match = None
                    for m_col in df_merge.columns:
                        if normalize_str(m_col) == clean_col:
                            match = m_col
                            break
                    if match:
                        df_final[col] = df_merge[match]

                # 5. Formateo y limpieza anti-nulos
                for col in df_final.columns:
                    if str(col).startswith('Rut') or str(col).startswith('Año'):
                        df_final[col] = df_final[col].fillna("")
                    else:
                        df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)
                        
                # 6. Reconstruir la estructura física estricta de Talana (Añadir headers decorativos)
                df_export = pd.DataFrame(columns=cols_in_sheet)
                
                # Insertar filas decorativas
                if not encabezados_talana.empty:
                    df_export.loc[0] = encabezados_talana.iloc[0].values
                    if header_row_idx > 1:
                        df_export.loc[1] = encabezados_talana.iloc[1].values
                        
                # Insertar los nombres de las columnas en la fila correspondiente
                df_export.loc[header_row_idx] = cols_in_sheet
                
                # Pegar la data real
                df_final.columns = df_export.columns
                df_export = pd.concat([df_export, df_final], ignore_index=True)
                
                # 7. Descargar
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, header=False) # Guardamos sin el header interno de Pandas
                output.seek(0)
                
                st.success("¡Datos cruzados exitosamente! La estructura y los montos están listos.")
                st.download_button(
                    label="📥 Descargar Subida Talana", 
                    data=output, 
                    file_name=f"Subida_Talana_Calculado_{mes_input}_{anio_input}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error procesando la información: {str(e)}")
