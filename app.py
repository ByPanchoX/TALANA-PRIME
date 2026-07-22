import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Consolidador de Remuneraciones", layout="wide")
st.title("Automatización de Consolidado de Remuneraciones")
st.write("Sube los archivos del mes (Libros de Remuneraciones, Costo por Trabajador, etc.) para generar el consolidado.")

# Interfaz de parámetros
col1, col2 = st.columns(2)
with col1:
    mes_input = st.selectbox("Mes a Procesar", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
with col2:
    anio_input = st.number_input("Año", min_value=2020, max_value=2050, value=2026)

rut_empresa_input = st.text_input("RUT de la Empresa", value="76.455.680-1")

# Subida múltiple de archivos
archivos_subidos = st.file_uploader("Selecciona los archivos Excel del mes", type=["xlsx", "xls"], accept_multiple_files=True)

def procesar_archivos(archivos, mes, anio, rut_empresa):
    df_list = []
    
    # Extraeremos la data principal de los Libros de Remuneraciones
    # ya que contienen el desglose tabular más completo
    for archivo in archivos:
        if "Libro de Remuneraciones" in archivo.name:
            try:
                df = pd.read_excel(archivo, sheet_name=0)
                df_list.append(df)
            except Exception as e:
                st.error(f"Error al leer {archivo.name}: {e}")
                
    if not df_list:
        return None
        
    df_consolidado = pd.concat(df_list, ignore_index=True)
    df_final = pd.DataFrame()
    
    # Estructura del Archivo: RUT Empresa estrictamente al inicio
    df_final['RUT EMPRESA'] = rut_empresa
    df_final['MES'] = mes
    df_final['AÑO'] = anio
    
    # Mapeo de Datos del Trabajador
    df_final['Rut Trabajador'] = df_consolidado.get('Rut Trabajador', '')
    df_final['Apellido Paterno'] = df_consolidado.get('Apellido Paterno', '')
    df_final['Apellido Materno'] = df_consolidado.get('Apellido Materno', '')
    df_final['Nombres'] = df_consolidado.get('Nombres', '')
    df_final['Cargo'] = df_consolidado.get('Cargo', '')
    df_final['Tipo de Contrato'] = df_consolidado.get('Tipo de Contrato', 'Indefinido')
    
    # Asistencia
    df_final['N° Dias Trabajados'] = df_consolidado.get('N° Dias Trabajados', 0)
    df_final['N° Dias Ausentes'] = df_consolidado.get('N° Dias Ausentes', 0)
    df_final['N° Dias Licencia'] = df_consolidado.get('N° Dias Licencia', 0)
    df_final['N° Dias Accidentes de Trabajo'] = df_consolidado.get('N° Dias Accidentes de Trabajo', 0)
    df_final['N° Dias No Contratado'] = df_consolidado.get('N° Dias No Contratado', 0)
    df_final['N° Cargas Familiares'] = df_consolidado.get('N° Cargas Familiares', 0)
    
    # Haberes
    df_final['Sueldo Base (H)'] = df_consolidado.get('Sueldo Base', 0)
    df_final['Gratificación Legal (H)'] = df_consolidado.get('Gratificacion', 0)
    df_final['Movilización (H)'] = df_consolidado.get('Movilizacion', 0)
    df_final['Colación (H)'] = df_consolidado.get('Colacion', 0)
    df_final['Total Imponible (H)'] = df_consolidado.get('Imponible', 0)
    
    # Descuentos del Trabajador
    df_final['Descuento Cotización AFP (D)'] = df_consolidado.get('Prevision', 0)
    df_final['Cotización Institución de Salud (D)'] = df_consolidado.get('Salud', 0)
    df_final['Descuento Seguro Cesantia (D)'] = df_consolidado.get('Seguro de Cesantia', 0)
    df_final['Impuesto Unico (D)'] = df_consolidado.get('Impuesto Unico', 0)
    df_final['Anticipos (D)'] = df_consolidado.get('Anticipos', 0)
    
    # Aportes y Seguros Patronales
    df_final['Aporte Accidentes de Trabajo Mutual (P)'] = df_consolidado.get('Aporte Accidentes de Trabajo Mutual', 0)
    df_final['Aporte Accidentes de Trabajo IPS (P)'] = df_consolidado.get('Aporte Accidentes de Trabajo IPS', 0)
    df_final['Seguro de Cesantía Empleador (P)'] = df_consolidado.get('Seguro de Cesantía Empleador', 0)
    df_final['Seguro de Invalidez y Sobrevivencia Empleador (P)'] = df_consolidado.get('Seguro de Invalidez y Sobrevivencia Empleador', 0)
    
    return df_final

if st.button("Generar Consolidado"):
    if archivos_subidos:
        with st.spinner("Procesando y cruzando archivos..."):
            df_resultado = procesar_archivos(archivos_subidos, mes_input, anio_input, rut_empresa_input)
            
            if df_resultado is not None:
                st.success("¡Consolidado generado exitosamente!")
                st.dataframe(df_resultado.head())
                
                # Exportar a Excel en memoria para descargar
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Sheet1')
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Descargar Consolidado Final (Excel)",
                    data=output,
                    file_name=f"Consolidado_Final_{mes_input}_{anio_input}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No se encontraron Libros de Remuneraciones válidos para procesar los datos.")
    else:
        st.error("Por favor, sube los archivos de Excel para comenzar.")
