import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador de Remuneraciones", layout="wide")
st.title("Automatización de Consolidado de Remuneraciones")

# Parámetros
col1, col2 = st.columns(2)
with col1:
    mes_input = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
with col2:
    anio_input = st.number_input("Año", min_value=2020, max_value=2050, value=2026)

rut_empresa = st.text_input("RUT de la Empresa", value="76.455.680-1")

archivos_subidos = st.file_uploader("Sube Libros de Remuneraciones e Informes de Haberes/Descuentos (Excel)", type=["xlsx", "xls"], accept_multiple_files=True)

# Lista completa de las 93 columnas objetivo
COLUMNAS_OBJETIVO = [
    'RUT EMPRESA', 'MES', 'AÑO', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 'Tipo de Contrato', 
    'N° Dias Trabajados', 'N° Dias Ausentes', 'N° Dias Licencia', 'N° Dias Accidentes de Trabajo', 'N° Dias No Contratado', 'N° Cargas Familiares', 
    'AGUINALDO (H)', 'ASIGNACION DE VIATICO (H)', 'ASIGNACION VIATICO MEU (H)', 'Aguinaldo (H)', 'Asignacion Familiar (H)', 'Asignación Capacitación (H)', 
    'Asignación Navidad (H)', 'Asignación de Cargo (H)', 'Asignación de Telefono (H)', 'BONO ADMINISTRACION (H)', 'BONO PMI (H)', 'BONO RECONOCIMIENTO (H)', 
    'Bono Calidad (H)', 'Bono Extraordinario (H)', 'Bono Producción (H)', 'Bono Reconocimiento (H)', 'Bono Vacaciones (H)', 'Bono años de servicio (H)', 
    'Bono de Faena (H)', 'Bono de Productividad (H)', 'Bono por Asistencia (H)', 'Cargas Familiares Normales (H)', 'Colación (H)', 'Compensacion dia feriado (H)', 
    'DIAS PARO (H)', 'Gastos Reembolsables (H)', 'Gratificación Legal (H)', 'HORAS EXTRAS DEL MES (H)', 'ICTP (H)', 'Incentivo de Desempeño (H)', 'Internet (H)', 
    'Llamado de Emergencia (H)', 'MOVILIZACION ENAP (H)', 'Movilización (H)', 'Otros Haberes no Imponibles (H)', 'Reconocimiento Mentoria (H)', 'Sobretiempo (H)', 
    'Sobretiempo Festivo (H)', 'Sueldo Base (H)', 'Total Imponible (H)', 'Viatico Alojamiento (H)', 'Visita Adicional (H)', 'ANTICIPO AGUINALDO (D)', 
    'Anticipo Aguinaldo (D)', 'Anticipo Asignación Capacitación (D)', 'Anticipo Asignación Navidad (D)', 'Anticipo Bono Calidad (D)', 'Anticipo Bono Producción (D)', 
    'Anticipo Bono Vacaciones (D)', 'Anticipo bono años de servicio (D)', 'Anticipos (D)', 'Cotización Institución de Salud (D)', 'Cotización Voluntaria a Isapre (D)', 
    'Crédito Fonasa (D)', 'Créditos Personales CCAF (D)', 'Cuenta de Ahorro AFP (D)', 'Descuento Cotización AFP (D)', 'Descuento Seguro Cesantia (D)', 
    'Descuento cuota Sindicato (D)', 'Descuentos por Leasing (D)', 'Días Ausentes (D)', 'Días Licencia (D)', 'Días No contratado (D)', 
    'Días de Licencia Accidentes de Trabajo (D)', 'Gastos Particulares (D)', 'Gastos Reembolsables (D)', 'Impuesto Unico (D)', 'Otros Descuentos (D)', 
    'Prestamo Empresa (D)', 'Retencion Pension de Alimentos (D)', 'Seguro Vida Camara (D)', 'Seguro de vida CCAF (D)', 'Adicional de Capitalización Individual AFP (P)', 
    'Aporte Accidentes de Trabajo IPS (P)', 'Aporte Accidentes de Trabajo Mutual (P)', 'Expectativa de Vida Seguro Social (P)', 'Seguro de Cesantía Empleador (P)', 
    'Seguro de Invalidez y Sobrevivencia Empleador (P)'
]

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
    df_libros = []
    df_infs = []
    
    for archivo in archivos_subidos:
        if "Libro de Remuneraciones" in archivo.name:
            df_libros.append(pd.read_excel(archivo))
        elif "Informe Haberes y Descuentos" in archivo.name:
            df_temp = pd.read_excel(archivo, skiprows=15)
            df_infs.append(procesar_informe_haberes(df_temp))
            
    if not df_libros:
        st.warning("Falta cargar los Libros de Remuneraciones.")
    else:
        df_libro = pd.concat(df_libros, ignore_index=True)
        df_inf = pd.concat(df_infs, ignore_index=True).groupby('Rut Trabajador').sum().reset_index() if df_infs else pd.DataFrame()
        
        df_merge = pd.merge(df_libro, df_inf, on='Rut Trabajador', how='left') if not df_inf.empty else df_libro.copy()
        df_final = pd.DataFrame(columns=COLUMNAS_OBJETIVO)
        
        df_final['RUT EMPRESA'] = rut_empresa
        df_final['MES'] = mes_input
        df_final['AÑO'] = anio_input
        
        # Mapeo inicial (Variables del Libro)
        map_libro = {
            'Rut Trabajador': 'Rut Trabajador', 'Apellido Paterno': 'Apellido Paterno', 'Apellido Materno': 'Apellido Materno',
            'Nombres': 'Nombres', 'Cargo': 'Cargo', 'Tipo de Contrato': 'Tipo de Contrato', 'N° Dias Trabajados': 'N° Dias Trabajados',
            'N° Dias Ausentes': 'N° Dias Ausentes', 'N° Dias Licencia': 'N° Dias Licencia', 'N° Dias Accidentes de Trabajo': 'N° Dias Accidentes de Trabajo',
            'N° Cargas Familiares': 'N° Cargas Familiares', 'Sueldo Base': 'Sueldo Base (H)', 'Imponible': 'Total Imponible (H)',
            'Aporte Accidentes de Trabajo Mutual': 'Aporte Accidentes de Trabajo Mutual (P)', 'Seguro de Cesantía Empleador': 'Seguro de Cesantía Empleador (P)'
        }
        for src, dst in map_libro.items():
            if src in df_merge.columns: df_final[dst] = df_merge[src]
            
        # Mapeo de Haberes y Descuentos cruzados
        for col in COLUMNAS_OBJETIVO:
            if df_final[col].notna().any(): continue
            clean_col = col.replace(' (H)', '').replace(' (D)', '').replace(' (P)', '').strip()
            
            match = None
            for m_col in df_merge.columns:
                if m_col.lower() == clean_col.lower():
                    match = m_col
                    break
                    
            if match:
                df_final[col] = df_merge[match]
            else:
                df_final[col] = 0
                
        df_final.fillna(0, inplace=True)
        
        # Guardar en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        output.seek(0)
        
        st.success("¡Consolidado Completo Generado!")
        st.download_button("📥 Descargar Consolidado", data=output, file_name=f"Consolidado_Final_{mes_input}.xlsx")
