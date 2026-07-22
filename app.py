import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
import numpy as np

# -------------------------------------------------------------------
# 1. FUNCIONES AUXILIARES
# -------------------------------------------------------------------

def normalize_text(text: str) -> str:
    """Elimina acentos, convierte a mayúsculas y limpia espacios."""
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("utf-8")
    text = re.sub(r"[()]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text.upper()


def clean_rut(rut: str) -> str:
    """Limpia un RUT eliminando puntos, guiones y espacios."""
    if not isinstance(rut, str):
        return ""
    rut = re.sub(r"[.\s-]", "", rut)
    return rut.upper()


def infer_file_type(filename: str) -> str:
    """Infiera el tipo de archivo según el nombre."""
    name = filename.lower()
    if "libro de remuneraciones" in name or "libro" in name:
        return "libro"
    if "costo por trabajador" in name or "costo" in name:
        return "costo"
    if "informe haberes" in name or "haberes y descuentos" in name:
        return "informe"
    return "unknown"


def normalize_column_name(col: str) -> str:
    """Normaliza un nombre de columna para comparación (sin acentos, sin paréntesis)."""
    if not isinstance(col, str):
        return ""
    # Quitar paréntesis y su contenido
    col = re.sub(r"\([^)]*\)", "", col)
    col = normalize_text(col)
    return col


def find_target_column(
    source_col: str,
    target_columns: list,
    explicit_mapping: dict,
    verbose: bool = False,
) -> str:
    """
    Encuentra la columna destino para una columna fuente.
    Primero usa el mapeo explícito, luego intenta coincidencia fonética.
    """
    source_norm = normalize_column_name(source_col)

    # 1. Mapeo explícito
    if source_col in explicit_mapping:
        candidates = explicit_mapping[source_col]
        for cand in candidates:
            if cand in target_columns:
                if verbose:
                    st.write(f"✅ Mapeo explícito: {source_col} -> {cand}")
                return cand
        # Si el mapeo explícito falla, intenta con la primera opción normalizada
        for cand in candidates:
            cand_norm = normalize_column_name(cand)
            for tcol in target_columns:
                if normalize_column_name(tcol) == cand_norm:
                    if verbose:
                        st.write(f"✅ Mapeo explícito (normalizado): {source_col} -> {tcol}")
                    return tcol

    # 2. Coincidencia fonética (normalizada)
    for tcol in target_columns:
        tcol_norm = normalize_column_name(tcol)
        if source_norm == tcol_norm:
            if verbose:
                st.write(f"✅ Coincidencia fonética: {source_col} -> {tcol}")
            return tcol

    # 3. Coincidencia parcial (contiene)
    for tcol in target_columns:
        tcol_norm = normalize_column_name(tcol)
        if source_norm in tcol_norm or tcol_norm in source_norm:
            if verbose:
                st.write(f"⚠️ Coincidencia parcial: {source_col} -> {tcol}")
            return tcol

    if verbose:
        st.write(f"❌ No se encontró destino para: {source_col}")
    return None


# -------------------------------------------------------------------
# 2. PROCESAMIENTO DEL INFORME DE HABERES Y DESCUENTOS
# -------------------------------------------------------------------

def process_informe_file(file) -> pd.DataFrame:
    """
    Procesa un archivo Informe Haberes y Descuentos.
    Los datos reales empiezan en fila 16 (skiprows=15).
    Estructura: col0 = categoría, col2 = RUT, col3 = Nombre, col9 = Monto.
    """
    # Leer el archivo para obtener los datos
    df_raw = pd.read_excel(file, skiprows=15, header=None)

    # Columnas según el mapeo: 0=categoría, 2=RUT, 3=Nombre, 9=Monto
    # También puede haber una columna 7 = N° de Cuota (opcional)
    if len(df_raw.columns) < 10:
        st.warning(f"El archivo {file.name} tiene menos columnas de lo esperado.")
        return pd.DataFrame()

    # Renombrar columnas para facilitar el trabajo
    df_raw.columns = ["categoria", "col1", "rut", "nombre", "col4", "col5", "col6", "cuota", "col8", "monto", "col10"] + [
        f"col{i}" for i in range(11, len(df_raw.columns))
    ]

    # Limpiar RUT
    df_raw["rut"] = df_raw["rut"].astype(str).apply(clean_rut)

    # Procesar filas para construir el pivot
    records = []
    current_category = None

    for idx, row in df_raw.iterrows():
        cat = row.get("categoria")
        if pd.notna(cat) and isinstance(cat, str):
            cat_clean = cat.strip()
            # Si la categoría empieza con "Total", ignorar
            if cat_clean.lower().startswith("total"):
                continue
            # Si la categoría es una categoría válida (no vacía), actualizar
            if cat_clean and not cat_clean.startswith("Total"):
                current_category = cat_clean
                continue

        # Si hay un RUT en esta fila, es un registro de datos
        rut = row.get("rut")
        if pd.notna(rut) and rut and str(rut).strip():
            monto = row.get("monto")
            if pd.isna(monto) or monto == "":
                monto = 0
            try:
                monto = float(monto)
            except (ValueError, TypeError):
                monto = 0

            if current_category and monto != 0:
                records.append({
                    "rut": str(rut).strip(),
                    "categoria": current_category,
                    "monto": monto,
                })

    if not records:
        return pd.DataFrame()

    # Crear DataFrame y pivotear
    df_records = pd.DataFrame(records)

    # Si hay duplicados (mismo RUT + misma categoría), sumar
    df_pivot = df_records.pivot_table(
        index="rut",
        columns="categoria",
        values="monto",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    # Asegurar que el RUT sea string
    df_pivot["rut"] = df_pivot["rut"].astype(str)

    return df_pivot


# -------------------------------------------------------------------
# 3. PROCESAMIENTO DE ARCHIVOS LIBRO Y COSTO
# -------------------------------------------------------------------

def process_libro_or_costo(file, file_type: str) -> pd.DataFrame:
    """
    Procesa un archivo Libro de Remuneraciones o Costo por Trabajador.
    Ambos tienen la misma estructura (primera fila = header).
    """
    df = pd.read_excel(file)

    # Limpiar RUT
    if "Rut Trabajador" in df.columns:
        df["Rut Trabajador"] = df["Rut Trabajador"].astype(str).apply(clean_rut)
    elif "Rut" in df.columns:
        df["Rut"] = df["Rut"].astype(str).apply(clean_rut)
        df.rename(columns={"Rut": "Rut Trabajador"}, inplace=True)

    # Si hay una columna "Rut *" renombrar a "Rut Trabajador"
    if "Rut *" in df.columns:
        df["Rut *"] = df["Rut *"].astype(str).apply(clean_rut)
        if "Rut Trabajador" not in df.columns:
            df.rename(columns={"Rut *": "Rut Trabajador"}, inplace=True)

    return df


# -------------------------------------------------------------------
# 4. FUNCIÓN PRINCIPAL DE PROCESAMIENTO
# -------------------------------------------------------------------

def process_all_files(
    template_file,
    data_files,
    mapping_dict,
    verbose: bool = True,
) -> pd.DataFrame:
    """
    Procesa todos los archivos y genera el consolidado final.
    """
    # ----- 4a. Leer plantilla destino -----
    # La plantilla tiene 2 filas decorativas, header está en fila 3 (header=2)
    try:
        template_df = pd.read_excel(template_file, header=2)
    except Exception as e:
        # Si falla, intentar sin header
        st.warning(f"Error al leer la plantilla con header=2, intentando sin header: {e}")
        template_df = pd.read_excel(template_file, header=0)

    target_columns = template_df.columns.tolist()
    if verbose:
        st.write(f"📋 Columnas destino ({len(target_columns)} columnas)")

    # ----- 4b. Clasificar archivos -----
    libro_files = []
    costo_files = []
    informe_files = []

    for f in data_files:
        ftype = infer_file_type(f.name)
        if ftype == "libro":
            libro_files.append(f)
        elif ftype == "costo":
            costo_files.append(f)
        elif ftype == "informe":
            informe_files.append(f)
        else:
            st.warning(f"⚠️ Tipo desconocido: {f.name}")

    if verbose:
        st.write(f"📂 Libros: {len(libro_files)}, Costos: {len(costo_files)}, Informes: {len(informe_files)}")

    # ----- 4c. Procesar Libros -----
    libro_dfs = []
    for f in libro_files:
        try:
            df = process_libro_or_costo(f, "libro")
            libro_dfs.append(df)
        except Exception as e:
            st.error(f"❌ Error procesando libro {f.name}: {e}")

    if not libro_dfs:
        st.error("No se pudo procesar ningún archivo de Libro de Remuneraciones.")
        return pd.DataFrame()

    libro_df = pd.concat(libro_dfs, ignore_index=True)

    # Agrupar por RUT (primero)
    if "Rut Trabajador" in libro_df.columns:
        libro_df = libro_df.groupby("Rut Trabajador").first().reset_index()

    if verbose:
        st.write(f"📊 Libro: {len(libro_df)} registros")

    # ----- 4d. Procesar Costos -----
    costo_dfs = []
    for f in costo_files:
        try:
            df = process_libro_or_costo(f, "costo")
            costo_dfs.append(df)
        except Exception as e:
            st.error(f"❌ Error procesando costo {f.name}: {e}")

    if costo_dfs:
        costo_df = pd.concat(costo_dfs, ignore_index=True)
        if "Rut Trabajador" in costo_df.columns:
            costo_df = costo_df.groupby("Rut Trabajador").first().reset_index()
        if verbose:
            st.write(f"📊 Costo: {len(costo_df)} registros")
    else:
        costo_df = pd.DataFrame()
        if verbose:
            st.write("📊 Costo: no se encontraron archivos")

    # ----- 4e. Fusionar Libro + Costo -----
    base_df = libro_df.copy()
    if not costo_df.empty and "Rut Trabajador" in costo_df.columns:
        # Eliminar columnas duplicadas de costo que ya están en libro
        cols_to_drop = []
        for col in costo_df.columns:
            if col != "Rut Trabajador" and col in base_df.columns:
                cols_to_drop.append(col)
        if cols_to_drop:
            costo_df = costo_df.drop(columns=cols_to_drop, errors="ignore")

        base_df = base_df.merge(costo_df, on="Rut Trabajador", how="left", suffixes=("", "_costo"))

    if verbose:
        st.write(f"📊 Base fusionada: {len(base_df)} registros, {len(base_df.columns)} columnas")

    # ----- 4f. Procesar Informes Haberes -----
    informe_dfs = []
    for f in informe_files:
        try:
            df_inf = process_informe_file(f)
            if not df_inf.empty:
                informe_dfs.append(df_inf)
        except Exception as e:
            st.error(f"❌ Error procesando informe {f.name}: {e}")

    if informe_dfs:
        # Concatenar todos los informes y agrupar por RUT sumando
        informe_df = pd.concat(informe_dfs, ignore_index=True)
        # Asegurar que rut sea string
        informe_df["rut"] = informe_df["rut"].astype(str)
        # Agrupar por RUT sumando todas las categorías
        informe_df = informe_df.groupby("rut").sum(numeric_only=True).reset_index()
        informe_df.rename(columns={"rut": "Rut Trabajador"}, inplace=True)

        if verbose:
            st.write(f"📊 Informe: {len(informe_df)} registros, {len(informe_df.columns)} columnas")

        # ----- 4g. Eliminar columnas del informe que ya existen en la base -----
        base_cols_norm = {normalize_column_name(c): c for c in base_df.columns}
        cols_to_drop_from_informe = []

        for col in informe_df.columns:
            if col == "Rut Trabajador":
                continue
            col_norm = normalize_column_name(col)
            if col_norm in base_cols_norm:
                cols_to_drop_from_informe.append(col)
                if verbose:
                    st.write(f"🗑️ Eliminando de informe (ya en base): {col} -> {base_cols_norm[col_norm]}")

        if cols_to_drop_from_informe:
            informe_df = informe_df.drop(columns=cols_to_drop_from_informe, errors="ignore")
            if verbose:
                st.write(f"🗑️ Eliminadas {len(cols_to_drop_from_informe)} columnas del informe")

        # ----- 4h. Fusionar base con informe -----
        final_df = base_df.merge(informe_df, on="Rut Trabajador", how="left", suffixes=("", "_informe"))

        # Rellenar NaN con 0 para columnas numéricas del informe
        for col in informe_df.columns:
            if col != "Rut Trabajador" and col in final_df.columns:
                final_df[col] = pd.to_numeric(final_df[col], errors="coerce").fillna(0)
    else:
        final_df = base_df
        if verbose:
            st.write("📊 Informe: no se encontraron archivos")

    if verbose:
        st.write(f"📊 Final: {len(final_df)} registros, {len(final_df.columns)} columnas")

    # ----- 4i. Mapear a columnas destino -----
    result_df = pd.DataFrame(columns=target_columns)

    # Rellenar con valores predeterminados para evitar errores
    for col in target_columns:
        result_df[col] = ""

    # Mapear columnas
    mapped_count = 0
    for source_col in final_df.columns:
        if source_col in ["Rut Trabajador", "Rut *", "Rut"]:
            # Tratamiento especial para RUT
            target_col = "Rut *"
            if target_col in target_columns:
                result_df[target_col] = final_df[source_col].astype(str).apply(clean_rut)
                mapped_count += 1
                if verbose:
                    st.write(f"✅ {source_col} -> {target_col}")
            continue

        target_col = find_target_column(source_col, target_columns, mapping_dict, verbose=verbose)
        if target_col:
            # Si es numérico, convertir
            try:
                # Intentar convertir a numérico
                values = pd.to_numeric(final_df[source_col], errors="coerce")
                result_df[target_col] = values.fillna(0)
            except:
                # Si falla, mantener como texto
                result_df[target_col] = final_df[source_col].fillna("")
            mapped_count += 1

    if verbose:
        st.write(f"✅ Mapeadas {mapped_count} columnas de {len(final_df.columns)}")

    # ----- 4j. Aplicar reglas de limpieza para Talana -----
    # Renombrar columnas
    if "ASIGNACIÓN VIÁTICO MEU" in result_df.columns:
        result_df.rename(columns={"ASIGNACIÓN VIÁTICO MEU": "ASIGNACION VIATICO MEU"}, inplace=True)

    if "BONO ADMINISTRACION" in result_df.columns:
        result_df.rename(columns={"BONO ADMINISTRACION": "BONO ADMINISTRACIÓN"}, inplace=True)

    # Eliminar columnas no válidas
    cols_to_drop_talana = ["BONO PMI", "BONO RECONOCIMIENTO", "DIAS PARO"]
    for col in cols_to_drop_talana:
        if col in result_df.columns:
            result_df.drop(columns=[col], inplace=True)
            if verbose:
                st.write(f"🗑️ Eliminada columna no válida: {col}")

    # ----- 4k. Manejar tipos de datos -----
    # Texto: fillna("")
    text_columns = ["Rut *", "Rut Razón Social", "Año_Mes (aaaamm)", "Rut Trabajador", "Nombres", "Cargo", "Tipo de Contrato"]
    for col in text_columns:
        if col in result_df.columns:
            result_df[col] = result_df[col].fillna("").astype(str)

    # Numérico: pd.to_numeric con fillna(0)
    for col in result_df.columns:
        if col not in text_columns:
            # Intentar convertir a numérico
            try:
                result_df[col] = pd.to_numeric(result_df[col], errors="coerce").fillna(0)
            except:
                # Si falla, dejar como texto
                pass

    # Asegurar que Año_Mes tenga el formato YYYYMM
    if "Año_Mes (aaaamm)" in result_df.columns:
        result_df["Año_Mes (aaaamm)"] = result_df["Año_Mes (aaaamm)"].astype(str).apply(
            lambda x: re.sub(r"[^0-9]", "", x)[:6] if x else ""
        )

    # Eliminar filas con RUT vacío
    if "Rut *" in result_df.columns:
        result_df = result_df[result_df["Rut *"].str.strip() != ""]

    if verbose:
        st.write(f"📊 Resultado final: {len(result_df)} registros, {len(result_df.columns)} columnas")

    return result_df


# -------------------------------------------------------------------
# 5. DEFINICIÓN DEL MAPEO EXPLÍCITO
# -------------------------------------------------------------------

MAPPING_DICT = {
    # RUT
    "Rut Trabajador": ["Rut *", "Rut Trabajador"],
    "Rut": ["Rut *", "Rut Trabajador"],

    # Sueldos y haberes
    "Sueldo Base": ["Sueldo Base *", "Sueldo Base"],
    "Imponible": ["Remuneración Imponible *", "Imponible"],
    "Base Tributable": ["Total Tributable (afecta a impuesto) *", "Base Tributable"],
    "Líquido": ["Sueldo Liquido *", "Alcance Liquido", "Líquido"],
    "Total Haberes": ["Remuneración Total o Sueldo Bruto *", "Total Haberes"],

    # Impuestos y descuentos
    "Impuesto Unico": ["Impuestos *", "Impuesto Unico"],
    "Salud": ["Total Isapre *", "Salud"],
    "Prevision": ["AFP *", "Prevision"],
    "Seguro de Cesantía": ["Seguro Cesantía Trabajador *", "Seguro de Cesantía"],

    # Aportes patronales
    "Aporte Accidentes de Trabajo Mutual": ["Seguro Accidente de Trabajo *", "Aporte Accidentes de Trabajo Mutual"],
    "Seguro de Cesantía Empleador": ["Seguro Cesantía Empleador *", "Seguro de Cesantía Empleador"],
    "Seguro de Invalidez y Sobrevivencia Empleador": [
        "Seguro Invalidez y Supervivencia (SIS) *",
        "Seguro de Invalidez y Sobrevivencia Empleador",
    ],

    # Días
    "N° Dias Trabajados": ["Días Trabajados del Mes *", "Días Trabajados"],
    "Dias Ausentes": ["Días de Ausentismo *", "Días de Ausentismo"],
    "Dias Licencia": ["Días de licencia médica *", "Días de licencia médica"],

    # Asignaciones y bonos
    "Asignacion Familiar": ["Asignación familiar y Maternal", "Asignacion Familiar"],
    "Asignación Familiar": ["Asignación familiar y Maternal", "Asignacion Familiar"],
    "Cargas Familiares Normales": ["Asignación familiar y Maternal", "Cargas Familiares Normales"],
    "Colación": ["Colación", "Colacion"],
    "Movilización": ["Movilización", "Movilizacion"],
    "Gastos Reembolsables": ["Gastos Reembolsables"],
    "Viatico Alojamiento": ["Viatico Alojamiento"],
    "Visita Adicional": ["Visita Adicional"],
    "Sobretiempo": ["Sobretiempo"],
    "Sobretiempo Festivo": ["Sobretiempo Festivo"],
    "Compensacion dia feriado": ["Compensacion dia feriado"],
    "Reconocimiento Mentoria": ["Reconocimiento Mentoria"],
    "Bono de Faena": ["Bono de Faena"],
    "Bono por Asistencia": ["Bono por Asistencia"],
    "Incentivo Excepcional": ["Incentivo de Desempeño", "Incentivo Excepcional"],
    "Asignación Capacitación": ["Asignación Capacitación", "Asignacion Capacitacion"],
    "Asignación de Cargo": ["Asignación de Cargo", "Asignacion de Cargo"],
    "Asignación de Telefono": ["Asignación de Telefono", "Asignacion de Telefono"],
    "Asignacion de Nacimiento": ["Asignación de Nacimiento", "Asignacion de Nacimiento"],
    "Bono por Matrimonio": ["Bono Matrimonio", "Bono por Matrimonio"],
    "Anticipo Bono": ["Anticipo Bono"],
    "Otros Haberes no Imponibles": ["Otros haberes exentos", "Otros Haberes no Imponibles"],

    # Descuentos
    "Anticipos": ["Descuento Anticipo", "Anticipos"],
    "Descuentos CCAF": ["Descuento Crédito Personal CCAF", "Descuentos CCAF"],
    "Cuenta de Ahorro AFP": ["Cuenta de Ahorro AFP"],
    "Retencion Pension de Alimentos": ["Retencion Pension de Alimentos"],
    "Seguro Vida Camara": ["Seguro Vida Camara"],
    "Descuento Seguro de Vida CCAF": ["Descuento Seguro de Vida CCAF"],
    "Prestamo Empresa": ["Prestamo Empresa"],
    "Gastos Particulares": ["Gastos Particulares"],
    "Descuento por Leasing o Ahorro CCAF": ["Descuento por Leasing o Ahorro CCAF"],
    "Ajuste Dias Descontados": ["Días de Ausentismo *", "Ajuste Dias Descontados"],

    # Otros
    "Gratificación Legal": ["Gratificación Legal", "Gratificacion"],
    "Gratificacion": ["Gratificación Legal", "Gratificacion"],
    "Aporte Empleador Trabajo Pesado": ["Trabajo Pesado Empleador", "Aporte Empleador Trabajo Pesado"],
    "Aporte Empleador Servicio Médico": ["Aporte Empleador Servicio Médico CCHC"],
    "Adicional de Capitalización Individual AFP": ["Capitalización Individual AFP", "Adicional de Capitalización Individual AFP"],
    "Expectativa de Vida Seguro Social": ["Expectativa de Vida", "Expectativa de Vida Seguro Social"],
}


# -------------------------------------------------------------------
# 6. APLICACIÓN STREAMLIT
# -------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Consolidación Remuneraciones", layout="wide")
    st.title("🧾 Consolidación de Remuneraciones para Talana")
    st.markdown("Sube la plantilla destino y los archivos del mes para generar el consolidado.")

    # --- Sidebar ---
    with st.sidebar:
        st.header("📤 Carga de archivos")

        # Plantilla destino
        template_file = st.file_uploader(
            "📄 Plantilla Destino (PLANILLA SUBIREEEEE.xlsx)",
            type=["xlsx"],
            key="template",
        )

        st.markdown("---")

        # Archivos de datos
        data_files = st.file_uploader(
            "📂 Archivos de datos (selecciona múltiples)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="data",
        )

        st.markdown("---")

        # Mostrar archivos cargados
        if data_files:
            st.write(f"**{len(data_files)} archivos cargados:**")
            for f in data_files:
                ftype = infer_file_type(f.name)
                icon = "📘" if ftype == "libro" else "📗" if ftype == "costo" else "📕" if ftype == "informe" else "📄"
                st.write(f"{icon} {f.name}")

        st.markdown("---")
        st.caption("""
        **Tipos de archivo detectados:**
        - 📘 Libro de Remuneraciones
        - 📗 Costo por Trabajador
        - 📕 Informe Haberes y Descuentos
        """)

    # --- Cuerpo principal ---
    if not template_file:
        st.info("👈 Sube la plantilla destino para comenzar.")
        return

    if not data_files:
        st.info("👈 Sube los archivos de datos del mes.")
        return

    # Botón de procesamiento
    if st.button("🚀 Procesar y Consolidar", type="primary"):
        with st.spinner("Procesando archivos..."):
            try:
                # Procesar
                result_df = process_all_files(
                    template_file=template_file,
                    data_files=data_files,
                    mapping_dict=MAPPING_DICT,
                    verbose=False,
                )

                if result_df.empty:
                    st.error("❌ No se generaron datos. Verifica los archivos.")
                    return

                # Mostrar vista previa
                st.success(f"✅ Consolidado generado con {len(result_df)} registros y {len(result_df.columns)} columnas.")

                col1, col2 = st.columns([2, 1])

                with col1:
                    st.subheader("📊 Vista previa del consolidado")
                    st.dataframe(result_df.head(10), use_container_width=True)

                with col2:
                    st.subheader("📋 Resumen")
                    st.write(f"**Registros:** {len(result_df)}")
                    st.write(f"**Columnas:** {len(result_df.columns)}")
                    # Mostrar algunas columnas clave
                    if "Rut *" in result_df.columns:
                        st.write(f"**RUTs únicos:** {result_df['Rut *'].nunique()}")

                # Descarga
                output = BytesIO()
                result_df.to_excel(output, index=False)
                output.seek(0)

                st.download_button(
                    label="📥 Descargar Consolidado Final",
                    data=output,
                    file_name="consolidado_remuneraciones.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

                # También permitir descargar con nombre del mes
                if "Año_Mes (aaaamm)" in result_df.columns:
                    mes_val = result_df["Año_Mes (aaaamm)"].iloc[0] if not result_df["Año_Mes (aaaamm)"].isna().all() else "YYYYMM"
                    if mes_val and mes_val != "":
                        st.download_button(
                            label=f"📥 Descargar Consolidado {mes_val}",
                            data=output,
                            file_name=f"consolidado_{mes_val}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

            except Exception as e:
                st.error(f"❌ Error en el procesamiento: {str(e)}")
                import traceback
                st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
