import re
import unicodedata
import pandas as pd
import numpy as np

from mapper import (
    construir_dataframe_destino,
    aplicar_mapeo
)


# ==========================================================
# NORMALIZACIÓN
# ==========================================================

def limpiar_rut(rut):

    if pd.isna(rut):
        return ""

    rut = str(rut)

    rut = rut.replace(".", "")
    rut = rut.replace(" ", "")
    rut = rut.upper()

    return rut


def normalizar_texto(texto):

    if texto is None:
        return ""

    texto = str(texto)

    texto = unicodedata.normalize(
        "NFKD",
        texto
    ).encode(
        "ascii",
        "ignore"
    ).decode()

    texto = texto.upper()

    texto = texto.replace("*", "")

    texto = re.sub(r"\(.*?\)", "", texto)

    texto = texto.replace("_", " ")

    texto = re.sub(r"\s+", " ", texto)

    texto = texto.strip()

    return texto


# ==========================================================
# LECTURA PLANTILLA
# ==========================================================

def leer_plantilla(file):

    prueba = pd.read_excel(
        file,
        header=None,
        nrows=5
    )

    encabezado = 0

    if prueba.shape[0] >= 3:

        fila3 = prueba.iloc[2].astype(str)

        if fila3.str.contains(
            "Rut",
            case=False
        ).any():

            encabezado = 2

    file.seek(0)

    plantilla = pd.read_excel(
        file,
        header=encabezado
    )

    plantilla.columns = [
        str(c).strip()
        for c in plantilla.columns
    ]

    return plantilla


# ==========================================================
# DETECTAR ARCHIVO
# ==========================================================

def detectar_tipo_archivo(file):

    nombre = file.name.lower()

    if "costo" in nombre:
        return "costo"

    if "haber" in nombre:
        return "haberes"

    if "descuento" in nombre:
        return "haberes"

    if "libro" in nombre:
        return "libro"

    muestra = pd.read_excel(
        file,
        nrows=20,
        header=None
    )

    file.seek(0)

    texto = " ".join(
        muestra.astype(str)
        .fillna("")
        .values.flatten()
    ).upper()

    if "SEGURO DE CESANTIA EMPLEADOR" in texto:
        return "costo"

    if "BONO" in texto and "RUT" in texto:
        return "haberes"

    return "libro"


# ==========================================================
# LECTURA MASIVA
# ==========================================================

def leer_archivos(lista_archivos):

    libros = []

    costos = []

    haberes = []

    for archivo in lista_archivos:

        tipo = detectar_tipo_archivo(archivo)

        if tipo == "libro":

            df = pd.read_excel(
                archivo
            )

            libros.append(df)

        elif tipo == "costo":

            df = pd.read_excel(
                archivo
            )

            costos.append(df)

        else:

            df = pd.read_excel(
                archivo,
                skiprows=15
            )

            haberes.append(df)

    return libros, costos, haberes

# ==========================================================
# PROCESAR LIBRO REMUNERACIONES
# ==========================================================

def procesar_libros(libros):

    if len(libros) == 0:
        return pd.DataFrame()

    base = pd.concat(
        libros,
        ignore_index=True
    )

    base.columns = [
        str(c).strip()
        for c in base.columns
    ]

    if "Rut Trabajador" not in base.columns:
        raise Exception(
            "No existe la columna 'Rut Trabajador' en Libro de Remuneraciones."
        )

    base["Rut Trabajador"] = (
        base["Rut Trabajador"]
        .astype(str)
        .apply(limpiar_rut)
    )

    base = (
        base
        .groupby("Rut Trabajador", as_index=False)
        .first()
    )

    return base


# ==========================================================
# PROCESAR COSTOS
# ==========================================================

def procesar_costos(costos):

    if len(costos) == 0:
        return pd.DataFrame()

    costo = pd.concat(
        costos,
        ignore_index=True
    )

    costo.columns = [
        str(c).strip()
        for c in costo.columns
    ]

    if "Rut Trabajador" not in costo.columns:
        raise Exception(
            "No existe la columna 'Rut Trabajador' en Costos."
        )

    costo["Rut Trabajador"] = (
        costo["Rut Trabajador"]
        .astype(str)
        .apply(limpiar_rut)
    )

    costo = (
        costo
        .groupby("Rut Trabajador", as_index=False)
        .first()
    )

    return costo


# ==========================================================
# UNIR LIBRO + COSTO
# ==========================================================

def construir_base(libros, costos):

    libro = procesar_libros(libros)

    costo = procesar_costos(costos)

    if costo.empty:
        return libro

    costo = costo.drop(
        columns=[
            c
            for c in costo.columns
            if c == "Rut Trabajador"
        ]
    )

    base = pd.concat(
        [
            libro.reset_index(drop=True),
            costo.reset_index(drop=True)
        ],
        axis=1
    )

    return base


# ==========================================================
# PROCESAR INFORME HABERES
# ==========================================================

def procesar_haberes(lista):

    if len(lista) == 0:
        return pd.DataFrame()

    todos = []

    for df in lista:

        df.columns = [
            str(c).strip()
            for c in df.columns
        ]

        categoria = df.columns[0]

        df[categoria] = (
            df[categoria]
            .fillna(method="ffill")
        )

        df = df[
            ~df[categoria]
            .astype(str)
            .str.startswith(
                "Total",
                na=False
            )
        ]

        rut_col = None

        for c in df.columns:

            if "rut" in c.lower():
                rut_col = c
                break

        monto_col = None

        for c in df.columns:

            if "monto" in c.lower():
                monto_col = c
                break

        if rut_col is None or monto_col is None:
            continue

        temp = df[
            [
                categoria,
                rut_col,
                monto_col
            ]
        ].copy()

        temp.columns = [
            "Concepto",
            "Rut Trabajador",
            "Monto"
        ]

        temp["Rut Trabajador"] = (
            temp["Rut Trabajador"]
            .astype(str)
            .apply(limpiar_rut)
        )

        temp["Monto"] = (
            pd.to_numeric(
                temp["Monto"],
                errors="coerce"
            )
            .fillna(0)
        )

        todos.append(temp)

    if len(todos) == 0:
        return pd.DataFrame()

    datos = pd.concat(
        todos,
        ignore_index=True
    )

    pivot = (
        datos
        .pivot_table(
            index="Rut Trabajador",
            columns="Concepto",
            values="Monto",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    pivot.columns.name = None

    return pivot


# ==========================================================
# MERGE FINAL
# ==========================================================

def unir_base_haberes(base, haberes):

    if haberes.empty:
        return base

    columnas_repetidas = []

    for c in haberes.columns:

        if c == "Rut Trabajador":
            continue

        if c in base.columns:
            columnas_repetidas.append(c)

    haberes = haberes.drop(
        columns=columnas_repetidas,
        errors="ignore"
    )

    resultado = base.merge(
        haberes,
        on="Rut Trabajador",
        how="left"
    )

    return resultado
