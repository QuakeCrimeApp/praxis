#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PRÁCTICA+
Sistema inteligente de asignación de prácticas
"""

import streamlit as st
import pandas as pd

# ==============================
# TODO TU PROGRAMA
# ==============================

# ...
# ...
# ...

st.download_button(
    "📥 Descargar asignaciones.xlsx",
    st.session_state.excel,
    "asignaciones.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary"
)

# ==============================
# CRÉDITOS
# ==============================

st.markdown("""
<div style="
    margin-top: 35px;
    padding: 15px 20px;
    border-top: 1px solid #cccccc;
    text-align: center;
    font-size: 14px;
    line-height: 1.6;
">
    <strong>Desarrollado por:</strong><br>
    <strong>Diego J. Maldonado Guzmán</strong><br>
    Profesor del Área de Derecho Penal de la Universidad de Málaga<br>
    Investigador del Instituto Andaluz Interuniversitario de Criminología – Sección Málaga
</div>
""", unsafe_allow_html=True)

import os
import re
from pathlib import Path
from io import BytesIO

import pandas as pd
import streamlit as st


# =========================================================
# CONFIG
# =========================================================

APP_NAME = "PRÁCTICA+"

COL = {

    # alumnos
    "id": "idalumno",
    "apellido": "apellido",
    "nombre": "nombre",
    "perfil": "perfil",
    "nota": "notamedia",

    # plazas
    "pid": "idplaza",
    "pname": "nombreplaza",

    "cm": "capacidadmanana",
    "ct": "capacidadtarde",

    "profiles": "perfilesadmitidos"
}

# =========================================================
# NORMALIZAR COLUMNAS
# =========================================================

def clean_col(x):

    return (
        str(x)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )

# =========================================================
# LEER EXCEL
# =========================================================

def leer_excel(source):
    if source is None:
        raise Exception("No se seleccionó archivo.")
    try:
        if isinstance(source, (str, Path)):
            path = Path(source)
            if not path.exists():
                raise Exception(f"No existe:\n{path}")
            df = pd.read_excel(path, keep_default_na=False, engine="openpyxl")
        else:
            if hasattr(source, "seek"):
                source.seek(0)
            df = pd.read_excel(source, keep_default_na=False, engine="openpyxl")
    except Exception as e:
        raise Exception(f"No se pudo leer Excel:\n{e}")
    df.columns = [clean_col(c) for c in df.columns]
    df = df.where(pd.notnull(df), "")
    return df.astype(str)


# =========================================================
# UTILIDADES
# =========================================================

def to_float(x):

    try:

        return float(
            str(x)
            .replace(",", ".")
        )

    except:

        return 0.0

def normalizar_perfil(p):

    p = str(p).lower().strip()

    mapa = {

        "educadora social":
            "educación social",

        "educador social":
            "educación social",

        "trabajadora social":
            "trabajo social",

        "trabajador social":
            "trabajo social",

        "criminóloga":
            "criminología",

        "criminologo":
            "criminología"
    }

    return mapa.get(p, p)

def parse_profiles(x):

    x = str(x).lower().strip()

    if x in ("", "todos", "all"):
        return None

    return {

        normalizar_perfil(p)

        for p in x.split("|")
    }

def normalize_shift(x):

    x = str(x).lower().strip()

    if x in ("", "any", "cualquiera"):
        return "any"

    if x in (
        "m",
        "mañana",
        "manana",
        "morning"
    ):
        return "morning"

    return "afternoon"

# =========================================================
# VALIDAR COLUMNAS
# =========================================================

def validar_columnas(df, cols, nombre):

    faltan = []

    for c in cols:

        if c not in df.columns:
            faltan.append(c)

    if faltan:

        raise Exception(

            f"Faltan columnas en {nombre}:\n\n"

            + "\n".join(faltan)
        )

# =========================================================
# PREPARAR PLAZAS
# =========================================================

def preparar_plazas(df):

    validar_columnas(

        df,

        [
            COL["pid"],
            COL["pname"],
            COL["cm"],
            COL["ct"]
        ],

        "PLAZAS"
    )

    plazas = {}
    alias = {}
    ocupacion = {}

    for _, r in df.iterrows():

        pid = str(
            r[COL["pid"]]
        ).strip().lower()

        nombre = str(
            r[COL["pname"]]
        ).strip().lower()

        if not pid:
            continue

        try:
            cm = int(
                float(
                    str(r[COL["cm"]])
                    .replace(",", ".")
                )
            )
        except:
            cm = 0

        try:
            ct = int(
                float(
                    str(r[COL["ct"]])
                    .replace(",", ".")
                )
            )
        except:
            ct = 0

        profiles = None

        if COL["profiles"] in df.columns:

            profiles = parse_profiles(
                r[COL["profiles"]]
            )

        plazas[pid] = {

            "id": pid,

            "nombre":
                str(r[COL["pname"]]),

            "cm_total":
                cm,

            "ct_total":
                ct,

            "profiles":
                profiles
        }

        ocupacion[pid] = {

            "cm": 0,
            "ct": 0
        }

        alias[nombre] = pid

    return plazas, alias, ocupacion

# =========================================================
# DETECTAR PREFERENCIAS
# =========================================================

def detectar_preferencias(df):

    prefs = []

    for c in df.columns:

        c = clean_col(c)

        m = re.match(
            r"preferencia([0-9]+)$",
            c
        )

        if m:

            prefs.append(
                int(m.group(1))
            )

    return sorted(set(prefs))

# =========================================================
# ASIGNAR
# =========================================================

def asignar(
    df,
    plazas,
    alias,
    ocupacion,
    prefer_morning=True
):

    validar_columnas(

        df,

        [
            COL["id"],
            COL["apellido"],
            COL["nombre"],
            COL["perfil"],
            COL["nota"]
        ],

        "ALUMNOS"
    )

    preferencias = detectar_preferencias(df)

    if not preferencias:

        raise Exception(
            "No existen columnas preferencia1, preferencia2..."
        )

    # =====================================================
    # ORDEN REAL POR NOTA
    # =====================================================

    df["_nota"] = df[
        COL["nota"]
    ].apply(to_float)

    df = df.sort_values(

        by=[
            "_nota",
            COL["apellido"],
            COL["nombre"]
        ],

        ascending=[False, True, True]
    )

    out = []

    # =====================================================
    # RECORRER ALUMNOS
    # =====================================================

    for _, s in df.iterrows():

        perfil = normalizar_perfil(
            s[COL["perfil"]]
        )

        asignado = False

        motivo = ""

        # =================================================
        # RECORRER PREFERENCIAS EN ORDEN
        # =================================================

        for i in preferencias:

            pref_col = f"preferencia{i}"
            turno_col = f"preferencia{i}turno"

            pref = str(
                s.get(pref_col, "")
            ).lower().strip()

            turno = normalize_shift(
                s.get(turno_col, "any")
            )

            if not pref:
                continue

            # nombre -> id

            if pref in alias:
                pref = alias[pref]

            if pref not in plazas:

                motivo = f"pref{i} no existe"

                continue

            plaza = plazas[pref]

            # =================================================
            # VALIDAR PERFIL
            # =================================================

            allowed = plaza["profiles"]

            if (
                allowed is not None
                and perfil not in allowed
            ):

                motivo = f"pref{i} perfil incompatible"

                continue

            # =================================================
            # TURNOS
            # =================================================

            if turno == "any":

                orden = (

                    ["cm", "ct"]

                    if prefer_morning

                    else ["ct", "cm"]
                )

            elif turno == "morning":

                orden = ["cm"]

            else:

                orden = ["ct"]

            # =================================================
            # INTENTAR ASIGNAR
            # =================================================

            asignado_turno = None

            for o in orden:

                capacidad_total = (

                    plaza["cm_total"]

                    if o == "cm"

                    else plaza["ct_total"]
                )

                ocupadas = ocupacion[pref][o]

                libres = (
                    capacidad_total
                    - ocupadas
                )

                if libres > 0:

                    ocupacion[pref][o] += 1

                    asignado = True

                    asignado_turno = (

                        "m"

                        if o == "cm"

                        else "t"
                    )

                    break

            if asignado:

                out.append({

                    "ID":
                        s[COL["id"]],

                    "Apellido":
                        s[COL["apellido"]],

                    "Nombre":
                        s[COL["nombre"]],

                    "Perfil":
                        perfil,

                    "Nota":
                        s[COL["nota"]],

                    "Empresa":
                        plaza["nombre"],

                    "Turno":
                        asignado_turno,

                    "Preferencia":
                        i,

                    "Estado":
                        "asignado"
                })

                break

            else:

                motivo = f"pref{i} sin plazas"

        # =====================================================
        # SIN ASIGNAR
        # =====================================================

        if not asignado:

            out.append({

                "ID":
                    s[COL["id"]],

                "Apellido":
                    s[COL["apellido"]],

                "Nombre":
                    s[COL["nombre"]],

                "Perfil":
                    perfil,

                "Nota":
                    s[COL["nota"]],

                "Empresa":
                    "",

                "Turno":
                    "",

                "Preferencia":
                    "",

                "Estado":
                    motivo or "sin plaza"
            })

    return pd.DataFrame(out)


# =========================================================
# INTERFAZ STREAMLIT
# =========================================================
st.set_page_config(page_title="PRAXIS", page_icon="🎓", layout="wide")

st.title("🎓 PRAXIS")
st.subheader("Sistema inteligente de asignación de prácticas")
st.caption("Desarrollado por Diego J. Maldonado Guzmán")

with st.sidebar:
    st.header("Configuración")
    alumnos = st.file_uploader("Archivo de alumnos", type=["xlsx"])
    plazas_file = st.file_uploader("Archivo de plazas", type=["xlsx"])
    prefer_morning = st.checkbox("Preferir turno de mañana", value=True)
    ejecutar = st.button(
        "▶ Ejecutar asignación",
        type="primary",
        use_container_width=True,
        disabled=not (alumnos and plazas_file),
    )

if "resultado" not in st.session_state:
    st.session_state.resultado = None
if "excel" not in st.session_state:
    st.session_state.excel = None

if ejecutar:
    try:
        with st.spinner("Calculando asignaciones..."):
            df_s = leer_excel(alumnos)
            df_p = leer_excel(plazas_file)
            plazas, alias, ocupacion = preparar_plazas(df_p)
            resultado = asignar(df_s.copy(), plazas, alias, ocupacion, prefer_morning)

            output = BytesIO()
            from openpyxl.styles import PatternFill, Font
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                resultado.to_excel(writer, sheet_name="Asignaciones", index=False)
                ws = writer.book["Asignaciones"]
                fill_ok = PatternFill("solid", fgColor="C6EFCE")
                fill_warn = PatternFill("solid", fgColor="FFF2CC")
                fill_bad = PatternFill("solid", fgColor="F8D7DA")
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                for row in ws.iter_rows(min_row=2):
                    estado, pref = row[-1].value, row[-2].value
                    fill = fill_ok if estado == "asignado" and pref == 1 else (
                        fill_warn if estado == "asignado" else fill_bad
                    )
                    for cell in row:
                        cell.fill = fill
                for cells in ws.columns:
                    ancho = max(len(str(c.value or "")) for c in cells) + 2
                    ws.column_dimensions[cells[0].column_letter].width = min(max(ancho, 10), 45)

            st.session_state.resultado = resultado
            st.session_state.excel = output.getvalue()
        st.success("Asignación completada.")
    except Exception as e:
        st.session_state.resultado = None
        st.session_state.excel = None
        st.error(f"Error: {e}")

resultado = st.session_state.resultado

if resultado is not None:
    total = len(resultado)

    asignados = int(
        (resultado["Estado"] == "asignado").sum()
    )

    primera = int(
        (
            (resultado["Estado"] == "asignado")
            &
            (
                pd.to_numeric(
                    resultado["Preferencia"],
                    errors="coerce"
                ) == 1
            )
        ).sum()
    )

    # ==========================================
    # MÉTRICAS
    # ==========================================

    a, b, c, d = st.columns(4)

    a.metric("Alumnos", total)
    b.metric("Asignados", asignados)
    c.metric("Sin plaza", total - asignados)
    d.metric("1.ª preferencia", primera)

    # ==========================================
    # COLORES DE LA TABLA
    # ==========================================

    def colorear(row):

        # Primera preferencia → VERDE
        if (
            row["Estado"] == "asignado"
            and str(row["Preferencia"]) in ("1", "1.0")
        ):
            estilo = (
                "background-color: #d4f4dd; "
                "color: #000000;"
            )

        # Otras preferencias → AMARILLO
        elif row["Estado"] == "asignado":

            estilo = (
                "background-color: #fff3cd; "
                "color: #000000;"
            )

        # Sin plaza → ROJO
        else:

            estilo = (
                "background-color: #f8d7da; "
                "color: #000000;"
            )

        return [estilo] * len(row)

    # ==========================================
    # RESULTADOS
    # ==========================================

    st.subheader("Resultados")

    tabla_estilizada = resultado.style.apply(
        colorear,
        axis=1
    )

    st.dataframe(
        tabla_estilizada,
        use_container_width=True,
        hide_index=True,
        height=520
    )

    # ==========================================
    # DESCARGAR EXCEL
    # ==========================================

    st.download_button(
        "📥 Descargar asignaciones.xlsx",
        st.session_state.excel,
        "asignaciones.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
