#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
═══════════════════════════════════════════════
 PRÁCTICA+
 Sistema inteligente de asignación de prácticas
═══════════════════════════════════════════════

✔ SOLO Excel (.xlsx)
✔ Prioridad REAL por nota
✔ Preferencias ilimitadas
✔ Compatible con nombres o IDs
✔ Turnos mañana/tarde
✔ Perfiles compatibles
✔ GUI visual
✔ Exportación Excel
✔ Corrección definitiva preferencias

Desarrollado por:
Diego J. Maldonado Guzmán
"""

import os
import re
import traceback
from pathlib import Path

import pandas as pd

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

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

def leer_excel(path):

    if path is None:
        raise Exception("No se seleccionó archivo.")

    # Streamlit entrega un UploadedFile; también mantenemos compatibilidad con rutas.
    if hasattr(path, "name") and hasattr(path, "read"):
        ext = Path(path.name).suffix.lower()
        source = path
    else:
        if not os.path.exists(path):
            raise Exception(f"No existe:\n{path}")
        ext = Path(path).suffix.lower()
        source = path

    if ext != ".xlsx":
        raise Exception("Solo se permiten archivos Excel .xlsx.")

    try:

        df = pd.read_excel(
            source,
            keep_default_na=False,
            engine="openpyxl"
        )

    except Exception as e:

        raise Exception(
            f"No se pudo leer Excel:\n{e}"
        )

    # limpiar columnas

    df.columns = [
        clean_col(c)
        for c in df.columns
    ]

    df = df.where(
        pd.notnull(df),
        ""
    )

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


import io
import streamlit as st
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title=APP_NAME, page_icon="🎓", layout="wide")

st.title("🎓 PRÁCTICA+")
st.caption("Sistema inteligente de asignación de prácticas · Desarrollado por Diego J. Maldonado Guzmán")

with st.sidebar:
    st.header("Archivos")
    students_file = st.file_uploader("Alumnos (.xlsx)", type=["xlsx"], key="students")
    places_file = st.file_uploader("Plazas (.xlsx)", type=["xlsx"], key="places")
    prefer_morning = st.checkbox("Preferir mañana cuando el turno sea indiferente", value=True)
    run_clicked = st.button("▶ Ejecutar asignación", type="primary", use_container_width=True)

st.info("Sube los dos Excel y pulsa **Ejecutar asignación**. Los archivos se procesan durante la sesión de la app.")

if "resultado" not in st.session_state:
    st.session_state.resultado = None

if run_clicked:
    if students_file is None or places_file is None:
        st.error("Debes subir el Excel de alumnos y el Excel de plazas.")
    else:
        try:
            df_s = leer_excel(students_file)
            df_p = leer_excel(places_file)
            plazas, alias, ocupacion = preparar_plazas(df_p)
            resultado = asignar(df_s, plazas, alias, ocupacion, prefer_morning)
            st.session_state.resultado = resultado
        except Exception as e:
            st.session_state.resultado = None
            st.exception(e)

resultado = st.session_state.resultado
if resultado is not None:
    total = len(resultado)
    ok = int((resultado["Estado"] == "asignado").sum())
    first = int(((resultado["Estado"] == "asignado") & (resultado["Preferencia"] == 1)).sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alumnos", total)
    c2.metric("Asignados", ok)
    c3.metric("1.ª preferencia", first)
    c4.metric("Sin plaza", total - ok)

    def marcar_fila(row):
        if row["Estado"] == "asignado" and row["Preferencia"] == 1:
            return ["background-color: #d4f4dd"] * len(row)
        if row["Estado"] == "asignado":
            return ["background-color: #fff3cd"] * len(row)
        return ["background-color: #f8d7da"] * len(row)

    st.subheader("Resultado")
    st.dataframe(resultado.style.apply(marcar_fila, axis=1), use_container_width=True, hide_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resultado.to_excel(writer, sheet_name="Asignaciones", index=False)
        ws = writer.book["Asignaciones"]
        fill_ok = PatternFill("solid", fgColor="C6EFCE")
        fill_warn = PatternFill("solid", fgColor="FFF2CC")
        fill_bad = PatternFill("solid", fgColor="F8D7DA")
        bold = Font(bold=True)
        for cell in ws[1]:
            cell.font = bold
        for row in ws.iter_rows(min_row=2):
            estado = row[-1].value
            pref = row[-2].value
            if estado == "asignado" and pref == 1:
                fill = fill_ok
            elif estado == "asignado":
                fill = fill_warn
            else:
                fill = fill_bad
            for cell in row:
                cell.fill = fill
        for column_cells in ws.columns:
            max_len = max(len(str(c.value or "")) for c in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 10), 45)
    output.seek(0)

    st.download_button(
        "📊 Descargar asignaciones.xlsx",
        data=output.getvalue(),
        file_name="asignaciones.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
