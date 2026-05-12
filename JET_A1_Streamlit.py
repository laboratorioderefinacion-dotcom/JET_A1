#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import numpy as np
from mailmerge import MailMerge
from datetime import datetime
import re
from pathlib import Path
import tempfile

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="JET A1 - Informe", layout="wide")

BASE_DIR = Path(__file__).parent

PLANTILLA_OUA = BASE_DIR / "JET_A1.docx"
PLANTILLA_SIN_OUA = BASE_DIR / "JET_A1_SIN_OUA.docx"
RUTA_ALCANCE = BASE_DIR / "alcance_acreditacion_JET.csv"

# =========================
# HELPERS
# =========================
def html_unescape_basic(x):
    """Convierte entidades HTML típicas que aparecen en CSV a símbolos."""
    if not isinstance(x, str):
        return x
    return (
        x.replace("&gt;", ">")
         .replace("&lt;", "<")
         .replace("&amp;", "&")
         .strip()
    )

def read_csv_lims(uploaded_file):
    df = pd.read_csv(uploaded_file, encoding="latin1", sep=";", header=None)
    # normaliza strings por si vienen con entidades HTML
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].apply(lambda v: html_unescape_basic(v) if isinstance(v, str) else v)
    return df

def read_csv_alcance():
    df = pd.read_csv(RUTA_ALCANCE, encoding="latin1", sep=";", header=None)
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].apply(lambda v: html_unescape_basic(v) if isinstance(v, str) else v)
    return df

def get_val(df, col_key, key, col_val, default=None):
    sub = df.loc[df[col_key] == key, col_val]
    if sub.empty:
        return default
    v = sub.values[0]
    if pd.isna(v):
        return default
    return html_unescape_basic(v) if isinstance(v, str) else v

def extraer_valor_celda_float(df, nombre_celda):
    sub = df.loc[df[1] == nombre_celda, 4]
    if sub.empty:
        return None
    v = sub.values[0]
    if pd.isna(v):
        return None
    if isinstance(v, str):
        v = html_unescape_basic(v)
        if v.startswith(">") or v.startswith("<"):
            return v
        try:
            return float(v.replace(",", "."))
        except ValueError:
            return None
    try:
        return float(v)
    except Exception:
        return None

def extraer_valor_celda_int(df, nombre_celda):
    sub = df.loc[df[1] == nombre_celda, 4]
    if sub.empty:
        return None
    v = sub.values[0]
    if pd.isna(v):
        return None
    if isinstance(v, str):
        v = html_unescape_basic(v)
        if v.startswith(">") or v.startswith("<"):
            return v
        try:
            return int(float(v.replace(",", ".")))
        except ValueError:
            return None
    try:
        return int(v)
    except Exception:
        return None

def extraer_valor_norma(df, nombre_celda):
    sub = df.loc[df[1] == nombre_celda, 2]
    if sub.empty:
        return None
    v = sub.values[0]
    if pd.isna(v):
        return None
    return html_unescape_basic(v) if isinstance(v, str) else str(v)

def fmt_num(v, decs, default="-----"):
    if v is None:
        return default
    if isinstance(v, str):  # ej: ">0,001" o "<0,1"
        return v
    try:
        return f"{v:.{decs}f}".replace(".", ",")
    except Exception:
        return default

def convertir_fecha_informe(celda_fecha_informe):
    # Tu script asumía formato MM/DD/YYYY y lo pasaba a DD/MM/YYYY
    try:
        dt = datetime.strptime(celda_fecha_informe, "%m/%d/%Y")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return celda_fecha_informe

def get_multi_nombre(df, patrones_regex):
    """
    Busca el primer match en df[1] para diferentes nombres posibles.
    Devuelve (valor, norma).
    """
    extracted = df[1].astype(str).str.extract(patrones_regex, expand=False)
    dfx = df[extracted.notna()]
    if dfx.empty:
        return None, None
    val = dfx.iloc[0, 4]
    norma = dfx.iloc[0, 2]
    if isinstance(val, str):
        val = html_unescape_basic(val)
        if val.startswith(">") or val.startswith("<"):
            return val, html_unescape_basic(str(norma))
        try:
            return float(val.replace(",", ".")), html_unescape_basic(str(norma))
        except ValueError:
            return None, html_unescape_basic(str(norma))
    return val, html_unescape_basic(str(norma))

def evaluar_acreditacion_y_plantilla(df_alcance, fecha_hoy, celdas_float, celdas_int):
    """
    Replica tu lógica:
    - Si fecha_hoy está entre inicio/fin => evaluar rangos por norma
    - Si hay norma acreditada => plantilla OUA, si no => sin OUA
    """
    try:
        fecha_inicio = datetime.strptime(str(df_alcance.iloc[1, 5]), "%d/%m/%Y")
        fecha_fin = datetime.strptime(str(df_alcance.iloc[1, 6]), "%d/%m/%Y")
    except Exception:
        return PLANTILLA_SIN_OUA, []

    if not (fecha_inicio <= fecha_hoy <= fecha_fin):
        return PLANTILLA_SIN_OUA, []

    normas_acreditadas = []

    # floats
    for norma, valor in celdas_float.items():
        if norma is None:
            continue
        if isinstance(valor, str) or valor is None:
            continue  # no se puede chequear rango
        if norma in df_alcance[0].values:
            vmin = df_alcance.loc[df_alcance[0] == norma, 1].values[0]
            vmax = df_alcance.loc[df_alcance[0] == norma, 2].values[0]
            try:
                vmin = float(str(vmin).replace(",", "."))
                vmax = float(str(vmax).replace(",", "."))
            except Exception:
                continue
            if vmin <= float(valor) <= vmax:
                normas_acreditadas.append(norma)

    # ints
    for norma, valor in celdas_int.items():
        if norma is None:
            continue
        if isinstance(valor, str) or valor is None:
            continue
        if norma in df_alcance[0].values:
            try:
                vmin = int(df_alcance.loc[df_alcance[0] == norma, 1].values[0])
                vmax = int(df_alcance.loc[df_alcance[0] == norma, 2].values[0])
            except Exception:
                continue
            if vmin <= int(valor) <= vmax:
                normas_acreditadas.append(norma)

    plantilla = PLANTILLA_OUA if len(normas_acreditadas) > 0 else PLANTILLA_SIN_OUA
    return plantilla, normas_acreditadas

def generate_docx_bytes(plantilla_path, merge_dict):
    """
    Genera el docx en un archivo temporal y devuelve bytes.
    (MailMerge trabaja con paths)
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        out_path = Path(tmpdir) / "informe.docx"
        doc = MailMerge(str(plantilla_path))
        # merge en una sola llamada (más rápido)
        doc.merge(**{k: ("" if v is None else str(v)) for k, v in merge_dict.items()})
        doc.write(str(out_path))
        return out_path.read_bytes()

# =========================
# UI
# =========================
st.title("✈️ JET A1 - Generación de Informe Word")

lims_file = st.file_uploader("Subí el CSV exportado del LIMS", type=["csv"])

if lims_file is None:
    st.info("Subí el CSV del LIMS para continuar.")
    st.stop()

# Validar que existan los archivos en el repo
if not PLANTILLA_OUA.exists() or not PLANTILLA_SIN_OUA.exists():
    st.error("No encuentro las plantillas en /templates. Verificá que estén en el repo.")
    st.stop()

if not RUTA_ALCANCE.exists():
    st.error("No encuentro el alcance en /data/alcance_acreditacion_JET.csv.")
    st.stop()

# =========================
# PROC
# =========================
df_a = read_csv_lims(lims_file)
df_alcance = read_csv_alcance()
fecha_hoy = datetime.now()

# Campos generales
celda_tanque = get_val(df_a, 0, "R-SamplePoint", 4, default="---")
celda_fecha_aprob = str(get_val(df_a, 0, "Fecha de Aprobación", 4, default="-----"))
celda_lims = get_val(df_a, 0, "Número de Muestra", 4, default="-----")
celda_numElab = get_val(df_a, 0, "Número de Elaboración", 4, default="-----")
celda_fecha = str(get_val(df_a, 0, "Fecha", 4, default="-----"))

celda_color = str(get_val(df_a, 1, "Color Saybolt final", 4, default="-----"))
celda_corrosion = str(get_val(df_a, 1, "Corrosion", 4, default="-----"))
celda_deposito = str(get_val(df_a, 1, "Depósitos en el tubo", 4, default="-----"))
celda_fecha_informe = str(get_val(df_a, 1, "Fecha de informe", 4, default="-----")).replace("-", "/")
celda_fecha_informe_2 = convertir_fecha_informe(celda_fecha_informe)

# Float celdas
celda_acidez = extraer_valor_celda_float(df_a, "Acidez Total")
celda_pto_inicial = extraer_valor_celda_float(df_a, "Punto Inicial")
celda_dest_10 = extraer_valor_celda_float(df_a, "10% vol")
celda_dest_50 = extraer_valor_celda_float(df_a, "50% vol")
celda_dest_90 = extraer_valor_celda_float(df_a, "90% vol")
celda_pto_final = extraer_valor_celda_float(df_a, "Punto Final")
celda_residuo = extraer_valor_celda_float(df_a, "Residuo")
celda_pto_inf = extraer_valor_celda_float(df_a, "Punto de Inflamación TAG")
celda_densidad = extraer_valor_celda_float(df_a, "Densidad promedio a 15º")
celda_congelacion = extraer_valor_celda_float(df_a, "Punto de Congelación")
celda_viscosidad = extraer_valor_celda_float(df_a, "Viscosidad Cinemática D445 (corrección)")
celda_calor = extraer_valor_celda_float(df_a, "Poder Calorífico Neto")
celda_pto_hum = extraer_valor_celda_float(df_a, "Punto de Humo (método automático)")
celda_JFTOT = extraer_valor_celda_float(df_a, "Caída de presión en el filtro")
celda_antiestatico = extraer_valor_celda_float(df_a, "Aditivo Antiestático (AAE)")
celda_no_hidrop = extraer_valor_celda_float(df_a, "Componentes no hidroprocesados")
celda_hidrop = extraer_valor_celda_float(df_a, "Componentes severamente hidroprocesados")
celda_sintetico = extraer_valor_celda_float(df_a, "Componentes sintéticos")
celda_copros = extraer_valor_celda_float(df_a, "Componentes coprocesados")
celda_particulado = extraer_valor_celda_float(df_a, "Particulado")
celda_vol_filt = extraer_valor_celda_float(df_a, "Volumen Filtrado")

densidad = celda_densidad * 1000 if isinstance(celda_densidad, (float, int)) else None

# Int celdas
celda_CP_4 = extraer_valor_celda_int(df_a, ">= 4 micrometros")
celda_CP_6 = extraer_valor_celda_int(df_a, ">= 6 micrometros")
celda_CP_14 = extraer_valor_celda_int(df_a, ">= 14 micrometros")
celda_CP_21 = extraer_valor_celda_int(df_a, ">= 21 micrometros")
celda_CP_25 = extraer_valor_celda_int(df_a, ">= 25 micrometros")
celda_CP_30 = extraer_valor_celda_int(df_a, ">= 30 micrometros")

celda_I_4 = extraer_valor_celda_int(df_a, "Código ISO >=4 micras")
celda_I_6 = extraer_valor_celda_int(df_a, "Código ISO >=6 micras")
celda_I_14 = extraer_valor_celda_int(df_a, "Código ISO >=14 micras")
celda_I_21 = extraer_valor_celda_int(df_a, "Código ISO >=21 micras")
celda_I_25 = extraer_valor_celda_int(df_a, "Código ISO >=25 micras")
celda_I_30 = extraer_valor_celda_int(df_a, "Código ISO >= 30 micras")

celda_temperatura = extraer_valor_celda_int(df_a, "Temperatura de Control")
celda_gomas = extraer_valor_celda_int(df_a, "Gomas Existentes")
celda_MSEP = extraer_valor_celda_int(df_a, "MSEP A rating")
celda_conductividad = extraer_valor_celda_int(df_a, "Conductividad eléctrica")
celda_volumen = extraer_valor_celda_int(df_a, "Volumen Total del Lote")

# Temperatura conductividad (búsqueda por 2 condiciones)
sub_temp = df_a[(df_a[0] == "R-JFCONDUC") & (df_a[1] == "Temperatura")][4]
celda_temp_cond = sub_temp.values[0] if not sub_temp.empty else "-----"

# Normas
norma_color = extraer_valor_norma(df_a, "Color Saybolt final")
norma_acidez = extraer_valor_norma(df_a, "Acidez Total")
norma_dest = extraer_valor_norma(df_a, "Punto Inicial")
norma_pto_inf = extraer_valor_norma(df_a, "Punto de Inflamación TAG")
norma_densidad = extraer_valor_norma(df_a, "Densidad promedio a 15º")
norma_congelacion = extraer_valor_norma(df_a, "Punto de Congelación")
norma_viscosidad = extraer_valor_norma(df_a, "Viscosidad Cinemática D445 (corrección)")
norma_calor = extraer_valor_norma(df_a, "Poder Calorífico Neto")
norma_pto_hum = extraer_valor_norma(df_a, "Punto de Humo (método automático)")
norma_JFTOT = extraer_valor_norma(df_a, "Caída de presión en el filtro")
norma_gomas = extraer_valor_norma(df_a, "Gomas Existentes")
norma_MSEP = extraer_valor_norma(df_a, "MSEP A rating")
norma_conductividad = extraer_valor_norma(df_a, "Conductividad eléctrica")
norma_corrosion = extraer_valor_norma(df_a, "Corrosion")
norma_conteo = extraer_valor_norma(df_a, "Código ISO >= 30 micras") or "IP_565"

# Color especial "+>"
if isinstance(celda_color, str) and celda_color.startswith("+>"):
    celda_color_2 = ">+" + celda_color[2:4]
else:
    celda_color_2 = celda_color

# Multi-nombres: Azufre / Aromáticos
celda_azufre, norma_azufre = get_multi_nombre(df_a, r"(Azufre \(% peso\)|Azufre %peso)")
celda_aromaticos, norma_aromaticos = get_multi_nombre(df_a, r"(Aromáticos|Aromáticos Totales)")

# Mercaptanes / Doctor
celda_mercaptanes = extraer_valor_celda_float(df_a, "Azufre Mercaptan")
norma_mercaptanes = extraer_valor_norma(df_a, "Azufre Mercaptan") or "ASTM_D_3227"
if celda_mercaptanes is None:
    celda_mercaptanes = "-----"

doctor_val = get_val(df_a, 1, "Reacción Doctor", 4, default=None)
if doctor_val is None:
    celda_doctor = "-----"
    norma_doctor = "ASTM_D_4952"
else:
    if str(doctor_val).strip().upper() == "NEGATIVA":
        celda_doctor = "Negativa"
        norma_doctor = extraer_valor_norma(df_a, "Reacción Doctor") or "ASTM_D_4952"
    else:
        celda_doctor = "-----"
        norma_doctor = "ASTM_D_4952"

# Naftalenos
celda_naftalenos = extraer_valor_celda_float(df_a, "Naftalenos")
norma_naftalenos = extraer_valor_norma(df_a, "Naftalenos") or "ASTM_D_1840"
if celda_naftalenos is None:
    celda_naftalenos = "-----"

# Pérdidas / Antioxidante
celda_perdidas = extraer_valor_celda_float(df_a, "Pérdidas")
if celda_perdidas is None:
    celda_perdidas = 0.0

celda_antioxidante = extraer_valor_celda_float(df_a, "Antioxidante en el batch final")
if celda_antioxidante is None:
    celda_antioxidante = 0.0

# =========================
# ACREDITACIÓN (igual idea)
# =========================
celdas_float = {
    norma_acidez: celda_acidez,
    norma_dest: celda_pto_inicial,
    norma_pto_inf: celda_pto_inf,
    norma_densidad: celda_densidad,
    norma_congelacion: celda_congelacion,
    norma_viscosidad: celda_viscosidad,
    norma_calor: celda_calor,
    norma_pto_hum: celda_pto_hum,
    norma_azufre: celda_azufre,
    norma_aromaticos: celda_aromaticos,
}

if isinstance(celda_naftalenos, (float, int)):
    celdas_float[norma_naftalenos] = celda_naftalenos
if isinstance(celda_mercaptanes, (float, int)):
    celdas_float[norma_mercaptanes] = celda_mercaptanes

celdas_int = {
    norma_conteo: celda_I_30,
    norma_gomas: celda_gomas,
    norma_MSEP: celda_MSEP,
    norma_conductividad: celda_conductividad,
}

plantilla, normas_acreditadas = evaluar_acreditacion_y_plantilla(df_alcance, fecha_hoy, celdas_float, celdas_int)

# =========================
# FORMATEOS (como tu script)
# =========================
celda_acidez_2 = fmt_num(celda_acidez, 3)
celda_pto_inicial_2 = fmt_num(celda_pto_inicial, 1)
celda_dest_10_2 = fmt_num(celda_dest_10, 1)
celda_dest_50_2 = fmt_num(celda_dest_50, 1)
celda_dest_90_2 = fmt_num(celda_dest_90, 1)
celda_residuo_2 = fmt_num(celda_residuo, 1)
celda_pto_final_2 = fmt_num(celda_pto_final, 1)
celda_perdidas_2 = fmt_num(celda_perdidas, 1, default="0,0")
celda_pto_inf_2 = fmt_num(celda_pto_inf, 1)
celda_densidad_2 = fmt_num(densidad, 1)
celda_congelacion_2 = fmt_num(celda_congelacion, 1)
celda_viscosidad_2 = fmt_num(celda_viscosidad, 3)
celda_calor_2 = fmt_num(celda_calor, 3)
celda_pto_hum_2 = fmt_num(celda_pto_hum, 1)
celda_JFTOT_2 = fmt_num(celda_JFTOT, 1)
celda_antiestatico_2 = fmt_num(celda_antiestatico, 2)
celda_antioxidante_2 = fmt_num(celda_antioxidante, 1)
celda_aromaticos_2 = fmt_num(celda_aromaticos, 1)
celda_naftalenos_2 = fmt_num(celda_naftalenos, 2)
celda_mercaptanes_2 = fmt_num(celda_mercaptanes, 4)

# Azufre con decimales variables
if isinstance(celda_azufre, (float, int)):
    x = float(celda_azufre)
    if x < 0.00001:   celda_azufre_2 = fmt_num(x, 8)
    elif x < 0.0001:  celda_azufre_2 = fmt_num(x, 7)
    elif x < 0.001:   celda_azufre_2 = fmt_num(x, 6)
    elif x < 0.01:    celda_azufre_2 = fmt_num(x, 5)
    elif x < 0.1:     celda_azufre_2 = fmt_num(x, 4)
    elif x < 1:       celda_azufre_2 = fmt_num(x, 3)
    else:             celda_azufre_2 = fmt_num(x, 2)
else:
    celda_azufre_2 = str(celda_azufre) if celda_azufre is not None else "-----"

# Composición
celda_no_hidrop_2 = fmt_num(celda_no_hidrop, 1)
celda_hidrop_2 = fmt_num(celda_hidrop, 1)
celda_sintetico_2 = fmt_num(celda_sintetico, 1)
celda_copros_2 = fmt_num(celda_copros, 1)

# Particulado
if celda_particulado is None or isinstance(celda_particulado, str):
    celda_cont_part = "-----"
else:
    celda_cont_part = (
        f"{fmt_num(celda_particulado, 2)}\n"
        f"(Vol. Filtrado: {fmt_num(celda_vol_filt, 2)} L)"
    )

# Conteo
def cont(cp, iso):
    if cp is None or iso is None:
        return "-----"
    if isinstance(cp, str) or isinstance(iso, str):
        return "-----"
    return f"{cp} / {int(iso):02d}"

celda_cont_I_4 = cont(celda_CP_4, celda_I_4)
celda_cont_I_6 = cont(celda_CP_6, celda_I_6)
celda_cont_I_14 = cont(celda_CP_14, celda_I_14)
celda_cont_I_21 = cont(celda_CP_21, celda_I_21)
celda_cont_I_25 = cont(celda_CP_25, celda_I_25)
celda_cont_I_30 = cont(celda_CP_30, celda_I_30)

# Observaciones: fecha MSEP
obs = ""
rep = ""
sub_obs = df_a.loc[df_a[1] == "Observaciones", 4]
if not sub_obs.empty and pd.notna(sub_obs.values[0]):
    observaciones = str(sub_obs.values[0])
    fecha_MSEP = re.search(r"MSEP.*?(\d{1,2}/\d{1,2}/\d{2,4})", observaciones, re.IGNORECASE)
    if fecha_MSEP:
        fecha_MSEP_2 = fecha_MSEP.group(1)
        obs = f"(4) La conductividad eléctrica y el índice de separación (MSEP) fueron analizados sobre muestra extraída el {fecha_MSEP_2}."
        rep = "(4)"

# Aromáticos FIA / HPLC
esp_arom = ""
tot = ""
if norma_aromaticos == "ASTM_D_1319":
    esp_arom = "25,0"
    tot = "(% vol.)"
elif norma_aromaticos == "ASTM_D_6379":
    esp_arom = "26,5"
    tot = "totales (% vol.)"

# =========================
# MERGE DICT (igual a tu template)
# =========================

# Nombre del archivo CSV (sin extensión)
nombre_csv = Path(lims_file.name).stem  # ej: "12345" si el archivo es "12345.csv"

tanque_short = str(celda_tanque)[3:] if isinstance(celda_tanque, str) and len(str(celda_tanque)) >= 4 else str(celda_tanque)

# Ahora el Word se llama con el nombre del CSV, no con el LIMS
nombre_archivo_nuevo = f"{nombre_csv} TK {tanque_short} JET A1"

merge = {
    "informe": str(nombre_csv),
    "tanque": tanque_short,
    "fecha_aprob": celda_fecha_aprob.replace("-", "/"),
    "fecha_informe": celda_fecha_informe_2,
    "lims": str(celda_lims),
    "numElab": str(celda_numElab),
    "fecha": str(celda_fecha).replace("-", "/"),
    "volumen": str(celda_volumen),

    "color": str(celda_color_2),
    "n_color": (norma_color or "").replace("_", " "),

    "cont_part": str(celda_cont_part),
    "cont_I_4": str(celda_cont_I_4),
    "cont_I_6": str(celda_cont_I_6),
    "cont_I_14": str(celda_cont_I_14),
    "cont_I_21": str(celda_cont_I_21),
    "cont_I_25": str(celda_cont_I_25),
    "cont_I_30": str(celda_cont_I_30),
    "n_conteo": (norma_conteo or "IP_565").replace("_", " "),

    "acidez": str(celda_acidez_2),
    "n_acidez": (norma_acidez or "").replace("_", " "),

    "aromaticos": str(celda_aromaticos_2),
    "n_aromaticos": (norma_aromaticos or "").replace("_", " "),

    "azufre": str(celda_azufre_2),
    "n_azuf": (norma_azufre or "").replace("_", " "),

    "mercaptanes": str(celda_mercaptanes_2),
    "n_mercaptanes": (norma_mercaptanes or "").replace("_", " "),

    "doctor": str(celda_doctor),
    "n_doctor": (norma_doctor or "").replace("_", " "),

    "no_hidrop": str(celda_no_hidrop_2),
    "hidrop": str(celda_hidrop_2),
    "sintetico": str(celda_sintetico_2),
    "copros": str(celda_copros_2),

    "pto_inicial": str(celda_pto_inicial_2),
    "n_dest": (norma_dest or "").replace("_", " "),
    "dest_10": str(celda_dest_10_2),
    "dest_50": str(celda_dest_50_2),
    "dest_90": str(celda_dest_90_2),
    "pto_final": str(celda_pto_final_2),
    "residuo": str(celda_residuo_2),
    "perdidas": str(celda_perdidas_2),

    "pto_inf": str(celda_pto_inf_2),
    "n_pto_inf": (norma_pto_inf or "").replace("_", " "),

    "densidad": str(celda_densidad_2),
    "n_dens": (norma_densidad or "").replace("_", " "),

    "congelacion": str(celda_congelacion_2),
    "n_congel": (norma_congelacion or "").replace("_", " "),

    "viscosidad": str(celda_viscosidad_2),
    "n_visco": (norma_viscosidad or "").replace("_", " "),

    "calor": str(celda_calor_2),
    "n_calor": (norma_calor or "").replace("_", " "),

    "n_pHum": (norma_pto_hum or "").replace("_", " "),
    "naftalenos": str(celda_naftalenos_2),
    "n_naft": (norma_naftalenos or "").replace("_", " "),

    "corrosion": str(celda_corrosion),
    "n_corr": (norma_corrosion or "").replace("_", " "),

    "temperatura": str(celda_temperatura),
    "JFTOT": str(celda_JFTOT_2),
    "deposito": str(celda_deposito),
    "n_JFTOT": (norma_JFTOT or "").replace("_", " "),

    "gomas": str(celda_gomas),
    "n_gomas": (norma_gomas or "").replace("_", " "),

    "MSEP": str(celda_MSEP),
    "n_MSEP": (norma_MSEP or "").replace("_", " "),

    "conductividad": str(celda_conductividad),
    "n_cond": (norma_conductividad or "").replace("_", " "),
    "temp_cond": str(celda_temp_cond),

    "antioxidante": str(celda_antioxidante_2),
    "antiestatico": str(celda_antiestatico_2),

    "observaciones": obs,
    "rep": rep,

    "esp_arom": esp_arom,
    "tot": tot,
}

# Condicional pto_hum_s / pto_hum_i
if isinstance(celda_pto_hum, (float, int)):
    merge["pto_hum_s"] = celda_pto_hum_2 if float(celda_pto_hum) >= 25 else "-----"
    merge["pto_hum_i"] = celda_pto_hum_2 if float(celda_pto_hum) < 25 else "-----"
else:
    merge["pto_hum_s"] = "-----"
    merge["pto_hum_i"] = "-----"

# Flags "(*)" si plantilla OUA y no acreditada
if plantilla == PLANTILLA_OUA:
    def mark(norma, field):
        merge[field] = "" if (norma in normas_acreditadas) else "(*)"
    mark(norma_color, "col")
    mark(norma_conteo, "conteo")
    mark(norma_acidez, "acid")
    mark(norma_aromaticos, "arom")
    mark(norma_azufre, "azuf")
    mark(norma_mercaptanes, "merc")
    mark(norma_doctor, "doc")
    mark(norma_dest, "dest")
    mark(norma_pto_inf, "PI")
    mark(norma_densidad, "dens")
    mark(norma_congelacion, "cong")
    mark(norma_viscosidad, "vis")
    mark(norma_calor, "cal")
    mark(norma_pto_hum, "PH")
    mark(norma_naftalenos, "naft")
    mark(norma_corrosion, "corr")
    mark(norma_gomas, "gom")
    mark(norma_MSEP, "MS")
    mark(norma_conductividad, "cond")

# =========================
# BOTÓN GENERAR
# =========================

if st.button("✅ Generar informe", type="primary"):
    try:
        doc_bytes = generate_docx_bytes(plantilla, merge)
        st.success("Informe generado correctamente.")
        st.download_button(
            "⬇️ Descargar informe Word",
            data=doc_bytes,
            file_name=f"{nombre_archivo_nuevo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"Error generando el Word: {e}")

