#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import numpy as np
from mailmerge import MailMerge
from docx import Document
from datetime import datetime
from io import BytesIO
import tempfile
import re
from pathlib import Path

# =========================
# Config general Streamlit
# =========================
st.set_page_config(page_title="JET A1 - Informe automático", layout="wide")

BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
DATA_DIR = BASE_DIR / "data"

TEMPLATE_OUA = TEMPLATES_DIR / "JET_A1.docx"
TEMPLATE_SIN_OUA = TEMPLATES_DIR / "JET_A1_SIN_OUA.docx"
ALCANCE_PATH = DATA_DIR / "alcance_acreditacion_JET.csv"

# =========================
# Helpers
# =========================

def html_unescape_basic(x: str) -> str:
    """Convierte entidades HTML típicas (&gt; &lt; &amp;) a símbolos reales."""
    if not isinstance(x, str):
        return x
    return (x.replace("&gt;", ">")
             .replace("&lt;", "<")
             .replace("&amp;", "&"))

def safe_value(df, col_key, col_val, key, default=None):
    """Devuelve df.loc[df[col_key]==key, col_val].values[0] con control de vacío."""
    sub = df.loc[df[col_key] == key, col_val]
    if sub.empty:
        return default
    v = sub.values[0]
    if pd.isna(v):
        return default
    if isinstance(v, str):
        v = html_unescape_basic(v).strip()
    return v

def safe_value_2cols(df, col0, val0, col1, val1, col_return, default=None):
    sub = df[(df[col0] == val0) & (df[col1] == val1)][col_return]
    if sub.empty:
        return default
    v = sub.values[0]
    if pd.isna(v):
        return default
    if isinstance(v, str):
        v = html_unescape_basic(v).strip()
    return v

def extraer_float(df, nombre_celda):
    sub = df.loc[df[1] == nombre_celda, 4]
    if sub.empty:
        return None
    v = sub.values[0]
    if pd.isna(v):
        return None
    if isinstance(v, str):
        v = html_unescape_basic(v).strip()
        if v.startswith(">") or v.startswith("<"):
            return v  # se conserva como string
        try:
            return float(v.replace(",", "."))
        except ValueError:
            return None
    try:
        return float(v)
    except Exception:
        return None

def extraer_int(df, nombre_celda):
    sub = df.loc[df[1] == nombre_celda, 4]
    if sub.empty:
        return None
    v = sub.values[0]
    if pd.isna(v):
        return None
    if isinstance(v, str):
        v = html_unescape_basic(v).strip()
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

def extraer_norma(df, nombre_celda):
    sub = df.loc[df[1] == nombre_celda, 2]
    if sub.empty:
        return None
    v = sub.values[0]
    if pd.isna(v):
        return None
    if isinstance(v, str):
        return html_unescape_basic(v).strip()
    return str(v)

def fmt_num(v, fmt, default="-----"):
    """Formatea floats con coma. Si string tipo '>0.1' lo deja tal cual."""
    if v is None:
        return default
    if isinstance(v, str):
        return v
    try:
        return format(v, fmt).replace(".", ",")
    except Exception:
        return default

def parse_fecha_mdy_a_dmy(fecha_str):
    """Convierte 'MM/DD/YYYY' o 'M/D/YYYY' a 'DD/MM/YYYY'."""
    if not fecha_str:
        return "-----"
    try:
        dt = datetime.strptime(fecha_str, "%m/%d/%Y")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        # Si ya viene como DD/MM/YYYY, lo devolvemos:
        return fecha_str

def get_azufre(df):
    azufre_extracted = df[1].astype(str).str.extract(r"(Azufre \(% peso\)|Azufre %peso)", expand=False)
    df_azuf = df[azufre_extracted.notna()]
    if df_azuf.empty:
        return None, None
    azufre = html_unescape_basic(str(df_azuf.iloc[0, 4])).strip()
    norma = html_unescape_basic(str(df_azuf.iloc[0, 2])).strip()
    if azufre.startswith(">") or azufre.startswith("<"):
        return azufre, norma
    try:
        return float(azufre.replace(",", ".")), norma
    except ValueError:
        return None, norma

def get_aromaticos(df):
    arom_extracted = df[1].astype(str).str.extract(r"(Aromáticos|Aromáticos Totales)", expand=False)
    df_arom = df[arom_extracted.notna()]
    if df_arom.empty:
        return None, None
    arom = html_unescape_basic(str(df_arom.iloc[0, 4])).strip()
    norma = html_unescape_basic(str(df_arom.iloc[0, 2])).strip()
    if arom.startswith(">") or arom.startswith("<"):
        return arom, norma
    try:
        return float(arom.replace(",", ".")), norma
    except ValueError:
        return None, norma

def acreditacion_y_plantilla(df_alcance, fecha_hoy, valores_por_norma):
    """
    Decide qué plantilla usar (OUA o sin OUA) y devuelve normas_acreditadas.
    valores_por_norma: dict {norma: valor (float/int)}
    """
    try:
        # Tu CSV de alcance usa fila 1 col 5 y 6 (según tu script)
        fecha_inicio = datetime.strptime(str(df_alcance.iloc[1, 5]), "%d/%m/%Y")
        fecha_fin = datetime.strptime(str(df_alcance.iloc[1, 6]), "%d/%m/%Y")
        dentro = (fecha_inicio <= fecha_hoy <= fecha_fin)
    except Exception:
        dentro = False

    if not dentro:
        return TEMPLATE_SIN_OUA, []

    normas_acreditadas = []

    for norma, valor in valores_por_norma.items():
        if norma is None or valor is None:
            continue

        # Si el valor es str (ej: ">0.001"), no podemos validar rango => no acreditar
        if isinstance(valor, str):
            continue

        if norma in df_alcance[0].values:
            vmin = df_alcance.loc[df_alcance[0] == norma, 1].values[0]
            vmax = df_alcance.loc[df_alcance[0] == norma, 2].values[0]

            # detecta si es int o float el rango
            try:
                if isinstance(valor, (int, np.integer)):
                    vmin = int(vmin)
                    vmax = int(vmax)
                else:
                    vmin = float(str(vmin).replace(",", "."))
                    vmax = float(str(vmax).replace(",", "."))
            except Exception:
                continue

            if vmin <= valor <= vmax:
                normas_acreditadas.append(norma)

    plantilla = TEMPLATE_OUA if len(normas_acreditadas) > 0 else TEMPLATE_SIN_OUA
    return plantilla, normas_acreditadas

def generar_docx(plantilla_path, datos_merge, extras_merge=None):
    """
    Genera DOCX (bytes) usando MailMerge.
    extras_merge: dict adicional para campos condicionales.
    """
    extras_merge = extras_merge or {}
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_out = Path(tmpdir) / "informe.docx"

        doc = MailMerge(str(plantilla_path))
        # merge principal
        for k, v in datos_merge.items():
            doc.merge(**{k: "" if v is None else str(v)})

        # merge adicional (condicional)
        for k, v in extras_merge.items():
            doc.merge(**{k: "" if v is None else str(v)})

        doc.write(str(tmp_out))

        b = tmp_out.read_bytes()
        return b

# =========================
# UI
# =========================

st.title("✈️ JET A1 - Generación automática de Informe (Word)")

col1, col2 = st.columns(2)
with col1:
    lims_file = st.file_uploader("Subí el CSV exportado desde LIMS", type=["csv"])
with col2:
    alcance_file = st.file_uploader("Subí (opcional) el CSV de Alcance Acreditación", type=["csv"])
    st.caption("Si no subís nada, se usa el alcance incluido en el repo (data/alcance_acreditacion_JET.csv).")

st.divider()

if lims_file is not None:
    # Leer LIMS
    try:
        df_a = pd.read_csv(lims_file, encoding="latin1", sep=";", header=None)
    except Exception as e:
        st.error(f"No pude leer el CSV LIMS. Detalle: {e}")
        st.stop()

    # Leer alcance
    try:
        if alcance_file is not None:
            df_alcance = pd.read_csv(alcance_file, encoding="latin1", sep=";", header=None)
        else:
            df_alcance = pd.read_csv(ALCANCE_PATH, encoding="latin1", sep=";", header=None)
    except Exception as e:
        st.error(f"No pude leer el CSV de alcance. Detalle: {e}")
        st.stop()

    fecha_hoy = datetime.now()

    # =========================
    # Extracciones principales
    # =========================
    celda_tanque = safe_value(df_a, 0, 4, "R-SamplePoint", default="---")
    celda_fecha_aprob = str(safe_value(df_a, 0, 4, "Fecha de Aprobación", default="-----"))
    celda_lims = safe_value(df_a, 0, 4, "Número de Muestra", default="-----")
    celda_numElab = safe_value(df_a, 0, 4, "Número de Elaboración", default="-----")
    celda_fecha = str(safe_value(df_a, 0, 4, "Fecha", default="-----"))

    celda_color = str(safe_value(df_a, 1, 4, "Color Saybolt final", default="-----"))
    celda_corrosion = str(safe_value(df_a, 1, 4, "Corrosion", default="-----"))
    celda_deposito = str(safe_value(df_a, 1, 4, "Depósitos en el tubo", default="-----"))
    celda_fecha_informe = str(safe_value(df_a, 1, 4, "Fecha de informe", default="-----")).replace("-", "/")
    celda_fecha_informe_2 = parse_fecha_mdy_a_dmy(celda_fecha_informe)

    # floats
    celda_acidez = extraer_float(df_a, "Acidez Total")
    celda_pto_inicial = extraer_float(df_a, "Punto Inicial")
    celda_dest_10 = extraer_float(df_a, "10% vol")
    celda_dest_50 = extraer_float(df_a, "50% vol")
    celda_dest_90 = extraer_float(df_a, "90% vol")
    celda_pto_final = extraer_float(df_a, "Punto Final")
    celda_residuo = extraer_float(df_a, "Residuo")
    celda_pto_inf = extraer_float(df_a, "Punto de Inflamación TAG")
    celda_densidad = extraer_float(df_a, "Densidad promedio a 15º")
    celda_congelacion = extraer_float(df_a, "Punto de Congelación")
    celda_viscosidad = extraer_float(df_a, "Viscosidad Cinemática D445 (corrección)")
    celda_calor = extraer_float(df_a, "Poder Calorífico Neto")
    celda_pto_hum = extraer_float(df_a, "Punto de Humo (método automático)")
    celda_JFTOT = extraer_float(df_a, "Caída de presión en el filtro")
    celda_antiestatico = extraer_float(df_a, "Aditivo Antiestático (AAE)")
    celda_no_hidrop = extraer_float(df_a, "Componentes no hidroprocesados")
    celda_hidrop = extraer_float(df_a, "Componentes severamente hidroprocesados")
    celda_sintetico = extraer_float(df_a, "Componentes sintéticos")
    celda_copros = extraer_float(df_a, "Componentes coprocesados")
    celda_particulado = extraer_float(df_a, "Particulado")
    celda_vol_filt = extraer_float(df_a, "Volumen Filtrado")

    densidad = None
    if isinstance(celda_densidad, (float, int)):
        densidad = celda_densidad * 1000

    # int
    celda_CP_4 = extraer_int(df_a, ">= 4 micrometros") or extraer_int(df_a, "&gt;= 4 micrometros")
    celda_CP_6 = extraer_int(df_a, ">= 6 micrometros") or extraer_int(df_a, "&gt;= 6 micrometros")
    celda_CP_14 = extraer_int(df_a, ">= 14 micrometros") or extraer_int(df_a, "&gt;= 14 micrometros")
    celda_CP_21 = extraer_int(df_a, ">= 21 micrometros") or extraer_int(df_a, "&gt;= 21 micrometros")
    celda_CP_25 = extraer_int(df_a, ">= 25 micrometros") or extraer_int(df_a, "&gt;= 25 micrometros")
    celda_CP_30 = extraer_int(df_a, ">= 30 micrometros") or extraer_int(df_a, "&gt;= 30 micrometros")

    celda_I_4 = extraer_int(df_a, "Código ISO >=4 micras") or extraer_int(df_a, "Código ISO &gt;=4 micras")
    celda_I_6 = extraer_int(df_a, "Código ISO >=6 micras") or extraer_int(df_a, "Código ISO &gt;=6 micras")
    celda_I_14 = extraer_int(df_a, "Código ISO >=14 micras") or extraer_int(df_a, "Código ISO &gt;=14 micras")
    celda_I_21 = extraer_int(df_a, "Código ISO >=21 micras") or extraer_int(df_a, "Código ISO &gt;=21 micras")
    celda_I_25 = extraer_int(df_a, "Código ISO >=25 micras") or extraer_int(df_a, "Código ISO &gt;=25 micras")
    celda_I_30 = extraer_int(df_a, "Código ISO >= 30 micras") or extraer_int(df_a, "Código ISO &gt;= 30 micras")

    celda_temperatura = extraer_int(df_a, "Temperatura de Control")
    celda_gomas = extraer_int(df_a, "Gomas Existentes")
    celda_MSEP = extraer_int(df_a, "MSEP A rating")
    celda_conductividad = extraer_int(df_a, "Conductividad eléctrica")
    celda_volumen = extraer_int(df_a, "Volumen Total del Lote")

    celda_temp_cond = safe_value_2cols(df_a, 0, "R-JFCONDUC", 1, "Temperatura", 4, default="-----")

    # normas
    norma_color = extraer_norma(df_a, "Color Saybolt final")
    norma_acidez = extraer_norma(df_a, "Acidez Total")
    norma_dest = extraer_norma(df_a, "Punto Inicial")
    norma_pto_inf = extraer_norma(df_a, "Punto de Inflamación TAG")
    norma_densidad = extraer_norma(df_a, "Densidad promedio a 15º")
    norma_congelacion = extraer_norma(df_a, "Punto de Congelación")
    norma_viscosidad = extraer_norma(df_a, "Viscosidad Cinemática D445 (corrección)")
    norma_calor = extraer_norma(df_a, "Poder Calorífico Neto")
    norma_pto_hum = extraer_norma(df_a, "Punto de Humo (método automático)")
    norma_JFTOT = extraer_norma(df_a, "Caída de presión en el filtro")
    norma_gomas = extraer_norma(df_a, "Gomas Existentes")
    norma_MSEP = extraer_norma(df_a, "MSEP A rating")
    norma_conductividad = extraer_norma(df_a, "Conductividad eléctrica")
    norma_corrosion = extraer_norma(df_a, "Corrosion")
    norma_conteo = extraer_norma(df_a, "Código ISO >= 30 micras") or extraer_norma(df_a, "Código ISO &gt;= 30 micras") or "IP_565"

    # Ajuste de color "+>"
    celda_color = html_unescape_basic(celda_color)
    if celda_color.startswith("+>"):
        celda_color_2 = ">+" + celda_color[2:4]
    else:
        celda_color_2 = celda_color

    # Azufre y aromáticos con multi-nombre
    celda_azufre, norma_azufre = get_azufre(df_a)
    celda_aromaticos, norma_aromaticos = get_aromaticos(df_a)

    # Mercaptanes / Doctor / Naftalenos
    celda_mercaptanes = extraer_float(df_a, "Azufre Mercaptan")
    norma_mercaptanes = extraer_norma(df_a, "Azufre Mercaptan") or "ASTM_D_3227"
    if celda_mercaptanes is None:
        celda_mercaptanes = "-----"

    doctor_raw = safe_value(df_a, 1, 4, "Reacción Doctor", default=None)
    if doctor_raw is None:
        celda_doctor = "-----"
        norma_doctor = "ASTM_D_4952"
    else:
        doctor_raw = str(doctor_raw).strip().upper()
        if doctor_raw == "NEGATIVA":
            celda_doctor = "Negativa"
            norma_doctor = extraer_norma(df_a, "Reacción Doctor") or "ASTM_D_4952"
        else:
            celda_doctor = "-----"
            norma_doctor = "ASTM_D_4952"

    celda_naftalenos = extraer_float(df_a, "Naftalenos")
    norma_naftalenos = extraer_norma(df_a, "Naftalenos") or "ASTM_D_1840"
    if celda_naftalenos is None:
        celda_naftalenos = "-----"

    # Pérdidas / antioxidante
    celda_perdidas = extraer_float(df_a, "Pérdidas")
    if celda_perdidas is None:
        celda_perdidas = 0.0

    celda_antioxidante = extraer_float(df_a, "Antioxidante en el batch final")
    if celda_antioxidante is None:
        celda_antioxidante = 0.0

    # =========================
    # Acreditación y plantilla
    # =========================
    valores_por_norma = {
        norma_conteo: celda_I_30 if isinstance(celda_I_30, (int, float)) else None,
        norma_gomas: celda_gomas if isinstance(celda_gomas, (int, float)) else None,
        norma_MSEP: celda_MSEP if isinstance(celda_MSEP, (int, float)) else None,
        norma_conductividad: celda_conductividad if isinstance(celda_conductividad, (int, float)) else None,

        norma_acidez: celda_acidez if isinstance(celda_acidez, (int, float)) else None,
        norma_dest: celda_pto_inicial if isinstance(celda_pto_inicial, (int, float)) else None,
        norma_pto_inf: celda_pto_inf if isinstance(celda_pto_inf, (int, float)) else None,
        norma_densidad: celda_densidad if isinstance(celda_densidad, (int, float)) else None,
        norma_congelacion: celda_congelacion if isinstance(celda_congelacion, (int, float)) else None,
        norma_viscosidad: celda_viscosidad if isinstance(celda_viscosidad, (int, float)) else None,
        norma_calor: celda_calor if isinstance(celda_calor, (int, float)) else None,
        norma_pto_hum: celda_pto_hum if isinstance(celda_pto_hum, (int, float)) else None,
        norma_azufre: celda_azufre if isinstance(celda_azufre, (int, float)) else None,
        norma_aromaticos: celda_aromaticos if isinstance(celda_aromaticos, (int, float)) else None,
        norma_naftalenos: celda_naftalenos if isinstance(celda_naftalenos, (int, float)) else None,
        norma_mercaptanes: celda_mercaptanes if isinstance(celda_mercaptanes, (int, float)) else None,
    }

    plantilla, normas_acreditadas = acreditacion_y_plantilla(df_alcance, fecha_hoy, valores_por_norma)

    with st.expander("🔎 Debug (opcional)"):
        st.write("Plantilla elegida:", plantilla.name)
        st.write("Normas acreditadas detectadas:", normas_acreditadas)

    # =========================
    # Formateos (como tu script)
    # =========================
    celda_acidez_2 = fmt_num(celda_acidez, ".3f")
    celda_pto_inicial_2 = fmt_num(celda_pto_inicial, ".1f")
    celda_dest_10_2 = fmt_num(celda_dest_10, ".1f")
    celda_dest_50_2 = fmt_num(celda_dest_50, ".1f")
    celda_dest_90_2 = fmt_num(celda_dest_90, ".1f")
    celda_residuo_2 = fmt_num(celda_residuo, ".1f")
    celda_pto_final_2 = fmt_num(celda_pto_final, ".1f")
    celda_perdidas_2 = fmt_num(celda_perdidas, ".1f", default="0,0")
    celda_pto_inf_2 = fmt_num(celda_pto_inf, ".1f")
    celda_densidad_2 = fmt_num(densidad, ".1f")
    celda_congelacion_2 = fmt_num(celda_congelacion, ".1f")
    celda_viscosidad_2 = fmt_num(celda_viscosidad, ".3f")
    celda_calor_2 = fmt_num(celda_calor, ".3f")
    celda_pto_hum_2 = fmt_num(celda_pto_hum, ".1f")
    celda_JFTOT_2 = fmt_num(celda_JFTOT, ".1f")
    celda_antiestatico_2 = fmt_num(celda_antiestatico, ".2f")
    celda_antioxidante_2 = fmt_num(celda_antioxidante, ".1f")
    celda_aromaticos_2 = fmt_num(celda_aromaticos, ".1f")
    celda_naftalenos_2 = fmt_num(celda_naftalenos, ".2f")
    celda_mercaptanes_2 = fmt_num(celda_mercaptanes, ".4f")

    # Azufre con lógica de decimales (igual a tu script, simplificada)
    if isinstance(celda_azufre, (float, int)):
        x = float(celda_azufre)
        if x < 0.00001:  celda_azufre_2 = fmt_num(x, ".8f")
        elif x < 0.0001: celda_azufre_2 = fmt_num(x, ".7f")
        elif x < 0.001:  celda_azufre_2 = fmt_num(x, ".6f")
        elif x < 0.01:   celda_azufre_2 = fmt_num(x, ".5f")
        elif x < 0.1:    celda_azufre_2 = fmt_num(x, ".4f")
        elif x < 1:      celda_azufre_2 = fmt_num(x, ".3f")
        else:            celda_azufre_2 = fmt_num(x, ".2f")
    else:
        celda_azufre_2 = str(celda_azufre) if celda_azufre is not None else "-----"

    # composición
    celda_no_hidrop_2 = fmt_num(celda_no_hidrop, ".1f")
    celda_hidrop_2 = fmt_num(celda_hidrop, ".1f")
    celda_sintetico_2 = fmt_num(celda_sintetico, ".1f")
    celda_copros_2 = fmt_num(celda_copros, ".1f")

    # particulado
    if celda_particulado is None or isinstance(celda_particulado, str):
        celda_cont_part = "-----"
    else:
        celda_cont_part = (
            f"{fmt_num(celda_particulado, '.2f')}\n"
            f"(Vol. Filtrado: {fmt_num(celda_vol_filt, '.2f')} L)"
        )

    # Conteos "CP / ISO"
    def cont(cp, iso):
        if iso is None or cp is None or isinstance(iso, str) or isinstance(cp, str):
            return "-----"
        return f"{cp} / {iso:02d}"

    celda_cont_I_4 = cont(celda_CP_4, celda_I_4)
    celda_cont_I_6 = cont(celda_CP_6, celda_I_6)
    celda_cont_I_14 = cont(celda_CP_14, celda_I_14)
    celda_cont_I_21 = cont(celda_CP_21, celda_I_21)
    celda_cont_I_25 = cont(celda_CP_25, celda_I_25)
    celda_cont_I_30 = cont(celda_CP_30, celda_I_30)

    # Norma conteo fallback
    norma_conteo_2 = norma_conteo or "IP_565"

    # Observaciones: fecha MSEP
    obs = ""
    rep = ""
    obs_raw = safe_value(df_a, 1, 4, "Observaciones", default=None)
    if obs_raw:
        m = re.search(r"MSEP.*?(\d{1,2}/\d{1,2}/\d{2,4})", str(obs_raw), re.IGNORECASE)
        if m:
            fecha_msep = m.group(1)
            obs = f"(4) La conductividad eléctrica y el índice de separación (MSEP) fueron analizados sobre muestra extraída el {fecha_msep}."
            rep = "(4)"

    # Aromáticos FIA / HPLC
    esp_arom, tot = "", ""
    if norma_aromaticos == "ASTM_D_1319":
        esp_arom, tot = "25,0", "(% vol.)"
    elif norma_aromaticos == "ASTM_D_6379":
        esp_arom, tot = "26,5", "totales (% vol.)"

    # =========================
    # Merge fields y condicionales
    # =========================
    # Nombre propuesto de salida (en cloud lo usa para download)
    tanque_short = str(celda_tanque)[3:] if isinstance(celda_tanque, str) and len(celda_tanque) >= 4 else str(celda_tanque)
    nombre_archivo_nuevo = f"{celda_lims} TK {tanque_short} JET A1".strip().replace("  ", " ")

    datos_fusion = {
        "informe": str(celda_lims),
        "tanque": tanque_short,
        "fecha_aprob": str(celda_fecha_aprob).replace("-", "/"),
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
        "n_conteo": (norma_conteo_2 or "").replace("_", " "),

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

    extras = {}

    # Campo condicional pto_hum
    if isinstance(celda_pto_hum, (float, int)):
        extras["pto_hum_s"] = celda_pto_hum_2 if float(celda_pto_hum) >= 25 else "-----"
        extras["pto_hum_i"] = celda_pto_hum_2 if float(celda_pto_hum) < 25 else "-----"
    else:
        extras["pto_hum_s"] = "-----"
        extras["pto_hum_i"] = "-----"

    # Marcas "(*)" cuando NO acreditada (solo si plantilla OUA)
    if plantilla.name == "JET_A1.docx":
        def flag(norma, key):
            extras[key] = "" if (norma in normas_acreditadas) else "(*)"

        flag(norma_color, "col")
        flag(norma_conteo, "conteo")
        flag(norma_acidez, "acid")
        flag(norma_aromaticos, "arom")
        flag(norma_azufre, "azuf")
        flag(norma_mercaptanes, "merc")
        flag(norma_doctor, "doc")
        flag(norma_dest, "dest")
        flag(norma_pto_inf, "PI")
        flag(norma_densidad, "dens")
        flag(norma_congelacion, "cong")
        flag(norma_viscosidad, "vis")
        flag(norma_calor, "cal")
        flag(norma_pto_hum, "PH")
        flag(norma_naftalenos, "naft")
        flag(norma_corrosion, "corr")
        flag(norma_gomas, "gom")
        flag(norma_MSEP, "MS")
        flag(norma_conductividad, "cond")

    # =========================
    # Botón generar + descarga
    # =========================
    st.subheader("📄 Generar informe")

    if st.button("✅ Generar Word", type="primary"):
        try:
            doc_bytes = generar_docx(plantilla, datos_fusion, extras_merge=extras)

            st.success("Informe generado.")
            st.download_button(
                label="⬇️ Descargar informe Word",
                data=doc_bytes,
                file_name=f"{nombre_archivo_nuevo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as e:
            st.error(f"Error generando el Word: {e}")

else:
    st.info("Subí el CSV del LIMS para comenzar.")

