# ============================================================
# WELLNESS HUB - EJECUTABLE LOCAL
# Generado desde el notebook Colab original
# Ejecuta: python wellness_hub_ejecutable.py
# Opcional: python wellness_hub_ejecutable.py "ruta/al/Wellness Export morning.xlsx"
# ============================================================

import os
import sys
import subprocess

def ensure_package(import_name, pip_name=None):
    try:
        __import__(import_name)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pip_name or import_name])

for import_name, pip_name in [
    ("pandas", "pandas"),
    ("numpy", "numpy"),
    ("plotly", "plotly"),
    ("dash", "dash"),
    ("matplotlib", "matplotlib"),
    ("fpdf", "fpdf2"),
    ("openpyxl", "openpyxl"),
]:
    ensure_package(import_name, pip_name)

def seleccionar_excel():
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        return sys.argv[1]
    env_path = os.environ.get("WELLNESS_EXCEL_PATH")
    if env_path and os.path.exists(env_path):
        return env_path
    local_path = os.path.join(os.getcwd(), "Wellness Export morning.xlsx")
    if os.path.exists(local_path):
        return local_path
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        selected = filedialog.askopenfilename(
            title="Selecciona Wellness Export morning.xlsx",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if selected:
            return selected
    except Exception:
        pass
    raise FileNotFoundError(
        "No encuentro el Excel. Pon 'Wellness Export morning.xlsx' junto al .py "
        "o ejecuta: python wellness_hub_ejecutable.py RUTA_DEL_EXCEL.xlsx"
    )

EXCEL_PATH = seleccionar_excel()

#Importación de librerías
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from datetime import timedelta
import warnings
import random
import string


warnings.simplefilter(action='ignore', category=FutureWarning)


# %%
# Carga local del Excel
ruta = EXCEL_PATH
df = pd.read_excel(ruta)

# %%

cols_alert = df.columns[[8, 9]]

# Crear columna alert
df['alert'] = df[cols_alert].bfill(axis=1).iloc[:, 0]

# Eliminar columnas originales
df.drop(columns=cols_alert, inplace=True)


# %%

column_mapping = {
    'Fecha': 'date',
    'Jugador': 'player',
    '¿Qué tal has dormido?': 'sleep_quality',
    '¿Cómo de fatigado te sientes?': 'fatigue_pre',
    '¿Tienes dolor muscular o de otro tipo?': 'muscle_pain_pre',
    '¿Te sientes estresado?': 'stress_pre',
    '¿Te apetece ir a entrenar?': 'training_desire',
    '¿Cómo te encuentras animicamente?': 'mood_pre',
    '¿Cómo de fatigado te sientes? (POST)': 'fatigue_post',
    '¿Tienes dolor muscular o de otro tipo? (POST)': 'muscle_pain_post',
    '¿Te sientes estresado? (POST)': 'stress_post',
    '¿Cómo te encuentras anímicamente? (POST)': 'mood_post'
}

df.rename(columns=column_mapping, inplace=True)


# %%

df['player_id'] = pd.factorize(df['player'])[0]


# %%

df['date'] = pd.to_datetime(df['date'])
df['day_of_week'] = df['date'].dt.dayofweek

print("DataFrame after adding 'day_of_week' column (first 5 rows):")

print("\nDataFrame info after adding 'day_of_week' column:")

# %%

negative_metrics = [
    'fatigue_pre',
    'muscle_pain_pre',
    'stress_pre',
    'fatigue_post',
    'muscle_pain_post',
    'stress_post'
]

for col in negative_metrics:
    min_val = df[col].min()
    max_val = df[col].max()
    # Ensure min_val and max_val are not NaN to avoid issues with transformation
    if pd.notna(min_val) and pd.notna(max_val):
        df[col] = (max_val + min_val) - df[col]


# %%

#ordenamos el DataFrame

df = df.sort_values(by=['player_id', 'date'])

wellness_cols = df.select_dtypes(include=np.number).columns.drop(['player_id', 'day_of_week'])

for col in wellness_cols:
    # ffill busca el último valor válido hacia atrás (el más cercano en el pasado)
    df[col] = df.groupby('player_id')[col].ffill()
    # bfill busca el primer valor válido hacia adelante (el más cercano en el futuro)
    # por si el jugador no tiene datos previos en su primer registro
    df[col] = df.groupby('player_id')[col].bfill()

# Si después de eso sigue habiendo NaNs (porque un jugador nunca ha tenido datos en esa columna)
# podemos usar la mediana global
for col in wellness_cols:
    df[col] = df[col].fillna(df[col].median())




# %%


#Hay que tener cuidado con esta celda si la ejecutas dos veces se estandariza mal.
wellness_metrics = df.select_dtypes(include=np.number).columns.drop(['player_id', 'day_of_week',"fatigue_pre","fatigue_post","muscle_pain_pre","muscle_pain_post"])

for col in wellness_metrics:
    min_val = 1
    max_val = 4

    # Esta es la línea que detecta valores mayores a 4 y los sustituye

    df[col] = df[col].apply(lambda x: max_val if x > max_val else x)
    if max_val == min_val:
        df[col] = 0.0 # Handle cases where all values are the same to avoid division by zero
    else:
        df[col] = ((df[col] - min_val) / (max_val - min_val)) * 10

    print(f"\nDescriptive statistics for scaled column: {col}")
    print(df[col].describe())



# %%

# Calculate wellness_score as the mean of all standardized wellness metrics
# First, identify the wellness metrics columns again, excluding identifiers and date.
wellness_metrics_for_score = df.select_dtypes(include=np.number).columns.drop(['player_id', 'day_of_week'])

# Calculate the mean across these columns for each row
df['wellness_score'] = df[wellness_metrics_for_score].mean(axis=1)


# %%

negative_metrics = [
    'fatigue_pre',
    'muscle_pain_pre',
    'stress_pre',
    'fatigue_post',
    'muscle_pain_post',
    'stress_post'
]

for col in negative_metrics:
    min_val = df[col].min()
    max_val = df[col].max()
    # Ensure min_val and max_val are not NaN to avoid issues with transformation
    if pd.notna(min_val) and pd.notna(max_val):
        df[col] = (max_val + min_val) - df[col]


# %%

# ============================================================
# WELLNESS HUB - WEB PROFESIONAL + PDF COMPLETO
# Ejecutar despues de tener cargado el df principal
# ============================================================

import sys, subprocess, threading, warnings, os, tempfile, random
from datetime import timedelta
warnings.filterwarnings("ignore")

try:
    import dash
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "dash"])
    import dash

try:
    from fpdf import FPDF
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "fpdf2"])
    from fpdf import FPDF

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from dash import Dash, dcc, html, Input, Output, State, no_update, ctx
import webbrowser
import time

COLOR_ROJO = "#E31F3F"
COLOR_AZUL = "#163A6F"
COLOR_AZUL_OSCURO = "#071D3A"
COLOR_ORO = "#C5A365"
COLOR_VERDE = "#2D8C3C"
COLOR_FONDO = "#F4F6FA"
COLOR_TEXTO = "#172033"

URL_LOGO_CLUB = "/assets/club-logo.svg"

traduccion_metricas = {
    'wellness_score': 'Wellness',
    'fatigue_pre': 'Fatiga Pre',
    'performance_pre': 'Rendimiento Pre',
    'muscle_pain_pre': 'Dolor Muscular Pre',
    'stress_pre': 'Estres Pre',
    'sleep_length': 'Horas Sueno',
    'sleep_quality': 'Sueno',
    'fatigue_post': 'Fatiga Post',
    'performance_post': 'Rendimiento Post',
    'muscle_pain_post': 'Dolor Muscular Post',
    'stress_post': 'Estres Post',
    'rpe_post': 'Esfuerzo RPE',
    'session_duration': 'Duracion'
}

dias_esp = {
    'Monday': 'Lunes',
    'Tuesday': 'Martes',
    'Wednesday': 'Miercoles',
    'Thursday': 'Jueves',
    'Friday': 'Viernes',
    'Saturday': 'Sabado',
    'Sunday': 'Domingo'
}

# ------------------------------------------------------------
# DF PRINCIPAL
# ------------------------------------------------------------

df_app = df.copy()
df_app["date"] = pd.to_datetime(df_app["date"])
df_app["day_of_week"] = df_app["date"].dt.dayofweek
df_app["year"] = df_app["date"].dt.year

wellness_metrics_raw = (
    df_app
    .select_dtypes(include=[np.number])
    .columns
    .drop(["player_id", "day_of_week", "year"], errors="ignore")
    .tolist()
)

metric_options = [
    {"label": traduccion_metricas.get(m, m), "value": m}
    for m in sorted(wellness_metrics_raw, key=lambda x: traduccion_metricas.get(x, x))
]

players_options = [{"label": "Todos - Media Equipo", "value": "TODOS"}] + [
    {"label": p, "value": p}
    for p in sorted(df_app["player"].dropna().unique().tolist())
]

players_informe_options = [
    {"label": p, "value": p}
    for p in sorted(df_app["player"].dropna().unique().tolist())
]

jugador_informe_inicial = players_informe_options[0]["value"] if players_informe_options else None

fecha_final = df_app["date"].max()
fecha_inicio_semana = fecha_final - timedelta(days=6)
media_global = round(df_app["wellness_score"].mean(), 2)
media_ultima_semana = round(df_app[df_app["date"] >= fecha_inicio_semana]["wellness_score"].mean(), 2)

# ------------------------------------------------------------
# CONTROL DIARIO INDEPENDIENTE
# ------------------------------------------------------------

ruta_diario = EXCEL_PATH
df_diario = pd.read_excel(ruta_diario)

dicc_columnas_diario = {
    "Fecha": "fecha",
    "Jugador": "jugador",
    "¿Qué tal has dormido?": "sueno",
    "¿Cómo de fatigado te sientes?": "fatiga_pre",
    "¿Tienes dolor muscular o de otro tipo?": "dolor_pre",
    "¿Te sientes estresado?": "estres_pre",
    "¿Te apetece ir a entrenar?": "ganas_entrenar",
    "¿Cómo te encuentras animicamente?": "animo_pre",
    "¿Ha sucedido algo fuera de lo normal que te gustaría comentar?": "comentarios_ext",
    "¿Sientes que hay factores externos al club que te estén afectando últimamente?": "factores_ext",
    "¿Cómo de fatigado te sientes? (POST)": "fatiga_post",
    "¿Tienes dolor muscular o de otro tipo? (POST)": "dolor_post",
    "¿Te sientes estresado? (POST)": "estres_post",
    "¿Cómo te encuentras anímicamente? (POST)": "animo_post"
}

df_diario = df_diario.rename(columns=dicc_columnas_diario)
df_diario = df_diario.drop(columns=["comentarios_ext", "factores_ext"], errors="ignore")
df_diario = df_diario.dropna(subset=["jugador", "fecha"]).dropna(how="all")
df_diario["fecha"] = pd.to_datetime(df_diario["fecha"])
df_diario = df_diario.sort_values(["jugador", "fecha"])

cols_numericas_diario = [
    "sueno", "fatiga_pre", "dolor_pre", "estres_pre", "ganas_entrenar",
    "animo_pre", "fatiga_post", "dolor_post", "estres_post", "animo_post"
]

for col in cols_numericas_diario:
    df_diario[col] = pd.to_numeric(df_diario[col], errors="coerce")

df_diario.loc[df_diario["sueno"] > 4, "sueno"] = np.nan
df_diario.loc[df_diario["dolor_post"] < 1, "dolor_post"] = np.nan

def moda_diario(x):
    m = x.mode()
    return m.iloc[0] if not m.empty else np.nan

for col in cols_numericas_diario:
    df_diario[col] = df_diario.groupby("jugador")[col].transform(lambda x: x.fillna(moda_diario(x)))

for col in cols_numericas_diario:
    moda_global_col = df_diario[col].mode()
    if not moda_global_col.empty:
        df_diario[col] = df_diario[col].fillna(moda_global_col.iloc[0])

for col in cols_numericas_diario:
    v_min = df_diario[col].min()
    v_max = df_diario[col].max()
    if pd.notna(v_min) and pd.notna(v_max) and v_max - v_min != 0:
        df_diario[col] = (df_diario[col] - v_min) / (v_max - v_min)
    else:
        df_diario[col] = 0.0

for col in ["fatiga_pre", "estres_pre", "dolor_pre", "fatiga_post", "estres_post", "dolor_post"]:
    df_diario[col] = 1 - df_diario[col]

df_diario["fis_pre"] = df_diario[["fatiga_pre", "dolor_pre"]].mean(axis=1)
df_diario["fis_post"] = df_diario[["fatiga_post", "dolor_post"]].mean(axis=1)
df_diario["men_pre"] = df_diario[["estres_pre", "animo_pre"]].mean(axis=1)
df_diario["men_post"] = df_diario[["estres_post", "animo_post"]].mean(axis=1)
df_diario["wellness"] = df_diario[["sueno", "fis_pre", "men_pre", "ganas_entrenar"]].mean(axis=1)

fechas_diario = sorted(df_diario["fecha"].dt.date.unique().tolist())
fecha_diario_inicial = fechas_diario[-1]

# ------------------------------------------------------------
# PREPROCESO PDF COMPLETO UNA SOLA VEZ
# ------------------------------------------------------------

column_mapping_inverso = {
    "date": "Fecha",
    "player": "Jugador",
    "sleep_quality": "¿Qué tal has dormido?",
    "fatigue_pre": "¿Cómo de fatigado te sientes?",
    "muscle_pain_pre": "¿Tienes dolor muscular o de otro tipo?",
    "stress_pre": "¿Te sientes estresado?",
    "training_desire": "¿Te apetece ir a entrenar?",
    "mood_pre": "¿Cómo te encuentras animicamente?",
    "fatigue_post": "¿Cómo de fatigado te sientes? (POST)",
    "muscle_pain_post": "¿Tienes dolor muscular o de otro tipo? (POST)",
    "stress_post": "¿Te sientes estresado? (POST)",
    "mood_post": "¿Cómo te encuentras anímicamente? (POST)"
}

cols_a_renombrar = {k: v for k, v in column_mapping_inverso.items() if k in df_app.columns}
df_pdf_base = df_app[list(cols_a_renombrar.keys())].copy().rename(columns=cols_a_renombrar)

variables_pdf = [
    "¿Qué tal has dormido?",
    "¿Cómo de fatigado te sientes?",
    "¿Tienes dolor muscular o de otro tipo?",
    "¿Te sientes estresado?",
    "¿Te apetece ir a entrenar?",
    "¿Cómo te encuentras animicamente?",
    "¿Cómo de fatigado te sientes? (POST)",
    "¿Tienes dolor muscular o de otro tipo? (POST)",
    "¿Te sientes estresado? (POST)",
    "¿Cómo te encuentras anímicamente? (POST)"
]

for var in variables_pdf:
    if var in df_pdf_base.columns:
        df_pdf_base[var] = pd.to_numeric(df_pdf_base[var], errors="coerce")

df_pdf_proc = df_pdf_base.copy()

for var in variables_pdf:
    if var not in df_pdf_proc.columns:
        continue
    modas_por_jugador = df_pdf_proc.groupby("Jugador")[var].transform(
        lambda x: x.mode().iloc[0] if not x.mode().empty else np.nan
    )
    df_pdf_proc[var] = df_pdf_proc[var].fillna(modas_por_jugador)
    moda_global = df_pdf_proc[var].mode()
    if not moda_global.empty:
        df_pdf_proc[var] = df_pdf_proc[var].fillna(moda_global.iloc[0])

for var in variables_pdf:
    if var not in df_pdf_proc.columns:
        continue
    min_val = df_pdf_proc[var].min()
    max_val = df_pdf_proc[var].max()
    if pd.notna(min_val) and pd.notna(max_val) and max_val - min_val != 0:
        df_pdf_proc[var] = (df_pdf_proc[var] - min_val) / (max_val - min_val)

variables_negativas_pdf = [
    "¿Cómo de fatigado te sientes?",
    "¿Tienes dolor muscular o de otro tipo?",
    "¿Te sientes estresado?",
    "¿Cómo de fatigado te sientes? (POST)",
    "¿Tienes dolor muscular o de otro tipo? (POST)",
    "¿Te sientes estresado? (POST)"
]

for var in variables_negativas_pdf:
    if var in df_pdf_proc.columns:
        df_pdf_proc[var] = 1 - df_pdf_proc[var]

vars_cansancio_pre = [v for v in [
    "¿Cómo de fatigado te sientes?",
    "¿Tienes dolor muscular o de otro tipo?",
    "¿Te sientes estresado?",
    "¿Cómo te encuentras animicamente?"
] if v in df_pdf_proc.columns]

vars_cansancio_post = [v for v in [
    "¿Cómo de fatigado te sientes? (POST)",
    "¿Tienes dolor muscular o de otro tipo? (POST)",
    "¿Te sientes estresado? (POST)",
    "¿Cómo te encuentras anímicamente? (POST)"
] if v in df_pdf_proc.columns]

df_pdf_proc["Bienestar_PRE"] = df_pdf_proc[vars_cansancio_pre].mean(axis=1)
df_pdf_proc["Bienestar_POST"] = df_pdf_proc[vars_cansancio_post].mean(axis=1)
df_pdf_proc["Impacto_Entrenamiento"] = df_pdf_proc["Bienestar_POST"] - df_pdf_proc["Bienestar_PRE"]

df_pdf_proc["tipo de impacto"] = np.where(
    df_pdf_proc["Impacto_Entrenamiento"] > 0, "Positivo",
    np.where(df_pdf_proc["Impacto_Entrenamiento"] < 0, "Negativo", "Neutral")
)

magnitud_pdf = df_pdf_proc["Impacto_Entrenamiento"].abs()
df_pdf_proc["nivel de impacto"] = np.where(
    magnitud_pdf <= 0.05, "Bajo",
    np.where(magnitud_pdf <= 0.15, "Moderado", "Alto")
)

dif_fatiga = df_pdf_proc["¿Cómo de fatigado te sientes? (POST)"] - df_pdf_proc["¿Cómo de fatigado te sientes?"]
dif_dolor = df_pdf_proc["¿Tienes dolor muscular o de otro tipo? (POST)"] - df_pdf_proc["¿Tienes dolor muscular o de otro tipo?"]
dif_estres = df_pdf_proc["¿Te sientes estresado? (POST)"] - df_pdf_proc["¿Te sientes estresado?"]
dif_animo = df_pdf_proc["¿Cómo te encuentras anímicamente? (POST)"] - df_pdf_proc["¿Cómo te encuentras animicamente?"]

df_pdf_proc["suma_fisica"] = (dif_fatiga + dif_dolor).abs()
df_pdf_proc["suma_mental"] = (dif_estres + dif_animo).abs()

df_pdf_proc["foco del impacto"] = np.where(
    df_pdf_proc["suma_fisica"] > df_pdf_proc["suma_mental"] + 0.1, "Fisico",
    np.where(df_pdf_proc["suma_mental"] > df_pdf_proc["suma_fisica"] + 0.1, "Mental", "Ambos")
)

vars_existentes_pdf = [v for v in variables_pdf if v in df_pdf_proc.columns]
df_pdf_proc["Wellness_General"] = df_pdf_proc[vars_existentes_pdf].mean(axis=1)
df_pdf_proc["Fecha"] = pd.to_datetime(df_pdf_proc["Fecha"])

df_pdf_proc["Estado fisico PRE"] = df_pdf_proc[
    ["¿Cómo de fatigado te sientes?", "¿Tienes dolor muscular o de otro tipo?"]
].mean(axis=1)
df_pdf_proc["Estado fisico POST"] = df_pdf_proc[
    ["¿Cómo de fatigado te sientes? (POST)", "¿Tienes dolor muscular o de otro tipo? (POST)"]
].mean(axis=1)
df_pdf_proc["Estado mental PRE"] = df_pdf_proc[
    ["¿Te sientes estresado?", "¿Cómo te encuentras animicamente?"]
].mean(axis=1)
df_pdf_proc["Estado mental POST"] = df_pdf_proc[
    ["¿Te sientes estresado? (POST)", "¿Cómo te encuentras anímicamente? (POST)"]
].mean(axis=1)

medias_equipo_pdf = df_pdf_proc.groupby("Fecha")[["Wellness_General", "Bienestar_PRE", "Bienestar_POST"]].mean().reset_index()
medias_equipo_pdf.columns = ["Fecha", "Media_General_Equipo", "Media_PRE_Equipo", "Media_POST_Equipo"]

# ------------------------------------------------------------
# FUNCIONES PDF
# ------------------------------------------------------------

def guardar_radar_pdf(df_jugador, df_equipo, nombre_jugador, ruta):
    variables_radar = {
        "Sueno": "¿Qué tal has dormido?",
        "Ganas": "¿Te apetece ir a entrenar?",
        "Fisico PRE": "Estado fisico PRE",
        "Mental PRE": "Estado mental PRE",
        "Wellness": "Wellness_General"
    }
    etiquetas = list(variables_radar.keys())
    val_jugador = [df_jugador[col].mean() for col in variables_radar.values()]
    val_equipo = [df_equipo[col].mean() for col in variables_radar.values()]
    angulos = np.linspace(0, 2 * np.pi, len(etiquetas), endpoint=False).tolist()
    val_jugador += val_jugador[:1]
    val_equipo += val_equipo[:1]
    angulos += angulos[:1]

    fig, ax = plt.subplots(figsize=(7, 7), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    plt.xticks(angulos[:-1], etiquetas, color="black", size=10)
    ax.plot(angulos, val_equipo, color="gray", linewidth=2, linestyle="--", label="Media Equipo")
    ax.fill(angulos, val_equipo, color="gray", alpha=0.1)
    ax.plot(angulos, val_jugador, color="#C4122D", linewidth=3, label=nombre_jugador)
    ax.fill(angulos, val_jugador, color="#C4122D", alpha=0.2)
    ax.set_ylim(0, 1)
    ax.set_yticks([0.2, 0.4, 0.6, 0.8, 1.0])
    ax.set_yticklabels(["0.2", "0.4", "0.6", "0.8", "1.0"], color="gray", size=8)
    plt.legend(loc="upper right", bbox_to_anchor=(1.2, 1.1))
    plt.title(f"Perfil: {nombre_jugador} vs Equipo", pad=20, weight="bold")
    plt.savefig(ruta, bbox_inches="tight", dpi=150)
    plt.close()

def guardar_quesitos_pdf(df_sub, ruta):
    fig, axs = plt.subplots(2, 2, figsize=(12, 11))
    rojo_alto, rojo_medio, rojo_bajo = "#7f1d1d", "#c4122d", "#f87171"
    verde_alto, verde_medio, verde_bajo = "#064e3b", "#10b981", "#a7f3d0"
    gris_oscuro, gris_claro = "#4b5563", "#d1d5db"

    df_sub = df_sub.copy()
    df_sub["status_color"] = df_sub["tipo de impacto"] + " " + df_sub["nivel de impacto"]

    orden = [
        "Positivo Alto", "Positivo Moderado", "Positivo Bajo",
        "Neutral Bajo", "Neutral Moderado", "Neutral Alto",
        "Negativo Bajo", "Negativo Moderado", "Negativo Alto"
    ]
    counts = df_sub["status_color"].value_counts()
    labels = [x for x in orden if x in counts.index]
    valores = [counts[x] for x in labels]
    colores_map = {
        "Negativo Alto": rojo_alto, "Negativo Moderado": rojo_medio, "Negativo Bajo": rojo_bajo,
        "Positivo Alto": verde_alto, "Positivo Moderado": verde_medio, "Positivo Bajo": verde_bajo,
        "Neutral Bajo": gris_claro, "Neutral Moderado": gris_oscuro, "Neutral Alto": "#1f2937"
    }

    if valores:
        axs[0, 0].pie(
            valores,
            labels=labels,
            autopct="%1.1f%%",
            pctdistance=0.75,
            colors=[colores_map[x] for x in labels],
            wedgeprops={"width": 0.4, "edgecolor": "w"}
        )
    axs[0, 0].set_title("Balance e Intensidad", pad=20, weight="bold")

    foco_counts = df_sub["foco del impacto"].value_counts()
    if not foco_counts.empty:
        axs[0, 1].pie(
            foco_counts,
            labels=foco_counts.index,
            autopct="%1.1f%%",
            pctdistance=0.75,
            colors=[rojo_medio, gris_oscuro, gris_claro],
            wedgeprops={"width": 0.4, "edgecolor": "w"}
        )
    axs[0, 1].set_title("Motivo Global del Impacto", pad=20, weight="bold")

    df_pos = df_sub[df_sub["tipo de impacto"] == "Positivo"]
    if not df_pos.empty:
        pos_counts = df_pos["foco del impacto"].value_counts()
        axs[1, 0].pie(
            pos_counts,
            labels=pos_counts.index,
            autopct="%1.1f%%",
            pctdistance=0.75,
            colors=[rojo_medio, gris_oscuro, gris_claro],
            wedgeprops={"width": 0.4, "edgecolor": "w"}
        )
    axs[1, 0].set_title("Motivos: Impactos Positivos", pad=20, weight="bold")

    df_neg = df_sub[df_sub["tipo de impacto"] == "Negativo"]
    if not df_neg.empty:
        neg_counts = df_neg["foco del impacto"].value_counts()
        axs[1, 1].pie(
            neg_counts,
            labels=neg_counts.index,
            autopct="%1.1f%%",
            pctdistance=0.75,
            colors=[rojo_medio, gris_oscuro, gris_claro],
            wedgeprops={"width": 0.4, "edgecolor": "w"}
        )
    axs[1, 1].set_title("Motivos: Impactos Negativos", pad=20, weight="bold")

    plt.tight_layout()
    plt.savefig(ruta, dpi=150)
    plt.close()

def guardar_comparativa_pdf(df_sub, df_equipo, nombre_jugador, ruta):
    df_plot = pd.merge(
        df_sub[["Fecha", "Wellness_General"]],
        df_equipo[["Fecha", "Media_General_Equipo"]],
        on="Fecha",
        how="left"
    )
    plt.figure(figsize=(10, 5))
    plt.plot(df_plot["Fecha"], df_plot["Media_General_Equipo"], label="Media Equipo", color="#fbbf24", linewidth=2, alpha=0.6)
    plt.plot(df_plot["Fecha"], df_plot["Wellness_General"], label=nombre_jugador, color="#10b981", linewidth=3, marker="o")
    plt.title(f"Comparativa Wellness Diario: {nombre_jugador} vs Equipo", weight="bold")
    plt.ylim(0, 1.1)
    plt.grid(True, alpha=0.2)
    plt.legend()
    plt.tight_layout()
    plt.savefig(ruta, dpi=150)
    plt.close()

def guardar_evolucion_pdf(df_sub, nombre, ruta):
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
    ax1.plot(df_sub["Fecha"], df_sub["Estado fisico PRE"], label="Fisico PRE", color="gray", marker="o")
    ax1.plot(df_sub["Fecha"], df_sub["Estado fisico POST"], label="Fisico POST", color="#C4122D", marker="o")
    ax1.set_title(f"Evolucion Fisica: {nombre}")
    ax1.set_ylim(0, 1.1)
    ax1.legend()

    ax2.plot(df_sub["Fecha"], df_sub["Estado mental PRE"], label="Mental PRE", color="gray", marker="s")
    ax2.plot(df_sub["Fecha"], df_sub["Estado mental POST"], label="Mental POST", color="#4B5563", marker="s")
    ax2.set_title(f"Evolucion Mental: {nombre}")
    ax2.set_ylim(0, 1.1)
    ax2.legend()

    plt.tight_layout()
    plt.savefig(ruta, dpi=150)
    plt.close()

class InformeWellness(FPDF):
    def header(self):
        self.set_fill_color(196, 18, 45)
        self.rect(0, 0, 210, 30, "F")
        self.set_font("Arial", "B", 16)
        self.set_text_color(255, 255, 255)
        self.cell(0, 10, "INFORME DE RENDIMIENTO Y WELLNESS", 0, 1, "C")
        self.ln(5)

def crear_pdf_completo(jugador, fecha_inicio, fecha_fin):
    workdir = tempfile.mkdtemp()
    f_ini = pd.to_datetime(fecha_inicio)
    f_fin = pd.to_datetime(fecha_fin)

    df_sub = df_pdf_proc[
        (df_pdf_proc["Jugador"] == jugador) &
        (df_pdf_proc["Fecha"] >= f_ini) &
        (df_pdf_proc["Fecha"] <= f_fin)
    ].sort_values("Fecha")

    if df_sub.empty:
        return None, "No hay datos para ese jugador en el periodo seleccionado."

    ruta_quesitos = os.path.join(workdir, "quesitos.png")
    ruta_evolucion = os.path.join(workdir, "evolucion.png")
    ruta_comp = os.path.join(workdir, "comp_equipo.png")
    ruta_radar = os.path.join(workdir, "radar.png")

    guardar_quesitos_pdf(df_sub, ruta_quesitos)
    guardar_evolucion_pdf(df_sub, jugador, ruta_evolucion)
    guardar_comparativa_pdf(df_sub, medias_equipo_pdf, jugador, ruta_comp)
    guardar_radar_pdf(df_sub, df_pdf_proc, jugador, ruta_radar)

    num_registros = len(df_sub)
    media_wellness = df_sub["Wellness_General"].mean()

    pdf = InformeWellness()

    pdf.add_page()
    pdf.set_y(35)
    pdf.set_font("helvetica", "B", 14)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, f"DATOS DEL JUGADOR: {jugador}", 0, 1)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 7, f"Periodo seleccionado: {f_ini.date()} al {f_fin.date()}", 0, 1)
    pdf.cell(0, 7, f"Sesiones registradas: {num_registros}", 0, 1)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 7, f"Media Wellness Global en el periodo: {media_wellness:.2f} / 1.00", 0, 1)
    pdf.set_draw_color(196, 18, 45)
    pdf.line(10, pdf.get_y() + 2, 200, pdf.get_y() + 2)
    pdf.set_y(pdf.get_y() + 8)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(196, 18, 45)
    pdf.cell(0, 10, "1. ANALISIS DE IMPACTOS E INTENSIDAD", 0, 1)
    pdf.image(ruta_quesitos, x=10, y=pdf.get_y(), w=190)

    pdf.add_page()
    pdf.set_y(35)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(196, 18, 45)
    pdf.cell(0, 10, "2. EVOLUCION BIENESTAR FISICO Y MENTAL", 0, 1)
    pdf.image(ruta_evolucion, x=20, y=pdf.get_y(), w=170)
    pdf.set_y(185)
    pdf.cell(0, 10, "3. COMPARATIVA DIARIA VS MEDIA DEL EQUIPO", 0, 1)
    pdf.image(ruta_comp, x=20, y=195, w=170)

    pdf.add_page()
    pdf.set_y(35)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(196, 18, 45)
    pdf.cell(0, 10, "4. PERFIL COMPARATIVO FINAL RADAR", 0, 1, "C")
    pdf.image(ruta_radar, x=40, y=55, w=130)

    nombre_archivo = f"Informe_Wellness_{jugador.replace(' ', '_')}.pdf"
    ruta_pdf = os.path.join(workdir, nombre_archivo)
    pdf.output(ruta_pdf)
    return ruta_pdf, f"Informe generado para {jugador}."

# ------------------------------------------------------------
# FUNCIONES DASH
# ------------------------------------------------------------

def get_status_diario(jugador_name, fecha_sel):
    fecha_dt = pd.to_datetime(fecha_sel)
    dia_semana_sel = fecha_dt.dayofweek
    nombre_dia = dias_esp[fecha_dt.day_name()]

    df_hist = df_diario[(df_diario["jugador"] == jugador_name) & (df_diario["fecha"] <= fecha_dt)]
    df_hist_mismo_dia = df_hist[df_hist["fecha"].dt.dayofweek == dia_semana_sel].sort_values("fecha")
    df_dia = df_diario[(df_diario["jugador"] == jugador_name) & (df_diario["fecha"] == fecha_dt)]

    if df_dia.empty:
        return None

    registro_hoy = df_dia.iloc[0]
    media_hist = df_hist_mismo_dia["wellness"].mean()
    num_registros = len(df_hist_mismo_dia)
    desv = 0 if pd.isna(media_hist) or media_hist == 0 else ((registro_hoy["wellness"] - media_hist) / media_hist) * 100

    if desv > -10:
        status, color, accion = "APTO", COLOR_VERDE, "Entrenamiento normal"
    elif desv > -15:
        status, color, accion = "PRECAUCION", COLOR_ORO, "Reducir cargas"
    else:
        status, color, accion = "ALERTA", COLOR_ROJO, "Necesita descansar"

    return {
        "status": status,
        "color": color,
        "accion": accion,
        "desv": desv,
        "val_actual": registro_hoy["wellness"],
        "media_hist": media_hist,
        "num_registros": num_registros,
        "registro": registro_hoy,
        "nombre_dia": nombre_dia
    }

def kpi_card(label, value, color=COLOR_AZUL, sub=None):
    return html.Div(className="kpi-card", children=[
        html.Div(label, className="kpi-label"),
        html.Div(value, className="kpi-value", style={"color": color}),
        html.Div(sub or "", className="kpi-sub")
    ])

# ------------------------------------------------------------
# APP
# ------------------------------------------------------------

app = Dash(__name__)
app.title = "Wellness Hub"
app.config.suppress_callback_exceptions = True

app.index_string = f"""
<!DOCTYPE html>
<html>
<head>
    {{%metas%}}
    <title>{{%title%}}</title>
    {{%favicon%}}
    {{%css%}}
    <style>
        body {{
            margin: 0;
            background: {COLOR_FONDO};
            color: {COLOR_TEXTO};
            font-family: Inter, Segoe UI, Arial, sans-serif;
        }}
        .app-shell {{
            max-width: 1440px;
            margin: 0 auto;
            padding: 28px;
        }}
        .hero {{
            background: linear-gradient(135deg, {COLOR_AZUL_OSCURO} 0%, {COLOR_AZUL} 65%, #244E86 100%);
            color: white;
            border-radius: 22px;
            padding: 26px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 18px 40px rgba(7,29,58,0.22);
            border-bottom: 5px solid {COLOR_ORO};
        }}
        .hero-title {{
            font-size: 38px;
            font-weight: 900;
            margin: 0;
            letter-spacing: 1px;
            text-transform: uppercase;
        }}
        .hero-subtitle {{
            margin: 8px 0 0 0;
            opacity: 0.86;
            font-size: 13px;
            letter-spacing: 3px;
            text-transform: uppercase;
            font-weight: 700;
        }}
        .hero-logo {{
            height: 74px;
            background: white;
            border-radius: 18px;
            padding: 10px;
        }}
        .help-button {{
            background: white;
            color: {COLOR_AZUL};
            border: 0;
            border-radius: 999px;
            padding: 11px 16px;
            font-weight: 900;
            cursor: pointer;
            box-shadow: 0 10px 24px rgba(0,0,0,0.18);
            margin-right: 18px;
        }}
        .help-button:hover {{
            background: {COLOR_FONDO};
        }}
        .modal-backdrop {{
            position: fixed;
            inset: 0;
            background: rgba(7, 29, 58, 0.58);
            z-index: 9999;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }}
        .modal-card {{
            width: min(860px, 92vw);
            background: white;
            border-radius: 24px;
            padding: 28px;
            box-shadow: 0 28px 80px rgba(0,0,0,0.28);
            border-top: 6px solid {COLOR_ORO};
        }}
        .modal-title {{
            color: {COLOR_AZUL};
            font-size: 22px;
            font-weight: 900;
            margin: 0 0 16px 0;
            text-transform: uppercase;
        }}
        .manual-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 18px;
            font-size: 14px;
            line-height: 1.55;
            color: {COLOR_TEXTO};
        }}
        .manual-section-title {{
            color: {COLOR_ROJO};
            font-weight: 900;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 1px;
            margin-bottom: 6px;
        }}
        .close-button {{
            margin-top: 22px;
            background: {COLOR_AZUL};
            color: white;
            border: 0;
            border-radius: 14px;
            padding: 12px 18px;
            font-weight: 900;
            cursor: pointer;
        }}
        .top-kpis {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 18px;
            margin-top: 22px;
        }}
        .kpi-card {{
            background: rgba(255,255,255,0.94);
            border: 1px solid #E5EAF2;
            border-radius: 18px;
            padding: 18px;
            box-shadow: 0 12px 28px rgba(15,30,60,0.06);
        }}
        .kpi-label {{
            color: #6B7280;
            font-size: 11px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 1.4px;
            margin-bottom: 8px;
        }}
        .kpi-value {{
            font-size: 34px;
            font-weight: 900;
            line-height: 1;
        }}
        .kpi-sub {{
            font-size: 12px;
            color: #6B7280;
            margin-top: 8px;
            min-height: 15px;
        }}
        .main-grid {{
            display: grid;
            grid-template-columns: 360px 1fr;
            gap: 22px;
            margin-top: 22px;
            align-items: start;
        }}
        .panel {{
            background: white;
            border: 1px solid #E5EAF2;
            border-radius: 22px;
            padding: 20px;
            box-shadow: 0 14px 34px rgba(15,30,60,0.07);
        }}
        .panel-title {{
            font-size: 14px;
            font-weight: 900;
            color: {COLOR_AZUL};
            text-transform: uppercase;
            letter-spacing: 1.2px;
            margin-bottom: 16px;
            padding-bottom: 10px;
            border-bottom: 3px solid {COLOR_ORO};
        }}
        .control-label {{
            font-size: 11px;
            font-weight: 900;
            color: #596273;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin: 14px 0 7px 0;
        }}
        .tabs-wrap {{
            margin-top: 22px;
        }}
        .section-spacer {{
            height: 14px;
        }}
        .primary-button {{
            width: 100%;
            margin-top: 14px;
            background: {COLOR_ROJO};
            color: white;
            border: 0;
            border-radius: 14px;
            padding: 13px 16px;
            font-weight: 900;
            letter-spacing: 1px;
            text-transform: uppercase;
            cursor: pointer;
            box-shadow: 0 10px 22px rgba(227,31,63,0.22);
        }}
        .primary-button:hover {{
            background: {COLOR_AZUL};
        }}
        .status-text {{
            font-size: 12px;
            color: #596273;
            margin-top: 10px;
            font-weight: 700;
            min-height: 18px;
        }}
        .status-pill {{
            display: inline-flex;
            padding: 9px 13px;
            border-radius: 999px;
            color: white;
            font-weight: 900;
            font-size: 12px;
            letter-spacing: 0.8px;
        }}
        .alert-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        .alert-table th {{
            text-align: left;
            background: {COLOR_AZUL};
            color: white;
            padding: 12px;
        }}
        .alert-table td {{
            border-bottom: 1px solid #EEF1F6;
            padding: 12px;
        }}
    </style>
</head>
<body>
    {{%app_entry%}}
    <footer>
        {{%config%}}
        {{%scripts%}}
        {{%renderer%}}
    </footer>
</body>
</html>
"""

app.layout = html.Div(className="app-shell", children=[
    html.Div(className="hero", children=[
        html.Div([
            html.P("Gestion de disponibilidad y bienestar", className="hero-subtitle"),
            html.H1("Wellness Hub", className="hero-title")
        ]),
        html.Div(
            style={"display": "flex", "alignItems": "center"},
            children=[
                html.Button("Manual de uso", id="btn-manual", n_clicks=0, className="help-button"),
                html.Img(src=URL_LOGO_CLUB, className="hero-logo")
            ]
        )
    ]),

    html.Div(
        id="modal-manual",
        className="modal-backdrop",
        style={"display": "none"},
        children=[
            html.Div(className="modal-card", children=[
                html.H2("Manual de uso", className="modal-title"),
                html.Div(className="manual-grid", children=[
                    html.Div([
                        html.Div("Filtros de analisis", className="manual-section-title"),
                        html.P("Selecciona las metricas y jugadores que quieres visualizar en la evolucion principal. Esta zona solo afecta a la grafica superior interactiva."),
                        html.Div("Grafica principal", className="manual-section-title"),
                        html.P("Permite comparar la evolucion temporal de una o varias metricas. Puedes pasar el raton por la grafica, ampliar zonas y activar o desactivar series desde la leyenda.")
                    ]),
                    html.Div([
                        html.Div("Filtro periodo", className="manual-section-title"),
                        html.P("Selecciona un jugador concreto y un rango de fechas para revisar el resumen del periodo y generar el informe PDF completo."),
                        html.Div("Control diario", className="manual-section-title"),
                        html.P("Elige una fecha y un jugador para ver su disponibilidad diaria, comparativa con el equipo, perfil radar, evolucion historica y alertas detectadas.")
                    ])
                ]),
                html.Button("Cerrar", id="btn-cerrar-manual", n_clicks=0, className="close-button")
            ])
        ]
    ),

    html.Div(className="top-kpis", children=[
        kpi_card("Media Wellness Global", f"{media_global}", COLOR_AZUL),
        kpi_card("Ultima Semana", f"{media_ultima_semana}", COLOR_ROJO),
        kpi_card("Jugadores", f"{df_app['player'].nunique()}", COLOR_AZUL),
        kpi_card("Ultima Fecha", fecha_final.strftime("%d/%m/%Y"), COLOR_ORO)
    ]),

    html.Div(className="main-grid", children=[
        html.Div(className="panel", children=[
            html.Div("Filtros de analisis", className="panel-title"),

            html.Div("Metricas", className="control-label"),
            dcc.Dropdown(
                id="metricas",
                options=metric_options,
                value=["wellness_score"],
                multi=True,
                clearable=False
            ),

            html.Div("Jugadores", className="control-label"),
            dcc.Dropdown(
                id="jugadores",
                options=players_options,
                value=["TODOS"],
                multi=True,
                clearable=False
            ),

            html.Div("Filtro periodo", className="panel-title", style={"marginTop": "24px"}),

            html.Div("Jugador del informe", className="control-label"),
            dcc.Dropdown(
                id="jugador-informe",
                options=players_informe_options,
                value=jugador_informe_inicial,
                clearable=False
            ),

            html.Div("Periodo", className="control-label"),
            dcc.DatePickerRange(
                id="periodo",
                start_date=fecha_inicio_semana.date(),
                end_date=fecha_final.date(),
                display_format="DD/MM/YYYY"
            ),

            html.Button("Generar informe PDF", id="btn-informe", n_clicks=0, className="primary-button"),
            dcc.Download(id="download-informe"),
            dcc.Loading(
                id="loading-informe",
                type="circle",
                color=COLOR_ROJO,
                children=html.Div(id="estado-informe", className="status-text")
            ),


            html.Div(className="section-spacer"),

            html.Div("Control diario", className="panel-title"),

            html.Div("Fecha", className="control-label"),
            dcc.DatePickerSingle(
                id="fecha-diaria",
                date=fecha_diario_inicial,
                display_format="DD/MM/YYYY"
            ),

            html.Div("Jugador", className="control-label"),
            dcc.Dropdown(id="jugador-diario", clearable=False)
        ]),

        html.Div(children=[
            html.Div(className="panel", children=[
                html.Div("Evolucion principal", className="panel-title"),
                dcc.Loading(
                    type="circle",
                    children=dcc.Graph(
                        id="grafica-principal",
                        config={"displayModeBar": True, "responsive": True},
                        style={"height": "520px"}
                    )
                )
            ]),

            html.Div(className="tabs-wrap panel", children=[
                dcc.Tabs(id="tabs", value="periodo", children=[
                    dcc.Tab(label="Informe de periodo", value="periodo"),
                    dcc.Tab(label="Control diario", value="diario"),
                ]),
                dcc.Loading(type="circle", children=html.Div(id="contenido-tabs"))
            ])
        ])
    ])
])

# ------------------------------------------------------------
# CALLBACKS
# ------------------------------------------------------------

@app.callback(
    Output("modal-manual", "style"),
    Input("btn-manual", "n_clicks"),
    Input("btn-cerrar-manual", "n_clicks"),
    prevent_initial_call=True
)
def controlar_modal_manual(n_abrir, n_cerrar):
    if ctx.triggered_id == "btn-manual":
        return {"display": "flex"}
    return {"display": "none"}

@app.callback(
    Output("grafica-principal", "figure"),
    Input("metricas", "value"),
    Input("jugadores", "value")
)
def actualizar_grafica(metricas, jugadores):
    if not metricas:
        metricas = ["wellness_score"]

    if not jugadores or "TODOS" in jugadores:
        df_filtered = df_app.copy()
    else:
        df_filtered = df_app[df_app["player"].isin(jugadores)]

    if df_filtered.empty:
        return go.Figure().update_layout(template="plotly_white", title="Sin datos")

    df_grouped = df_filtered.groupby("date", as_index=False)[metricas].mean()

    fig = go.Figure()
    colores = [COLOR_ROJO, COLOR_ORO, COLOR_VERDE, "#7B1FA2", COLOR_AZUL]

    for i, metrica in enumerate(metricas):
        fig.add_trace(go.Scatter(
            x=df_grouped["date"],
            y=df_grouped[metrica],
            mode="lines+markers",
            name=traduccion_metricas.get(metrica, metrica),
            line=dict(width=4, color=colores[i % len(colores)]),
            marker=dict(size=7, color="white", line=dict(width=2, color=colores[i % len(colores)]))
        ))

    for f in df_grouped[df_grouped["date"].dt.dayofweek == 0]["date"]:
        fig.add_vline(x=f, line_width=1, line_dash="dot", line_color=COLOR_AZUL, opacity=0.18)

    fig.update_layout(
        template="plotly_white",
        height=520,
        autosize=True,
        margin=dict(l=40, r=25, t=25, b=45),
        legend=dict(orientation="h", y=1.12, x=0),
        yaxis=dict(range=[-0.5, 10.5], gridcolor="#EEF1F6", title="Puntuacion"),
        xaxis=dict(gridcolor="#F4F6FA"),
        hovermode="x unified",
        font=dict(family="Inter, Segoe UI, Arial")
    )

    return fig

@app.callback(
    Output("jugador-diario", "options"),
    Output("jugador-diario", "value"),
    Input("fecha-diaria", "date")
)
def cargar_jugadores_diario(fecha):
    fecha_dt = pd.to_datetime(fecha)
    jugadores = sorted(df_diario[df_diario["fecha"] == fecha_dt]["jugador"].unique().tolist())
    return [{"label": j, "value": j} for j in jugadores], (jugadores[0] if jugadores else None)

@app.callback(
    Output("download-informe", "data"),
    Output("estado-informe", "children"),
    Input("btn-informe", "n_clicks"),
    State("periodo", "start_date"),
    State("periodo", "end_date"),
    State("jugador-informe", "value"),
    prevent_initial_call=True
)
def descargar_informe_pdf(n_clicks, start_date, end_date, jugador):
    if not jugador:
        return no_update, "Selecciona un jugador concreto para generar el informe."

    ruta_pdf, mensaje = crear_pdf_completo(jugador, start_date, end_date)

    if ruta_pdf is None:
        return no_update, mensaje

    return dcc.send_file(ruta_pdf, filename=os.path.basename(ruta_pdf)), mensaje

@app.callback(
    Output("contenido-tabs", "children"),
    Input("tabs", "value"),
    Input("periodo", "start_date"),
    Input("periodo", "end_date"),
    Input("jugadores", "value"),
    Input("fecha-diaria", "date"),
    Input("jugador-diario", "value"),
    Input("jugador-informe", "value")
)
def actualizar_tabs(tab, start_date, end_date, jugadores, fecha_diaria, jugador_diario, jugador_informe):
    if tab == "periodo":
        f_ini = pd.to_datetime(start_date)
        f_fin = pd.to_datetime(end_date)

        if not jugador_informe:
            return html.Div("Selecciona un jugador para el informe de periodo.", style={"padding": "22px"})

        df_periodo = df_app[
            (df_app["date"] >= f_ini) &
            (df_app["date"] <= f_fin) &
            (df_app["player"] == jugador_informe)
        ]

        titulo = jugador_informe

        if df_periodo.empty:
            return html.Div("No hay datos en el periodo seleccionado.", style={"padding": "22px"})

        media_periodo = round(df_periodo["wellness_score"].mean(), 2)
        sesiones = len(df_periodo)
        jugadores_periodo = df_periodo["player"].nunique()

        alertas = []
        if "alert" in df_periodo.columns:
            df_alertas = df_periodo[df_periodo["alert"].notna()].sort_values("date", ascending=False)
            for _, r in df_alertas.head(10).iterrows():
                alertas.append(html.Tr([
                    html.Td(r["player"]),
                    html.Td(r["date"].strftime("%d/%m/%Y")),
                    html.Td(str(r["alert"]))
                ]))

        return html.Div([
            html.Div(className="top-kpis", children=[
                kpi_card("Analisis", titulo, COLOR_AZUL),
                kpi_card("Media Periodo", f"{media_periodo}", COLOR_ROJO),
                kpi_card("Sesiones", f"{sesiones}", COLOR_AZUL),
                kpi_card("Jugadores", f"{jugadores_periodo}", COLOR_ORO)
            ]),
            html.Div(className="section-spacer"),
            html.Div("Alertas del periodo", className="panel-title"),
            html.Table(className="alert-table", children=[
                html.Thead(html.Tr([html.Th("Jugador"), html.Th("Fecha"), html.Th("Alerta")])),
                html.Tbody(alertas if alertas else [html.Tr([html.Td("Sin alertas activas", colSpan=3)])])
            ])
        ])

    fecha_dt = pd.to_datetime(fecha_diaria)

    if not jugador_diario:
        return html.Div("No hay jugadores disponibles para la fecha seleccionada.", style={"padding": "22px"})

    data = get_status_diario(jugador_diario, fecha_dt)

    if not data:
        return html.Div("Sin datos para ese jugador y fecha.", style={"padding": "22px"})

    dia_semana_sel = fecha_dt.dayofweek
    df_hist_radar = df_diario[
        (df_diario["jugador"] == jugador_diario) &
        (df_diario["fecha"] <= fecha_dt) &
        (df_diario["fecha"].dt.dayofweek == dia_semana_sel)
    ]

    medias = df_hist_radar[["wellness", "fis_pre", "men_pre", "ganas_entrenar", "sueno"]].mean()

    labels = ["Bienestar", "Fisico PRE", "Animico PRE", "Ganas", "Sueno"]
    hoy = [
        data["registro"]["wellness"],
        data["registro"]["fis_pre"],
        data["registro"]["men_pre"],
        data["registro"]["ganas_entrenar"],
        data["registro"]["sueno"]
    ]
    hist = [
        medias["wellness"],
        medias["fis_pre"],
        medias["men_pre"],
        medias["ganas_entrenar"],
        medias["sueno"]
    ]

    fig_radar = go.Figure()
    fig_radar.add_trace(go.Scatterpolar(
        r=hist + [hist[0]],
        theta=labels + [labels[0]],
        fill="toself",
        name="Media historica",
        line=dict(color="#A1AAB8", dash="dash")
    ))
    fig_radar.add_trace(go.Scatterpolar(
        r=hoy + [hoy[0]],
        theta=labels + [labels[0]],
        fill="toself",
        name="Estado dia",
        line=dict(color=COLOR_ROJO, width=4)
    ))
    fig_radar.update_layout(
        template="plotly_white",
        height=430,
        margin=dict(l=35, r=35, t=40, b=30),
        polar=dict(radialaxis=dict(visible=True, range=[0, 1])),
        title=dict(text="Perfil individual", font=dict(color=COLOR_AZUL, size=18)),
        showlegend=True
    )

    df_dia = df_diario[df_diario["fecha"] == fecha_dt].sort_values("wellness", ascending=False)

    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=df_dia["jugador"],
        y=df_dia["wellness"],
        marker_color=[COLOR_ROJO if j == jugador_diario else COLOR_AZUL for j in df_dia["jugador"]],
        name="Wellness"
    ))
    fig_bar.add_hline(
        y=df_dia["wellness"].mean(),
        line_dash="dash",
        line_color="#111827",
        annotation_text=f"Media equipo {df_dia['wellness'].mean():.2f}"
    )
    fig_bar.update_layout(
        template="plotly_white",
        height=430,
        margin=dict(l=35, r=20, t=40, b=80),
        title=dict(text="Ranking grupal", font=dict(color=COLOR_AZUL, size=18)),
        yaxis=dict(range=[0, 1.1]),
        showlegend=False
    )

    df_hist = df_diario[
        (df_diario["jugador"] == jugador_diario) &
        (df_diario["fecha"] <= fecha_dt) &
        (df_diario["fecha"].dt.dayofweek == dia_semana_sel)
    ].sort_values("fecha")

    fig_hist = go.Figure()
    fig_hist.add_trace(go.Scatter(
        x=df_hist["fecha"],
        y=df_hist["wellness"],
        mode="lines+markers",
        line=dict(color=COLOR_AZUL, width=4),
        marker=dict(size=8),
        name="Wellness"
    ))
    fig_hist.add_hline(
        y=df_hist["wellness"].mean(),
        line_dash="dash",
        line_color=COLOR_ROJO,
        annotation_text="Media acumulada"
    )
    fig_hist.update_layout(
        template="plotly_white",
        height=340,
        margin=dict(l=35, r=20, t=40, b=35),
        title=dict(text=f"Evolucion historica de los {data['nombre_dia']}s", font=dict(color=COLOR_AZUL, size=18)),
        yaxis=dict(range=[0, 1.1]),
        hovermode="x unified"
    )

    filas_alertas = []
    for p in sorted(df_dia["jugador"].unique()):
        d = get_status_diario(p, fecha_dt)
        if d and d["status"] != "APTO":
            filas_alertas.append(html.Tr([
                html.Td(p),
                html.Td(html.Span(d["status"], className="status-pill", style={"background": d["color"]})),
                html.Td(f"{d['desv']:.1f}%"),
                html.Td(fecha_dt.strftime("%d/%m/%Y"))
            ]))

    return html.Div([
        html.Div(className="top-kpis", children=[
            kpi_card("Disponibilidad", data["status"], data["color"], f"Desviacion {data['desv']:.1f}%"),
            kpi_card("Wellness del dia", f"{data['val_actual']:.2f}", COLOR_AZUL),
            kpi_card(f"Media historica {data['nombre_dia']}", f"{data['media_hist']:.2f}", "#666"),
            kpi_card("Accion recomendada", data["accion"], data["color"])
        ]),
        html.Div(style={"display": "grid", "gridTemplateColumns": "1fr 1.3fr", "gap": "18px", "marginTop": "18px"}, children=[
            dcc.Graph(figure=fig_radar, config={"displayModeBar": True}),
            dcc.Graph(figure=fig_bar, config={"displayModeBar": True})
        ]),
        dcc.Graph(figure=fig_hist, config={"displayModeBar": True}),
        html.Div("Alertas detectadas ese dia", className="panel-title"),
        html.Table(className="alert-table", children=[
            html.Thead(html.Tr([html.Th("Jugador"), html.Th("Estado"), html.Th("Desviacion"), html.Th("Fecha")])),
            html.Tbody(filas_alertas if filas_alertas else [html.Tr([html.Td("Ninguna alerta detectada", colSpan=4)])])
        ])
    ])

# ------------------------------------------------------------
# LANZAR EN NUEVA PAGINA
# ------------------------------------------------------------

# Puerto aleatorio por defecto para no chocar con otra instancia; se puede
# fijar con WELLNESS_PORT cuando hace falta una URL estable.
PORT = int(os.environ.get("WELLNESS_PORT", 8050 + random.randint(0, 999)))

def abrir_navegador():
    time.sleep(1.5)
    webbrowser.open(f"http://127.0.0.1:{PORT}")

if __name__ == "__main__":
    threading.Thread(target=abrir_navegador, daemon=True).start()
    print(f"Dashboard listo. Abriendo http://127.0.0.1:{PORT}")
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)
