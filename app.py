import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from io import BytesIO
import hashlib
import re

from calibrador import calcular, pesos_iguales_porcentaje
from export_excel import crear_excel_resultados

st.set_page_config(
    page_title="Calibrador de Perfiles · evaluar.com",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Logo cargado dinámicamente desde el directorio del proyecto
import base64, os as _os
_logo_path = _os.path.join(_os.path.dirname(__file__), "brand_evaluar_on_dark.svg")
_logo_fallback = _os.path.join(_os.path.dirname(__file__), "brand_evaluar_white.svg")
_lp = _logo_path if _os.path.exists(_logo_path) else _logo_fallback
with open(_lp, "rb") as _f:
    SVG_LOGO = f"data:image/svg+xml;base64,{base64.b64encode(_f.read()).decode()}"
C_DARK   = "#22194e"
C_ACCENT = "#ff4298"
C_ORANGE = "#ffab48"
C_BG     = "#f9f7fc"
C_WHITE  = "#ffffff"
C_BORDER = "#e8e2f5"

def clasificar(cap_val):
    if cap_val >= 85:   return "Adecuado", "#4a9e6b"
    elif cap_val >= 70: return "Cercano",  "#d4893a"
    else:               return "Alejado",  "#c95f5f"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&family=Nunito+Sans:wght@300;400;600&display=swap');
html,body,[class*="css"]{{font-family:'Nunito Sans',sans-serif;background:{C_BG};}}
.stApp{{background:{C_BG};}}
section[data-testid="stSidebar"]>div{{background:{C_DARK} !important;border-right:none;}}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] .stSlider label,
section[data-testid="stSidebar"] span{{color:white !important;}}
section[data-testid="stSidebar"] .stButton>button{{
    background:rgba(255,255,255,0.12) !important;
    color:white !important;
    border:1px solid rgba(255,255,255,0.3) !important;
    font-weight:600 !important;
    font-size:0.82rem !important;
}}
section[data-testid="stSidebar"] .stButton>button:hover{{
    background:rgba(255,255,255,0.22) !important;
    border-color:rgba(255,255,255,0.5) !important;
}}
section[data-testid="stSidebar"] [data-testid="stNumberInput"] input{{
    color:{C_DARK} !important;
    -webkit-text-fill-color:{C_DARK} !important;
    text-align:center;
    font-size:0.76rem !important;
    font-weight:800;
    padding:0 4px !important;
}}
section[data-testid="stSidebar"] [data-testid="stNumberInput"]{{
    max-width:68px;
    margin-left:auto;
}}
section[data-testid="stSidebar"] [data-testid="stNumberInput"] [data-baseweb="input"]{{
    background:white !important;
    min-height:30px !important;
    height:30px !important;
    border:1px solid rgba(255,255,255,0.55) !important;
    border-radius:7px !important;
}}
section[data-testid="stSidebar"] [data-testid="stNumberInput"] button{{
    display:none !important;
}}
section[data-testid="stSidebar"] .stDownloadButton>button,
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button{{
    background:{C_ACCENT} !important;
    color:white !important;
    border:1px solid {C_ACCENT} !important;
    font-weight:700 !important;
    font-size:0.78rem !important;
}}
section[data-testid="stSidebar"] .stDownloadButton>button:hover,
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button:hover{{
    background:#e93686 !important;
    border-color:#e93686 !important;
}}
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button:disabled{{
    opacity:0.42 !important;
    cursor:not-allowed !important;
}}
/* File uploader estilizado */
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]{{
    background:rgba(255,255,255,0.05) !important;
    border:1px dashed rgba(255,255,255,0.2) !important;
    border-radius:8px !important;
}}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button{{
    background:rgba(255,255,255,0.12) !important;
    color:white !important;
    border:1px solid rgba(255,255,255,0.3) !important;
    font-size:0.8rem !important;
    font-weight:600 !important;
}}
/* Ocultar todo el bloque de instrucciones nativo */
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"]{{
    display:none !important;
}}
/* Nombre del archivo subido: hacerlo visible */
section[data-testid="stSidebar"] [data-testid="stFileUploader"] span,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] small,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] p{{
    color:rgba(255,255,255,0.75) !important;
}}
.evl-header{{background:linear-gradient(135deg,{C_DARK} 0%,#3d2980 100%);border-radius:14px;
    padding:1.1rem 1.8rem;margin-bottom:1.2rem;display:flex;align-items:center;justify-content:space-between;}}
.evl-title{{font-family:'Nunito',sans-serif;font-size:1.35rem;font-weight:800;color:white;margin:0;}}
.evl-sub{{font-size:0.75rem;color:rgba(255,255,255,0.5);margin-top:0.15rem;}}
.evl-badge{{background:rgba(255,255,255,0.12);border:1px solid rgba(255,255,255,0.2);
    color:{C_ORANGE};font-size:0.68rem;font-weight:700;letter-spacing:1px;
    padding:3px 11px;border-radius:20px;text-transform:uppercase;white-space:nowrap;}}
.evl-metric{{background:{C_WHITE};border-radius:10px;padding:0.75rem 1rem;
    border:1px solid {C_BORDER};text-align:center;}}
.evl-metric .lbl{{font-size:0.63rem;text-transform:uppercase;letter-spacing:0.8px;color:#999;font-weight:600;margin-bottom:0.2rem;}}
.evl-metric .val{{font-family:'Nunito',sans-serif;font-size:1.65rem;font-weight:800;color:{C_DARK};line-height:1;}}
.evl-metric .sub{{font-size:0.65rem;color:#bbb;margin-top:0.15rem;}}
.evl-section{{font-family:'Nunito',sans-serif;font-size:0.88rem;font-weight:700;color:{C_DARK};
    margin:1rem 0 0.6rem;padding-bottom:0.3rem;border-bottom:2px solid {C_BORDER};}}
.evl-card{{background:{C_WHITE};border-radius:10px;border:1px solid {C_BORDER};padding:0.8rem 1rem;margin-bottom:0.5rem;}}
.pill-v{{background:#e8f5ee;color:#2d6a4a;padding:2px 8px;border-radius:20px;font-size:0.7rem;font-weight:700;}}
.pill-n{{background:#fdf0e0;color:#8a5a2a;padding:2px 8px;border-radius:20px;font-size:0.7rem;font-weight:700;}}
.pill-r{{background:#faeaea;color:#8b3a3a;padding:2px 8px;border-radius:20px;font-size:0.7rem;font-weight:700;}}
</style>
""", unsafe_allow_html=True)


@st.cache_data
def parse_excel(file_bytes):
    df_raw = pd.read_excel(BytesIO(file_bytes), header=None)

    # ── Detectar formato ──────────────────────────────────────────────────────
    # Formato evaluar.com nativo: celda A1 contiene "Nombre del Proceso"
    # Formato nuevo (API/export): fila 0 contiene headers tipo "processId", "Cap", etc.
    es_formato_nativo = str(df_raw.iloc[0, 0]).strip() == "Nombre del Proceso"

    if es_formato_nativo:
        return _parse_formato_nativo(df_raw)
    else:
        return _parse_formato_api(df_raw)


def _parse_formato_nativo(df_raw):
    """Formato estándar descargado desde evaluar.com."""
    meta = {}
    for i in range(5):
        k = str(df_raw.iloc[i, 0]).strip()
        v = str(df_raw.iloc[i, 1]).strip() if pd.notna(df_raw.iloc[i, 1]) else ""
        if k not in ["nan", ""]:
            meta[k] = v

    row5 = df_raw.iloc[5].tolist()
    row6 = df_raw.iloc[6].tolist()

    col_estado    = next((i for i, v in enumerate(row5) if str(v).strip() == "Estado"),    6)
    col_nombres   = next((i for i, v in enumerate(row5) if str(v).strip() == "Nombres"),   3)
    col_apellidos = next((i for i, v in enumerate(row5) if str(v).strip() == "Apellidos"), 4)
    primera_comp_col = next((i for i, v in enumerate(row6) if str(v).strip() == "Valor"), 21)

    competencias = {}
    for idx, val in enumerate(row5):
        if idx >= primera_comp_col and pd.notna(val):
            nombre = str(val).strip()
            if nombre in ["", "Detalles del candidato", "TRUST", "Disc"]:
                continue
            if str(row6[idx]).strip() == "Valor":
                competencias[nombre] = {
                    "valor": idx, "esperado": idx+1, "brecha": idx+2, "cumplimiento": idx+3
                }

    rows = []
    for i in range(7, len(df_raw)):
        row = df_raw.iloc[i]
        if str(row[col_estado]).strip() != "TERMINADO":
            continue
        nombres   = str(row[col_nombres]).strip()   if pd.notna(row[col_nombres])   else ""
        apellidos = str(row[col_apellidos]).strip()  if pd.notna(row[col_apellidos])  else ""
        c = {"Candidato": f"{nombres} {apellidos}".strip(),
             "CAP_archivo": float(row[1]) if pd.notna(row[1]) else np.nan}
        for comp, cols in competencias.items():
            c[f"{comp}__valor"]    = float(row[cols["valor"]])    if pd.notna(row[cols["valor"]])    else np.nan
            c[f"{comp}__esperado"] = float(row[cols["esperado"]]) if pd.notna(row[cols["esperado"]]) else np.nan
        rows.append(c)
    return pd.DataFrame(rows), competencias, meta


def _parse_formato_api(df_raw):
    """Formato exportado vía API: fila 0 = headers, competencias como columnas pares."""
    # Leer con header real
    df = df_raw.copy()
    df.columns = df.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)

    # Metadatos desde columnas
    meta = {}
    for col, key in [("processName", "Nombre del Proceso"),
                     ("positionName", "Nombre del Perfil"),
                     ("publicCompanyName", "Empresa")]:
        if col in df.columns:
            vals = df[col].dropna()
            if len(vals) > 0:
                meta[key] = str(vals.iloc[0]).strip()

    # Detectar competencias: columnas que tienen su par "_expected"
    excluir = {"personId","processId","publicCompanyName","processName",
               "positionName","firstName","lastName","identification",
               "createdAt","degree","Cap"}
    competencias = {}
    for col in df.columns:
        if str(col).endswith("_expected") or col in excluir:
            continue
        exp_col = f"{col}_expected"
        if exp_col in df.columns:
            competencias[str(col).strip()] = {"valor_col": col, "esperado_col": exp_col}

    # Construir filas
    rows = []
    for _, row in df.iterrows():
        nombre = f"{str(row.get('firstName','')).strip()} {str(row.get('lastName','')).strip()}".strip()
        cap = float(row["Cap"]) if pd.notna(row.get("Cap")) else np.nan
        c = {"Candidato": nombre, "CAP_archivo": cap}
        for comp, cols in competencias.items():
            v = row.get(cols["valor_col"])
            e = row.get(cols["esperado_col"])
            c[f"{comp}__valor"]    = float(v) if pd.notna(v) else np.nan
            c[f"{comp}__esperado"] = float(e) if pd.notna(e) else np.nan
        rows.append(c)

    # Convertir competencias al mismo esquema que formato nativo
    comp_normalizado = {k: {"valor": k, "esperado": f"{k}_expected"} for k in competencias}
    return pd.DataFrame(rows), comp_normalizado, meta


def dist_rangos(series):
    s = series.dropna()
    total = len(s)
    if total == 0:
        return {k: {"n":0,"pct":0.0,"color":"#ccc"} for k in ["Adecuado","Cercano","Alejado"]}
    def p(n): return round(n/total*100, 1)
    na = int((s >= 85).sum())
    nc = int(((s >= 70) & (s < 85)).sum())
    nr = int((s <  70).sum())
    return {
        "Adecuado": {"n":na, "pct":p(na), "color":"#4a9e6b"},
        "Cercano":  {"n":nc, "pct":p(nc), "color":"#d4893a"},
        "Alejado":  {"n":nr, "pct":p(nr), "color":"#c95f5f"},
    }


def barra_dist(dist, height=55):
    fig = go.Figure()
    for label in ["Alejado", "Cercano", "Adecuado"]:
        d = dist[label]
        txt = f"<b>{d['pct']}%</b>" if d["pct"] >= 8 else ""
        fig.add_trace(go.Bar(
            name=label, x=[d["pct"]], y=[""],
            orientation="h", marker_color=d["color"],
            text=txt, textposition="inside", insidetextanchor="middle",
            textfont=dict(color="white", size=10, family="Nunito"),
            hovertemplate=f"<b>{label}</b>: {d['n']} cand. ({d['pct']}%)<extra></extra>",
            showlegend=False,
        ))
    fig.update_layout(
        barmode="stack", height=height,
        margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="white", plot_bgcolor="white",
        xaxis=dict(range=[0,100], showticklabels=False, showgrid=False, zeroline=False),
        yaxis=dict(showticklabels=False, showgrid=False),
        bargap=0,
    )
    return fig


def dona_dist(dist, height=200):
    labels = ["Alejado", "Cercano", "Adecuado"]
    values = [dist[l]["pct"] for l in labels]
    colors = [dist[l]["color"] for l in labels]
    fig = go.Figure(go.Pie(
        labels=labels,
        values=values,
        hole=0.62,
        marker=dict(colors=colors, line=dict(color="white", width=2)),
        textinfo="percent",
        textfont=dict(size=11, family="Nunito", color="white"),
        hovertemplate="<b>%{label}</b>: %{value}%<extra></extra>",
        sort=False,
    ))
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="white",
        showlegend=True,
        legend=dict(
            orientation="v", x=1.0, y=0.5, xanchor="left", yanchor="middle",
            font=dict(size=10, family="Nunito"),
        ),
    )
    return fig


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style='padding:1rem 0 1.1rem;border-bottom:1px solid rgba(255,255,255,0.1);margin-bottom:1rem;'>
        <img src='{SVG_LOGO}' style='height:20px;'>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader("Subir archivo", type=["xlsx"],
                                label_visibility="visible",
                                help="Descarga directamente desde evaluar.com sin modificar el archivo")
    if not uploaded:
        st.markdown("""
        <div style='background:rgba(255,255,255,0.05);border-radius:10px;padding:0.9rem;
                    font-size:0.75rem;color:rgba(255,255,255,0.5);line-height:1.7;margin-top:0.8rem;'>
            <b style='color:rgba(255,255,255,0.75);'>Cómo usar:</b><br>
            1. Descarga el reporte desde evaluar.com<br>
            2. Súbelo aquí sin modificarlo<br>
            3. Ajusta el esperado y el peso<br>
            4. Observa el impacto en tiempo real
        </div>
        """, unsafe_allow_html=True)
        st.stop()

file_bytes = uploaded.getvalue()
file_id = hashlib.sha1(file_bytes).hexdigest()[:10]
df_base, competencias, meta = parse_excel(file_bytes)
if df_base.empty:
    st.error("No se encontraron candidatos con estado TERMINADO.")
    st.stop()
if not competencias:
    st.error("No se encontraron competencias en el archivo.")
    st.stop()

esperados_actuales = {}
for comp in competencias:
    col = f"{comp}__esperado"
    if col in df_base.columns:
        v = df_base[col].dropna()
        esperados_actuales[comp] = float(v.iloc[0]) if len(v) > 0 else 5.0
pesos_iniciales = pesos_iguales_porcentaje(competencias)

with st.sidebar:
    if st.button("↩ Restaurar originales", width="stretch"):
        for comp in competencias:
            st.session_state[f"sl_{file_id}_{comp}"] = int(round(esperados_actuales.get(comp, 5.0)))
            st.session_state[f"wt_{file_id}_{comp}"] = pesos_iniciales[comp]
        st.rerun()

    st.markdown("<div style='font-size:0.62rem;text-transform:uppercase;letter-spacing:1px;color:rgba(255,255,255,0.35);margin:0.6rem 0 0.2rem;'>Calibración por competencia</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:0.62rem;color:rgba(255,255,255,0.45);margin-bottom:0.5rem;'>Distribuye el peso total entre las competencias.</div>", unsafe_allow_html=True)
    st.markdown("<div style='display:grid;grid-template-columns:1fr 68px;column-gap:0.5rem;font-size:0.58rem;color:rgba(255,255,255,0.55);margin-bottom:0.35rem;text-align:center;'><span style='text-align:right;padding-right:1rem;'>ESPERADO</span><span>PESO %</span></div>", unsafe_allow_html=True)
    esperados_sim = {}
    pesos_sim = {}
    for comp in competencias:
        act = int(round(esperados_actuales.get(comp, 5.0)))
        key_esp = f"sl_{file_id}_{comp}"
        key_peso = f"wt_{file_id}_{comp}"
        if key_esp not in st.session_state:
            st.session_state[key_esp] = act
        if key_peso not in st.session_state:
            st.session_state[key_peso] = pesos_iniciales[comp]
        st.markdown(f"<div style='font-size:0.75rem;font-weight:600;color:white;margin-top:0.45rem;'>{comp}</div>", unsafe_allow_html=True)
        col_esp, col_peso = st.columns([3.15, 1.0], gap="small")
        with col_esp:
            esperados_sim[comp] = st.slider(
                f"Esperado · {comp}", 1, 10, key=key_esp, label_visibility="collapsed"
            )
        with col_peso:
            pesos_sim[comp] = st.number_input(
                f"Peso · {comp}", min_value=0.0, max_value=100.0, step=0.1,
                format="%.1f", key=key_peso, label_visibility="collapsed"
            )

    total_pesos = round(sum(pesos_sim.values()), 1)
    pesos_validos = abs(total_pesos - 100.0) < 0.05
    color_total = "#7ee2a8" if pesos_validos else "#ff8dbf"
    texto_total = "✓ Total correcto" if pesos_validos else "Ajusta hasta 100%"
    st.markdown(
        f"<div style='margin:0.7rem 0 0.25rem;padding:0.45rem 0.65rem;border-radius:7px;"
        f"background:rgba(255,255,255,0.07);display:flex;justify-content:space-between;"
        f"font-size:0.69rem;font-weight:700;'>"
        f"<span style='color:{color_total}!important;'>{texto_total}</span>"
        f"<span style='color:{color_total}!important;'>{total_pesos:.1f}%</span></div>",
        unsafe_allow_html=True,
    )

    pesos_calculo = pesos_sim if total_pesos > 0 else pesos_iniciales

    cambio_esperados = any(
        abs(esperados_sim[k] - esperados_actuales.get(k, 0)) > 0.01
        for k in esperados_sim
    )
    cambio_pesos = any(
        abs(pesos_sim[k] - pesos_iniciales[k]) > 0.05 for k in competencias
    )
    en_sim = cambio_esperados or cambio_pesos
    if en_sim:
        st.markdown("""
        <div style='background:rgba(255,66,152,0.12);border:1px solid rgba(255,66,152,0.3);
                    border-radius:8px;padding:0.45rem 0.7rem;margin-top:0.8rem;
                    font-size:0.72rem;color:#ff4298;text-align:center;font-weight:700;'>
            ⚡ SIMULACIÓN ACTIVA
        </div>""", unsafe_allow_html=True)

# ── CALCULAR ──────────────────────────────────────────────────────────────────
pesos_actuales = pesos_iniciales
df_actual = calcular(df_base, competencias, esperados_actuales, pesos_actuales)
df       = calcular(df_base, competencias, esperados_sim, pesos_calculo)
n_total  = len(df)
proceso  = meta.get("Nombre del Perfil", meta.get("Nombre del Proceso", ""))
dist_g   = dist_rangos(df["CAP_global"])
cap_prom = df["CAP_global"].mean()

# ── DESCARGA EN SIDEBAR ──────────────────────────────────────────────────────
nombre_base = meta.get("Nombre del Perfil") or meta.get("Nombre del Proceso") or "calibracion"
nombre_base = re.sub(r"[^A-Za-z0-9_-]+", "-", nombre_base.strip()).strip("-").lower()
excel_bytes = b""
if pesos_validos:
    excel_bytes = crear_excel_resultados(
        meta=meta,
        df_base=df_base,
        df_actual=df_actual,
        df_simulado=df,
        competencias=competencias,
        esperados_actuales=esperados_actuales,
        esperados_simulados=esperados_sim,
        pesos_simulados=pesos_sim,
    )
with st.sidebar:
    st.markdown("<div style='height:1px;background:rgba(255,255,255,0.12);margin:0.75rem 0;'></div>", unsafe_allow_html=True)
    st.download_button(
        "⬇ Descargar análisis Excel",
        data=excel_bytes,
        file_name=f"calibracion-{nombre_base or 'perfil'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        disabled=not pesos_validos,
        help=None if pesos_validos else "Los pesos deben sumar 100% para descargar.",
    )

# ── ALERTA DE BALANCE ─────────────────────────────────────────────────────────  ← NUEVO
pct_adecuados = dist_g["Adecuado"]["pct"]
_mostrar_alerta = pct_adecuados > 30 or pct_adecuados < 10

# ── HEADER ────────────────────────────────────────────────────────────────────
# Metadatos con etiquetas
_meta_items = [
    ("Proceso",     meta.get("Nombre del Proceso", "")),
    ("Perfil",      meta.get("Nombre del Perfil",  "")),
    ("Inicio",      meta.get("Inicio",     "")),
    ("Fin",         meta.get("Fin",        "")),
    ("Reclutador",  meta.get("Reclutador", "")),
]
_meta_html = "".join(
    f"<div style='font-size:0.67rem;color:rgba(255,255,255,0.4);margin-top:1px;'>"
    f"<span style='color:rgba(255,255,255,0.25);'>{k}:</span> "
    f"<span style='color:rgba(255,255,255,0.65);'>{v}</span></div>"
    for k, v in _meta_items if v
)

st.markdown(f"""
<div class='evl-header'>
    <div style='display:flex;align-items:center;gap:1.2rem;'>
        <img src='{SVG_LOGO}' style='height:22px;opacity:0.95;flex-shrink:0;'>
        <div style='width:1px;height:36px;background:rgba(255,255,255,0.2);'></div>
        <div>
            <div class='evl-title'>Calibrador de Perfiles</div>
            {_meta_html}
        </div>
    </div>
    <div class='evl-badge'>{'⚡ Simulación activa' if en_sim else '✓ Datos reales'}</div>
</div>
""", unsafe_allow_html=True)

# métricas removidas

# ── DISTRIBUCIÓN GLOBAL: barras izq + dona der ────────────────────────────────
st.markdown("<div class='evl-section'>Distribución global — adecuación al perfil</div>", unsafe_allow_html=True)

col_barras, col_dona = st.columns([3, 2], gap="large")

with col_barras:
    # Barras por competencia (compactas)
    st.markdown("<div style='font-size:0.75rem;font-weight:600;color:#888;margin:0.8rem 0 0.4rem;text-transform:uppercase;letter-spacing:0.5px;'>Por competencia</div>", unsafe_allow_html=True)
    for comp in competencias:
        cap_col = f"{comp}__cap"
        if cap_col not in df.columns:
            continue
        dist_c  = dist_rangos(df[cap_col].dropna())
        esp_act = esperados_actuales.get(comp, 5.0)
        esp_sim = esperados_sim.get(comp, esp_act)
        cambio  = f" → <b style='color:#ff4298;'>{esp_sim}</b>" if abs(esp_sim - esp_act) > 0.01 else ""
        peso_mostrado = pesos_sim[comp]

        st.markdown(f"""
        <div style='margin-bottom:0.6rem;'>
            <div style='font-size:0.75rem;font-weight:600;color:{C_DARK};margin-bottom:1px;
                        display:flex;justify-content:space-between;align-items:center;'>
                <span>{comp}</span>
                <span style='font-size:0.68rem;color:#aaa;font-weight:400;'>
                    Esp: <b style='color:{C_DARK};'>{esp_act}</b>{cambio}
                    &nbsp;·&nbsp; Peso: <b style='color:{C_DARK};'>{peso_mostrado:.1f}%</b>
                    &nbsp;·&nbsp; CAP: <b style='color:{C_DARK};'>{df[cap_col].mean():.1f}%</b>
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.plotly_chart(barra_dist(dist_c, height=38), width="stretch",
                        config={"displayModeBar": False}, key=f"barra_{comp}")

with col_dona:
    st.markdown("<div class='evl-card' style='display:flex;flex-direction:column;align-items:center;'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:0.82rem;font-weight:700;color:#555;text-align:center;margin-bottom:0.5rem;'>CAP Global · <span style='color:{C_DARK};font-size:1rem;'>n={n_total}</span></div>", unsafe_allow_html=True)
    st.plotly_chart(dona_dist(dist_g, height=230), width="stretch",
                    config={"displayModeBar": False}, key="dona_global")
    # Stats bajo la dona
    sc1, sc2, sc3 = st.columns(3)
    for scol, label, pill in zip([sc1, sc2, sc3],
                                  ["Alejado","Cercano","Adecuado"],
                                  ["pill-r","pill-n","pill-v"]):
        d = dist_g[label]
        scol.markdown(
            f"<div style='text-align:center;'>"
            f"<span class='{pill}' style='font-size:0.82rem;padding:3px 10px;'>{d['pct']}%</span>"
            f"<div style='font-size:0.75rem;font-weight:600;color:#666;margin-top:4px;'>{d['n']} cand.</div>"
            f"</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ── ALERTA ────────────────────────────────────────────────────────────────────
if _mostrar_alerta:
    if pct_adecuados > 30:
        _alerta_msg = (
            f"<b>Perfil posiblemente desbalanceado</b> — el porcentaje de candidatos "
            f"Adecuados es de <b>{pct_adecuados:.1f}%</b>. Probablemente el perfil sea "
            f"poco exigente o los candidatos estén sobreajustados al mismo."
        )
    else:
        _alerta_msg = (
            f"<b>Perfil posiblemente desbalanceado</b> — el porcentaje de candidatos "
            f"Adecuados es de <b>{pct_adecuados:.1f}%</b>. El perfil puede ser "
            f"demasiado exigente para el pool actual de candidatos."
        )
    st.markdown(f"""
    <div style='
        background:#fffbea;
        border-left:4px solid #f0a500;
        border-radius:6px;
        padding:0.75rem 1rem;
        margin-bottom:1rem;
        display:flex;
        align-items:flex-start;
        gap:0.75rem;
        font-size:0.82rem;
        color:#5a4200;
        line-height:1.55;
    '>
        <span style='font-size:1.1rem;margin-top:1px;flex-shrink:0;'>⚠️</span>
        <span>{_alerta_msg}</span>
    </div>
    """, unsafe_allow_html=True)

# ── TABLA CANDIDATOS ──────────────────────────────────────────────────────────
st.markdown("<div class='evl-section'>Detalle de candidatos</div>", unsafe_allow_html=True)
with st.expander("Ver tabla completa"):
    comp_cap_cols = [f"{c}__cap" for c in competencias if f"{c}__cap" in df.columns]
    df_tabla = df[["Candidato", "CAP_archivo", "CAP_global"] + comp_cap_cols].copy()
    df_tabla.insert(2, "CAP_actual", df_actual["CAP_global"])
    rename = {
        "CAP_archivo": "CAP archivo (%)",
        "CAP_actual": "CAP actual (%)",
        "CAP_global": "CAP simulado (%)",
    }
    rename.update({c: c.replace("__cap","") for c in comp_cap_cols})
    df_tabla = df_tabla.rename(columns=rename)
    df_tabla["Clasificación"] = df_tabla["CAP simulado (%)"].apply(lambda x: clasificar(x)[0])
    df_tabla = df_tabla.sort_values("CAP simulado (%)", ascending=False).reset_index(drop=True)
    df_tabla.index += 1

    def color_row(row):
        _, color = clasificar(row["CAP simulado (%)"])
        return [f"background-color:{color}18"] * len(row)

    num_cols = {c: "{:.1f}" for c in df_tabla.columns if c not in ["Candidato","Clasificación"]}
    st.dataframe(
        df_tabla.style.apply(color_row, axis=1).format(num_cols),
        width="stretch", height=400
    )
