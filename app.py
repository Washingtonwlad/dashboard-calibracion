import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from io import BytesIO

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


def cumplimiento_sim(valor, esperado):
    """
    Replica la fórmula de evaluar.com:
    - Si valor >= esperado: cumplimiento = 1.0 (100%)
    - Si valor < esperado:  cumplimiento = 1.0 - round(abs(brecha) / 10, 2)
      donde brecha = valor - esperado, escala base = 10
    """
    if pd.isna(valor) or pd.isna(esperado) or esperado <= 0:
        return np.nan
    brecha = valor - esperado
    if brecha >= 0:
        return 1.0
    return 1.0 - round(abs(brecha) / 10, 2)


def calcular(df, competencias, esperados_sim):
    df2 = df.copy()
    cumpl_cols = []
    for comp in competencias:
        esp = esperados_sim[comp]
        ccol, capcol = f"{comp}__cumpl_sim", f"{comp}__cap"
        df2[ccol]   = df2[f"{comp}__valor"].apply(lambda v: cumplimiento_sim(v, esp))
        df2[capcol] = df2[ccol] * 100
        cumpl_cols.append(ccol)
    if cumpl_cols:
        df2["CAP_global"] = df2[cumpl_cols].mean(axis=1) * 100
    return df2.round(3)


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
            3. Ajusta el puntaje esperado<br>
            4. Observa el impacto en tiempo real
        </div>
        """, unsafe_allow_html=True)
        st.stop()

df_base, competencias, meta = parse_excel(uploaded.read())
if df_base.empty:
    st.error("No se encontraron candidatos con estado TERMINADO.")
    st.stop()

esperados_actuales = {}
for comp in competencias:
    col = f"{comp}__esperado"
    if col in df_base.columns:
        v = df_base[col].dropna()
        esperados_actuales[comp] = float(v.iloc[0]) if len(v) > 0 else 5.0

with st.sidebar:
    if st.button("↩ Restaurar originales", use_container_width=True):
        for comp in competencias:
            key = f"sl_{comp}"
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

    st.markdown("<div style='font-size:0.62rem;text-transform:uppercase;letter-spacing:1px;color:rgba(255,255,255,0.35);margin:0.6rem 0 0.5rem;'>Puntaje esperado por competencia</div>", unsafe_allow_html=True)
    esperados_sim = {}
    for comp in competencias:
        act = esperados_actuales.get(comp, 5.0)
        esperados_sim[comp] = st.slider(comp, 1, 10, int(round(act)), 1, key=f"sl_{comp}")

    en_sim = any(abs(esperados_sim[k] - esperados_actuales.get(k, 0)) > 0.01 for k in esperados_sim)
    if en_sim:
        st.markdown("""
        <div style='background:rgba(255,66,152,0.12);border:1px solid rgba(255,66,152,0.3);
                    border-radius:8px;padding:0.45rem 0.7rem;margin-top:0.8rem;
                    font-size:0.72rem;color:#ff4298;text-align:center;font-weight:700;'>
            ⚡ SIMULACIÓN ACTIVA
        </div>""", unsafe_allow_html=True)

# ── CALCULAR ──────────────────────────────────────────────────────────────────
df       = calcular(df_base, competencias, esperados_sim)
n_total  = len(df)
proceso  = meta.get("Nombre del Perfil", meta.get("Nombre del Proceso", ""))
dist_g   = dist_rangos(df["CAP_global"])
cap_prom = df["CAP_global"].mean()

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

        st.markdown(f"""
        <div style='margin-bottom:0.6rem;'>
            <div style='font-size:0.75rem;font-weight:600;color:{C_DARK};margin-bottom:1px;
                        display:flex;justify-content:space-between;align-items:center;'>
                <span>{comp}</span>
                <span style='font-size:0.68rem;color:#aaa;font-weight:400;'>
                    Esp: <b style='color:{C_DARK};'>{esp_act}</b>{cambio}
                    &nbsp;·&nbsp; CAP: <b style='color:{C_DARK};'>{df[cap_col].mean():.1f}%</b>
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.plotly_chart(barra_dist(dist_c, height=38), use_container_width=True,
                        config={"displayModeBar": False})

with col_dona:
    st.markdown("<div class='evl-card' style='display:flex;flex-direction:column;align-items:center;'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:0.82rem;font-weight:700;color:#555;text-align:center;margin-bottom:0.5rem;'>CAP Global · <span style='color:{C_DARK};font-size:1rem;'>n={n_total}</span></div>", unsafe_allow_html=True)
    st.plotly_chart(dona_dist(dist_g, height=230), use_container_width=True,
                    config={"displayModeBar": False})
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

# ── TABLA CANDIDATOS ──────────────────────────────────────────────────────────
st.markdown("<div class='evl-section'>Detalle de candidatos</div>", unsafe_allow_html=True)
with st.expander("Ver tabla completa"):
    comp_cap_cols = [f"{c}__cap" for c in competencias if f"{c}__cap" in df.columns]
    df_tabla = df[["Candidato", "CAP_archivo", "CAP_global"] + comp_cap_cols].copy()
    rename = {"CAP_archivo": "CAP archivo (%)", "CAP_global": "CAP simulado (%)"}
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
        use_container_width=True, height=400
    )