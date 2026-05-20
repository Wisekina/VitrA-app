import warnings
warnings.filterwarnings("ignore")

import base64
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import os

st.set_page_config(
    page_title="VitrA Karo · Talep Tahmin",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Login ─────────────────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <style>
    .stApp { background: #0D1B2A; }
    </style>
    """, unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style='text-align:center; margin-bottom:32px;'>
          <p style='color:rgba(255,255,255,.4); font-size:.75rem; letter-spacing:3px;
                    text-transform:uppercase; font-weight:700; margin:0;'>
            Eczacıbaşı VitrA Karo
          </p>
          <h2 style='color:white; font-size:1.6rem; font-weight:800; margin:8px 0 0;'>
            Talep Tahmin Sistemi
          </h2>
        </div>
        """, unsafe_allow_html=True)
        username = st.text_input("Kullanıcı Adı", placeholder="Kullanıcı adınızı girin")
        password = st.text_input("Şifre", type="password", placeholder="Şifrenizi girin")
        if st.button("Giriş Yap", use_container_width=True):
            if (username == st.secrets["credentials"]["username"] and
                    password == st.secrets["credentials"]["password"]):
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Kullanıcı adı veya şifre hatalı.")
    st.stop()

# ── Logo ──────────────────────────────────────────────────────────────────────
_LOGO_PATH = os.path.join(os.path.dirname(__file__), "vitra.jpeg")
with open(_LOGO_PATH, "rb") as _f:
    logo_b64 = base64.b64encode(_f.read()).decode()
LOGO = f'<img src="data:image/jpeg;base64,{logo_b64}"'

# ── Sabitler ──────────────────────────────────────────────────────────────────
BOLGE_FULL = {
    "AKDENIZ"     : "Akdeniz",
    "BATI_MARMARA": "Batı Marmara",
    "DOGU_MARMARA": "Doğu Marmara",
    "EGE"         : "Ege",
    "GUNEYDOGU"   : "Güneydoğu Anadolu",
    "KARADENIZ"   : "Karadeniz",
    "IC_ANADOLU"  : "İç Anadolu",
}
BOLGE_RENAME = {
    "AKDENİZ BÖLGESİ"          : "AKDENIZ",
    "BATI MARMARA BÖLGESİ"     : "BATI_MARMARA",
    "DOĞU MARMARA BÖLGESİ"     : "DOGU_MARMARA",
    "EGE BÖLGESİ"              : "EGE",
    "GÜNEYDOĞU ANADOLU BÖLGESİ": "GUNEYDOGU",
    "KARADENİZ BÖLGESİ"        : "KARADENIZ",
    "İÇ ANADOLU BÖLGESİ"       : "IC_ANADOLU",
}
BOLGE_COORDS = {
    "AKDENIZ"     : (37.0, 31.5), "BATI_MARMARA": (40.3, 28.0),
    "DOGU_MARMARA": (40.5, 30.5), "EGE"         : (38.2, 27.5),
    "GUNEYDOGU"   : (37.5, 38.5), "KARADENIZ"   : (41.0, 35.5),
    "IC_ANADOLU"  : (39.2, 33.0),
}
AY_TR = {1:"Ocak",2:"Şubat",3:"Mart",4:"Nisan",5:"Mayıs",6:"Haziran",
          7:"Temmuz",8:"Ağustos",9:"Eylül",10:"Ekim",11:"Kasım",12:"Aralık"}
_DIR      = os.path.dirname(__file__)
DATA_PATH = os.path.join(_DIR, "EBAT-BOLGE.xlsx")
FC_PATH   = os.path.join(_DIR, "tahmin_raporu_bolge_2026_augmented.xlsx")

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
* { font-family: 'Inter', sans-serif !important; box-sizing: border-box; }

/* Sayfa arka plan: sade beyaz */
.stApp { background: #F2F3F5; }

/* Ana içerik */
.main .block-container {
    padding: 0 2rem 3rem !important;
    max-width: 1400px;
}

/* Sidebar — tek renk banyo karosu */
section[data-testid="stSidebar"] {
    background-color: #0D1B2A !important;
    background-image:
        linear-gradient(rgba(255,255,255,0.12) 2px, transparent 2px),
        linear-gradient(90deg, rgba(255,255,255,0.12) 2px, transparent 2px) !important;
    background-size: 62px 62px !important;
    border-right: none;
    box-shadow: 2px 0 24px rgba(0,0,0,0.25);
}
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stSelectbox label {
    color: rgba(255,255,255,0.55) !important;
    font-size: 0.7rem !important;
    letter-spacing: 2px;
    text-transform: uppercase;
    font-weight: 600 !important;
    margin-bottom: 6px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] {
    background: rgba(255,255,255,0.07) !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
    border-radius: 10px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] * { color: white !important; }
section[data-testid="stSidebar"] [data-baseweb="select"] svg {
    fill: #E05C2A !important;
    width: 20px !important;
    height: 20px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] svg path {
    fill: #E05C2A !important;
}

/* Sekmeler */
button[data-baseweb="tab"] {
    font-size: 0.85rem !important; font-weight: 600 !important;
    color: #94A3B8 !important; padding: 12px 24px !important;
    border-bottom: 2px solid transparent !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #0D1B2A !important;
    border-bottom: 2px solid #E05C2A !important;
}

/* Metrik kutuları */
div[data-testid="metric-container"] {
    background: white;
    border-radius: 16px;
    padding: 22px 24px;
    border: 1px solid #EAECF0;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
div[data-testid="stMetricLabel"] p {
    font-size: 0.68rem !important;
    font-weight: 700 !important;
    letter-spacing: 1.8px !important;
    color: #94A3B8 !important;
    text-transform: uppercase !important;
    margin-bottom: 4px !important;
}
div[data-testid="stMetricValue"] {
    font-size: 1.8rem !important;
    font-weight: 800 !important;
    color: #0D1B2A !important;
    line-height: 1.1 !important;
}
div[data-testid="stMetricDelta"] { font-size: 0.8rem !important; }

/* Expander */
details {
    background: white;
    border-radius: 12px;
    border: 1px solid #EAECF0 !important;
    padding: 4px 16px;
}
details summary p { font-weight: 700 !important; color: #0D1B2A !important; }

/* Plotly grafik container */
.stPlotlyChart { border-radius: 16px; overflow: hidden; }

/* ── Tablet & mobil uyumu ── */
@media (max-width: 900px) {
    .main .block-container {
        padding: 0 0.5rem 2rem !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.2rem !important;
    }
    div[data-testid="stMetricLabel"] p {
        font-size: 0.58rem !important;
        letter-spacing: 1px !important;
    }
    button[data-baseweb="tab"] {
        font-size: 0.72rem !important;
        padding: 10px 10px !important;
    }
}
</style>
""", unsafe_allow_html=True)

# ── Veri yükleme ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Veri yükleniyor…")
def load_raw(path):
    df = pd.read_excel(path, sheet_name="Sheet1")
    df["Tarih"] = pd.to_datetime(df["Tarih"]).dt.to_period("M").dt.to_timestamp("M")
    df = df.rename(columns={"Dist. Grup": "Bolge"})
    df["Bolge"] = df["Bolge"].map(BOLGE_RENAME).fillna(df["Bolge"])
    ecols = [c for c in df.columns if c not in ["Tarih","Bolge"]]
    df = df.melt(id_vars=["Tarih","Bolge"], value_vars=ecols,
                 var_name="Ebat", value_name="Satis")
    df["Satis"] = df["Satis"].fillna(0).clip(lower=0)
    return df.sort_values(["Bolge","Ebat","Tarih"]).reset_index(drop=True), ecols

@st.cache_data(show_spinner="Tahminler yükleniyor…")
def load_forecast(path):
    xl = pd.read_excel(path, sheet_name=None)
    fc = xl["Tahminler_Long"] if "Tahminler_Long" in xl else xl[list(xl.keys())[0]]
    fc["Tarih"] = pd.to_datetime(fc["Tarih"])
    metrics = xl.get("Model_Metrikleri")
    if metrics is not None and "En_Iyi_MASE" not in metrics.columns and "En_Iyi" in metrics.columns:
        def _best_mase(row):
            col = f"{row['En_Iyi']}_MASE"
            return row[col] if col in row.index else 999
        metrics = metrics.copy()
        metrics["En_Iyi_MASE"] = metrics.apply(_best_mase, axis=1)
    return fc, metrics

@st.cache_data(show_spinner=False)
def garch_band(satis_t, fc_t, horizon=12):
    vals, fv = np.array(satis_t, float), np.array(fc_t, float)
    if len(vals) < 24 or vals.std() < 1e-6:
        return None, None
    try:
        from arch import arch_model as am
        res = am(vals, vol="GARCH", p=1, q=1, dist="normal", rescale=True).fit(
            disp="off", show_warning=False)
        fc  = res.forecast(horizon=horizon, reindex=False)
        cs  = np.sqrt(fc.variance.values[-1]) / res.scale
        return fv + 1.96*cs, np.clip(fv - 1.96*cs, 0, None)
    except Exception:
        return None, None

# ── Veri yükle ────────────────────────────────────────────────────────────────
if not os.path.exists(DATA_PATH):
    st.error(f"Veri dosyası bulunamadı: {DATA_PATH}")
    st.stop()

df_raw, EBAT_COLS = load_raw(DATA_PATH)
BOLGELER = sorted(df_raw["Bolge"].unique())
has_fc   = os.path.exists(FC_PATH)
df_fc, df_metrics = load_forecast(FC_PATH) if has_fc else (None, None)

# Varsayılan: en yüksek satış hacmi
_db, _de = 0, 0
try:
    _hacim = (df_raw.groupby(["Bolge","Ebat"])["Satis"].sum()
              .reset_index().sort_values("Satis", ascending=False))
    _b, _e = str(_hacim.iloc[0]["Bolge"]), str(_hacim.iloc[0]["Ebat"])
    if _b in BOLGELER: _db = BOLGELER.index(_b)
    if _e in EBAT_COLS: _de = EBAT_COLS.index(_e)
except Exception:
    pass

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="padding: 32px 16px 24px; text-align:center;
                border-bottom: 1px solid rgba(255,255,255,0.08);
                margin-bottom: 28px;">
      {LOGO} style="width: 200px; border-radius: 10px;
                    box-shadow: 0 6px 24px rgba(0,0,0,0.4);"/>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<p style="font-size:.68rem;letter-spacing:2px;color:rgba(255,255,255,.4);margin:0 0 14px;text-transform:uppercase;font-weight:700;">Bölge Seçin</p>', unsafe_allow_html=True)
    sel_bolge = st.selectbox("_bolge", BOLGELER, index=_db,
                              format_func=lambda x: BOLGE_FULL.get(x, x),
                              label_visibility="collapsed")

    st.markdown('<p style="font-size:.68rem;letter-spacing:2px;color:rgba(255,255,255,.4);margin:16px 0 14px;text-transform:uppercase;font-weight:700;">Ebat Seçin</p>', unsafe_allow_html=True)
    sel_ebat  = st.selectbox("_ebat", EBAT_COLS, index=_de,
                              label_visibility="collapsed")

    bolge_ad = BOLGE_FULL.get(sel_bolge, sel_bolge)
    st.markdown(f"""
    <div style="margin-top:36px; padding:18px 16px; background:rgba(255,255,255,0.05);
                border-radius:12px; border:1px solid rgba(255,255,255,0.08);">
      <p style="margin:0 0 10px; font-size:.68rem; letter-spacing:2px;
                color:rgba(255,255,255,.4); text-transform:uppercase; font-weight:700;">
        Seçili Filtre
      </p>
      <p style="margin:0 0 4px; font-size:.95rem; font-weight:700; color:white;">
        {bolge_ad}
      </p>
      <p style="margin:0; font-size:1.05rem; font-weight:800; color:#E05C2A;">
        {sel_ebat}
      </p>
    </div>
    <div style="margin-top:20px; padding:16px; background:rgba(255,255,255,0.04);
                border-radius:12px; border:1px solid rgba(255,255,255,0.07);">
      <p style="margin:0; font-size:.72rem; color:rgba(255,255,255,.35); line-height:2;">
        📍 7 Bölge · 33 Ebat<br>
        📊 231 Zaman Serisi<br>
        📅 2021 – 2025 Geçmiş<br>
        <span style="color:#E05C2A; font-weight:600;">🔮 2026 Projeksiyonu</span>
      </p>
    </div>
    """, unsafe_allow_html=True)

# ── Veri filtrele ─────────────────────────────────────────────────────────────
mask   = (df_raw["Bolge"] == sel_bolge) & (df_raw["Ebat"] == sel_ebat)
seri   = df_raw[mask].sort_values("Tarih")
train  = seri[seri["Tarih"] <= "2024-12-31"]["Satis"]
test_s = seri[seri["Tarih"] >= "2025-01-01"]

fc_seri, fc_col = pd.DataFrame(), "Ensemble"
if has_fc and df_fc is not None:
    mfc = (df_fc["Bolge"] == sel_bolge) & (df_fc["Ebat"] == sel_ebat)
    fc_seri = df_fc[mfc].sort_values("Tarih")
    fc_col  = "Ensemble" if "Ensemble" in df_fc.columns else df_fc.columns[-1]

toplam_2025 = float(test_s["Satis"].sum())
ort_2021_24 = float(train.mean()) if len(train) > 0 else 0.0
max_ay_str  = (seri.loc[seri["Satis"].idxmax(),"Tarih"].strftime("%b %Y")
               if len(seri) > 0 else "—")
has_tahmin  = has_fc and len(fc_seri) > 0

tahmin_2026 = tahmin_aylik = delta_pct = 0.0
if has_tahmin:
    tahmin_2026  = float(fc_seri[fc_col].sum())
    tahmin_aylik = float(fc_seri[fc_col].mean())
    delta_pct    = (tahmin_2026 - toplam_2025) / toplam_2025 * 100 if toplam_2025 > 0 else 0.0

# ── BAŞLIK BAR ────────────────────────────────────────────────────────────────
d_clr  = "#10B981" if delta_pct >= 0 else "#EF4444"
d_arr  = "▲" if delta_pct >= 0 else "▼"
d_sign = "+" if delta_pct >= 0 else ""

st.markdown(f"""
<div style="background:white; border-radius:0 0 20px 20px; padding:16px 20px;
            border-bottom:3px solid #E05C2A;
            box-shadow:0 2px 16px rgba(13,27,42,0.08);
            display:flex; align-items:center; justify-content:space-between;
            flex-wrap:wrap; gap:10px;
            margin-bottom:20px;">
  <div style="min-width:0; flex:1;">
    <p style="margin:0; font-size:clamp(.55rem,.9vw,.7rem); font-weight:700; letter-spacing:2px;
              color:#94A3B8; text-transform:uppercase; white-space:nowrap;">
      Talep Tahmin Sistemi · Karo
    </p>
    <h2 style="margin:4px 0 0; font-size:clamp(1rem,2.5vw,1.45rem); font-weight:800;
               color:#0D1B2A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
      {bolge_ad}
      <span style="color:#E05C2A; margin:0 6px;">·</span>
      {sel_ebat}
    </h2>
  </div>
  {LOGO} style="height:clamp(44px,6vw,72px); border-radius:8px; object-fit:contain; flex-shrink:0;"/>
</div>
""", unsafe_allow_html=True)

# ── KPI SATIRI ────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns([1.6, 1, 1, 1])

# Hero KPI
with c1:
    if has_tahmin:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg,#0D1B2A 0%,#1A3560 100%);
                    border-radius:16px; padding:18px 20px;
                    box-shadow:0 4px 20px rgba(13,27,42,0.2);
                    border-left:4px solid #E05C2A; min-height:110px;
                    display:flex; flex-direction:column; justify-content:center;">
          <p style="margin:0; font-size:clamp(.55rem,.8vw,.65rem); font-weight:700; letter-spacing:2px;
                    color:rgba(255,255,255,.5); text-transform:uppercase;">
            2026 Tahmini Yıllık Satış (m²)
          </p>
          <p style="margin:6px 0 4px; font-size:clamp(1.4rem,3vw,2.4rem); font-weight:900;
                    color:white; line-height:1; letter-spacing:-1px;">
            {tahmin_2026:,.0f}
          </p>
          <p style="margin:0; font-size:clamp(.7rem,1vw,.8rem); color:rgba(255,255,255,.55);">
            Aylık ort.&nbsp;<b style="color:white;">{tahmin_aylik:,.0f} m²</b>
            &ensp;
            <span style="color:{d_clr}; font-weight:700;">{d_arr}&nbsp;%{abs(delta_pct):.1f}</span>
            <span style="color:rgba(255,255,255,.35);">&nbsp;vs 2025</span>
          </p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="background:#F8FAFC; border-radius:16px; padding:24px 28px;
                    border:2px dashed #CBD5E1; min-height:110px;
                    display:flex; align-items:center; justify-content:center;">
          <p style="margin:0; color:#94A3B8; font-size:.9rem;">Tahmin verisi yok</p>
        </div>""", unsafe_allow_html=True)

with c2:
    if has_tahmin:
        st.markdown(f"""
        <div style="background:white;border-radius:16px;padding:18px 16px;
                    border:1px solid #EAECF0;box-shadow:0 1px 4px rgba(0,0,0,0.06);
                    min-height:110px;display:flex;flex-direction:column;justify-content:center;">
          <p style="margin:0 0 4px;font-size:clamp(.55rem,.8vw,.68rem);font-weight:700;letter-spacing:1.5px;
                    color:#94A3B8;text-transform:uppercase;">2025 Gerçek (m²)</p>
          <p style="margin:0 0 8px;font-size:clamp(1.1rem,2vw,1.8rem);font-weight:800;color:#0D1B2A;line-height:1.1;">
            {toplam_2025:,.0f}
          </p>
          <p style="margin:0;font-size:.82rem;font-weight:600;color:{d_clr};">
            <span style="font-size:1.15rem;margin-right:4px;">&#10132;</span>
            {d_sign}%{abs(delta_pct):.1f} · 2026 tahmini
          </p>
        </div>""", unsafe_allow_html=True)
    else:
        st.metric("2025 Gerçek (m²)", f"{toplam_2025:,.0f}")
with c3:
    st.metric("Aylık Ort. 2021–24 (m²)", f"{ort_2021_24:,.0f}")
with c4:
    st.metric("En Yüksek Ay", max_ay_str)

st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

# ── SEKMELER ──────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📅  Aylık Tahmin",
    "📈  Geçmiş & Trend",
    "🗺️  Bölge Haritası",
    "📋  Tahmin Kalitesi",
])

# ─── TAB 1: AYLIK TAHMİN ──────────────────────────────────────────────────────
with tab1:
    if not has_tahmin:
        st.info("Tahmin verisi bulunamadı. Modeli çalıştırıp Excel'i üretin.")
    else:
        fc26 = fc_seri.copy()
        fc26["AyNo"] = fc26["Tarih"].dt.month
        fc26["AyAd"] = fc26["AyNo"].map(AY_TR)
        vals26  = [float(v) for v in fc26[fc_col].fillna(0).values]
        aylar   = list(fc26["AyAd"])
        etiket  = [str(int(round(v))) for v in vals26]

        te_ay = test_s.copy()
        te_ay["AyNo"] = te_ay["Tarih"].dt.month
        te_ay["AyAd"] = te_ay["AyNo"].map(AY_TR)
        vals25 = [float(v) for v in te_ay["Satis"].fillna(0).values]

        # GARCH güven bandı
        gu1, gl1 = garch_band(tuple(seri["Satis"].values),
                               tuple(fc_seri[fc_col].values), len(fc_seri))
        has_garch = gu1 is not None and len(gu1) == len(vals26)

        mx = max(vals26) if max(vals26) > 0 else 1
        clrs = []
        for v in vals26:
            t = v / mx
            clrs.append(f"rgb({int(13+(224-13)*t)},{int(27+(92-27)*(1-t))},42)")

        fig = go.Figure()

        err_kw = {}
        if has_garch:
            err_kw = dict(error_y=dict(
                type="data",
                array=[float(u - v) for u, v in zip(gu1, vals26)],
                arrayminus=[float(max(v - l, 0)) for v, l in zip(vals26, gl1)],
                visible=True, color="#94A3B8", thickness=1.5, width=5,
            ))

        fig.add_trace(go.Bar(
            x=aylar, y=vals26, name="2026 Tahmini",
            marker=dict(color=clrs, line=dict(color="rgba(13,27,42,0.2)", width=0.5)),
            text=etiket, textposition="outside",
            **err_kw,
        ))
        fig.update_layout(
            barmode="group", height=420, bargap=0.22,
            paper_bgcolor="white", plot_bgcolor="white",
            legend=dict(orientation="h", y=1.04, x=0),
            yaxis=dict(title="Satış (m²)", gridcolor="#F1F5F9", zeroline=False),
            xaxis=dict(showgrid=False),
            margin=dict(l=50, r=20, t=55, b=10),
            hovermode="x unified",
        )
        st.markdown('<div style="background:white;border-radius:18px;padding:20px 20px 8px;box-shadow:0 1px 6px rgba(0,0,0,0.06);">', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # GARCH stok yorum kartı
        if has_garch:
            min_stok = float(min(gl1))
            max_stok = float(max(gu1))
            yillik_min = float(sum(gl1))
            yillik_max = float(sum(gu1))
            st.markdown(f"""
            <div style="margin-top:16px; display:flex; gap:16px;">
              <div style="flex:1; background:#EFF6FF; border-radius:14px; padding:18px 20px;
                          border-left:4px solid #1B4F8A;">
                <p style="margin:0 0 4px; font-size:.68rem; font-weight:700; letter-spacing:2px;
                           color:#1B4F8A; text-transform:uppercase;">📦 Minimum Stok Tavsiyesi</p>
                <p style="margin:0; font-size:1.4rem; font-weight:800; color:#0D1B2A;">
                  {min_stok:,.0f} m² <span style="font-size:.8rem;font-weight:500;color:#64748B;">/ ay</span>
                </p>
                <p style="margin:6px 0 0; font-size:.78rem; color:#475569;">
                  Yıllık alt sınır: <b>{yillik_min:,.0f} m²</b> — düşük talep senaryosu
                </p>
              </div>
              <div style="flex:1; background:#FFF7ED; border-radius:14px; padding:18px 20px;
                          border-left:4px solid #E05C2A;">
                <p style="margin:0 0 4px; font-size:.68rem; font-weight:700; letter-spacing:2px;
                           color:#E05C2A; text-transform:uppercase;">🏭 Maksimum Stok Tavsiyesi</p>
                <p style="margin:0; font-size:1.4rem; font-weight:800; color:#0D1B2A;">
                  {max_stok:,.0f} m² <span style="font-size:.8rem;font-weight:500;color:#64748B;">/ ay</span>
                </p>
                <p style="margin:6px 0 0; font-size:.78rem; color:#475569;">
                  Yıllık üst sınır: <b>{yillik_max:,.0f} m²</b> — yüksek talep senaryosu
                </p>
              </div>
            </div>
            """, unsafe_allow_html=True)

        # Sade 2026 tablosu — 3 sütun yan yana
        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
        max_v = max(vals26) if max(vals26) > 0 else 1
        rows_html = ""
        for ay, v in zip(aylar, vals26):
            bar_w = int(v / max_v * 100)
            rows_html += f"""
            <tr>
              <td style="padding:10px 16px;font-weight:600;color:#0D1B2A;width:90px;">{ay}</td>
              <td style="padding:10px 16px;font-size:1.05rem;font-weight:700;color:#E05C2A;width:110px;text-align:right;">{int(round(v)):,} m²</td>
              <td style="padding:10px 20px 10px 8px;">
                <div style="background:#EFF6FF;border-radius:6px;height:10px;width:100%;">
                  <div style="background:linear-gradient(90deg,#1B4F8A,#E05C2A);border-radius:6px;height:10px;width:{bar_w}%;"></div>
                </div>
              </td>
            </tr>"""
        st.markdown(f"""
        <div style="background:white;border-radius:18px;padding:16px 4px;
                    box-shadow:0 1px 6px rgba(0,0,0,0.06);overflow:hidden;">
          <p style="margin:0 0 12px 20px;font-size:.68rem;font-weight:700;
                    letter-spacing:2px;color:#94A3B8;text-transform:uppercase;">
            2026 · Aylık Satış Tahmini
          </p>
          <table style="width:100%;border-collapse:collapse;font-family:Inter,sans-serif;">
            <tbody>{rows_html}</tbody>
          </table>
        </div>
        """, unsafe_allow_html=True)

# ─── TAB 2: TARİHSEL TREND ────────────────────────────────────────────────────
with tab2:
    fig2 = go.Figure()
    tr = seri[seri["Tarih"] <= "2024-12-31"]
    te = seri[seri["Tarih"] >= "2025-01-01"]

    fig2.add_trace(go.Scatter(
        x=tr["Tarih"], y=tr["Satis"], name="Geçmiş (2021–2024)",
        mode="lines", line=dict(color="#CBD5E1", width=1.8),
    ))
    fig2.add_trace(go.Scatter(
        x=te["Tarih"], y=te["Satis"], name="Gerçek (2025)",
        mode="lines+markers", line=dict(color="#1B4F8A", width=2.5),
        marker=dict(size=7, color="#1B4F8A", line=dict(color="white", width=1.5)),
    ))

    if has_tahmin:
        gu, gl = garch_band(tuple(seri["Satis"].values),
                            tuple(fc_seri[fc_col].values), len(fc_seri))
        if gu is not None:
            bx = list(fc_seri["Tarih"]) + list(fc_seri["Tarih"])[::-1]
            by = list(gu) + list(gl)[::-1]
            fig2.add_trace(go.Scatter(x=bx, y=by, fill="toself",
                fillcolor="rgba(224,92,42,0.10)",
                line=dict(color="rgba(0,0,0,0)"),
                name="GARCH %95 Bandı", hoverinfo="skip"))
            for arr in [gu, gl]:
                fig2.add_trace(go.Scatter(
                    x=fc_seri["Tarih"], y=arr, mode="lines",
                    line=dict(color="rgba(224,92,42,0.35)", width=1, dash="dot"),
                    showlegend=False, hoverinfo="skip"))
        fig2.add_trace(go.Scatter(
            x=fc_seri["Tarih"], y=fc_seri[fc_col],
            name="2026 Tahmini", mode="lines+markers",
            line=dict(color="#E05C2A", width=3),
            marker=dict(size=9, symbol="diamond", color="#E05C2A",
                        line=dict(color="white", width=2)),
        ))

    fig2.add_vline(x=pd.Timestamp("2025-01-01").timestamp()*1000,
                   line_dash="dot", line_color="#CBD5E1", annotation_text="2025")
    if has_fc:
        fig2.add_vline(x=pd.Timestamp("2026-01-01").timestamp()*1000,
                       line_dash="dash", line_color="#E05C2A", annotation_text="Tahmin")

    fig2.update_layout(
        height=450, paper_bgcolor="white", plot_bgcolor="white",
        legend=dict(orientation="h", y=-0.18),
        yaxis=dict(title="Satış (m²)", gridcolor="#F1F5F9", zeroline=False),
        xaxis=dict(gridcolor="#F1F5F9"),
        margin=dict(l=50, r=20, t=20, b=70),
        hovermode="x unified",
    )
    st.markdown('<div style="background:white;border-radius:18px;padding:20px;box-shadow:0 1px 6px rgba(0,0,0,0.06);">', unsafe_allow_html=True)
    st.plotly_chart(fig2, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ─── TAB 3: TÜRKİYE HARİTASI ─────────────────────────────────────────────────
with tab3:
    if not (has_fc and df_fc is not None):
        st.info("Harita için önce modeli çalıştırın.")
    else:
        bt = (df_fc.groupby("Bolge")[fc_col].sum()
              .reset_index().rename(columns={fc_col:"T2026"}))
        bt["Lat"] = bt["Bolge"].map(lambda b: BOLGE_COORDS.get(b,(39,35))[0])
        bt["Lon"] = bt["Bolge"].map(lambda b: BOLGE_COORDS.get(b,(39,35))[1])
        bt["Ad"]  = bt["Bolge"].map(lambda b: BOLGE_FULL.get(b,b))
        mx = bt["T2026"].max()

        fig3 = go.Figure()
        for _, r in bt.iterrows():
            sec   = r["Bolge"] == sel_bolge
            renk  = "#E05C2A" if sec else "#1B4F8A"
            boyut = 18 + 46 * (r["T2026"] / mx)
            fig3.add_trace(go.Scattergeo(
                lat=[r["Lat"]], lon=[r["Lon"]],
                mode="markers+text",
                marker=dict(size=boyut, color=renk, opacity=0.85,
                            line=dict(color="white", width=2.5 if sec else 1.5)),
                text=[f"{r['Ad']}\n{r['T2026']:,.0f} m²"],
                textposition="top center",
                textfont=dict(size=10),
                showlegend=False,
                hovertemplate=f"<b>{r['Ad']}</b><br>2026: <b>{r['T2026']:,.0f} m²</b><extra></extra>",
            ))
        fig3.update_layout(
            geo=dict(scope="asia", center=dict(lat=39,lon=35), projection_scale=4.5,
                     showland=True, landcolor="#F1F5F9",
                     showocean=True, oceancolor="#DBEAFE",
                     showcoastlines=True, coastlinecolor="#CBD5E1", showframe=False),
            height=460, margin=dict(l=0,r=0,t=10,b=0), paper_bgcolor="white",
        )
        st.caption("🟠 Seçili bölge  ·  🔵 Diğer bölgeler  ·  Daire = tahmin hacmi")
        st.markdown('<div style="background:white;border-radius:18px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,0.06);">', unsafe_allow_html=True)
        st.plotly_chart(fig3, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ─── TAB 4: TAHMİN KALİTESİ ──────────────────────────────────────────────────
with tab4:
    if not (has_fc and df_metrics is not None):
        st.info("Tahmin kalite verisi yok.")
    else:
        try:
            q = df_metrics.copy()
            q["Bölge"] = q["Bolge"].map(lambda b: BOLGE_FULL.get(b, b))
            q["MASE"]  = pd.to_numeric(q["En_Iyi_MASE"], errors="coerce")

            def guven_etiketi(v):
                if v < 0.8:  return "🟢 Çok İyi"
                if v < 1.0:  return "🟡 İyi"
                if v < 1.5:  return "🟠 Orta"
                return               "🔴 Zayıf"

            q["Tahmin Güvenilirliği"] = q["MASE"].apply(guven_etiketi)

            # Özet sayaçlar
            cats = q["Tahmin Güvenilirliği"].value_counts()
            c1, c2, c3, c4 = st.columns(4)
            col_map = {
                "🟢 Çok İyi": (c1, "#D1FAE5", "#065F46"),
                "🟡 İyi":     (c2, "#FEF3C7", "#92400E"),
                "🟠 Orta":    (c3, "#FFEDD5", "#9A3412"),
                "🔴 Zayıf":   (c4, "#FEE2E2", "#7F1D1D"),
            }
            for label, (col, bg, fg) in col_map.items():
                cnt = int(cats.get(label, 0))
                with col:
                    st.markdown(f"""
                    <div style="background:{bg};border-radius:14px;padding:18px 16px;text-align:center;">
                      <p style="margin:0;font-size:1.7rem;font-weight:900;color:{fg};">{cnt}</p>
                      <p style="margin:4px 0 0;font-size:.72rem;font-weight:700;color:{fg};letter-spacing:.5px;">{label}</p>
                    </div>""", unsafe_allow_html=True)

            st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

            # Seri tablosu — sadece anlaşılır kolonlar
            goster = (q.sort_values("MASE")
                       [["Bölge", "Ebat", "Tahmin Güvenilirliği"]]
                       .reset_index(drop=True))
            goster.index = goster.index + 1
            st.caption("Her satır bir bölge–ebat kombinasyonuna ait tahmin güvenilirliğini gösterir.")
            st.dataframe(goster, use_container_width=True)
        except Exception as e:
            st.info(f"Tahmin kalite verisi gösterilemiyor. ({e})")

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="margin-top:40px; padding:18px 32px; background:#0D1B2A;
            border-radius:16px; display:flex; align-items:center;
            justify-content:space-between;">
  {LOGO} style="height:48px; border-radius:7px; object-fit:contain;"/>
  <p style="margin:0; font-size:.72rem; color:rgba(255,255,255,.3);">
    Eczacıbaşı VitrA Karo · Talep Tahmin Sistemi · 2026 Projeksiyonu
  </p>
</div>
""", unsafe_allow_html=True)
