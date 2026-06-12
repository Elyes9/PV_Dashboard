"""
PROFITPV Dashboard — Solar PV Installation Analyzer
Based on: PROFITPV_MT CRA2E v0.2 TLS (real data extracted from Excel)
Libraries: streamlit, pandas, numpy, altair only
Run: streamlit run profitpv_dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="PROFITPV · Analyse PV",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'Syne', sans-serif; }
.stApp { background:#0d1117; color:#e8eaf0; }
[data-testid="stSidebar"] { background:#111820; border-right:1px solid #1e2d3d; }
h1,h2,h3 { font-family:'Syne',sans-serif; font-weight:800; }
.kpi { background:linear-gradient(135deg,#111820,#0d1117); border:1px solid #1e2d3d;
       border-radius:12px; padding:18px 22px; margin-bottom:10px; position:relative; overflow:hidden; }
.kpi::before { content:''; position:absolute; top:0; left:0; width:4px; height:100%;
               background:linear-gradient(180deg,#f5a623,#e67e22); border-radius:4px 0 0 4px; }
.kpi-label { font-family:'DM Mono',monospace; font-size:10px; color:#5a7a9a;
             text-transform:uppercase; letter-spacing:1.5px; margin-bottom:5px; }
.kpi-value { font-size:26px; font-weight:800; color:#f5a623; line-height:1.1; }
.kpi-unit  { font-family:'DM Mono',monospace; font-size:11px; color:#5a7a9a; margin-top:3px; }
.kpi-green .kpi-value { color:#2ecc71; }
.kpi-green::before { background:linear-gradient(180deg,#2ecc71,#27ae60); }
.kpi-blue .kpi-value { color:#3498db; }
.kpi-blue::before { background:linear-gradient(180deg,#3498db,#2980b9); }
.kpi-red .kpi-value { color:#e74c3c; }
.kpi-red::before { background:linear-gradient(180deg,#e74c3c,#c0392b); }
.sec { font-size:11px; font-weight:700; letter-spacing:2px; text-transform:uppercase;
       color:#f5a623; margin:22px 0 10px; padding-bottom:6px; border-bottom:1px solid #1e2d3d; }
.info { background:#111820; border:1px solid #1e2d3d; border-radius:8px;
        padding:12px 16px; font-size:13px; color:#8aa4be; margin-bottom:8px; }
.info strong { color:#c5d8ea; }
.badge-ok   { background:#0d2b1a; border:1px solid #2ecc71; border-radius:10px; padding:16px; text-align:center; }
.badge-warn { background:#2b1f0d; border:1px solid #f5a623; border-radius:10px; padding:16px; text-align:center; }
.badge-bad  { background:#2b0d0d; border:1px solid #e74c3c; border-radius:10px; padding:16px; text-align:center; }
.excel-tag { background:#1e2d3d; border-radius:4px; padding:2px 7px;
             font-family:'DM Mono',monospace; font-size:10px; color:#f5a623; margin-left:6px; }
.footer { text-align:center; color:#2a4a6a; font-family:'DM Mono',monospace;
          font-size:11px; padding:10px 0; margin-top:20px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ALTAIR DARK THEME
# ─────────────────────────────────────────────
def pv_theme():
    return {
        "config": {
            "background": "#111820",
            "view": {"stroke": "transparent", "fill": "#111820"},
            "axis": {
                "gridColor": "#1e2d3d", "domainColor": "#1e2d3d",
                "tickColor": "#1e2d3d", "labelColor": "#8aa4be",
                "titleColor": "#8aa4be", "labelFont": "DM Mono",
                "titleFont": "Syne", "labelFontSize": 11, "titleFontSize": 12,
            },
            "legend": {
                "labelColor": "#8aa4be", "titleColor": "#8aa4be",
                "fillColor": "#111820", "strokeColor": "#1e2d3d",
                "labelFont": "DM Mono", "titleFont": "Syne",
            },
            "title": {"color": "#8aa4be", "font": "Syne", "fontSize": 13, "anchor": "start"},
            "mark": {"color": "#f5a623"},
            "range": {
                "category": ["#f5a623", "#2ecc71", "#3498db", "#e74c3c", "#8e44ad", "#1abc9c"],
            },
        }
    }
alt.themes.register("pv_theme", pv_theme)
alt.themes.enable("pv_theme")

# ─────────────────────────────────────────────
# REAL EXCEL VALUES (pre-loaded as defaults)
# ─────────────────────────────────────────────
# From sheet "Résultats" and "Entrées Utilisateur" — PROFITPV_MT CRA2E v0.2 TLS
EXCEL = {
    "kwc": 197.21,
    "prix_kwc": 2200.0,
    "cout_ht": 433862.0,
    "tva": 56402.06,
    "cout_ttc": 490264.06,
    "sub_dt": 0.0,
    "rendement": 1659.09,
    "degradation": 0.004,
    "region": "Monastir",
    "orientation": "Sud",
    "conso": 315628.0,
    "prod_y1": 327189.14,
    "autoconso_y1": 208812.32,
    "cedee_y1": 118376.82,
    "taux_couv": 0.6616,
    "facture_sans": 93741.52,
    "facture_avec": 31724.26,
    "vente_exced": 7852.54,
    "payback": 6.89,
    "tri_projet": 0.15168,
    "tri_equity": 0.19312,
    "van": 282804.74,
    "lcoe": 0.20275,
    "co2_total": 3729.96,
    "tarif_kwh": 0.291,
    "surtaxe": 0.006,
    "tarif_achat": 0.08,
    "tarif_transport": 0.039,
    "limite_exced": 0.30,
    "om_pct": 0.04,
    "om_y1": 17354.48,
    "montant_dette": 303703.40,
    "cap_propres": 130158.60,
    "taux_interet": 0.10,
    "taux_disc": 0.08,
    "duree_credit": 10,
    "grace": 1,
    "hausse_tarif": 0.05,
    "inflation": 0.0102,
    "duree": 20,
    "marge": 0.0,
    # Real yearly revenues from Flux de trésorerie sheet
    "rev_years": [0, 69869.80, 73069.84, 76416.43, 79916.31, 83129.19,
                  86491.01, 90008.58, 93689.03, 97539.82, 101568.73,
                  105783.92, 110193.91, 114807.62, 119634.37, 124683.91,
                  129966.44, 135492.62, 141273.61, 147321.06, 153647.18],
    # Real yearly opex from Flux de trésorerie sheet
    "opex_years": [0, 17354.48, 17701.57, 18055.60, 18416.71, 18785.05,
                   19160.75, 19543.96, 19934.84, 20333.54, 20740.21,
                   21155.01, 21578.11, 22009.68, 22449.87, 22898.87,
                   23356.85, 23823.98, 24300.46, 24786.47, 25282.20],
    # Real debt service (interest + principal) from sheet
    "ds_years": [0, 30370.34, 52735.22, 52735.22, 52735.22, 52735.22,
                 52735.22, 52735.22, 52735.22, 52735.22, 52735.22,
                 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    # CFADS from sheet (net operating cashflow)
    "cfads_years": [0, 52515.32, 55368.27, 58360.83, 61499.59, 64344.14,
                    67330.26, 70464.62, 73754.19, 77206.28, 80828.52,
                    84628.91, 88615.80, 92797.94, 97184.50, 101785.04,
                    106609.60, 111668.64, 116973.15, 122534.59, 128364.98],
    # Equity cashflow from sheet
    "eq_cf_years": [0, 22144.98, 2633.04, 5625.61, 8764.37, 11608.92,
                    14595.04, 17729.39, 21018.97, 24471.06, 28093.30,
                    84628.91, 88615.80, 92797.94, 97184.50, 101785.04,
                    106609.60, 111668.64, 116973.15, 122534.59, 128364.98],
    "dscr_min": 1.0499,
    "llcr_min": 1.2461,
}

# ─────────────────────────────────────────────
# SIDEBAR INPUTS
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ☀️ PROFITPV")
    st.markdown("<div style='color:#5a7a9a;font-size:12px;margin-bottom:4px;font-family:DM Mono,monospace;'>CRA2E · MT · v0.2 TLS</div>", unsafe_allow_html=True)
    st.markdown("<div style='color:#2ecc71;font-size:11px;margin-bottom:18px;font-family:DM Mono,monospace;'>✓ Données Excel chargées</div>", unsafe_allow_html=True)

    use_excel = st.toggle("Utiliser les valeurs Excel réelles", value=True)

    st.markdown('<div class="sec">🔧 Installation PV</div>', unsafe_allow_html=True)
    kwc        = st.number_input("Taille installation (kWc)", 1.0, 5000.0, EXCEL["kwc"], 1.0)
    prix_kwc   = st.number_input("Prix unitaire HT (DT/kWc)", 500.0, 5000.0, EXCEL["prix_kwc"], 50.0)
    subvention = st.selectbox("Subvention FTE ?", ["Non", "Oui"])
    sub_dt     = st.number_input("Montant subvention (DT)", 0.0, value=EXCEL["sub_dt"], step=1000.0) if subvention == "Oui" else 0.0
    tva_pct    = st.slider("TVA (%)", 0, 20, 13)   # 56402/433862 ≈ 13%

    st.markdown('<div class="sec">📍 Site & Panneaux</div>', unsafe_allow_html=True)
    region      = st.selectbox("Région", ["Monastir","Tunis","Sfax","Bizerte","Gafsa","Jendouba","Jerba","Kairouan","Tataouine","Tozeur"], index=0)
    orientation = st.selectbox("Orientation", ["Sud","Sud-ouest"])
    yield_map   = {"Monastir":1659,"Sfax":1680,"Gafsa":1720,"Jerba":1700,"Kairouan":1690,
                   "Tataouine":1730,"Tozeur":1740,"Bizerte":1580,"Jendouba":1560,"Tunis":1600}
    rendement   = st.number_input("Rendement spécifique (kWh/kWc/an)", 1000, 2000, int(EXCEL["rendement"]))
    degradation = st.slider("Dégradation annuelle (%)", 0.0, 1.5, EXCEL["degradation"]*100, 0.1) / 100

    st.markdown('<div class="sec">⚡ Consommation</div>', unsafe_allow_html=True)
    conso        = st.number_input("Consommation annuelle (kWh/an)", 1000.0, 5e6, EXCEL["conso"], 1000.0)
    tarif_kwh    = st.number_input("Tarif uniforme STEG (DT/kWh)", 0.1, 1.0, EXCEL["tarif_kwh"], 0.001, format="%.3f")
    tarif_j      = st.number_input("Tarif jour (DT/kWh)", 0.1, 1.0, 0.290, 0.001, format="%.3f")
    tarif_pointe = st.number_input("Tarif pointe soir (DT/kWh)", 0.1, 1.0, 0.377, 0.001, format="%.3f")
    tarif_nuit   = st.number_input("Tarif nuit/dim (DT/kWh)", 0.1, 1.0, 0.222, 0.001, format="%.3f")
    surtaxe      = st.number_input("Surtaxe municipale (DT/kWh)", 0.0, 0.05, EXCEL["surtaxe"], 0.001, format="%.3f")
    limite_exced = st.slider("Limite excédent annuel (%)", 10, 50, int(EXCEL["limite_exced"]*100)) / 100
    tarif_achat  = st.number_input("Tarif achat excédent (DT/kWh)", 0.01, 0.30, EXCEL["tarif_achat"], 0.001, format="%.3f")
    tarif_transport = st.number_input("Tarif transport (DT/kWh)", 0.01, 0.20, EXCEL["tarif_transport"], 0.001, format="%.3f")

    st.markdown('<div class="sec">💰 Financement</div>', unsafe_allow_html=True)
    dette_pct    = st.slider("Part dette (%)", 0, 100, 70) / 100
    duree_credit = st.slider("Durée crédit (ans)", 1, 20, EXCEL["duree_credit"])
    grace        = st.slider("Délai de grâce (ans)", 0, 5, EXCEL["grace"])
    taux_ref     = st.number_input("Taux référence BCT (%)", 0.0, 30.0, 10.0, 0.5) / 100
    marge        = st.number_input("Marge bancaire (%)", 0.0, 10.0, 0.0, 0.25) / 100
    taux_disc    = st.number_input("Taux actualisation (%)", 1.0, 20.0, EXCEL["taux_disc"]*100, 0.5) / 100

    st.markdown('<div class="sec">📈 Hypothèses Prix</div>', unsafe_allow_html=True)
    hausse_tarif = st.slider("Hausse tarif STEG après Y2 (%/an)", 0.0, 10.0, EXCEL["hausse_tarif"]*100, 0.5) / 100
    inflation    = st.number_input("Inflation (%/an)", 0.5, 15.0, EXCEL["inflation"]*100, 0.1) / 100
    om_pct       = st.slider("O&M (% investissement/an)", 1, 10, int(EXCEL["om_pct"]*100)) / 100
    duree        = st.slider("Durée du projet (ans)", 10, 30, EXCEL["duree"])

# ─────────────────────────────────────────────
# CALCULATIONS — use real Excel data if toggled
# ─────────────────────────────────────────────
tva_frac = tva_pct / 100
cout_ht  = kwc * prix_kwc
cout_ttc = cout_ht * (1 + tva_frac) - sub_dt
cout_net = max(cout_ttc, 0)

prod_y1      = kwc * rendement
taux_couv    = min(prod_y1 / conso, 1.0)
autoconso_y1 = prod_y1 * min(taux_couv, 0.85)
cedee_y1     = max(0, prod_y1 - autoconso_y1)
cedee_lim_y1 = min(cedee_y1, prod_y1 * limite_exced)

facture_sans  = conso * (tarif_kwh + surtaxe)
eco_autoconso = autoconso_y1 * (tarif_kwh + surtaxe)
rev_exced_y1  = cedee_lim_y1 * tarif_achat
facture_avec  = facture_sans - eco_autoconso - rev_exced_y1
eco_annuelle  = facture_sans - facture_avec

montant_dette = min(cout_net * dette_pct, 200_000)
cap_propres   = cout_net - montant_dette
taux_interet  = taux_ref + marge

years = list(range(1, duree + 1))

if use_excel and duree == 20:
    # Use real numbers from the Excel file
    rev_list  = EXCEL["rev_years"][1:]
    opex_list = EXCEL["opex_years"][1:]
    ds_list   = EXCEL["ds_years"][1:]
    cfads_list= EXCEL["cfads_years"][1:]
    eq_cf_list= EXCEL["eq_cf_years"][1:]
    ncf_list  = cfads_list   # net operating CF = CFADS
    prod_list = [EXCEL["prod_y1"] * (1 - EXCEL["degradation"]) ** (y - 1) for y in years]
    cum_list  = list(np.cumsum(np.array(ncf_list) - EXCEL["cout_ht"]/20))   # simplified
    # Recompute proper cumulative from equity perspective
    inv = EXCEL["cap_propres"]
    cum_list = []
    running = -inv
    for ec in eq_cf_list:
        running += ec
        cum_list.append(running)
    # Override key indicators with real Excel values
    cout_net     = EXCEL["cout_ttc"]
    cout_ht      = EXCEL["cout_ht"]
    prod_y1      = EXCEL["prod_y1"]
    autoconso_y1 = EXCEL["autoconso_y1"]
    cedee_lim_y1 = EXCEL["cedee_y1"]
    facture_sans = EXCEL["facture_sans"]
    facture_avec = EXCEL["facture_avec"]
    eco_annuelle = facture_sans - facture_avec
    montant_dette= EXCEL["montant_dette"]
    cap_propres  = EXCEL["cap_propres"]
    taux_couv    = EXCEL["taux_couv"]
    project_irr  = EXCEL["tri_projet"]
    equity_irr   = EXCEL["tri_equity"]
    project_npv  = EXCEL["van"]
    lcoe         = EXCEL["lcoe"]
    payback_val  = EXCEL["payback"]
    dscr_min     = EXCEL["dscr_min"]
    llcr_min     = EXCEL["llcr_min"]
    rev_exced_y1 = EXCEL["vente_exced"]
    om_y1        = EXCEL["om_y1"]
else:
    # Compute from sidebar inputs
    prod_list, rev_list, opex_list, ds_list, ncf_list, cum_list = [], [], [], [], [], []
    cfads_list, eq_cf_list = [], []
    outstanding = montant_dette
    cum = -cap_propres

    for y in years:
        p  = prod_y1 * (1 - degradation) ** (y - 1)
        t  = tarif_kwh if y <= 2 else tarif_kwh * (1 + hausse_tarif) ** (y - 2)
        ac = p * min(conso / (p + 1e-9), 0.85)
        ce = min(max(0, p - ac), p * limite_exced)
        r  = ac * (t + surtaxe) + ce * tarif_achat * (1 + hausse_tarif) ** max(0, y - 2)
        o  = cout_net * om_pct * (1 + inflation) ** (y - 1)
        if y <= grace:
            ds = outstanding * taux_interet; principal = 0.0
        elif y <= duree_credit:
            principal = montant_dette / max(duree_credit - grace, 1)
            ds = outstanding * taux_interet + principal
            outstanding = max(0, outstanding - principal)
        else:
            ds = 0.0
        ncf = r - o
        eq_cf = ncf - ds
        cum += eq_cf
        prod_list.append(p); rev_list.append(r); opex_list.append(o)
        ds_list.append(ds); ncf_list.append(ncf); cum_list.append(cum)
        cfads_list.append(ncf); eq_cf_list.append(eq_cf)

    cf_arr = np.array([-cap_propres] + eq_cf_list)
    def calc_npv(rate, cfs):
        t = np.arange(len(cfs))
        return np.sum(cfs / (1 + rate) ** t)
    def calc_irr(cfs, lo=-0.5, hi=5.0, tol=1e-7, maxiter=200):
        for _ in range(maxiter):
            mid = (lo + hi) / 2
            if calc_npv(mid, cfs) > 0: lo = mid
            else: hi = mid
            if hi - lo < tol: break
        return (lo + hi) / 2

    project_irr  = calc_irr(np.array([-cout_net] + ncf_list))
    equity_irr   = calc_irr(cf_arr)
    project_npv  = calc_npv(taux_disc, np.array([-cout_net] + ncf_list))
    disc_costs   = cout_net + sum((cout_net * om_pct * (1 + inflation)**(y-1)) / (1+taux_disc)**y for y in range(1, duree+1))
    disc_prod    = sum((prod_y1 * (1-degradation)**(y-1)) / (1+taux_disc)**y for y in range(1, duree+1))
    lcoe         = disc_costs / disc_prod if disc_prod > 0 else 0
    payback_val  = next((i+1 for i,v in enumerate(cum_list) if v > 0), None) or duree+1
    om_y1        = cout_net * om_pct
    dscr_min     = 0.0; llcr_min = 0.0

irr_pct    = project_irr * 100
eq_irr_pct = equity_irr * 100 if "equity_irr" in dir() else EXCEL["tri_equity"] * 100
co2_per_year = [(EXCEL["prod_y1"] if use_excel and duree==20 else prod_y1) * (1 - (EXCEL["degradation"] if use_excel else degradation))**(y-1) * 0.57 / 1000 for y in years]
co2_total  = sum(co2_per_year)
co2_cum    = list(np.cumsum(co2_per_year))
total_prod = sum(prod_list)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
if irr_pct > 12:
    badge_cls, badge_color, badge_label = "badge-ok",   "#2ecc71", "TRÈS RENTABLE"
elif irr_pct > 8:
    badge_cls, badge_color, badge_label = "badge-ok",   "#2ecc71", "RENTABLE"
elif irr_pct > 5:
    badge_cls, badge_color, badge_label = "badge-warn", "#f5a623", "MARGINAL"
else:
    badge_cls, badge_color, badge_label = "badge-bad",  "#e74c3c", "DÉFICITAIRE"

col_title, col_badge = st.columns([4, 1])
with col_title:
    excel_tag = '<span class="excel-tag">EXCEL ✓</span>' if use_excel and duree==20 else ''
    st.markdown(f"<h1 style='color:#f5a623;margin-bottom:4px;'>☀️ Analyse Installation PV {excel_tag}</h1>", unsafe_allow_html=True)
    st.markdown(f"<div style='color:#5a7a9a;font-family:DM Mono,monospace;font-size:13px;'>"
                f"{kwc:.2f} kWc · {region} · {orientation} · {duree} ans · "
                f"Monastir 2025 · PROFITPV MT CRA2E v0.2</div>", unsafe_allow_html=True)
with col_badge:
    st.markdown(f"""<div class="{badge_cls}" style="margin-top:10px;">
        <div style='color:{badge_color};font-size:16px;font-weight:800;'>{badge_label}</div>
        <div style='color:{badge_color};font-family:DM Mono,monospace;font-size:13px;'>TRI = {irr_pct:.1f}%</div>
        <div style='color:{badge_color};font-family:DM Mono,monospace;font-size:11px;'>Equity = {eq_irr_pct:.1f}%</div>
    </div>""", unsafe_allow_html=True)

st.markdown("---")

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Résumé Financier",
    "⚡ Énergie & Production",
    "💵 Flux de Trésorerie",
    "🏗️ Investissement & Financement",
    "📐 Ratios Bancaires",
    "🌍 Impact Environnemental",
])

# ══════════════════════════════════
# TAB 1 — FINANCIAL SUMMARY
# ══════════════════════════════════
with tab1:
    st.markdown('<div class="sec">Indicateurs Clés de Rentabilité</div>', unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    kpi_data = [
        (c1, "TRI Projet",        f"{irr_pct:.2f}%",          "Taux rentabilité interne",    ""),
        (c2, "TRI Capitaux Propres", f"{eq_irr_pct:.2f}%",    "Equity IRR",                  "kpi-green"),
        (c3, "VAN Projet",        f"{project_npv/1000:+.0f}k DT", f"Taux disc. {taux_disc*100:.0f}%", "kpi-blue" if project_npv > 0 else "kpi-red"),
        (c4, "Temps de Retour",   f"{payback_val:.1f} ans",   "Retour sur investissement",   ""),
        (c5, "LCOE",              f"{lcoe:.4f} DT/kWh",       f"vs STEG {tarif_kwh:.3f} DT/kWh", "kpi-green" if lcoe < tarif_kwh else "kpi-red"),
    ]
    for col, label, val, unit, cls in kpi_data:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    c6, c7, c8, c9 = st.columns(4)
    reduction = (facture_sans - facture_avec) / facture_sans * 100 if facture_sans > 0 else 0
    kpi_data2 = [
        (c6, "Économie Année 1",    f"{eco_annuelle:,.0f} DT",      f"Facture: {facture_avec:,.0f} → {facture_sans:,.0f} DT", ""),
        (c7, "Réduction Facture",   f"{reduction:.1f}%",             "de la facture STEG",          "kpi-green"),
        (c8, f"Gain Total {duree}a",f"{sum(ncf_list)/1000:.0f}k DT","Flux nets cumulés opérations", ""),
        (c9, "Investissement Net",  f"{cout_net/1000:.0f}k DT",      "TTC après subvention",         "kpi-blue"),
    ]
    for col, label, val, unit, cls in kpi_data2:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="sec">Décomposition de la Facture STEG — Année 1</div>', unsafe_allow_html=True)

    col_wf, col_lcoe = st.columns([3, 2])
    with col_wf:
        wf_df = pd.DataFrame([
            {"catégorie": "Facture sans PV",      "base": 0,                              "val": facture_sans,   "type": "facture"},
            {"catégorie": "Économies autoconso",   "base": facture_sans - eco_autoconso,   "val": eco_autoconso,  "type": "économie"},
            {"catégorie": "Vente excédent STEG",   "base": facture_sans - eco_autoconso - rev_exced_y1, "val": rev_exced_y1, "type": "vente"},
            {"catégorie": "Facture avec PV",       "base": 0,                              "val": facture_avec,   "type": "résultat"},
        ])
        base_b = alt.Chart(wf_df).mark_bar(opacity=0).encode(
            x=alt.X("catégorie:N", sort=None, axis=alt.Axis(labelAngle=-10, title="")),
            y=alt.Y("base:Q", title="DT/an"),
        )
        color_b = alt.Chart(wf_df).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, size=50).encode(
            x=alt.X("catégorie:N", sort=None),
            y=alt.Y("base:Q"),
            y2=alt.Y2("top:Q"),
            color=alt.Color("type:N", scale=alt.Scale(
                domain=["facture","économie","vente","résultat"],
                range=["#5a7a9a","#2ecc71","#3498db","#f5a623"]
            ), legend=alt.Legend(title="Type")),
            tooltip=["catégorie:N", alt.Tooltip("montant:Q", format=",.0f", title="DT")],
        ).transform_calculate(top="datum.base + datum.val", montant="datum.val")
        label_b = alt.Chart(wf_df).mark_text(dy=-8, color="#e8eaf0", fontSize=12, fontWeight="bold", font="DM Mono").encode(
            x=alt.X("catégorie:N", sort=None),
            y=alt.Y("top2:Q"),
            text=alt.Text("montant2:Q", format=",.0f"),
        ).transform_calculate(top2="datum.base + datum.val", montant2="datum.val")
        st.altair_chart((base_b + color_b + label_b).properties(
            title="Décomposition de la facture STEG (DT/an)", height=320
        ), use_container_width=True)

    with col_lcoe:
        lcoe_df = pd.DataFrame([
            {"label": "LCOE projet",       "value": lcoe,                  "type": "lcoe"},
            {"label": "Tarif STEG 2025",   "value": tarif_kwh,             "type": "tarif"},
            {"label": "Tarif STEG +5%×10a","value": tarif_kwh*(1.05**10),  "type": "proj10"},
            {"label": "Tarif STEG +5%×20a","value": tarif_kwh*(1.05**20),  "type": "proj20"},
        ])
        lcoe_chart = alt.Chart(lcoe_df).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            x=alt.X("value:Q", title="DT/kWh"),
            y=alt.Y("label:N", sort=None, title=""),
            color=alt.Color("type:N", scale=alt.Scale(
                domain=["lcoe","tarif","proj10","proj20"],
                range=["#f5a623","#e74c3c","#8e44ad","#c0392b"]
            ), legend=None),
            tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=".4f", title="DT/kWh")],
        ).properties(title="LCOE vs Projections Tarif STEG", height=220)
        st.altair_chart(lcoe_chart, use_container_width=True)

        st.markdown(f'<div class="info"><strong>Vente excédent (Y1) :</strong> {rev_exced_y1:,.0f} DT</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="info"><strong>Remboursement STEG :</strong> 1 617,61 DT/an</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="info"><strong>Économies/kWh autoconsommé :</strong> {(tarif_kwh+surtaxe):.3f} DT/kWh</div>', unsafe_allow_html=True)


# ══════════════════════════════════
# TAB 2 — ENERGY & PRODUCTION
# ══════════════════════════════════
with tab2:
    st.markdown('<div class="sec">Indicateurs de Production</div>', unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    en_kpis = [
        (c1, "Production Y1",     f"{prod_y1:,.0f} kWh",       f"{kwc:.2f} kWc × {rendement} kWh/kWc",  ""),
        (c2, "Taux Couverture",   f"{taux_couv*100:.1f}%",      "conso couverte par PV",                  "kpi-green"),
        (c3, "Autoconsommée",     f"{autoconso_y1:,.0f} kWh",   "sur site Y1",                            ""),
        (c4, "Excédent STEG",     f"{cedee_lim_y1:,.0f} kWh",  f"limité à {limite_exced*100:.0f}%",       "kpi-blue"),
        (c5, "% Prod. cédée",     f"{cedee_lim_y1/prod_y1*100:.1f}%", "de la production totale",         ""),
    ]
    for col, label, val, unit, cls in en_kpis:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    # 20-year production bars
    df_prod = pd.DataFrame({
        "Année": years,
        "Production (kWh)": prod_list,
        "Revenue (DT)": rev_list,
    })
    prod_bars = alt.Chart(df_prod).mark_bar(
        color="#f5a623", opacity=0.82, cornerRadiusTopLeft=2, cornerRadiusTopRight=2
    ).encode(
        x=alt.X("Année:O", title="Année"),
        y=alt.Y("Production (kWh):Q", title="kWh/an"),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Production (kWh):Q", format=",.0f", title="Production kWh")],
    )
    conso_rule = alt.Chart(pd.DataFrame({"y": [conso]})).mark_rule(
        color="#e74c3c", strokeDash=[6,3], strokeWidth=2
    ).encode(y="y:Q")
    conso_label = alt.Chart(pd.DataFrame({"y": [conso], "label": [f"Conso: {conso:,.0f} kWh/an"]})).mark_text(
        align="left", dx=6, dy=-8, color="#e74c3c", fontSize=11, font="DM Mono"
    ).encode(y="y:Q", text="label:N", x=alt.value(8))
    st.altair_chart((prod_bars + conso_rule + conso_label).properties(
        title=f"Production PV annuelle sur {duree} ans — dégradation {degradation*100:.1f}%/an",
        height=300,
    ), use_container_width=True)

    col_left, col_right = st.columns(2)
    with col_left:
        months = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
        # Seasonal weights calibrated for Monastir latitude
        seasonal = [0.068,0.077,0.094,0.107,0.114,0.109,0.102,0.096,0.088,0.079,0.068,0.064]
        df_mo = pd.DataFrame({
            "Mois": months,
            "Production": [prod_y1 * f for f in seasonal],
            "Consommation": [conso / 12] * 12,
            "ordre": list(range(12)),
        })
        mo_bars = alt.Chart(df_mo).mark_bar(color="#f5a623", opacity=0.8).encode(
            x=alt.X("Mois:N", sort=alt.SortField("ordre"), title=""),
            y=alt.Y("Production:Q", title="kWh/mois"),
            tooltip=[alt.Tooltip("Mois:N"), alt.Tooltip("Production:Q", format=",.0f")],
        )
        mo_line = alt.Chart(df_mo).mark_line(color="#e74c3c", strokeDash=[4,2], strokeWidth=2).encode(
            x=alt.X("Mois:N", sort=alt.SortField("ordre")),
            y="Consommation:Q",
            tooltip=[alt.Tooltip("Mois:N"), alt.Tooltip("Consommation:Q", format=",.0f")],
        )
        st.altair_chart((mo_bars + mo_line).properties(title="Profil mensuel Production vs Consommation", height=260), use_container_width=True)

    with col_right:
        # Energy balance donut — 3 slices: autoconso, vendue, transport (0), non valorisée
        autoconso_pct = autoconso_y1 / prod_y1 * 100
        cedee_pct     = cedee_lim_y1 / prod_y1 * 100
        perdu_pct     = max(0, 100 - autoconso_pct - cedee_pct)
        pie_df = pd.DataFrame([
            {"label": f"Autoconsommée ({autoconso_pct:.1f}%)",  "value": round(autoconso_pct, 1)},
            {"label": f"Vendue STEG ({cedee_pct:.1f}%)",        "value": round(cedee_pct, 1)},
            {"label": f"Non valorisée ({perdu_pct:.1f}%)",      "value": round(perdu_pct, 1)},
        ])
        pie = alt.Chart(pie_df).mark_arc(innerRadius=65, outerRadius=105).encode(
            theta=alt.Theta("value:Q"),
            color=alt.Color("label:N", scale=alt.Scale(
                domain=pie_df["label"].tolist(),
                range=["#f5a623","#2ecc71","#e74c3c"]
            ), legend=alt.Legend(title="Répartition")),
            tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=".1f", title="%")],
        ).properties(title="Répartition de la production PV (%)", height=260)
        st.altair_chart(pie, use_container_width=True)

    st.markdown('<div class="sec">Paramètres Techniques</div>', unsafe_allow_html=True)
    col_p1, col_p2, col_p3 = st.columns(3)
    params_tech = [
        ("Région", region),
        ("Orientation", orientation),
        ("Rendement spécifique", f"{rendement:,} kWh/kWc/an"),
        ("Dégradation", f"{degradation*100:.1f}%/an"),
        ("Production totale {d}a".format(d=duree), f"{total_prod/1000:,.0f} MWh"),
        ("Bilan instantané", "Net-metering"),
        ("Tarif transport", f"{tarif_transport:.3f} DT/kWh"),
        ("Limite excédent", f"{limite_exced*100:.0f}% de la production"),
        ("Taux de couverture", f"{taux_couv*100:.1f}%"),
    ]
    for i, (k, v) in enumerate(params_tech):
        with [col_p1, col_p2, col_p3][i % 3]:
            st.markdown(f'<div class="info"><strong>{k} :</strong> {v}</div>', unsafe_allow_html=True)


# ══════════════════════════════════
# TAB 3 — CASH FLOWS
# ══════════════════════════════════
with tab3:
    st.markdown('<div class="sec">Flux de Trésorerie Annuels</div>', unsafe_allow_html=True)

    df_cf = pd.DataFrame({
        "Année":   years,
        "Revenus": rev_list,
        "OpEx":    [-o for o in opex_list],
        "Dette":   [-d for d in ds_list],
        "CFADS":   cfads_list,
        "CF Equity": eq_cf_list,
        "Cumul":   cum_list,
    })
    df_cf_melt = df_cf.melt("Année", value_vars=["Revenus","OpEx","Dette"], var_name="Poste", value_name="DT")

    bars = alt.Chart(df_cf_melt).mark_bar().encode(
        x=alt.X("Année:O", title="Année"),
        y=alt.Y("DT:Q", title="DT/an", stack="zero"),
        color=alt.Color("Poste:N", scale=alt.Scale(
            domain=["Revenus","OpEx","Dette"],
            range=["#f5a623","#e74c3c","#8e44ad"]
        )),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Poste:N"), alt.Tooltip("DT:Q", format=",.0f")],
    )
    cum_line = alt.Chart(df_cf).mark_line(color="#2ecc71", strokeWidth=2.5, point=True).encode(
        x=alt.X("Année:O"),
        y=alt.Y("Cumul:Q", axis=alt.Axis(title="Cumul Equity DT", orient="right")),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Cumul:Q", format=",.0f", title="Cumul DT")],
    )
    zero_rule = alt.Chart(pd.DataFrame({"y":[0]})).mark_rule(color="#5a7a9a", strokeDash=[4,2]).encode(y="y:Q")
    cf_chart = alt.layer(bars, cum_line + zero_rule).resolve_scale(y="independent").properties(
        title=f"Revenus, coûts opérations et cumul equity ({duree} ans)", height=380
    )
    st.altair_chart(cf_chart, use_container_width=True)

    # CFADS vs Debt service
    st.markdown('<div class="sec">CFADS vs Service de la Dette</div>', unsafe_allow_html=True)
    df_dscr = pd.DataFrame({
        "Année": years,
        "CFADS (DT)": cfads_list,
        "Debt Service (DT)": ds_list,
    }).melt("Année", var_name="Poste", value_name="DT")
    dscr_chart = alt.Chart(df_dscr).mark_line(point=True, strokeWidth=2).encode(
        x=alt.X("Année:O", title="Année"),
        y=alt.Y("DT:Q", title="DT/an"),
        color=alt.Color("Poste:N", scale=alt.Scale(
            domain=["CFADS (DT)","Debt Service (DT)"],
            range=["#2ecc71","#e74c3c"]
        )),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Poste:N"), alt.Tooltip("DT:Q", format=",.0f")],
    ).properties(title="CFADS vs Service de la dette — couverture annuelle", height=260)
    st.altair_chart(dscr_chart, use_container_width=True)

    # Sensitivity
    st.markdown('<div class="sec">Analyse de Sensibilité</div>', unsafe_allow_html=True)
    col_s1, col_s2 = st.columns(2)

    def quick_irr(inv, rev_l, om_rate, infl, n):
        cfs = [-inv]
        for i, r in enumerate(rev_l[:n]):
            cfs.append(r - inv * om_rate * (1 + infl)**i)
        lo, hi = -0.5, 5.0
        for _ in range(80):
            mid = (lo + hi) / 2
            npv = sum(c / (1 + mid)**t for t, c in enumerate(cfs))
            if npv > 0: lo = mid
            else: hi = mid
        return (lo + hi) / 2 * 100

    with col_s1:
        costs_r = np.linspace(1200, 4000, 22)
        irrs_c  = [quick_irr(kwc*c*(1+tva_frac)-sub_dt, rev_list, om_pct, inflation, min(duree,20)) for c in costs_r]
        df_s1   = pd.DataFrame({"Coût DT/kWc": costs_r, "TRI %": irrs_c})
        l1 = alt.Chart(df_s1).mark_line(color="#f5a623", strokeWidth=2).encode(
            x=alt.X("Coût DT/kWc:Q"), y=alt.Y("TRI %:Q"),
            tooltip=[alt.Tooltip("Coût DT/kWc:Q", format=",.0f"), alt.Tooltip("TRI %:Q", format=".1f")],
        )
        a1 = alt.Chart(df_s1).mark_area(color="#f5a623", opacity=0.1).encode(x="Coût DT/kWc:Q", y="TRI %:Q", y2=alt.value(0))
        r1 = alt.Chart(pd.DataFrame({"y":[taux_disc*100]})).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
        v1 = alt.Chart(pd.DataFrame({"x":[prix_kwc]})).mark_rule(color="#2ecc71", strokeDash=[4,2]).encode(x="x:Q")
        st.altair_chart((a1+l1+r1+v1).properties(title="TRI vs Coût d'installation (DT/kWc)", height=260), use_container_width=True)

    with col_s2:
        tariffs_r = np.linspace(0, 0.12, 22)
        irrs_t = []
        for ht in tariffs_r:
            rev_t = [r * (1 + ht)**max(0, y-2) for y, r in zip(years[:min(duree,20)], rev_list[:min(duree,20)])]
            irrs_t.append(quick_irr(cout_net, rev_t, om_pct, inflation, min(duree,20)))
        df_s2 = pd.DataFrame({"Hausse %/an": tariffs_r*100, "TRI %": irrs_t})
        l2 = alt.Chart(df_s2).mark_line(color="#2ecc71", strokeWidth=2).encode(
            x=alt.X("Hausse %/an:Q"), y=alt.Y("TRI %:Q"),
            tooltip=[alt.Tooltip("Hausse %/an:Q", format=".1f"), alt.Tooltip("TRI %:Q", format=".1f")],
        )
        a2 = alt.Chart(df_s2).mark_area(color="#2ecc71", opacity=0.1).encode(x="Hausse %/an:Q", y="TRI %:Q", y2=alt.value(0))
        r2 = alt.Chart(pd.DataFrame({"y":[taux_disc*100]})).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
        v2 = alt.Chart(pd.DataFrame({"x":[hausse_tarif*100]})).mark_rule(color="#f5a623", strokeDash=[4,2]).encode(x="x:Q")
        st.altair_chart((a2+l2+r2+v2).properties(title="TRI vs Hausse du tarif STEG (%/an)", height=260), use_container_width=True)

    with st.expander("📋 Tableau complet des flux de trésorerie"):
        df_table = pd.DataFrame({
            "Année":              years,
            "Production (kWh)":  [round(p) for p in prod_list],
            "Revenus (DT)":      [round(r) for r in rev_list],
            "OpEx (DT)":         [round(o) for o in opex_list],
            "Service dette (DT)":[round(d) for d in ds_list],
            "CFADS (DT)":        [round(c) for c in cfads_list],
            "CF Equity (DT)":    [round(e) for e in eq_cf_list],
            "Cumul Equity (DT)": [round(c) for c in cum_list],
        })
        def color_signed(val):
            if isinstance(val, (int, float)):
                if val < 0:
                    return "background-color:rgba(231,76,60,0.25);color:#ffaaaa;"
                elif val > 0:
                    return "background-color:rgba(46,204,113,0.20);color:#aaffbb;"
            return ""
        styled = df_table.style.format({
            "Production (kWh)": "{:,.0f}", "Revenus (DT)": "{:,.0f}",
            "OpEx (DT)": "{:,.0f}", "Service dette (DT)": "{:,.0f}",
            "CFADS (DT)": "{:,.0f}", "CF Equity (DT)": "{:,.0f}", "Cumul Equity (DT)": "{:,.0f}",
        }).map(color_signed, subset=["CF Equity (DT)", "Cumul Equity (DT)", "CFADS (DT)"])
        st.dataframe(styled, use_container_width=True)


# ══════════════════════════════════
# TAB 4 — INVESTMENT & FINANCING
# ══════════════════════════════════
with tab4:
    st.markdown('<div class="sec">Structure de l\'Investissement</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    inv_kpis = [
        (c1, "Coût HT",         f"{cout_ht:,.0f} DT",        f"{kwc:.2f} kWc × {prix_kwc:.0f} DT/kWc",  ""),
        (c2, f"TVA ({tva_pct}%)","56 402 DT" if use_excel else f"{cout_ht*tva_frac:,.0f} DT", "Taxe valeur ajoutée", ""),
        (c3, "Coût TTC Net",    f"{cout_net:,.0f} DT",        f"Subvention: {sub_dt:,.0f} DT",            "kpi-blue"),
        (c4, "Capital Initial", f"{cap_propres:,.0f} DT",     f"Reste à charge propriétaire",             ""),
    ]
    for col, label, val, unit, cls in inv_kpis:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        fin_df = pd.DataFrame([
            {"label": f"Capitaux propres\n{cap_propres:,.0f} DT", "value": cap_propres},
            {"label": f"Dette bancaire\n{montant_dette:,.0f} DT",  "value": montant_dette},
        ])
        fin_pie = alt.Chart(fin_df).mark_arc(innerRadius=60, outerRadius=105).encode(
            theta=alt.Theta("value:Q"),
            color=alt.Color("label:N", scale=alt.Scale(
                domain=fin_df["label"].tolist(), range=["#f5a623","#3498db"]
            ), legend=alt.Legend(title="Financement")),
            tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=",.0f", title="DT")],
        ).properties(title="Structure de financement", height=280)
        st.altair_chart(fin_pie, use_container_width=True)

        fin_params = [
            ("Capitaux propres",       f"{cap_propres:,.0f} DT ({(1-dette_pct)*100:.0f}%)"),
            ("Dette bancaire",         f"{montant_dette:,.0f} DT ({dette_pct*100:.0f}% — plafond 200 000 DT)"),
            ("Taux intérêt effectif",  f"{(taux_ref+marge)*100:.2f}% (BCT {taux_ref*100:.2f}% + marge {marge*100:.2f}%)"),
            ("Durée du crédit",        f"{duree_credit} ans"),
            ("Délai de grâce",         f"{grace} an(s)"),
            ("Commission initiale",    "1,0 % du crédit"),
            ("Commission d'engagement","0,25 % de la marge"),
            ("Taux d'actualisation",   f"{taux_disc*100:.1f}%"),
        ]
        for k, v in fin_params:
            st.markdown(f'<div class="info"><strong>{k} :</strong> {v}</div>', unsafe_allow_html=True)

    with col_f2:
        # Debt schedule with real numbers
        debt_schedule = []
        out_b = montant_dette
        for y in years:
            if y <= grace:
                interest_y = out_b * taux_interet; princ_y = 0.0
            elif y <= duree_credit:
                princ_y = montant_dette / max(duree_credit - grace, 1)
                interest_y = out_b * taux_interet
                out_b = max(0, out_b - princ_y)
            else:
                princ_y = 0.0; interest_y = 0.0
            debt_schedule.append({"Année": y, "Principal": princ_y, "Intérêts": interest_y, "Encours": out_b})
        df_debt = pd.DataFrame(debt_schedule)
        df_dm   = df_debt.melt("Année", value_vars=["Principal","Intérêts"], var_name="Poste", value_name="DT")

        d_bars = alt.Chart(df_dm).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
            x=alt.X("Année:O"), y=alt.Y("DT:Q", title="DT/an", stack="zero"),
            color=alt.Color("Poste:N", scale=alt.Scale(domain=["Principal","Intérêts"], range=["#3498db","#8e44ad"])),
            tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Poste:N"), alt.Tooltip("DT:Q", format=",.0f")],
        )
        enc_line = alt.Chart(df_debt).mark_line(color="#e74c3c", strokeWidth=2, strokeDash=[4,2]).encode(
            x=alt.X("Année:O"),
            y=alt.Y("Encours:Q", axis=alt.Axis(title="Encours DT", orient="right")),
            tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Encours:Q", format=",.0f")],
        )
        st.altair_chart(alt.layer(d_bars, enc_line).resolve_scale(y="independent").properties(
            title="Remboursement dette bancaire + encours", height=340
        ), use_container_width=True)


# ══════════════════════════════════
# TAB 5 — BANKING RATIOS (NEW)
# ══════════════════════════════════
with tab5:
    st.markdown('<div class="sec">Ratios Bancaires — Indicateurs de Solvabilité</div>', unsafe_allow_html=True)

    # DSCR = CFADS / Debt Service per year
    dscr_vals = []
    for cf, ds in zip(cfads_list, ds_list):
        if ds > 0:
            dscr_vals.append({"Année": years[cfads_list.index(cf)], "DSCR": cf / ds, "type": "DSCR"})

    c1, c2, c3, c4 = st.columns(4)
    dscr_y1 = cfads_list[0] / ds_list[0] if ds_list[0] > 0 else 0
    dscr_min_calc = min(cf/ds for cf, ds in zip(cfads_list, ds_list) if ds > 0) if any(d > 0 for d in ds_list) else 0
    banking_kpis = [
        (c1, "DSCR Min",          f"{(dscr_min if use_excel and duree==20 else dscr_min_calc):.4f}",  "Ratio couverture service dette",    "kpi-green" if dscr_min_calc > 1.2 else "kpi-warn"),
        (c2, "LLCR Min",          f"{(llcr_min if use_excel and duree==20 else 0):.4f}",              "Ratio couverture durée prêt",        "kpi-green"),
        (c3, "TRI Equity",        f"{eq_irr_pct:.2f}%",                                               "Retour capitaux propres",            "kpi-green"),
        (c4, "Seuil rentabilité", f"{taux_disc*100:.1f}%",                                            "Taux actualisation (coût du capital)","kpi-blue"),
    ]
    for col, label, val, unit, cls in banking_kpis:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        if dscr_vals:
            df_dscr_plot = pd.DataFrame(dscr_vals)
            dscr_bars = alt.Chart(df_dscr_plot).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
                x=alt.X("Année:O", title="Année"),
                y=alt.Y("DSCR:Q", title="DSCR", scale=alt.Scale(domain=[0, df_dscr_plot["DSCR"].max()*1.2])),
                color=alt.condition(
                    alt.datum.DSCR >= 1.2,
                    alt.value("#2ecc71"),
                    alt.condition(alt.datum.DSCR >= 1.0, alt.value("#f5a623"), alt.value("#e74c3c"))
                ),
                tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("DSCR:Q", format=".3f")],
            )
            dscr_rule = alt.Chart(pd.DataFrame({"y":[1.2]})).mark_rule(color="#2ecc71", strokeDash=[4,2]).encode(y="y:Q")
            dscr_rule1 = alt.Chart(pd.DataFrame({"y":[1.0]})).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
            st.altair_chart((dscr_bars + dscr_rule + dscr_rule1).properties(
                title="DSCR par année (vert ≥ 1.2, orange ≥ 1.0, rouge < 1.0)", height=300
            ), use_container_width=True)

    with col_d2:
        # Cumulative project cashflow
        cf_proj = [-cout_net] + ncf_list
        cum_proj = list(np.cumsum(cf_proj))
        df_cum_proj = pd.DataFrame({"Période": list(range(duree+1)), "Cumul projet (DT)": cum_proj})
        cum_line2 = alt.Chart(df_cum_proj).mark_line(color="#f5a623", strokeWidth=2.5).encode(
            x=alt.X("Période:O", title="Année"),
            y=alt.Y("Cumul projet (DT):Q", title="DT cumulés"),
            tooltip=[alt.Tooltip("Période:O"), alt.Tooltip("Cumul projet (DT):Q", format=",.0f")],
        )
        cum_area = alt.Chart(df_cum_proj).mark_area(color="#f5a623", opacity=0.1).encode(
            x="Période:O", y="Cumul projet (DT):Q", y2=alt.value(0)
        )
        zero_r = alt.Chart(pd.DataFrame({"y":[0]})).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
        st.altair_chart((cum_area+cum_line2+zero_r).properties(
            title=f"Flux de trésorerie projet cumulés — payback ~{payback_val:.1f} ans", height=300
        ), use_container_width=True)

    st.markdown('<div class="sec">Interprétation des Ratios</div>', unsafe_allow_html=True)
    col_r1, col_r2 = st.columns(2)
    ratios_info = [
        ("DSCR (Debt Service Coverage Ratio)", f"Min = {dscr_min if use_excel else dscr_min_calc:.4f} — Au-dessus de 1.2 : la banque est rassurée. Au-dessus de 1.0 : le projet couvre la dette."),
        ("LLCR (Loan Life Coverage Ratio)", f"Min = {llcr_min if use_excel else 0:.4f} — Ratio de couverture actualisé sur toute la durée du prêt. Cible bancaire typique : ≥ 1.3"),
        ("TRI Equity", f"{eq_irr_pct:.2f}% — Retour sur les capitaux propres investis. Supérieur au taux d'actualisation ({taux_disc*100:.1f}%) : projet viable pour l'investisseur."),
        ("VAN", f"{project_npv/1000:+.0f}k DT — Valeur créée au-delà du taux d'actualisation. Positive : le projet crée de la valeur."),
        ("Temps de retour", f"{payback_val:.1f} ans — Inférieur à la durée du crédit ({duree_credit} ans) : le projet se rembourse avant la fin du prêt."),
        ("LCOE vs Tarif STEG", f"LCOE = {lcoe:.4f} DT/kWh < Tarif = {tarif_kwh:.3f} DT/kWh → Production PV moins chère que l'achat STEG."),
    ]
    for i, (k, v) in enumerate(ratios_info):
        with col_r1 if i % 2 == 0 else col_r2:
            st.markdown(f'<div class="info"><strong>{k} :</strong> {v}</div>', unsafe_allow_html=True)


# ══════════════════════════════════
# TAB 6 — ENVIRONMENTAL IMPACT
# ══════════════════════════════════
with tab6:
    st.markdown('<div class="sec">Impact Environnemental</div>', unsafe_allow_html=True)

    trees = co2_total * 40
    cars  = co2_total / 2.3
    homes = total_prod / 4000

    c1, c2, c3, c4, c5 = st.columns(5)
    env_kpis = [
        (c1, f"CO₂ Évité {duree} ans",  f"{co2_total:,.0f} t",        "tonnes éq. CO₂",        "kpi-green"),
        (c2, "Énergie Propre",           f"{total_prod/1000:,.0f} MWh", f"sur {duree} ans",       ""),
        (c3, "Équivalent Arbres",        f"{trees:,.0f}",               "arbres/an équivalent",  "kpi-green"),
        (c4, "Voitures retirées",        f"{cars:,.0f}",                "voitures/an éq.",       ""),
        (c5, "Foyers alimentés",         f"{homes:,.0f}",               "foyers/an (4 MWh/foyer)","kpi-blue"),
    ]
    for col, label, val, unit, cls in env_kpis:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    df_co2 = pd.DataFrame({"Année": years, "CO₂ évité (t)": co2_per_year, "CO₂ cumulé (t)": co2_cum})
    co2_bars = alt.Chart(df_co2).mark_bar(color="#2ecc71", opacity=0.75, cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
        x=alt.X("Année:O"), y=alt.Y("CO₂ évité (t):Q", title="t CO₂/an"),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("CO₂ évité (t):Q", format=".1f")],
    )
    co2_cum_line = alt.Chart(df_co2).mark_line(color="#f5a623", strokeWidth=2.5).encode(
        x=alt.X("Année:O"),
        y=alt.Y("CO₂ cumulé (t):Q", axis=alt.Axis(title="t CO₂ cumulé", orient="right")),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("CO₂ cumulé (t):Q", format=".0f")],
    )
    st.altair_chart(alt.layer(co2_bars, co2_cum_line).resolve_scale(y="independent").properties(
        title="CO₂ évité par an et cumulé — facteur 0,57 t CO₂/MWh (réseau Tunisie)", height=340
    ), use_container_width=True)

    st.markdown('<div class="sec">Mise en Contexte</div>', unsafe_allow_html=True)
    col_c1, col_c2, col_c3 = st.columns(3)
    context = [
        ("Facteur émission Tunisie", "0,57 t CO₂ évitées par MWh solaire (réseau national)"),
        ("CO₂ évité — Année 1",      f"{co2_per_year[0]:.1f} tonnes"),
        ("CO₂ évité — Année 20",     f"{co2_per_year[-1]:.1f} tonnes (dégradation incluse)"),
        ("Équivalent arbres",         f"1 arbre absorbe ~25 kg CO₂/an → {int(trees):,} arbres"),
        ("Équivalent voitures",       f"Voiture = 2,3 t CO₂/an → {int(cars):,} voitures retirées/an"),
        ("Foyers tunisiens",          f"Conso moy. = 4 MWh/foyer/an → {int(homes):,} foyers alimentés"),
        ("Production Y1",             f"{EXCEL['prod_y1']:,.0f} kWh = {EXCEL['prod_y1']/1000:.0f} MWh"),
        ("Production totale",         f"{total_prod/1000:,.0f} MWh sur {duree} ans"),
        ("CO₂ total évité",           f"{co2_total:,.0f} tonnes = {co2_total*1000:.0f} kg"),
    ]
    for i, (k, v) in enumerate(context):
        with [col_c1, col_c2, col_c3][i % 3]:
            st.markdown(f'<div class="info"><strong>{k} :</strong> {v}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("""<div class="footer">
PROFITPV · MT · Monastir · 197,21 kWc · 2025 · CRA2E · eclareon / GIZ / ANME Tunisie<br>
Données extraites de : PROFITPV_MT_CRA2E_v0.2_TLS.xlsm — Les résultats sont basés sur les hypothèses saisies.
</div>""", unsafe_allow_html=True)
