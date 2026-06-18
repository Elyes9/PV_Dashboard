"""
PROFITPV Dashboard — Solar PV Installation Analyzer
Source: PROFITPV_MT_CRA2E_v0_TLS.xlsm (updated file)
Libraries: streamlit, pandas, numpy, altair only
Run: streamlit run profitpv_dashboard.py
"""

import streamlit as st
import streamlit.components.v1 as components
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
.kpi-blue .kpi-value  { color:#3498db; }
.kpi-blue::before  { background:linear-gradient(180deg,#3498db,#2980b9); }
.kpi-red .kpi-value   { color:#e74c3c; }
.kpi-red::before   { background:linear-gradient(180deg,#e74c3c,#c0392b); }
.sec { font-size:11px; font-weight:700; letter-spacing:2px; text-transform:uppercase;
       color:#f5a623; margin:22px 0 10px; padding-bottom:6px; border-bottom:1px solid #1e2d3d; }
.info { background:#111820; border:1px solid #1e2d3d; border-radius:8px;
        padding:12px 16px; font-size:13px; color:#8aa4be; margin-bottom:8px; }
.info strong { color:#c5d8ea; }
.badge-ok   { background:#0d2b1a; border:1px solid #2ecc71; border-radius:10px; padding:16px; text-align:center; }
.badge-warn { background:#2b1f0d; border:1px solid #f5a623; border-radius:10px; padding:16px; text-align:center; }
.badge-bad  { background:#2b0d0d; border:1px solid #e74c3c; border-radius:10px; padding:16px; text-align:center; }
.excel-tag { background:#1a3a1a; border-radius:4px; padding:2px 8px;
             font-family:'DM Mono',monospace; font-size:10px; color:#2ecc71; margin-left:8px;
             border:1px solid #2ecc71; }
.footer { text-align:center; color:#2a4a6a; font-family:'DM Mono',monospace;
          font-size:11px; padding:10px 0; margin-top:20px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ALTAIR DARK THEME
# ─────────────────────────────────────────────
def pv_theme():
    return {"config": {
        "background": "#111820",
        "view": {"stroke": "transparent", "fill": "#111820"},
        "axis": {"gridColor":"#1e2d3d","domainColor":"#1e2d3d","tickColor":"#1e2d3d",
                 "labelColor":"#8aa4be","titleColor":"#8aa4be","labelFont":"DM Mono",
                 "titleFont":"Syne","labelFontSize":11,"titleFontSize":12},
        "legend": {"labelColor":"#8aa4be","titleColor":"#8aa4be","fillColor":"#111820",
                   "strokeColor":"#1e2d3d","labelFont":"DM Mono","titleFont":"Syne"},
        "title": {"color":"#8aa4be","font":"Syne","fontSize":13,"anchor":"start"},
        "mark": {"color":"#f5a623"},
        "range": {"category":["#f5a623","#2ecc71","#3498db","#e74c3c","#8e44ad","#1abc9c"]},
    }}
alt.themes.register("pv_theme", pv_theme)
alt.themes.enable("pv_theme")

# ─────────────────────────────────────────────
# EXCEL DATA — PROFITPV_MT_CRA2E_v0_TLS.xlsm
# All values extracted directly from the file
# ─────────────────────────────────────────────
EXCEL = {
    # ── Installation ─────────────────────────
    "kwc":          157.80,
    "prix_kwc":     2200.0,
    "cout_ht":      347160.0,
    "tva_dt":       45130.8,
    "cout_ttc":     392290.8,
    "sub_dt":       0.0,
    "rendement":    1659.09,
    "degradation":  0.004,
    "region":       "Monastir",
    "orientation":  "Sud",
    "annee_debut":  2025,
    "duree":        20,

    # ── Consumption & Tariffs ─────────────────
    "conso":        315628.0,
    "tarif_kwh":    0.291,
    "tarif_jour":   0.290,
    "tarif_pointe_matin": 0.417,
    "tarif_pointe_soir":  0.377,
    "tarif_nuit":   0.222,
    "surtaxe":      0.006,
    "tarif_achat":  0.08,
    "tarif_transport": 0.039,
    "limite_exced": 0.30,
    "tva_conso":    0.0,
    "hausse_tarif": 0.05,   # after year 2

    # ── Energy results ────────────────────────
    "prod_y1":        261804.40,
    "autoconso_y1":   183752.49,
    "cedee_y1":       78051.91,
    "taux_couv":      0.5822,
    "pct_cedee":      0.2981,
    "facture_sans":   93741.52,
    "facture_avec":   39167.03,
    "vente_exced":    6244.15,
    "remb_steg":      0.0,
    "eco_kwh":        0.23231,  # économies facturation DT/kWh

    # ── Financial indicators ──────────────────
    "tri_projet":     0.16890,
    "tri_equity":     0.22881,
    "van":            294108.42,
    "lcoe":           0.20275,
    "payback":        6.2882,
    "dscr_min":       1.17165,
    "llcr_min":       1.38870,
    "om_y1":          13886.4,

    # ── Financing ─────────────────────────────
    "cout_ht_inv":    347160.0,
    "cap_propres":    104148.0,
    "montant_dette":  243012.0,
    "dette_pct":      0.70,
    "taux_interet":   0.10,
    "taux_disc":      0.08,
    "duree_credit":   10,
    "grace":          1,
    "inflation":      0.0102,
    "om_pct":         0.04,

    # ── Annual revenues (Y1–Y20) from Flux trésorerie ──
    "rev_years": [
        60818.64, 63604.14, 66517.21, 69563.69, 72394.04,
        75355.44, 78453.89, 81695.65, 85087.30, 88635.68,
        92347.98, 96231.68, 100294.65, 104545.07, 108991.54,
        113643.02, 118508.90, 123599.00, 128923.56, 134493.34,
    ],
    # ── Annual OpEx (Y1–Y20) ──────────────────
    "opex_years": [
        13886.40, 14164.13, 14447.41, 14736.36, 15031.09,
        15331.71, 15638.34, 15951.11, 16270.13, 16595.53,
        16927.44, 17266.00, 17611.31, 17963.54, 18322.81,
        18689.27, 19063.05, 19444.31, 19833.20, 20229.86,
    ],
    # ── Annual CFADS (Y1–Y20) ─────────────────
    "cfads_years": [
        46932.24, 49440.01, 52069.79, 54827.33, 57362.95,
        60023.73, 62815.54, 65744.54, 68817.17, 72040.15,
        75420.53, 78965.69, 82683.33, 86581.53, 90668.73,
        94953.75, 99445.85, 104154.68, 109090.37, 114263.48,
    ],
    # ── Annual Debt Service (Y1–Y20) ─────────
    "ds_years": [
        24301.20, 42196.73, 42196.73, 42196.73, 42196.73,
        42196.73, 42196.73, 42196.73, 42196.73, 42196.73,
        0.0, 0.0, 0.0, 0.0, 0.0,
        0.0, 0.0, 0.0, 0.0, 0.0,
    ],
    # ── Equity CF (Y1–Y20) ───────────────────
    "eq_cf_years": [
        22631.04, 7243.27, 9873.06, 12630.60, 15166.22,
        17826.99, 20618.81, 23547.81, 26620.43, 29843.41,
        75420.53, 78965.69, 82683.33, 86581.53, 90668.73,
        94953.75, 99445.85, 104154.68, 109090.37, 114263.48,
    ],
    # ── CO2 reduction ────────────────────────
    "co2_total": 2984.57,
}

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ☀️ PROFITPV")
    st.markdown("<div style='color:#5a7a9a;font-size:12px;margin-bottom:4px;font-family:DM Mono,monospace;'>CRA2E · MT · v0 TLS</div>", unsafe_allow_html=True)
    st.markdown("<div style='color:#2ecc71;font-size:11px;margin-bottom:16px;font-family:DM Mono,monospace;'>✓ 157,80 kWc · Monastir 2025</div>", unsafe_allow_html=True)

    use_excel = st.toggle("🔒 Utiliser valeurs Excel réelles", value=True)

    st.markdown('<div class="sec">🔧 Installation PV</div>', unsafe_allow_html=True)
    kwc        = st.number_input("Taille installation (kWc)", 1.0, 5000.0, EXCEL["kwc"], 1.0)
    prix_kwc   = st.number_input("Prix unitaire HT (DT/kWc)", 500.0, 5000.0, EXCEL["prix_kwc"], 50.0)
    subvention = st.selectbox("Subvention FTE ?", ["Non", "Oui"])
    sub_dt     = st.number_input("Montant subvention (DT)", 0.0, value=0.0, step=1000.0) if subvention == "Oui" else 0.0
    tva_pct    = st.slider("TVA (%)", 0, 20, 13)

    st.markdown('<div class="sec">📍 Site & Panneaux</div>', unsafe_allow_html=True)
    region      = st.selectbox("Région", ["Monastir","Tunis","Sfax","Bizerte","Gafsa","Jendouba","Jerba","Kairouan","Tataouine","Tozeur"])
    orientation = st.selectbox("Orientation", ["Sud","Sud-ouest"])
    yield_map   = {"Monastir":1659,"Sfax":1680,"Gafsa":1720,"Jerba":1700,"Kairouan":1690,
                   "Tataouine":1730,"Tozeur":1740,"Bizerte":1580,"Jendouba":1560,"Tunis":1600}
    rendement   = st.number_input("Rendement spécifique (kWh/kWc/an)", 1000, 2000, int(EXCEL["rendement"]))
    degradation = st.slider("Dégradation annuelle (%)", 0.0, 1.5, EXCEL["degradation"]*100, 0.1) / 100

    st.markdown('<div class="sec">⚡ Consommation & Tarifs</div>', unsafe_allow_html=True)
    conso        = st.number_input("Consommation annuelle (kWh/an)", 1000.0, 5e6, EXCEL["conso"], 1000.0)
    tarif_kwh    = st.number_input("Tarif uniforme STEG (DT/kWh)", 0.1, 1.0, EXCEL["tarif_kwh"], 0.001, format="%.3f")
    tarif_jour   = st.number_input("Tarif jour (DT/kWh)", 0.1, 1.0, EXCEL["tarif_jour"], 0.001, format="%.3f")
    tarif_pointe = st.number_input("Tarif pointe soir (DT/kWh)", 0.1, 1.0, EXCEL["tarif_pointe_soir"], 0.001, format="%.3f")
    tarif_nuit   = st.number_input("Tarif nuit/dim (DT/kWh)", 0.1, 1.0, EXCEL["tarif_nuit"], 0.001, format="%.3f")
    surtaxe      = st.number_input("Surtaxe municipale (DT/kWh)", 0.0, 0.05, EXCEL["surtaxe"], 0.001, format="%.3f")
    limite_exced = st.slider("Limite excédent annuel (%)", 10, 50, int(EXCEL["limite_exced"]*100)) / 100
    tarif_achat  = st.number_input("Tarif achat excédent (DT/kWh)", 0.01, 0.30, EXCEL["tarif_achat"], 0.001, format="%.3f")
    tarif_transport = st.number_input("Tarif transport (DT/kWh)", 0.01, 0.20, EXCEL["tarif_transport"], 0.001, format="%.3f")

    st.markdown('<div class="sec">💰 Financement</div>', unsafe_allow_html=True)
    dette_pct    = st.slider("Part dette (%)", 0, 100, int(EXCEL["dette_pct"]*100)) / 100
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
# COMPUTATIONS
# ─────────────────────────────────────────────
tva_frac  = tva_pct / 100
cout_ht   = kwc * prix_kwc
cout_ttc  = cout_ht * (1 + tva_frac) - sub_dt
cout_net  = max(cout_ttc, 0)

years = list(range(1, duree + 1))

if use_excel and duree == 20:
    # ── Use exact Excel values ────────────────
    kwc_val       = EXCEL["kwc"]
    prod_y1       = EXCEL["prod_y1"]
    autoconso_y1  = EXCEL["autoconso_y1"]
    cedee_lim_y1  = EXCEL["cedee_y1"]
    taux_couv     = EXCEL["taux_couv"]
    facture_sans  = EXCEL["facture_sans"]
    facture_avec  = EXCEL["facture_avec"]
    eco_annuelle  = facture_sans - facture_avec
    rev_exced_y1  = EXCEL["vente_exced"]
    om_y1_val     = EXCEL["om_y1"]
    cout_net      = EXCEL["cout_ttc"]
    cout_ht       = EXCEL["cout_ht"]
    montant_dette = EXCEL["montant_dette"]
    cap_propres   = EXCEL["cap_propres"]
    rev_list      = EXCEL["rev_years"]
    opex_list     = EXCEL["opex_years"]
    ds_list       = EXCEL["ds_years"]
    cfads_list    = EXCEL["cfads_years"]
    eq_cf_list    = EXCEL["eq_cf_years"]
    prod_list     = [EXCEL["prod_y1"] * (1 - EXCEL["degradation"])**(y-1) for y in years]
    project_irr   = EXCEL["tri_projet"]
    equity_irr    = EXCEL["tri_equity"]
    project_npv   = EXCEL["van"]
    lcoe          = EXCEL["lcoe"]
    payback_val   = EXCEL["payback"]
    dscr_min      = EXCEL["dscr_min"]
    llcr_min      = EXCEL["llcr_min"]
    # Cumulative equity CF (starting from -cap_propres)
    cum_list = []
    running = -EXCEL["cap_propres"]
    for ec in eq_cf_list:
        running += ec
        cum_list.append(running)

else:
    # ── Recalculate from sidebar inputs ──────
    kwc_val      = kwc
    prod_y1      = kwc * rendement
    taux_couv    = min(prod_y1 / conso, 1.0)
    autoconso_y1 = prod_y1 * min(taux_couv, 0.85)
    cedee_y1     = max(0, prod_y1 - autoconso_y1)
    cedee_lim_y1 = min(cedee_y1, prod_y1 * limite_exced)
    facture_sans = conso * (tarif_kwh + surtaxe)
    eco_ac       = autoconso_y1 * (tarif_kwh + surtaxe)
    rev_exced_y1 = cedee_lim_y1 * tarif_achat
    facture_avec = facture_sans - eco_ac - rev_exced_y1
    eco_annuelle = facture_sans - facture_avec
    montant_dette= min(cout_net * dette_pct, 200_000)
    cap_propres  = cout_net - montant_dette
    om_y1_val    = cout_net * om_pct

    prod_list, rev_list, opex_list, ds_list, cfads_list, eq_cf_list, cum_list = [], [], [], [], [], [], []
    outstanding = montant_dette
    running = -cap_propres

    for y in years:
        p  = prod_y1 * (1 - degradation)**(y-1)
        t  = tarif_kwh if y <= 2 else tarif_kwh * (1 + hausse_tarif)**(y-2)
        ac = p * min(conso / (p + 1e-9), 0.85)
        ce = min(max(0, p - ac), p * limite_exced)
        r  = ac * (t + surtaxe) + ce * tarif_achat * (1 + hausse_tarif)**max(0, y-2)
        o  = cout_net * om_pct * (1 + inflation)**(y-1)
        if y <= grace:
            ds = outstanding * (taux_ref + marge); princ = 0.0
        elif y <= duree_credit:
            princ = montant_dette / max(duree_credit - grace, 1)
            ds = outstanding * (taux_ref + marge) + princ
            outstanding = max(0, outstanding - princ)
        else:
            ds = 0.0
        cfads = r - o
        eq_cf = cfads - ds
        running += eq_cf
        prod_list.append(p); rev_list.append(r); opex_list.append(o)
        ds_list.append(ds); cfads_list.append(cfads)
        eq_cf_list.append(eq_cf); cum_list.append(running)

    def calc_npv(rate, cfs):
        return np.sum(np.array(cfs) / (1 + rate)**np.arange(len(cfs)))
    def calc_irr(cfs):
        lo, hi = -0.5, 5.0
        for _ in range(200):
            mid = (lo + hi) / 2
            if calc_npv(mid, cfs) > 0: lo = mid
            else: hi = mid
            if hi - lo < 1e-8: break
        return (lo + hi) / 2

    project_irr = calc_irr([-cout_net] + cfads_list)
    equity_irr  = calc_irr([-cap_propres] + eq_cf_list)
    project_npv = calc_npv(taux_disc, [-cout_net] + cfads_list)
    dc = cout_net + sum((cout_net*om_pct*(1+inflation)**(y-1))/(1+taux_disc)**y for y in range(1, duree+1))
    dp = sum((prod_y1*(1-degradation)**(y-1))/(1+taux_disc)**y for y in range(1, duree+1))
    lcoe = dc / dp if dp > 0 else 0
    payback_val = next((i+1.0 for i,v in enumerate(cum_list) if v > 0), float(duree+1))
    dscr_min = min(c/d for c,d in zip(cfads_list, ds_list) if d > 0) if any(d>0 for d in ds_list) else 0
    llcr_min = 0.0

# Derived
irr_pct    = project_irr * 100
eq_irr_pct = equity_irr * 100
total_prod = sum(prod_list)
co2_list   = [prod_list[i] * 0.57 / 1000 for i in range(len(years))]
co2_total  = sum(co2_list)
co2_cum    = list(np.cumsum(co2_list))
reduction  = (facture_sans - facture_avec) / facture_sans * 100 if facture_sans > 0 else 0

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
if irr_pct > 12:
    bcls, bcol, blbl = "badge-ok",   "#2ecc71", "TRÈS RENTABLE"
elif irr_pct > 8:
    bcls, bcol, blbl = "badge-ok",   "#2ecc71", "RENTABLE"
elif irr_pct > 5:
    bcls, bcol, blbl = "badge-warn", "#f5a623", "MARGINAL"
else:
    bcls, bcol, blbl = "badge-bad",  "#e74c3c", "DÉFICITAIRE"

xtag = '<span class="excel-tag">EXCEL ✓</span>' if use_excel and duree == 20 else ""
c_hdr, c_bdg = st.columns([4, 1])
with c_hdr:
    st.markdown(f"<h1 style='color:#f5a623;margin-bottom:4px;'>☀️ Analyse Installation PV {xtag}</h1>", unsafe_allow_html=True)
    st.markdown(
        f"<div style='color:#5a7a9a;font-family:DM Mono,monospace;font-size:13px;'>"
        f"{EXCEL['kwc'] if use_excel and duree==20 else kwc:.2f} kWc · "
        f"{EXCEL['region']} · {EXCEL['orientation']} · {EXCEL['annee_debut']} · "
        f"{duree} ans · PROFITPV MT CRA2E v0 TLS</div>",
        unsafe_allow_html=True,
    )
with c_bdg:
    st.markdown(f"""<div class="{bcls}" style="margin-top:10px;">
        <div style='color:{bcol};font-size:15px;font-weight:800;'>{blbl}</div>
        <div style='color:{bcol};font-family:DM Mono,monospace;font-size:13px;'>TRI = {irr_pct:.2f}%</div>
        <div style='color:{bcol};font-family:DM Mono,monospace;font-size:11px;'>Equity = {eq_irr_pct:.2f}%</div>
    </div>""", unsafe_allow_html=True)

st.markdown("---")

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📊 Résumé Financier",
    "⚡ Énergie & Production",
    "💵 Flux de Trésorerie",
    "🏗️ Investissement & Financement",
    "📐 Ratios Bancaires",
    "🌍 Impact Environnemental",
    "🧩 Vue 3D Installation",
])

# ── helpers ──────────────────────────────────
def kpi(col, label, val, unit, cls=""):
    with col:
        st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div>'
                    f'<div class="kpi-value">{val}</div>'
                    f'<div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

def info(label, val):
    st.markdown(f'<div class="info"><strong>{label} :</strong> {val}</div>', unsafe_allow_html=True)

# ══════════════════════════════════
# TAB 1 — FINANCIAL SUMMARY
# ══════════════════════════════════
with tab1:
    st.markdown('<div class="sec">Indicateurs Clés de Rentabilité</div>', unsafe_allow_html=True)

    c1,c2,c3,c4,c5 = st.columns(5)
    kpi(c1, "TRI Projet",        f"{irr_pct:.2f}%",               "Taux de rentabilité interne",          "")
    kpi(c2, "TRI Capitaux Propres", f"{eq_irr_pct:.2f}%",         "Equity IRR",                           "kpi-green")
    kpi(c3, "VAN Projet",        f"{project_npv/1000:+.1f}k DT",  f"Taux disc. {taux_disc*100:.0f}%",     "kpi-blue" if project_npv > 0 else "kpi-red")
    kpi(c4, "Temps de Retour",   f"{payback_val:.2f} ans",         "Retour sur investissement",            "")
    kpi(c5, "LCOE",              f"{lcoe:.4f} DT/kWh",            f"vs STEG {EXCEL['tarif_kwh']:.3f} DT/kWh", "kpi-green" if lcoe < EXCEL["tarif_kwh"] else "kpi-red")

    c6,c7,c8,c9 = st.columns(4)
    kpi(c6, "Économie Année 1",   f"{eco_annuelle:,.0f} DT",       f"Facture: {facture_avec:,.0f} → {facture_sans:,.0f} DT", "")
    kpi(c7, "Réduction Facture",  f"{reduction:.1f}%",              "de la facture STEG annuelle",          "kpi-green")
    kpi(c8, f"Revenus {duree} ans", f"{sum(rev_list)/1000:.0f}k DT","Revenus bruts cumulés",               "")
    kpi(c9, "Investissement TTC", f"{(EXCEL['cout_ttc'] if use_excel and duree==20 else cout_net)/1000:.1f}k DT", "Coût net après subvention", "kpi-blue")

    st.markdown('<div class="sec">Décomposition de la Facture STEG — Année 1</div>', unsafe_allow_html=True)

    cl, cr = st.columns([3, 2])
    with cl:
        eco_ac_y1 = autoconso_y1 * (tarif_kwh + surtaxe) if not (use_excel and duree==20) else (facture_sans - facture_avec - rev_exced_y1)
        wf = pd.DataFrame([
            {"cat": "Facture sans PV",    "base": 0,                           "val": facture_sans,  "type": "facture"},
            {"cat": "Économies autoconso","base": facture_sans - eco_ac_y1,    "val": eco_ac_y1,     "type": "économie"},
            {"cat": "Vente excédent STEG","base": facture_sans - eco_ac_y1 - rev_exced_y1,"val": rev_exced_y1, "type": "vente"},
            {"cat": "Facture avec PV",    "base": 0,                           "val": facture_avec,  "type": "résultat"},
        ])
        b0 = alt.Chart(wf).mark_bar(opacity=0).encode(
            x=alt.X("cat:N", sort=None, axis=alt.Axis(labelAngle=-10, title="")),
            y=alt.Y("base:Q", title="DT/an"),
        )
        b1 = alt.Chart(wf).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, size=52).encode(
            x=alt.X("cat:N", sort=None),
            y="base:Q", y2=alt.Y2("top:Q"),
            color=alt.Color("type:N", scale=alt.Scale(
                domain=["facture","économie","vente","résultat"],
                range=["#5a7a9a","#2ecc71","#3498db","#f5a623"]),
                legend=alt.Legend(title="Type")),
            tooltip=["cat:N", alt.Tooltip("montant:Q", format=",.0f", title="DT")],
        ).transform_calculate(top="datum.base+datum.val", montant="datum.val")
        lb = alt.Chart(wf).mark_text(dy=-10, color="#e8eaf0", fontSize=11, fontWeight="bold", font="DM Mono").encode(
            x=alt.X("cat:N", sort=None),
            y=alt.Y("top2:Q"),
            text=alt.Text("m2:Q", format=",.0f"),
        ).transform_calculate(top2="datum.base+datum.val", m2="datum.val")
        st.altair_chart((b0+b1+lb).properties(title="Décomposition facture STEG (DT/an)", height=320), use_container_width=True)

    with cr:
        lcoe_df = pd.DataFrame([
            {"label": "LCOE projet",        "value": lcoe,                   "type": "lcoe"},
            {"label": "Tarif STEG 2025",    "value": EXCEL["tarif_kwh"],     "type": "tarif"},
            {"label": "Tarif STEG +5%×10a", "value": EXCEL["tarif_kwh"]*(1.05**10), "type": "proj10"},
            {"label": "Tarif STEG +5%×20a", "value": EXCEL["tarif_kwh"]*(1.05**20), "type": "proj20"},
        ])
        lc = alt.Chart(lcoe_df).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            x=alt.X("value:Q", title="DT/kWh"),
            y=alt.Y("label:N", sort=None, title=""),
            color=alt.Color("type:N", scale=alt.Scale(
                domain=["lcoe","tarif","proj10","proj20"],
                range=["#f5a623","#e74c3c","#8e44ad","#c0392b"]), legend=None),
            tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=".4f", title="DT/kWh")],
        ).properties(title="LCOE vs Projections Tarif STEG", height=220)
        st.altair_chart(lc, use_container_width=True)
        info("Vente excédent (Y1)",     f"{rev_exced_y1:,.0f} DT")
        info("Remb. STEG dépassement",  f"{EXCEL['remb_steg']:,.0f} DT/an")
        info("Économies/kWh autoconsommé", f"{EXCEL['eco_kwh']:.5f} DT/kWh")
        info("% production cédée",      f"{EXCEL['pct_cedee']*100:.1f}%")


# ══════════════════════════════════
# TAB 2 — ENERGY & PRODUCTION
# ══════════════════════════════════
with tab2:
    st.markdown('<div class="sec">Indicateurs de Production</div>', unsafe_allow_html=True)

    c1,c2,c3,c4,c5 = st.columns(5)
    kpi(c1, "Production Y1",    f"{prod_y1:,.0f} kWh",         f"{kwc_val:.2f} kWc × {rendement} kWh/kWc",  "")
    kpi(c2, "Taux Couverture",  f"{taux_couv*100:.1f}%",        "conso couverte par PV",                      "kpi-green")
    kpi(c3, "Autoconsommée",    f"{autoconso_y1:,.0f} kWh",     "sur site Y1",                                "")
    kpi(c4, "Excédent STEG",    f"{cedee_lim_y1:,.0f} kWh",    f"limité à {limite_exced*100:.0f}%",           "kpi-blue")
    kpi(c5, "% Prod. cédée",    f"{EXCEL['pct_cedee']*100:.1f}%" if use_excel and duree==20 else f"{cedee_lim_y1/prod_y1*100:.1f}%",
        "de la production totale", "")

    df_p = pd.DataFrame({"Année": years, "Production (kWh)": prod_list})
    pb = alt.Chart(df_p).mark_bar(color="#f5a623", opacity=0.82, cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
        x=alt.X("Année:O"), y=alt.Y("Production (kWh):Q", title="kWh/an"),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Production (kWh):Q", format=",.0f")],
    )
    cr2 = alt.Chart(pd.DataFrame({"y":[conso]})).mark_rule(color="#e74c3c", strokeDash=[6,3], strokeWidth=2).encode(y="y:Q")
    cl2 = alt.Chart(pd.DataFrame({"y":[conso],"label":[f"Conso: {conso:,.0f} kWh/an"]})).mark_text(
        align="left", dx=6, dy=-8, color="#e74c3c", fontSize=11, font="DM Mono"
    ).encode(y="y:Q", text="label:N", x=alt.value(8))
    st.altair_chart((pb+cr2+cl2).properties(
        title=f"Production PV sur {duree} ans — dégradation {EXCEL['degradation']*100:.1f}%/an", height=290
    ), use_container_width=True)

    col_l, col_r = st.columns(2)
    with col_l:
        months = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
        seasonal = [0.068,0.077,0.094,0.107,0.114,0.109,0.102,0.096,0.088,0.079,0.068,0.064]
        df_mo = pd.DataFrame({
            "Mois": months, "ordre": list(range(12)),
            "Production": [prod_y1 * f for f in seasonal],
            "Consommation": [conso / 12] * 12,
        })
        mb = alt.Chart(df_mo).mark_bar(color="#f5a623", opacity=0.8).encode(
            x=alt.X("Mois:N", sort=alt.SortField("ordre"), title=""),
            y=alt.Y("Production:Q", title="kWh/mois"),
            tooltip=[alt.Tooltip("Mois:N"), alt.Tooltip("Production:Q", format=",.0f")],
        )
        ml = alt.Chart(df_mo).mark_line(color="#e74c3c", strokeDash=[4,2], strokeWidth=2).encode(
            x=alt.X("Mois:N", sort=alt.SortField("ordre")),
            y="Consommation:Q",
            tooltip=[alt.Tooltip("Mois:N"), alt.Tooltip("Consommation:Q", format=",.0f")],
        )
        st.altair_chart((mb+ml).properties(title="Profil mensuel Production vs Consommation", height=260), use_container_width=True)

    with col_r:
        a_pct = autoconso_y1 / prod_y1 * 100
        c_pct = cedee_lim_y1 / prod_y1 * 100
        p_pct = max(0, 100 - a_pct - c_pct)
        pie_df = pd.DataFrame([
            {"label": f"Autoconsommée ({a_pct:.1f}%)",  "value": round(a_pct,1)},
            {"label": f"Vendue STEG ({c_pct:.1f}%)",    "value": round(c_pct,1)},
            {"label": f"Non valorisée ({p_pct:.1f}%)",  "value": round(p_pct,1)},
        ])
        pie = alt.Chart(pie_df).mark_arc(innerRadius=65, outerRadius=105).encode(
            theta=alt.Theta("value:Q"),
            color=alt.Color("label:N", scale=alt.Scale(
                domain=pie_df["label"].tolist(), range=["#f5a623","#2ecc71","#e74c3c"]),
                legend=alt.Legend(title="Répartition")),
            tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=".1f", title="%")],
        ).properties(title="Répartition de la production PV (%)", height=260)
        st.altair_chart(pie, use_container_width=True)

    st.markdown('<div class="sec">Paramètres Techniques</div>', unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    tech = [
        ("Région", EXCEL["region"]),
        ("Orientation", EXCEL["orientation"]),
        ("Rendement spécifique", f"{EXCEL['rendement']:,} kWh/kWc/an"),
        ("Dégradation", f"{EXCEL['degradation']*100:.1f}%/an"),
        ("Production totale 20 ans", f"{total_prod/1000:,.1f} MWh"),
        ("Bilan instantané", "Net-metering"),
        ("Tarif transport STEG", f"{EXCEL['tarif_transport']:.3f} DT/kWh"),
        ("Limite excédent annuel", f"{EXCEL['limite_exced']*100:.0f}%"),
        ("Remboursement STEG", f"{EXCEL['remb_steg']:,.0f} DT/an"),
    ]
    for i,(k,v) in enumerate(tech):
        with [c1,c2,c3][i%3]: info(k, v)


# ══════════════════════════════════
# TAB 3 — CASH FLOWS
# ══════════════════════════════════
with tab3:
    st.markdown('<div class="sec">Flux de Trésorerie Annuels</div>', unsafe_allow_html=True)

    df_cf = pd.DataFrame({
        "Année": years,
        "Revenus": rev_list, "OpEx": [-o for o in opex_list],
        "Dette": [-d for d in ds_list], "Cumul": cum_list,
    })
    melt_cf = df_cf.melt("Année", value_vars=["Revenus","OpEx","Dette"], var_name="Poste", value_name="DT")
    bars_cf = alt.Chart(melt_cf).mark_bar().encode(
        x=alt.X("Année:O"), y=alt.Y("DT:Q", title="DT/an", stack="zero"),
        color=alt.Color("Poste:N", scale=alt.Scale(
            domain=["Revenus","OpEx","Dette"], range=["#f5a623","#e74c3c","#8e44ad"])),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Poste:N"), alt.Tooltip("DT:Q", format=",.0f")],
    )
    cum_ln = alt.Chart(df_cf).mark_line(color="#2ecc71", strokeWidth=2.5, point=True).encode(
        x=alt.X("Année:O"),
        y=alt.Y("Cumul:Q", axis=alt.Axis(title="Cumul Equity DT", orient="right")),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Cumul:Q", format=",.0f")],
    )
    z_rule = alt.Chart(pd.DataFrame({"y":[0]})).mark_rule(color="#5a7a9a", strokeDash=[4,2]).encode(y="y:Q")
    st.altair_chart(alt.layer(bars_cf, cum_ln+z_rule).resolve_scale(y="independent").properties(
        title=f"Revenus, coûts et cumul equity ({duree} ans)", height=370
    ), use_container_width=True)

    # CFADS vs Debt Service
    st.markdown('<div class="sec">CFADS vs Service de la Dette</div>', unsafe_allow_html=True)
    df_dscr = pd.DataFrame({
        "Année": years,
        "CFADS": cfads_list,
        "Service Dette": ds_list,
    }).melt("Année", var_name="Poste", value_name="DT")
    dscr_ch = alt.Chart(df_dscr).mark_line(point=True, strokeWidth=2).encode(
        x=alt.X("Année:O"), y=alt.Y("DT:Q", title="DT/an"),
        color=alt.Color("Poste:N", scale=alt.Scale(
            domain=["CFADS","Service Dette"], range=["#2ecc71","#e74c3c"])),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Poste:N"), alt.Tooltip("DT:Q", format=",.0f")],
    ).properties(title="CFADS vs Service de la dette — couverture annuelle", height=250)
    st.altair_chart(dscr_ch, use_container_width=True)

    # Sensitivity
    st.markdown('<div class="sec">Analyse de Sensibilité</div>', unsafe_allow_html=True)
    def quick_irr(inv, rvs, om_r, infl, n):
        cfs = [-inv] + [r - inv*om_r*(1+infl)**i for i,r in enumerate(rvs[:n])]
        lo, hi = -0.5, 5.0
        for _ in range(80):
            mid=(lo+hi)/2
            if sum(c/(1+mid)**t for t,c in enumerate(cfs)) > 0: lo=mid
            else: hi=mid
        return (lo+hi)/2*100

    cs1, cs2 = st.columns(2)
    with cs1:
        costs_r = np.linspace(1200, 4000, 22)
        irrs_c  = [quick_irr(kwc_val*c*(1+tva_frac)-sub_dt, rev_list, om_pct, inflation, min(duree,20)) for c in costs_r]
        df_s1 = pd.DataFrame({"Coût DT/kWc": costs_r, "TRI %": irrs_c})
        l1 = alt.Chart(df_s1).mark_line(color="#f5a623", strokeWidth=2).encode(
            x=alt.X("Coût DT/kWc:Q"), y=alt.Y("TRI %:Q"),
            tooltip=[alt.Tooltip("Coût DT/kWc:Q", format=",.0f"), alt.Tooltip("TRI %:Q", format=".2f")],
        )
        a1 = alt.Chart(df_s1).mark_area(color="#f5a623", opacity=0.1).encode(x="Coût DT/kWc:Q", y="TRI %:Q", y2=alt.value(0))
        r1 = alt.Chart(pd.DataFrame({"y":[taux_disc*100]})).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
        v1 = alt.Chart(pd.DataFrame({"x":[prix_kwc]})).mark_rule(color="#2ecc71", strokeDash=[4,2]).encode(x="x:Q")
        st.altair_chart((a1+l1+r1+v1).properties(title="TRI vs Coût installation (DT/kWc)", height=250), use_container_width=True)

    with cs2:
        tariffs_r = np.linspace(0, 0.12, 22)
        irrs_t = []
        for ht in tariffs_r:
            rvt = [r*(1+ht)**max(0,y-2) for y,r in zip(years[:min(duree,20)], rev_list[:min(duree,20)])]
            irrs_t.append(quick_irr(EXCEL["cout_ttc"] if use_excel and duree==20 else cout_net, rvt, om_pct, inflation, min(duree,20)))
        df_s2 = pd.DataFrame({"Hausse %/an": tariffs_r*100, "TRI %": irrs_t})
        l2 = alt.Chart(df_s2).mark_line(color="#2ecc71", strokeWidth=2).encode(
            x=alt.X("Hausse %/an:Q"), y=alt.Y("TRI %:Q"),
            tooltip=[alt.Tooltip("Hausse %/an:Q", format=".1f"), alt.Tooltip("TRI %:Q", format=".2f")],
        )
        a2 = alt.Chart(df_s2).mark_area(color="#2ecc71", opacity=0.1).encode(x="Hausse %/an:Q", y="TRI %:Q", y2=alt.value(0))
        r2 = alt.Chart(pd.DataFrame({"y":[taux_disc*100]})).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
        v2 = alt.Chart(pd.DataFrame({"x":[hausse_tarif*100]})).mark_rule(color="#f5a623", strokeDash=[4,2]).encode(x="x:Q")
        st.altair_chart((a2+l2+r2+v2).properties(title="TRI vs Hausse tarif STEG (%/an)", height=250), use_container_width=True)

    with st.expander("📋 Tableau complet des flux de trésorerie"):
        df_tbl = pd.DataFrame({
            "Année":              years,
            "Production (kWh)":  [round(p) for p in prod_list],
            "Revenus (DT)":      [round(r) for r in rev_list],
            "OpEx (DT)":         [round(o) for o in opex_list],
            "Service dette (DT)":[round(d) for d in ds_list],
            "CFADS (DT)":        [round(c) for c in cfads_list],
            "CF Equity (DT)":    [round(e) for e in eq_cf_list],
            "Cumul Equity (DT)": [round(c) for c in cum_list],
        })
        def csign(val):
            if isinstance(val, (int,float)):
                if val < 0: return "background-color:rgba(231,76,60,0.25);color:#ffaaaa;"
                if val > 0: return "background-color:rgba(46,204,113,0.20);color:#aaffbb;"
            return ""
        st.dataframe(
            df_tbl.style.format({
                "Production (kWh)":"{:,.0f}", "Revenus (DT)":"{:,.0f}",
                "OpEx (DT)":"{:,.0f}", "Service dette (DT)":"{:,.0f}",
                "CFADS (DT)":"{:,.0f}", "CF Equity (DT)":"{:,.0f}", "Cumul Equity (DT)":"{:,.0f}",
            }).map(csign, subset=["CF Equity (DT)","Cumul Equity (DT)","CFADS (DT)"]),
            use_container_width=True,
        )


# ══════════════════════════════════
# TAB 4 — INVESTMENT & FINANCING
# ══════════════════════════════════
with tab4:
    st.markdown('<div class="sec">Structure de l\'Investissement</div>', unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    kpi(c1, "Coût HT",          f"{EXCEL['cout_ht'] if use_excel and duree==20 else cout_ht:,.0f} DT",
        f"{kwc_val:.2f} kWc × {EXCEL['prix_kwc']:.0f} DT/kWc", "")
    kpi(c2, "TVA (13%)",        f"{EXCEL['tva_dt']:,.0f} DT",       "Taxe valeur ajoutée",          "")
    kpi(c3, "Coût TTC Net",     f"{EXCEL['cout_ttc'] if use_excel and duree==20 else cout_net:,.0f} DT",
        f"Subvention: {EXCEL['sub_dt']:,.0f} DT", "kpi-blue")
    kpi(c4, "Capitaux Propres", f"{EXCEL['cap_propres'] if use_excel and duree==20 else cap_propres:,.0f} DT",
        "Reste à charge propriétaire", "")

    cf1, cf2 = st.columns(2)
    with cf1:
        cp = EXCEL["cap_propres"] if use_excel and duree==20 else cap_propres
        md = EXCEL["montant_dette"] if use_excel and duree==20 else montant_dette
        fd = pd.DataFrame([
            {"label": f"Capitaux propres\n{cp:,.0f} DT",  "value": cp},
            {"label": f"Dette bancaire\n{md:,.0f} DT",    "value": md},
        ])
        fp = alt.Chart(fd).mark_arc(innerRadius=60, outerRadius=105).encode(
            theta=alt.Theta("value:Q"),
            color=alt.Color("label:N", scale=alt.Scale(
                domain=fd["label"].tolist(), range=["#f5a623","#3498db"]),
                legend=alt.Legend(title="Financement")),
            tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=",.0f", title="DT")],
        ).properties(title="Structure de financement", height=280)
        st.altair_chart(fp, use_container_width=True)

        fp_params = [
            ("Capitaux propres",        f"{cp:,.0f} DT ({(1-EXCEL['dette_pct'])*100:.0f}%)"),
            ("Dette bancaire",          f"{md:,.0f} DT ({EXCEL['dette_pct']*100:.0f}% — plafond 200 000 DT)"),
            ("Taux intérêt effectif",   f"{EXCEL['taux_interet']*100:.1f}% (BCT)"),
            ("Durée du crédit",         f"{EXCEL['duree_credit']} ans"),
            ("Délai de grâce",          f"{EXCEL['grace']} an"),
            ("Commission initiale",     "1,0 % du crédit"),
            ("Commission d'engagement", "0,25 % de la marge"),
            ("Taux d'actualisation",    f"{EXCEL['taux_disc']*100:.0f}%"),
        ]
        for k,v in fp_params: info(k, v)

    with cf2:
        out_b = EXCEL["montant_dette"] if use_excel and duree==20 else montant_dette
        ti    = EXCEL["taux_interet"] if use_excel and duree==20 else (taux_ref+marge)
        dc_y  = EXCEL["duree_credit"] if use_excel and duree==20 else duree_credit
        gc_y  = EXCEL["grace"] if use_excel and duree==20 else grace
        md_v  = EXCEL["montant_dette"] if use_excel and duree==20 else montant_dette
        dsched = []
        for y in years:
            if y <= gc_y:
                intY=out_b*ti; prY=0.0
            elif y <= dc_y:
                prY = md_v / max(dc_y-gc_y,1); intY = out_b*ti; out_b=max(0,out_b-prY)
            else:
                prY=0.0; intY=0.0
            dsched.append({"Année":y,"Principal":prY,"Intérêts":intY,"Encours":out_b})
        df_ds = pd.DataFrame(dsched)
        df_dm = df_ds.melt("Année", value_vars=["Principal","Intérêts"], var_name="Poste", value_name="DT")
        db = alt.Chart(df_dm).mark_bar(cornerRadiusTopLeft=2,cornerRadiusTopRight=2).encode(
            x=alt.X("Année:O"), y=alt.Y("DT:Q", title="DT/an", stack="zero"),
            color=alt.Color("Poste:N", scale=alt.Scale(
                domain=["Principal","Intérêts"], range=["#3498db","#8e44ad"])),
            tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Poste:N"), alt.Tooltip("DT:Q", format=",.0f")],
        )
        el = alt.Chart(df_ds).mark_line(color="#e74c3c", strokeWidth=2, strokeDash=[4,2]).encode(
            x=alt.X("Année:O"),
            y=alt.Y("Encours:Q", axis=alt.Axis(title="Encours DT", orient="right")),
            tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Encours:Q", format=",.0f")],
        )
        st.altair_chart(alt.layer(db, el).resolve_scale(y="independent").properties(
            title="Remboursement dette bancaire + encours", height=340
        ), use_container_width=True)


# ══════════════════════════════════
# TAB 5 — BANKING RATIOS
# ══════════════════════════════════
with tab5:
    st.markdown('<div class="sec">Ratios Bancaires — Indicateurs de Solvabilité</div>', unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    kpi(c1, "DSCR Minimum",      f"{dscr_min:.4f}",            "Ratio couverture service dette",   "kpi-green" if dscr_min >= 1.2 else "kpi-warn")
    kpi(c2, "LLCR Minimum",      f"{llcr_min:.4f}",            "Ratio couverture durée du prêt",   "kpi-green" if llcr_min >= 1.3 else "kpi-warn")
    kpi(c3, "TRI Equity",        f"{eq_irr_pct:.2f}%",         "Retour capitaux propres",          "kpi-green")
    kpi(c4, "VAN",               f"{project_npv/1000:+.1f}k DT", f"Taux disc. {taux_disc*100:.0f}%","kpi-blue" if project_npv>0 else "kpi-red")

    cd1, cd2 = st.columns(2)
    with cd1:
        dscr_rows = []
        for y, cf, ds in zip(years, cfads_list, ds_list):
            if ds > 0: dscr_rows.append({"Année": y, "DSCR": cf/ds})
        if dscr_rows:
            df_dscr = pd.DataFrame(dscr_rows)
            df_dscr["Statut"] = df_dscr["DSCR"].apply(
                lambda v: ">=1.2 Sûr" if v >= 1.2 else (">=1.0 Limite" if v >= 1.0 else "<1.0 Risque")
            )
            dscr_b = alt.Chart(df_dscr).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
                x=alt.X("Année:O"), y=alt.Y("DSCR:Q", title="DSCR"),
                color=alt.Color("Statut:N", scale=alt.Scale(
                    domain=[">=1.2 Sûr", ">=1.0 Limite", "<1.0 Risque"],
                    range=["#2ecc71", "#f5a623", "#e74c3c"]
                ), legend=alt.Legend(title="Couverture")),
                tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("DSCR:Q", format=".4f"), alt.Tooltip("Statut:N")],
            )
            r12 = alt.Chart(pd.DataFrame({"y":[1.2]})).mark_rule(color="#2ecc71", strokeDash=[4,2]).encode(y="y:Q")
            r10 = alt.Chart(pd.DataFrame({"y":[1.0]})).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
            st.altair_chart((dscr_b+r12+r10).properties(
                title="DSCR par année (vert ≥ 1.2 · orange ≥ 1.0 · rouge < 1.0)", height=290
            ), use_container_width=True)

    with cd2:
        inv0 = EXCEL["cout_ttc"] if use_excel and duree==20 else cout_net
        cf_proj = [-inv0] + cfads_list
        cum_proj = list(np.cumsum(cf_proj))
        df_cp = pd.DataFrame({"Période": list(range(duree+1)), "Cumul projet (DT)": cum_proj})
        cl_p = alt.Chart(df_cp).mark_line(color="#f5a623", strokeWidth=2.5).encode(
            x=alt.X("Période:O"), y=alt.Y("Cumul projet (DT):Q"),
            tooltip=[alt.Tooltip("Période:O"), alt.Tooltip("Cumul projet (DT):Q", format=",.0f")],
        )
        ar_p = alt.Chart(df_cp).mark_area(color="#f5a623", opacity=0.1).encode(
            x="Période:O", y="Cumul projet (DT):Q", y2=alt.value(0)
        )
        zr_p = alt.Chart(pd.DataFrame({"y":[0]})).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
        st.altair_chart((ar_p+cl_p+zr_p).properties(
            title=f"Flux cumulés projet — payback {payback_val:.2f} ans", height=290
        ), use_container_width=True)

    st.markdown('<div class="sec">Interprétation des Ratios</div>', unsafe_allow_html=True)
    cr1, cr2 = st.columns(2)
    ratios = [
        ("DSCR",        f"Min = {dscr_min:.4f} — Cible bancaire ≥ 1.2 · {'✅ OK' if dscr_min>=1.2 else '⚠️ Surveiller'}"),
        ("LLCR",        f"Min = {llcr_min:.4f} — Cible bancaire ≥ 1.3 · {'✅ OK' if llcr_min>=1.3 else '⚠️ Surveiller'}"),
        ("TRI Equity",  f"{eq_irr_pct:.2f}% > taux disc. {taux_disc*100:.1f}% → ✅ Projet viable pour l'investisseur"),
        ("TRI Projet",  f"{irr_pct:.2f}% — Comparaison avec taux sans risque (BCT {EXCEL['taux_interet']*100:.0f}%)"),
        ("Payback",     f"{payback_val:.2f} ans < durée crédit ({EXCEL['duree_credit']} ans) → ✅ Remboursé avant fin prêt"),
        ("LCOE",        f"{lcoe:.4f} DT/kWh < Tarif STEG {EXCEL['tarif_kwh']:.3f} DT/kWh → ✅ PV moins cher que réseau"),
    ]
    for i,(k,v) in enumerate(ratios):
        with cr1 if i%2==0 else cr2: info(k, v)


# ══════════════════════════════════
# TAB 6 — ENVIRONMENTAL IMPACT
# ══════════════════════════════════
with tab6:
    st.markdown('<div class="sec">Impact Environnemental</div>', unsafe_allow_html=True)

    trees = co2_total * 40
    cars  = co2_total / 2.3
    homes = total_prod / 4000

    c1,c2,c3,c4,c5 = st.columns(5)
    kpi(c1, f"CO₂ Évité {duree} ans", f"{co2_total:,.0f} t",        "tonnes éq. CO₂",          "kpi-green")
    kpi(c2, "Énergie Propre",         f"{total_prod/1000:,.1f} MWh",  f"produite sur {duree} ans","")
    kpi(c3, "Équivalent Arbres",      f"{trees:,.0f}",                "arbres/an équivalent",    "kpi-green")
    kpi(c4, "Voitures retirées",      f"{cars:,.0f}",                 "voitures/an équivalent",  "")
    kpi(c5, "Foyers alimentés",       f"{homes:,.0f}",                "foyers (4 MWh/foyer/an)", "kpi-blue")

    df_co2 = pd.DataFrame({"Année": years, "CO₂ évité (t)": co2_list, "CO₂ cumulé (t)": co2_cum})
    co2_b = alt.Chart(df_co2).mark_bar(color="#2ecc71", opacity=0.75, cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
        x=alt.X("Année:O"), y=alt.Y("CO₂ évité (t):Q", title="t CO₂/an"),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("CO₂ évité (t):Q", format=".2f")],
    )
    co2_cl = alt.Chart(df_co2).mark_line(color="#f5a623", strokeWidth=2.5).encode(
        x=alt.X("Année:O"),
        y=alt.Y("CO₂ cumulé (t):Q", axis=alt.Axis(title="t CO₂ cumulé", orient="right")),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("CO₂ cumulé (t):Q", format=".1f")],
    )
    st.altair_chart(alt.layer(co2_b, co2_cl).resolve_scale(y="independent").properties(
        title="CO₂ évité par an et cumulé — facteur 0,57 t CO₂/MWh (réseau Tunisie)", height=340
    ), use_container_width=True)

    st.markdown('<div class="sec">Mise en Contexte</div>', unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    ctx = [
        ("Facteur émission Tunisie",  "0,57 t CO₂ évitées par MWh solaire"),
        ("CO₂ évité — Année 1",       f"{co2_list[0]:.2f} tonnes"),
        ("CO₂ évité — Année 20",      f"{co2_list[-1]:.2f} tonnes (dégradation incluse)"),
        ("CO₂ total évité",           f"{co2_total:,.1f} tonnes = {EXCEL['co2_total']:,.2f} t (Excel)"),
        ("Équivalent arbres",         f"1 arbre ≈ 25 kg CO₂/an → {int(trees):,} arbres"),
        ("Équivalent voitures",       f"Voiture ≈ 2,3 t CO₂/an → {int(cars):,} voitures retirées/an"),
        ("Foyers tunisiens",          f"Conso moy. 4 MWh/foyer → {int(homes):,} foyers alimentés"),
        ("Production totale 20 ans",  f"{total_prod/1000:,.1f} MWh"),
        ("Production Y1",             f"{prod_y1:,.0f} kWh"),
    ]
    for i,(k,v) in enumerate(ctx):
        with [c1,c2,c3][i%3]: info(k, v)

# ══════════════════════════════════
# TAB 7 — 3D INSTALLATION VIEW (NEW)
# ══════════════════════════════════
with tab7:
    st.markdown('<div class="sec">Vue 3D de l\'Installation PV</div>', unsafe_allow_html=True)

    n_strings = 16
    panels_per_string = 15
    n_panels = n_strings * panels_per_string
    panel_wp = round((kwc * 1000) / n_panels)

    c1, c2, c3, c4 = st.columns(4)
    pv3d_kpis = [
        (c1, "Panneaux total",   f"{n_panels}",               "unités",                          ""),
        (c2, "Strings",          f"{n_strings}",               f"× {panels_per_string} panneaux/string", "kpi-blue"),
        (c3, "Puissance/panneau",f"{panel_wp} Wc",             f"{kwc:.2f} kWc total",             ""),
        (c4, "Surface estimée",  f"~{n_panels*5.6:,.0f} m²",   "à 5,6 m²/panneau",                 ""),
    ]
    for col, label, val, unit, cls in pv3d_kpis:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    pv3d_html = f"""
<div style="font-family:'Syne',sans-serif;color:#e8eaf0;">
<style>
#wrap3d{{width:100%;user-select:none;background:#0d1117;border-radius:12px;padding:16px;}}
#hdr3d{{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px;}}
#hdr3d h2{{font-size:15px;font-weight:700;color:#f5a623;margin:0;}}
.badge3d{{font-size:11px;padding:3px 9px;border-radius:8px;border:1px solid #1e2d3d;color:#8aa4be;font-family:'DM Mono',monospace;}}
#controls3d{{display:flex;gap:18px;align-items:center;flex-wrap:wrap;margin-bottom:14px;}}
.ctrl3d{{display:flex;align-items:center;gap:8px;font-size:12px;color:#8aa4be;font-family:'DM Mono',monospace;}}
.ctrl3d label{{min-width:75px;}}
.ctrl3d span{{min-width:30px;font-weight:700;color:#f5a623;font-size:12px;}}
.ctrl3d input[type=range]{{accent-color:#f5a623;}}
#scene-wrap3d{{width:100%;overflow:hidden;border-radius:10px;border:1px solid #1e2d3d;background:#e5e7eb;position:relative;height:460px;}}
#cv3d{{display:block;width:100%;height:100%;}}
#tooltip3d{{position:absolute;pointer-events:none;background:#111820;border:1px solid #f5a623;border-radius:8px;padding:8px 12px;font-size:12px;color:#e8eaf0;display:none;min-width:150px;z-index:10;font-family:'DM Mono',monospace;}}
#tooltip3d b{{color:#f5a623;}}
#legend3d{{display:flex;gap:16px;margin-top:12px;flex-wrap:wrap;}}
.leg3d{{display:flex;align-items:center;gap:6px;font-size:11px;color:#8aa4be;font-family:'DM Mono',monospace;}}
.leg-sq3d{{width:12px;height:8px;border-radius:2px;}}
</style>

<div id="wrap3d">
  <div id="hdr3d">
    <h2>Vue isométrique interactive</h2>
    <span class="badge3d">{kwc:.2f} kWc · {n_strings} strings · {panels_per_string} panneaux/string</span>
  </div>

  <div id="controls3d">
    <div class="ctrl3d"><label>Angle soleil</label><input type="range" id="sun3d" min="0" max="180" value="60" step="1" style="width:100px"><span id="sun-v3d">60°</span></div>
    <div class="ctrl3d"><label>Rotation vue</label><input type="range" id="rot3d" min="-60" max="60" value="0" step="1" style="width:100px"><span id="rot-v3d">0°</span></div>
    <div class="ctrl3d"><label>Élévation</label><input type="range" id="elev3d" min="20" max="60" value="35" step="1" style="width:100px"><span id="elev-v3d">35°</span></div>
  </div>

  <div id="scene-wrap3d">
    <canvas id="cv3d"></canvas>
    <div id="tooltip3d"></div>
  </div>

  <div id="legend3d">
    <div class="leg3d"><div class="leg-sq3d" style="background:#2563a8"></div>Face PV (active)</div>
    <div class="leg3d"><div class="leg-sq3d" style="background:#1a3d6e"></div>Face PV (ombre)</div>
    <div class="leg3d"><div class="leg-sq3d" style="background:#4b5563"></div>Cadre aluminium</div>
    <div class="leg3d"><div class="leg-sq3d" style="background:#d97706"></div>String sélectionné</div>
  </div>
</div>
</div>

<script>
(function(){{
const cv = document.getElementById('cv3d');
const ctx = cv.getContext('2d');
const wrap = document.getElementById('scene-wrap3d');
const tip = document.getElementById('tooltip3d');

const STRINGS = {n_strings};
const PANELS_PER_STRING = {panels_per_string};
const PANEL_WP = {panel_wp};
const PW = 1.0, PH = 1.65, TILT = 32;
const GAP_X = 0.15, GAP_Y = 0.35, GROUP_GAP = 0.55;

let sunAngle = 60, rotY = 0, elevAngle = 35;
let selectedString = -1, hoveredPanel = null;
let W, H;

function resize(){{
  W = wrap.clientWidth; H = wrap.clientHeight;
  const dpr = window.devicePixelRatio || 1;
  cv.width = W * dpr; cv.height = H * dpr;
  cv.style.width = W+'px'; cv.style.height = H+'px';
  ctx.setTransform(1,0,0,1,0,0);
  ctx.scale(dpr, dpr);
  draw();
}}

function project(x, y, z){{
  const rRot = rotY * Math.PI/180;
  const rElev = elevAngle * Math.PI/180;
  const rx = x * Math.cos(rRot) + z * Math.sin(rRot);
  const ry = y;
  const rz = -x * Math.sin(rRot) + z * Math.cos(rRot);
  const sx = rx - rz * Math.cos(Math.PI/6);
  const sy = -ry - (rx + rz) * Math.sin(rElev * 0.6);
  const scale = Math.min(W, H) * 0.062;
  return {{ x: W/2 + sx * scale, y: H * 0.58 + sy * scale }};
}}

function getSunDir(){{
  const az = sunAngle * Math.PI/180;
  return {{ x: Math.cos(az), y: Math.abs(Math.sin(az))*0.9+0.1, z: Math.sin(az)*0.5 }};
}}

function panelNormal(tiltRad){{ return {{ x:0, y:Math.cos(tiltRad), z:Math.sin(tiltRad) }}; }}
function dot(a,b){{ return a.x*b.x+a.y*b.y+a.z*b.z; }}

function lerpColor(a,b,t){{
  const ah=parseInt(a.slice(1),16), bh=parseInt(b.slice(1),16);
  const r=Math.round(((ah>>16)&255)*(1-t)+((bh>>16)&255)*t);
  const g=Math.round(((ah>>8)&255)*(1-t)+((bh>>8)&255)*t);
  const bl2=Math.round((ah&255)*(1-t)+(bh&255)*t);
  return `rgb(${{r}},${{g}},${{bl2}})`;
}}
function hexToRgb(h,a){{
  a = a===undefined?1:a;
  const n=parseInt(h.slice(1),16);
  return `rgba(${{(n>>16)&255}},${{(n>>8)&255}},${{n&255}},${{a}})`;
}}

function panelCorners(sx, sz, tiltRad){{
  const cosT=Math.cos(tiltRad), sinT=Math.sin(tiltRad);
  const hw=PW/2, hh=PH/2;
  return [
    {{x:sx-hw, y:-hh*cosT, z:sz+hh*sinT}},
    {{x:sx+hw, y:-hh*cosT, z:sz+hh*sinT}},
    {{x:sx+hw, y: hh*cosT, z:sz-hh*sinT}},
    {{x:sx-hw, y: hh*cosT, z:sz-hh*sinT}},
  ];
}}

function drawPolygon(pts3d, fill, stroke, strokeW){{
  strokeW = strokeW===undefined?0.5:strokeW;
  const pts = pts3d.map(p=>project(p.x,p.y,p.z));
  ctx.beginPath(); ctx.moveTo(pts[0].x,pts[0].y);
  for(let i=1;i<pts.length;i++) ctx.lineTo(pts[i].x,pts[i].y);
  ctx.closePath();
  ctx.fillStyle = fill; ctx.fill();
  if(stroke){{ ctx.strokeStyle=stroke; ctx.lineWidth=strokeW; ctx.stroke(); }}
}}

function drawLine(a3,b3,color,w){{
  w = w===undefined?0.8:w;
  const a=project(a3.x,a3.y,a3.z), b=project(b3.x,b3.y,b3.z);
  ctx.beginPath(); ctx.moveTo(a.x,a.y); ctx.lineTo(b.x,b.y);
  ctx.strokeStyle=color; ctx.lineWidth=w; ctx.stroke();
}}

function lerp2D(a,b,t){{ return {{x:a.x+(b.x-a.x)*t, y:a.y+(b.y-a.y)*t}}; }}

function pointInPoly(px,py,pts){{
  let inside=false;
  for(let i=0,j=pts.length-1;i<pts.length;j=i++){{
    const xi=pts[i].x, yi=pts[i].y, xj=pts[j].x, yj=pts[j].y;
    if(((yi>py)!=(yj>py)) && (px<(xj-xi)*(py-yi)/(yj-yi)+xi)) inside=!inside;
  }}
  return inside;
}}

function draw(){{
  ctx.clearRect(0,0,W,H);
  const tiltRad = TILT*Math.PI/180;
  const sun = getSunDir();
  const norm = panelNormal(tiltRad);
  const diffuse = Math.max(0, dot(sun, norm));

  const groundPts = [project(-12,0,-3), project(12,0,-3), project(12,0,22), project(-12,0,22)];
  ctx.beginPath(); ctx.moveTo(groundPts[0].x,groundPts[0].y);
  groundPts.forEach(p=>ctx.lineTo(p.x,p.y));
  ctx.closePath(); ctx.fillStyle='#e5e7eb'; ctx.fill();

  ctx.strokeStyle='#d1d5db'; ctx.lineWidth=0.4;
  for(let xi=-12; xi<=12; xi+=2){{
    const a=project(xi,0,-3), b=project(xi,0,22);
    ctx.beginPath(); ctx.moveTo(a.x,a.y); ctx.lineTo(b.x,b.y); ctx.stroke();
  }}
  for(let zi=-3; zi<=22; zi+=2){{
    const a=project(-12,0,zi), b=project(12,0,zi);
    ctx.beginPath(); ctx.moveTo(a.x,a.y); ctx.lineTo(b.x,b.y); ctx.stroke();
  }}

  for(let grp=0; grp<4; grp++){{
    for(let col=0; col<4; col++){{
      const si = grp*4+col;
      const gapOffset = grp*(4*(PW+GAP_X)+GROUP_GAP);
      const sx = gapOffset + col*(PW+GAP_X) - 9.2;
      const isSelected = si === selectedString;

      for(let row=0; row<PANELS_PER_STRING; row++){{
        const sz = row*(PH*Math.cos(tiltRad)+GAP_Y)+0.5;
        const corners = panelCorners(sx+PW/2, sz, tiltRad);
        const topZ = sz + PH/2*Math.sin(tiltRad);
        const topY = -PH/2*Math.cos(tiltRad);

        const faceLight = 0.35 + diffuse*0.65;
        let panelColor, frameColor;
        if(isSelected){{
          panelColor = lerpColor('#a16207','#fef3c7', faceLight);
          frameColor = '#d97706';
        }} else {{
          panelColor = lerpColor('#1a3d6e','#5b9bd5', faceLight);
          frameColor = '#4b5563';
        }}

        const backCorners = corners.map(c=>({{...c, y:c.y+0.03, z:c.z-0.03}}));
        drawPolygon(backCorners, '#1a1a2e', null);
        drawPolygon(corners, panelColor, hexToRgb(frameColor,0.6), 0.6);

        if(!isSelected){{
          const p = corners.map(c=>project(c.x,c.y,c.z));
          ctx.save();
          ctx.beginPath(); ctx.moveTo(p[0].x,p[0].y);
          p.forEach(pt=>ctx.lineTo(pt.x,pt.y));
          ctx.closePath(); ctx.clip();
          ctx.strokeStyle = hexToRgb('#1a3d6e',0.4); ctx.lineWidth=0.4;
          for(let li=1; li<6; li++){{
            const t=li/6;
            const la=lerp2D(p[0],p[3],t), lb=lerp2D(p[1],p[2],t);
            ctx.beginPath(); ctx.moveTo(la.x,la.y); ctx.lineTo(lb.x,lb.y); ctx.stroke();
          }}
          for(let li=1; li<10; li++){{
            const t=li/10;
            const la=lerp2D(p[0],p[1],t), lb=lerp2D(p[3],p[2],t);
            ctx.beginPath(); ctx.moveTo(la.x,la.y); ctx.lineTo(lb.x,lb.y); ctx.stroke();
          }}
          ctx.restore();
        }}

        drawPolygon(corners, 'rgba(0,0,0,0)', frameColor, 1.0);

        if(row===0 || row===PANELS_PER_STRING-1){{
          const postY = 0.5;
          drawLine({{x:sx,y:postY,z:topZ}}, {{x:sx,y:topY,z:topZ}}, '#6b7280', 1.2);
          drawLine({{x:sx+PW,y:postY,z:topZ}}, {{x:sx+PW,y:topY,z:topZ}}, '#6b7280', 1.2);
        }}
        if(row===0){{
          drawLine({{x:sx,y:topY+0.1,z:topZ}}, {{x:sx+PW,y:topY+0.1,z:topZ}}, '#9ca3af', 0.8);
        }}
      }}
    }}
  }}

  for(let si=0; si<STRINGS; si++){{
    const grp=Math.floor(si/4), col=si%4;
    const gapOffset = grp*(4*(PW+GAP_X)+GROUP_GAP);
    const sx = gapOffset + col*(PW+GAP_X) - 9.2 + PW/2;
    const labelPt = project(sx, 0.1, -1.5);
    ctx.fillStyle = si===selectedString ? '#d97706' : '#6b7280';
    ctx.font = (si===selectedString?'bold 11px':'bold 9px') + ` 'DM Mono', monospace`;
    ctx.textAlign = 'center';
    ctx.fillText('S'+(si+1), labelPt.x, labelPt.y);
  }}

  const invPts = [{{x:-11,y:0,z:8}},{{x:-10,y:0,z:8}},{{x:-10,y:0,z:10}},{{x:-11,y:0,z:10}}];
  drawPolygon(invPts, '#374151', '#1f2937', 1);
  const invTop = invPts.map(p=>({{...p, y:-1.2}}));
  drawPolygon(invTop, '#4b5563', '#374151', 1);
  drawPolygon([invPts[0],invPts[1],invTop[1],invTop[0]], '#374151', '#1f2937', 0.5);
  drawPolygon([invPts[1],invPts[2],invTop[2],invTop[1]], '#4b5563', '#374151', 0.5);
  const invLabel = project(-10.5, -1.5, 9);
  ctx.fillStyle='#9ca3af'; ctx.font = "9px 'DM Mono', monospace"; ctx.textAlign='center';
  ctx.fillText('Onduleur', invLabel.x, invLabel.y);
  ctx.fillText('SG110CX×2', invLabel.x, invLabel.y+11);

  for(let si=0; si<STRINGS; si++){{
    const grp=Math.floor(si/4), col=si%4;
    const gapOffset = grp*(4*(PW+GAP_X)+GROUP_GAP);
    const sx = gapOffset + col*(PW+GAP_X) - 9.2;
    drawLine({{x:sx,y:-0.05,z:1.0}}, {{x:-10.5,y:-0.05,z:9}}, hexToRgb('#6b7280',0.3), 0.5);
  }}
}}

cv.addEventListener('mousemove', function(e){{
  const rect = cv.getBoundingClientRect();
  const mx = e.clientX - rect.left;
  const my = e.clientY - rect.top;
  const tiltRad = TILT*Math.PI/180;
  let found = null;
  for(let si=STRINGS-1; si>=0 && !found; si--){{
    for(let row=PANELS_PER_STRING-1; row>=0; row--){{
      const grp=Math.floor(si/4), col=si%4;
      const gapOffset = grp*(4*(PW+GAP_X)+GROUP_GAP);
      const sx = gapOffset + col*(PW+GAP_X) - 9.2;
      const sz = row*(PH*Math.cos(tiltRad)+GAP_Y)+0.5;
      const corners = panelCorners(sx+PW/2, sz, tiltRad);
      const pts = corners.map(c=>project(c.x,c.y,c.z));
      if(pointInPoly(mx,my,pts)){{ found={{si:si,row:row}}; break; }}
    }}
  }}
  if(found){{
    cv.style.cursor='pointer';
    tip.style.display='block';
    tip.style.left=(e.clientX-rect.left+14)+'px';
    tip.style.top=(e.clientY-rect.top-10)+'px';
    tip.innerHTML='<b>String '+(found.si+1)+'</b><br>Panneau #'+(found.row+1)+' / '+PANELS_PER_STRING+'<br>Puissance: '+PANEL_WP+' Wc';
    hoveredPanel = found;
  }} else {{
    cv.style.cursor='default'; tip.style.display='none'; hoveredPanel = null;
  }}
}});

cv.addEventListener('click', function(){{
  if(hoveredPanel){{
    selectedString = selectedString===hoveredPanel.si ? -1 : hoveredPanel.si;
    draw();
  }}
}});
cv.addEventListener('mouseleave', function(){{ tip.style.display='none'; }});

document.getElementById('sun3d').addEventListener('input', function(e){{
  sunAngle=+e.target.value;
  document.getElementById('sun-v3d').textContent=sunAngle+'°';
  draw();
}});
document.getElementById('rot3d').addEventListener('input', function(e){{
  rotY=+e.target.value;
  document.getElementById('rot-v3d').textContent=rotY+'°';
  draw();
}});
document.getElementById('elev3d').addEventListener('input', function(e){{
  elevAngle=+e.target.value;
  document.getElementById('elev-v3d').textContent=elevAngle+'°';
  draw();
}});

const ro = new ResizeObserver(function(){{ resize(); }});
ro.observe(wrap);
setTimeout(resize, 50);
}})();
</script>
"""
    components.html(pv3d_html, height=620, scrolling=False)

    st.markdown('<div class="sec">Disposition Électrique</div>', unsafe_allow_html=True)
    col_e1, col_e2 = st.columns(2)
    electrical_info = [
        ("Configuration",         f"{n_strings} strings × {panels_per_string} panneaux en série"),
        ("Puissance par string",  f"{panels_per_string * panel_wp / 1000:.2f} kWc"),
        ("Total MPPT requis",     f"{n_strings} (2 onduleurs × ~8 MPPT chacun, selon modèle)"),
        ("Panneau unitaire",      f"~{panel_wp} Wc, Si-mono"),
        ("Inclinaison",           "32° (orientation Sud)"),
        ("Surface au sol estimée",f"~{n_panels*5.6:,.0f} m² (5,6 m²/panneau avec espacement)"),
    ]
    for i, (k, v) in enumerate(electrical_info):
        with col_e1 if i % 2 == 0 else col_e2:
            st.markdown(f'<div class="info"><strong>{k} :</strong> {v}</div>', unsafe_allow_html=True)

    st.markdown('<div class="info" style="margin-top:8px;">💡 <strong>Astuce :</strong> cliquez sur un panneau dans la vue 3D pour mettre en surbrillance son string complet. Utilisez les curseurs pour faire pivoter la vue et simuler l\'angle du soleil à différentes heures de la journée.</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("""<div class="footer">
PROFITPV · MT · Monastir · 157,80 kWc · 2025 · CRA2E · eclareon / GIZ / ANME Tunisie<br>
Source : PROFITPV_MT_CRA2E_v0_TLS.xlsm — Résultats basés sur les hypothèses saisies.
</div>""", unsafe_allow_html=True)
