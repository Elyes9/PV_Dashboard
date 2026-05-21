"""
PROFITPV Dashboard — Solar PV Installation Analyzer
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
.sec { font-size:11px; font-weight:700; letter-spacing:2px; text-transform:uppercase;
       color:#f5a623; margin:22px 0 10px; padding-bottom:6px; border-bottom:1px solid #1e2d3d; }
.info { background:#111820; border:1px solid #1e2d3d; border-radius:8px;
        padding:12px 16px; font-size:13px; color:#8aa4be; margin-bottom:8px; }
.info strong { color:#c5d8ea; }
.badge-ok   { background:#0d2b1a; border:1px solid #2ecc71; border-radius:10px; padding:14px; text-align:center; }
.badge-warn { background:#2b1f0d; border:1px solid #f5a623; border-radius:10px; padding:14px; text-align:center; }
.badge-bad  { background:#2b0d0d; border:1px solid #e74c3c; border-radius:10px; padding:14px; text-align:center; }
.footer { text-align:center; color:#2a4a6a; font-family:'DM Mono',monospace;
          font-size:11px; padding:10px 0; margin-top:20px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ALTAIR THEME  (dark, matches CSS palette)
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
# SIDEBAR INPUTS
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ☀️ PROFITPV")
    st.markdown("<div style='color:#5a7a9a;font-size:12px;margin-bottom:18px;font-family:DM Mono,monospace;'>CRA2E · MT · v0.2</div>", unsafe_allow_html=True)

    st.markdown('<div class="sec">🔧 Installation PV</div>', unsafe_allow_html=True)
    kwc        = st.number_input("Taille installation (kWc)", 1.0, 5000.0, 200.0, 10.0)
    prix_kwc   = st.number_input("Prix unitaire HT (DT/kWc)", 500.0, 5000.0, 2000.0, 50.0)
    subvention = st.selectbox("Subvention FTE ?", ["Non", "Oui"])
    sub_dt     = st.number_input("Montant subvention (DT)", 0.0, value=0.0, step=1000.0) if subvention == "Oui" else 0.0
    tva_pct    = st.slider("TVA (%)", 0, 20, 0) / 100

    st.markdown('<div class="sec">📍 Site & Panneaux</div>', unsafe_allow_html=True)
    region      = st.selectbox("Région", ["Monastir","Tunis","Sfax","Bizerte","Gafsa","Jendouba","Jerba","Kairouan","Tataouine","Tozeur"])
    orientation = st.selectbox("Orientation", ["Sud","Sud-ouest"])
    yield_map   = {"Monastir":1659,"Sfax":1680,"Gafsa":1720,"Jerba":1700,"Kairouan":1690,
                   "Tataouine":1730,"Tozeur":1740,"Bizerte":1580,"Jendouba":1560,"Tunis":1600}
    rendement   = st.number_input("Rendement spécifique (kWh/kWc/an)", 1000, 2000, yield_map.get(region, 1659))
    degradation = st.slider("Dégradation annuelle (%)", 0.0, 1.5, 0.4, 0.1) / 100

    st.markdown('<div class="sec">⚡ Consommation</div>', unsafe_allow_html=True)
    conso       = st.number_input("Consommation annuelle (kWh/an)", 1000.0, 5e6, 315628.0, 5000.0)
    tarif_type  = st.selectbox("Tarif STEG", ["Tarif Uniforme","Tarif par Postes Horaires"])
    tarif_kwh   = st.number_input("Tarif STEG (DT/kWh)", 0.1, 1.0, 0.291, 0.001, format="%.3f")
    surtaxe     = st.number_input("Surtaxe municipale (DT/kWh)", 0.0, 0.05, 0.006, 0.001, format="%.3f")
    limite_exced= st.slider("Limite excédent (%)", 10, 50, 30) / 100
    tarif_achat = st.number_input("Tarif achat excédent (DT/kWh)", 0.01, 0.30, 0.08, 0.001, format="%.3f")
    tarif_transport = st.number_input("Tarif transport (DT/kWh)", 0.01, 0.20, 0.039, 0.001, format="%.3f")

    st.markdown('<div class="sec">💰 Financement</div>', unsafe_allow_html=True)
    dette_pct    = st.slider("Part dette (%)", 0, 100, 70) / 100
    duree_credit = st.slider("Durée crédit (ans)", 1, 20, 10)
    grace        = st.slider("Délai de grâce (ans)", 0, 5, 1)
    taux_ref     = st.number_input("Taux référence BCT (%)", 0.0, 30.0, 10.0, 0.5) / 100
    marge        = st.number_input("Marge bancaire (%)", 0.0, 10.0, 1.25, 0.25) / 100
    taux_disc    = st.number_input("Taux actualisation (%)", 1.0, 20.0, 8.0, 0.5) / 100

    st.markdown('<div class="sec">📈 Hypothèses Prix</div>', unsafe_allow_html=True)
    hausse_tarif = st.slider("Hausse tarif STEG/an après Y5 (%)", 0.0, 10.0, 5.0, 0.5) / 100
    inflation    = st.number_input("Inflation (%/an)", 0.5, 15.0, 1.02, 0.1) / 100
    om_pct       = st.slider("O&M (% investissement/an)", 1, 10, 4) / 100
    duree        = st.slider("Durée du projet (ans)", 10, 30, 20)

# ─────────────────────────────────────────────
# CALCULATIONS
# ─────────────────────────────────────────────
cout_ht  = kwc * prix_kwc
cout_ttc = cout_ht * (1 + tva_pct) - sub_dt
cout_net = max(cout_ttc, 0)

prod_y1       = kwc * rendement
taux_couv     = min(prod_y1 / conso, 1.0)
autoconso_y1  = prod_y1 * min(taux_couv, 0.85)
cedee_y1      = max(0, prod_y1 - autoconso_y1)
cedee_lim_y1  = min(cedee_y1, prod_y1 * limite_exced)

facture_sans  = conso * (tarif_kwh + surtaxe)
eco_autoconso = autoconso_y1 * (tarif_kwh + surtaxe)
rev_exced     = cedee_lim_y1 * tarif_achat
facture_avec  = facture_sans - eco_autoconso - rev_exced
eco_annuelle  = facture_sans - facture_avec
om_y1         = cout_net * om_pct

montant_dette = min(cout_net * dette_pct, 200_000)
cap_propres   = cout_net - montant_dette
taux_interet  = taux_ref + marge

years = list(range(1, duree + 1))
prod_list, rev_list, opex_list, ds_list, ncf_list, cum_list = [], [], [], [], [], []
outstanding = montant_dette
cum = -cout_net

for y in years:
    p = prod_y1 * (1 - degradation) ** (y - 1)
    t = tarif_kwh if y <= 5 else tarif_kwh * (1 + hausse_tarif) ** (y - 5)
    ac = p * min(conso / (p + 1e-9), 0.85)
    ce = min(max(0, p - ac), p * limite_exced)
    r  = ac * (t + surtaxe) + ce * tarif_achat * (1 + hausse_tarif) ** max(0, y - 5)
    o  = cout_net * om_pct * (1 + inflation) ** (y - 1)

    if y <= grace:
        ds = outstanding * taux_interet
        principal = 0.0
    elif y <= duree_credit:
        principal = montant_dette / max(duree_credit - grace, 1)
        ds = outstanding * taux_interet + principal
        outstanding = max(0, outstanding - principal)
    else:
        ds = 0.0

    ncf = r - o
    cum += ncf

    prod_list.append(p)
    rev_list.append(r)
    opex_list.append(o)
    ds_list.append(ds)
    ncf_list.append(ncf)
    cum_list.append(cum)

# NPV / IRR via numpy
cf_arr = np.array([-cout_net] + ncf_list)

def calc_npv(rate, cfs):
    t = np.arange(len(cfs))
    return np.sum(cfs / (1 + rate) ** t)

def calc_irr(cfs, lo=-0.5, hi=5.0, tol=1e-7, maxiter=200):
    for _ in range(maxiter):
        mid = (lo + hi) / 2
        if calc_npv(mid, cfs) > 0:
            lo = mid
        else:
            hi = mid
        if hi - lo < tol:
            break
    return (lo + hi) / 2

project_npv = calc_npv(taux_disc, cf_arr)
project_irr = calc_irr(cf_arr)

# LCOE
disc_costs = cout_net + sum((cout_net * om_pct * (1 + inflation) ** (y - 1)) / (1 + taux_disc) ** y for y in range(1, duree + 1))
disc_prod  = sum((prod_y1 * (1 - degradation) ** (y - 1)) / (1 + taux_disc) ** y for y in range(1, duree + 1))
lcoe = disc_costs / disc_prod if disc_prod > 0 else 0

payback = next((i + 1 for i, v in enumerate(cum_list) if v > 0), None)
irr_pct = project_irr * 100

co2_per_year = [prod_y1 * (1 - degradation) ** (y - 1) * 0.57 / 1000 for y in years]
co2_total    = sum(co2_per_year)
co2_cum      = list(np.cumsum(co2_per_year))
total_prod   = sum(prod_list)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
if irr_pct > 10:
    badge_cls, badge_color, badge_label = "badge-ok", "#2ecc71", "RENTABLE"
elif irr_pct > 5:
    badge_cls, badge_color, badge_label = "badge-warn", "#f5a623", "MARGINAL"
else:
    badge_cls, badge_color, badge_label = "badge-bad", "#e74c3c", "DÉFICITAIRE"

col_title, col_badge = st.columns([4, 1])
with col_title:
    st.markdown(f"<h1 style='color:#f5a623;margin-bottom:4px;'>☀️ Analyse Installation PV</h1>", unsafe_allow_html=True)
    st.markdown(f"<div style='color:#5a7a9a;font-family:DM Mono,monospace;font-size:13px;'>{kwc:.0f} kWc · {region} · {orientation} · {duree} ans</div>", unsafe_allow_html=True)
with col_badge:
    st.markdown(f"""<div class="{badge_cls}" style="margin-top:10px;">
        <div style='color:{badge_color};font-size:18px;font-weight:800;'>{badge_label}</div>
        <div style='color:{badge_color};font-family:DM Mono,monospace;font-size:12px;'>TRI = {irr_pct:.1f}%</div>
    </div>""", unsafe_allow_html=True)

st.markdown("---")

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Résumé Financier",
    "⚡ Énergie & Production",
    "💵 Flux de Trésorerie",
    "🏗️ Investissement & Financement",
    "🌍 Impact Environnemental",
])

# ══════════════════════════════════
# TAB 1 — FINANCIAL SUMMARY
# ══════════════════════════════════
with tab1:
    st.markdown('<div class="sec">Indicateurs Clés de Rentabilité</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    kpis = [
        (c1, "TRI Projet", f"{irr_pct:.1f}%",       "Taux de rentabilité interne"),
        (c2, "VAN",        f"{project_npv/1000:+.0f}k DT", f"Taux actualisation {taux_disc*100:.0f}%"),
        (c3, "Retour Invest.", f"{payback or '>'+str(duree)} ans", "Temps de retour dynamique"),
        (c4, "LCOE",       f"{lcoe:.3f} DT/kWh",    f"vs tarif STEG {tarif_kwh:.3f} DT/kWh"),
    ]
    for col, label, val, unit in kpis:
        with col:
            st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    c5, c6, c7, c8 = st.columns(4)
    reduction = eco_annuelle / facture_sans * 100 if facture_sans > 0 else 0
    kpis2 = [
        (c5, "Économie Année 1",   f"{eco_annuelle:,.0f} DT",  f"Facture: {facture_avec:,.0f} → {facture_sans:,.0f} DT"),
        (c6, "Réduction Facture",  f"{reduction:.0f}%",         "de la facture STEG annuelle"),
        (c7, f"Gain Total {duree}a", f"{sum(ncf_list)/1000:.0f}k DT", "Flux nets cumulés"),
        (c8, "Investissement Net", f"{cout_net/1000:.0f}k DT",  "TTC après subvention"),
    ]
    for col, label, val, unit in kpis2:
        with col:
            st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="sec">Décomposition de la Facture STEG (Année 1)</div>', unsafe_allow_html=True)

    # Waterfall using stacked bars in Altair
    waterfall_data = pd.DataFrame([
        {"catégorie": "1. Facture sans PV",         "base": 0,             "val": facture_sans,     "type": "total"},
        {"catégorie": "2. Économies autoconso",      "base": facture_sans - eco_autoconso, "val": eco_autoconso,   "type": "économie"},
        {"catégorie": "3. Vente excédent",           "base": facture_sans - eco_autoconso - rev_exced, "val": rev_exced, "type": "vente"},
        {"catégorie": "4. Facture avec PV",          "base": 0,             "val": facture_avec,     "type": "résultat"},
    ])

    color_map = {"total": "#5a7a9a", "économie": "#2ecc71", "vente": "#3498db", "résultat": "#f5a623"}

    base_bars = alt.Chart(waterfall_data).mark_bar(opacity=0).encode(
        x=alt.X("catégorie:N", sort=None, axis=alt.Axis(labelAngle=-10)),
        y=alt.Y("base:Q", title="DT/an"),
    )
    color_bars = alt.Chart(waterfall_data).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X("catégorie:N", sort=None, title=""),
        y=alt.Y("base:Q"),
        y2=alt.Y2("val2:Q"),
        color=alt.Color("type:N", scale=alt.Scale(domain=list(color_map.keys()), range=list(color_map.values())), legend=alt.Legend(title="Type")),
        tooltip=["catégorie:N", alt.Tooltip("montant:Q", format=",.0f", title="Montant (DT)")],
    ).transform_calculate(
        val2="datum.base + datum.val",
        montant="datum.val",
    )
    wf_chart = (base_bars + color_bars).properties(
        title="Décomposition de la facture (DT/an)",
        height=300,
    )
    st.altair_chart(wf_chart, use_container_width=True)

    # LCOE comparison bar
    lcoe_df = pd.DataFrame([
        {"label": "LCOE (votre projet)", "value": lcoe,      "type": "lcoe"},
        {"label": f"Tarif STEG actuel",  "value": tarif_kwh, "type": "tarif"},
        {"label": "Tarif STEG +5%×15a",  "value": tarif_kwh * (1.05 ** 15), "type": "projection"},
    ])
    lcoe_chart = alt.Chart(lcoe_df).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X("value:Q", title="DT/kWh"),
        y=alt.Y("label:N", sort=None, title=""),
        color=alt.Color("type:N", scale=alt.Scale(
            domain=["lcoe","tarif","projection"],
            range=["#f5a623","#e74c3c","#8e44ad"]
        ), legend=None),
        tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=".3f", title="DT/kWh")],
    ).properties(title="LCOE vs Tarifs STEG (DT/kWh)", height=160)
    st.altair_chart(lcoe_chart, use_container_width=True)


# ══════════════════════════════════
# TAB 2 — ENERGY & PRODUCTION
# ══════════════════════════════════
with tab2:
    st.markdown('<div class="sec">Indicateurs de Production</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    en_kpis = [
        (c1, "Production Année 1", f"{prod_y1:,.0f} kWh",  f"{kwc:.0f} kWc × {rendement} kWh/kWc"),
        (c2, "Taux de Couverture",  f"{taux_couv*100:.0f}%", "consommation couverte"),
        (c3, "Autoconsommation",    f"{autoconso_y1:,.0f} kWh", "utilisée sur site"),
        (c4, "Excédent STEG",       f"{cedee_lim_y1:,.0f} kWh", f"limité à {limite_exced*100:.0f}%"),
    ]
    for col, label, val, unit in en_kpis:
        with col:
            st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    # Production over project life
    df_prod = pd.DataFrame({"Année": years, "Production (kWh)": prod_list, "Consommation (kWh)": [conso] * duree})
    df_prod_melt = df_prod.melt("Année", var_name="Série", value_name="kWh")

    prod_bars = alt.Chart(df_prod).mark_bar(
        color="#f5a623", opacity=0.85, cornerRadiusTopLeft=2, cornerRadiusTopRight=2
    ).encode(
        x=alt.X("Année:O", title="Année"),
        y=alt.Y("Production (kWh):Q", title="kWh/an"),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Production (kWh):Q", format=",.0f")],
    )
    conso_line = alt.Chart(df_prod).mark_rule(
        color="#e74c3c", strokeDash=[6, 3], strokeWidth=2
    ).encode(y="mean(Consommation (kWh)):Q")
    conso_text = alt.Chart(pd.DataFrame({"y": [conso], "label": [f"Consommation: {conso:,.0f} kWh/an"]})).mark_text(
        align="left", dx=5, dy=-8, color="#e74c3c", fontSize=11, font="DM Mono"
    ).encode(y="y:Q", text="label:N", x=alt.value(10))

    prod_chart = (prod_bars + conso_line + conso_text).properties(
        title=f"Production PV annuelle — dégradation {degradation*100:.1f}%/an",
        height=320,
    )
    st.altair_chart(prod_chart, use_container_width=True)

    col_left, col_right = st.columns(2)

    with col_left:
        # Monthly profile
        months = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
        seasonal = [0.070,0.078,0.095,0.108,0.113,0.108,0.100,0.095,0.088,0.080,0.070,0.065]
        df_monthly = pd.DataFrame({
            "Mois": months,
            "Production": [prod_y1 * f for f in seasonal],
            "Consommation": [conso / 12] * 12,
            "ordre": list(range(12)),
        })
        monthly_bars = alt.Chart(df_monthly).mark_bar(color="#f5a623", opacity=0.8).encode(
            x=alt.X("Mois:N", sort=alt.SortField("ordre"), title=""),
            y=alt.Y("Production:Q", title="kWh/mois"),
            tooltip=[alt.Tooltip("Mois:N"), alt.Tooltip("Production:Q", format=",.0f", title="Production kWh")],
        )
        monthly_line = alt.Chart(df_monthly).mark_line(color="#e74c3c", strokeDash=[4,2], strokeWidth=2).encode(
            x=alt.X("Mois:N", sort=alt.SortField("ordre")),
            y=alt.Y("Consommation:Q"),
            tooltip=[alt.Tooltip("Mois:N"), alt.Tooltip("Consommation:Q", format=",.0f", title="Conso kWh")],
        )
        monthly_chart = (monthly_bars + monthly_line).properties(title="Profil mensuel de production vs consommation", height=280)
        st.altair_chart(monthly_chart, use_container_width=True)

    with col_right:
        # Energy balance donut via stacked arc (Altair arc mark)
        autoconso_pct = autoconso_y1 / prod_y1 * 100
        cedee_pct     = cedee_lim_y1 / prod_y1 * 100
        perdu_pct     = max(0, 100 - autoconso_pct - cedee_pct)

        pie_df = pd.DataFrame([
            {"label": "Autoconsommée", "value": round(autoconso_pct, 1)},
            {"label": "Vendue STEG",   "value": round(cedee_pct, 1)},
            {"label": "Non valorisée", "value": round(perdu_pct, 1)},
        ])
        pie_chart = alt.Chart(pie_df).mark_arc(innerRadius=70, outerRadius=110).encode(
            theta=alt.Theta("value:Q"),
            color=alt.Color("label:N", scale=alt.Scale(
                domain=["Autoconsommée","Vendue STEG","Non valorisée"],
                range=["#f5a623","#2ecc71","#e74c3c"]
            ), legend=alt.Legend(title="Répartition")),
            tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=".1f", title="%")],
        ).properties(title="Répartition de la production PV (%)", height=280)
        st.altair_chart(pie_chart, use_container_width=True)

        st.markdown(f'<div class="info"><strong>Production totale {duree} ans :</strong> {total_prod/1000:,.0f} MWh</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="info"><strong>Rendement :</strong> {rendement} kWh/kWc/an ({region})</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="info"><strong>Dégradation :</strong> {degradation*100:.2f}%/an</div>', unsafe_allow_html=True)


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
    cum_line = alt.Chart(df_cf).mark_line(color="#2ecc71", strokeWidth=2.5).encode(
        x=alt.X("Année:O"),
        y=alt.Y("Cumul:Q", axis=alt.Axis(title="Cumul DT", orient="right")),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Cumul:Q", format=",.0f", title="Cumul DT")],
    )
    zero_rule = alt.Chart(pd.DataFrame({"y": [0]})).mark_rule(
        color="#5a7a9a", strokeDash=[4, 2], strokeWidth=1
    ).encode(y="y:Q")

    cf_chart = alt.layer(bars, cum_line + zero_rule).resolve_scale(y="independent").properties(
        title=f"Revenus, coûts et cumul des flux nets ({duree} ans)", height=380
    )
    st.altair_chart(cf_chart, use_container_width=True)

    # Sensitivity: IRR vs cost
    st.markdown('<div class="sec">Analyse de Sensibilité</div>', unsafe_allow_html=True)
    col_s1, col_s2 = st.columns(2)

    with col_s1:
        costs = np.linspace(1000, 4000, 25)
        irrs_c = []
        for c in costs:
            inv = kwc * c * (1 + tva_pct) - sub_dt
            cfs = np.array([-inv] + [r - inv * om_pct for r in rev_list])
            try:
                irrs_c.append(calc_irr(cfs) * 100)
            except Exception:
                irrs_c.append(np.nan)
        df_sens1 = pd.DataFrame({"Coût DT/kWc": costs, "TRI %": irrs_c})
        threshold1 = pd.DataFrame({"y": [taux_disc * 100]})

        line1 = alt.Chart(df_sens1).mark_line(color="#f5a623", strokeWidth=2).encode(
            x=alt.X("Coût DT/kWc:Q", title="Coût installation (DT/kWc)"),
            y=alt.Y("TRI %:Q", title="TRI (%)"),
            tooltip=[alt.Tooltip("Coût DT/kWc:Q", format=",.0f"), alt.Tooltip("TRI %:Q", format=".1f")],
        )
        area1 = alt.Chart(df_sens1).mark_area(color="#f5a623", opacity=0.1).encode(
            x="Coût DT/kWc:Q", y="TRI %:Q", y2=alt.value(0),
        )
        rule1 = alt.Chart(threshold1).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
        current1 = alt.Chart(pd.DataFrame({"x": [prix_kwc]})).mark_rule(color="#2ecc71", strokeDash=[4,2]).encode(x="x:Q")
        sens1_chart = (area1 + line1 + rule1 + current1).properties(title="TRI vs Coût d'installation", height=280)
        st.altair_chart(sens1_chart, use_container_width=True)

    with col_s2:
        tariffs = np.linspace(0, 0.10, 25)
        irrs_t = []
        for ht in tariffs:
            cfs2 = [-cout_net]
            for y in years:
                p2 = prod_y1 * (1 - degradation) ** (y - 1)
                t2 = tarif_kwh if y <= 5 else tarif_kwh * (1 + ht) ** (y - 5)
                ac2 = p2 * min(conso / (p2 + 1e-9), 0.85)
                ce2 = min(max(0, p2 - ac2), p2 * limite_exced)
                r2  = ac2 * (t2 + surtaxe) + ce2 * tarif_achat
                o2  = cout_net * om_pct * (1 + inflation) ** (y - 1)
                cfs2.append(r2 - o2)
            try:
                irrs_t.append(calc_irr(np.array(cfs2)) * 100)
            except Exception:
                irrs_t.append(np.nan)
        df_sens2 = pd.DataFrame({"Hausse %/an": tariffs * 100, "TRI %": irrs_t})
        threshold2 = pd.DataFrame({"y": [taux_disc * 100]})

        line2 = alt.Chart(df_sens2).mark_line(color="#2ecc71", strokeWidth=2).encode(
            x=alt.X("Hausse %/an:Q", title="Hausse tarif STEG (%/an)"),
            y=alt.Y("TRI %:Q", title="TRI (%)"),
            tooltip=[alt.Tooltip("Hausse %/an:Q", format=".1f"), alt.Tooltip("TRI %:Q", format=".1f")],
        )
        area2 = alt.Chart(df_sens2).mark_area(color="#2ecc71", opacity=0.1).encode(
            x="Hausse %/an:Q", y="TRI %:Q", y2=alt.value(0),
        )
        rule2 = alt.Chart(threshold2).mark_rule(color="#e74c3c", strokeDash=[4,2]).encode(y="y:Q")
        current2 = alt.Chart(pd.DataFrame({"x": [hausse_tarif * 100]})).mark_rule(color="#f5a623", strokeDash=[4,2]).encode(x="x:Q")
        sens2_chart = (area2 + line2 + rule2 + current2).properties(title="TRI vs Hausse du tarif STEG", height=280)
        st.altair_chart(sens2_chart, use_container_width=True)

    # Detailed table
    with st.expander("📋 Tableau détaillé des flux de trésorerie"):
        df_table = pd.DataFrame({
            "Année":              years,
            "Production (kWh)":  [round(p) for p in prod_list],
            "Revenus (DT)":      [round(r) for r in rev_list],
            "OpEx (DT)":         [round(o) for o in opex_list],
            "Service dette (DT)":[round(d) for d in ds_list],
            "Flux net (DT)":     [round(n) for n in ncf_list],
            "Cumul (DT)":        [round(c) for c in cum_list],
        })
        st.dataframe(df_table.style.format({
            "Production (kWh)": "{:,.0f}", "Revenus (DT)": "{:,.0f}",
            "OpEx (DT)": "{:,.0f}", "Service dette (DT)": "{:,.0f}",
            "Flux net (DT)": "{:,.0f}", "Cumul (DT)": "{:,.0f}",
        }).background_gradient(subset=["Flux net (DT)","Cumul (DT)"], cmap="RdYlGn"),
        use_container_width=True)


# ══════════════════════════════════
# TAB 4 — INVESTMENT & FINANCING
# ══════════════════════════════════
with tab4:
    st.markdown('<div class="sec">Structure de l\'Investissement</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    inv_kpis = [
        (c1, "Coût total HT",      f"{cout_ht:,.0f} DT",  f"{kwc:.0f} kWc × {prix_kwc:.0f} DT/kWc"),
        (c2, f"TVA ({tva_pct*100:.0f}%)", f"{cout_ht*tva_pct:,.0f} DT", "Taxe sur la valeur ajoutée"),
        (c3, "Coût Net TTC",       f"{cout_net:,.0f} DT",  f"Subvention déduite: {sub_dt:,.0f} DT"),
    ]
    for col, label, val, unit in inv_kpis:
        with col:
            st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    col_f1, col_f2 = st.columns(2)

    with col_f1:
        fin_pie_df = pd.DataFrame([
            {"label": "Capitaux propres", "value": cap_propres},
            {"label": "Dette bancaire",   "value": montant_dette},
        ])
        fin_pie = alt.Chart(fin_pie_df).mark_arc(innerRadius=65, outerRadius=105).encode(
            theta=alt.Theta("value:Q"),
            color=alt.Color("label:N", scale=alt.Scale(
                domain=["Capitaux propres","Dette bancaire"],
                range=["#f5a623","#3498db"]
            ), legend=alt.Legend(title="Financement")),
            tooltip=[alt.Tooltip("label:N"), alt.Tooltip("value:Q", format=",.0f", title="DT")],
        ).properties(title="Structure de financement", height=280)
        st.altair_chart(fin_pie, use_container_width=True)

        params = [
            ("Capitaux propres",      f"{cap_propres:,.0f} DT"),
            ("Montant de la dette",   f"{montant_dette:,.0f} DT ({dette_pct*100:.0f}% — plafond 200 000 DT)"),
            ("Taux d'intérêt effectif", f"{(taux_ref+marge)*100:.2f}% (BCT {taux_ref*100:.2f}% + marge {marge*100:.2f}%)"),
            ("Durée du crédit",        f"{duree_credit} ans"),
            ("Délai de grâce",         f"{grace} an(s)"),
            ("Commission initiale",    "1,0 % du crédit"),
            ("Commission d'engagement","0,25 % de la marge"),
        ]
        for k, v in params:
            st.markdown(f'<div class="info"><strong>{k} :</strong> {v}</div>', unsafe_allow_html=True)

    with col_f2:
        # Debt repayment schedule
        out_bal = montant_dette
        dp_rows = []
        for y in years:
            if y <= grace:
                interest = out_bal * taux_interet
                principal_p = 0.0
            elif y <= duree_credit:
                principal_p = montant_dette / max(duree_credit - grace, 1)
                interest = out_bal * taux_interet
                out_bal = max(0, out_bal - principal_p)
            else:
                principal_p = 0.0
                interest = 0.0
            dp_rows.append({"Année": y, "Principal": principal_p, "Intérêts": interest, "Encours": out_bal})
        df_debt = pd.DataFrame(dp_rows)
        df_debt_melt = df_debt.melt("Année", value_vars=["Principal","Intérêts"], var_name="Poste", value_name="DT")

        debt_bars = alt.Chart(df_debt_melt).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
            x=alt.X("Année:O", title="Année"),
            y=alt.Y("DT:Q", title="DT/an", stack="zero"),
            color=alt.Color("Poste:N", scale=alt.Scale(
                domain=["Principal","Intérêts"],
                range=["#3498db","#8e44ad"]
            )),
            tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Poste:N"), alt.Tooltip("DT:Q", format=",.0f")],
        )
        encours_line = alt.Chart(df_debt).mark_line(color="#e74c3c", strokeWidth=2, strokeDash=[4,2]).encode(
            x=alt.X("Année:O"),
            y=alt.Y("Encours:Q", axis=alt.Axis(title="Encours DT", orient="right")),
            tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("Encours:Q", format=",.0f", title="Encours")],
        )
        debt_chart = alt.layer(debt_bars, encours_line).resolve_scale(y="independent").properties(
            title="Remboursement de la dette bancaire", height=340
        )
        st.altair_chart(debt_chart, use_container_width=True)


# ══════════════════════════════════
# TAB 5 — ENVIRONMENTAL IMPACT
# ══════════════════════════════════
with tab5:
    st.markdown('<div class="sec">Impact Environnemental</div>', unsafe_allow_html=True)

    trees = co2_total * 40
    cars  = co2_total / 2.3

    c1, c2, c3, c4 = st.columns(4)
    env_kpis = [
        (c1, f"CO₂ Évité ({duree} ans)", f"{co2_total:,.0f} t",   "tonnes équivalent CO₂"),
        (c2, "Énergie Propre",           f"{total_prod/1000:,.0f} MWh", f"produite sur {duree} ans"),
        (c3, "Équivalent Arbres",        f"{trees:,.0f}",           "arbres/an équivalent"),
        (c4, "Voitures retirées",        f"{cars:,.0f}",            "voitures/an équivalent"),
    ]
    for col, label, val, unit in env_kpis:
        with col:
            st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-unit">{unit}</div></div>', unsafe_allow_html=True)

    df_co2 = pd.DataFrame({"Année": years, "CO₂ évité (t)": co2_per_year, "CO₂ cumulé (t)": co2_cum})

    co2_bars = alt.Chart(df_co2).mark_bar(
        color="#2ecc71", opacity=0.75, cornerRadiusTopLeft=2, cornerRadiusTopRight=2
    ).encode(
        x=alt.X("Année:O", title="Année"),
        y=alt.Y("CO₂ évité (t):Q", title="t CO₂/an"),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("CO₂ évité (t):Q", format=".1f", title="t CO₂ évité")],
    )
    co2_cum_line = alt.Chart(df_co2).mark_line(color="#f5a623", strokeWidth=2.5).encode(
        x=alt.X("Année:O"),
        y=alt.Y("CO₂ cumulé (t):Q", axis=alt.Axis(title="t CO₂ cumulé", orient="right")),
        tooltip=[alt.Tooltip("Année:O"), alt.Tooltip("CO₂ cumulé (t):Q", format=".0f", title="Cumul t CO₂")],
    )
    co2_chart = alt.layer(co2_bars, co2_cum_line).resolve_scale(y="independent").properties(
        title="CO₂ évité par an et cumulé (0,57 t CO₂/MWh solaire — Tunisie)",
        height=360,
    )
    st.altair_chart(co2_chart, use_container_width=True)

    st.markdown('<div class="sec">Mise en Contexte</div>', unsafe_allow_html=True)
    col_c1, col_c2 = st.columns(2)
    context = [
        ("Facteur d'émission Tunisie", "0,57 t CO₂ évitées par MWh solaire produit (réseau national)"),
        ("CO₂ évité — Année 1",        f"{co2_per_year[0]:.1f} tonnes"),
        ("CO₂ évité — Année 20",       f"{co2_per_year[-1]:.1f} tonnes (dégradation incluse)"),
        ("Équivalent arbres",           f"1 arbre absorbe ~25 kg CO₂/an → équiv. {int(trees):,} arbres"),
        ("Équivalent voitures",         f"Voiture moyenne: 2,3 t CO₂/an → {int(cars):,} voitures retirées"),
        ("Production totale",           f"{total_prod/1000:,.0f} MWh sur {duree} ans"),
    ]
    for i, (k, v) in enumerate(context):
        with col_c1 if i % 2 == 0 else col_c2:
            st.markdown(f'<div class="info"><strong>{k} :</strong> {v}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("""<div class="footer">
PROFITPV · Analyse de rentabilité PV MT/HT · Tunisie · CRA2E · eclareon / GIZ / ANME<br>
Les résultats sont basés sur les hypothèses saisies — consultez un expert pour valider l'étude de faisabilité.
</div>""", unsafe_allow_html=True)
