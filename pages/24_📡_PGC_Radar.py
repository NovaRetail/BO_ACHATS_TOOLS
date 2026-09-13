import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="PGC Radar", page_icon="📡", layout="wide")

# ============================================================
# Constantes
# ============================================================

FORMAT_MAP = {
    "10301": "Hyper", "10202": "Hyper", "10203": "Hyper",
    "10705": "Market", "10208": "Market", "10209": "Market",
    "10206": "Market", "10605": "Market", "10604": "Market",
    "10601": "Supeco", "10602": "Supeco", "10603": "Supeco",
}
SEUIL_CA_CRITIQUE = -0.05
SEUIL_CA_ATTENTION = -0.02
SEUIL_MARGE_CRITIQUE_PT = -1.0
SEUIL_MARGE_ATTENTION_PT = -0.5
MATERIALITE_FCFA = 1_000_000
SEUIL_POIDS_SITE = 0.10  # en dessous, les sites sont regroupés en "autres"

CAUSES = ["— À sélectionner —", "Rupture stock fournisseur", "Rupture stock magasin",
          "Prix non compétitif", "Exécution magasin", "Action concurrente",
          "Effet promo / cannibalisation", "Saisonnalité", "Autre"]

BLEU, ROUGE, VERT, ORANGE = "#007AFF", "#FF3B30", "#34C759", "#FF9500"
GRIS, GRIS_CLAIR, NOIR = "#8E8E93", "#F2F2F7", "#1C1C1E"

# ============================================================
# Charte visuelle — Apple style, police arrondie
# ============================================================

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&display=swap');

html, body, [class*="css"], .stApp {{
    font-family: 'Nunito', -apple-system, BlinkMacSystemFont, sans-serif;
}}
.stApp {{ background-color: {GRIS_CLAIR}; }}

.kpi-card {{
    background: #fff; border-radius: 16px; padding: 18px 20px;
    height: 100%;
}}
.kpi-label {{ font-size: 13px; color: {GRIS}; font-weight: 600; margin-bottom: 4px; }}
.kpi-value {{ font-size: 30px; color: {NOIR}; font-weight: 800; line-height: 1.1; }}
.kpi-delta-neg {{ font-size: 13px; color: {ROUGE}; font-weight: 700; margin-top: 4px; }}
.kpi-delta-pos {{ font-size: 13px; color: {VERT}; font-weight: 700; margin-top: 4px; }}
.kpi-delta-mut {{ font-size: 13px; color: {GRIS}; font-weight: 600; margin-top: 4px; }}

.ref-line {{
    background: #fff; border-radius: 14px; padding: 10px 16px;
    display: flex; justify-content: space-between; align-items: center;
    font-size: 13px; color: {GRIS}; font-weight: 600; margin: 10px 0 18px 0;
}}

.list-card {{ background: #fff; border-radius: 16px; overflow: hidden; margin-bottom: 18px; }}
.list-row {{
    display: flex; justify-content: space-between; align-items: center;
    padding: 12px 16px; border-bottom: 0.5px solid #E5E5EA;
    font-size: 14px; color: {NOIR}; font-weight: 600;
}}
.list-row:last-child {{ border-bottom: none; }}
.list-row.total {{ background: {GRIS_CLAIR}; font-weight: 800; }}

.fmt-block {{ background: #fff; border-radius: 16px; padding: 14px 16px; margin-bottom: 14px; }}
.fmt-head {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }}
.fmt-name {{ font-size: 15px; font-weight: 800; color: {NOIR}; }}
.fmt-sub {{ font-size: 12px; color: {GRIS}; font-weight: 600; }}
.site-row {{
    display: flex; justify-content: space-between; padding: 6px 0 6px 10px;
    font-size: 13px; color: {NOIR}; font-weight: 600;
    border-top: 0.5px solid #F0F0F3;
}}
.autres {{ font-size: 12px; color: {GRIS}; font-weight: 600; padding: 6px 0 0 10px; }}

.badge {{ font-size: 11px; font-weight: 800; padding: 3px 12px; border-radius: 12px; }}
.badge-crit {{ background: #FFE5E3; color: {ROUGE}; }}
.badge-att  {{ background: #FFF2DE; color: {ORANGE}; }}
.badge-ok   {{ background: #E4F8EA; color: {VERT}; }}

.alert-card {{ background: #fff; border-radius: 16px; padding: 14px 16px; margin-bottom: 12px; }}
.alert-head {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }}
.alert-title {{ font-size: 14px; font-weight: 800; color: {NOIR}; }}
.alert-line {{ font-size: 12px; color: {GRIS}; font-weight: 600; margin-top: 2px; }}

.neg {{ color: {ROUGE}; font-weight: 700; }}
.pos {{ color: {VERT}; font-weight: 700; }}
.att {{ color: {ORANGE}; font-weight: 700; }}
.mut {{ color: {GRIS}; font-weight: 600; }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# Chargement et calculs
# ============================================================

@st.cache_data(show_spinner=False)
def load_export(file):
    df = pd.read_excel(file, sheet_name="Export", header=0)
    df = df[~df["Département"].astype(str).str.startswith("Filtres appliqués", na=False)]
    df = df.dropna(how="all")
    df["Département"] = df["Département"].ffill()
    df["Rayon"] = df["Rayon"].ffill()

    grand_total = df[df["Département"] == "Total"]
    ca_reseau = grand_total["CA"].sum()

    df = df[df["Département"] != "Total"]
    dept = df[df["Rayon"] == "Total"].copy()
    site = df[(df["Site"] != "Total") & (df["Rayon"] != "Total") & df["Site"].notna()
              & (df["Département"] == "01 - PGC")].copy()
    site["code_site"] = site["Site"].str.split(" - ").str[0].str.strip()
    site["site_court"] = site["Site"].str.split(" - ").str[1].str.replace(
        r"^(Hyper|Market|Supeco)\s+", "", regex=True)
    site["format"] = site["code_site"].map(FORMAT_MAP).fillna("Autre")
    return dept, site, ca_reseau


def alerte(vs_n1, marge_pt):
    md = marge_pt if pd.notna(marge_pt) else 0
    if (pd.notna(vs_n1) and vs_n1 < SEUIL_CA_CRITIQUE) or md < SEUIL_MARGE_CRITIQUE_PT:
        return "Critique"
    if (pd.notna(vs_n1) and vs_n1 < SEUIL_CA_ATTENTION) or md < SEUIL_MARGE_ATTENTION_PT:
        return "Attention"
    return "OK"


BADGE = {"Critique": ("badge-crit", "Critique"), "Attention": ("badge-att", "Attention"),
         "OK": ("badge-ok", "OK")}


def fmt_m(x):
    return "—" if pd.isna(x) else f"{x/1_000_000:,.0f} M".replace(",", " ")

def fmt_pct(x, d=1):
    return "—" if pd.isna(x) else f"{x*100:+.{d}f}%".replace(".", ",")

def fmt_pt(x):
    return "—" if pd.isna(x) else f"{x:+.2f} pt".replace(".", ",")

def cls(x):
    if pd.isna(x):
        return "mut"
    return "neg" if x < 0 else "pos"


def agg_niveau(site_df, by):
    g = site_df.groupby(by).agg(
        nb_sites=("code_site", "nunique"),
        CA=("CA", "sum"), CA_N1=("CA N-1", "sum"),
        Budget=("Budget", lambda s: np.nan if s.isna().all() else s.sum()),
        Marge=("Marge", "sum"), Marge_N1=("Marge N-1", "sum"),
    ).reset_index()
    g["ecart_n1"] = g["CA"] - g["CA_N1"]
    g["Vs_N1"] = g["ecart_n1"] / g["CA_N1"]
    g["Vs_Bgt"] = np.where(g["Budget"].isna(), np.nan, (g["CA"] - g["Budget"]) / g["Budget"])
    g["Marge_pt"] = (g["Marge"] / g["CA"] - g["Marge_N1"] / g["CA_N1"]) * 100
    pool = g.loc[g["ecart_n1"] < 0, "ecart_n1"].sum()
    g["poids"] = g["ecart_n1"].apply(lambda x: x / pool if (x < 0 and pool < 0) else 0.0)
    return g


# ============================================================
# UI
# ============================================================

st.markdown(f"<div style='font-size:24px; font-weight:800; color:{NOIR}; padding:4px 0;'>📡 PGC Radar</div>",
            unsafe_allow_html=True)
st.markdown(f"<div style='font-size:13px; color:{GRIS}; font-weight:600; margin-bottom:14px;'>"
            "Performance réseau PGC — Vs N-1 et marge en premier plan, budget en référence</div>",
            unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📡 PGC Radar")
    st.caption("Import fichier")
    up = st.file_uploader("Export Département / Rayon / Site (avec Budget)", type=["xlsx"])

if not up:
    st.info("Dépose l'export PBI Département / Rayon / Site pour démarrer.")
    st.stop()

dept, site, ca_reseau = load_export(up)
pgc = dept[dept["Département"] == "01 - PGC"].iloc[0]

fmt_lvl = agg_niveau(site, "format").set_index("format")
fmt_lvl = fmt_lvl.reindex(["Hyper", "Market", "Supeco"]).reset_index()
site_lvl = agg_niveau(site, ["code_site", "site_court", "format"])

# Alertes Rayon x Site
al = site.copy()
al["Vs_N1"] = al["Vs N-1 (%)"]
al["Vs_Bgt"] = al["Vs Bgt (%)"]
al["Marge_pt"] = al["Taux de Marge N Vs N-1"]
al["Debit_vs_n1"] = al["Vs N-1 (%).1"]
al["Panier_vs_n1"] = al["Panier N Vs N-1"]
al["ecart_n1"] = al["CA"] - al["CA N-1"]
al["Alerte"] = al.apply(lambda r: alerte(r["Vs_N1"], r["Marge_pt"]), axis=1)
pool_al = al.loc[al["ecart_n1"] < 0, "ecart_n1"].sum()
al["poids"] = al["ecart_n1"].apply(lambda x: x / pool_al if (x < 0 and pool_al < 0) else 0.0)
al_actives = al[(al["Alerte"] != "OK") & (al["ecart_n1"].abs() >= MATERIALITE_FCFA)] \
    .sort_values("ecart_n1")

tab1, tab2, tab3 = st.tabs(["Vue d'ensemble", "Format", "Alertes"])

# ---------------- Écran 1 — Vue d'ensemble ----------------
with tab1:
    taux_marge = pgc["Taux de Marge"]
    c1, c2 = st.columns(2)
    with c1:
        d = pgc["Vs N-1 (%)"]
        st.markdown(f"""<div class='kpi-card'>
            <div class='kpi-label'>CA PGC</div>
            <div class='kpi-value'>{fmt_m(pgc["CA"])}</div>
            <div class='kpi-delta-{'neg' if d < 0 else 'pos'}'>{'▼' if d < 0 else '▲'} {fmt_pct(abs(d))[1:]} vs N-1</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        mpt = pgc["Taux de Marge N Vs N-1"]
        st.markdown(f"""<div class='kpi-card'>
            <div class='kpi-label'>Taux de marge</div>
            <div class='kpi-value'>{taux_marge*100:.1f}%</div>
            <div class='kpi-delta-{'neg' if mpt < 0 else 'pos'}'>{'▼' if mpt < 0 else '▲'} {fmt_pt(abs(mpt))[1:]} vs N-1</div>
        </div>""".replace(".", ","), unsafe_allow_html=True)

    st.markdown(f"""<div class='ref-line'>
        <span>Budget (référence)</span>
        <span>{fmt_m(pgc["Budget"])} · écart {fmt_pct(pgc["Vs Bgt (%)"])}</span>
    </div>""", unsafe_allow_html=True)

    rows_html = ""
    for _, r in dept.iterrows():
        if r["Département"] == "01 - PGC":
            continue
        d = r["Vs N-1 (%)"]
        rows_html += (f"<div class='list-row'><span>{r['Département'].title()} "
                      f"<span class='mut'>{fmt_m(r['CA'])}</span></span>"
                      f"<span class='{cls(d)}'>{fmt_pct(d)} N-1</span></div>")
    rows_html += (f"<div class='list-row total'><span>Total magasin "
                  f"<span class='mut'>{fmt_m(ca_reseau)}</span></span><span></span></div>")
    st.markdown(f"<div class='list-card'>{rows_html}</div>", unsafe_allow_html=True)

# ---------------- Écran 2 — Format ----------------
with tab2:
    for _, f in fmt_lvl.iterrows():
        b_cls, b_txt = BADGE[alerte(f["Vs_N1"], f["Marge_pt"])]
        sites_f = site_lvl[site_lvl["format"] == f["format"]].sort_values("ecart_n1")
        gros = sites_f[sites_f["poids"] >= SEUIL_POIDS_SITE]
        autres = sites_f[sites_f["poids"] < SEUIL_POIDS_SITE]

        sites_html = ""
        for _, s_ in gros.iterrows():
            sites_html += (f"<div class='site-row'><span>{s_['site_court']}</span>"
                           f"<span><span class='{cls(s_['Vs_N1'])}'>{fmt_pct(s_['Vs_N1'])}</span>"
                           f" · {fmt_m(s_['ecart_n1']).replace(' M', ' M')}"
                           f" · <b>{s_['poids']*100:.0f}% du recul réseau</b></span></div>")
        if len(autres) and autres["poids"].sum() > 0:
            sites_html += (f"<div class='autres'>+ {len(autres)} autres sites, "
                           f"poids cumulé {autres['poids'].sum()*100:.0f}% du recul</div>")
        elif not len(gros):
            sites_html += "<div class='autres'>Aucun site significatif dans le recul</div>"

        st.markdown(f"""<div class='fmt-block'>
            <div class='fmt-head'>
                <span class='fmt-name'>{f['format']}
                    <span class='fmt-sub'>{int(f['nb_sites'])} sites · {fmt_m(f['CA'])}</span></span>
                <span class='badge {b_cls}'>{b_txt}</span>
            </div>
            <div class='fmt-sub'>Vs N-1 <span class='{cls(f['Vs_N1'])}'>{fmt_pct(f['Vs_N1'])}</span>
                · Vs Budget <span class='mut'>{fmt_pct(f['Vs_Bgt'])}</span></div>
            {sites_html}
        </div>""", unsafe_allow_html=True)

# ---------------- Écran 3 — Alertes ----------------
with tab3:
    st.markdown(f"<div class='fmt-sub' style='margin-bottom:10px;'>"
                f"{len(al_actives)} lignes en alerte · déclencheur N-1 + marge · "
                f"écart ≥ {MATERIALITE_FCFA/1e6:.0f} M FCFA</div>", unsafe_allow_html=True)

    if al_actives.empty:
        st.success("Aucune alerte cette semaine.")
    else:
        for _, r in al_actives.iterrows():
            b_cls, b_txt = BADGE[r["Alerte"]]
            site_court = r["Site"].split(" - ")[1]
            rayon_court = r["Rayon"].split(" - ")[1].title()
            st.markdown(f"""<div class='alert-card'>
                <div class='alert-head'>
                    <span class='alert-title'>{site_court} — {rayon_court}</span>
                    <span class='badge {b_cls}'>{b_txt}</span>
                </div>
                <div class='alert-line'>CA {fmt_m(r['CA'])} ·
                    <span class='{cls(r['Vs_N1'])}'>{fmt_pct(r['Vs_N1'])} N-1</span> ·
                    <span class='mut'>Budget {fmt_pct(r['Vs_Bgt'])}</span> ·
                    Marge <span class='{cls(r['Marge_pt'])}'>{fmt_pt(r['Marge_pt'])}</span></div>
                <div class='alert-line'>Débit <span class='{cls(r['Debit_vs_n1'])}'>{fmt_pct(r['Debit_vs_n1'])} N-1</span>
                    · Panier <span class='{cls(r['Panier_vs_n1'])}'>{fmt_pct(r['Panier_vs_n1'])} N-1</span>
                    · <b>{r['poids']*100:.0f}% du recul réseau</b></div>
            </div>""", unsafe_allow_html=True)

        with st.expander("✏️ Saisir les causes (à transmettre aux acheteurs)"):
            saisie = al_actives[["Site", "Rayon", "Alerte"]].copy()
            saisie["Écart (M)"] = (al_actives["ecart_n1"] / 1e6).round(1)
            saisie["Cause"] = CAUSES[0]
            edited = st.data_editor(
                saisie, hide_index=True, width="stretch",
                column_config={"Cause": st.column_config.SelectboxColumn(options=CAUSES)},
                key="causes_editor",
            )
            csv = edited.to_csv(index=False).encode("utf-8-sig")
            st.download_button("Télécharger la liste (CSV)", data=csv,
                               file_name="alertes_pgc_radar.csv", mime="text/csv")
