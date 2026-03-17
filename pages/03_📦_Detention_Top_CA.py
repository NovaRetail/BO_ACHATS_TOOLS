# ─── TRAITEMENT ──────────────────────────────────────────────────────────────
with st.spinner("Calcul des taux de détention…"):
    grp, absents = compute_detention(df_stock, top_codes)
    grp["Alerte"] = grp.apply(compute_alerte, axis=1)
    taux_df = compute_taux(grp, top_codes)

# ─── KPI BASE ────────────────────────────────────────────────────────────────
n_sites   = df_stock["Libellé site"].nunique()
taux_all  = taux_df[taux_df["flux"]=="ALL"]
taux_im   = taux_df[taux_df["flux"]=="IM"]
taux_lo   = taux_df[taux_df["flux"]=="LO"]

taux_moy  = taux_all["taux"].mean()
taux_im_m = taux_im["taux"].mean()
taux_lo_m = taux_lo["taux"].mean()

n_urgences= (grp["Alerte"] != "✅ OK").sum()

# ─── NOUVEAU : VALORISATION STOCK ───────────────────────────────────────────
grp["val_stock_achat"] = grp["stock"] * grp.get("Prix d'achat", 0)

val_total = grp["val_stock_achat"].sum()

df_b = grp[grp["code_etat"] == "B"].copy()
val_bloques = df_b["val_stock_achat"].sum()
nb_bloques  = len(df_b)

pct_bloques = (val_bloques / val_total * 100) if val_total > 0 else 0

top_set = set(top_codes)
df_b_top = df_b[df_b["Code article"].isin(top_set)]

df_actifs = grp[grp["code_etat"] == "2"].copy()

# ─── KPIs ────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-label'>Indicateurs globaux · " + str(n_sites) + " magasin(s) · " + str(len(top_codes)) + " références Top CA</div>", unsafe_allow_html=True)

k1,k2,k3,k4,k5,k6 = st.columns(6)

k1.metric("Réf Top CA", str(len(top_codes)))

k2.metric("Taux détention moy",
          f"{taux_moy:.1f}%" if taux_moy else "—",
          f"cible {cible_taux}%")

k3.metric("Taux IM (Import)",
          f"{taux_im_m:.1f}%" if taux_im_m else "—",
          f"{taux_im_m-cible_taux:+.1f} pt vs cible" if taux_im_m else "")

k4.metric("Taux LO (Local)",
          f"{taux_lo_m:.1f}%" if taux_lo_m else "—",
          f"{taux_lo_m-cible_taux:+.1f} pt vs cible" if taux_lo_m else "")

k5.metric("Urgences", str(n_urgences), "articles à traiter")

# KPI STOCK BLOQUÉ
if pct_bloques > 15:
    delta_b = "🔴 Critique"
elif pct_bloques > 5:
    delta_b = "🟡 À surveiller"
else:
    delta_b = "🟢 OK"

k6.metric(
    "€ Stock bloqué (B)",
    fmt(val_bloques),
    f"{pct_bloques:.1f}% du stock · {nb_bloques} réf"
)

# ─── ALERTES ─────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Alertes & actions prioritaires</div>", unsafe_allow_html=True)

sites_sous_cible = taux_all[taux_all["taux"] < cible_taux].sort_values("taux")
if not sites_sous_cible.empty:
    liste = ", ".join([f"{r['site']} ({r['taux']:.0f}%)" for _,r in sites_sous_cible.iterrows()])
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>⚠️ {len(sites_sous_cible)} magasin(s) sous la cible {cible_taux}%</strong> — {liste}
</div>""", unsafe_allow_html=True)

# ─── TABS ─────────────────────────────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Synthèse réseau", "🔄 IM vs LO", "🚨 Plan d'action", "🚫 Absents ERP"
])

# ═══ TAB 3 ═══════════════════════════════════════════════════════════════════
with tab3:

    # 🔴 STOCK BLOQUÉ
    st.markdown("## 🔴 Analyse du stock bloqué")

    st.markdown("### Top 10 Bloqués — Top CA")
    top_b_topca = (
        df_b_top.groupby(["Code article","lib_article"])
        .agg({"val_stock_achat": "sum","stock": "sum","Libellé site": "nunique"})
        .reset_index()
        .sort_values("val_stock_achat", ascending=False)
        .head(10)
    )
    top_b_topca.columns = ["Code","Libellé","€ Stock","Stock","Nb magasins"]
    top_b_topca["€ Stock"] = top_b_topca["€ Stock"].apply(fmt)
    st.dataframe(top_b_topca, use_container_width=True, hide_index=True)

    st.markdown("### Top 10 Bloqués — Global")
    top_b_global = (
        df_b.groupby(["Code article","lib_article"])
        .agg({"val_stock_achat": "sum","stock": "sum","Libellé site": "nunique"})
        .reset_index()
        .sort_values("val_stock_achat", ascending=False)
        .head(10)
    )
    top_b_global.columns = ["Code","Libellé","€ Stock","Stock","Nb magasins"]
    top_b_global["€ Stock"] = top_b_global["€ Stock"].apply(fmt)
    st.dataframe(top_b_global, use_container_width=True, hide_index=True)

    # 🟢 STOCK ACTIF
    st.markdown("## 🟢 Analyse stock actif (Code 2)")

    st.markdown("### Top 10 Actifs — Top CA")
    top_actifs_topca = (
        df_actifs[df_actifs["Code article"].isin(top_set)]
        .groupby(["Code article","lib_article"])
        .agg({"val_stock_achat": "sum","stock": "sum","Libellé site": "nunique"})
        .reset_index()
        .sort_values("val_stock_achat", ascending=False)
        .head(10)
    )
    top_actifs_topca.columns = ["Code","Libellé","€ Stock","Stock","Nb magasins"]
    top_actifs_topca["€ Stock"] = top_actifs_topca["€ Stock"].apply(fmt)
    st.dataframe(top_actifs_topca, use_container_width=True, hide_index=True)

    st.markdown("### Top 10 Actifs — Global")
    top_actifs_global = (
        df_actifs.groupby(["Code article","lib_article"])
        .agg({"val_stock_achat": "sum","stock": "sum","Libellé site": "nunique"})
        .reset_index()
        .sort_values("val_stock_achat", ascending=False)
        .head(10)
    )
    top_actifs_global.columns = ["Code","Libellé","€ Stock","Stock","Nb magasins"]
    top_actifs_global["€ Stock"] = top_actifs_global["€ Stock"].apply(fmt)
    st.dataframe(top_actifs_global, use_container_width=True, hide_index=True)

    # 🚨 PLAN D'ACTION
    st.markdown("## 🚨 Plan d’action")

    urgences = grp[grp["Alerte"] != "✅ OK"].copy()

    if urgences.empty:
        st.success("✅ Aucune urgence.")
    else:
        disp3 = urgences[["Code article","lib_article","Libellé site","Code marketing",
                           "code_etat","stock","ral","Alerte"]]
        disp3.columns = ["Code","Libellé","Magasin","Flux","Code état","Stock","RAL","Alerte"]
        st.dataframe(disp3, use_container_width=True, hide_index=True)
