"""
Module 18 - Reporting Vente CA
================================
(Remplace la précédente version du module 18_📊_Reporting_Vente CA.py)

Vue BI exploratoire (lecture seule) du reporting hebdomadaire de
performance commerciale. Cette page NE remplace PAS le fichier Excel
généré par `reporting_achats.main` : elle sert à explorer les mêmes
données de façon interactive avant / après l'export. La saisie des
commentaires (Cause / Commentaire acheteur) se fait exclusivement dans
le fichier Excel, pas ici.

Charte graphique : cohérente avec le reste de SmartBuyer Hub
(#F2F2F7, #007AFF, #34C759, #FF3B30, SF Pro/Inter, radius 14px).
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from reporting_achats import diagnostics, drivers, excel_writer, prioritization
from reporting_achats.config import ReportingConfig
from reporting_achats.loaders import load_all
from reporting_achats.metrics import (
    RESEAU,
    build_famille_kpis,
    build_rayon_format_kpis,
    build_site_famille_kpis,
)
from reporting_achats.transforms import detail_rows, enrich_detail_export, enrich_site_export
from reporting_achats.validation import build_quality_report

st.set_page_config(page_title="Reporting Vente CA", page_icon="💸", layout="wide")

# -----------------------------------------------------------------------------
# Charte graphique Apple (cohérente avec les autres modules du hub)
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .kpi-card {
        background: #FFFFFF;
        border-radius: 14px;
        padding: 16px 20px;
        border: 1px solid #E5E5EA;
        margin-bottom: 8px;
    }
    .kpi-label { font-size: 12px; color: #6E6E73; margin-bottom: 4px; }
    .kpi-value { font-size: 22px; font-weight: 600; color: #1D1D1F; }
    .kpi-evo-pos { font-size: 12px; color: #34C759; margin-top: 4px; }
    .kpi-evo-neg { font-size: 12px; color: #FF3B30; margin-top: 4px; }
    .kpi-evo-neutral { font-size: 12px; color: #6E6E73; margin-top: 4px; }
    .diag-box {
        background: #F2F2F7;
        border-radius: 14px;
        padding: 16px 20px;
        font-size: 14px;
        color: #1D1D1F;
        margin: 12px 0 20px 0;
    }
    .alert-box {
        background: #FFF3E0;
        border-left: 4px solid #FF9500;
        border-radius: 8px;
        padding: 10px 16px;
        font-size: 13px;
        color: #6E6E73;
        margin: 12px 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def _fmt_millions(v):
    if v is None or pd.isna(v):
        return "N/D"
    return f"{v / 1_000_000:,.1f} M".replace(",", " ")


def _fmt_pct(v, signed=True):
    if v is None or pd.isna(v):
        return "N/D"
    return f"{v * 100:+.1f} %" if signed else f"{v * 100:.1f} %"


def _fmt_pts(v):
    if v is None or pd.isna(v):
        return "N/D"
    return f"{v * 100:+.2f} pt"


def _evo_css_class(v):
    if v is None or pd.isna(v):
        return "kpi-evo-neutral"
    return "kpi-evo-pos" if v > 0 else ("kpi-evo-neg" if v < 0 else "kpi-evo-neutral")


def kpi_card(col, label, value_str, evo_val, evo_str):
    css = _evo_css_class(evo_val)
    col.markdown(
        f"""<div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value_str}</div>
            <div class="{css}">{evo_str}</div>
        </div>""",
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_and_process(site_bytes, current_bytes, previous_bytes, config_bytes):
    import tempfile
    from pathlib import Path

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        site_path = tmp_path / "site.xlsx"
        current_path = tmp_path / "current.xlsx"
        previous_path = tmp_path / "previous.xlsx"
        site_path.write_bytes(site_bytes)
        current_path.write_bytes(current_bytes)
        previous_path.write_bytes(previous_bytes)

        cfg_path = None
        if config_bytes:
            cfg_path = tmp_path / "config.yaml"
            cfg_path.write_bytes(config_bytes)

        cfg = ReportingConfig.load(cfg_path)
        exports = load_all(site_path, current_path, previous_path)

        site_enriched = enrich_site_export(exports.site.detail_rows, cfg.format_mapping)
        n_enriched = enrich_detail_export(exports.current.detail_rows, cfg.format_mapping)
        n1_enriched = enrich_detail_export(exports.previous_iso.detail_rows, cfg.format_mapping)

        site_d = detail_rows(site_enriched)
        n_d = detail_rows(n_enriched)
        n1_d = detail_rows(n1_enriched)

        quality = build_quality_report(exports, site_enriched, n_enriched, n1_enriched, cfg)

        rayons = (
            n_d[["rayon_code", "rayon_label"]].drop_duplicates()
            .set_index("rayon_code")["rayon_label"].to_dict()
        )
        display_label = {c: cfg.rayon_label_overrides.get(l, l) for c, l in rayons.items()}

        all_kpis = build_rayon_format_kpis(
            n_d, n1_d, site_d,
            site_coverage_by_scope={
                RESEAU: quality.site_coverage_global.coverage_pct,
                **{fmt: cov.coverage_pct for fmt, cov in quality.site_coverage_by_format.items()},
            },
        )
        kpis_lookup = {(k.rayon_code, k.scope): k for k in all_kpis}

        familles_by_rayon = {}
        site_famille_by_rayon = {}
        diagnostics_by_rayon = {}
        for code, label in display_label.items():
            n_r = n_d[n_d["rayon_code"] == code]
            n1_r = n1_d[n1_d["rayon_code"] == code]
            familles = build_famille_kpis(n_r, n1_r)
            prioritization.apply_priorities(familles, cfg.prioritization)
            drivers.apply_signals_famille(familles, cfg.drivers)
            sf_rows = build_site_famille_kpis(n_r, n1_r)
            drivers.apply_signals_site_famille(sf_rows, cfg.drivers)
            prioritization.assign_site_cle(familles, sf_rows)
            familles_sorted = prioritization.sort_familles_by_impact(familles)
            familles_by_rayon[label] = familles_sorted
            site_famille_by_rayon[label] = sorted(sf_rows, key=lambda r: abs(r.gap_ca_fcfa), reverse=True)

            scope_kpis = {s: kpis_lookup.get((code, s)) for s in [RESEAU, "HYPER", "MARKET", "SUPECO"]}
            reseau_kpi = scope_kpis.get(RESEAU)
            if reseau_kpi is not None:
                diagnostics_by_rayon[label] = diagnostics.build_rayon_diagnostic(
                    reseau_kpi, familles_sorted, [k for k in scope_kpis.values() if k is not None]
                )

        return {
            "cfg": cfg,
            "quality": quality,
            "display_label": display_label,
            "kpis_lookup": kpis_lookup,
            "familles_by_rayon": familles_by_rayon,
            "site_famille_by_rayon": site_famille_by_rayon,
            "diagnostics_by_rayon": diagnostics_by_rayon,
            "period_label": exports.current.period.label,
            "exports": exports,
            "site_enriched": site_enriched,
            "n_enriched": n_enriched,
            "n1_enriched": n1_enriched,
        }


st.title("Reporting Vente CA")
st.caption("Vue BI exploratoire - la saisie Cause / Commentaire se fait dans le fichier Excel exporté")

with st.sidebar:
    st.subheader("Imports")
    site_file = st.file_uploader("Export DATA_SITE", type=["xlsx"])
    current_file = st.file_uploader("Export DATA_N (semaine courante)", type=["xlsx"])
    previous_file = st.file_uploader("Export DATA_N1_ISO", type=["xlsx"])
    config_file = st.file_uploader("Config (optionnel)", type=["yaml"])

if not (site_file and current_file and previous_file):
    st.info("Charge les 3 exports PBI dans la barre latérale pour démarrer l'analyse.")
    st.stop()

data = load_and_process(
    site_file.getvalue(), current_file.getvalue(), previous_file.getvalue(),
    config_file.getvalue() if config_file else None,
)

cfg = data["cfg"]
quality = data["quality"]
rayon_labels = sorted(set(data["display_label"].values()))

tab_reseau, tab_rayon, tab_familles, tab_qualite = st.tabs(
    ["Vue réseau", "Vue rayon", "Vue familles", "Contrôles qualité"]
)

# -----------------------------------------------------------------------------
# Onglet Vue réseau
# -----------------------------------------------------------------------------
with tab_reseau:
    st.caption(f"Semaine : {data['period_label']}")
    cols = st.columns(len(rayon_labels) or 1)
    for col, label in zip(cols, rayon_labels):
        code = next(c for c, l in data["display_label"].items() if l == label)
        kpi = data["kpis_lookup"].get((code, RESEAU))
        if kpi:
            kpi_card(col, label, _fmt_millions(kpi.ca_n), kpi.ca_evo_pct, _fmt_pct(kpi.ca_evo_pct) + " vs N-1 ISO")

    if quality.has_perimeter_issue:
        st.markdown(f'<div class="alert-box">{quality.summary_line}</div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Onglet Vue rayon
# -----------------------------------------------------------------------------
with tab_rayon:
    selected_rayon = st.selectbox("Rayon", rayon_labels, key="rayon_select")
    code = next(c for c, l in data["display_label"].items() if l == selected_rayon)

    scope_cols = st.columns(4)
    for col, scope in zip(scope_cols, [RESEAU, "HYPER", "MARKET", "SUPECO"]):
        kpi = data["kpis_lookup"].get((code, scope))
        if kpi:
            kpi_card(col, scope.capitalize(), _fmt_millions(kpi.ca_n), kpi.ca_evo_pct,
                     _fmt_pct(kpi.ca_evo_pct) + " vs N-1 ISO")

    diag_text = data["diagnostics_by_rayon"].get(selected_rayon, "")
    if diag_text:
        st.markdown(f'<div class="diag-box">{diag_text}</div>', unsafe_allow_html=True)

    familles = data["familles_by_rayon"].get(selected_rayon, [])
    if familles:
        df = pd.DataFrame([{
            "Priorité": f.priority or "-",
            "Famille": f.famille_label,
            "CA N": _fmt_millions(f.ca_n),
            "Écart CA": _fmt_millions(f.gap_ca_fcfa),
            "vs N-1": _fmt_pct(f.evo_pct),
            "Δ Tx marge": _fmt_pts(f.delta_tx_marge_pts),
            "Signal": f.signal,
            "Site clé": f.site_cle,
        } for f in familles])
        st.dataframe(df, use_container_width=True, hide_index=True)

    if quality.has_perimeter_issue:
        st.markdown(f'<div class="alert-box">{quality.summary_line}</div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Onglet Vue familles (multi-rayons, triable)
# -----------------------------------------------------------------------------
with tab_familles:
    all_familles = []
    for label, familles in data["familles_by_rayon"].items():
        for f in familles:
            all_familles.append({
                "Rayon": label, "Priorité": f.priority or "-", "Famille": f.famille_label,
                "Écart CA (FCFA)": f.gap_ca_fcfa, "vs N-1 (%)": (f.evo_pct or 0) * 100,
                "Δ Tx marge (pt)": (f.delta_tx_marge_pts or 0) * 100, "Signal": f.signal,
            })
    df_all = pd.DataFrame(all_familles)
    if not df_all.empty:
        prio_filter = st.multiselect("Filtrer par priorité", ["P1", "P2", "-"], default=["P1", "P2"])
        filtered = df_all[df_all["Priorité"].isin(prio_filter)] if prio_filter else df_all
        st.dataframe(
            filtered.sort_values("Écart CA (FCFA)", key=abs, ascending=False),
            use_container_width=True, hide_index=True,
        )

# -----------------------------------------------------------------------------
# Onglet Contrôles qualité
# -----------------------------------------------------------------------------
with tab_qualite:
    st.subheader("Périmètre de sites")
    cov_rows = [{"Scope": "RÉSEAU", "Couverture": quality.site_coverage_global.label,
                 "%": f"{quality.site_coverage_global.coverage_pct:.0%}"}]
    for fmt, cov in quality.site_coverage_by_format.items():
        cov_rows.append({"Scope": fmt, "Couverture": cov.label, "%": f"{cov.coverage_pct:.0%}"})
    st.dataframe(pd.DataFrame(cov_rows), use_container_width=True, hide_index=True)

    if quality.site_coverage_global.missing_in_site_export:
        st.warning(
            "Sites absents de DATA_SITE (Budget/Débit/Panier indisponibles) : "
            + ", ".join(sorted(quality.site_coverage_global.missing_in_site_export))
        )

    st.subheader("Réconciliation CA (DATA_SITE vs DATA_N, sites communs)")
    recon_rows = [{
        "Scope": "RÉSEAU", "CA DATA_SITE": _fmt_millions(quality.ca_reconciliation_global.ca_site_export),
        "CA DATA_N": _fmt_millions(quality.ca_reconciliation_global.ca_detail_export),
        "Alignement": f"{quality.ca_reconciliation_global.alignment_ratio:.1%}",
        "Statut": quality.ca_reconciliation_global.status,
    }]
    for fmt, recon in quality.ca_reconciliation_by_format.items():
        recon_rows.append({
            "Scope": fmt, "CA DATA_SITE": _fmt_millions(recon.ca_site_export),
            "CA DATA_N": _fmt_millions(recon.ca_detail_export),
            "Alignement": f"{recon.alignment_ratio:.1%}", "Statut": recon.status,
        })
    st.dataframe(pd.DataFrame(recon_rows), use_container_width=True, hide_index=True)

    st.subheader("Schéma")
    for check in quality.schema_checks:
        if check.ok:
            st.success(f"{check.export_name} : toutes les colonnes obligatoires sont présentes.")
        else:
            st.error(f"{check.export_name} : colonnes manquantes -> {', '.join(check.missing_columns)}")

# -----------------------------------------------------------------------------
# Export Excel
# -----------------------------------------------------------------------------
st.divider()
if st.button("Générer le reporting Excel", type="primary"):
    with st.spinner("Génération du fichier Excel..."):
        rayons_sheet_info = [
            (f"{i+1:02d}_{label.replace(' ', '_')[:25]}", label)
            for i, label in enumerate(rayon_labels)
        ]
        coverage_labels = {RESEAU: quality.site_coverage_global.label}
        coverage_labels.update({fmt: cov.label for fmt, cov in quality.site_coverage_by_format.items()})

        kpis_by_rayon = {}
        for label in rayon_labels:
            code = next(c for c, l in data["display_label"].items() if l == label)
            kpis_by_rayon[label] = {
                s: data["kpis_lookup"][(code, s)]
                for s in [RESEAU, "HYPER", "MARKET", "SUPECO"]
                if (code, s) in data["kpis_lookup"]
            }

        wb = excel_writer.build_workbook(
            rayons=rayons_sheet_info,
            week_label=data["period_label"],
            quality=quality,
            coverage_labels_by_rayon={label: coverage_labels for label in rayon_labels},
            kpis_by_rayon=kpis_by_rayon,
            familles_by_rayon=data["familles_by_rayon"],
            site_famille_by_rayon=data["site_famille_by_rayon"],
            diagnostics_by_rayon=data["diagnostics_by_rayon"],
            cfg=cfg,
        )
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.download_button(
            "Télécharger le fichier Excel",
            data=buffer,
            file_name=f"reporting_achats_{data['period_label'].split(' ')[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
