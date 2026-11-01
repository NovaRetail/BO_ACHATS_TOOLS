"""SmartBuyer Hub · page d'accueil.

Les modules sont découverts dans pages/ ; le registre META donne leur nom, leur usage et leur rubrique.
Un module retiré du dossier pages/ (rangé sur la branche « archive ») disparaît de l'accueil sans autre changement.
"""
from pathlib import Path

import streamlit as st

from core import theme as T

st.set_page_config(page_title="SmartBuyer Hub", page_icon="🛍️", layout="wide")
T.apply()

PAGES_DIR = Path(__file__).resolve().parent / "pages"

# num : (nom, à quoi il sert, quand on l'ouvre, rubrique) — l'ordre est l'ordre d'affichage
META = {
    "27": ("Flash Quotidien", "Ventes de la veille par secteur, retours attendus et PowerPoint.", "Chaque matin", "pilotage"),
    "26": ("Revue CODIR", "Entonnoir de la semaine Vs N-1 et rapport PowerPoint du CODIR.", "Lundi", "pilotage"),
    "24": ("PGC Radar", "Réel vs N-1 et budget, format, rayon et alertes rayon × magasin.", "Chaque semaine", "pilotage"),
    "06": ("Marge négative", "Ventes à perte article × magasin, Flop 100 et effet promo.", "Chaque semaine", "marge"),
    "05": ("Suivi Implantation", "Taux d'implantation T1, actions par article et cessions.", "À chaque vague T1", "assort"),
    "07": ("OTIF", "Taux de service fournisseurs et plan d'action acheteurs.", "Chaque semaine", "assort"),
}
SECTIONS = [("pilotage", "📋 Pilotage"), ("marge", "💸 Marge"), ("assort", "🏪 Assortiment et fournisseurs"), ("autres", "🧩 Autres")]


def discover():
    out = []
    for p in sorted(PAGES_DIR.glob("*.py")):
        parts = p.stem.split("_")
        if p.name.startswith("_") or not parts[0].isdigit():
            continue
        num, emoji = parts[0], (parts[1] if len(parts) > 1 else "•")
        name, desc, when, sec = META.get(num, (" ".join(parts[2:]) or p.stem, "", "", "autres"))
        out.append(dict(num=num, emoji=emoji, name=name, desc=desc, when=when, sec=sec, path=f"pages/{p.name}"))
    order = list(META)
    return sorted(out, key=lambda m: order.index(m["num"]) if m["num"] in order else 99)


MODS = discover()

st.markdown("""<style>
[data-testid="stSidebarNav"] { display: none !important; }
[data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"] {
  background: #fff; border-radius: 16px !important; border: 1px solid #E3E5EE !important;
  box-shadow: 0 1px 2px rgba(20,20,60,.04), 0 8px 22px rgba(20,20,60,.05); padding: 6px 8px 10px !important; }
[data-testid="stMain"] [data-testid="stPageLink"] p { font-size: 15px !important; font-weight: 800 !important; color: #24235C !important; }
.mod-when { display:inline-block; font-size: 11px; font-weight: 700; color: #B4610A; background: #FDF1E2;
  border-radius: 999px; padding: 2px 9px; margin: 0 0 4px 10px; }
.mod-desc { font-size: 12.5px; color: #6B7090; margin: 0 0 4px 10px; line-height: 1.45; min-height: 36px; }
</style>""", unsafe_allow_html=True)

# ── Barre latérale : navigation par rubrique ──────────────────────────────────
T.sidebar_brand("🛍️", "SmartBuyer Hub", "Achats et category management")
for key, label in SECTIONS:
    group = [m for m in MODS if m["sec"] == key]
    if group:
        T.sidebar_label(label)
        with st.sidebar:
            for m in group:
                st.page_link(m["path"], label=f"{m['emoji']}  {m['name']}")
T.sidebar_footer("Carrefour Côte d'Ivoire · CFAO Consumer<br>Usage personnel · les équipes reçoivent les exports Excel et PowerPoint")

# ── Bandeau ───────────────────────────────────────────────────────────────────
rows = [(f"{m['emoji']} {m['name']}", "ok", m["when"]) for m in MODS if m["sec"] == "pilotage"]
T.landing_hero("SMARTBUYER <b>HUB</b>",
               "Les outils de pilotage achats et category management : on dépose les exports Power BI ou ERP, "
               "on lit, puis on diffuse un Excel ou un PowerPoint aux équipes.",
               eyebrow="Carrefour Côte d'Ivoire · CFAO Consumer",
               status=T.landing_status("Rituels de pilotage", rows, "modules de pilotage"))

T.kpi_row([
    T.kpi("Modules actifs", str(len(MODS)), "", "mut", "même charte CFAO", "blue", "🧩"),
    T.kpi("Rubriques", str(sum(1 for k, _ in SECTIONS if any(m["sec"] == k for m in MODS))), "", "mut", "pilotage, marge, assortiment", "violet", "🗂️"),
    T.kpi("Sources", "PBI · ERP", "", "mut", "exports déposés à la main", "teal", "📥"),
    T.kpi("Sorties", "Excel · PPT", "", "mut", "valeurs uniquement, sans formule", "", "📤"),
])

T.steps([("Choisir", "le module selon le rituel du jour"), ("Exporter", "les fichiers Power BI ou ERP indiqués sur sa page d'accueil"),
         ("Déposer", "dans la barre latérale du module"), ("Lire", "le bandeau, les indicateurs puis les onglets"),
         ("Diffuser", "l'Excel ou le PowerPoint généré")])

# ── Modules ───────────────────────────────────────────────────────────────────
for key, label in SECTIONS:
    group = [m for m in MODS if m["sec"] == key]
    if not group:
        continue
    T.section(label)
    for i in range(0, len(group), 3):
        cols = st.columns(3)
        for col, m in zip(cols, group[i:i + 3]):
            with col, st.container(border=True):
                st.page_link(m["path"], label=f"{m['emoji']}  {m['name']}  ›")
                T.html((f"<div class='mod-when'>{m['when']}</div>" if m["when"] else "") + f"<div class='mod-desc'>{m['desc'] or '&nbsp;'}</div>")

T.html("<div class='note' style='text-align:center;margin-top:18px'>Les modules retirés (Reporting Sous-Familles, Reporting Vente CA, "
       "Suivi Rayon vs Pays et les plus anciens) sont conservés sur la branche « archive » du dépôt.</div>")
