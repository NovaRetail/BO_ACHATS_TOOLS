"""
patch_sidebar.py
Mise à jour automatique du bloc st.page_link() dans toutes les pages SmartBuyer.
À exécuter depuis la racine du projet : python patch_sidebar.py
"""

import re
from pathlib import Path

# ─── Navigation complète à injecter dans toutes les pages ─────────────────────
NAV_BLOCK = '''    st.page_link("app.py",                                       label="🏠  Accueil")
    st.page_link("pages/01_📊_Analyse_Scoring_ABC.py",           label="📊  Scoring ABC")
    st.page_link("pages/02_📈_Ventes_PBI.py",                    label="📈  Ventes PBI")
    st.page_link("pages/03_📦_Detention_Top_CA.py",              label="📦  Détention Top CA")
    st.page_link("pages/04_💸_Performance_Promo.py",             label="💸  Performance Promo")
    st.page_link("pages/05_🏪_Suivi_Implantation.py",            label="🏪  Suivi Implantation")
    st.page_link("pages/06_💸_Marges_Negatives.py",              label="💸  Marges Négatives")'''

# ─── Pattern : bloc de st.page_link consécutifs ───────────────────────────────
# Capture tout ce qui est une suite de lignes st.page_link(...)
PAGE_LINK_PATTERN = re.compile(
    r'([ \t]*st\.page_link\([^\n]+\)\n)+',
    re.MULTILINE
)

# ─── Fichiers à patcher ───────────────────────────────────────────────────────
PAGES_DIR = Path("pages")
TARGET_FILES = [
    PAGES_DIR / "01_📊_Analyse_Scoring_ABC.py",
    PAGES_DIR / "02_📈_Ventes_PBI.py",
    PAGES_DIR / "03_📦_Detention_Top_CA.py",
    PAGES_DIR / "04_💸_Performance_Promo.py",
    PAGES_DIR / "05_🏪_Suivi_Implantation.py",
    PAGES_DIR / "06_💸_Marges_Negatives.py",
]

# ─── Patch ────────────────────────────────────────────────────────────────────
def patch_file(path: Path) -> str:
    if not path.exists():
        return f"  ⚠️  Fichier introuvable : {path}"

    source = path.read_text(encoding="utf-8")

    matches = list(PAGE_LINK_PATTERN.finditer(source))
    if not matches:
        return f"  ⚠️  Aucun bloc st.page_link trouvé dans {path.name}"

    # Remplacer le premier bloc trouvé (celui de la sidebar navigation)
    match = matches[0]
    new_source = (
        source[: match.start()]
        + NAV_BLOCK + "\n"
        + source[match.end() :]
    )

    if new_source == source:
        return f"  ✅  Déjà à jour : {path.name}"

    path.write_text(new_source, encoding="utf-8")
    return f"  ✅  Patché : {path.name}"


if __name__ == "__main__":
    print("=" * 55)
    print("  SmartBuyer — Patch sidebar navigation")
    print("=" * 55)

    results = []
    for f in TARGET_FILES:
        results.append(patch_file(f))

    for r in results:
        print(r)

    print("=" * 55)
    print("  Terminé. Recharge l'app Streamlit pour voir les changements.")
    print("=" * 55)
