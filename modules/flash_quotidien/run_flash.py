"""Génère le flash commercial sans passer par Streamlit.

Usage :
    python -m modules.flash_quotidien.run_flash HIER.xlsx SEMAINE.xlsx SITES.xlsx MASQUE.pptx [--date 2026-10-30] [--out dossier]

Produit dans le dossier de sortie :
    Flash_commercial_AAAA-MM-JJ.pptx  (support de communication pour les secteurs)
    Flash_quotidien_AAAA-MM-JJ.xlsx   (détail en valeurs, sans formules)
"""
from __future__ import annotations

import argparse
import datetime as dt
import pathlib
import sys

from modules.flash_quotidien import engine as E
from modules.flash_quotidien.excel_report import build_excel
from modules.flash_quotidien.ppt_report import build_ppt


def main() -> int:
    ap = argparse.ArgumentParser(description="Flash commercial quotidien")
    ap.add_argument("hier"); ap.add_argument("semaine"); ap.add_argument("sites"); ap.add_argument("masque")
    ap.add_argument("--date", help="date des ventes (défaut : hier)")
    ap.add_argument("--out", default=".")
    ap.add_argument("--force", action="store_true", help="générer malgré un contrôle bloquant")
    a = ap.parse_args()

    d = dt.date.fromisoformat(a.date) if a.date else dt.date.today() - dt.timedelta(days=1)
    h = E.read_export(a.hier, pathlib.Path(a.hier).name)
    s = E.read_export(a.semaine, pathlib.Path(a.semaine).name)
    t = E.read_export(a.sites, pathlib.Path(a.sites).name)

    checks = E.check_exports(h, s, t)
    for c in checks:
        print(f"[{c.statut}] {c.message}")
    if E.bloquant(checks) and not a.force:
        print("Contrôle bloquant : corrigez les extractions (ou --force).")
        return 1

    R = E.compute(h, s, t, d)
    out = pathlib.Path(a.out); out.mkdir(parents=True, exist_ok=True)
    with open(a.masque, "rb") as m:
        (out / f"Flash_commercial_{d.isoformat()}.pptx").write_bytes(build_ppt(R, m))
    (out / f"Flash_quotidien_{d.isoformat()}.xlsx").write_bytes(build_excel(R))
    print("OK :", out.resolve())
    return 0


if __name__ == "__main__":
    sys.exit(main())
