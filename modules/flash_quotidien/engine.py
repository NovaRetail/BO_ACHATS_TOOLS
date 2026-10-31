"""Moteur du flash quotidien : lecture des 3 exports Power BI, contrôles, calculs.

Un seul moteur alimente l'écran Streamlit, l'Excel des équipes et le PowerPoint.
Les règles reprennent celles du fichier Excel validé (Flash_quotidien_Carrefour.xlsx).
"""
from __future__ import annotations

import datetime as dt
import io
import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from . import config as C

# ====================================================================== lecture
COLS = {
    "Département": "dep", "Rayon": "rayon", "Site": "site",
    "CA N-1": "ca_n1", "Budget": "budget", "CA": "ca",
    "Marge N-1": "marge_n1", "Marge": "marge",
    "Débit N-1": "debit_n1", "Débit": "debit",
}


@dataclass
class Export:
    name: str
    kind: str                 # "rayon" (Département > Rayon > Site) ou "site"
    periode: str              # texte du pied de page : "Hier", "Cette Semaine", "Libre - Semaine"...
    ht: bool
    footer: str
    df: pd.DataFrame          # lignes standardisées avec colonne "niveau"


def _footer(raw: pd.DataFrame) -> str:
    for v in raw.iloc[:, 0].fillna("").astype(str):
        if v.startswith("Filtres"):
            return v
    return ""


def read_export(file, name: str = "") -> Export:
    """Lit un export Power BI (.xlsx) et le met au format du moteur.

    Gère les cellules fusionnées (Département et Rayon écrits sur la 1re ligne du groupe).
    """
    if isinstance(file, (bytes, bytearray)):
        file = io.BytesIO(file)
    raw = pd.read_excel(file, header=0)
    footer = _footer(raw)
    m = re.search(r"Type d'affichage est ([^\n]+)", footer)
    periode = m.group(1).strip() if m else "?"
    ht = "Libelle est HT" in footer
    # on coupe au pied de page
    first = raw.columns[0]
    stop = raw.index[raw[first].fillna("").astype(str).str.startswith("Filtres")]
    if len(stop):
        raw = raw.loc[: stop[0] - 1]
    keep = {c: COLS[c] for c in raw.columns if c in COLS}
    df = raw[list(keep)].rename(columns=keep).copy()
    kind = "rayon" if "dep" in df.columns else "site"
    if kind == "rayon":
        df["dep"] = df["dep"].ffill()
        df["rayon"] = df["rayon"].ffill()
        site = df["site"].fillna("").astype(str)
        df["niveau"] = np.select(
            [df["dep"].eq("Total").to_numpy(bool), df["rayon"].eq("Total").to_numpy(bool),
             site.eq("Total").to_numpy(bool), site.str[:2].str.isdigit().to_numpy(bool)],
            ["total", "departement", "rayon", "detail"], default="vide")
    else:
        site = df["site"].fillna("").astype(str)
        df["niveau"] = np.select([site.eq("Total").to_numpy(bool), site.str[:2].str.isdigit().to_numpy(bool)],
                                 ["total", "detail"], default="vide")
    df = df[df["niveau"] != "vide"].copy()
    for c in ("ca_n1", "budget", "ca", "marge_n1", "marge", "debit_n1", "debit"):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return Export(name=name, kind=kind, periode=periode, ht=ht, footer=footer, df=df)


# ====================================================================== contrôles
@dataclass
class Check:
    statut: str   # "ok", "alerte", "erreur"
    message: str


EXTRACTIONS = {
    "hier": {
        "titre": "1. Ventes d'hier par rayon et par site",
        "attendu_kind": "rayon", "attendu_periode": "Hier",
        "consigne": [
            "Power BI, page « Ventes CA J-1 », bouton « Hier », affichage HT.",
            "Matrice : Département > Rayon > Site, dépliée jusqu'au site (« Développer tout »).",
            "Tous les départements et les 12 sites du périmètre, sans filtre rayon.",
            "Menu « … » du visuel > Exporter les données > Données avec la disposition actuelle > .xlsx.",
        ],
        "colonnes": ["Département", "Rayon", "Site", "CA N-1", "Budget", "CA", "Marge N-1", "Marge", "Débit N-1", "Débit"],
        "pied": "Type d'affichage est Hier · Libelle est HT",
        "fichier": "data_hier_rayon.xlsx",
        "pieges": "Ne pas prendre « Libre - Semaine » ni une semaine passée (W43…) : le pied de page doit dire « Hier ».",
    },
    "semaine": {
        "titre": "2. Cumul de la semaine par rayon et par site",
        "attendu_kind": "rayon", "attendu_periode": "Cette Semaine",
        "consigne": [
            "Même page et même matrice que l'export 1.",
            "Bouton « Cette semaine » au lieu de « Hier » (cumul depuis lundi).",
            "Même export : disposition actuelle, .xlsx.",
            "Le lundi, cet export ne contient que la journée de lundi : c'est normal.",
        ],
        "colonnes": ["Département", "Rayon", "Site", "CA N-1", "Budget", "CA", "Marge N-1", "Marge"],
        "pied": "Type d'affichage est Cette Semaine · Libelle est HT",
        "fichier": "data_semaine_rayon.xlsx",
        "pieges": "La colonne Budget doit être remplie : c'est elle qui donne l'écart au budget de la semaine.",
    },
    "sites": {
        "titre": "3. Ventes d'hier par site (débit et panier)",
        "attendu_kind": "site", "attendu_periode": "Hier",
        "consigne": [
            "Même page, bouton « Hier », affichage HT.",
            "Matrice repliée au niveau Site uniquement (une ligne par magasin, pas de rayon).",
            "C'est la seule vue qui donne le débit et le panier par magasin.",
            "Exporter les données > disposition actuelle > .xlsx.",
        ],
        "colonnes": ["Site", "CA N-1", "CA", "Marge N-1", "Marge", "Débit N-1", "Débit", "Panier N-1", "Panier"],
        "pied": "Type d'affichage est Hier · Libelle est HT · 1 - Axes est Site",
        "fichier": "data_hier_site.xlsx",
        "pieges": "Pas de colonne Département ni Rayon : sinon le débit est compté plusieurs fois par ticket.",
    },
}


def check_exports(hier: Export | None, semaine: Export | None, sites: Export | None) -> list[Check]:
    out: list[Check] = []
    for key, exp in (("hier", hier), ("semaine", semaine), ("sites", sites)):
        spec = EXTRACTIONS[key]
        if exp is None:
            out.append(Check("erreur", f"{spec['titre']} : fichier manquant."))
            continue
        if exp.kind != spec["attendu_kind"]:
            vue = "Département > Rayon > Site" if spec["attendu_kind"] == "rayon" else "Site seul"
            out.append(Check("erreur", f"{spec['titre']} : mauvaise vue, il faut la vue {vue}."))
            continue
        if exp.periode.lower() != spec["attendu_periode"].lower():
            out.append(Check("erreur", f"{spec['titre']} : période « {exp.periode} » au lieu de « {spec['attendu_periode']} »."))
            continue
        if not exp.ht:
            out.append(Check("alerte", f"{spec['titre']} : l'affichage HT n'est pas confirmé dans le pied de page."))
        if exp.kind == "rayon":
            tot = exp.df.loc[exp.df.niveau == "total", "ca"].sum()
            det = exp.df.loc[exp.df.niveau == "detail", "ca"].sum()
            if tot and abs(det / tot - 1) > 0.001:
                out.append(Check("erreur", f"{spec['titre']} : le détail ({det:,.0f}) ne retombe pas sur le total ({tot:,.0f})."))
                continue
        out.append(Check("ok", f"{spec['titre']} : conforme."))
    if hier is not None and sites is not None and hier.kind == "rayon" and sites.kind == "site":
        th = hier.df.loc[hier.df.niveau == "total", "ca"].sum()
        ts = sites.df.loc[sites.df.niveau == "total", "ca"].sum()
        if th and abs(ts / th - 1) > 0.001:
            out.append(Check("erreur", "Les exports 1 et 3 ne portent pas sur la même journée (totaux de CA différents)."))
        else:
            out.append(Check("ok", "Exports 1 et 3 : même journée, totaux concordants."))
    if hier is not None and hier.kind == "rayon":
        presents = set(hier.df.loc[hier.df.niveau == "detail", "site"])
        attendus = {s[0] for s in C.SITES}
        manquants = attendus - presents
        inconnus = presents - attendus
        if manquants:
            out.append(Check("alerte", "Sites absents de l'export : " + ", ".join(sorted(manquants))))
        if inconnus:
            out.append(Check("alerte", "Sites hors paramètres (ignorés) : " + ", ".join(sorted(inconnus))))
    return out


def bloquant(checks: list[Check]) -> bool:
    return any(c.statut == "erreur" for c in checks)


# ====================================================================== indicateurs
def kpi(ca_n1, ca, marge_n1, marge, debit_n1=None, debit=None, budget=None, ca_budget=None) -> dict:
    """Indicateurs standard d'un périmètre."""
    d = {"ca": ca, "ca_n1": ca_n1, "marge": marge, "marge_n1": marge_n1}
    d["vs_n1"] = ca / ca_n1 - 1 if ca_n1 else None
    d["marge_vs_n1"] = marge / marge_n1 - 1 if marge_n1 else None
    d["dmarge"] = marge - marge_n1
    d["taux"] = marge / ca if ca else None
    d["taux_n1"] = marge_n1 / ca_n1 if ca_n1 else None
    d["dpts"] = (d["taux"] - d["taux_n1"]) * 100 if d["taux"] is not None and d["taux_n1"] is not None else None
    d["effet_volume"] = (ca - ca_n1) * d["taux_n1"] if d["taux_n1"] is not None else None
    d["effet_taux"] = (d["taux"] - d["taux_n1"]) * ca if d["dpts"] is not None else None
    if debit_n1 is not None and debit is not None:
        d["debit_vs_n1"] = debit / debit_n1 - 1 if debit_n1 else None
        d["panier_vs_n1"] = (ca / debit) / (ca_n1 / debit_n1) - 1 if debit and debit_n1 and ca_n1 else None
    if budget is not None:
        d["vs_budget"] = ca_budget / budget - 1 if budget else None
    return d


def _sum(df: pd.DataFrame) -> dict:
    s = df[["ca_n1", "ca", "marge_n1", "marge"]].sum()
    return dict(ca_n1=s.ca_n1, ca=s.ca, marge_n1=s.marge_n1, marge=s.marge)


def is_alert(k: dict, s=C.SEUILS) -> bool:
    return bool((k.get("vs_n1") is not None and k["vs_n1"] < s["ca_vs_n1"])
                or (k.get("dpts") is not None and k["dpts"] < s["taux_pts"])
                or (k.get("marge_vs_n1") is not None and k["marge_vs_n1"] < s["marge_vs_n1"]))


def breach_text(k: dict, s=C.SEUILS) -> str:
    """Écarts hors seuil uniquement, sous forme courte."""
    parts = []
    if k.get("vs_n1") is not None and k["vs_n1"] < s["ca_vs_n1"]:
        parts.append(f"CA {pct(k['vs_n1'], 0)}")
    if k.get("marge_vs_n1") is not None and k["marge_vs_n1"] < s["marge_vs_n1"]:
        parts.append(f"Marge {pct(k['marge_vs_n1'], 0)}")
    if k.get("dpts") is not None and k["dpts"] < s["taux_pts"]:
        parts.append(f"Taux {pts(k['dpts'])}")
    return " · ".join(parts)


# ====================================================================== formats français
def _fr(x: float, dec: int) -> str:
    # espace fine insécable pour les milliers, virgule décimale
    s = f"{abs(x):,.{dec}f}".replace(",", "\u202f").replace(".", ",")
    return s


NEG = "\u2212"   # vrai signe moins : ne se sépare jamais du chiffre à la ligne
NB = "\u00a0"    # espace insécable avant %, pt, M, k


def pct(x, dec=1, sign=True) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "n.d."
    v = round(x * 100, dec)
    sg = ("+" if v >= 0 else NEG) if sign else (NEG if v < 0 else "")
    return f"{sg}{_fr(v, dec)}{NB}%"


def pts(x, dec=1) -> str:
    if x is None:
        return "n.d."
    v = round(x, dec)
    if v == 0:
        return f"{_fr(0, dec)}{NB}pt"
    return f"{'+' if v > 0 else NEG}{_fr(v, dec)}{NB}pt"


def mfcfa(x, dec=1, sign=False) -> str:
    if x is None:
        return "n.d."
    sg = ("+" if x >= 0 else NEG) if sign else (NEG if x < 0 else "")
    return f"{sg}{_fr(x / 1e6, dec)}{NB}M"


def kfcfa(x, sign=True) -> str:
    if x is None:
        return "n.d."
    sg = ("+" if x >= 0 else NEG) if sign else ""
    return f"{sg}{_fr(x / 1e3, 0)}{NB}k"


def short_site(site: str) -> str:
    """'10208 - Market Riviera' -> 'Riviera'."""
    name = site.split(" - ", 1)[-1]
    for p in ("Hyper ", "Market ", "Supeco "):
        if name.startswith(p):
            return name[len(p):].strip().title() if name[len(p):].isupper() else name[len(p):].strip()
    return name


def short_rayon(r: str) -> str:
    """'014 - EPICERIE' -> 'Épicerie'."""
    name = r.split(" - ", 1)[-1].strip().capitalize()
    return (name.replace("Epicerie", "Épicerie").replace("Electro", "électro").replace("electro", "électro")
            .replace("Legumes", "légumes").replace("Patisserie", "pâtisserie").replace("detente", "détente"))


# ====================================================================== calcul
def date_reponse(d: dt.date) -> dt.date:
    """Lendemain, ou lundi si le flash porte sur un vendredi ou un samedi."""
    nxt = d + dt.timedelta(days=1)
    while nxt.weekday() >= 5:
        nxt += dt.timedelta(days=1)
    return nxt


def compute(hier: Export, semaine: Export, sites: Export, date_flash: dt.date, s=C.SEUILS) -> dict:
    site_info = pd.DataFrame(C.SITES, columns=["site", "format", "zone", "converti_le"])
    site_info["zone"] = "Zone " + site_info["zone"]
    sec_map = C.DEPARTEMENTS
    sites_ok = sites is not None and abs(
        sites.df.loc[sites.df.niveau == "total", "ca"].sum() / hier.df.loc[hier.df.niveau == "total", "ca"].sum() - 1) < 0.001

    H, W = hier.df, semaine.df
    hd = H[(H.niveau == "detail") & H.site.isin(site_info.site)].copy()
    wd = W[(W.niveau == "detail") & W.site.isin(site_info.site)].copy()
    for d in (hd, wd):
        d["secteur"] = d["dep"].map(sec_map)
        d["format"] = d["site"].map(site_info.set_index("site")["format"])
        d["zone"] = d["site"].map(site_info.set_index("site")["zone"])
        d["converti"] = d["site"].map(site_info.set_index("site")["converti_le"].notna())

    # budget : seuls les sites budgétés (budget semaine > 0)
    bud_site = wd.groupby("site")["budget"].sum()
    budgetes = set(bud_site[bud_site > 0].index)
    wd["ca_budgete"] = np.where(wd.site.isin(budgetes), wd.ca, 0.0)

    def sem_kpi(mask):
        d = wd[mask]
        return kpi(**_sum(d), budget=d.budget.sum(), ca_budget=d.ca_budgete.sum())

    # -------- réseau
    tot = H[H.niveau == "total"].iloc[0]
    reseau = kpi(tot.ca_n1, tot.ca, tot.marge_n1, tot.marge, tot.debit_n1, tot.debit)
    reseau_sem = sem_kpi(wd.site.notna())
    hors_conv = kpi(**_sum(hd[~hd.converti]))

    # -------- débit / panier par site (export Site)
    deb = pd.DataFrame(columns=["debit_n1", "debit"])
    if sites_ok:
        deb = sites.df[sites.df.niveau == "detail"].set_index("site")[["debit_n1", "debit"]]

    def grp_kpi(mask_sites: pd.Series):
        d = hd[mask_sites]
        base = _sum(d)
        ss = d.site.unique()
        if sites_ok and len(ss):
            base.update(debit_n1=deb.reindex(ss).debit_n1.sum(), debit=deb.reindex(ss).debit.sum())
        k = kpi(**base)
        return k

    # -------- secteurs
    secteurs = {}
    for sec in C.SECTEURS:
        k = kpi(**_sum(hd[hd.secteur == sec]))
        k["sem"] = sem_kpi(wd.secteur == sec)
        secteurs[sec] = k

    # -------- formats, zones
    formats = {f: grp_kpi(hd.format == f) for f in C.FORMATS}
    zones = {z: grp_kpi(hd.zone == z) for z in ("Zone AB", "Zone CD")}

    # -------- secteur x format (texte hors seuil seulement)
    croise = {}
    for sec in C.SECTEURS:
        croise[sec] = {}
        for f in C.FORMATS:
            k = kpi(**_sum(hd[(hd.secteur == sec) & (hd.format == f)]))
            croise[sec][f] = {"kpi": k, "texte": breach_text(k, s)}

    # -------- sites
    site_rows = []
    for _, si in site_info.iterrows():
        k = grp_kpi(hd.site == si.site)
        k["sem"] = sem_kpi(wd.site == si.site)
        k["budgete"] = si.site in budgetes
        conv = si.converti_le if pd.notna(si.converti_le) else None
        k.update(site=si.site, nom=short_site(si.site), format=si.format, zone=si.zone,
                 converti_le=conv, alerte=is_alert(k, s))
        site_rows.append(k)
    sites_alerte = [r for r in site_rows if r["converti_le"] is None and r["alerte"]]
    convertis = [r for r in site_rows if r["converti_le"] is not None]

    # -------- alertes rayon x site
    key = ["dep", "rayon", "site"]
    wkey = wd.set_index(key)[["ca_n1", "ca", "marge_n1", "marge"]]
    al = []
    for _, r in hd.iterrows():
        if r.ca_n1 < s["ca_n1_min"]:
            continue
        k = kpi(r.ca_n1, r.ca, r.marge_n1, r.marge, r.debit_n1, r.debit)
        a_ca = k["vs_n1"] is not None and k["vs_n1"] < s["ca_vs_n1"]
        a_m = r.ca > 0 and ((k["dpts"] is not None and k["dpts"] < s["taux_pts"])
                            or (k["marge_vs_n1"] is not None and k["marge_vs_n1"] < s["marge_vs_n1"]))
        if not (a_ca or a_m):
            continue
        if not a_ca:
            diag = "Marge en recul : prix, promo ou mix"
        elif k.get("debit_vs_n1") is not None and k["debit_vs_n1"] < s["debit_vs_n1"]:
            diag = "Trafic en baisse"
        elif k.get("panier_vs_n1") is not None and k["panier_vs_n1"] < s["panier_vs_n1"]:
            diag = "Panier en baisse : offre ou ruptures"
        else:
            diag = "CA en baisse"
        qui = "Magasin" if diag == "Trafic en baisse" else ("Acheteur + Supply" if diag.startswith("Panier") else "Acheteur")
        try:
            w = wkey.loc[(r.dep, r.rayon, r.site)]
            wk = kpi(w.ca_n1, w.ca, w.marge_n1, w.marge)
            persist = is_alert(wk, s)
        except KeyError:
            persist = False
        al.append(dict(k, rayon=r.rayon, site=r.site, secteur=r.secteur, format=r.format,
                       type="CA et marge" if a_ca and a_m else ("CA" if a_ca else "Marge"),
                       diagnostic=diag, qui=qui, persistant=persist))
    alertes = sorted(al, key=lambda x: x["dmarge"])

    # -------- rayons (tous sites)
    rt = H[H.niveau == "rayon"].copy()
    rt["secteur"] = rt.dep.map(sec_map)
    wr = W[W.niveau == "rayon"].set_index(["dep", "rayon"])[["ca_n1", "ca", "marge_n1", "marge"]]
    rayons = {sec: [] for sec in C.SECTEURS}
    for _, r in rt.iterrows():
        if r.secteur not in rayons:
            continue
        k = kpi(r.ca_n1, r.ca, r.marge_n1, r.marge, r.debit_n1, r.debit)
        try:
            w = wr.loc[(r.dep, r.rayon)]
            k["sem"] = kpi(w.ca_n1, w.ca, w.marge_n1, w.marge)
        except KeyError:
            k["sem"] = {}
        a_traiter = ((k["vs_n1"] is not None and k["vs_n1"] < 0) or k["dmarge"] < 0
                     or (k["dpts"] is not None and k["dpts"] < s["taux_pts"] / 2))
        if not a_traiter:
            continue
        if (k["dpts"] is not None and k["dpts"] < s["taux_pts"] / 2) or (k["dmarge"] < 0 and (k["vs_n1"] or 0) >= 0):
            action = "Contrôler prix de vente, promos sous marge et mix"
        elif k.get("debit_vs_n1") is not None and k["debit_vs_n1"] < s["debit_vs_n1"]:
            action = "Trafic rayon en baisse : visibilité, implantation, concurrence"
        elif k.get("panier_vs_n1") is not None and k["panier_vs_n1"] < s["panier_vs_n1"]:
            action = "Panier en baisse : ruptures et disponibilité des tops"
        else:
            action = "Vérifier disponibilité et prix des tops"
        rayons[r.secteur].append(dict(k, rayon=r.rayon, nom=short_rayon(r.rayon), action=action))
    for sec in rayons:
        rayons[sec].sort(key=lambda x: x["dmarge"])

    # -------- retours attendus
    rep = date_reponse(date_flash)
    retours = build_retours(rayons, alertes, sites_alerte, convertis, s)

    return dict(
        date=date_flash, date_reponse=rep, sites_ok=sites_ok,
        reseau=reseau, reseau_sem=reseau_sem, hors_convertis=hors_conv,
        secteurs=secteurs, formats=formats, zones=zones, croise=croise,
        sites=site_rows, sites_alerte=sites_alerte, convertis=convertis,
        alertes=alertes, rayons=rayons, retours=retours,
        lecture=lecture_reseau(reseau, s), messages=messages_cles(reseau, reseau_sem, secteurs, formats),
    )


# ====================================================================== textes
def lecture_reseau(k: dict, s=C.SEUILS) -> str:
    if k["dmarge"] < 0:
        return "La marge en valeur recule : priorité absolue."
    if k.get("debit_vs_n1") is not None and k["debit_vs_n1"] >= s["debit_vs_n1"] and (k.get("panier_vs_n1") or 0) < s["panier_vs_n1"]:
        return "Débit stable, panier en baisse : problème d'offre."
    if k.get("debit_vs_n1") is not None and k["debit_vs_n1"] < s["debit_vs_n1"]:
        return "Débit en baisse : problème de trafic."
    if (k["dpts"] or 0) < 0:
        return "Le volume compense la baisse du taux : surveiller prix, promo et mix."
    return "Volume et taux progressent."


def messages_cles(res: dict, sem: dict, secteurs: dict, formats: dict) -> list[str]:
    """Trois points à retenir, rédigés à partir des écarts les plus forts."""
    out = []
    ev, et = res["effet_volume"], res["effet_taux"]
    if et is not None and et < 0 < ev:
        out.append(f"Le volume rapporte {mfcfa(ev, sign=True)} de marge, la baisse du taux en coûte {mfcfa(-et)}.")
    else:
        out.append(f"Marge {mfcfa(res['dmarge'], sign=True)} vs N-1 : volume {mfcfa(ev, sign=True)}, taux {mfcfa(et, sign=True)}.")
    perdants = [n for n, k in secteurs.items() if k["dmarge"] < 0]
    if perdants:
        n = perdants[0]
        out.append(f"Le {n.lower()} est le seul secteur à perdre de la marge ({pct(secteurs[n]['marge_vs_n1'])})."
                   if len(perdants) == 1 else f"{len(perdants)} secteurs perdent de la marge : {', '.join(perdants)}.")
    else:
        worst = min(secteurs, key=lambda n: secteurs[n]["dpts"] if secteurs[n]["dpts"] is not None else 0)
        out.append(f"Tous les secteurs gagnent de la marge ; le taux recule le plus en {worst.lower()} ({pts(secteurs[worst]['dpts'])}).")
    vb = [k["sem"].get("vs_budget") for k in secteurs.values() if k["sem"].get("vs_budget") is not None]
    if vb and max(vb) < 0:
        out.append(f"Sur la semaine, tous les secteurs sont sous le budget (de {pct(max(vb), 0)} à {pct(min(vb), 0)}).")
    elif sem.get("vs_budget") is not None:
        out.append(f"Sur la semaine, le réseau est à {pct(sem['vs_budget'])} du budget.")
    return out[:3]


def build_retours(rayons: dict, alertes: list, sites_alerte: list, convertis: list, s=C.SEUILS) -> list[dict]:
    """Questions à poser, à qui, avec les sujets hors périmètre en « pour information »."""
    out = []
    P = C.POLES
    for sec in C.SECTEURS:
        for r in rayons[sec][:2]:
            if r["dpts"] is not None and r["dpts"] < s["taux_extreme_pts"]:
                q = f"Taux {pts(r['dpts'])} : déstockage voulu ou erreur de prix ?"
                qui = f"{sec} · {P['bdd']}"
            elif r["dpts"] is not None and r["dpts"] < s["taux_pts"] / 2:
                q = f"Taux {pts(r['dpts'])} : promos vendues sous marge ou prix à revoir ?"
                qui = f"{sec} · {P['pricing']}"
            elif r["action"].startswith("Panier"):
                q = f"CA {pct(r['vs_n1'])}, panier {pct(r.get('panier_vs_n1'))} : ruptures ou prix ?"
                qui = sec
            elif r["action"].startswith("Trafic"):
                q = f"CA {pct(r['vs_n1'])}, débit {pct(r.get('debit_vs_n1'))} : visibilité ou concurrence ?"
                qui = sec
            else:
                q = f"CA {pct(r['vs_n1'])} : disponibilité et prix des tops ?"
                qui = sec
            out.append(dict(sujet=f"{r['nom']} · {sec}", question=q, qui=qui, type="Réponse attendue",
                            poids=r["dmarge"]))
    out.sort(key=lambda x: x["poids"])
    out = out[: C.MAX_RETOURS - 1]
    # magasins et supply : pour information (hors périmètre)
    infos = []
    for a in alertes:
        if a["diagnostic"].startswith("Panier") and len(infos) < 1:
            infos.append(dict(sujet=f"{short_rayon(a['rayon'])} · {short_site(a['site'])}",
                              question=f"Panier {pct(a.get('panier_vs_n1'))} : ruptures possibles", qui="Supply", type="Pour information"))
    for r in sites_alerte:
        if r.get("debit_vs_n1") is not None and r["debit_vs_n1"] < s["debit_vs_n1"] and len(infos) < 2:
            infos.append(dict(sujet=r["nom"], question=f"Débit {pct(r['debit_vs_n1'])} : travaux, concurrence, horaires ?",
                              qui="Direction des magasins", type="Pour information"))
    return out + infos
