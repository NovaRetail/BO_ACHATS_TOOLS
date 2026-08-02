"""
utils_promo.py — SmartBuyer Hub / Reporting Promo Performance
=============================================================
ÉTAPE 1 : ingestion robuste + contrôles qualité.
Aucune dépendance à Streamlit ici (pur pandas) -> testable en CLI.
Charte : réutilise l'esprit utils_io (fallback encodage, safe_sort).
"""

from __future__ import annotations
import re
import io
import pandas as pd
import numpy as np

# --------------------------------------------------------------------------- #
# Constantes
# --------------------------------------------------------------------------- #
SHEET_DEFAULT = "Export"

# Colonnes attendues dans l'extraction ventes (clé interne -> libellé source)
VENTES_MAP = {
    "departement":      "Departement",      # optionnel : info contextuelle, n'intervient pas dans le filtre de grain
    "rayon":            "Rayon",
    "famille":          "Famille",
    "sous_famille":     "Sous Famille",
    "article":          "Article",
    "site":             "Site nom long",
    "ca":               "CA",
    "marge":            "Marge",
    "pct_marge":        "%Marge",
    "ca_promo":         "CA Promo",
    "marge_promo":      "Marge Promo",
    "pct_marge_promo":  "%Marge Promo",
    "poids_promo":      "%CA Poids Promo",
    "qte":              "Qté Vente",
    "qte_n1":           "Qté Vente N-1",
    "casse_qte":        "Casse (Qté)",
}
# Colonnes minimales sans lesquelles on ne peut rien faire.
# "ca" et "marge" sont indispensables : base du % Poids CA et de la carte Marge en en-tête.
VENTES_REQUISES = ["article", "site", "ca", "marge", "ca_promo", "marge_promo", "qte", "poids_promo"]

# Colonnes attendues dans la liste prévisions.
# Triple axe : prev_qte (unités) + prev_val (CA) + prev_marge (marge). Code obligatoire.
PREV_ALIASES = {
    "code":    ["code article", "code", "codeart", "code art"],
    "libelle": ["libelle", "designation", "libelle article"],
    "prev_qte": ["prevision qte", "prevision quantite", "prev qte", "prevision de vente qte",
                 "prevision vente qte", "prevision de vente", "prevision vente",
                 "prevision", "prev vente", "qte prevue", "quantite prevue"],
    "prev_val": ["prevision ca", "prevision valeur", "prev ca", "prev valeur",
                 "prevision de vente ca", "prevision de vente valeur",
                 "prevision ca fcfa", "ca prevu", "valeur prevue"],
    "prev_marge": ["prevision marge", "marge prevue", "prev marge", "marge prevision",
                   "prevision de vente marge", "marge prev"],
}

# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _norm(s: str) -> str:
    """Normalise un intitulé de colonne pour matching souple."""
    s = str(s).strip().lower()
    s = (s.replace("é", "e").replace("è", "e").replace("ê", "e")
           .replace("à", "a").replace("â", "a").replace("ô", "o")
           .replace("û", "u").replace("î", "i").replace("ç", "c"))
    s = re.sub(r"\s+", " ", s)
    return s


def safe_sort(df: pd.DataFrame, by, ascending=True) -> pd.DataFrame:
    """Tri robuste (ne casse pas si la colonne est absente ou vide)."""
    cols = [c for c in ([by] if isinstance(by, str) else by) if c in df.columns]
    if not cols:
        return df
    return df.sort_values(by=cols, ascending=ascending, kind="mergesort").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Chargement
# --------------------------------------------------------------------------- #
def load_ventes(file, sheet: str | None = None) -> pd.DataFrame:
    """Charge l'extraction ventes (xlsx ou csv) avec fallback encodage."""
    name = getattr(file, "name", str(file)).lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        xls = pd.ExcelFile(file)
        sh = sheet or (SHEET_DEFAULT if SHEET_DEFAULT in xls.sheet_names else xls.sheet_names[0])
        return pd.read_excel(xls, sheet_name=sh)
    # CSV : essais séparateur + encodage
    raw = file.read() if hasattr(file, "read") else open(file, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        for sep in (";", ",", "\t"):
            try:
                df = pd.read_csv(io.BytesIO(raw), sep=sep, encoding=enc)
                if df.shape[1] > 3:
                    return df
            except Exception:
                continue
    raise ValueError("Impossible de lire le fichier ventes (séparateur/encodage).")


_DATE_RE = re.compile(r"\d{2}/\d{2}/\d{4}")


def _split_period_row(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Isole une éventuelle ligne 'période' en dernière position de la 1ère colonne
    (texte libre, ex. dates ou phrase de filtre) et la retire des données.
    """
    meta = {"periode_texte": None, "periode_debut": None, "periode_fin": None}
    if df.empty:
        return df, meta
    first_col = df.columns[0]
    last_idx = df[first_col].last_valid_index()
    if last_idx is None:
        return df, meta
    val = str(df.at[last_idx, first_col]).strip()
    # une vraie ligne de code commence par des chiffres (avec éventuellement un séparateur) :
    # si ce n'est pas le cas, on considère que c'est la ligne de période.
    if not re.match(r"^\d", val):
        meta["periode_texte"] = val
        dates = _DATE_RE.findall(val)
        if len(dates) >= 2:
            meta["periode_debut"], meta["periode_fin"] = dates[0], dates[-1]
        elif len(dates) == 1:
            meta["periode_debut"] = dates[0]
        df = df.drop(index=last_idx)
    return df.reset_index(drop=True), meta


def load_previsions(file) -> tuple[pd.DataFrame, dict, dict]:
    """Charge la liste prévision, isole la période (dernière cellule colonne A),
    et mappe les colonnes (code/libelle/prev_qte/prev_val/prev_marge).
    Retourne (df_mappé, found, meta_periode)."""
    name = getattr(file, "name", str(file)).lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(file)
    else:
        raw = file.read() if hasattr(file, "read") else open(file, "rb").read()
        df = None
        for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
            for sep in (";", ",", "\t"):
                try:
                    tmp = pd.read_csv(io.BytesIO(raw), sep=sep, encoding=enc)
                    if tmp.shape[1] >= 2:
                        df = tmp; break
                except Exception:
                    continue
            if df is not None:
                break
        if df is None:
            raise ValueError("Impossible de lire le fichier prévisions.")

    df, meta = _split_period_row(df)

    norm = {_norm(c): c for c in df.columns}
    resolved, found, used = {}, {}, set()
    for key, aliases in PREV_ALIASES.items():
        col = next((norm[a] for a in aliases if a in norm and norm[a] not in used), None)
        found[key] = col
        if col:
            used.add(_norm(col))
            resolved[key] = df[col]
    out = pd.DataFrame(resolved)
    if "code" in out.columns:
        out["code"] = out["code"].astype(str).str.strip()
    for k in ("prev_qte", "prev_val", "prev_marge"):
        if k in out.columns:
            out[k] = pd.to_numeric(out[k], errors="coerce")
    return out, found, meta


# --------------------------------------------------------------------------- #
# Nettoyage / grain article x réseau
# --------------------------------------------------------------------------- #
_CODE_RE = re.compile(r"^\s*(\d+)\s*-\s*(.+)$")
_PERIODE_RE = re.compile(r"apr[eè]s le\s*(\d{2}/\d{2}/\d{4}).*?avant le\s*(\d{2}/\d{2}/\d{4})",
                         re.IGNORECASE | re.DOTALL)


def extract_meta(df_raw: pd.DataFrame) -> dict:
    """Récupère la période depuis la ligne de filtres, si présente."""
    meta = {"periode_debut": None, "periode_fin": None, "filtre_brut": None}
    for col in df_raw.columns:
        s = df_raw[col].astype(str)
        hit = s[s.str.contains("Filtres appliqu", case=False, na=False)]
        if len(hit):
            txt = hit.iloc[0]
            meta["filtre_brut"] = txt
            m = _PERIODE_RE.search(txt)
            if m:
                meta["periode_debut"] = m.group(1)
                # 'avant le' = borne exclusive -> fin réelle = veille
                fin = pd.to_datetime(m.group(2), format="%d/%m/%Y") - pd.Timedelta(days=1)
                meta["periode_fin"] = fin.strftime("%d/%m/%Y")
            break
    return meta


def to_article_reseau(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Isole le grain Article x Total réseau, code parsé, casse NaN->0."""
    art_col, site_col, qte_col = VENTES_MAP["article"], VENTES_MAP["site"], VENTES_MAP["qte"]

    # nb de sites actifs par article (lignes site-level avec Qté > 0)
    site_rows = df_raw[(df_raw[site_col].astype(str) != "Total")
                       & df_raw[art_col].astype(str).str.match(_CODE_RE)].copy()
    if len(site_rows):
        site_rows["_c"] = site_rows[art_col].astype(str).str.extract(_CODE_RE)[0].str.strip()
        site_rows["_q"] = pd.to_numeric(site_rows[qte_col], errors="coerce").fillna(0)
        nb_sites = site_rows[site_rows["_q"] > 0].groupby("_c")[site_col].nunique()
    else:
        nb_sites = pd.Series(dtype=int)

    d = df_raw.copy()
    mask = (d[site_col].astype(str) == "Total") & d[art_col].astype(str).str.match(_CODE_RE)
    d = d[mask].copy()
    codes = d[art_col].astype(str).str.extract(_CODE_RE)
    d["Code"] = codes[0].str.strip()
    d["Libellé_src"] = codes[1].str.strip()
    d["Nb sites actifs"] = d["Code"].map(nb_sites).fillna(0).astype(int)
    if VENTES_MAP["casse_qte"] in d.columns:
        d[VENTES_MAP["casse_qte"]] = d[VENTES_MAP["casse_qte"]].fillna(0)
    return d.reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Moteur de contrôles
# --------------------------------------------------------------------------- #
def _chk(label, statut, valeur="", message=""):
    return {"contrôle": label, "statut": statut, "valeur": valeur, "message": message}


def run_controls(df_raw: pd.DataFrame,
                 df_art: pd.DataFrame,
                 df_prev: pd.DataFrame,
                 prev_found: dict,
                 meta: dict,
                 prev_meta: dict | None = None) -> tuple[list[dict], bool]:
    """
    Retourne (liste_de_contrôles, bloquant).
    statut ∈ {'ok','warn','err'} ; 'err' => bloquant (on n'exploite pas).
    """
    R = []
    V = VENTES_MAP

    # --- A. STRUCTURE ---------------------------------------------------------
    manquantes = [V[k] for k in VENTES_REQUISES if V[k] not in df_raw.columns]
    R.append(_chk("A1 · Colonnes ventes requises",
                  "err" if manquantes else "ok",
                  f"{len(VENTES_REQUISES)-len(manquantes)}/{len(VENTES_REQUISES)}",
                  ("Manquantes : " + ", ".join(manquantes)) if manquantes else "Toutes présentes"))

    R.append(_chk("A2 · Grain Article × réseau isolé",
                  "ok" if len(df_art) else "err",
                  f"{len(df_art)} articles",
                  f"{len(df_raw)} lignes brutes → {len(df_art)} articles nets (sous-totaux/footer retirés)"))

    if meta.get("periode_debut"):
        R.append(_chk("A3 · Période détectée", "ok",
                      f"{meta['periode_debut']} → {meta['periode_fin']}",
                      "Extraite du pied de page"))
    else:
        R.append(_chk("A3 · Période détectée", "warn", "—",
                      "Pied de page absent : à renseigner manuellement"))

    # --- B. CLÉ / JOINTURE ----------------------------------------------------
    if "Code" in df_art.columns:
        non_num = (~df_art["Code"].str.match(r"^\d+$")).sum()
        R.append(_chk("B1 · Format des codes ventes",
                      "ok" if non_num == 0 else "warn",
                      f"{non_num} non conformes",
                      "Tous numériques" if non_num == 0 else "Codes non numériques détectés"))
        dup = df_art["Code"].duplicated(keep=False).sum()
        R.append(_chk("B2 · Doublons de code (ventes)",
                      "ok" if dup == 0 else "err",
                      f"{dup}", "Aucun" if dup == 0 else "Codes en double au niveau réseau"))

    # Prévisions : code obligatoire + les 3 axes (Qté, Valeur, Marge) requis
    has_code = bool(prev_found.get("code"))
    axes = [a for a in ("prev_qte", "prev_val", "prev_marge") if prev_found.get(a)]
    lib = {"prev_qte": "Qté", "prev_val": "Valeur", "prev_marge": "Marge"}
    miss_axes = [lib[a] for a in ("prev_qte", "prev_val", "prev_marge") if a not in axes]
    miss_prev = (not has_code) or bool(miss_axes)
    R.append(_chk("B3 · Colonnes prévision (code + Qté + Valeur + Marge)",
                  "err" if miss_prev else "ok",
                  ("code manquant" if not has_code else " + ".join(lib[a] for a in axes)),
                  "Les 3 axes sont mappés" if not miss_prev
                  else "Manquant : " + ", ".join((["code"] if not has_code else []) + miss_axes)))

    if prev_meta:
        if prev_meta.get("periode_debut"):
            R.append(_chk("B0 · Période promo détectée (col. A)", "ok",
                          f"{prev_meta['periode_debut']}"
                          + (f" → {prev_meta['periode_fin']}" if prev_meta.get("periode_fin") else ""),
                          "Lue dans la dernière cellule de la colonne A"))
        else:
            R.append(_chk("B0 · Période promo détectée (col. A)", "warn", "—",
                          "Aucune date reconnue dans la dernière cellule de colonne A"))

    coverage_ok = (not miss_prev) and ("Code" in df_art.columns)
    if coverage_ok:
        pc = df_prev.copy()
        dup_prev = pc["code"].duplicated(keep=False).sum()
        R.append(_chk("B4 · Doublons de code (prévision)",
                      "ok" if dup_prev == 0 else "warn",
                      f"{dup_prev}", "Aucun" if dup_prev == 0 else "À dédoublonner (somme des prév ?)"))

        codes_v = set(df_art["Code"])
        codes_p = set(pc["code"])
        inter = codes_p & codes_v
        cov = len(inter) / len(codes_p) if codes_p else 0
        orphelins_prev = codes_p - codes_v          # planifiés, pas retrouvés en ventes
        promo_codes = set(df_art.loc[df_art[V["ca_promo"]].fillna(0) > 0, "Code"])
        promo_hors_prev = promo_codes - codes_p     # vendus en promo mais absents du plan
        R.append(_chk("B5 · Couverture jointure prévision→ventes",
                      "ok" if cov >= 0.95 else "warn" if cov >= 0.8 else "err",
                      f"{cov:.0%}",
                      f"{len(inter)}/{len(codes_p)} codes prévision retrouvés"))
        R.append(_chk("B6 · Prévisions orphelines (planifié non vendu)",
                      "ok" if not orphelins_prev else "warn",
                      f"{len(orphelins_prev)}",
                      "0" if not orphelins_prev else f"Ex. {list(orphelins_prev)[:5]}"))
        R.append(_chk("B7 · Promo hors périmètre prévision",
                      "ok" if not promo_hors_prev else "warn",
                      f"{len(promo_hors_prev)}",
                      "Articles vendus en promo sans ligne de prévision (à ajouter au plan ?)"))

    # --- C. VALEURS -----------------------------------------------------------
    if coverage_ok:
        for ax, nom in (("prev_qte", "Qté"), ("prev_val", "Valeur"), ("prev_marge", "Marge")):
            if ax in df_prev.columns:
                v = pd.to_numeric(df_prev[ax], errors="coerce")
                nan_v = int(v.isna().sum())
                neg_v = int((v <= 0).sum()) if ax != "prev_marge" else int((v == 0).sum())
                R.append(_chk(f"C1 · Prévision {nom} : vides / {'≤0' if ax!='prev_marge' else '=0'}",
                              "err" if nan_v else "warn" if neg_v else "ok",
                              f"{nan_v} vides · {neg_v}" + (" ≤0" if ax != "prev_marge" else " =0"),
                              "Valide" if not (nan_v or neg_v) else "À corriger avant calcul"))

    qte_neg = (pd.to_numeric(df_art[V["qte"]], errors="coerce") < 0).sum()
    ca_neg = (pd.to_numeric(df_art[V["ca_promo"]], errors="coerce") < 0).sum()
    R.append(_chk("C2 · Quantités / CA promo négatifs",
                  "warn" if (qte_neg or ca_neg) else "ok",
                  f"{qte_neg} qté · {ca_neg} CA",
                  "Aucun" if not (qte_neg or ca_neg) else "Valeurs négatives (retours ?) à vérifier"))

    if V["casse_qte"] in df_art.columns:
        R.append(_chk("C3 · Casse (NaN → 0)", "ok",
                      f"{int((df_art[V['casse_qte']]==0).sum())} à 0",
                      "NaN traités comme absence de casse"))

    if V["marge_promo"] in df_art.columns:
        promo = df_art[df_art[V["ca_promo"]].fillna(0) > 0]
        neg_marge = (promo[V["marge_promo"]].fillna(0) < 0).sum()
        out_pct = 0
        if V["pct_marge_promo"] in promo.columns:
            mp = pd.to_numeric(promo[V["pct_marge_promo"]], errors="coerce").dropna()
            out_pct = int(((mp < -1) | (mp > 1)).sum())
        R.append(_chk("C4 · Marges promo négatives (signal métier)",
                      "warn" if neg_marge else "ok",
                      f"{neg_marge}/{len(promo)}",
                      f"{neg_marge} articles détruisent de la marge en promo"
                      + (f" · {out_pct} taux extrêmes à vérifier" if out_pct else "")))

    bloquant = any(r["statut"] == "err" for r in R)
    return R, bloquant


# --------------------------------------------------------------------------- #
# ÉTAPE 2 — Jointure & périmètre (piloté par la prévision, double axe)
# --------------------------------------------------------------------------- #
def build_perimetre(df_art: pd.DataFrame, df_prev: pd.DataFrame) -> pd.DataFrame:
    """
    Périmètre = liste prévision (left). On raccroche les ventes réelles.
    Décisions validées : périmètre prévision-driven ; promo hors plan ignorée.
    Axes réalisés (définition à valider) :
      - Qté  : Qté Vente TOTALE période (pas de split promo/hors-promo en quantité dans la source)
      - Valeur : CA TOTAL période  (CA Promo conservé comme part 'santé promo')
    """
    V = VENTES_MAP
    keep = ["Code", "Libellé_src", V["rayon"], V["sous_famille"],
            V["qte"], V["ca"], V["marge"], V["pct_marge"],
            V["ca_promo"], V["marge_promo"], V["pct_marge_promo"],
            V["poids_promo"], V["casse_qte"], "Nb sites actifs"]
    keep = [c for c in keep if c in df_art.columns]
    ventes = df_art[keep].copy()

    perim = df_prev.merge(ventes, left_on="code", right_on="Code", how="left")

    # libellé : prévision prioritaire, sinon source ventes
    perim["Libellé"] = perim.get("libelle")
    if "Libellé_src" in perim.columns:
        perim["Libellé"] = perim["Libellé"].fillna(perim["Libellé_src"])

    # ventes réelles (0 si article planifié non vendu = orphelin)
    perim["Ventes_Qte"]   = pd.to_numeric(perim.get(V["qte"]),   errors="coerce").fillna(0)
    perim["Ventes_CA"]    = pd.to_numeric(perim.get(V["ca"]),    errors="coerce").fillna(0)
    perim["Ventes_Marge"] = pd.to_numeric(perim.get(V["marge"]), errors="coerce").fillna(0)
    perim["_vendu"] = perim["Code"].notna()          # False => orphelin (non trouvé en ventes)
    perim["Orphelin"] = ~perim["_vendu"]
    return perim


def controls_jointure(perim: pd.DataFrame, df_prev: pd.DataFrame) -> tuple[list[dict], bool]:
    """Contrôles spécifiques à la jointure avant tout calcul de KPI."""
    R = []
    n = len(perim)
    orph = int(perim["Orphelin"].sum())
    vendus = n - orph
    R.append(_chk("J1 · Articles au périmètre (= prévision)",
                  "ok" if n else "err", f"{n}", "Périmètre piloté par la liste prévision"))
    R.append(_chk("J2 · Articles raccrochés aux ventes",
                  "ok" if vendus else "err", f"{vendus}/{n}",
                  f"{orph} orphelin(s) = planifié(s) non vendu(s) → réalisation 0%"))
    dup = int(perim["code"].duplicated(keep=False).sum())
    R.append(_chk("J3 · Unicité après jointure",
                  "ok" if dup == 0 else "err", f"{dup} doublon(s)",
                  "1 ligne par code" if dup == 0 else "La jointure a dupliqué des lignes (code ventes en double ?)"))

    # cohérence volumétrie par axe
    for ax, real, nom in (("prev_qte", "Ventes_Qte", "Qté"), ("prev_val", "Ventes_CA", "Valeur")):
        if ax in perim.columns:
            sp = pd.to_numeric(perim[ax], errors="coerce").sum()
            sr = perim[real].sum()
            ratio = (sr / sp) if sp else 0
            R.append(_chk(f"J4 · Réalisation globale {nom} (contrôle)",
                          "ok" if 0.3 <= ratio <= 3 else "warn",
                          f"{ratio:.0%}",
                          f"Σ réel {sr:,.0f} / Σ prév {sp:,.0f}".replace(",", " ")))
    bloquant = any(r["statut"] == "err" for r in R)
    return R, bloquant


# --------------------------------------------------------------------------- #
# ÉTAPE 3 — Calcul des KPI + statut RAG (double axe)
# --------------------------------------------------------------------------- #
# Seuils RAG (esprit Tesco) — modifiables
RAG_SEUIL_BAS   = 0.70   # < -> SOUS-PERF
RAG_SEUIL_CIBLE = 0.90   # [0.90 ; 1.15] -> CIBLE
RAG_SEUIL_HAUT  = 1.15   # > -> SUR-PERF


def _statut(ratio: pd.Series, prev: pd.Series) -> pd.Series:
    r = pd.to_numeric(ratio, errors="coerce")
    valide = pd.to_numeric(prev, errors="coerce").fillna(0) > 0
    lab = pd.Series("", index=r.index, dtype=object)
    lab[valide & (r < RAG_SEUIL_BAS)] = "SOUS-PERF"
    lab[valide & (r >= RAG_SEUIL_BAS) & (r < RAG_SEUIL_CIBLE)] = "SOUS TENDANCE"
    lab[valide & (r >= RAG_SEUIL_CIBLE) & (r <= RAG_SEUIL_HAUT)] = "CIBLE"
    lab[valide & (r > RAG_SEUIL_HAUT)] = "SUR-PERF"
    lab[~valide] = "À prévoir"
    return lab


def compute_kpi(perim: pd.DataFrame) -> pd.DataFrame:
    """
    Enrichit le périmètre avec écarts, % réalisation et statut RAG.
    Axe Qté  : Ventes_Qte / prev_qte
    Axe Valeur (base TOTAL validée) : Ventes_CA / prev_val
    Statut commercial de référence = axe Qté (volume), l'axe valeur en second.
    """
    d = perim.copy()
    V = VENTES_MAP

    def axe(prev_col, real_col, suf):
        if prev_col in d.columns:
            prev = pd.to_numeric(d[prev_col], errors="coerce")
            real = pd.to_numeric(d[real_col], errors="coerce").fillna(0)
            d[f"Ecart_{suf}"] = real - prev
            d[f"Real_{suf}"] = np.where(prev > 0, real / prev, np.nan)
            d[f"Statut_{suf}"] = _statut(d[f"Real_{suf}"], prev)

    axe("prev_qte", "Ventes_Qte", "Qte")
    axe("prev_val", "Ventes_CA",  "Val")

    # statut commercial de référence = Qté si dispo, sinon valeur
    if "Statut_Qte" in d.columns:
        d["Statut"] = d["Statut_Qte"]
    elif "Statut_Val" in d.columns:
        d["Statut"] = d["Statut_Val"]
    else:
        d["Statut"] = ""

    # alerte (marge promo négative prioritaire)
    mp = pd.to_numeric(d.get(V["marge_promo"]), errors="coerce").fillna(0)
    ref = d.get("Real_Qte", d.get("Real_Val"))
    ref = pd.to_numeric(ref, errors="coerce")
    d["Alerte"] = np.select(
        [mp < 0, ref < RAG_SEUIL_BAS, ref > RAG_SEUIL_HAUT],
        ["Marge nég.", "Surstock / casse", "Réassort"], default="")
    return d


def controls_kpi(kpi: pd.DataFrame) -> tuple[list[dict], bool]:
    """Contrôles post-calcul avant restitution."""
    R = []
    ref_axis = "Qte" if "Statut_Qte" in kpi.columns else "Val"
    st_col = f"Statut_{ref_axis}"

    calc = (kpi[st_col] != "").sum() if st_col in kpi.columns else 0
    a_prevoir = (kpi[st_col] == "À prévoir").sum() if st_col in kpi.columns else 0
    R.append(_chk("K1 · Statuts attribués",
                  "ok" if calc == len(kpi) else "warn",
                  f"{calc}/{len(kpi)}",
                  "Chaque ligne a un statut" if calc == len(kpi) else "Lignes sans statut"))
    R.append(_chk("K2 · Prévisions non calculables (≤0/vide)",
                  "ok" if a_prevoir == 0 else "warn",
                  f"{a_prevoir}", "Aucune" if a_prevoir == 0 else "Prévision absente/≤0"))

    if st_col in kpi.columns:
        vc = kpi[st_col].value_counts()
        dist = " · ".join(f"{k}:{int(vc.get(k,0))}"
                          for k in ["SOUS-PERF", "SOUS TENDANCE", "CIBLE", "SUR-PERF"])
        R.append(_chk(f"K3 · Répartition RAG ({'Qté' if ref_axis=='Qte' else 'Valeur'})",
                      "ok", dist, "Distribution des statuts commerciaux"))

    if "Alerte" in kpi.columns:
        nmarge = int((kpi["Alerte"] == "Marge nég.").sum())
        R.append(_chk("K4 · Alertes marge négative",
                      "warn" if nmarge else "ok", f"{nmarge}",
                      "Articles en destruction de marge (priorité COPIL)" if nmarge else "Aucune"))

    bloquant = any(r["statut"] == "err" for r in R)
    return R, bloquant


# --------------------------------------------------------------------------- #
# ÉTAPE 4 — Table scorecard finale + Poids CA + classement magasin
# --------------------------------------------------------------------------- #
def build_scorecard(kpi: pd.DataFrame, df_art: pd.DataFrame) -> pd.DataFrame:
    """
    Table finale demandée :
    Rang | Sous Famille | Code Article | Libellé | Qté prév. | CA prév. | Marge prév.
    | Qté réal. | CA réal. | Marge réal. | Poids CA | % Atteinte
    - Rang : CA réel décroissant
    - Poids CA : CA réel article ÷ CA TOTAL complet de l'extraction (tous articles,
      promo + hors promo, tous départements/rayons présents) — PAS seulement le périmètre promo.
    - % Atteinte : CA réal. ÷ CA prév. (axe de référence validé), avec statut RAG.
    """
    V = VENTES_MAP
    ca_total_extraction = pd.to_numeric(df_art[V["ca"]], errors="coerce").fillna(0).sum()

    d = kpi[kpi["_vendu"] | kpi["Orphelin"]].copy()  # tout le périmètre (orphelins = 0)
    sf_col = V["sous_famille"]
    d["Sous Famille"] = d.get(sf_col, "").astype(str).str.split(" - ", n=1).str[-1].str.strip()

    out = pd.DataFrame({
        "Sous Famille": d["Sous Famille"],
        "Code Article": d["code"],
        "Libellé": d["Libellé"],
        "Qté prév.": pd.to_numeric(d.get("prev_qte"), errors="coerce"),
        "CA prév.": pd.to_numeric(d.get("prev_val"), errors="coerce"),
        "Marge prév.": pd.to_numeric(d.get("prev_marge"), errors="coerce"),
        "Qté réal.": d["Ventes_Qte"],
        "CA réal.": d["Ventes_CA"],
        "Marge réal.": d["Ventes_Marge"],
        "Poids CA": (d["Ventes_CA"] / ca_total_extraction * 100) if ca_total_extraction else 0,
        "% Atteinte": d.get("Real_Val"),      # ratio CA réal / CA prév
        "Statut": d.get("Statut_Val"),
        "Alerte": d.get("Alerte"),
    })
    out = out.sort_values("CA réal.", ascending=False).reset_index(drop=True)
    out.insert(0, "Rang", range(1, len(out) + 1))
    return out


RAG_ICON = {
    "SOUS-PERF":     "🔴",
    "SOUS TENDANCE": "🟠",
    "CIBLE":         "🟢",
    "SUR-PERF":      "🔵",
    "À prévoir":     "⚪",
}


def store_ranking(df_raw: pd.DataFrame, codes_perimetre: list[str]) -> pd.DataFrame:
    """
    Classement magasin — sur tout le périmètre promo suivi (codes de la prévision).
    Valeurs réelles agrégées par site (pas de % réalisation : pas de prévision par magasin).
    """
    V = VENTES_MAP
    art_col, site_col = V["article"], V["site"]
    d = df_raw[(df_raw[site_col].astype(str) != "Total")
              & df_raw[art_col].astype(str).str.match(_CODE_RE)].copy()
    d["_code"] = d[art_col].astype(str).str.extract(_CODE_RE)[0].str.strip()
    d = d[d["_code"].isin(set(codes_perimetre))]

    agg = d.groupby(site_col).agg(
        CA=(V["ca"], "sum"), Marge=(V["marge"], "sum"), Qté=(V["qte"], "sum")
    ).reset_index().rename(columns={site_col: "Magasin"})
    total_ca = agg["CA"].sum()
    agg["Part réseau"] = (agg["CA"] / total_ca * 100) if total_ca else 0
    agg = agg.sort_values("CA", ascending=False).reset_index(drop=True)
    agg.insert(0, "Rang", range(1, len(agg) + 1))
    return agg
