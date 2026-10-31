"""SmartBuyer Hub · Flash quotidien.

Page de pilotage (usage personnel) : on charge les 3 extractions Power BI et le masque
PowerPoint CFAO, on contrôle, on lit, puis on télécharge le PPT pour les secteurs
et l'Excel de détail (valeurs uniquement).
"""
from __future__ import annotations

import datetime as dt
import pathlib
import shutil
import subprocess
import sys
import tempfile

import pandas as pd
import streamlit as st

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from modules.flash_quotidien import config as C  # noqa: E402
from modules.flash_quotidien import engine as E  # noqa: E402
from modules.flash_quotidien.excel_report import build_excel  # noqa: E402
from modules.flash_quotidien.ppt_report import build_ppt  # noqa: E402

NAVY, ORANGE, GREY, RED, TEAL, INK, MUTE = "#24235C", "#F09F39", "#F3F4F8", "#E62244", "#08868D", "#363C5A", "#6B7090"
MASQUE_DEFAUT = ROOT / "modules" / "flash_quotidien" / "assets" / "masque_cfao_consumer.pptx"  # non versionné (dépôt public)

st.set_page_config(page_title="Flash quotidien", page_icon="⚡", layout="wide")

st.markdown(f"""
<style>
.block-container {{padding-top: 1.6rem; max-width: 1400px;}}
.hero {{background:{NAVY}; color:#fff; border-radius:14px; padding:22px 28px; margin-bottom:14px;}}
.hero h1 {{color:#fff; font-size:1.7rem; margin:0; font-weight:300; letter-spacing:.02em;}}
.hero h1 b {{font-weight:800;}}
.hero p {{color:#D6D7E6; margin:.3rem 0 0 0;}}
.kpi {{background:{GREY}; border-radius:12px; padding:14px 18px; height:100%;}}
.kpi .l {{color:{MUTE}; font-size:.85rem;}}
.kpi .v {{color:{NAVY}; font-size:1.9rem; font-weight:800; line-height:1.2;}}
.kpi .d {{font-weight:700; font-size:.95rem;}}
.pos {{color:{TEAL};}} .neg {{color:{RED};}}
.card {{border:1px solid #DADDE6; border-radius:12px; padding:14px 18px; margin-bottom:10px; background:#fff;}}
.card h4 {{margin:0 0 6px 0; padding:0; color:{NAVY}; font-size:1.05rem;}}
.tag {{display:inline-block; background:{ORANGE}; color:#fff; border-radius:6px; padding:1px 8px; font-size:.75rem; font-weight:700;}}
.msg {{border-left:4px solid {ORANGE}; padding:6px 12px; margin:6px 0; background:{GREY}; border-radius:0 8px 8px 0; color:{INK};}}
.ext h5 {{margin:.2rem 0; color:{NAVY};}}
.cols span {{display:inline-block; background:{GREY}; border:1px solid #DADDE6; border-radius:6px; padding:1px 7px; margin:2px; font-size:.78rem; color:{INK};}}
</style>""", unsafe_allow_html=True)


def tone(v):
    return "pos" if v is not None and v >= 0 else "neg"


def kpi_card(col, label, value, delta, d_val):
    col.markdown(f"<div class='kpi'><div class='l'>{label}</div><div class='v'>{value}</div>"
                 f"<div class='d {tone(d_val)}'>{delta}</div></div>", unsafe_allow_html=True)


def guide_extraction(key: str, expanded: bool = False):
    spec = E.EXTRACTIONS[key]
    with st.expander(spec["titre"], expanded=expanded):
        st.markdown("\n".join(f"- {c}" for c in spec["consigne"]))
        st.markdown("**Colonnes attendues**")
        st.markdown("<div class='cols'>" + "".join(f"<span>{c}</span>" for c in spec["colonnes"]) + "</div>",
                    unsafe_allow_html=True)
        st.caption(f"Pied de page attendu : {spec['pied']}  ·  Nom conseillé : {spec['fichier']}")
        st.warning(spec["pieges"], icon="⚠️")


# ------------------------------------------------------------------ barre latérale
with st.sidebar:
    st.markdown(f"<h3 style='color:{NAVY};margin-bottom:0'>⚡ Flash quotidien</h3>"
                f"<span style='color:{MUTE}'>Category management, pricing et données</span>", unsafe_allow_html=True)
    st.divider()
    st.markdown("**Import fichiers**")
    date_flash = st.date_input("Ventes du", value=dt.date.today() - dt.timedelta(days=1), format="DD/MM/YYYY")
    f_hier = st.file_uploader("1. Hier · rayon × site", type="xlsx", key="hier")
    f_sem = st.file_uploader("2. Cette semaine · rayon × site", type="xlsx", key="sem")
    f_site = st.file_uploader("3. Hier · par site (débit, panier)", type="xlsx", key="site")
    f_masque = st.file_uploader("Masque PowerPoint CFAO", type="pptx", key="masque",
                                help="Le modèle 16-9 CFAO Consumer 2025. Seule la couverture est conservée, "
                                     "les autres slides sont construites sur la disposition « Titre seul ».")
    if f_masque is None and MASQUE_DEFAUT.exists():
        st.caption(f"Masque par défaut : {MASQUE_DEFAUT.name}")
    st.divider()
    st.markdown("**Extractions attendues**")
    for k in ("hier", "semaine", "sites"):
        guide_extraction(k)

# ------------------------------------------------------------------ page d'accueil
if not (f_hier and f_sem and f_site):
    st.markdown("<div class='hero'><h1>FLASH <b>QUOTIDIEN</b></h1>"
                "<p>Trois extractions Power BI, un masque PowerPoint : contrôle, lecture, puis PPT pour les secteurs "
                "et Excel de détail.</p></div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Les 3 extractions à préparer chaque matin")
        st.caption("Power BI, page des ventes CA, affichage HT. Le N-1 est le même jour de semaine de l'an dernier.")
        for k in ("hier", "semaine", "sites"):
            guide_extraction(k, expanded=True)
    with c2:
        st.markdown("#### Ce que produit la page")
        st.markdown(f"""
<div class='card'><h4>Contrôles automatiques</h4>Bonne vue, bonne période, HT, totaux cohérents entre exports 1 et 3,
sites manquants ou inconnus. Un contrôle en rouge bloque la génération.</div>
<div class='card'><h4>PowerPoint pour les secteurs (6 slides)</h4>Couverture · L'essentiel · Les secteurs · Où agir ·
Retours attendus (réponse le lendemain avant {C.HEURE_REPONSE}, le lundi pour un flash du vendredi ou du samedi) ·
Comment lire ce flash.</div>
<div class='card'><h4>Excel de détail</h4>Valeurs uniquement, sans formules : réseau, formats, zones, magasins,
convertis, alertes, un onglet par secteur et les retours attendus.</div>
<div class='card'><h4>Seuils d'alerte</h4>CA vs N-1 &lt; {C.SEUILS['ca_vs_n1']:.0%} · marge vs N-1 &lt;
{C.SEUILS['marge_vs_n1']:.0%} · taux &lt; {C.SEUILS['taux_pts']:.0f} pts · rayon × site avec CA N-1 ≥
{C.SEUILS['ca_n1_min']:,} FCFA. Réglables dans <code>config.py</code>.</div>""".replace(f"{C.SEUILS['ca_n1_min']:,}", f"{C.SEUILS['ca_n1_min']:,}".replace(",", " ")),
                    unsafe_allow_html=True)
    st.info("Chargez les 3 extractions dans la barre latérale pour lancer le flash.")
    st.stop()


# ------------------------------------------------------------------ calcul
@st.cache_data(show_spinner="Lecture des extractions…")
def run(b_hier, n_hier, b_sem, n_sem, b_site, n_site, d):
    import io
    h = E.read_export(io.BytesIO(b_hier), n_hier)
    s = E.read_export(io.BytesIO(b_sem), n_sem)
    t = E.read_export(io.BytesIO(b_site), n_site)
    checks = E.check_exports(h, s, t)
    R = None if E.bloquant(checks) else E.compute(h, s, t, d)
    return checks, R


checks, R = run(f_hier.getvalue(), f_hier.name, f_sem.getvalue(), f_sem.name, f_site.getvalue(), f_site.name, date_flash)

icon = {"ok": "✅", "alerte": "⚠️", "erreur": "⛔"}
nb_pb = sum(c.statut != "ok" for c in checks)
with st.expander(f"Contrôle des extractions · {'tout est conforme' if not nb_pb else f'{nb_pb} point(s) à vérifier'}",
                 expanded=bool(nb_pb)):
    for c in checks:
        st.markdown(f"{icon.get(c.statut, '•')} {c.message}")
if R is None:
    st.error("Génération bloquée : corrigez l'extraction signalée en rouge (voir les consignes dans la barre latérale).")
    st.stop()

# ------------------------------------------------------------------ en-tête
r = R["reseau"]
st.markdown(f"<div class='hero'><h1>FLASH <b>COMMERCIAL</b></h1>"
            f"<p>Ventes du {R['date'].strftime('%d/%m/%Y')} · {R['lecture']}</p></div>", unsafe_allow_html=True)
k1, k2, k3, k4, k5 = st.columns(5)
kpi_card(k1, "Chiffre d'affaires", E.mfcfa(r["ca"]), E.pct(r["vs_n1"]) + " vs N-1", r["vs_n1"])
kpi_card(k2, "Marge", E.mfcfa(r["marge"]), f"{E.pct(r['marge_vs_n1'])} · {E.mfcfa(r['dmarge'], sign=True)}", r["marge_vs_n1"])
kpi_card(k3, "Taux de marge", E.pct(r["taux"], sign=False), E.pts(r["dpts"]), r["dpts"])
kpi_card(k4, "Fréquentation", E.pct(r.get("debit_vs_n1")), "panier " + E.pct(r.get("panier_vs_n1")), r.get("debit_vs_n1"))
sem = R["reseau_sem"]
kpi_card(k5, "Semaine vs budget", E.pct(sem.get("vs_budget")), E.pct(sem["vs_n1"]) + " vs N-1", sem.get("vs_budget"))

st.markdown("")
for m in R["messages"]:
    st.markdown(f"<div class='msg'>{m}</div>", unsafe_allow_html=True)


def kpi_df(items: dict | list, label: str, trafic=False) -> pd.DataFrame:
    rows = []
    it = items.items() if isinstance(items, dict) else [(x["nom"], x) for x in items]
    for n, k in it:
        row = {label: n, "CA (M)": k["ca"] / 1e6, "CA vs N-1": k["vs_n1"] * 100, "Marge (M)": k["marge"] / 1e6,
               "Marge vs N-1": k["marge_vs_n1"] * 100, "Écart marge (M)": k["dmarge"] / 1e6,
               "Taux": k["taux"] * 100, "Écart taux (pt)": k["dpts"],
               "Effet volume (M)": k["effet_volume"] / 1e6, "Effet taux (M)": k["effet_taux"] / 1e6}
        if trafic:
            row["Débit vs N-1"] = (k.get("debit_vs_n1") or 0) * 100
            row["Panier vs N-1"] = (k.get("panier_vs_n1") or 0) * 100
        rows.append(row)
    return pd.DataFrame(rows)


CFG = {
    "CA (M)": st.column_config.NumberColumn(format="%.1f"),
    "Marge (M)": st.column_config.NumberColumn(format="%.1f"),
    "CA vs N-1": st.column_config.NumberColumn(format="%+.1f %%"),
    "Marge vs N-1": st.column_config.NumberColumn(format="%+.1f %%"),
    "Écart marge (M)": st.column_config.NumberColumn(format="%+.2f"),
    "Taux": st.column_config.NumberColumn(format="%.1f %%"),
    "Écart taux (pt)": st.column_config.NumberColumn(format="%+.1f"),
    "Effet volume (M)": st.column_config.NumberColumn(format="%+.2f"),
    "Effet taux (M)": st.column_config.NumberColumn(format="%+.2f"),
    "Débit vs N-1": st.column_config.NumberColumn(format="%+.1f %%"),
    "Panier vs N-1": st.column_config.NumberColumn(format="%+.1f %%"),
}


def show(df):
    st.dataframe(df, hide_index=True, width="stretch", column_config=CFG)


# ------------------------------------------------------------------ onglets
t_sec, t_fmt, t_ou, t_ret, t_ppt = st.tabs(["Secteurs", "Formats et zones", "Où agir", "Retours attendus", "Aperçu PPT"])

with t_sec:
    cols = st.columns(len(R["secteurs"]))
    for col, (sec, k) in zip(cols, R["secteurs"].items()):
        rayons = R["rayons"].get(sec, [])[:3]
        lst = "".join(f"<li>{x['nom']} : marge {E.kfcfa(x['dmarge'])}, taux {E.pts(x['dpts'])}</li>" for x in rayons)
        col.markdown(f"""<div class='card'><span class='tag'>{sec}</span>
<div class='kpi' style='background:none;padding:6px 0'><div class='l'>Marge vs N-1</div>
<div class='v {tone(k['marge_vs_n1'])}'>{E.pct(k['marge_vs_n1'])}</div>
<div class='l'>soit {E.mfcfa(k['dmarge'], sign=True)} FCFA · CA {E.pct(k['vs_n1'])} · taux {E.pct(k['taux'], sign=False)}
({E.pts(k['dpts'])}) · semaine vs budget {E.pct(k['sem'].get('vs_budget'))}</div></div>
<b>À regarder</b><ul style='margin:4px 0 0 0'>{lst or '<li>Rien hors seuil</li>'}</ul></div>""", unsafe_allow_html=True)
    show(kpi_df(R["secteurs"], "Secteur"))
    st.markdown("**Secteur × format** (écarts hors seuil)")
    rows = [{"Secteur": s, "Format": f, "Hors seuil": v["texte"]}
            for s, d in R["croise"].items() for f, v in d.items() if v["texte"]]
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")

with t_fmt:
    show(kpi_df(R["formats"], "Format", trafic=True))
    show(kpi_df(R["zones"], "Zone de pricing", trafic=True))
    st.markdown("**Magasins**")
    df = kpi_df(R["sites"], "Magasin", trafic=True)
    df.insert(1, "Alerte", ["⚠️" if x["alerte"] else "" for x in R["sites"]])
    show(df)

with t_ou:
    a, b = st.columns([3, 2])
    with a:
        st.markdown(f"**Alertes rayon × magasin** · top {C.TOP_ALERTES} par perte de marge")
        al = pd.DataFrame([{
            "Rayon": E.short_rayon(x["rayon"]), "Magasin": E.short_site(x["site"]), "Secteur": x["secteur"],
            "CA vs N-1": x["vs_n1"] * 100, "Écart marge (M)": x["dmarge"] / 1e6, "Écart taux (pt)": x["dpts"],
            "Diagnostic": x["diagnostic"], "Qui": x["qui"], "Aussi sur la semaine": "oui" if x.get("persistant") else ""}
            for x in R["alertes"]])
        show(al)
    with b:
        st.markdown("**Magasins convertis**")
        for x in R["convertis"]:
            st.markdown(f"""<div class='card'><h4>{x['nom']} <span style='color:{MUTE};font-weight:400;font-size:.85rem'>
· converti le {x['converti_le'].strftime('%d/%m/%Y') if x.get('converti_le') else 'n.d.'}</span></h4>
Marge <b class='{tone(x['marge_vs_n1'])}'>{E.pct(x['marge_vs_n1'])}</b> · Débit <b class='{tone(x.get('debit_vs_n1'))}'>
{E.pct(x.get('debit_vs_n1'))}</b> · Panier <b class='{tone(x.get('panier_vs_n1'))}'>{E.pct(x.get('panier_vs_n1'))}</b>
· Budget sem. <b class='{tone(x['sem'].get('vs_budget'))}'>{E.pct(x['sem'].get('vs_budget'))}</b></div>""",
                        unsafe_allow_html=True)

with t_ret:
    jr = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"][R["date_reponse"].weekday()]
    st.markdown(f"**Réponse attendue le {jr} {R['date_reponse'].strftime('%d/%m')} avant {C.HEURE_REPONSE}**")
    st.caption("Vous pouvez reformuler une question avant de générer le PPT.")
    ed = st.data_editor(pd.DataFrame(R["retours"])[["sujet", "question", "qui", "type"]],
                        hide_index=True, width="stretch", num_rows="dynamic",
                        column_config={"sujet": "Sujet", "question": "Question", "qui": "Qui répond",
                                       "type": st.column_config.SelectboxColumn(
                                           "Type", options=["Réponse attendue", "Pour information"])})
    R = {**R, "retours": ed.fillna("").to_dict("records")}

# ------------------------------------------------------------------ génération
masque = f_masque.getvalue() if f_masque else (MASQUE_DEFAUT.read_bytes() if MASQUE_DEFAUT.exists() else None)
ppt_bytes = build_ppt(R, masque) if masque else None
xls_bytes = build_excel(R)
stamp = R["date"].isoformat()


@st.cache_data(show_spinner="Rendu de l'aperçu…")
def render_preview(b: bytes) -> list[bytes]:
    """Images des slides si LibreOffice est installé sur le poste, sinon liste vide."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice or not shutil.which("pdftoppm"):
        return []
    with tempfile.TemporaryDirectory() as tmp:
        p = pathlib.Path(tmp) / "flash.pptx"; p.write_bytes(b)
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(p)],
                       capture_output=True, timeout=120)
        pdf = p.with_suffix(".pdf")
        if not pdf.exists():
            return []
        subprocess.run(["pdftoppm", "-png", "-r", "70", str(pdf), str(pathlib.Path(tmp) / "s")], capture_output=True)
        return [f.read_bytes() for f in sorted(pathlib.Path(tmp).glob("s-*.png"))]


with t_ppt:
    if not ppt_bytes:
        st.warning("Chargez le masque PowerPoint CFAO dans la barre latérale pour générer le PPT.")
    else:
        imgs = render_preview(ppt_bytes)
        if imgs:
            c = st.columns(2)
            for i, im in enumerate(imgs):
                c[i % 2].image(im, width="stretch")
        else:
            st.info("Aperçu image indisponible sur ce poste (LibreOffice absent). Plan du PPT généré :")
            st.markdown(f"""1. **Couverture** · Flash commercial, ventes du {R['date'].strftime('%d/%m/%Y')}
2. **L'essentiel** · 4 chiffres clés, 3 messages, marge vs N-1 par secteur
3. **Les secteurs** · une carte par secteur et les rayons à regarder
4. **Où agir** · {len(R['sites_alerte'])} magasins en alerte, écarts hors seuil, magasins convertis
5. **Retours attendus** · {sum(x['type'] == 'Réponse attendue' for x in R['retours'])} questions
6. **Comment lire ce flash**""")

st.divider()
d1, d2, _ = st.columns([1, 1, 2])
d1.download_button("⬇️ PowerPoint pour les secteurs", ppt_bytes or b"", file_name=f"Flash_commercial_{stamp}.pptx",
                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                   disabled=ppt_bytes is None, type="primary", width="stretch")
d2.download_button("⬇️ Excel de détail", xls_bytes, file_name=f"Flash_quotidien_{stamp}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
