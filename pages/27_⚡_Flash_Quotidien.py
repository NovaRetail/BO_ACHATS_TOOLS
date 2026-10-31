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
JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]

LANDING_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
.lp, .lp * {{font-family:'Inter', system-ui, sans-serif; box-sizing:border-box;}}
.lp {{color:{INK};}}
.lp-hero {{position:relative; overflow:hidden; border-radius:18px; padding:30px 34px;
  background:linear-gradient(120deg,#1B1A4A 0%,{NAVY} 45%,#3B3A8F 100%); color:#fff; display:flex; gap:28px;
  align-items:stretch; box-shadow:0 12px 30px rgba(36,35,92,.22);}}
.lp-hero:before {{content:""; position:absolute; inset:0; opacity:.10;
  background-image:linear-gradient(#fff 1px,transparent 1px),linear-gradient(90deg,#fff 1px,transparent 1px);
  background-size:28px 28px; mask-image:linear-gradient(90deg,transparent 30%,#000 100%);}}
.lp-hero:after {{content:""; position:absolute; right:-80px; top:-80px; width:280px; height:280px; border-radius:50%;
  background:radial-gradient(circle,{ORANGE}55 0%,transparent 70%);}}
.lp-hl {{flex:1.4; position:relative; z-index:1;}}
.lp-eyebrow {{font-size:.72rem; letter-spacing:.18em; font-weight:700; color:{ORANGE}; text-transform:uppercase;}}
.lp-title {{font-size:2.3rem; font-weight:300; margin:.35rem 0 .2rem 0; letter-spacing:.01em; line-height:1.1;}}
.lp-title b {{font-weight:800;}}
.lp-sub {{color:#CFD0E6; font-size:.98rem; max-width:620px; line-height:1.5;}}
.lp-meta {{display:flex; gap:10px; margin-top:18px; flex-wrap:wrap;}}
.lp-chip {{background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.18); border-radius:999px;
  padding:6px 14px; font-size:.82rem; color:#fff; backdrop-filter:blur(4px);}}
.lp-chip b {{color:{ORANGE}; font-weight:700;}}
.lp-status {{flex:1; position:relative; z-index:1; background:rgba(255,255,255,.07); border:1px solid rgba(255,255,255,.16);
  border-radius:14px; padding:16px 18px; min-width:280px;}}
.lp-status h6 {{margin:0 0 4px 0; font-size:.75rem; letter-spacing:.12em; text-transform:uppercase; color:#B9BAD6; font-weight:600;}}
.lp-prog {{display:flex; align-items:baseline; gap:8px; margin-bottom:8px;}}
.lp-prog .n {{font-size:2rem; font-weight:800; color:#fff;}}
.lp-prog .t {{color:#B9BAD6; font-size:.85rem;}}
.lp-bar {{height:6px; border-radius:6px; background:rgba(255,255,255,.14); overflow:hidden; margin-bottom:12px;}}
.lp-bar span {{display:block; height:100%; background:linear-gradient(90deg,{ORANGE},#FFC27A); border-radius:6px;}}
.lp-row {{display:flex; align-items:center; gap:10px; padding:6px 0; font-size:.86rem; color:#E7E8F5;
  border-top:1px solid rgba(255,255,255,.08);}}
.lp-dot {{width:20px; height:20px; border-radius:50%; display:flex; align-items:center; justify-content:center;
  font-size:.7rem; font-weight:800; flex:none;}}
.lp-dot.ok {{background:{TEAL}; color:#fff;}} .lp-dot.wait {{border:1.5px dashed #8E8FB8; color:#8E8FB8;}}
.lp-row .r {{margin-left:auto; font-size:.74rem; color:#B9BAD6;}}

.lp-steps {{display:grid; grid-template-columns:repeat(4,1fr); gap:0; margin:22px 0 8px 0; position:relative;}}
.lp-step {{position:relative; padding:0 14px 0 0;}}
.lp-step:not(:last-child):after {{content:""; position:absolute; top:17px; left:44px; right:8px; height:2px;
  background:linear-gradient(90deg,{ORANGE},#E3E5EE);}}
.lp-step .num {{width:36px; height:36px; border-radius:10px; background:#fff; border:2px solid {ORANGE}; color:{NAVY};
  display:flex; align-items:center; justify-content:center; font-weight:800; position:relative; z-index:1;}}
.lp-step .lab {{font-weight:700; color:{NAVY}; margin-top:10px; font-size:.95rem;}}
.lp-step .txt {{color:{MUTE}; font-size:.8rem; line-height:1.45; margin-top:2px; padding-right:10px;}}

.lp-sec {{display:flex; align-items:center; gap:10px; margin:26px 0 12px 0;}}
.lp-sec h3 {{margin:0; padding:0; font-size:1.05rem; font-weight:700; color:{NAVY};}}
.lp-sec .line {{flex:1; height:1px; background:#E3E5EE;}}
.lp-sec .hint {{font-size:.78rem; color:{MUTE};}}

.lp-grid3 {{display:grid; grid-template-columns:repeat(3,1fr); gap:16px;}}
.lp-ext {{background:#fff; border:1px solid #E3E5EE; border-radius:16px; padding:18px 18px 16px; position:relative;
  box-shadow:0 1px 2px rgba(20,20,60,.04),0 8px 22px rgba(20,20,60,.05); display:flex; flex-direction:column;}}
.lp-ext.done {{border-color:{TEAL}66; box-shadow:0 0 0 3px {TEAL}14,0 8px 22px rgba(20,20,60,.05);}}
.lp-ext .top {{display:flex; align-items:center; gap:12px;}}
.lp-ext .ic {{width:42px; height:42px; border-radius:12px; display:flex; align-items:center; justify-content:center;
  font-size:1.05rem; font-weight:800; color:#fff; background:linear-gradient(135deg,{NAVY},#4B4AA8); flex:none;}}
.lp-ext .tt {{font-weight:700; color:{NAVY}; font-size:.98rem; line-height:1.25;}}
.lp-ext .st {{margin-left:auto; font-size:.7rem; font-weight:700; padding:3px 9px; border-radius:999px; white-space:nowrap;}}
.lp-ext .st.ok {{background:{TEAL}1A; color:{TEAL};}} .lp-ext .st.wait {{background:{GREY}; color:{MUTE};}}
.lp-badges {{display:flex; gap:6px; margin:12px 0 10px 0; flex-wrap:wrap;}}
.lp-b {{font-size:.72rem; font-weight:600; padding:3px 9px; border-radius:6px; background:{ORANGE}1F; color:#B4610A;}}
.lp-b.v {{background:#24235C14; color:{NAVY};}}
.lp-ext ol {{margin:0 0 10px 0; padding-left:18px; font-size:.82rem; line-height:1.5; color:{INK};}}
.lp-ext ol li {{margin-bottom:2px;}}
.lp-k {{font-size:.68rem; letter-spacing:.1em; text-transform:uppercase; color:{MUTE}; font-weight:700; margin:6px 0 5px 0;}}
.lp-cols {{display:flex; flex-wrap:wrap; gap:4px;}}
.lp-cols span {{font-size:.72rem; padding:2px 7px; border-radius:5px; background:{GREY}; border:1px solid #E3E5EE; color:{INK};}}
.lp-foot {{font-family:ui-monospace,SFMono-Regular,Menlo,monospace !important; font-size:.72rem; color:{NAVY};
  background:#F7F7FB; border:1px dashed #D5D7E4; border-radius:8px; padding:6px 9px; margin-top:4px;}}
.lp-trap {{margin-top:auto; padding-top:10px;}}
.lp-trap div {{display:flex; gap:8px; font-size:.78rem; line-height:1.45; color:#7A4A0E; background:#FFF6EA;
  border-left:3px solid {ORANGE}; border-radius:0 8px 8px 0; padding:8px 10px;}}

.lp-grid2 {{display:grid; grid-template-columns:1.25fr 1fr; gap:16px;}}
.lp-card {{background:#fff; border:1px solid #E3E5EE; border-radius:16px; padding:18px 20px;
  box-shadow:0 1px 2px rgba(20,20,60,.04),0 8px 22px rgba(20,20,60,.05);}}
.lp-card h4 {{margin:0 0 4px 0; padding:0; font-size:.98rem; font-weight:700; color:{NAVY};}}
.lp-card p {{margin:0 0 12px 0; font-size:.8rem; color:{MUTE};}}
.lp-slides {{display:grid; grid-template-columns:repeat(3,1fr); gap:10px;}}
.lp-sl {{aspect-ratio:16/9; border-radius:8px; overflow:hidden; border:1px solid #E3E5EE; background:#fff;
  display:flex; flex-direction:column; position:relative;}}
.lp-sl .bar {{height:24%; background:{NAVY}; display:flex; align-items:center; padding:0 8px; color:#fff;
  font-size:.62rem; font-weight:700; letter-spacing:.04em; text-transform:uppercase;}}
.lp-sl .body {{flex:1; padding:6px 8px; display:flex; gap:4px; align-items:flex-end;}}
.lp-sl .body i {{flex:1; background:#E3E5EE; border-radius:2px;}}
.lp-sl .body i.o {{background:{ORANGE};}} .lp-sl .body i.t {{background:{TEAL};}}
.lp-sl.cover {{background:linear-gradient(135deg,{NAVY},#3B3A8F);}}
.lp-sl.cover .ct {{color:#fff; font-size:.7rem; font-weight:300; padding:10px; margin-top:auto;}}
.lp-sl.cover .ct b {{font-weight:800;}}
.lp-sl .pg {{position:absolute; right:6px; bottom:4px; font-size:.55rem; color:{MUTE};}}
.lp-tabs {{display:flex; flex-wrap:wrap; gap:6px; margin-top:10px;}}
.lp-tabs span {{font-size:.74rem; padding:4px 10px; border-radius:6px 6px 0 0; background:{GREY};
  border:1px solid #E3E5EE; border-bottom:2px solid #1D6F42; color:{INK}; font-weight:600;}}
.lp-th {{display:grid; grid-template-columns:repeat(2,1fr); gap:10px; margin-top:14px;}}
.lp-t {{background:{GREY}; border-radius:12px; padding:10px 12px;}}
.lp-t .v {{font-size:1.35rem; font-weight:800; color:{RED};}}
.lp-t .l {{font-size:.74rem; color:{MUTE}; line-height:1.3;}}
.lp-cta {{margin-top:22px; display:flex; align-items:center; gap:14px; background:#fff; border:1px solid #E3E5EE;
  border-left:4px solid {ORANGE}; border-radius:12px; padding:14px 18px; font-size:.9rem; color:{INK};}}
.lp-cta b {{color:{NAVY};}}
@media (max-width: 1100px) {{ .lp-grid3, .lp-grid2 {{grid-template-columns:1fr;}} .lp-hero {{flex-direction:column;}}
  .lp-steps {{grid-template-columns:repeat(2,1fr); row-gap:18px;}} .lp-step:after {{display:none;}} }}
</style>"""


def _h(s: str) -> str:
    """HTML sur une ligne (évite que Markdown transforme l'indentation en bloc de code)."""
    return " ".join(line.strip() for line in s.splitlines())


def render_landing(files: dict, masque_ok: bool, d: dt.date):
    rep = E.date_reponse(d)
    n_ok = sum(1 for k in ("hier", "semaine", "sites") if files[k])
    pct_ok = int(n_ok / 3 * 100)

    def row(label, ok, right):
        return (f"<div class='lp-row'><div class='lp-dot {'ok' if ok else 'wait'}'>{'✓' if ok else ''}</div>"
                f"{label}<span class='r'>{right}</span></div>")

    status = (f"<div class='lp-status'><h6>Préparation du flash</h6>"
              f"<div class='lp-prog'><span class='n'>{n_ok}/3</span><span class='t'>extractions chargées</span></div>"
              f"<div class='lp-bar'><span style='width:{max(pct_ok, 3)}%'></span></div>"
              + row("Hier · rayon × site", files["hier"], "Export 1")
              + row("Cette semaine · rayon × site", files["semaine"], "Export 2")
              + row("Hier · par site", files["sites"], "Export 3")
              + row("Masque PowerPoint CFAO", masque_ok, "par défaut" if masque_ok and not files["masque"] else "")
              + "</div>")

    hero = f"""<div class='lp-hero'><div class='lp-hl'>
<div class='lp-eyebrow'>Pilotage · Quotidien</div>
<div class='lp-title'>FLASH <b>COMMERCIAL</b></div>
<div class='lp-sub'>Les ventes de la veille lues en entonnoir — réseau, secteurs, formats, magasins —
avec les écarts à traiter et les questions à poser aux secteurs. Trois extractions Power BI suffisent.</div>
<div class='lp-meta'>
<span class='lp-chip'>Ventes du <b>{JOURS[d.weekday()]} {d.strftime('%d/%m')}</b></span>
<span class='lp-chip'>Retours attendus <b>{JOURS[rep.weekday()]} {rep.strftime('%d/%m')} · {C.HEURE_REPONSE}</b></span>
<span class='lp-chip'>Comparaison <b>même jour N-1 · HT</b></span>
<span class='lp-chip'><b>{len(C.SITES)}</b> magasins · <b>{len(C.SECTEURS)}</b> secteurs</span>
</div></div>{status}</div>"""

    steps = """<div class='lp-steps'>
<div class='lp-step'><div class='num'>1</div><div class='lab'>Extraire</div>
<div class='txt'>3 exports Power BI au format .xlsx, disposition actuelle.</div></div>
<div class='lp-step'><div class='num'>2</div><div class='lab'>Contrôler</div>
<div class='txt'>Période, vue, HT et totaux vérifiés ; une erreur bloque la génération.</div></div>
<div class='lp-step'><div class='num'>3</div><div class='lab'>Lire</div>
<div class='txt'>L'essentiel, les secteurs, où agir et les retours à demander.</div></div>
<div class='lp-step'><div class='num'>4</div><div class='lab'>Diffuser</div>
<div class='txt'>PowerPoint pour les secteurs et Excel de détail en valeurs.</div></div>
</div>"""

    meta = {"hier": ("Hier", "Département › Rayon › Site", "hier"),
            "semaine": ("Cette semaine", "Département › Rayon › Site", "semaine"),
            "sites": ("Hier", "Site uniquement", "sites")}
    cards = []
    for i, k in enumerate(("hier", "semaine", "sites"), 1):
        spec = E.EXTRACTIONS[k]
        per, vue, fk = meta[k]
        ok = files[fk]
        titre = spec["titre"].split(". ", 1)[-1]
        cards.append(f"""<div class='lp-ext {'done' if ok else ''}'>
<div class='top'><div class='ic'>{i}</div><div class='tt'>{titre}</div>
<span class='st {'ok' if ok else 'wait'}'>{'Chargé' if ok else 'En attente'}</span></div>
<div class='lp-badges'><span class='lp-b'>Période · {per}</span><span class='lp-b v'>{vue}</span><span class='lp-b v'>HT</span></div>
<ol>{''.join(f'<li>{c}</li>' for c in spec['consigne'])}</ol>
<div class='lp-k'>Colonnes attendues</div>
<div class='lp-cols'>{''.join(f'<span>{c}</span>' for c in spec['colonnes'])}</div>
<div class='lp-k'>Pied de page à vérifier</div>
<div class='lp-foot'>{spec['pied']}</div>
<div class='lp-trap'><div>⚠️<span>{spec['pieges']}</span></div></div>
</div>""")

    slides = [("cover", "", ""), ("L'essentiel", "o t t i", "2"), ("Les secteurs", "t t i", "3"),
              ("Où agir", "i o i i", "4"), ("Retours attendus", "o i i", "5"), ("Comment lire", "i i i i", "6")]
    sl_html = ""
    heights = [70, 45, 85, 30, 60]
    for name, bars, pg in slides:
        if name == "cover":
            sl_html += "<div class='lp-sl cover'><div class='ct'>FLASH <b>COMMERCIAL</b></div></div>"
            continue
        bh = "".join(f"<i class='{c}' style='height:{heights[j % 5]}%'></i>" for j, c in enumerate(bars.split()))
        sl_html += f"<div class='lp-sl'><div class='bar'>{name}</div><div class='body'>{bh}</div><span class='pg'>{pg}</span></div>"

    sv = C.SEUILS
    seuils = "".join(f"<div class='lp-t'><div class='v'>{v}</div><div class='l'>{l}</div></div>" for v, l in [
        (E.pct(sv['ca_vs_n1'], 0), "Alerte si le CA vs N-1 passe sous ce seuil"),
        (E.pct(sv['marge_vs_n1'], 0), "Alerte si la marge vs N-1 passe sous ce seuil"),
        (E.pts(sv['taux_pts'], 0).replace("pt", "pts"), "Alerte si le taux de marge recule de plus"),
        ("≥ " + f"{sv['ca_n1_min'] / 1000:,.0f} k".replace(",", " "), "CA N-1 minimum d'un rayon × site suivi (FCFA)"),
    ])

    outputs = f"""<div class='lp-grid2'>
<div class='lp-card'><h4>PowerPoint pour les secteurs</h4>
<p>6 slides sur le masque CFAO · réponse le lendemain avant {C.HEURE_REPONSE} (lundi pour un flash du vendredi ou du samedi)</p>
<div class='lp-slides'>{sl_html}</div></div>
<div class='lp-card'><h4>Excel de détail &amp; seuils d'alerte</h4>
<p>Valeurs uniquement, sans formules · seuils réglables dans <code>config.py</code></p>
<div class='lp-tabs'><span>Flash</span><span>PGC</span><span>Non-alimentaire</span><span>Frais</span><span>Retours attendus</span></div>
<div class='lp-th'>{seuils}</div></div>
</div>"""

    cta = ("<div class='lp-cta'>⬅️<span>Chargez les <b>3 extractions</b> dans la barre latérale : "
           "le flash se calcule dès que les trois fichiers sont présents.</span></div>")

    st.markdown(LANDING_CSS, unsafe_allow_html=True)
    st.markdown(_h(f"<div class='lp'>{hero}{steps}"
                   f"<div class='lp-sec'><h3>Les 3 extractions à préparer chaque matin</h3><div class='line'></div>"
                   f"<span class='hint'>Power BI · page des ventes CA</span></div>"
                   f"<div class='lp-grid3'>{''.join(cards)}</div>"
                   f"<div class='lp-sec'><h3>Ce que vous obtenez</h3><div class='line'></div></div>"
                   f"{outputs}{cta}</div>"), unsafe_allow_html=True)


if not (f_hier and f_sem and f_site):
    render_landing({"hier": bool(f_hier), "semaine": bool(f_sem), "sites": bool(f_site), "masque": bool(f_masque)},
                   bool(f_masque) or MASQUE_DEFAUT.exists(), date_flash)
    st.stop()


# ------------------------------------------------------------------ styles des résultats (même charte que l'accueil)
APP_CSS = f"""
<style>
.bi, .bi * {{font-family:'Inter', system-ui, sans-serif; box-sizing:border-box;}}
.bi {{color:{INK};}}
/* contrôle des extractions */
.bi-check {{display:flex; align-items:center; gap:10px; flex-wrap:wrap; background:#fff; border:1px solid #E3E5EE;
  border-radius:12px; padding:10px 14px; margin-bottom:14px; font-size:.8rem;}}
.bi-check .h {{font-weight:700; color:{NAVY}; margin-right:4px;}}
.bi-check .p {{display:inline-flex; align-items:center; gap:6px; padding:4px 10px; border-radius:999px; font-weight:600;}}
.bi-check .p.ok {{background:{TEAL}14; color:{TEAL};}} .bi-check .p.alerte {{background:{ORANGE}1F; color:#B4610A;}}
.bi-check .p.erreur {{background:{RED}14; color:{RED};}}
/* cascade de marge dans le bandeau */
.wf {{display:flex; flex-direction:column; gap:7px; margin-top:4px;}}
.wf-r {{display:grid; grid-template-columns:92px 1fr 70px; align-items:center; gap:8px; font-size:.78rem; color:#E7E8F5;}}
.wf-t {{position:relative; height:16px; background:rgba(255,255,255,.06); border-radius:4px;}}
.wf-t span {{position:absolute; top:0; bottom:0; border-radius:4px;}}
.wf-r .v {{text-align:right; font-weight:700; color:#fff; white-space:nowrap;}}
.wf-n {{font-size:.68rem; color:#9C9DC4; margin-top:6px;}}
/* cartes KPI */
.kgrid {{display:grid; grid-template-columns:repeat(5,1fr); gap:14px; margin:18px 0 6px 0;}}
.kc {{background:#fff; border:1px solid #E3E5EE; border-radius:16px; padding:16px 16px 14px; position:relative; overflow:hidden;
  box-shadow:0 1px 2px rgba(20,20,60,.04),0 8px 22px rgba(20,20,60,.05);}}
.kc:before {{content:""; position:absolute; left:0; top:0; right:0; height:3px; background:var(--acc);}}
.kc .hd {{display:flex; align-items:center; gap:8px; font-size:.78rem; color:{MUTE}; font-weight:600;}}
.kc .ic {{width:26px; height:26px; border-radius:8px; display:flex; align-items:center; justify-content:center;
  font-size:.8rem; background:var(--tint); color:var(--acc); font-weight:800;}}
.kc .v {{font-size:1.75rem; font-weight:800; color:{NAVY}; margin:10px 0 6px 0; letter-spacing:-.01em; white-space:nowrap;}}
.kc .v.pos {{color:{TEAL};}} .kc .v.neg {{color:{RED};}}
.kc .d {{display:inline-flex; align-items:center; gap:4px; font-size:.76rem; font-weight:700; padding:3px 9px; border-radius:999px;}}
.d.pos {{background:{TEAL}17; color:{TEAL};}} .d.neg {{background:{RED}14; color:{RED};}}
.kc .s {{font-size:.74rem; color:{MUTE}; margin-top:8px;}}
.kc .cmp {{margin-top:10px; display:flex; flex-direction:column; gap:3px;}}
.kc .cmp div {{display:grid; grid-template-columns:28px 1fr; align-items:center; gap:6px; font-size:.64rem; color:{MUTE};}}
.kc .cmp i {{display:block; height:5px; border-radius:3px;}}
/* à retenir */
.ret {{display:grid; grid-template-columns:repeat(3,1fr); gap:14px; margin:12px 0 4px 0;}}
.ret div {{display:flex; gap:12px; align-items:flex-start; background:#fff; border:1px solid #E3E5EE; border-radius:14px;
  padding:14px 16px; font-size:.86rem; line-height:1.45;}}
.ret b.n {{flex:none; width:26px; height:26px; border-radius:50%; background:{ORANGE}; color:#fff; display:flex;
  align-items:center; justify-content:center; font-size:.8rem;}}
/* onglets en contrôle segmenté */
div[data-testid="stTabs"] [role="tablist"] {{gap:4px; background:{GREY}; padding:5px; border-radius:12px;
  width:fit-content; border:none; box-shadow:none;}}
div[data-testid="stTabs"] [data-testid="stTab"] {{height:36px; padding:0 18px; border-radius:9px; background:transparent;
  font-family:'Inter',sans-serif; font-weight:600; color:{MUTE}; border:none; margin:0;}}
div[data-testid="stTabs"] [data-testid="stTab"] p {{font-weight:600; font-size:.88rem;}}
div[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] {{background:#fff; color:{NAVY};
  box-shadow:0 1px 3px rgba(20,20,60,.14);}}
div[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] p {{color:{NAVY};}}
div[data-testid="stTabs"] [role="tablist"] > :not([data-testid="stTab"]) {{display:none;}}
div[data-testid="stTabs"] [role="tablist"] ~ * {{border:none;}}
div[data-testid="stTabs"] [data-testid="stTab"] div {{background:transparent !important;}}
div[data-testid="stTabs"] [data-testid="stTab"]:after, div[data-testid="stTabs"] [data-testid="stTab"]:before {{display:none;}}
/* titres de section */
.sec {{display:flex; align-items:center; gap:10px; margin:22px 0 10px 0;}}
.sec h3 {{margin:0; padding:0; font-size:1rem; font-weight:700; color:{NAVY};}}
.sec .line {{flex:1; height:1px; background:#E3E5EE;}} .sec .hint {{font-size:.76rem; color:{MUTE};}}
/* tableaux */
.tw {{background:#fff; border:1px solid #E3E5EE; border-radius:14px; overflow:auto;
  box-shadow:0 1px 2px rgba(20,20,60,.04),0 8px 22px rgba(20,20,60,.04);}}
table.bt {{width:100%; border-collapse:collapse; font-size:.8rem;}}
.bt th {{position:sticky; top:0; background:#F7F7FB; color:{MUTE}; font-weight:700; font-size:.66rem; letter-spacing:.06em;
  text-transform:uppercase; text-align:right; padding:10px 10px; border-bottom:1px solid #E3E5EE; white-space:nowrap;}}
.bt th.l, .bt td.l {{text-align:left;}}
.bt td {{padding:9px 10px; border-bottom:1px solid #F0F1F5; text-align:right; white-space:nowrap; color:{INK};}}
.bt tr:last-child td {{border-bottom:none;}}
.bt tr:hover td {{background:#FAFAFD;}}
.bt td.first {{font-weight:700; color:{NAVY};}}
.bt td .sub {{font-weight:500; color:{MUTE}; font-size:.72rem; margin-top:1px;}}
.bt tr.al td.first {{box-shadow:inset 3px 0 0 {RED};}}
.bt tr.tot td {{background:#F7F7FB; font-weight:700;}}
.bt .pos {{color:{TEAL}; font-weight:600;}} .bt .neg {{color:{RED}; font-weight:600;}}
.db {{display:flex; align-items:center; justify-content:flex-end; gap:8px;}}
.db .tr {{position:relative; width:52px; height:8px; background:{GREY}; border-radius:4px;}}
.db .tr:after {{content:""; position:absolute; left:50%; top:-2px; bottom:-2px; width:1px; background:#C9CBD8;}}
.db .tr span {{position:absolute; top:0; bottom:0; border-radius:4px;}}
.pill {{display:inline-block; font-size:.7rem; font-weight:700; padding:2px 8px; border-radius:6px; white-space:nowrap;}}
.pill.nav {{background:#24235C12; color:{NAVY};}} .pill.or {{background:{ORANGE}1F; color:#B4610A;}}
.pill.rd {{background:{RED}14; color:{RED};}} .pill.tl {{background:{TEAL}17; color:{TEAL};}}
/* cartes secteur */
.g3 {{display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px;}}
.sc {{background:#fff; border:1px solid #E3E5EE; border-radius:16px; padding:16px 18px;
  box-shadow:0 1px 2px rgba(20,20,60,.04),0 8px 22px rgba(20,20,60,.05);}}
.sc .top {{display:flex; align-items:center; justify-content:space-between;}}
.sc .nm {{font-weight:800; color:#fff; background:{NAVY}; border-radius:8px; padding:4px 10px; font-size:.78rem; letter-spacing:.04em;}}
.sc .big {{font-size:2rem; font-weight:800; margin:12px 0 0 0; letter-spacing:-.01em;}}
.sc .bl {{font-size:.74rem; color:{MUTE};}}
.sc .mg {{display:grid; grid-template-columns:repeat(3,1fr); gap:8px; margin:14px 0;}}
.sc .mg div {{background:{GREY}; border-radius:10px; padding:8px 10px;}}
.sc .mg .l {{font-size:.66rem; color:{MUTE}; text-transform:uppercase; letter-spacing:.06em; font-weight:700;}}
.sc .mg .v {{font-size:.95rem; font-weight:800; color:{NAVY}; margin-top:2px;}}
.sc .mg .x {{font-size:.72rem; font-weight:700;}}
.sc .k {{font-size:.66rem; letter-spacing:.1em; text-transform:uppercase; color:{MUTE}; font-weight:700; margin-bottom:6px;}}
.sc .ry {{display:flex; align-items:center; justify-content:space-between; gap:8px; padding:7px 0; border-top:1px solid #F0F1F5;
  font-size:.8rem;}}
.sc .ry b {{color:{NAVY};}}
.pos {{color:{TEAL};}} .neg {{color:{RED};}}
/* carte de chaleur secteur × format */
.hm {{display:grid; gap:6px;}}
.hm .c {{border-radius:10px; padding:10px 12px; min-height:62px; font-size:.74rem;}}
.hm .c .v {{font-size:1.05rem; font-weight:800;}}
.hm .h {{font-size:.66rem; letter-spacing:.08em; text-transform:uppercase; color:{MUTE}; font-weight:700; align-self:end; padding:0 4px 2px;}}
.hm .rh {{display:flex; align-items:center; font-weight:800; color:{NAVY}; font-size:.85rem;}}
.leg {{display:flex; align-items:center; gap:8px; font-size:.7rem; color:{MUTE}; margin-top:8px;}}
.leg i {{display:block; width:160px; height:8px; border-radius:4px;
  background:linear-gradient(90deg,{RED},#F6C3CD,#F3F4F8,#B9E1E3,{TEAL});}}
/* convertis */
.cv {{background:#fff; border:1px solid #E3E5EE; border-radius:14px; padding:14px 16px;
  box-shadow:0 1px 2px rgba(20,20,60,.04),0 8px 22px rgba(20,20,60,.05);}}
.cv .t {{display:flex; justify-content:space-between; align-items:center;}}
.cv .t b {{color:{NAVY}; font-size:.95rem;}}
.cv .m {{display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:6px; margin-top:10px;}}
.cv .m div {{background:{GREY}; border-radius:8px; padding:6px 8px;}}
.cv .m .l {{font-size:.58rem; color:{MUTE}; text-transform:uppercase; letter-spacing:.05em; font-weight:700;}}
.cv .m .v {{font-size:.88rem; font-weight:800;}}
/* bandeau d'export */
.dl {{display:flex; align-items:center; gap:14px; background:linear-gradient(120deg,#1B1A4A,{NAVY}); color:#fff;
  border-radius:14px; padding:16px 20px; margin-top:24px;}}
.dl b {{color:{ORANGE};}} .dl span {{color:#CFD0E6; font-size:.85rem;}}
div[data-testid="stDownloadButton"] button {{border-radius:10px; font-weight:700; font-family:'Inter',sans-serif; height:44px;}}
div[data-testid="stDownloadButton"] button[kind="primary"] {{background:{NAVY}; border-color:{NAVY};}}
div[data-testid="stDownloadButton"] button[kind="primary"]:hover {{background:#34338A; border-color:#34338A;}}
.frame {{background:#fff; border:1px solid #E3E5EE; border-radius:12px; padding:6px; margin-bottom:12px;
  box-shadow:0 8px 22px rgba(20,20,60,.06);}}
@media (max-width:1100px) {{ .kgrid {{grid-template-columns:repeat(2,1fr);}} .g3, .ret {{grid-template-columns:1fr;}} }}
</style>"""

st.markdown(LANDING_CSS + APP_CSS, unsafe_allow_html=True)


def html(s: str):
    st.markdown(_h(s), unsafe_allow_html=True)


def sec_title(t: str, hint: str = ""):
    html(f"<div class='bi'><div class='sec'><h3>{t}</h3><div class='line'></div><span class='hint'>{hint}</span></div></div>")


def cl(v) -> str:
    return "" if v is None else ("pos" if v >= 0 else "neg")


def dbar(v, m, label) -> str:
    """Barre divergente centrée sur zéro, échelle commune à la colonne."""
    if v is None or not m:
        return label
    w = min(abs(v) / m, 1) * 50
    side = f"left:50%;width:{w:.1f}%;background:{TEAL}" if v >= 0 else f"right:50%;width:{w:.1f}%;background:{RED}"
    return f"<div class='db'><span class='{cl(v)}'>{label}</span><div class='tr'><span style='{side}'></span></div></div>"


# colonnes : (titre, clé, type) ; types : txt, m (millions), pct, pts, taux, dm (écart en M avec barre), dpct (pct avec barre)
KPI_COLS = [("CA", "ca", "m"), ("CA vs N-1", "vs_n1", "dpct"), ("Marge", "marge", "m"), ("Marge vs N-1", "marge_vs_n1", "pct"),
            ("Écart marge", "dmarge", "dm"), ("Taux", "taux", "taux"), ("Écart taux", "dpts", "pts"),
            ("Effet volume", "effet_volume", "sm"), ("Effet taux", "effet_taux", "sm")]
TRAFIC_COLS = [("Débit", "debit_vs_n1", "pct"), ("Panier", "panier_vs_n1", "pct")]
SITE_COLS = KPI_COLS[:7] + TRAFIC_COLS   # sans les effets volume/taux, pour tenir à l'écran


def fmt_cell(t, v, m=None):
    if t == "txt":
        return v if v is not None else ""
    if v is None or (isinstance(v, float) and v != v):
        return "<span style='color:#B0B3C6'>n.d.</span>"
    if t == "m":
        return E.mfcfa(v)
    if t == "sm":
        return f"<span class='{cl(v)}'>{E.mfcfa(v, dec=2, sign=True)}</span>"
    if t == "taux":
        return E.pct(v, sign=False)
    if t == "pct":
        return f"<span class='{cl(v)}'>{E.pct(v)}</span>"
    if t == "pts":
        return f"<span class='{cl(v)}'>{E.pts(v)}</span>"
    if t == "dm":
        return dbar(v, m, E.mfcfa(v, dec=2, sign=True))
    if t == "dpct":
        return dbar(v, m, E.pct(v))
    return str(v)


def bi_table(rows: list[dict], cols: list[tuple], first: tuple, extra_left: list[tuple] = (), alert_key=None, total=None):
    """rows : dictionnaires de KPI ; first : (titre, clé) de la 1re colonne ; extra_left : colonnes texte à gauche."""
    allrows = rows + ([total] if total else [])
    mx = {k: max([abs(r.get(k) or 0) for r in allrows] or [0]) for _, k, t in cols if t in ("dm", "dpct")}
    head = f"<th class='l'>{first[0]}</th>" + "".join(f"<th class='l'>{h}</th>" for h, _ in extra_left) \
        + "".join(f"<th class='{'l' if t == 'txt' else ''}'>{h}</th>" for h, _, t in cols)
    body = ""
    for r in allrows:
        trc = "tot" if r is total else ("al" if alert_key and r.get(alert_key) else "")
        body += f"<tr class='{trc}'><td class='l first'>{r.get(first[1], '')}</td>"
        body += "".join(f"<td class='l'>{r.get(k, '')}</td>" for _, k in extra_left)
        body += "".join(f"<td class='{'l' if t == 'txt' else ''}'>{fmt_cell(t, r.get(k), mx.get(k))}</td>"
                        for _, k, t in cols) + "</tr>"
    html(f"<div class='bi'><div class='tw'><table class='bt'><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div></div>")


def heat(v):
    if v is None:
        return GREY, INK
    a = min(abs(v) / 0.30, 1)
    base = (8, 134, 141) if v >= 0 else (230, 34, 68)
    bg = f"rgba({base[0]},{base[1]},{base[2]},{0.10 + 0.80 * a:.2f})"
    return bg, ("#fff" if a > 0.45 else INK)


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

# ------------------------------------------------------------------ contrôle des extractions
icon = {"ok": "✓", "alerte": "!", "erreur": "✕"}
nb_ok = sum(c.statut == "ok" for c in checks)
pills = "".join(f"<span class='p {c.statut}' title='{c.message}'>{icon.get(c.statut, '•')} "
                f"{c.message.split(' : ')[0] if c.statut == 'ok' else c.message}</span>" for c in checks)
html(f"<div class='bi'><div class='bi-check'><span class='h'>Contrôle des extractions · {nb_ok}/{len(checks)} conformes</span>"
     f"{pills}</div></div>")
if R is None:
    st.error("Génération bloquée : corrigez l'extraction signalée en rouge (voir les consignes dans la barre latérale).")
    st.stop()

# ------------------------------------------------------------------ bandeau
r, sem = R["reseau"], R["reseau_sem"]
rep = R["date_reponse"]
n_quest = sum(x["type"] == "Réponse attendue" for x in R["retours"])
m0, ev, et = r["marge_n1"], r["effet_volume"], r["effet_taux"]
lv = [m0, m0 + ev, m0 + ev + et]
lo, hi = min(lv) * 0.80, max(lv) * 1.02


def _x(v):
    return (v - lo) / (hi - lo) * 100


def wf_row(lab, a, b, color, val):
    x0, x1 = sorted((_x(a), _x(b)))
    return (f"<div class='wf-r'><span>{lab}</span><div class='wf-t'><span style='left:{x0:.1f}%;width:{max(x1 - x0, 0.8):.1f}%;"
            f"background:{color}'></span></div><span class='v'>{val}</span></div>")


wf = ("<div class='lp-status'><h6>D'où vient l'écart de marge</h6><div class='wf'>"
      + wf_row("Marge N-1", lo, m0, "#8E8FB8", E.mfcfa(m0))
      + wf_row("Effet volume", m0, m0 + ev, TEAL if ev >= 0 else RED, E.mfcfa(ev, sign=True))
      + wf_row("Effet taux", m0 + ev, m0 + ev + et, TEAL if et >= 0 else RED, E.mfcfa(et, sign=True))
      + wf_row("Marge N", lo, r["marge"], ORANGE, E.mfcfa(r["marge"]))
      + "</div><div class='wf-n'>Échelle zoomée sur l'écart · effet volume = ventes en plus au taux N-1</div></div>")

html(f"""<div class='lp'><div class='lp-hero'><div class='lp-hl'>
<div class='lp-eyebrow'>Flash commercial · ventes du {JOURS[R['date'].weekday()]} {R['date'].strftime('%d/%m/%Y')}</div>
<div class='lp-title'>CA <b>{E.pct(r['vs_n1'], 0)}</b>, MARGE <b>{E.pct(r['marge_vs_n1'], 0)}</b></div>
<div class='lp-sub'>{R['lecture']}</div>
<div class='lp-meta'>
<span class='lp-chip'><b>{len(R['sites_alerte'])}</b> magasins en alerte</span>
<span class='lp-chip'><b>{len(R['alertes'])}</b> rayons × magasins hors seuil</span>
<span class='lp-chip'><b>{n_quest}</b> questions aux secteurs · réponse <b>{JOURS[rep.weekday()]} {rep.strftime('%d/%m')} {C.HEURE_REPONSE}</b></span>
</div></div>{wf}</div></div>""")


def kcard(acc, ic, lab, val, dlt, dval, sub, cmp=None, val_tone=None):
    arrow = "▲" if (dval or 0) >= 0 else "▼"
    c = ""
    if cmp:
        a, b = cmp
        mx = max(a, b) or 1
        c = (f"<div class='cmp'><div>N<i style='width:{a / mx * 100:.0f}%;background:{acc}'></i></div>"
             f"<div>N-1<i style='width:{b / mx * 100:.0f}%;background:#D5D7E4'></i></div></div>")
    return (f"<div class='kc' style='--acc:{acc};--tint:{acc}1A'><div class='hd'><span class='ic'>{ic}</span>{lab}</div>"
            f"<div class='v {cl(val_tone)}'>{val}</div><span class='d {cl(dval)}'>{arrow} {dlt}</span>"
            f"<div class='s'>{sub}</div>{c}</div>")


cards = [
    kcard(NAVY, "F", "Chiffre d'affaires", E.mfcfa(r["ca"]), E.pct(r["vs_n1"]) + " vs N-1", r["vs_n1"],
          f"N-1 : {E.mfcfa(r['ca_n1'])} FCFA HT", (r["ca"], r["ca_n1"])),
    kcard(ORANGE, "Σ", "Marge", E.mfcfa(r["marge"]), E.pct(r["marge_vs_n1"]) + " vs N-1", r["marge_vs_n1"],
          f"soit {E.mfcfa(r['dmarge'], sign=True)} FCFA", (r["marge"], r["marge_n1"])),
    kcard(TEAL, "%", "Taux de marge", E.pct(r["taux"], sign=False), E.pts(r["dpts"]), r["dpts"],
          f"N-1 : {E.pct(r['taux_n1'], sign=False)}", (r["taux"], r["taux_n1"])),
    kcard("#4B4AA8", "↗", "Fréquentation", E.pct(r.get("debit_vs_n1")), "panier " + E.pct(r.get("panier_vs_n1")),
          r.get("panier_vs_n1"), "Débit (tickets) vs N-1", val_tone=r.get("debit_vs_n1")),
    kcard(RED if (sem.get("vs_budget") or 0) < 0 else TEAL, "◎", "Semaine vs budget", E.pct(sem.get("vs_budget")),
          E.pct(sem["vs_n1"]) + " vs N-1", sem["vs_n1"], f"Cumul semaine : {E.mfcfa(sem['ca'])} FCFA",
          val_tone=sem.get("vs_budget")),
]
html("<div class='bi'><div class='kgrid'>" + "".join(cards) + "</div>"
     "<div class='ret'>" + "".join(f"<div><b class='n'>{i}</b><span>{m}</span></div>"
                                   for i, m in enumerate(R["messages"], 1)) + "</div></div>")

# ------------------------------------------------------------------ onglets
t_sec, t_fmt, t_ou, t_ret, t_ppt = st.tabs(["Secteurs", "Formats et zones", "Où agir", "Retours attendus", "Aperçu PPT"])

with t_sec:
    sc = ""
    for sec, k in R["secteurs"].items():
        vb = k["sem"].get("vs_budget")
        rays = "".join(f"<div class='ry'><b>{x['nom']}</b><span><span class='{cl(x['dmarge'])}'>{E.kfcfa(x['dmarge'])}</span>"
                       f" · <span class='{cl(x['dpts'])}'>{E.pts(x['dpts'])}</span></span></div>"
                       for x in R["rayons"].get(sec, [])[:3]) or "<div class='ry'>Rien hors seuil</div>"
        sc += f"""<div class='sc'><div class='top'><span class='nm'>{sec.upper()}</span>
<span class='pill {'tl' if k['dmarge'] >= 0 else 'rd'}'>{E.mfcfa(k['dmarge'], sign=True)} FCFA</span></div>
<div class='big {cl(k['marge_vs_n1'])}'>{E.pct(k['marge_vs_n1'])}</div><div class='bl'>marge vs N-1 · {E.mfcfa(k['marge'])} FCFA</div>
<div class='mg'><div><div class='l'>CA</div><div class='v'>{E.mfcfa(k['ca'])}</div><div class='x {cl(k['vs_n1'])}'>{E.pct(k['vs_n1'])}</div></div>
<div><div class='l'>Taux</div><div class='v'>{E.pct(k['taux'], sign=False)}</div><div class='x {cl(k['dpts'])}'>{E.pts(k['dpts'])}</div></div>
<div><div class='l'>Sem. vs bgt</div><div class='v {cl(vb)}'>{E.pct(vb)}</div><div class='x' style='color:{MUTE}'>cumul</div></div></div>
<div class='k'>Rayons à regarder</div>{rays}</div>"""
    html(f"<div class='bi'><div class='g3' style='margin-top:14px'>{sc}</div></div>")

    sec_title("Secteurs · détail", "hier vs même jour N-1")
    bi_table([{"nom": n, **k} for n, k in R["secteurs"].items()], KPI_COLS, ("Secteur", "nom"),
             total={"nom": "Réseau", **R["reseau"]})

    sec_title("Secteur × format", "couleur = marge vs N-1 · texte = écarts hors seuil")
    fmts = C.FORMATS
    grid = f"<div class='hm' style='grid-template-columns:140px repeat({len(fmts)},1fr)'><div></div>"
    grid += "".join(f"<div class='h'>{f}</div>" for f in fmts)
    for s_, d_ in R["croise"].items():
        grid += f"<div class='rh'>{s_}</div>"
        for f in fmts:
            v = d_.get(f)
            if not v:
                grid += f"<div class='c' style='background:{GREY}'></div>"
                continue
            bg, fg = heat(v["kpi"].get("marge_vs_n1"))
            grid += (f"<div class='c' style='background:{bg};color:{fg}'><div class='v'>{E.pct(v['kpi'].get('marge_vs_n1'), 0)}</div>"
                     f"<div>{v['texte'] or 'Dans les seuils'}</div></div>")
    grid += "</div><div class='leg'>−30 %<i></i>+30 %</div>"
    html(f"<div class='bi'>{grid}</div>")

with t_fmt:
    sec_title("Formats", "débit et panier vs N-1")
    bi_table([{"nom": n, **k} for n, k in R["formats"].items()], SITE_COLS, ("Format", "nom"))
    sec_title("Zones de pricing")
    bi_table([{"nom": n, **k} for n, k in R["zones"].items()], SITE_COLS, ("Zone", "nom"))
    sec_title("Magasins", "liseré rouge = magasin en alerte")
    bi_table([{**x, "fz": f"<span class='pill nav'>{x['format']}</span> <span class='pill or'>{x['zone']}</span>"}
              for x in R["sites"]], SITE_COLS, ("Magasin", "nom"), [("Format · zone", "fz")], alert_key="alerte")

with t_ou:
    if True:
        sec_title("Alertes rayon × magasin", f"triées par perte de marge · {len(R['alertes'])} lignes")
        rows = [{**x, "ry": f"{E.short_rayon(x['rayon'])}<div class='sub'>{E.short_site(x['site'])} · {x['secteur']}</div>",
                 "dg": f"<span class='pill or'>{x['diagnostic']}</span>",
                 "qu": f"<span class='pill tl'>{x['qui']}</span>",
                 "pe": "<span class='pill rd'>Oui</span>" if x.get("persistant") else ""} for x in R["alertes"]]
        bi_table(rows, [("CA vs N-1", "vs_n1", "dpct"), ("Écart marge", "dmarge", "dm"), ("Écart taux", "dpts", "pts"),
                        ("Diagnostic", "dg", "txt"), ("Qui", "qu", "txt"), ("Persiste", "pe", "txt")],
                 ("Rayon · magasin", "ry"))
    if True:
        sec_title("Magasins convertis", "comparés à leur ancien format Supeco : on lit surtout la tendance")
        cv = ""
        for x in R["convertis"]:
            vb = x["sem"].get("vs_budget")
            m = "".join(f"<div><div class='l'>{l}</div><div class='v {cl(v)}'>{E.pct(v, 0) if v is not None else 'n.d.'}</div></div>"
                        for l, v in [("Marge", x["marge_vs_n1"]), ("Débit", x.get("debit_vs_n1")),
                                     ("Panier", x.get("panier_vs_n1")), ("Budget", vb)])
            dte = x["converti_le"].strftime("%d/%m/%Y") if x.get("converti_le") else "n.d."
            cv += (f"<div class='cv'><div class='t'><b>{x['nom']}</b><span class='pill nav'>converti le {dte}</span></div>"
                   f"<div class='m'>{m}</div></div>")
        html(f"<div class='bi'><div class='g3'>{cv}</div></div>")

with t_ret:
    sec_title(f"{n_quest} questions aux secteurs",
              f"réponse le {JOURS[rep.weekday()]} {rep.strftime('%d/%m')} avant {C.HEURE_REPONSE} · modifiable avant génération")
    ed = st.data_editor(pd.DataFrame(R["retours"])[["sujet", "question", "qui", "type"]],
                        hide_index=True, width="stretch", num_rows="dynamic",
                        column_config={"sujet": st.column_config.TextColumn("Sujet", width="medium"),
                                       "question": st.column_config.TextColumn("Question", width="large"),
                                       "qui": st.column_config.TextColumn("Qui répond", width="medium"),
                                       "type": st.column_config.SelectboxColumn(
                                           "Type", options=["Réponse attendue", "Pour information"], width="small")})
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
            sec_title("Aperçu du PowerPoint", f"{len(imgs)} slides sur le masque CFAO")
            c = st.columns(2, gap="medium")
            for i, im in enumerate(imgs):
                with c[i % 2]:
                    st.image(im, width="stretch")
        else:
            st.info("Aperçu image indisponible sur ce poste (LibreOffice absent). Plan du PPT généré :")
            st.markdown(f"""1. **Couverture** · Flash commercial, ventes du {R['date'].strftime('%d/%m/%Y')}
2. **L'essentiel** · 4 chiffres clés, 3 messages, marge vs N-1 par secteur
3. **Les secteurs** · une carte par secteur et les rayons à regarder
4. **Où agir** · {len(R['sites_alerte'])} magasins en alerte, écarts hors seuil, magasins convertis
5. **Retours attendus** · {sum(x['type'] == 'Réponse attendue' for x in R['retours'])} questions
6. **Comment lire ce flash**""")

html(f"<div class='bi'><div class='dl'><div><b>Diffuser</b><br><span>PowerPoint pour les secteurs · Excel de détail "
     f"en valeurs · ventes du {R['date'].strftime('%d/%m/%Y')}</span></div></div></div>")
d1, d2, _ = st.columns([1, 1, 2])
d1.download_button("⬇  PowerPoint pour les secteurs", ppt_bytes or b"", file_name=f"Flash_commercial_{stamp}.pptx",
                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                   disabled=ppt_bytes is None, type="primary", width="stretch")
d2.download_button("⬇  Excel de détail", xls_bytes, file_name=f"Flash_quotidien_{stamp}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
