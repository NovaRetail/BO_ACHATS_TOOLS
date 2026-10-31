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
