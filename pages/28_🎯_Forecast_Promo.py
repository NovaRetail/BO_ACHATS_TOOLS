"""SmartBuyer Hub · Forecast Promo — challenger les prévisions des acheteurs, avant et après la promo.

Page de pilotage (usage personnel). Deux temps, mêmes fichiers :
  ① Avant la promo : la prévision de chaque article est confrontée au CA N-1 de la même période.
  ② Après la promo : prévu vs réalisé sans compensation entre articles, et comparaison à N-1.
Lecture niveau par niveau (Rayon, Famille, Sous-famille, Article) : top 15 de ce qu'il faut challenger, clic pour descendre.
Export Excel : valeurs uniquement, aucune formule.

Les calculs et l'export sont dans modules/forecast_promo/engine.py.
"""
from __future__ import annotations

import json
import pathlib
import re
import sys
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import theme as T  # noqa: E402
from modules.forecast_promo import engine as E  # noqa: E402

st.set_page_config(page_title="Forecast Promo · SmartBuyer", page_icon="🎯", layout="wide", initial_sidebar_state="expanded")
T.apply()

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ══════════════════════════════════════════════════════════════════════════════
# CACHE (les fichiers ne sont relus que s'ils changent)
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Lecture de l'extraction…")
def load_sales(data: bytes, name: str):
    return E.parse_sales(data, name)


@st.cache_data(show_spinner=False)
def load_promo(data: bytes, name: str, price_ttc: bool):
    return E.parse_promo(data, name, price_ttc)


@st.cache_data(show_spinner=False)
def run(sales, promo, params):
    return E.compute(sales, promo, dict(params))


@st.cache_data(show_spinner=False)
def export(a, tot, controls, filters, meta, params):
    return E.build_excel(a, tot, controls, filters, meta, dict(params))


# ══════════════════════════════════════════════════════════════════════════════
# LECTURE NIVEAU PAR NIVEAU (top 15 de ce qu'il faut challenger, drill au clic)
# ══════════════════════════════════════════════════════════════════════════════
def _clean(v):
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return None
    return float(v)


@st.cache_data(show_spinner=False)
def drill_html(a: pd.DataFrame, mode: str, params) -> str:
    p = dict(params)
    levels = []
    if a["dep"].nunique() > 1:
        levels.append(("dep", "Département"))
    levels += [("ray", "Rayon"), ("fam", "Famille")]
    if (a["sf"] != E.NO_SF).any():
        levels.append(("sf", "Sous-famille"))
    keys = [k for k, _ in levels] + ["label"]
    names = [lab for _, lab in levels] + ["Article"]
    rows = [[str(r[k]) for k in keys] + [r["code"], _clean(r["prev_ca"]), _clean(r["CA N-1"]), _clean(r["CA"]), r["control"],
                                         r["result"], _clean(r["short"]), bool(r["over"])] for r in a.to_dict("records")]
    data = json.dumps({"rows": rows, "names": names, "p": p, "mode": mode}, ensure_ascii=False).replace("</", "<\\/")
    c = {"ok": (f"{T.TEAL}17", T.VERT_TEXTE), "att": (f"{T.ORANGE}1F", T.ORANGE_TEXTE), "crit": (f"{T.RED}14", T.ROUGE_TEXTE),
         "mut": ("#24235C12", T.NAVY)}
    badges = "".join(f".{k}{{background:{bg};color:{fg}}}" for k, (bg, fg) in c.items())
    return f"""<html><head><meta charset='utf-8'>
<link href='https://fonts.googleapis.com/css2?family=Inter:wght@500;600;700;800&display=swap' rel='stylesheet'>
<style>
body{{margin:0;font-family:{T.PLOTLY_FONT};color:{T.NAVY};background:transparent}}
.top{{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:10px}}
.seg{{display:flex;gap:4px;background:#ECEDF4;border-radius:12px;padding:4px;flex-wrap:wrap;align-items:center}}
.seg span{{font-size:12px;font-weight:700;color:{T.GRIS_TEXTE};padding:0 6px 0 8px}}
.seg button{{font:inherit;font-size:13px;font-weight:700;border:0;background:transparent;color:{T.GRIS_TEXTE};padding:7px 14px;border-radius:9px;cursor:pointer}}
.seg button.on{{background:#fff;color:{T.NAVY};box-shadow:0 1px 2px rgba(20,20,60,.08)}}
.tgl{{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:600;color:{T.GRIS_TEXTE};cursor:pointer}}.tgl input{{accent-color:{T.NAVY};width:16px;height:16px}}
.card{{background:#fff;border:1px solid {T.LIGNE};border-radius:16px;padding:10px 14px 12px;box-shadow:{T.SHADOW}}}
.crumb{{font-size:13px;font-weight:600;color:{T.GRIS_TEXTE};padding:4px 2px 4px}}.crumb a{{color:{T.BLEU};cursor:pointer;font-weight:700}}.crumb b{{color:{T.NAVY}}}
.scope{{font-size:12.5px;color:{T.GRIS_TEXTE};padding:0 2px 8px}}.scope b{{color:{T.NAVY}}}
table{{width:100%;border-collapse:collapse;font-size:13.5px}}
th{{font-size:10.5px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{T.GRIS_TEXTE};text-align:right;padding:8px;border-bottom:1px solid {T.LIGNE}}}th:first-child{{text-align:left}}
td{{text-align:right;padding:10px 8px;border-bottom:1px solid #F0F1F5;font-weight:600}}td:first-child{{text-align:left}}
td small{{display:block;color:{T.GRIS_TEXTE};font-weight:500;font-size:11.5px}}
tr.go{{cursor:pointer}}tr.go:hover td{{background:{T.FOND}}}tr.go td:first-child:after{{content:" ›";color:{T.GRIS}}}
.b{{display:inline-block;border-radius:999px;padding:2px 10px;font-size:12px;font-weight:700}}{badges}
.neg{{color:{T.ROUGE_TEXTE}}}.more{{font-size:12.5px;color:{T.GRIS_TEXTE};padding:10px 8px 2px}}
</style></head><body>
<div class='top'><div class='seg' id='lvl'><span>Lire par</span></div><label class='tgl'><input type='checkbox' id='all'> Tout afficher</label></div>
<div class='card'><div class='crumb' id='crumb'></div><div class='scope' id='scope'></div><table><thead id='th'></thead><tbody id='tb'></tbody></table><div class='more' id='more'></div></div>
<script>
const X={data};const D=X.rows,N=X.names,P=X.p,MODE=X.mode,NL=N.length;
const I={{code:NL,p:NL+1,n1:NL+2,ca:NL+3,k:NL+4,res:NL+5,sh:NL+6,o:NL+7}};
let lvl=0,path=[],showAll=false;
const esc=s=>String(s).replace(/[&<>"']/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c]));
const M=v=>{{if(v==null)return'—';const a=Math.abs(v);return a>=1e9?(v/1e9).toFixed(2).replace('.',',')+' Md':a>=1e6?(v/1e6).toFixed(2).replace('.',',')+' M':a>=1e3?Math.round(v/1e3)+' k':Math.round(v)+''}};
const PC=v=>v==null||!isFinite(v)?'—':Math.round(v*100)+' %';
const R=v=>'× '+v.toFixed(1).replace('.',',');
const B=(t,k)=>`<span class="b ${{k}}">${{esc(t)}}</span>`;
const KC={{'À justifier':'crit','À surveiller':'att','Sans historique':'att','Cohérente':'ok','Surestimée':'crit','Pas de vente':'crit','Fiable':'ok','Sous-estimée':'att'}};
function agg(rs){{const t={{n:rs.length,p:0,ca:0,n1:0,sh:0,o:0,j:0,w:0,h:0,pj:0}};rs.forEach(r=>{{t.p+=r[I.p]||0;t.ca+=r[I.ca]||0;t.n1+=r[I.n1]||0;t.sh+=r[I.sh]||0;t.o+=r[I.o]?1:0;
 if(r[I.k]==='À justifier'){{t.j++;t.pj+=r[I.p]||0}}if(r[I.k]==='À surveiller')t.w++;if(r[I.k]==='Sans historique')t.h++}});
 t.ratio=t.n1?t.p/t.n1:null;t.real=t.p?t.ca/t.p:null;t.bias=t.n?t.o/t.n:null;return t}}
const lb=document.getElementById('lvl');N.forEach((n,i)=>{{const b=document.createElement('button');b.textContent=n;b.onclick=()=>{{if(i<path.length)path=path.slice(0,i);lvl=i;render()}};lb.appendChild(b)}});
function render(){{
 const sc=D.filter(r=>path.every((v,i)=>r[i]===v)),T=agg(sc),lv=Math.max(lvl,path.length);
 let c=`<a onclick="go(0)">Tout le périmètre</a>`;path.forEach((v,i)=>c+=` › <a onclick="go(${{i+1}})">${{esc(v)}}</a>`);
 document.getElementById('crumb').innerHTML=c+` › <b>par ${{N[lv].toLowerCase()}}</b>`;
 document.getElementById('scope').innerHTML=MODE==='avant'
  ?`Sur ce périmètre : CA prévu <b>${{M(T.p)}}</b> pour <b>${{M(T.n1)}}</b> en N-1 (${{T.ratio?R(T.ratio):'—'}}) · <b>${{T.j}}</b> à justifier, ${{T.w}} à surveiller, ${{T.h}} sans historique`
  :`Sur ce périmètre : CA réalisé <b>${{M(T.ca)}}</b> pour <b>${{M(T.p)}}</b> prévus (${{PC(T.real)}}) · CA promis non réalisé <b>${{M(T.sh)}}</b> · biais <b>${{PC(T.bias)}}</b>`;
 const g=new Map();sc.forEach(r=>{{const k=r[lv];if(!g.has(k))g.set(k,[]);g.get(k).push(r)}});
 let L=[...g].map(([name,rs])=>({{name,rs,t:agg(rs),art:lv===NL-1?rs[0]:null}}));
 const total=L.length;
 if(!showAll)L=L.filter(x=>MODE==='avant'?(x.t.j+x.t.w+x.t.h)>0:(x.t.o>0&&x.t.sh>0));
 const nb=L.length;
 L.sort(MODE==='avant'?(a,b)=>(b.t.j-a.t.j)||(b.t.pj-a.t.pj)||((b.t.ratio??99)-(a.t.ratio??99)):(a,b)=>b.t.sh-a.t.sh);
 const S=showAll?L:L.slice(0,15);
 document.getElementById('th').innerHTML=MODE==='avant'?`<tr><th>${{N[lv]}}</th><th>CA prévu</th><th>CA N-1</th><th>Prévu / N-1</th><th>Contrôle</th></tr>`
  :`<tr><th>${{N[lv]}}</th><th>CA prévu</th><th>CA réalisé</th><th>Réal.</th><th>CA promis non réalisé</th><th>Biais</th></tr>`;
 document.getElementById('tb').innerHTML=S.map((x,ix)=>{{const t=x.t,a=x.art,go=lv<NL-1?` class="go" data-i="${{ix}}"`:'';
  const sub=a?`<small>${{esc(a[I.code])}}</small>`:`<small>${{t.n}} art.</small>`;
  if(MODE==='avant'){{const rk=t.ratio==null?B('pas de N-1','att'):B(R(t.ratio),t.ratio>P.justify?'crit':t.ratio>P.watch?'att':'ok');
   const ct=a?B(a[I.k],KC[a[I.k]]):(t.j?B(t.j+' à justifier','crit'):t.w?B(t.w+' à surveiller','att'):t.h?B(t.h+' sans N-1','att'):B('RAS','ok'));
   return `<tr${{go}}><td>${{esc(x.name)}}${{sub}}</td><td>${{M(t.p)}}</td><td>${{M(t.n1)}}</td><td>${{rk}}</td><td>${{ct}}</td></tr>`}}
  const rr=t.real==null?B('—','mut'):B(PC(t.real),t.real<P.low?'crit':t.real>P.high?'att':'ok');
  const bi=a?B(a[I.res],KC[a[I.res]]):B(PC(t.bias),t.bias>=.5?'crit':t.bias>=.25?'att':'ok');
  return `<tr${{go}}><td>${{esc(x.name)}}${{sub}}</td><td>${{M(t.p)}}</td><td><b>${{M(t.ca)}}</b></td><td>${{rr}}</td><td class="neg">${{t.sh?M(t.sh):'—'}}</td><td>${{bi}}</td></tr>`}}).join('')
  ||`<tr><td colspan="6" style="text-align:center;color:{T.GRIS_TEXTE}">Rien à challenger à ce niveau · cocher « Tout afficher »</td></tr>`;
 document.querySelectorAll('#tb tr.go').forEach(tr=>tr.onclick=()=>{{path.push(S[+tr.dataset.i].name);lvl=path.length;render()}});
 const hid=total-S.length;
 document.getElementById('more').textContent=showAll||hid<=0?'':`${{S.length}} ligne(s) à challenger affichée(s) sur ${{total}} · ${{total-nb}} sans problème masquée(s)`+(nb>15?` · ${{nb-15}} hors top 15`:'');
 [...lb.querySelectorAll('button')].forEach((b,i)=>b.classList.toggle('on',i===lv));
}}
function go(i){{path=path.slice(0,i);lvl=i;render()}}
document.getElementById('all').onchange=e=>{{showAll=e.target.checked;render()}};
render();
</script></body></html>"""


def drill(a, mode, params):
    page = drill_html(a, mode, params)  # libellés échappés côté navigateur
    if hasattr(st, "iframe"):
        st.iframe(page, height=820)
    else:  # versions de Streamlit antérieures à st.iframe
        components.html(page, height=820, scrolling=True)


# ══════════════════════════════════════════════════════════════════════════════
# BARRE LATÉRALE
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_brand("🎯", "Forecast Promo")
T.sidebar_label("1 · Extraction ventes")
with st.sidebar:
    f_sales = st.file_uploader("Export PBI ventes par article (xlsx / csv)", type=["xlsx", "xlsm", "csv"], key="fp_sales",
                               help="Export hiérarchique Département › Rayon › Famille › Sous-famille › Article › Site, sur les dates de la promo.")
T.sidebar_label("2 · Liste promo des acheteurs")
with st.sidebar:
    f_promo = st.file_uploader("Code, Prévision Qté, Prix promo, PA, TVA", type=["xlsx", "xlsm", "csv"], key="fp_promo",
                               help="Une ligne par article. CA prévu et marge prévue sont calculés par le module.")
    price_ttc = st.toggle("Prix promo saisi TTC", value=True, help="Activé : le prix promo est ramené en HT avec la colonne TVA.")
T.sidebar_label("3 · Règles")
with st.sidebar:
    watch = st.slider("① À surveiller au-delà de (× N-1)", 1.0, 2.0, 1.3, 0.1)
    justify = st.slider("① À justifier au-delà de (× N-1)", 1.2, 5.0, 2.0, 0.1)
    low = st.slider("② Surestimée sous", 50, 100, 80, 5, format="%d %%") / 100
    high = st.slider("② Sous-estimée au-delà de", 100, 200, 120, 5, format="%d %%") / 100
    stable = st.slider("② Stable vs N-1 (±)", 0, 20, 5, 1, format="%d %%") / 100
    converted = st.multiselect("Sites au N-1 non comparable", E.DEFAULT_CONVERTED_SITES, default=E.DEFAULT_CONVERTED_SITES)
params = (("watch", watch), ("justify", max(justify, watch)), ("low", low), ("high", max(high, low)), ("stable", stable),
          ("converted_sites", tuple(converted)))
P = dict(params)


# ══════════════════════════════════════════════════════════════════════════════
# ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════
if f_sales is None or f_promo is None:
    T.sidebar_footer("CA prévu HT = Qté × Prix promo ÷ (1 + TVA)<br>Marge prévue = Qté × (Prix HT − PA)")
    rows = [("Extraction ventes PBI", "ok" if f_sales else "wait", "reçue" if f_sales else "à déposer"),
            ("Liste promo des acheteurs", "ok" if f_promo else "wait", "reçue" if f_promo else "à déposer")]
    T.landing_hero("FORECAST <b>PROMO</b>",
                   "La prévision de l'acheteur tient-elle face à l'an dernier, et a-t-elle été tenue ? "
                   "On la contrôle avant la promo, on la mesure après, et on ne montre que ce qu'il faut challenger.",
                   eyebrow="Forecast Promo · SmartBuyer Hub",
                   status=T.landing_status("Préparation de l'analyse", rows, "fichiers obligatoires reçus", count=2))
    T.steps([("Collecter", "la liste promo de chaque acheteur : Code, Prévision Qté, Prix promo, PA, TVA"),
             ("Extraire", "les ventes PBI sur les dates de la promo (avant : seules les colonnes N-1 sont remplies)"),
             ("① Contrôler", "avant validation : les prévisions à × 2 le N-1 sont à justifier"),
             ("② Mesurer", "après la promo : CA promis non réalisé et biais, niveau par niveau"),
             ("Exporter", "le classeur pour la revue avec les acheteurs")])
    T.section("Format des fichiers attendus")
    c1, c2, c3 = st.columns(3)
    with c1:
        T.columns_card("Colonnes attendues", [
            ("Liste promo des acheteurs", ["Code", "Prévision Qté", "Prix promo (TTC par défaut)", "PA (HT)", "TVA (18, 0,18 ou 18 %)",
                                           "Libellé (facultatif)"]),
            ("Extraction ventes (export PBI)", ["Departement · Rayon · Famille · Sous-Famille", "Article", "Site nom long", "CA · Marge · Qté Vente",
                                                "CA N-1 · Qté Vente N-1 · Marge Hors Promo N-1 · Marge Promo N-1"]),
        ], "Article au format « 12003043 - LIBELLÉ » : le code avant le tiret sert de clé. Les lignes « Total » sont ignorées. "
           "Sans l'axe Sous-Famille, la lecture passe directement de Famille à Article.")
    with c2:
        T.rules_card("📐 Méthode de calcul", [
            ("CA prévu HT", "Qté × Prix promo ÷ (1 + TVA)"),
            ("Marge prévue", "Qté × (Prix promo HT − PA)"),
            ("① Contrôle", f"{T.badge('Cohérente', 'ok')} {T.badge('× 1,3', 'att')} {T.badge('× 2', 'crit')}"),
            ("② Prévision", f"{T.badge('Fiable', 'ok')} {T.badge('Surestimée', 'crit')} {T.badge('Sous-estimée', 'att')}"),
            ("CA promis non réalisé", "somme des manques article par article"),
            ("Biais", "part d'articles surestimés ou sans vente"),
        ])
    with c3:
        T.rules_card("⚠️ Points d'attention", [
            ("Période", "l'extraction couvre uniquement les dates de la promo"),
            ("Montants", "exports PBI en HT : le prix promo TTC est ramené en HT"),
            ("Quantité réalisée", "quantité totale vendue (pas de quantité promo dans l'export)"),
            ("Marge N-1", "Marge Hors Promo N-1 + Marge Promo N-1"),
            ("Sites convertis", "Cité verte, Aboboté, Toit rouge : N-1 non comparable"),
            ("Exports Excel", "valeurs uniquement, aucune formule"),
        ])
    T.section("📖 Lexique")
    half = (len(E.LEXIQUE) + 1) // 2
    c1, c2 = st.columns(2)
    with c1:
        T.rules_card("Avant la promo et calculs", [(f"<b>{t}</b>", d) for t, d in E.LEXIQUE[:half]], tall=False)
    with c2:
        T.rules_card("Après la promo et comparaison N-1", [(f"<b>{t}</b>", d) for t, d in E.LEXIQUE[half:]], tall=False)
    T.section("Exemple de liste promo")
    st.dataframe(pd.DataFrame({"Code": [12003029, 12003043, 12503300],
                               "Libellé": ["X42 COUCH STD PACK MINI JOYA", "X40 COUCH STD PACK MAXI SMILEY", "X55 MOUCHOIRS FACIALE 4P STYLE"],
                               "Prévision Qté": [200, 1500, 900], "Prix promo": [3900, 2950, 450], "PA": [2700, 2250, 260], "TVA": [18, 18, 18]}),
                 width="stretch", hide_index=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT (en cache)
# ══════════════════════════════════════════════════════════════════════════════
try:
    sales, filters, smeta = load_sales(f_sales.getvalue(), f_sales.name)
except Exception as exc:  # noqa: BLE001
    T.page_header("🎯", "Forecast Promo")
    T.alert("Extraction ventes illisible", str(exc), "crit")
    st.stop()
try:
    promo, meta = load_promo(f_promo.getvalue(), f_promo.name, price_ttc)
except Exception as exc:  # noqa: BLE001
    T.page_header("🎯", "Forecast Promo")
    T.alert("Liste promo illisible", f"{exc} Colonnes attendues : Code, Prévision Qté, Prix promo, PA, TVA.", "crit")
    st.stop()
if sales.empty:
    T.page_header("🎯", "Forecast Promo")
    T.alert("Aucune ligne exploitable", "L'export ne contient aucune ligne article × site.", "crit")
    st.stop()

a, tot, controls = run(sales, promo, params)
per = f"du {filters['period_from']:%d/%m/%Y} au {filters['period_to']:%d/%m/%Y}" if filters["period_from"] else "période non lue"
before = tot["phase"] == "avant"
T.sidebar_footer(f"{tot['n']} articles · {tot['sites']} sites<br>{per}<br>CA prévu : {meta['ca_mode']}")


# ══════════════════════════════════════════════════════════════════════════════
# BANDEAU
# ══════════════════════════════════════════════════════════════════════════════
if before:
    title = f"FORECAST · <b>{E.fmt_ratio(tot['ratio'])}</b> LE N-1"
    chips = [f"{tot['n']} articles", (f"{tot['justify']} à justifier", "alert") if tot["justify"] else ("Aucune prévision à justifier", "ok"),
             f"{tot['watch']} à surveiller", f"{tot['nohist']} sans historique", ("Promo pas encore vendue", "")]
else:
    title = f"CA RÉALISÉ · <b>{E.fmt_pct(tot['real_ca'])}</b> DU FORECAST"
    chips = [f"{tot['n']} articles", (f"CA promis non réalisé {E.fmt_fcfa(tot['short'])}", "alert") if tot["short"] else ("Forecast tenu", "ok"),
             (f"Biais {E.fmt_pct(tot['bias'])}", "alert" if (tot["bias"] or 0) >= .25 else "ok"),
             f"{tot['justify']} prévision(s) étaient à justifier", f"CA vs N-1 {E.fmt_pct(tot['ca_evo'], 0, True)}"]
T.hero(title, f"{filters['supplier'] or 'Fournisseur non lu'} · extraction {per}", chips, eyebrow="Forecast Promo")
if not smeta["has_sf"]:
    T.alert("Sous-famille absente de l'extraction", "Ajoutez l'axe Sous-Famille entre Famille et Article dans l'export PBI pour lire aussi à ce niveau.", "info")

n_ctrl = sum(len(controls[k]) for k in ["not_found", "duplicates", "no_price", "neg_marge"]) + (len(controls["no_pa"]) if meta["has_pa"] else 1)
tab1, tab2, tab3, tab4 = st.tabs(["① Avant la promo", "② Après la promo", f"🧪 Contrôles · {n_ctrl}", "⬇️ Export"])

# ── ① Avant la promo ──────────────────────────────────────────────────────────
with tab1:
    T.kpi_row([
        T.kpi("CA prévu HT", E.fmt_fcfa(tot["prev_ca"]), E.fmt_ratio(tot["ratio"]) + " le N-1",
              "neg" if (tot["ratio"] or 0) > P["justify"] else "mut", f"CA N-1 même période {E.fmt_fcfa(tot['ca_n1'])}", "blue", "🎯", "FCFA"),
        T.kpi("Prévisions à justifier", str(tot["justify"]), f"{tot['watch']} à surveiller", "neg" if tot["justify"] else "pos",
              f"prévision > × {P['justify']:.1f} le N-1".replace(".", ","), "violet", "🚩"),
        T.kpi("CA prévu à justifier", E.fmt_fcfa(tot["justify_ca"]), E.fmt_pct(E.div(tot["justify_ca"], tot["prev_ca"])) + " du forecast",
              "neg" if tot["justify_ca"] else "pos", "part du forecast portée par ces prévisions", "violet", "⚖️", "FCFA"),
        T.kpi("Sans historique", str(tot["nohist"]), "", "mut", "aucune vente N-1 sur la période", "teal", "🆕"),
    ])
    T.card(f"<div style='font-size:14px;font-weight:600;line-height:1.55'>{E.sentence_before(tot)}</div>")
    T.section("À challenger avant validation")
    drill(a, "avant", params)

# ── ② Après la promo ──────────────────────────────────────────────────────────
with tab2:
    if before:
        T.alert("La promo n'a pas encore de ventes", "L'extraction ne contient que le N-1 : utilisez l'onglet ① pour contrôler le forecast. "
                "Rechargez la même extraction après l'opération pour mesurer le résultat.", "info")
    else:
        sg = lambda x: "pos" if (x or 0) >= P["low"] else "neg"  # noqa: E731
        cards = [
            T.kpi("CA réalisé HT", E.fmt_fcfa(tot["ca"]), E.fmt_pct(tot["real_ca"]) + " du forecast", sg(tot["real_ca"]),
                  f"forecast {E.fmt_fcfa(tot['prev_ca'])}", "blue", "🧾", "FCFA"),
            T.kpi("CA promis non réalisé", E.fmt_fcfa(tot["short"]), "à challenger" if tot["short"] else "", "neg" if tot["short"] else "pos",
                  "somme des manques, sans compensation", "violet", "📉", "FCFA"),
            T.kpi("Biais", E.fmt_pct(tot["bias"]).replace(" %", ""), f"{tot['over']} / {tot['n']} articles",
                  "neg" if (tot["bias"] or 0) >= .25 else "pos", "surestimés ou sans vente", "violet", "🎯", "%"),
            T.kpi("CA vs N-1", E.fmt_pct(tot["ca_evo"], 0, True).replace(" %", ""), f"marge {E.fmt_pct(tot['marge_evo'], 0, True)}",
                  "pos" if (tot["ca_evo"] or 0) >= 0 else "neg",
                  f"taux {E.fmt_pct(tot['rate'], 1)} · prévu {E.fmt_pct(tot['prev_rate'], 1)} · N-1 {E.fmt_pct(tot['rate_n1'], 1)}", "teal", "📅", "%"),
        ]
        T.kpi_row(cards)
        T.card(f"<div style='font-size:14px;font-weight:600;line-height:1.55'>{E.sentence_after(tot)}</div>")
        if not meta["has_pa"]:
            T.alert("Marge prévue non calculée", "Ajoutez PA et TVA dans la liste promo pour comparer la marge au forecast.", "att")
        if controls["converted"]:
            T.alert("N-1 partiellement non comparable", f"Sites convertis présents : {', '.join(controls['converted'])}. "
                    "Évolutions vs N-1 à lire avec prudence ; détail par article dans l'export.", "att")
        T.section("À challenger avec les acheteurs")
        drill(a, "apres", params)

# ── Contrôles ─────────────────────────────────────────────────────────────────
with tab3:
    def rr(ok, lab, det):
        return (f"<div class='rule-row'><span><b>{lab}</b><small style='display:block;color:{T.GRIS_TEXTE};font-weight:500'>{det}</small></span>"
                f"<span style='align-self:center'>{T.badge('OK', 'ok') if ok else T.badge('À vérifier', 'att')}</span></div>")
    days = (filters["period_to"] - filters["period_from"]).days + 1 if filters["period_from"] else None
    T.section("Extraction")
    T.card(rr(days is not None and days <= 31, "Période", per + (f", soit {days} jours : au-delà de la promo, les résultats sont dilués." if days else ""))
           + rr(bool(filters["supplier"]), "Fournisseur", filters["supplier"] or "Non lu dans les filtres.")
           + rr(True, "Sites", f"{tot['sites']} sites" + (f" : {', '.join(filters['sites'])}" if filters["sites"] else ""))
           + rr(smeta["has_sf"], "Niveaux de lecture", " › ".join(smeta["levels_found"]) or "aucun"))
    T.section("Liste promo")
    T.card(rr(not controls["not_found"], "Codes trouvés", f"{int(a['found'].sum())} / {len(a)}"
              + (f" · introuvables : {', '.join(controls['not_found'])}" if controls["not_found"] else ""))
           + rr(not controls["duplicates"], "Doublons", ", ".join(controls["duplicates"]) or "Aucun.")
           + rr(not controls["no_price"], "Prix promo manquant", ", ".join(controls["no_price"]) or "Aucun.")
           + (rr(not controls["no_pa"], "PA manquant", ", ".join(controls["no_pa"]) or "Aucun.") if meta["has_pa"]
              else rr(False, "PA et TVA", "Colonnes absentes : la marge prévue n'est pas calculée."))
           + rr(not controls["neg_marge"], "Marge prévue négative (prix promo HT < PA)", ", ".join(controls["neg_marge"]) or "Aucune.")
           + rr(meta["has_vat"] or not price_ttc, "TVA", "Prix ramené en HT." if meta["has_vat"] and price_ttc else
                "Prix promo déclaré HT." if not price_ttc else "Colonne TVA absente : CA prévu surévalué si le prix est TTC."))

# ── Export ────────────────────────────────────────────────────────────────────
with tab4:
    T.section("Classeur Excel · valeurs uniquement")
    T.card("<div style='font-size:13px;line-height:1.55'>Onglets : <b>Synthèse</b>, <b>① Avant la promo</b> (forecast vs N-1 par article), "
           "<b>② Après la promo</b> (forecast vs réalisé vs N-1, marge en valeur et taux), <b>Contrôles</b>, <b>Lexique</b>. "
           "Chaque ligne porte Département, Rayon, Famille et Sous-famille pour filtrer ou faire un TCD.</div>"
           "<div class='note'>Généré avec les règles de la barre latérale · aucune formule.</div>")
    if st.button("Générer le classeur", type="primary", key="fp_btn", width="stretch"):
        st.session_state["fp_xlsx"] = export(a, tot, controls, filters, meta, params)
        sup = re.sub(r"[^A-Za-z0-9]+", "_", (filters["supplier"] or "promo").split(" - ")[-1]).strip("_")
        st.session_state["fp_name"] = f"SmartBuyer_Forecast_Promo_{sup}_{datetime.now():%Y%m%d}.xlsx"
    if st.session_state.get("fp_xlsx"):
        st.download_button("⬇️ Télécharger le classeur", data=st.session_state["fp_xlsx"], file_name=st.session_state["fp_name"],
                           mime=XLSX, key="fp_dl", width="stretch")
