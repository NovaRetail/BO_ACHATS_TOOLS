"""
09 ✅ Tasks Tracker
Suivi des tâches acheteurs · Google Sheets · SmartBuyer
"""

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
import plotly.graph_objects as go

# ── CONFIG PAGE ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Tasks Tracker", page_icon="✅", layout="wide")

# ── CHARTE SMARTBUYER ─────────────────────────────────────────────────────────
PL_BG    = "#F2F2F7"
PL_FONT  = "#1C1C1E"
PL_GRID  = "#E5E5EA"
WHITE    = "#FFFFFF"
BORDER   = "#E5E5EA"

BLUE     = "#007AFF"
BLUE_S   = "#E8F1FF"
GREEN    = "#34C759"
GREEN_S  = "#E8FAF0"
ORANGE   = "#FF9500"
ORANGE_S = "#FFF4E0"
RED      = "#FF3B30"
RED_S    = "#FFF0EF"
PURPLE   = "#AF52DE"
PURPLE_S = "#F5EEFF"
GRAY_S   = "#F2F2F7"

STATUTS   = ["À faire", "En cours", "Bloqué", "Terminé"]
PRIORITES = ["Haute", "Moyenne", "Basse"]

STATUT_CFG = {
    "À faire":  {"color": BLUE,   "soft": BLUE_S,   "icon": "●"},
    "En cours": {"color": ORANGE, "soft": ORANGE_S, "icon": "◐"},
    "Bloqué":   {"color": RED,    "soft": RED_S,    "icon": "■"},
    "Terminé":  {"color": GREEN,  "soft": GREEN_S,  "icon": "✓"},
}
PRIORITE_CFG = {
    "Haute":   {"color": RED,    "soft": RED_S},
    "Moyenne": {"color": ORANGE, "soft": ORANGE_S},
    "Basse":   {"color": GREEN,  "soft": GREEN_S},
}

AVATAR_PALETTE = [
    (BLUE,   BLUE_S),
    (PURPLE, PURPLE_S),
    (GREEN,  GREEN_S),
    (ORANGE, ORANGE_S),
    (RED,    RED_S),
]

def avatar_color(name):
    return AVATAR_PALETTE[hash(str(name)) % len(AVATAR_PALETTE)]

def initiales(name):
    parts = str(name or "?").strip().split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return str(name)[:2].upper() if name else "?"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  [data-testid="stAppViewContainer"] {{ background-color: {PL_BG}; }}
  [data-testid="stHeader"]           {{ background-color: {PL_BG}; }}
  .block-container {{ padding-top: 1.5rem; max-width: 1400px; }}

  .kpi-box {{
    background: {WHITE};
    border: 1px solid {BORDER};
    border-radius: 16px;
    padding: 16px 18px;
    text-align: center;
  }}
  .kpi-val {{ font-size: 1.9rem; font-weight: 700; line-height: 1.1; }}
  .kpi-lbl {{ font-size: 0.72rem; color: #636366; margin-top: 3px; }}

  .col-head {{
    border-radius: 10px;
    padding: 7px 12px;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 0.8rem;
    font-weight: 700;
  }}
  .col-badge {{
    font-size: 0.68rem;
    font-weight: 700;
    border-radius: 20px;
    padding: 2px 8px;
    color: {WHITE};
  }}

  .tcard {{
    background: {WHITE};
    border: 1px solid {BORDER};
    border-radius: 14px;
    padding: 13px 14px;
    margin-bottom: 9px;
  }}
  .tcard.late {{
    border-left: 3px solid {RED};
    background: {RED_S};
    border-radius: 0 14px 14px 0;
  }}
  .tcard.done {{ opacity: 0.55; }}

  .tcard-title {{
    font-size: 0.85rem;
    font-weight: 600;
    color: {PL_FONT};
    margin-bottom: 4px;
    line-height: 1.35;
  }}
  .tcard-desc {{
    font-size: 0.74rem;
    color: #636366;
    margin-bottom: 8px;
    line-height: 1.4;
  }}
  .tcard-footer {{
    display: flex;
    align-items: center;
    gap: 5px;
    flex-wrap: wrap;
  }}
  .av {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 22px;
    height: 22px;
    border-radius: 50%;
    font-size: 0.6rem;
    font-weight: 700;
    flex-shrink: 0;
  }}
  .pill {{
    display: inline-flex;
    align-items: center;
    padding: 2px 7px;
    border-radius: 20px;
    font-size: 0.68rem;
    font-weight: 600;
  }}
  .dchip {{
    font-size: 0.68rem;
    color: #8E8E93;
    margin-left: auto;
  }}
  .dchip.late {{ color: {RED}; font-weight: 600; }}
  .empty-col {{
    background: {GRAY_S};
    border-radius: 12px;
    padding: 20px;
    text-align: center;
    font-size: 0.76rem;
    color: #8E8E93;
  }}
</style>
""", unsafe_allow_html=True)

# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=30)
def load_tasks():
    try:
        gc = get_client()
        sh = gc.open_by_key(st.secrets["sheet_id"])
        ws = sh.worksheet("Tâches")
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        if df.empty:
            return _empty_df()
        df["Échéance"] = pd.to_datetime(df["Échéance"], errors="coerce").dt.date
        df["Créé le"]  = pd.to_datetime(df["Créé le"],  errors="coerce").dt.date
        df["ID"]       = df["ID"].astype(str)
        return df
    except Exception as e:
        st.error(f"Erreur connexion Google Sheets : {e}")
        return _empty_df()

def _empty_df():
    return pd.DataFrame(columns=[
        "ID","Titre","Description","Responsable",
        "Statut","Priorité","Échéance","Créé le","Commentaire"
    ])

def save_task(row_data, row_index=None):
    gc = get_client()
    sh = gc.open_by_key(st.secrets["sheet_id"])
    ws = sh.worksheet("Tâches")
    if row_index is None:
        ws.append_row(list(row_data.values()), value_input_option="USER_ENTERED")
    else:
        headers = ws.get_all_values()[0]
        for col_idx, header in enumerate(headers, start=1):
            if header in row_data:
                ws.update_cell(row_index + 1, col_idx, str(row_data[header]))
    load_tasks.clear()

def next_id(df):
    if df.empty or df["ID"].dropna().empty:
        return "T001"
    nums = df["ID"].str.extract(r"(\d+)")[0].dropna().astype(int)
    return f"T{(nums.max() + 1):03d}"

# ── HEADER ────────────────────────────────────────────────────────────────────
c_title, c_btn = st.columns([6, 1])
with c_title:
    st.markdown("## ✅ Tasks Tracker")
    st.caption("Suivi des tâches en temps réel · Équipe Achats · Google Sheets")
with c_btn:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔄 Rafraîchir", use_container_width=True):
        load_tasks.clear()
        st.rerun()

st.divider()

df    = load_tasks()
today = date.today()

# ── KPIs ──────────────────────────────────────────────────────────────────────
total    = len(df)
en_cours = len(df[df["Statut"] == "En cours"])  if not df.empty else 0
bloques  = len(df[df["Statut"] == "Bloqué"])    if not df.empty else 0
termines = len(df[df["Statut"] == "Terminé"])   if not df.empty else 0
retards  = len(
    df[(df["Échéance"] < today) & (df["Statut"] != "Terminé")]
) if not df.empty else 0
taux     = round(termines / total * 100) if total > 0 else 0

k1, k2, k3, k4, k5 = st.columns(5)
for col, val, lbl, color in [
    (k1, total,      "Total tâches", PL_FONT),
    (k2, en_cours,   "En cours",     ORANGE),
    (k3, bloques,    "Bloquées",     RED),
    (k4, f"{taux}%", "Complétées",   GREEN),
    (k5, retards,    "En retard",    RED if retards > 0 else GREEN),
]:
    col.markdown(f"""
    <div class="kpi-box">
      <div class="kpi-val" style="color:{color}">{val}</div>
      <div class="kpi-lbl">{lbl}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── ONGLETS ───────────────────────────────────────────────────────────────────
tab_kanban, tab_form, tab_liste = st.tabs([
    "🗂  Kanban", "➕  Nouvelle tâche", "📋  Liste complète"
])

# ═══════════════════════════════════════════════════════════════════════════════
# KANBAN
# ═══════════════════════════════════════════════════════════════════════════════
with tab_kanban:
    st.markdown("<br>", unsafe_allow_html=True)

    resps_list = sorted(df["Responsable"].dropna().unique().tolist()) if not df.empty else []
    fc1, fc2, _ = st.columns([2, 2, 4])
    f_resp = fc1.selectbox("Responsable", ["Tous"] + resps_list, label_visibility="collapsed")
    f_prio = fc2.selectbox("Priorité", ["Toutes"] + PRIORITES, label_visibility="collapsed")

    df_view = df.copy() if not df.empty else df
    if f_resp != "Tous" and not df_view.empty:
        df_view = df_view[df_view["Responsable"] == f_resp]
    if f_prio != "Toutes" and not df_view.empty:
        df_view = df_view[df_view["Priorité"] == f_prio]

    st.markdown("<br>", unsafe_allow_html=True)
    cols = st.columns(4)

    for col, statut in zip(cols, STATUTS):
        cfg = STATUT_CFG[statut]
        sub = df_view[df_view["Statut"] == statut] if not df_view.empty else pd.DataFrame()

        col.markdown(f"""
        <div class="col-head" style="background:{cfg['soft']};color:{cfg['color']}">
          <span>{cfg['icon']} {statut}</span>
          <span class="col-badge" style="background:{cfg['color']}">{len(sub)}</span>
        </div>""", unsafe_allow_html=True)

        if sub.empty:
            col.markdown('<div class="empty-col">Aucune tâche</div>', unsafe_allow_html=True)
            continue

        for _, row in sub.iterrows():
            ech    = row.get("Échéance", "")
            retard = bool(ech and ech < today and statut != "Terminé")
            css    = "late" if retard else ("done" if statut == "Terminé" else "")
            prio   = row.get("Priorité", "")
            pcfg   = PRIORITE_CFG.get(prio, {"color": BLUE, "soft": BLUE_S})
            resp   = str(row.get("Responsable", ""))
            av_c, av_bg = avatar_color(resp)
            ini    = initiales(resp)
            desc   = str(row.get("Description", "")).strip()
            desc_html = (
                f'<div class="tcard-desc">{desc[:65]}{"…" if len(desc)>65 else ""}</div>'
                if desc else ""
            )
            date_lbl = f"📅 {ech}" + (" 🔴" if retard else "") if ech else ""
            date_cls = "late" if retard else ""

            col.markdown(f"""
            <div class="tcard {css}">
              <div class="tcard-title">{row.get("Titre","")}</div>
              {desc_html}
              <div class="tcard-footer">
                <span class="av" style="background:{av_bg};color:{av_c}">{ini}</span>
                <span class="pill" style="background:{pcfg['soft']};color:{pcfg['color']}">{prio}</span>
                <span class="dchip {date_cls}">{date_lbl}</span>
              </div>
            </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# FORMULAIRE
# ═══════════════════════════════════════════════════════════════════════════════
with tab_form:
    st.markdown("<br>", unsafe_allow_html=True)
    mode = st.radio("Mode", ["Nouvelle tâche", "Modifier une tâche existante"], horizontal=True)

    row_sel   = None
    row_excel = None

    if mode == "Modifier une tâche existante" and not df.empty:
        choices  = df.apply(lambda r: f"{r['ID']} — {r['Titre']}", axis=1).tolist()
        selected = st.selectbox("Tâche à modifier", choices)
        sel_id   = selected.split(" — ")[0]
        row_sel  = df[df["ID"] == sel_id].iloc[0]
        row_excel= df.index.get_loc(df[df["ID"] == sel_id].index[0]) + 1

    st.markdown("<br>", unsafe_allow_html=True)

    with st.form("task_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        titre       = c1.text_input(
            "Titre *",
            value=row_sel["Titre"] if row_sel is not None else "",
            placeholder="Ex : Négociation remise MIPA"
        )
        responsable = c2.text_input(
            "Responsable *",
            value=row_sel["Responsable"] if row_sel is not None else "",
            placeholder="Ex : Grace, Carine, Yves..."
        )
        description = st.text_area(
            "Description",
            value=row_sel["Description"] if row_sel is not None else "",
            height=90,
            placeholder="Contexte, objectif, détails..."
        )
        c3, c4, c5 = st.columns(3)
        statut = c3.selectbox(
            "Statut", STATUTS,
            index=STATUTS.index(row_sel["Statut"])
            if row_sel is not None and row_sel["Statut"] in STATUTS else 0
        )
        priorite = c4.selectbox(
            "Priorité", PRIORITES,
            index=PRIORITES.index(row_sel["Priorité"])
            if row_sel is not None and row_sel["Priorité"] in PRIORITES else 1
        )
        echeance = c5.date_input(
            "Échéance",
            value=row_sel["Échéance"]
            if row_sel is not None and row_sel["Échéance"] else today
        )
        commentaire = st.text_input(
            "Commentaire",
            value=row_sel["Commentaire"] if row_sel is not None else "",
            placeholder="Bloquant, lien utile, note..."
        )

        submitted = st.form_submit_button("💾 Enregistrer", type="primary", use_container_width=True)

        if submitted:
            if not titre or not responsable:
                st.error("Le titre et le responsable sont obligatoires.")
            else:
                task_id  = row_sel["ID"] if row_sel is not None else next_id(df)
                row_data = {
                    "ID":          task_id,
                    "Titre":       titre,
                    "Description": description,
                    "Responsable": responsable,
                    "Statut":      statut,
                    "Priorité":    priorite,
                    "Échéance":    str(echeance),
                    "Créé le":     str(row_sel["Créé le"]) if row_sel is not None else str(today),
                    "Commentaire": commentaire,
                }
                try:
                    save_task(row_data, row_index=row_excel)
                    st.success(f"✅ Tâche **{task_id}** enregistrée avec succès !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur sauvegarde : {e}")

# ═══════════════════════════════════════════════════════════════════════════════
# LISTE COMPLÈTE
# ═══════════════════════════════════════════════════════════════════════════════
with tab_liste:
    st.markdown("<br>", unsafe_allow_html=True)

    if df.empty:
        st.info("Aucune tâche enregistrée.")
    else:
        resps  = sorted(df["Responsable"].dropna().unique().tolist())
        f1, f2, f3, f4 = st.columns(4)
        f_resp2   = f1.multiselect("Responsable", resps,     default=resps)
        f_stat2   = f2.multiselect("Statut",      STATUTS,   default=STATUTS)
        f_prio2   = f3.multiselect("Priorité",    PRIORITES, default=PRIORITES)
        f_retard2 = f4.checkbox("Retards uniquement")

        df_filt = df[
            df["Responsable"].isin(f_resp2) &
            df["Statut"].isin(f_stat2) &
            df["Priorité"].isin(f_prio2)
        ].copy()

        if f_retard2:
            df_filt = df_filt[
                (df_filt["Échéance"] < today) & (df_filt["Statut"] != "Terminé")
            ]

        df_filt["⚠️"] = df_filt["Échéance"].apply(
            lambda e: "🔴" if (e and e < today) else ""
        )

        st.dataframe(
            df_filt[["ID","Titre","Responsable","Statut","Priorité","Échéance","⚠️","Commentaire"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Échéance": st.column_config.DateColumn("Échéance", format="DD/MM/YYYY"),
            }
        )

        nb_retards = len(df_filt[
            (df_filt["Échéance"] < today) & (df_filt["Statut"] != "Terminé")
        ]) if not df_filt.empty else 0
        st.caption(
            f"{len(df_filt)} tâche(s) · "
            f"{len(df_filt[df_filt['Statut']=='Terminé'])} terminée(s) · "
            f"{nb_retards} en retard"
        )

        if not df_filt.empty and len(df_filt) > 1:
            st.markdown("<br>", unsafe_allow_html=True)
            agg = df_filt.groupby(["Responsable","Statut"]).size().reset_index(name="n")
            fig = go.Figure()
            for statut in STATUTS:
                sub = agg[agg["Statut"] == statut]
                if sub.empty:
                    continue
                fig.add_trace(go.Bar(
                    name=statut,
                    x=sub["Responsable"],
                    y=sub["n"],
                    marker_color=STATUT_CFG[statut]["color"],
                    marker_line_width=0,
                ))
            fig.update_layout(
                barmode="stack",
                paper_bgcolor=WHITE,
                plot_bgcolor=WHITE,
                height=240,
                margin=dict(l=10, r=10, t=10, b=10),
                legend=dict(orientation="h", y=-0.3, font_size=11, font_color=PL_FONT),
                font=dict(
                    family="SF Pro Display, -apple-system, sans-serif",
                    color=PL_FONT, size=12
                ),
                xaxis=dict(gridcolor=PL_GRID, linecolor=PL_GRID),
                yaxis=dict(gridcolor=PL_GRID, tickformat="d"),
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
