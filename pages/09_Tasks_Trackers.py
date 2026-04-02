"""
Module 08 — Task Tracker
Suivi des tâches acheteurs connecté à Google Sheets
"""

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date, datetime
import plotly.graph_objects as go

# ── CONFIG PAGE ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Task Tracker", page_icon="✅", layout="wide")

# ── CONSTANTES VISUELLES ─────────────────────────────────────────────────────
PL_BG    = "#F2F2F7"
PL_FONT  = "#1C1C1E"
PL_GRID  = "#E5E5EA"
BLUE     = "#007AFF"
GREEN    = "#34C759"
ORANGE   = "#FF9500"
RED      = "#FF3B30"
PURPLE   = "#AF52DE"

ACHETEURS = ["GB", "CK", "AC"]
STATUTS   = ["À faire", "En cours", "Bloqué", "Terminé"]
PRIORITES = ["Haute", "Moyenne", "Basse"]

STATUT_COLORS = {
    "À faire":  BLUE,
    "En cours": ORANGE,
    "Bloqué":   RED,
    "Terminé":  GREEN,
}

PRIORITE_COLORS = {
    "Haute":   RED,
    "Moyenne": ORANGE,
    "Basse":   GREEN,
}

# ── STYLES CSS ───────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  [data-testid="stAppViewContainer"] {{ background-color: {PL_BG}; }}
  [data-testid="stHeader"] {{ background-color: {PL_BG}; }}
  .block-container {{ padding-top: 1.5rem; max-width: 1400px; }}

  .card {{
    background: white;
    border-radius: 14px;
    padding: 16px 18px;
    margin-bottom: 10px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    border-left: 4px solid {BLUE};
  }}
  .card.bloque  {{ border-left-color: {RED}; }}
  .card.en_cours{{ border-left-color: {ORANGE}; }}
  .card.termine {{ border-left-color: {GREEN}; opacity: 0.7; }}
  .card.retard  {{ border-left-color: {RED}; background: #FFF2F2; }}

  .task-title {{ font-weight: 600; font-size: 0.95rem; color: {PL_FONT}; }}
  .task-meta  {{ font-size: 0.78rem; color: #636366; margin-top: 4px; }}
  .badge {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
    color: white;
  }}
  .kpi-box {{
    background: white;
    border-radius: 14px;
    padding: 18px 20px;
    text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }}
  .kpi-val  {{ font-size: 2rem; font-weight: 700; }}
  .kpi-lbl  {{ font-size: 0.78rem; color: #636366; margin-top: 2px; }}
  .section-title {{
    font-size: 1.05rem;
    font-weight: 700;
    color: {PL_FONT};
    margin: 18px 0 10px 2px;
    letter-spacing: -0.3px;
  }}
</style>
""", unsafe_allow_html=True)

# ── CONNEXION GOOGLE SHEETS ───────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def get_gsheet_client():
    """Crée le client gspread à partir des secrets Streamlit."""
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES,
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=30)
def load_tasks():
    """Charge les tâches depuis Google Sheets (cache 30 s)."""
    try:
        gc     = get_gsheet_client()
        sh     = gc.open_by_key(st.secrets["sheet_id"])
        ws     = sh.worksheet("Tâches")
        data   = ws.get_all_records()
        df     = pd.DataFrame(data)
        if df.empty:
            return _empty_df()
        df["Échéance"]  = pd.to_datetime(df["Échéance"], errors="coerce").dt.date
        df["Créé le"]   = pd.to_datetime(df["Créé le"],  errors="coerce").dt.date
        df["ID"]        = df["ID"].astype(str)
        return df
    except Exception as e:
        st.error(f"Erreur connexion Google Sheets : {e}")
        return _empty_df()

def _empty_df():
    return pd.DataFrame(columns=[
        "ID", "Titre", "Description", "Acheteur", "Statut",
        "Priorité", "Échéance", "Créé le", "Commentaire"
    ])

def save_task(row_data: dict, row_index: int = None):
    """Ajoute ou met à jour une ligne dans Google Sheets."""
    gc  = get_gsheet_client()
    sh  = gc.open_by_key(st.secrets["sheet_id"])
    ws  = sh.worksheet("Tâches")
    if row_index is None:
        # Nouvelle tâche
        ws.append_row(list(row_data.values()), value_input_option="USER_ENTERED")
    else:
        # Mise à jour (row_index = ligne Excel 1-indexed, headers en ligne 1)
        all_data = ws.get_all_values()
        headers  = all_data[0]
        for col_idx, header in enumerate(headers, start=1):
            if header in row_data:
                ws.update_cell(row_index + 1, col_idx, str(row_data[header]))
    load_tasks.clear()

def delete_task(row_index: int):
    """Supprime une ligne dans Google Sheets."""
    gc  = get_gsheet_client()
    sh  = gc.open_by_key(st.secrets["sheet_id"])
    ws  = sh.worksheet("Tâches")
    ws.delete_rows(row_index + 1)
    load_tasks.clear()

def next_id(df):
    if df.empty or df["ID"].dropna().empty:
        return "T001"
    nums = df["ID"].str.extract(r"(\d+)")[0].dropna().astype(int)
    return f"T{(nums.max() + 1):03d}"

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("## ✅ Task Tracker — Équipe Achats")
st.caption("Suivi des tâches en temps réel · Synchronisé avec Google Sheets")
st.divider()

df = load_tasks()
today = date.today()

# ── KPIs ──────────────────────────────────────────────────────────────────────
total      = len(df)
en_cours   = len(df[df["Statut"] == "En cours"])  if not df.empty else 0
bloques    = len(df[df["Statut"] == "Bloqué"])     if not df.empty else 0
termines   = len(df[df["Statut"] == "Terminé"])    if not df.empty else 0
retards    = len(df[(df["Échéance"] < today) & (df["Statut"] != "Terminé")]) if not df.empty else 0
taux_compl = round(termines / total * 100, 0) if total > 0 else 0

k1, k2, k3, k4, k5 = st.columns(5)
kpi_data = [
    (k1, str(total),       "Total tâches",      PL_FONT),
    (k2, str(en_cours),    "En cours",           ORANGE),
    (k3, str(bloques),     "Bloquées 🔴",        RED),
    (k4, f"{taux_compl}%", "Taux complétion",    GREEN),
    (k5, str(retards),     "En retard ⚠️",       RED if retards > 0 else GREEN),
]
for col, val, lbl, color in kpi_data:
    col.markdown(f"""
    <div class="kpi-box">
      <div class="kpi-val" style="color:{color}">{val}</div>
      <div class="kpi-lbl">{lbl}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("")

# ── GRAPHIQUE RÉPARTITION PAR ACHETEUR ───────────────────────────────────────
if not df.empty:
    col_chart, col_space = st.columns([2, 1])
    with col_chart:
        st.markdown('<div class="section-title">Répartition par acheteur</div>', unsafe_allow_html=True)
        agg = df.groupby(["Acheteur", "Statut"]).size().reset_index(name="n")
        colors_map = [BLUE, ORANGE, RED, GREEN]
        fig = go.Figure()
        for i, statut in enumerate(STATUTS):
            sub = agg[agg["Statut"] == statut]
            fig.add_trace(go.Bar(
                name=statut,
                x=sub["Acheteur"],
                y=sub["n"],
                marker_color=list(STATUT_COLORS.values())[i],
            ))
        fig.update_layout(
            barmode="stack",
            paper_bgcolor="white",
            plot_bgcolor="white",
            height=220,
            margin=dict(l=10, r=10, t=10, b=10),
            legend=dict(orientation="h", y=-0.25, font_size=11),
            font=dict(family="SF Pro Display, -apple-system, sans-serif", color=PL_FONT),
            xaxis=dict(gridcolor=PL_GRID),
            yaxis=dict(gridcolor=PL_GRID, tickformat="d"),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

st.divider()

# ── ONGLETS : KANBAN + FORMULAIRE ────────────────────────────────────────────
tab_kanban, tab_form, tab_liste = st.tabs(["🗂 Kanban", "➕ Ajouter / Modifier", "📋 Liste complète"])

# ── TAB KANBAN ────────────────────────────────────────────────────────────────
with tab_kanban:
    acheteur_filter = st.radio(
        "Acheteur", ["Tous"] + ACHETEURS,
        horizontal=True, label_visibility="collapsed"
    )
    df_view = df if acheteur_filter == "Tous" else df[df["Acheteur"] == acheteur_filter]

    cols = st.columns(4)
    for col, statut in zip(cols, STATUTS):
        color = STATUT_COLORS[statut]
        sub   = df_view[df_view["Statut"] == statut] if not df_view.empty else pd.DataFrame()
        col.markdown(f"""
        <div style="background:{color}18;border-radius:12px;padding:10px 12px;margin-bottom:8px;">
          <span style="font-weight:700;color:{color};font-size:0.9rem">{statut}</span>
          <span style="float:right;background:{color};color:white;
                border-radius:20px;padding:1px 8px;font-size:0.75rem">{len(sub)}</span>
        </div>""", unsafe_allow_html=True)

        if sub.empty:
            col.caption("Aucune tâche")
        else:
            for _, row in sub.iterrows():
                echeance = row.get("Échéance", "")
                retard   = echeance and echeance < today and statut != "Terminé"
                css_card = "retard" if retard else statut.lower().replace(" ", "_").replace("é", "e")
                prio_col = PRIORITE_COLORS.get(row.get("Priorité", ""), BLUE)
                badge_html = f'<span class="badge" style="background:{prio_col}">{row.get("Priorité","")}</span>'
                acheteur_b = f'<span class="badge" style="background:{PURPLE};margin-left:4px">{row.get("Acheteur","")}</span>'
                date_str   = f"📅 {echeance}" if echeance else ""
                retard_str = " 🔴 <b>EN RETARD</b>" if retard else ""
                col.markdown(f"""
                <div class="card {css_card}">
                  <div class="task-title">{row.get("Titre","")}</div>
                  <div class="task-meta">{badge_html}{acheteur_b}</div>
                  <div class="task-meta">{date_str}{retard_str}</div>
                  <div class="task-meta" style="margin-top:5px;color:#48484A">{row.get("Description","")[:80]}</div>
                </div>""", unsafe_allow_html=True)

# ── TAB FORMULAIRE ────────────────────────────────────────────────────────────
with tab_form:
    mode = st.radio("Mode", ["Nouvelle tâche", "Modifier une tâche existante"], horizontal=True)

    if mode == "Modifier une tâche existante" and not df.empty:
        task_choices = df.apply(lambda r: f"{r['ID']} — {r['Titre']}", axis=1).tolist()
        selected     = st.selectbox("Sélectionner la tâche", task_choices)
        sel_id       = selected.split(" — ")[0]
        row_sel      = df[df["ID"] == sel_id].iloc[0]
        row_excel    = df.index.get_loc(df[df["ID"] == sel_id].index[0]) + 1  # +1 pour header
    else:
        row_sel  = None
        row_excel = None

    with st.form("task_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        titre   = c1.text_input("Titre *", value=row_sel["Titre"] if row_sel is not None else "")
        acheteur= c2.selectbox("Acheteur *", ACHETEURS,
                               index=ACHETEURS.index(row_sel["Acheteur"]) if row_sel is not None and row_sel["Acheteur"] in ACHETEURS else 0)

        description = st.text_area("Description", value=row_sel["Description"] if row_sel is not None else "", height=80)

        c3, c4, c5 = st.columns(3)
        statut   = c3.selectbox("Statut", STATUTS,
                                index=STATUTS.index(row_sel["Statut"]) if row_sel is not None and row_sel["Statut"] in STATUTS else 0)
        priorite = c4.selectbox("Priorité", PRIORITES,
                                index=PRIORITES.index(row_sel["Priorité"]) if row_sel is not None and row_sel["Priorité"] in PRIORITES else 1)
        echeance = c5.date_input("Échéance",
                                 value=row_sel["Échéance"] if row_sel is not None and row_sel["Échéance"] else today)

        commentaire = st.text_input("Commentaire", value=row_sel["Commentaire"] if row_sel is not None else "")

        submitted = st.form_submit_button("💾 Enregistrer", type="primary", use_container_width=True)

        if submitted:
            if not titre:
                st.error("Le titre est obligatoire.")
            else:
                task_id = row_sel["ID"] if row_sel is not None else next_id(df)
                row_data = {
                    "ID":          task_id,
                    "Titre":       titre,
                    "Description": description,
                    "Acheteur":    acheteur,
                    "Statut":      statut,
                    "Priorité":    priorite,
                    "Échéance":    str(echeance),
                    "Créé le":     str(row_sel["Créé le"]) if row_sel is not None else str(today),
                    "Commentaire": commentaire,
                }
                try:
                    save_task(row_data, row_index=row_excel)
                    st.success(f"✅ Tâche {task_id} enregistrée avec succès !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur lors de la sauvegarde : {e}")

# ── TAB LISTE COMPLÈTE ────────────────────────────────────────────────────────
with tab_liste:
    if df.empty:
        st.info("Aucune tâche enregistrée.")
    else:
        f1, f2, f3 = st.columns(3)
        f_ach  = f1.multiselect("Acheteur", ACHETEURS, default=ACHETEURS)
        f_stat = f2.multiselect("Statut",   STATUTS,   default=STATUTS)
        f_prio = f3.multiselect("Priorité", PRIORITES, default=PRIORITES)

        df_filt = df[
            df["Acheteur"].isin(f_ach) &
            df["Statut"].isin(f_stat) &
            df["Priorité"].isin(f_prio)
        ]

        # Colonne retard
        df_filt = df_filt.copy()
        df_filt["⚠️"] = df_filt["Échéance"].apply(
            lambda e: "🔴" if (e and e < today) else ""
        )

        st.dataframe(
            df_filt[["ID","Titre","Acheteur","Statut","Priorité","Échéance","⚠️","Commentaire"]],
            use_container_width=True,
            hide_index=True,
        )
        st.caption(f"{len(df_filt)} tâche(s) affichée(s)")
