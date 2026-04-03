"""
Module 09 — Task Tracker (v2)
Suivi des tâches acheteurs · Design premium · Google Sheets
"""

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
import plotly.graph_objects as go

# ── CONFIG PAGE ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Task Tracker", page_icon="✅", layout="wide")

# ── PALETTE PREMIUM ───────────────────────────────────────────────────────────
BG        = "#F7F7F8"
WHITE     = "#FFFFFF"
BORDER    = "#E8E8EC"
TEXT_PRI  = "#1A1A2E"
TEXT_SEC  = "#6B6B80"
BLUE      = "#4361EE"
BLUE_SOFT = "#EEF1FD"
GREEN     = "#2DC653"
GREEN_SOFT= "#E8F8ED"
ORANGE    = "#F77F00"
ORANGE_SOFT="#FEF3E2"
RED       = "#E63946"
RED_SOFT  = "#FDECEA"
PURPLE    = "#7B2FBE"
PURPLE_SOFT="#F3EAFA"
GRAY_SOFT = "#F0F0F3"

STATUTS   = ["À faire", "En cours", "Bloqué", "Terminé"]
PRIORITES = ["Haute", "Moyenne", "Basse"]

STATUT_CFG = {
    "À faire":  {"color": BLUE,   "soft": BLUE_SOFT,   "icon": "○"},
    "En cours": {"color": ORANGE, "soft": ORANGE_SOFT, "icon": "◑"},
    "Bloqué":   {"color": RED,    "soft": RED_SOFT,    "icon": "✕"},
    "Terminé":  {"color": GREEN,  "soft": GREEN_SOFT,  "icon": "✓"},
}
PRIORITE_CFG = {
    "Haute":   {"color": RED,    "soft": RED_SOFT},
    "Moyenne": {"color": ORANGE, "soft": ORANGE_SOFT},
    "Basse":   {"color": GREEN,  "soft": GREEN_SOFT},
}

def avatar_color(name):
    colors = [
        (BLUE, BLUE_SOFT), (PURPLE, PURPLE_SOFT),
        (GREEN, GREEN_SOFT), (ORANGE, ORANGE_SOFT), (RED, RED_SOFT),
    ]
    return colors[hash(name or "?") % len(colors)]

def initiales(name):
    parts = (name or "?").strip().split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return name[:2].upper() if name else "?"

# ── CSS PREMIUM ───────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  [data-testid="stAppViewContainer"] {{ background: {BG}; }}
  [data-testid="stHeader"] {{ background: {BG}; }}
  .block-container {{ padding-top: 1.8rem; max-width: 1400px; }}
  div[data-testid="stTabs"] button {{ font-size: 0.82rem; font-weight: 600; }}

  .kpi-card {{
    background: {WHITE};
    border: 1px solid {BORDER};
    border-radius: 16px;
    padding: 18px 20px;
    text-align: center;
  }}
  .kpi-val {{ font-size: 2rem; font-weight: 700; line-height: 1.1; }}
  .kpi-lbl {{ font-size: 0.73rem; color: {TEXT_SEC}; margin-top: 4px; letter-spacing: 0.3px; }}

  .col-header {{
    border-radius: 10px;
    padding: 7px 12px;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 0.82rem;
    font-weight: 700;
    letter-spacing: 0.2px;
  }}
  .col-count {{
    font-size: 0.72rem;
    font-weight: 700;
    border-radius: 20px;
    padding: 2px 8px;
    color: white;
  }}

  .task-card {{
    background: {WHITE};
    border: 1px solid {BORDER};
    border-radius: 12px;
    padding: 12px 14px;
    margin-bottom: 8px;
    transition: box-shadow 0.15s;
  }}
  .task-card:hover {{ box-shadow: 0 4px 12px rgba(0,0,0,0.07); }}
  .task-card.retard {{ border-left: 3px solid {RED}; background: {RED_SOFT}; }}
  .task-card.termine {{ opacity: 0.6; }}

  .task-title {{
    font-weight: 600;
    font-size: 0.88rem;
    color: {TEXT_PRI};
    margin-bottom: 6px;
    line-height: 1.3;
  }}
  .task-desc {{
    font-size: 0.76rem;
    color: {TEXT_SEC};
    margin-bottom: 7px;
    line-height: 1.4;
  }}
  .task-footer {{
    display: flex;
    align-items: center;
    gap: 6px;
    flex-wrap: wrap;
  }}
  .pill {{
    display: inline-flex;
    align-items: center;
    gap: 3px;
    padding: 2px 8px;
    border-radius: 20px;
    font-size: 0.69rem;
    font-weight: 600;
  }}
  .avatar {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 22px;
    height: 22px;
    border-radius: 50%;
    font-size: 0.62rem;
    font-weight: 700;
    flex-shrink: 0;
  }}
  .date-chip {{
    font-size: 0.69rem;
    color: {TEXT_SEC};
    margin-left: auto;
  }}
  .date-chip.retard {{ color: {RED}; font-weight: 600; }}

  .section-label {{
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.8px;
    color: {TEXT_SEC};
    text-transform: uppercase;
    margin-bottom: 10px;
    margin-top: 4px;
  }}
  .page-title {{
    font-size: 1.4rem;
    font-weight: 700;
    color: {TEXT_PRI};
    letter-spacing: -0.5px;
  }}
  .page-sub {{
    font-size: 0.78rem;
    color: {TEXT_SEC};
    margin-top: 2px;
  }}
  .empty-col {{
    background: {GRAY_SOFT};
    border-radius: 10px;
    padding: 18px;
    text-align: center;
    font-size: 0.76rem;
    color: {TEXT_SEC};
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
        "ID","Titre","Description","Responsable","Statut",
        "Priorité","Échéance","Créé le","Commentaire"
    ])

def save_task(row_data, row_index=None):
    gc = get_client()
    sh = gc.open_by_key(st.secrets["sheet_id"])
    ws = sh.worksheet("Tâches")
    if row_index is None:
        ws.append_row(list(row_data.values()), value_input_option="USER_ENTERED")
    else:
        all_data = ws.get_all_values()
        headers  = all_data[0]
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
col_title, col_refresh = st.columns([5, 1])
with col_title:
    st.markdown(f'<div class="page-title">✅ Task Tracker</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-sub">Suivi des tâches · Équipe Achats · Synchronisé Google Sheets</div>', unsafe_allow_html=True)
with col_refresh:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔄 Rafraîchir", use_container_width=True):
        load_tasks.clear()
        st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

df    = load_tasks()
today = date.today()

# ── KPIs ──────────────────────────────────────────────────────────────────────
total      = len(df)
en_cours   = len(df[df["Statut"] == "En cours"])  if not df.empty else 0
bloques    = len(df[df["Statut"] == "Bloqué"])     if not df.empty else 0
termines   = len(df[df["Statut"] == "Terminé"])    if not df.empty else 0
retards    = len(df[(df["Échéance"] < today) & (df["Statut"] != "Terminé")]) if not df.empty else 0
taux       = round(termines / total * 100) if total > 0 else 0

k1,k2,k3,k4,k5 = st.columns(5)
for col, val, lbl, color in [
    (k1, total,        "Total tâches",     TEXT_PRI),
    (k2, en_cours,     "En cours",         ORANGE),
    (k3, bloques,      "Bloquées",         RED),
    (k4, f"{taux}%",   "Complétées",       GREEN),
    (k5, retards,      "En retard",        RED if retards > 0 else GREEN),
]:
    col.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-val" style="color:{color}">{val}</div>
      <div class="kpi-lbl">{lbl}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── ONGLETS ───────────────────────────────────────────────────────────────────
tab_kanban, tab_form, tab_liste = st.tabs(["🗂  Kanban", "➕  Nouvelle tâche", "📋  Liste complète"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB KANBAN
# ═══════════════════════════════════════════════════════════════════════════════
with tab_kanban:
    st.markdown("<br>", unsafe_allow_html=True)

    # Filtres
    responsables = ["Tous"] + sorted(df["Responsable"].dropna().unique().tolist()) if not df.empty else ["Tous"]
    f1, f2, _ = st.columns([2, 2, 4])
    filtre_resp = f1.selectbox("Responsable", responsables, label_visibility="collapsed")
    filtre_prio = f2.selectbox("Priorité", ["Toutes"] + PRIORITES, label_visibility="collapsed")

    df_view = df.copy() if not df.empty else df
    if filtre_resp != "Tous" and not df_view.empty:
        df_view = df_view[df_view["Responsable"] == filtre_resp]
    if filtre_prio != "Toutes" and not df_view.empty:
        df_view = df_view[df_view["Priorité"] == filtre_prio]

    st.markdown("<br>", unsafe_allow_html=True)
    cols = st.columns(4)

    for col, statut in zip(cols, STATUTS):
        cfg   = STATUT_CFG[statut]
        sub   = df_view[df_view["Statut"] == statut] if not df_view.empty else pd.DataFrame()

        col.markdown(f"""
        <div class="col-header" style="background:{cfg['soft']};color:{cfg['color']}">
          <span>{cfg['icon']} {statut}</span>
          <span class="col-count" style="background:{cfg['color']}">{len(sub)}</span>
        </div>""", unsafe_allow_html=True)

        if sub.empty:
            col.markdown(f'<div class="empty-col">Aucune tâche</div>', unsafe_allow_html=True)
        else:
            for _, row in sub.iterrows():
                echeance   = row.get("Échéance", "")
                retard     = bool(echeance and echeance < today and statut != "Terminé")
                css_extra  = "retard" if retard else ("termine" if statut == "Terminé" else "")
                prio       = row.get("Priorité", "")
                pcfg       = PRIORITE_CFG.get(prio, {"color": BLUE, "soft": BLUE_SOFT})
                resp       = str(row.get("Responsable", ""))
                av_color, av_bg = avatar_color(resp)
                ini        = initiales(resp)
                date_str   = f"📅 {echeance}" if echeance else ""
                date_cls   = "retard" if retard else ""
                desc       = str(row.get("Description",""))[:70]

                col.markdown(f"""
                <div class="task-card {css_extra}">
                  <div class="task-title">{row.get("Titre","")}</div>
                  {"<div class='task-desc'>" + desc + ("…" if len(str(row.get("Description",""))) > 70 else "") + "</div>" if desc else ""}
                  <div class="task-footer">
                    <span class="avatar" style="background:{av_bg};color:{av_color}">{ini}</span>
                    <span class="pill" style="background:{pcfg['soft']};color:{pcfg['color']}">{prio}</span>
                    <span class="date-chip {date_cls}">{date_str}{"  🔴" if retard else ""}</span>
                  </div>
                </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB FORMULAIRE
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
        titre       = c1.text_input("Titre *", value=row_sel["Titre"] if row_sel is not None else "", placeholder="Ex : Négociation remise MIPA")
        responsable = c2.text_input("Responsable *", value=row_sel["Responsable"] if row_sel is not None else "", placeholder="Ex : Grace, Carine, Yves...")

        description = st.text_area("Description", value=row_sel["Description"] if row_sel is not None else "", height=90, placeholder="Contexte, objectif, détails...")

        c3, c4, c5 = st.columns(3)
        statut   = c3.selectbox("Statut", STATUTS,
                                index=STATUTS.index(row_sel["Statut"]) if row_sel is not None and row_sel["Statut"] in STATUTS else 0)
        priorite = c4.selectbox("Priorité", PRIORITES,
                                index=PRIORITES.index(row_sel["Priorité"]) if row_sel is not None and row_sel["Priorité"] in PRIORITES else 1)
        echeance = c5.date_input("Échéance", value=row_sel["Échéance"] if row_sel is not None and row_sel["Échéance"] else today)

        commentaire = st.text_input("Commentaire / Note", value=row_sel["Commentaire"] if row_sel is not None else "", placeholder="Bloquant, lien utile, remarque...")

        submitted = st.form_submit_button("💾 Enregistrer la tâche", type="primary", use_container_width=True)

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
# TAB LISTE
# ═══════════════════════════════════════════════════════════════════════════════
with tab_liste:
    st.markdown("<br>", unsafe_allow_html=True)
    if df.empty:
        st.info("Aucune tâche enregistrée.")
    else:
        f1, f2, f3, f4 = st.columns(4)
        resps    = sorted(df["Responsable"].dropna().unique().tolist())
        f_resp   = f1.multiselect("Responsable", resps, default=resps)
        f_stat   = f2.multiselect("Statut",      STATUTS,   default=STATUTS)
        f_prio   = f3.multiselect("Priorité",    PRIORITES, default=PRIORITES)
        f_retard = f4.checkbox("Retards uniquement")

        df_filt = df[
            df["Responsable"].isin(f_resp) &
            df["Statut"].isin(f_stat) &
            df["Priorité"].isin(f_prio)
        ].copy()

        if f_retard:
            df_filt = df_filt[(df_filt["Échéance"] < today) & (df_filt["Statut"] != "Terminé")]

        df_filt["⚠️"] = df_filt["Échéance"].apply(
            lambda e: "🔴" if (e and e < today) else "")

        st.dataframe(
            df_filt[["ID","Titre","Responsable","Statut","Priorité","Échéance","⚠️","Commentaire"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Échéance": st.column_config.DateColumn("Échéance", format="DD/MM/YYYY"),
            }
        )
        st.caption(f"{len(df_filt)} tâche(s) · {len(df_filt[df_filt['Statut']=='Terminé'])} terminée(s)")

        # Graphique répartition
        if len(df_filt) > 0:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="section-label">Répartition par responsable</div>', unsafe_allow_html=True)
            agg = df_filt.groupby(["Responsable","Statut"]).size().reset_index(name="n")
            fig = go.Figure()
            for statut in STATUTS:
                sub = agg[agg["Statut"] == statut]
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
                legend=dict(orientation="h", y=-0.3, font_size=11),
                font=dict(family="-apple-system, sans-serif", color=TEXT_PRI, size=12),
                xaxis=dict(gridcolor=BORDER, linecolor=BORDER),
                yaxis=dict(gridcolor=BORDER, tickformat="d"),
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
