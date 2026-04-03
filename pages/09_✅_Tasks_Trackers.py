"""
09 ✅ Tasks Tracker
Style Trello · Carte cliquable · Archives · SmartBuyer
"""

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
import plotly.graph_objects as go

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Tasks Tracker", page_icon="✅", layout="wide")

# ── CHARTE SMARTBUYER ─────────────────────────────────────────────────────────
PL_BG    = "#F2F2F7"
PL_FONT  = "#1C1C1E"
PL_GRID  = "#E5E5EA"
WHITE    = "#FFFFFF"
BORDER   = "#E5E5EA"
BLUE     = "#007AFF"; BLUE_S   = "#E8F1FF"
GREEN    = "#34C759"; GREEN_S  = "#E8FAF0"
ORANGE   = "#FF9500"; ORANGE_S = "#FFF4E0"
RED      = "#FF3B30"; RED_S    = "#FFF0EF"
PURPLE   = "#AF52DE"; PURPLE_S = "#F5EEFF"
GRAY_S   = "#F2F2F7"

STATUTS      = ["À faire", "En cours", "Bloqué", "Terminé"]
PRIORITES    = ["Haute", "Moyenne", "Basse"]
ARCHIVE_DAYS = 30

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
    (BLUE, BLUE_S), (PURPLE, PURPLE_S), (GREEN, GREEN_S),
    (ORANGE, ORANGE_S), (RED, RED_S),
]

def avatar_color(name):
    return AVATAR_PALETTE[hash(str(name)) % len(AVATAR_PALETTE)]

def initiales(name):
    parts = str(name or "?").strip().split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return str(name)[:2].upper() if name else "?"

def is_archived(row, today):
    if row.get("Statut") != "Terminé":
        return False
    for key in ["Échéance", "Créé le"]:
        val = row.get(key)
        if val and isinstance(val, date):
            return (today - val).days > ARCHIVE_DAYS
    return False

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  [data-testid="stAppViewContainer"] {{ background-color: {PL_BG}; }}
  [data-testid="stHeader"]           {{ background-color: {PL_BG}; }}
  .block-container {{ padding-top: 1.5rem; max-width: 1400px; }}

  /* KPIs */
  .kpi-box {{
    background: {WHITE}; border: 1px solid {BORDER};
    border-radius: 16px; padding: 16px 18px; text-align: center;
  }}
  .kpi-val {{ font-size: 1.9rem; font-weight: 700; line-height: 1.1; }}
  .kpi-lbl {{ font-size: 0.72rem; color: #636366; margin-top: 3px; }}

  /* Colonne Kanban */
  .kol {{ background: {GRAY_S}; border-radius: 14px; padding: 10px; }}
  .kol-head {{
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 10px; padding: 0 2px;
  }}
  .kol-name {{ font-size: 0.8rem; font-weight: 700; display: flex; align-items: center; gap: 6px; }}
  .kol-dot  {{ width: 8px; height: 8px; border-radius: 50%; display: inline-block; }}
  .kol-badge {{ font-size: 0.68rem; font-weight: 700; border-radius: 20px; padding: 2px 8px; color: {WHITE}; }}

  /* ── Carte = bouton Streamlit déguisé ── */
  div[data-testid="stButton"].card-btn > button {{
    background: {WHITE} !important;
    border: 0.5px solid {BORDER} !important;
    border-radius: 12px !important;
    padding: 12px 13px !important;
    margin-bottom: 8px !important;
    width: 100% !important;
    text-align: left !important;
    cursor: pointer !important;
    color: {PL_FONT} !important;
    font-size: 0.84rem !important;
    font-weight: 400 !important;
    line-height: 1.5 !important;
    white-space: pre-wrap !important;
    height: auto !important;
    min-height: unset !important;
    transition: border-color 0.12s, box-shadow 0.12s !important;
  }}
  div[data-testid="stButton"].card-btn > button:hover {{
    border-color: #C7C7CC !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07) !important;
    background: {WHITE} !important;
  }}
  div[data-testid="stButton"].card-btn > button:focus {{
    box-shadow: none !important;
    outline: none !important;
  }}
  div[data-testid="stButton"].card-btn.sel > button {{
    border: 1.5px solid {BLUE} !important;
    background: {BLUE_S} !important;
  }}
  div[data-testid="stButton"].card-btn.late > button {{
    border-left: 3px solid {RED} !important;
    border-radius: 0 12px 12px 0 !important;
  }}
  div[data-testid="stButton"].card-btn.done > button {{
    opacity: 0.5 !important;
  }}

  /* Colonne vide */
  .empty-kol {{
    border: 0.5px dashed {BORDER}; border-radius: 10px;
    padding: 18px; text-align: center;
    font-size: 0.76rem; color: #8E8E93;
  }}

  /* Archive banner */
  .arch-banner {{
    background: {GREEN_S}; border: 0.5px solid #A3D9B1;
    border-radius: 10px; padding: 9px 14px;
    font-size: 0.78rem; color: #1A6B35; margin-bottom: 12px;
  }}

  /* Panneau édition */
  .edit-header {{
    font-size: 0.88rem; font-weight: 700;
    color: {PL_FONT}; margin-bottom: 6px;
    padding: 12px 16px;
    background: {WHITE}; border: 1px solid {BORDER};
    border-radius: 14px 14px 0 0;
    border-bottom: none;
  }}
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ─────────────────────────────────────────────────────────────
for k, v in {"edit_id": None, "view_mode": "active"}.items():
    if k not in st.session_state:
        st.session_state[k] = v

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
        gc   = get_client()
        sh   = gc.open_by_key(st.secrets["sheet_id"])
        ws   = sh.worksheet("Tâches")
        data = ws.get_all_records()
        df   = pd.DataFrame(data)
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

def delete_task(row_index):
    gc = get_client()
    sh = gc.open_by_key(st.secrets["sheet_id"])
    ws = sh.worksheet("Tâches")
    ws.delete_rows(row_index + 1)
    load_tasks.clear()

def next_id(df):
    if df.empty or df["ID"].dropna().empty:
        return "T001"
    nums = df["ID"].str.extract(r"(\d+)")[0].dropna().astype(int)
    return f"T{(nums.max() + 1):03d}"

def card_label(row, today):
    """Texte multi-ligne affiché dans le bouton-carte."""
    prio   = row.get("Priorité", "")
    resp   = str(row.get("Responsable", ""))
    ini    = initiales(resp)
    ech    = row.get("Échéance", "")
    retard = bool(ech and ech < today and row.get("Statut") != "Terminé")
    desc   = str(row.get("Description", "")).strip()
    desc_s = f"\n{desc[:55]}{'…' if len(desc)>55 else ''}" if desc else ""
    date_s = f"\n📅 {ech}{'  🔴 EN RETARD' if retard else ''}" if ech else ""
    return f"**{row.get('Titre','')}**{desc_s}\n\n`{ini}`  ·  {prio}{date_s}"

# ── DONNÉES ───────────────────────────────────────────────────────────────────
df    = load_tasks()
today = date.today()

# ── HEADER ────────────────────────────────────────────────────────────────────
ch, cb = st.columns([6, 1])
with ch:
    st.markdown("## ✅ Tasks Tracker")
    st.caption("Suivi des tâches · Équipe Achats · Google Sheets · Cliquer sur une carte pour l'éditer")
with cb:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔄 Rafraîchir", use_container_width=True):
        load_tasks.clear()
        st.rerun()

st.divider()

# ── KPIs ──────────────────────────────────────────────────────────────────────
total    = len(df)
en_cours = len(df[df["Statut"] == "En cours"])  if not df.empty else 0
bloques  = len(df[df["Statut"] == "Bloqué"])    if not df.empty else 0
termines = len(df[df["Statut"] == "Terminé"])   if not df.empty else 0
retards  = len(df[(df["Échéance"] < today) & (df["Statut"] != "Terminé")]) if not df.empty else 0
taux     = round(termines / total * 100) if total > 0 else 0
nb_arch  = len(df[df.apply(lambda r: is_archived(r, today), axis=1)]) if not df.empty else 0

k1,k2,k3,k4,k5 = st.columns(5)
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
tab_kanban, tab_new, tab_liste = st.tabs([
    "🗂  Kanban", "➕  Nouvelle tâche", "📋  Liste complète"
])

# ═══════════════════════════════════════════════════════════════════════════════
# KANBAN
# ═══════════════════════════════════════════════════════════════════════════════
with tab_kanban:
    st.markdown("<br>", unsafe_allow_html=True)

    # Toggles + filtres
    t1, t2, t3, _, f1, f2 = st.columns([1.2, 1, 1.6, 1.5, 1.8, 1.8])
    if t1.button("Actives",
                 type="primary" if st.session_state.view_mode=="active" else "secondary",
                 use_container_width=True):
        st.session_state.view_mode="active"; st.session_state.edit_id=None; st.rerun()
    if t2.button("Toutes",
                 type="primary" if st.session_state.view_mode=="all" else "secondary",
                 use_container_width=True):
        st.session_state.view_mode="all"; st.session_state.edit_id=None; st.rerun()
    if t3.button(f"📦 Archives ({nb_arch})",
                 type="primary" if st.session_state.view_mode=="archive" else "secondary",
                 use_container_width=True):
        st.session_state.view_mode="archive"; st.session_state.edit_id=None; st.rerun()

    resps_list = sorted(df["Responsable"].dropna().unique().tolist()) if not df.empty else []
    f_resp = f1.selectbox("Responsable", ["Tous"] + resps_list, label_visibility="collapsed")
    f_prio = f2.selectbox("Priorité", ["Toutes"] + PRIORITES,  label_visibility="collapsed")

    st.markdown("<br>", unsafe_allow_html=True)

    if st.session_state.view_mode == "archive" and nb_arch > 0:
        st.markdown(f'<div class="arch-banner">📦 <b>{nb_arch} tâche(s)</b> archivée(s) — terminées depuis +{ARCHIVE_DAYS} jours</div>',
                    unsafe_allow_html=True)

    # Filtrage
    def row_visible(row):
        archived = is_archived(row, today)
        mode = st.session_state.view_mode
        if mode == "archive": return archived
        if mode == "active":  return not archived
        return True

    df_view = df[df.apply(row_visible, axis=1)].copy() if not df.empty else df
    if f_resp != "Tous" and not df_view.empty:
        df_view = df_view[df_view["Responsable"] == f_resp]
    if f_prio != "Toutes" and not df_view.empty:
        df_view = df_view[df_view["Priorité"] == f_prio]

    # Colonnes Kanban
    kanban_cols = st.columns(4)
    for col, statut in zip(kanban_cols, STATUTS):
        cfg = STATUT_CFG[statut]
        sub = df_view[df_view["Statut"] == statut] if not df_view.empty else pd.DataFrame()

        col.markdown(f"""
        <div class="kol">
          <div class="kol-head">
            <div class="kol-name">
              <span class="kol-dot" style="background:{cfg['color']}"></span>{statut}
            </div>
            <span class="kol-badge" style="background:{cfg['color']}">{len(sub)}</span>
          </div>
        </div>""", unsafe_allow_html=True)

        if sub.empty:
            col.markdown('<div class="empty-kol">Aucune tâche</div>', unsafe_allow_html=True)
            continue

        for _, row in sub.iterrows():
            ech    = row.get("Échéance","")
            retard = bool(ech and ech < today and statut != "Terminé")
            done   = statut == "Terminé"
            is_sel = str(row["ID"]) == str(st.session_state.edit_id)

            # CSS classes sur le conteneur du bouton via JS injection
            css_classes = " ".join(filter(None,[
                "card-btn",
                "sel"  if is_sel else "",
                "late" if retard else "",
                "done" if done   else "",
            ]))

            label = card_label(row, today)

            # Injecter les classes CSS sur le prochain bouton
            col.markdown(
                f'<style>div[data-testid="stButton"]:has(+ div[data-testid="stButton"]) {{ display:none }}</style>',
                unsafe_allow_html=True
            )
            # Tag CSS appliqué via data attribute sur le wrapper
            col.markdown(
                f'<div class="{css_classes}" style="margin:0">',
                unsafe_allow_html=True
            )
            clicked = col.button(label, key=f"card_{row['ID']}", use_container_width=True)
            col.markdown("</div>", unsafe_allow_html=True)

            if clicked:
                st.session_state.edit_id = None if is_sel else str(row["ID"])
                st.rerun()

    # ── Panneau édition ────────────────────────────────────────────────────────
    if st.session_state.edit_id:
        st.markdown("<br>", unsafe_allow_html=True)
        mask     = df["ID"] == st.session_state.edit_id
        row_sel  = df[mask].iloc[0]   if mask.any() else None
        row_excel= df.index.get_loc(df[mask].index[0]) + 1 if mask.any() else None

        if row_sel is not None:
            st.markdown(f"#### ✏️ {row_sel['Titre']}")
            with st.form("edit_form"):
                ec1, ec2 = st.columns(2)
                titre       = ec1.text_input("Titre *",       value=str(row_sel["Titre"]))
                responsable = ec2.text_input("Responsable *", value=str(row_sel["Responsable"]))
                description = st.text_area("Description",     value=str(row_sel["Description"]), height=80)
                ec3, ec4, ec5 = st.columns(3)
                statut   = ec3.selectbox("Statut", STATUTS,
                    index=STATUTS.index(row_sel["Statut"]) if row_sel["Statut"] in STATUTS else 0)
                priorite = ec4.selectbox("Priorité", PRIORITES,
                    index=PRIORITES.index(row_sel["Priorité"]) if row_sel["Priorité"] in PRIORITES else 1)
                echeance = ec5.date_input("Échéance",
                    value=row_sel["Échéance"] if row_sel["Échéance"] else today)
                commentaire = st.text_input("Commentaire", value=str(row_sel["Commentaire"]))

                bs, bd, bx = st.columns([3, 1, 1])
                save_btn   = bs.form_submit_button("💾 Enregistrer", type="primary", use_container_width=True)
                delete_btn = bd.form_submit_button("🗑 Supprimer",   use_container_width=True)
                cancel_btn = bx.form_submit_button("✕ Annuler",      use_container_width=True)

                if save_btn:
                    if not titre or not responsable:
                        st.error("Titre et responsable obligatoires.")
                    else:
                        save_task({
                            "ID": row_sel["ID"], "Titre": titre,
                            "Description": description, "Responsable": responsable,
                            "Statut": statut, "Priorité": priorite,
                            "Échéance": str(echeance),
                            "Créé le": str(row_sel["Créé le"]),
                            "Commentaire": commentaire,
                        }, row_index=row_excel)
                        st.session_state.edit_id = None
                        st.success("✅ Tâche mise à jour !")
                        st.rerun()

                if delete_btn:
                    delete_task(row_excel)
                    st.session_state.edit_id = None
                    st.success("🗑 Tâche supprimée.")
                    st.rerun()

                if cancel_btn:
                    st.session_state.edit_id = None
                    st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# NOUVELLE TÂCHE
# ═══════════════════════════════════════════════════════════════════════════════
with tab_new:
    st.markdown("<br>", unsafe_allow_html=True)
    with st.form("new_task_form", clear_on_submit=True):
        nc1, nc2 = st.columns(2)
        titre       = nc1.text_input("Titre *",       placeholder="Ex : Négociation remise MIPA")
        responsable = nc2.text_input("Responsable *", placeholder="Ex : Grace, Carine, Yves...")
        description = st.text_area("Description", height=90, placeholder="Contexte, objectif, détails...")
        nc3, nc4, nc5 = st.columns(3)
        statut   = nc3.selectbox("Statut",   STATUTS)
        priorite = nc4.selectbox("Priorité", PRIORITES, index=1)
        echeance = nc5.date_input("Échéance", value=today)
        commentaire = st.text_input("Commentaire", placeholder="Bloquant, lien utile, note...")

        if st.form_submit_button("➕ Ajouter la tâche", type="primary", use_container_width=True):
            if not titre or not responsable:
                st.error("Titre et responsable obligatoires.")
            else:
                task_id = next_id(df)
                save_task({
                    "ID": task_id, "Titre": titre, "Description": description,
                    "Responsable": responsable, "Statut": statut, "Priorité": priorite,
                    "Échéance": str(echeance), "Créé le": str(today),
                    "Commentaire": commentaire,
                })
                st.success(f"✅ Tâche **{task_id}** ajoutée !")
                st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# LISTE COMPLÈTE
# ═══════════════════════════════════════════════════════════════════════════════
with tab_liste:
    st.markdown("<br>", unsafe_allow_html=True)
    if df.empty:
        st.info("Aucune tâche enregistrée.")
    else:
        resps = sorted(df["Responsable"].dropna().unique().tolist())
        lf1,lf2,lf3,lf4 = st.columns(4)
        lf_resp   = lf1.multiselect("Responsable", resps,     default=resps)
        lf_stat   = lf2.multiselect("Statut",      STATUTS,   default=STATUTS)
        lf_prio   = lf3.multiselect("Priorité",    PRIORITES, default=PRIORITES)
        lf_retard = lf4.checkbox("Retards uniquement")

        df_filt = df[
            df["Responsable"].isin(lf_resp) &
            df["Statut"].isin(lf_stat) &
            df["Priorité"].isin(lf_prio)
        ].copy()

        if lf_retard:
            df_filt = df_filt[(df_filt["Échéance"] < today) & (df_filt["Statut"] != "Terminé")]

        df_filt["⚠️"]      = df_filt["Échéance"].apply(lambda e: "🔴" if (e and e < today) else "")
        df_filt["Archive"] = df_filt.apply(lambda r: "📦" if is_archived(r, today) else "", axis=1)

        st.dataframe(
            df_filt[["ID","Titre","Responsable","Statut","Priorité","Échéance","⚠️","Archive","Commentaire"]],
            use_container_width=True, hide_index=True,
            column_config={"Échéance": st.column_config.DateColumn("Échéance", format="DD/MM/YYYY")},
        )

        nb_ret = len(df_filt[(df_filt["Échéance"] < today) & (df_filt["Statut"] != "Terminé")]) if not df_filt.empty else 0
        nb_arc = len(df_filt[df_filt["Archive"] == "📦"])
        st.caption(f"{len(df_filt)} tâche(s) · {len(df_filt[df_filt['Statut']=='Terminé'])} terminée(s) · {nb_ret} en retard · {nb_arc} archivée(s)")

        if not df_filt.empty and len(df_filt) > 1:
            st.markdown("<br>", unsafe_allow_html=True)
            agg = df_filt.groupby(["Responsable","Statut"]).size().reset_index(name="n")
            fig = go.Figure()
            for statut in STATUTS:
                sub = agg[agg["Statut"] == statut]
                if sub.empty: continue
                fig.add_trace(go.Bar(
                    name=statut, x=sub["Responsable"], y=sub["n"],
                    marker_color=STATUT_CFG[statut]["color"], marker_line_width=0,
                ))
            fig.update_layout(
                barmode="stack", paper_bgcolor=WHITE, plot_bgcolor=WHITE,
                height=240, margin=dict(l=10,r=10,t=10,b=10),
                legend=dict(orientation="h", y=-0.3, font_size=11, font_color=PL_FONT),
                font=dict(family="SF Pro Display,-apple-system,sans-serif", color=PL_FONT, size=12),
                xaxis=dict(gridcolor=PL_GRID, linecolor=PL_GRID),
                yaxis=dict(gridcolor=PL_GRID, tickformat="d"),
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
