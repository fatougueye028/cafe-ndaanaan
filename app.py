# ============================================================
#  ☕ Kafe Ndaanaan — Application de Gestion
#  app.py — MVP 1.0
#  Backend : Google Sheets via gspread
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# ─── Configuration page ────────────────────────────────────────
st.set_page_config(
    page_title="☕ Kafe Ndaanaan",
    page_icon="☕",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ───────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    section[data-testid="stSidebar"] { background-color: #3d1a00; }
    section[data-testid="stSidebar"] * { color: #f5e6d0 !important; }
    section[data-testid="stSidebar"] hr { border-color: #7a3d00; }
    .kpi-box {
        background: #fff8f0;
        border: 1px solid #e8c99a;
        border-left: 5px solid #8B4513;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 6px;
        text-align: center;
    }
    .kpi-box .val { font-size: 1.7rem; font-weight: 700; color: #8B4513; }
    .kpi-box .lbl { font-size: 0.78rem; color: #777; margin-top: 2px; }
    .alert-rouge { background:#fde8e8; border-left:4px solid #e53e3e;
                   border-radius:6px; padding:10px 14px; margin-bottom:8px; }
    .alert-orange { background:#fff3cd; border-left:4px solid #f6993f;
                    border-radius:6px; padding:10px 14px; margin-bottom:8px; }
    @media(max-width:640px){ .block-container{padding:0.5rem 0.4rem;} }
</style>
""", unsafe_allow_html=True)

# ─── Constantes métier ─────────────────────────────────────────
GAMMES          = ["Signature", "Original", "Prestige", "Épicé", "Ñooket"]
FORMATS         = ["250g", "500g", "1kg"]
ZONES           = ["Touba", "Dakar", "France", "Canada", "Autre"]
LOCALISATIONS   = ["Touba", "Dakar", "France", "Canada"]
STATUTS_PAY     = ["Non payé", "Partiel", "Payé"]
SOURCES         = ["WhatsApp", "Facebook", "TikTok", "YouTube",
                   "Recommandation", "Famille", "Autre"]

# ── Pipeline commercial ────────────────────────────────────────
TYPES_DEMANDE = [
    "Prospect / À rappeler",
    "Précommande",
    "Commande confirmée",
    "Préparée",
    "Livrée",
    "Annulée",
]
# Types qui entrent dans le CA réel
TYPES_CA_REEL   = ["Commande confirmée", "Préparée", "Livrée"]
# Types qui déclenchent le décrément stock
TYPES_STOCK     = ["Préparée", "Livrée"]
# Préfixes ID par type
ID_PREFIXES     = {
    "Prospect / À rappeler": "PRO",
    "Précommande":           "PRE",
    "Commande confirmée":    "CMD",
    "Préparée":              "CMD",
    "Livrée":                "CMD",
    "Annulée":               "CMD",
}
DATES_PREVUES   = ["Aujourd'hui", "Cette semaine", "Ce mois",
                   "Magal Touba", "Dans 1 mois", "Dans 2 mois", "À définir"]

# Rétrocompatibilité (ancien champ Statut_Livraison)
STATUTS_LIV = ["À préparer", "Préparée", "Livrée", "Annulée"]

PRIX_FCFA = {
    ("Signature","250g"): 1500, ("Signature","500g"): 3000, ("Signature","1kg"): 6000,
    ("Original", "250g"): 1500, ("Original", "500g"): 3000, ("Original", "1kg"): 6000,
    ("Prestige", "250g"): 2000, ("Prestige", "500g"): 3500, ("Prestige", "1kg"): 7000,
    ("Épicé",    "250g"): 2000, ("Épicé",    "500g"): 3500, ("Épicé",    "1kg"): 7000,
    ("Ñooket",   "250g"): 2000, ("Ñooket",   "500g"): 3500, ("Ñooket",   "1kg"): 7000,
}
PRIX_EUR = {
    ("Signature","250g"): 6.50,  ("Signature","500g"): 10.90, ("Signature","1kg"): 20.00,
    ("Original", "250g"): 6.50,  ("Original", "500g"): 10.90, ("Original", "1kg"): 20.00,
    ("Prestige", "250g"): 6.50,  ("Prestige", "500g"): 10.90, ("Prestige", "1kg"): 20.00,
    ("Épicé",    "250g"): 6.50,  ("Épicé",    "500g"): 10.90, ("Épicé",    "1kg"): 20.00,
    ("Ñooket",   "250g"): 6.50,  ("Ñooket",   "500g"): 10.90, ("Ñooket",   "1kg"): 20.00,
}

COULEURS_GAMME = {
    "Signature": "#8B4513", "Original": "#D2691E",
    "Prestige":  "#A0522D", "Épicé":    "#C0392B", "Ñooket": "#27AE60",
}
COULEURS_STATUT = {
    "À préparer": "#E74C3C", "Préparée": "#F39C12",
    "Livrée":     "#27AE60", "Annulée":  "#95A5A6",
}

# ─── Google Sheets ─────────────────────────────────────────────
@st.cache_resource
def _gs_client():
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    return gspread.authorize(creds)

def _ws(name: str):
    return _gs_client().open(st.secrets["SHEET_NAME"]).worksheet(name)

@st.cache_data(ttl=10)
def load(sheet: str) -> pd.DataFrame:
    records = _ws(sheet).get_all_records(value_render_option="UNFORMATTED_VALUE")
    return pd.DataFrame(records) if records else pd.DataFrame()

def bust():
    load.clear()

def append(sheet: str, row: list):
    _ws(sheet).append_row(row, value_input_option="RAW")
    bust()

def set_cell(sheet: str, df_row_idx: int, col_name: str, df: pd.DataFrame, value):
    """df_row_idx est l'index pandas (0-based). +2 = header + base 1."""
    col_idx = list(df.columns).index(col_name) + 1
    _ws(sheet).update_cell(df_row_idx + 2, col_idx, value)
    bust()

def next_id(df: pd.DataFrame, type_demande: str = "Commande confirmée") -> str:
    year   = date.today().year
    prefix = ID_PREFIXES.get(type_demande, "CMD")
    full_prefix = f"{prefix}-{year}-"
    nums = []
    if not df.empty and "ID" in df.columns:
        for v in df["ID"].dropna().astype(str):
            if v.startswith(full_prefix):
                try: nums.append(int(v.split("-")[-1]))
                except ValueError: pass
    return f"{full_prefix}{(max(nums) + 1 if nums else 1):03d}"

# ─── Navigation ────────────────────────────────────────────────
PAGES = {
    "🏠 Dashboard":         "dashboard",
    "➕ Nouvelle commande":  "new_order",
    "📋 Commandes":         "orders",
    "📦 Stock":             "stock",
    "🏭 Production":        "production",
    "🎨 Sachets & Affiches": "sachets",
    "🚚 Livraisons":        "livreurs",
}

def sidebar_nav() -> str:
    st.sidebar.markdown(
        "<h2 style='text-align:center;font-size:1.4rem;margin-bottom:0'>☕ Kafe Ndaanaan</h2>"
        "<p style='text-align:center;font-size:0.75rem;opacity:0.7;margin-top:2px'>Gestion & Pilotage</p>",
        unsafe_allow_html=True,
    )
    st.sidebar.markdown("---")
    choice = PAGES[st.sidebar.radio("", list(PAGES.keys()), label_visibility="collapsed")]
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Rafraîchir les données", use_container_width=True):
        bust()
        st.rerun()
    return choice

# ─── Helpers d'affichage ───────────────────────────────────────
def kpi(val: str, label: str):
    st.markdown(
        f'<div class="kpi-box"><div class="val">{val}</div><div class="lbl">{label}</div></div>',
        unsafe_allow_html=True,
    )

# ─── Page : Dashboard ──────────────────────────────────────────
def page_dashboard():
    st.title("🏠 Tableau de Bord")

    df = load("Commandes")
    df_stock = load("Stock")

    if df.empty:
        st.info("Aucune commande pour l'instant.")
        return

    df["CA"]       = pd.to_numeric(df["CA"],       errors="coerce").fillna(0)
    df["Quantité"] = pd.to_numeric(df["Quantité"], errors="coerce").fillna(0)

    ca_fcfa   = df[df["Devise"] == "FCFA"]["CA"].sum()
    ca_eur    = df[df["Devise"] == "EUR"]["CA"].sum()
    # Colonne de statut (Type_Demande en priorité, sinon Statut_Livraison)
    statut_col = "Type_Demande" if "Type_Demande" in df.columns else "Statut_Livraison"

    # CA réel = uniquement commandes confirmées/préparées/livrées
    # Les lignes offertes restent dans le CA (comptées comme coût marketing)
    df_reel = df[df[statut_col].isin(TYPES_CA_REEL + ["Livrée", "Préparée"])]
    ca_fcfa  = df_reel[df_reel["Devise"] == "FCFA"]["CA"].sum()
    ca_eur   = df_reel[df_reel["Devise"] == "EUR"]["CA"].sum()

    # Pipeline prévisionnel
    df_prev  = df[df[statut_col].isin(["Prospect / À rappeler", "Précommande"])]
    ca_prev  = df_prev["CA"].sum()

    nb_attente = df[df[statut_col].isin(["Commande confirmée", "À préparer", "Préparée"])]["ID"].nunique()
    nb_livrees = df[df[statut_col] == "Livrée"]["ID"].nunique()
    non_payes  = df_reel[df_reel["Statut_Paiement"].isin(["Non payé", "Partiel"])]["CA"]

    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi(f"{ca_fcfa:,.0f}", "CA réel FCFA")
    with c2: kpi(str(nb_attente), "À préparer / livrer")
    with c3: kpi(f"{non_payes.sum():,.0f}", "Paiements en attente")
    with c4: kpi(f"{ca_prev:,.0f}", "Pipeline prévisionnel")

    # Alerte prospects/précommandes
    if not df_prev.empty:
        nb_prev = df_prev["ID"].nunique()
        st.markdown(
            f'<div class="alert-orange">📋 <b>{nb_prev} prospect(s)/précommande(s)</b> en attente de confirmation</div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    col_a, col_b = st.columns(2)

    with col_a:
        vg = df.groupby("Gamme")["Quantité"].sum().reset_index()
        vg.columns = ["Gamme", "Unités"]
        fig1 = px.bar(
            vg, x="Gamme", y="Unités", color="Gamme",
            title="Unités vendues par gamme",
            color_discrete_map=COULEURS_GAMME,
        )
        fig1.update_layout(showlegend=False, height=300, margin=dict(t=40, b=0))
        st.plotly_chart(fig1, use_container_width=True)

    with col_b:
        zc = df.groupby("Zone").size().reset_index(name="Commandes")
        fig2 = px.pie(
            zc, values="Commandes", names="Zone",
            title="Commandes par zone",
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig2.update_layout(height=300, margin=dict(t=40, b=0))
        st.plotly_chart(fig2, use_container_width=True)

    col_c, col_d = st.columns(2)

    with col_c:
        sc = df["Statut_Livraison"].value_counts().reset_index()
        sc.columns = ["Statut", "Nb"]
        fig3 = px.bar(
            sc, x="Statut", y="Nb", color="Statut",
            title="Statut livraisons",
            color_discrete_map=COULEURS_STATUT,
        )
        fig3.update_layout(showlegend=False, height=280, margin=dict(t=40, b=0))
        st.plotly_chart(fig3, use_container_width=True)

    with col_d:
        if "Source" in df.columns and df["Source"].notna().any():
            src = df["Source"].value_counts().reset_index()
            src.columns = ["Source", "Nb"]
            fig4 = px.pie(
                src, values="Nb", names="Source",
                title="Sources clients",
                color_discrete_sequence=px.colors.qualitative.Pastel,
            )
            fig4.update_layout(height=280, margin=dict(t=40, b=0))
            st.plotly_chart(fig4, use_container_width=True)

    # Alertes stock
    st.markdown("---")
    st.subheader("📦 Alertes Stock")
    if not df_stock.empty:
        col_stock = "Stock_Restant" if "Stock_Restant" in df_stock.columns else "Stock_Dispo" if "Stock_Dispo" in df_stock.columns else None
        if col_stock and "Lot" in df_stock.columns:
            df_stock[col_stock] = pd.to_numeric(df_stock[col_stock], errors="coerce").fillna(0)
            # Garder uniquement le dernier lot
            dernier_lot = df_stock["Lot"].dropna().astype(str).unique().tolist()
            # Trier par numéro de lot (extraire le chiffre)
            def lot_num(l):
                import re
                m = re.search(r'\d+', str(l))
                return int(m.group()) if m else 0
            dernier_lot = max(dernier_lot, key=lot_num) if dernier_lot else None

            if dernier_lot:
                df_dernier = df_stock[df_stock["Lot"].astype(str) == dernier_lot]
                faible = df_dernier[df_dernier[col_stock] <= 5]
                if not faible.empty:
                    st.markdown(
                        f'<div class="alert-orange">🟠 <b>{len(faible)} références ≤ 5 unités ({dernier_lot})</b></div>',
                        unsafe_allow_html=True,
                    )
                    cols_alerte = [c for c in ["Lot","Gamme","Format",col_stock] if c in faible.columns]
                    st.dataframe(faible[cols_alerte], hide_index=True)
                else:
                    st.success(f"✅ Stock OK sur le {dernier_lot}.")
    else:
        st.info("Aucun stock renseigné.")

# ─── Page : Nouvelle Commande ──────────────────────────────────
def page_new_order():
    st.title("➕ Nouvelle Commande")

    df_cmd = load("Commandes")

    with st.form("form_cmd", clear_on_submit=True):
        st.subheader("Type de demande")
        t1, t2 = st.columns(2)
        with t1:
            type_dem = st.selectbox("Type *", TYPES_DEMANDE, index=2,
                help="Prospect = intention non confirmée | Précommande = réservation | Commande confirmée = vente réelle")
        with t2:
            date_prevue = st.selectbox("Date prévue", DATES_PREVUES, index=6)

        st.subheader("Client")
        c1, c2 = st.columns(2)
        with c1:
            client  = st.text_input("Nom *", placeholder="Aminata Diallo")
            tel     = st.text_input("Téléphone", placeholder="77 123 45 67")
            source  = st.selectbox("Source", SOURCES)
        with c2:
            zone    = st.selectbox("Zone *", ZONES)
            adresse = st.text_input("Adresse (France/Canada)", placeholder="12 rue des Lilas, 75010 Paris")
            comm    = st.text_area("Commentaire / Notes", height=70)

        st.subheader("Produit")
        p1, p2, p3, p4 = st.columns(4)
        with p1:
            gamme  = st.selectbox("Gamme *", GAMMES)
        with p2:
            fmt    = st.selectbox("Format *", FORMATS)
        with p3:
            qty    = st.number_input("Quantité *", min_value=1, value=1)
        with p4:
            devise = "EUR" if zone == "France" else ("CAD" if zone == "Canada" else "FCFA")
            prix_d = PRIX_EUR.get((gamme, fmt), 0) if devise == "EUR" else PRIX_FCFA.get((gamme, fmt), 0)
            prix   = st.number_input(f"Prix unitaire ({devise})", value=float(prix_d), min_value=0.0, step=0.5)

        ca = round(prix * qty, 2)
        est_reel = type_dem in TYPES_CA_REEL
        if est_reel:
            st.info(f"💰 **CA réel : {ca:,.2f} {devise}**")
        else:
            st.warning(f"📋 **Prévisionnel : {ca:,.2f} {devise}** — n'entre pas dans le CA tant que non confirmé")

        l1, l2 = st.columns(2)
        with l1:
            lot_prevu = st.text_input("Lot prévu", placeholder="Lot 4, Lot Magal, À définir")
        with l2:
            offert = st.checkbox("🎁 Offre commerciale")

        submitted = st.form_submit_button("✅ Enregistrer", use_container_width=True, type="primary")

    if submitted:
        if not client.strip():
            st.error("Le nom du client est obligatoire.")
        else:
            new_id = next_id(df_cmd, type_dem)
            # Statut_Livraison déduit du Type_Demande
            statut_liv = {"Livrée": "Livrée", "Préparée": "Préparée",
                          "Annulée": "Annulée"}.get(type_dem, "À préparer")
            append("Commandes", [
                date.today().strftime("%d/%m/%Y"),
                new_id, client.strip(), tel, adresse, zone,
                gamme, fmt, qty, prix, ca, devise,
                type_dem, statut_liv, "Non payé",
                source, lot_prevu, comm, date_prevue,
                "Offre commerciale" if offert else "",
            ])
            st.success(f"✅ {type_dem} **{new_id}** enregistrée !")
            if est_reel:
                st.balloons()

# ─── Page : Commandes ──────────────────────────────────────────
def page_orders():
    st.title("📋 Commandes")

    df = load("Commandes")
    if df.empty:
        st.info("Aucune commande enregistrée.")
        return

    df["CA"]       = pd.to_numeric(df["CA"],       errors="coerce").fillna(0)
    df["Quantité"] = pd.to_numeric(df["Quantité"], errors="coerce").fillna(0)

    # Colonne de statut active
    statut_col  = "Type_Demande" if "Type_Demande" in df.columns else "Statut_Livraison"
    statut_vals = sorted(df[statut_col].dropna().unique().tolist()) if statut_col in df.columns else TYPES_DEMANDE

    # CA réel = Type_Demande confirmé/préparé/livré
    REEL = ["Commande confirmée", "Préparée", "Livrée", "À préparer"]
    df_reel = df[df[statut_col].isin(REEL)]

    # Prospects / précommandes
    PREV = ["Prospect / À rappeler", "Précommande"]
    df_prev = df[df[statut_col].isin(PREV)]

    # ── KPIs — utilise Statut_Livraison pour les statuts opérationnels ──
    col_liv = "Statut_Livraison" if "Statut_Livraison" in df.columns else statut_col
    nb_total   = df["ID"].nunique()
    nb_livrees = df[df[col_liv] == "Livrée"]["ID"].nunique()
    # À préparer : uniquement commandes confirmées (pas prospects/précommandes)
    mask_prep = (
        df[col_liv].isin(["À préparer", "Préparée"])
        & df[statut_col].isin(["Commande confirmée", "À préparer", "Préparée"])
    )
    nb_a_prep  = df[mask_prep]["ID"].nunique()

    k1, k2, k3, k4 = st.columns(4)
    with k1: kpi(str(nb_total),   "Total commandes")
    with k2: kpi(str(nb_livrees), "CMD livrées")
    with k3: kpi(str(nb_a_prep),  "À préparer / Préparées")
    with k4:
        kpi(f"{df_reel[df_reel['Statut_Paiement'].isin(['Non payé','Partiel'])]['CA'].sum():,.0f}", "Paiements en attente")

    if not df_prev.empty:
        st.markdown(
            f'<div class="alert-orange">📋 <b>{df_prev["ID"].nunique()} prospect(s)/précommande(s)</b> à confirmer</div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Filtres ──
    with st.expander("🔍 Filtres", expanded=False):
        f1, f2, f3, f4 = st.columns(4)
        with f1: fz = st.multiselect("Zone",      ZONES,        default=ZONES)
        with f2: fl = st.multiselect("Statut",    statut_vals,  default=statut_vals)
        with f3: fp = st.multiselect("Paiement",  STATUTS_PAY,  default=STATUTS_PAY)
        with f4: fg = st.multiselect("Gamme",     GAMMES,       default=GAMMES)
        recherche = st.text_input("🔎 Rechercher un client", placeholder="Nom du client…")

    mask = (
        df["Zone"].isin(fz)
        & df[statut_col].isin(fl)
        & df["Statut_Paiement"].isin(fp)
        & df["Gamme"].isin(fg)
    )
    df_f = df[mask].copy()
    if recherche.strip():
        df_f = df_f[df_f["Client"].str.contains(recherche.strip(), case=False, na=False)]

    ca_fcfa = df_f[df_f["Devise"] == "FCFA"]["CA"].sum()
    st.caption(f"**{df_f['ID'].nunique()} commande(s)** — CA : {ca_fcfa:,.0f} FCFA")

    # Colonnes affichées — toutes les colonnes présentes dans le sheet
    COLS_PRIORITE = ["Date","Lot","ID","Client","Zone","Gamme","Format",
                     "Quantité","CA","Type_Demande","Statut_Livraison",
                     "Statut_Paiement","Source","Date_Prevue","Offre_Commerciale","Commentaire"]
    COLS = [c for c in COLS_PRIORITE if c in df_f.columns]

    st.dataframe(
        df_f[COLS],
        use_container_width=True,
        hide_index=True,
        column_config={
            "CA":               st.column_config.NumberColumn("CA (FCFA)", format="%.0f"),
            "Type_Demande":     st.column_config.SelectboxColumn("Type", options=TYPES_DEMANDE),
            "Statut_Livraison": st.column_config.SelectboxColumn("Livraison", options=STATUTS_LIV),
            "Statut_Paiement":  st.column_config.SelectboxColumn("Paiement", options=STATUTS_PAY),
            "Date_Prevue":      st.column_config.TextColumn("Date prévue"),
            "Offre_Commerciale":st.column_config.TextColumn("Offre comm."),
        },
    )

    # ── Modifier ──
    st.markdown("---")
    st.subheader("✏️ Modifier une commande")

    ids = df_f["ID"].dropna().astype(str).unique().tolist()
    if not ids:
        return

    sel_id = st.selectbox("Choisir une commande", ids)
    df_full = load("Commandes")
    row = df_full[df_full["ID"] == sel_id]
    if row.empty:
        st.warning("Commande introuvable.")
        return
    row = row.iloc[0]
    row_idx = df_full.index[df_full["ID"] == sel_id][0]

    c1, c2, c3 = st.columns(3)
    with c1:
        idx_l = STATUTS_LIV.index(row["Statut_Livraison"]) if row["Statut_Livraison"] in STATUTS_LIV else 0
        new_liv = st.selectbox("Statut livraison", STATUTS_LIV, index=idx_l)
    with c2:
        idx_p = STATUTS_PAY.index(row["Statut_Paiement"]) if row["Statut_Paiement"] in STATUTS_PAY else 0
        new_pay = st.selectbox("Statut paiement", STATUTS_PAY, index=idx_p)
    with c3:
        new_comm = st.text_input("Commentaire", value=str(row.get("Commentaire", "") or ""))

    if st.button("💾 Sauvegarder", use_container_width=True, type="primary"):
        ws = _ws("Commandes")
        col_l = list(df_full.columns).index("Statut_Livraison") + 1
        col_p = list(df_full.columns).index("Statut_Paiement") + 1
        col_c = list(df_full.columns).index("Commentaire") + 1
        sheet_row = row_idx + 2
        ws.update_cell(sheet_row, col_l, new_liv)
        ws.update_cell(sheet_row, col_p, new_pay)
        ws.update_cell(sheet_row, col_c, new_comm)

        # Décrémenter stock uniquement si passage à "Préparée" ou "Livrée"
        statut_col_local = "Type_Demande" if "Type_Demande" in df_full.columns else "Statut_Livraison"
        ancien_statut = row.get(statut_col_local, "")
        if new_liv in TYPES_STOCK and ancien_statut not in TYPES_STOCK:
            _decrement_stock(row)

        bust()
        st.success(f"✅ Commande {sel_id} mise à jour !")
        st.rerun()

def _update_stock_row(df_s, idx, qty_delta_vend=0, qty_delta_prod=0, qty_delta_cmd=0):
    """Met à jour une ligne du stock et recalcule Stock_Restant."""
    ws_s    = _ws("Stock")
    cols    = list(df_s.columns)
    today_s = date.today().strftime("%d/%m/%Y")

    def _update(col_name, delta):
        if col_name in cols:
            col_idx = cols.index(col_name) + 1
            current = int(pd.to_numeric(df_s.loc[idx, col_name], errors="coerce") or 0)
            new_val = max(0, current + delta)
            ws_s.update_cell(idx + 2, col_idx, new_val)
            return new_val
        return 0

    _update("Unites_Vendues",   qty_delta_vend)
    _update("Unites_Commandees", qty_delta_cmd)
    _update("Unites_Produites", qty_delta_prod)

    # Recalcul Stock_Restant
    prod  = int(pd.to_numeric(df_s.loc[idx, "Unites_Produites"],  errors="coerce") or 0) + qty_delta_prod
    vend  = int(pd.to_numeric(df_s.loc[idx, "Unites_Vendues"],    errors="coerce") or 0) + qty_delta_vend
    restant = max(0, prod - vend)
    if "Stock_Restant" in cols:
        ws_s.update_cell(idx + 2, cols.index("Stock_Restant") + 1, restant)
    if "Derniere_MAJ" in cols:
        ws_s.update_cell(idx + 2, cols.index("Derniere_MAJ") + 1, today_s)

def _decrement_stock(row):
    """Incrémente Unites_Vendues quand une commande passe à Livrée."""
    try:
        df_s = load("Stock")
        if df_s.empty:
            return
        qty  = int(pd.to_numeric(row.get("Quantité", 0), errors="coerce") or 0)
        mask = (df_s["Gamme"] == row["Gamme"]) & (df_s["Format"] == row["Format"])
        if not mask.any():
            return
        # Prendre la ligne avec le plus de stock restant
        idx = df_s[mask].sort_values("Stock_Restant", ascending=False).index[0] \
              if "Stock_Restant" in df_s.columns else df_s.index[mask][0]
        _update_stock_row(df_s, idx, qty_delta_vend=qty, qty_delta_cmd=qty)
        bust()
    except Exception as e:
        st.warning(f"⚠️ Stock non décrémenté : {e}")

# ─── Page : Stock ──────────────────────────────────────────────
def page_stock():
    st.title("📦 Suivi des Stocks")

    df_s = load("Stock")

    NUM_COLS = ["Unites_Produites", "Unites_Commandees", "Unites_Vendues", "Stock_Restant"]
    if not df_s.empty:
        for c in NUM_COLS:
            if c in df_s.columns:
                df_s[c] = pd.to_numeric(df_s[c], errors="coerce").fillna(0).astype(int)

    # ── Tableau principal ────────────────────────────────────────
    if df_s.empty:
        st.info("Aucun stock renseigné. Lance setup_sheets.py puis migrate_stock.py.")
    else:
        gam_f = st.multiselect("Filtrer par gamme", GAMMES, default=GAMMES)
        df_f  = df_s[df_s["Gamme"].isin(gam_f)].copy()

        COLS_SHOW = ["Date", "Lot", "Gamme", "Format",
                     "Unites_Produites", "Unites_Commandees",
                     "Unites_Vendues",   "Stock_Restant"]
        COLS_SHOW = [c for c in COLS_SHOW if c in df_f.columns]

        def style_restant(val):
            try:
                v = int(val)
                if v <= 0: return "background-color:#fde8e8;color:#c0392b;font-weight:bold"
                if v <= 5: return "background-color:#fff3cd;color:#e67e22;font-weight:bold"
                return "background-color:#eafaf1;color:#27ae60"
            except:
                return ""

        rest_col = [c for c in ["Stock_Restant"] if c in COLS_SHOW]
        styled = df_f[COLS_SHOW].style.map(style_restant, subset=rest_col) if rest_col \
                 else df_f[COLS_SHOW]

        st.dataframe(styled, use_container_width=True, hide_index=True,
            column_config={
                "Unites_Produites":  st.column_config.NumberColumn("Unités Produites"),
                "Unites_Commandees": st.column_config.NumberColumn("Unités Commandées"),
                "Unites_Vendues":    st.column_config.NumberColumn("Unités Vendues"),
                "Stock_Restant":     st.column_config.NumberColumn("Stock Restant"),
            }
        )

        # Graphique
        if "Stock_Restant" in df_f.columns and df_f["Stock_Restant"].sum() > 0:
            fig = px.bar(
                df_f.groupby(["Gamme", "Format"])["Stock_Restant"].sum().reset_index(),
                x="Gamme", y="Stock_Restant", color="Format",
                title="Stock restant par gamme", barmode="group",
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig.update_layout(height=280, margin=dict(t=40, b=0))
            st.plotly_chart(fig, use_container_width=True)

    # ── Ajout manuel d'une ligne ─────────────────────────────────
    st.markdown("---")
    st.subheader("➕ Ajouter / mettre à jour un stock")

    with st.form("form_stock", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1: gam_s = st.selectbox("Gamme",  GAMMES,  key="sg")
        with c2: fmt_s = st.selectbox("Format", FORMATS, key="sf")
        with c3: lot_s = st.text_input("Lot", placeholder="Lot 4")
        with c4: date_s = st.date_input("Date", value=date.today())

        n1, n2, n3, n4 = st.columns(4)
        with n1: prod_s = st.number_input("Unités Produites",   min_value=0, value=0)
        with n2: cmd_s  = st.number_input("Unités Commandées",  min_value=0, value=0)
        with n3: vend_s = st.number_input("Unités Vendues",     min_value=0, value=0)
        with n4: rest_s = st.number_input("Stock Restant",      min_value=0, value=0)

        comm_s = st.text_input("Commentaire")

        if st.form_submit_button("💾 Enregistrer", use_container_width=True, type="primary"):
            append("Stock", [
                date_s.strftime("%d/%m/%Y"), lot_s, gam_s, fmt_s,
                prod_s, cmd_s, vend_s, rest_s, comm_s
            ])
            bust()
            st.success("✅ Ligne stock enregistrée !")
            st.rerun()

# ─── Page : Livreurs ───────────────────────────────────────────
def page_livreurs():
    st.title("🚚 Livreurs")

    tab1, tab2 = st.tabs(["🛵 Courses livreurs", "📋 Contacts transporteurs"])

    # ── Tab 1 : Courses ──────────────────────────────────────────
    with tab1:
        df = load("Livreurs")

        if not df.empty:
            df["Montant"] = pd.to_numeric(df["Montant"], errors="coerce").fillna(0)
            df["Courses"] = pd.to_numeric(df["Courses"], errors="coerce").fillna(0)

            resume = (
                df.groupby("Livreur")
                .agg(Courses=("Courses", "sum"), Montant_Total=("Montant", "sum"))
                .reset_index()
            )
            c1, c2 = st.columns(2)
            with c1:
                st.metric("Total courses", int(resume["Courses"].sum()))
            with c2:
                st.metric("Total payé", f"{resume['Montant_Total'].sum():,.0f} FCFA")

            st.subheader("Résumé par livreur")
            st.dataframe(resume, use_container_width=True, hide_index=True,
                column_config={"Montant_Total": st.column_config.NumberColumn("Montant Payé (FCFA)", format="%.0f")}
            )

            st.subheader("Historique des courses")
            st.dataframe(df, use_container_width=True, hide_index=True,
                column_config={"Montant": st.column_config.NumberColumn("Montant Payé (FCFA)", format="%.0f")}
            )

        st.markdown("---")
        st.subheader("➕ Nouvelle course")

    with st.form("form_liv", clear_on_submit=True):
        l1, l2 = st.columns(2)
        with l1:
            livreur  = st.text_input("Livreur", placeholder="Thierno")
            zone_l   = st.text_input("Zone", placeholder="Liberté 6, Golf, Parcelles…")
        with l2:
            tarif    = st.number_input("Tarif par course (FCFA)", min_value=0, value=1500, step=500)
            nb       = st.number_input("Nombre de courses", min_value=1, value=1)

        montant = tarif * nb
        st.info(f"💵 Montant total : {montant:,} FCFA")

        if st.form_submit_button("✅ Enregistrer", use_container_width=True, type="primary"):
            if livreur.strip():
                append("Livreurs", [
                    livreur.strip(),
                    date.today().strftime("%d/%m/%Y"),
                    zone_l, tarif, nb, montant,
                ])
                bust()
                st.success("✅ Course enregistrée !")
                st.rerun()
            else:
                st.error("Le nom du livreur est obligatoire.")

    # ── Tab 2 : Contacts transporteurs ──────────────────────────
    with tab2:
        df_t = load("Transporteurs")

        if not df_t.empty:
            st.dataframe(df_t, use_container_width=True, hide_index=True,
                column_config={
                    "Nom":       st.column_config.TextColumn("Transporteur"),
                    "Type":      st.column_config.TextColumn("Type"),
                    "Zone":      st.column_config.TextColumn("Zone couverte"),
                    "Tarif":     st.column_config.TextColumn("Tarif"),
                    "Telephone": st.column_config.TextColumn("Téléphone"),
                    "Delai":     st.column_config.TextColumn("Délai"),
                    "Notes":     st.column_config.TextColumn("Notes"),
                }
            )

        st.markdown("---")
        st.subheader("➕ Ajouter un contact")
        with st.form("form_trans", clear_on_submit=True):
            t1, t2 = st.columns(2)
            with t1:
                nom_t  = st.text_input("Nom du transporteur *")
                type_t = st.selectbox("Type", ["Livraison locale", "Export France", "Export Canada", "Autre"])
                zone_t = st.text_input("Zone couverte")
            with t2:
                tarif_t = st.text_input("Tarif", placeholder="Ex: 2000 FCFA / 5€ par kg")
                tel_t   = st.text_input("Téléphone")
                delai_t = st.text_input("Délai", placeholder="Ex: 1 mois")
            notes_t = st.text_input("Notes")

            if st.form_submit_button("💾 Enregistrer", use_container_width=True, type="primary"):
                if nom_t.strip():
                    append("Transporteurs", [nom_t.strip(), type_t, zone_t, tarif_t, tel_t, delai_t, notes_t])
                    bust()
                    st.success("✅ Contact enregistré !")
                    st.rerun()
                else:
                    st.error("Le nom est obligatoire.")

# ─── Main ──────────────────────────────────────────────────────
# ─── Page : Production ─────────────────────────────────────────
def page_production():
    st.title("🏭 Production")

    df_prod = load("Production")

    # ── Historique des lots ──
    if not df_prod.empty:
        st.subheader("Historique des sous-lots")

        for col in ["Cout_Total", "Cout_Revient_kg", "Cafe_Brut_kg", "Cafe_Net_kg"]:
            if col in df_prod.columns:
                df_prod[col] = pd.to_numeric(df_prod[col], errors="coerce").fillna(0)

        COLS_DISPLAY = ["Lot", "Gamme", "Date", "Cafe_Brut_kg", "Cafe_Net_kg",
                        "Cout_Total", "Cout_Revient_kg", "Notes"]
        # Vue résumée
        cols_ok = [c for c in COLS_DISPLAY if c in df_prod.columns]
        st.dataframe(
            df_prod[cols_ok],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Cout_Total":       st.column_config.NumberColumn("Coût total (FCFA)", format="%.0f"),
                "Cout_Revient_kg":  st.column_config.NumberColumn("Coût/kg fini (FCFA)", format="%.0f"),
                "Cafe_Brut_kg":     st.column_config.NumberColumn("Café brut (kg)", format="%.1f"),
                "Cafe_Net_kg":      st.column_config.NumberColumn("Café net (kg)", format="%.1f"),
            }
        )

        # Vue détaillée
        with st.expander("📊 Voir le détail complet des charges"):
            DETAIL_COLS = {
                "Lot":              "Lot",
                "Gamme":            "Gamme",
                "Date":             "Date",
                "Cafe_Brut_kg":     "Café brut (kg)",
                "Prix_Cafe_Brut":   "Prix café brut (FCFA/kg)",
                "Cout_Cafe_Brut":   "Coût achat café brut",
                "Jar_kg":           "Baies de Selim (kg)",
                "Prix_Jar":         "Prix Jar (FCFA/kg)",
                "Cout_Jar":         "Coût achat Jar",
                "Clous_FCFA":       "Clous de girofle",
                "Poivre_FCFA":      "Poivre",
                "Gingembre_FCFA":   "Gingembre",
                "Frais_Torrefaction": "Torréfaction & moulage",
                "Frais_Transport":  "Transport",
                "Sachets_FCFA":     "Sachets / contenants",
                "Main_Oeuvre":      "Main d'œuvre",
                "Affiches_FCFA":    "Affiches & impression",
                "Emballage_FCFA":   "Emballage gros commandes",
                "Marketing_FCFA":   "Coût marketing",
                "Cafe_Net_kg":      "Café net vendable (kg)",
                "Cout_Total":       "COÛT TOTAL (FCFA)",
                "Cout_Revient_kg":  "Coût de revient/kg",
                "Notes":            "Notes",
            }
            cols_detail = [c for c in DETAIL_COLS.keys() if c in df_prod.columns]
            df_detail = df_prod[cols_detail].rename(columns=DETAIL_COLS)

            # Mettre en forme les colonnes numériques
            num_cols = [v for k, v in DETAIL_COLS.items()
                        if k not in ("Lot","Gamme","Date","Notes") and v in df_detail.columns]

            col_cfg_detail = {c: st.column_config.NumberColumn(c, format="%.0f") for c in num_cols}
            col_cfg_detail["Café brut (kg)"]       = st.column_config.NumberColumn("Café brut (kg)", format="%.1f")
            col_cfg_detail["Baies de Selim (kg)"]  = st.column_config.NumberColumn("Baies de Selim (kg)", format="%.1f")
            col_cfg_detail["Café net vendable (kg)"] = st.column_config.NumberColumn("Café net vendable (kg)", format="%.1f")
            col_cfg_detail["Coût de revient/kg"]   = st.column_config.NumberColumn("Coût de revient/kg", format="%.0f")

            st.dataframe(df_detail, use_container_width=True, hide_index=True, column_config=col_cfg_detail)

        # Graph coût de revient par gamme
        if "Cout_Revient_kg" in df_prod.columns and df_prod["Cout_Revient_kg"].sum() > 0:
            fig = px.bar(
                df_prod[df_prod["Cout_Revient_kg"] > 0],
                x="Gamme", y="Cout_Revient_kg", color="Lot",
                title="Coût de revient au kg par gamme (FCFA)",
                barmode="group",
            )
            fig.update_layout(height=300, margin=dict(t=40, b=0))
            st.plotly_chart(fig, use_container_width=True)

    # ── Nouveau sous-lot ──
    st.markdown("---")
    st.subheader("➕ Nouveau sous-lot de production")

    with st.form("form_prod", clear_on_submit=True):

        # Identification
        st.markdown("**Identification**")
        id1, id2, id3 = st.columns(3)
        with id1:
            lot_num = st.text_input("Numéro de lot *", placeholder="Lot 4")
        with id2:
            gamme_p = st.selectbox("Gamme *", GAMMES)
        with id3:
            date_p  = st.date_input("Date de lancement", value=date.today())

        st.markdown("**Matières premières**")
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            cafe_brut_kg   = st.number_input("Café brut alloué (kg)", min_value=0.0, value=0.0, step=0.5)
        with m2:
            prix_cafe_brut = st.number_input("Prix café brut (FCFA/kg)", min_value=0, value=2300, step=50)
        with m3:
            jar_kg         = st.number_input("Baies de Selim / Jar (kg)", min_value=0.0, value=0.0, step=0.1)
        with m4:
            prix_jar       = st.number_input("Prix Jar (FCFA/kg)", min_value=0, value=2000, step=100)

        st.markdown("**Épices** (laisser à 0 si non utilisé)")
        e1, e2, e3 = st.columns(3)
        with e1:
            clous_fcfa    = st.number_input("Clous de girofle (FCFA)", min_value=0, value=0, step=100)
        with e2:
            poivre_fcfa   = st.number_input("Poivre (FCFA)", min_value=0, value=0, step=100)
        with e3:
            gingembre_fcfa = st.number_input("Gingembre (FCFA)", min_value=0, value=0, step=100)

        st.markdown("**Charges de fabrication**")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            frais_torref   = st.number_input("Torréfaction & moulage (FCFA)", min_value=0, value=0, step=500)
        with c2:
            frais_transport = st.number_input("Transport (FCFA)", min_value=0, value=0, step=500)
        with c3:
            sachets_fcfa   = st.number_input("Sachets / contenants (FCFA)", min_value=0, value=0, step=500)
        with c4:
            main_oeuvre    = st.number_input("Main d'œuvre (FCFA)", min_value=0, value=0, step=500)

        c5, c6, c7 = st.columns(3)
        with c5:
            affiches_fcfa  = st.number_input("Affiches & impression (FCFA)", min_value=0, value=0, step=500)
        with c6:
            emballage_fcfa = st.number_input("Emballage gros commandes (FCFA)", min_value=0, value=0, step=500)
        with c7:
            marketing_fcfa = st.number_input("Coût marketing (FCFA)", min_value=0, value=0, step=500)

        st.markdown("**Résultat de production**")
        r1, r2 = st.columns(2)
        with r1:
            cafe_net_kg = st.number_input("Café net vendable obtenu (kg)", min_value=0.0, value=0.0, step=0.5)
        with r2:
            notes_p = st.text_input("Notes techniques", placeholder="Ex: café de meilleure qualité")

        st.markdown("**Stock produit** (unités emballées)")
        s1, s2, s3 = st.columns(3)
        with s1:
            qte_250g = st.number_input("Sachets 250g", min_value=0, value=0)
        with s2:
            qte_500g = st.number_input("Sachets 500g", min_value=0, value=0)
        with s3:
            qte_1kg  = st.number_input("Sachets 1kg",  min_value=0, value=0)

        # Calculs automatiques
        cout_cafe_brut = cafe_brut_kg * prix_cafe_brut
        cout_jar       = jar_kg * prix_jar
        cout_total     = (cout_cafe_brut + cout_jar + clous_fcfa + poivre_fcfa +
                          gingembre_fcfa + frais_torref + frais_transport +
                          sachets_fcfa + main_oeuvre + affiches_fcfa +
                          emballage_fcfa + marketing_fcfa)
        cout_revient_kg = round(cout_total / cafe_net_kg, 2) if cafe_net_kg > 0 else 0

        st.info(
            f"💰 **Coût de production total : {cout_total:,.0f} FCFA** | "
            f"Coût de revient au kg fini : **{cout_revient_kg:,.0f} FCFA/kg**"
        )

        submitted = st.form_submit_button("✅ Enregistrer ce sous-lot", use_container_width=True, type="primary")

    if submitted:
        if not lot_num.strip():
            st.error("Le numéro de lot est obligatoire.")
        elif cafe_brut_kg == 0:
            st.error("Le café brut alloué doit être > 0.")
        else:
            append("Production", [
                lot_num.strip(), gamme_p, date_p.strftime("%d/%m/%Y"),
                cafe_brut_kg, prix_cafe_brut, round(cout_cafe_brut),
                jar_kg, prix_jar, round(cout_jar),
                clous_fcfa, poivre_fcfa, gingembre_fcfa,
                frais_torref, frais_transport, sachets_fcfa,
                main_oeuvre, affiches_fcfa, emballage_fcfa, marketing_fcfa,
                cafe_net_kg, round(cout_total), cout_revient_kg,
                qte_250g, qte_500g, qte_1kg, notes_p,
            ])

            # Incrémenter le stock
            _increment_stock(gamme_p, "250g", qte_250g)
            _increment_stock(gamme_p, "500g", qte_500g)
            _increment_stock(gamme_p, "1kg",  qte_1kg)

            st.success(f"✅ Sous-lot {lot_num} — {gamme_p} enregistré ! "
                       f"Stock incrémenté : {qte_250g}×250g, {qte_500g}×500g, {qte_1kg}×1kg")
            st.balloons()

def _increment_stock(gamme: str, fmt: str, qty: int):
    """Ajoute au stock Touba quand un lot est produit."""
    if qty <= 0:
        return
    try:
        df_s = load("Stock")
        if df_s.empty:
            return
        mask = (
            (df_s["Gamme"]          == gamme)
            & (df_s["Format"]       == fmt)
            & (df_s["Localisation"] == "Touba")
        )
        if not mask.any():
            return
        idx = df_s.index[mask][0]
        _update_stock_row(df_s, idx, qty_delta_prod=qty)
        _log_mouvement(gamme, fmt, "Touba", "Entrée (production)", qty)
        bust()
    except Exception as e:
        st.warning(f"⚠️ Stock non incrémenté pour {gamme} {fmt} : {e}")


# ─── Main ──────────────────────────────────────────────────────
# ─── Page : Sachets & Affiches ─────────────────────────────────
def page_sachets():
    st.title("🎨 Sachets & Affiches")

    COULEURS = ["Blanc", "Noir", "Doré", "Doré vif", "Argenté", "Autre"]

    tab1, tab2 = st.tabs(["🖼️ Stock Affiches (par gamme)", "📦 Stock Sachets (par couleur)"])

    def style_stock(val):
        try:
            v = int(val)
            if v <= 0:  return "background-color:#fde8e8;color:#c0392b;font-weight:bold"
            if v <= 20: return "background-color:#fff3cd;color:#e67e22;font-weight:bold"
            return "background-color:#eafaf1;color:#27ae60"
        except: return ""

    # ── Tab 1 : Affiches par gamme ───────────────────────────────
    with tab1:
        df_a = load("Affiches")

        if not df_a.empty:
            for c in ["Qte_Imprimee", "Qte_Utilisee", "Stock_Restant"]:
                if c in df_a.columns:
                    df_a[c] = pd.to_numeric(df_a[c], errors="coerce").fillna(0).astype(int)

            COLS_A = ["Date", "Gamme", "Format", "Qte_Imprimee", "Qte_Utilisee", "Stock_Restant"]
            COLS_A = [c for c in COLS_A if c in df_a.columns]
            rest_a = [c for c in ["Stock_Restant"] if c in COLS_A]
            styled_a = df_a[COLS_A].style.map(style_stock, subset=rest_a) if rest_a else df_a[COLS_A]
            st.dataframe(styled_a, use_container_width=True, hide_index=True,
                column_config={
                    "Qte_Imprimee":  st.column_config.NumberColumn("Imprimées"),
                    "Qte_Utilisee":  st.column_config.NumberColumn("Utilisées"),
                    "Stock_Restant": st.column_config.NumberColumn("Stock Restant"),
                }
            )
            if "Stock_Restant" in df_a.columns and df_a["Stock_Restant"].sum() > 0:
                fig = px.bar(df_a.groupby("Gamme")["Stock_Restant"].sum().reset_index(),
                             x="Gamme", y="Stock_Restant", color="Gamme",
                             title="Stock affiches par gamme",
                             color_discrete_map=COULEURS_GAMME)
                fig.update_layout(showlegend=False, height=260, margin=dict(t=40,b=0))
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.subheader("➕ Nouvel achat d'affiches")
        with st.form("form_affiche", clear_on_submit=True):
            a1, a2, a3 = st.columns(3)
            with a1: gam_a  = st.selectbox("Gamme",  GAMMES)
            with a2: fmt_a  = st.selectbox("Format", FORMATS)
            with a3: date_a = st.date_input("Date", value=date.today())
            b1, b2, b3 = st.columns(3)
            with b1: qte_imp  = st.number_input("Qté imprimée",  min_value=0, value=0)
            with b2: qte_util = st.number_input("Qté utilisée",  min_value=0, value=0)
            with b3: stock_a  = st.number_input("Stock restant", min_value=0, value=0)
            notes_a = st.text_input("Notes")
            if st.form_submit_button("💾 Enregistrer", use_container_width=True, type="primary"):
                append("Affiches", [date_a.strftime("%d/%m/%Y"), gam_a, fmt_a,
                                    qte_imp, qte_util, stock_a, notes_a])
                bust(); st.success("✅ Enregistré !"); st.rerun()

    # ── Tab 2 : Sachets par couleur ──────────────────────────────
    with tab2:
        df_s = load("Sachets")

        if not df_s.empty:
            for c in ["Qte_Achetee", "Qte_Utilisee", "Stock_Restant"]:
                if c in df_s.columns:
                    df_s[c] = pd.to_numeric(df_s[c], errors="coerce").fillna(0).astype(int)

            COLS_S = ["Date", "Gamme", "Couleur", "Format", "Qte_Achetee", "Qte_Utilisee", "Stock_Restant"]
            COLS_S = [c for c in COLS_S if c in df_s.columns]
            rest_s = [c for c in ["Stock_Restant"] if c in COLS_S]
            styled_s = df_s[COLS_S].style.map(style_stock, subset=rest_s) if rest_s else df_s[COLS_S]
            st.dataframe(styled_s, use_container_width=True, hide_index=True,
                column_config={
                    "Qte_Achetee":   st.column_config.NumberColumn("Achetés"),
                    "Qte_Utilisee":  st.column_config.NumberColumn("Utilisés"),
                    "Stock_Restant": st.column_config.NumberColumn("Stock Restant"),
                }
            )
            if "Stock_Restant" in df_s.columns and df_s["Stock_Restant"].sum() > 0:
                fig = px.bar(df_s.groupby(["Couleur","Format"])["Stock_Restant"].sum().reset_index(),
                             x="Couleur", y="Stock_Restant", color="Format",
                             title="Stock sachets par couleur", barmode="group")
                fig.update_layout(height=260, margin=dict(t=40,b=0))
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.subheader("➕ Nouvel achat de sachets")
        with st.form("form_sachet", clear_on_submit=True):
            c1, c2, c3, c4 = st.columns(4)
            with c1: gam_s2 = st.selectbox("Gamme",   GAMMES,   key="gs2")
            with c2: coul_s = st.selectbox("Couleur", COULEURS)
            with c3: fmt_s  = st.selectbox("Format",  FORMATS)
            with c4: date_s = st.date_input("Date", value=date.today(), key="ds")
            n1, n2, n3 = st.columns(3)
            with n1: qte_ach  = st.number_input("Qté achetée",   min_value=0, value=0)
            with n2: qte_util = st.number_input("Qté utilisée",  min_value=0, value=0)
            with n3: stock_r  = st.number_input("Stock restant", min_value=0, value=0)
            notes_s = st.text_input("Notes", key="ns")
            if st.form_submit_button("💾 Enregistrer", use_container_width=True, type="primary"):
                append("Sachets", [date_s.strftime("%d/%m/%Y"), gam_s2, coul_s, fmt_s,
                                   qte_ach, qte_util, stock_r, notes_s])
                bust(); st.success("✅ Enregistré !"); st.rerun()


# ─── Main ──────────────────────────────────────────────────────
def main():
    page = sidebar_nav()

    if page == "dashboard":
        page_dashboard()
    elif page == "new_order":
        page_new_order()
    elif page == "orders":
        page_orders()
    elif page == "stock":
        page_stock()
    elif page == "production":
        page_production()
    elif page == "sachets":
        page_sachets()
    elif page == "livreurs":
        page_livreurs()

if __name__ == "__main__":
    main()
