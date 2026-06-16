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
STATUTS_LIV     = ["À préparer", "Préparée", "Livrée", "Annulée"]
STATUTS_PAY     = ["Non payé", "Partiel", "Payé"]
SOURCES         = ["WhatsApp", "Facebook", "TikTok", "YouTube",
                   "Recommandation", "Famille", "Autre"]

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

@st.cache_data(ttl=30)
def load(sheet: str) -> pd.DataFrame:
    records = _ws(sheet).get_all_records()
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

def next_id(df: pd.DataFrame) -> str:
    if df.empty or "ID" not in df.columns:
        return "CMD-001"
    nums = []
    for v in df["ID"].dropna().astype(str):
        if v.startswith("CMD-"):
            try:
                nums.append(int(v.split("-")[1]))
            except ValueError:
                pass
    return f"CMD-{(max(nums) + 1 if nums else 1):03d}"

# ─── Navigation ────────────────────────────────────────────────
PAGES = {
    "🏠 Dashboard":         "dashboard",
    "➕ Nouvelle commande":  "new_order",
    "📋 Commandes":         "orders",
    "📦 Stock":             "stock",
    "🏭 Production":        "production",
    "🚚 Livreurs":          "livreurs",
}

def sidebar_nav() -> str:
    st.sidebar.markdown(
        "<h2 style='text-align:center;font-size:1.4rem;margin-bottom:0'>☕ Kafe Ndaanaan</h2>"
        "<p style='text-align:center;font-size:0.75rem;opacity:0.7;margin-top:2px'>Gestion & Pilotage</p>",
        unsafe_allow_html=True,
    )
    st.sidebar.markdown("---")
    return PAGES[st.sidebar.radio("", list(PAGES.keys()), label_visibility="collapsed")]

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
    en_attente = df[df["Statut_Livraison"].isin(["À préparer", "Préparée"])]
    non_payes  = df[df["Statut_Paiement"].isin(["Non payé", "Partiel"])]["CA"]

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: kpi(f"{ca_fcfa:,.0f}", "CA FCFA total")
    with c2: kpi(f"{ca_eur:,.2f} €", "CA France total")
    with c3: kpi(str(len(en_attente)), "Commandes en attente")
    with c4: kpi(f"{non_payes.sum():,.0f}", "Paiements à recevoir (FCFA)")
    with c5: kpi(str(len(df[df["Statut_Livraison"] == "Livrée"])), "Commandes livrées")

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
        df_stock["Stock_Dispo"] = pd.to_numeric(df_stock["Stock_Dispo"], errors="coerce").fillna(0)
        critique = df_stock[df_stock["Stock_Dispo"] == 0]
        faible   = df_stock[(df_stock["Stock_Dispo"] > 0) & (df_stock["Stock_Dispo"] <= 5)]
        if not critique.empty:
            st.markdown(
                f'<div class="alert-rouge">🔴 <b>{len(critique)} références en rupture de stock</b></div>',
                unsafe_allow_html=True,
            )
            st.dataframe(critique[["Gamme", "Format", "Localisation", "Stock_Dispo"]], hide_index=True)
        if not faible.empty:
            st.markdown(
                f'<div class="alert-orange">🟠 <b>{len(faible)} références avec stock ≤ 5</b></div>',
                unsafe_allow_html=True,
            )
            st.dataframe(faible[["Gamme", "Format", "Localisation", "Stock_Dispo"]], hide_index=True)
        if critique.empty and faible.empty:
            st.success("✅ Tous les stocks sont OK.")
    else:
        st.info("Aucun stock renseigné.")

# ─── Page : Nouvelle Commande ──────────────────────────────────
def page_new_order():
    st.title("➕ Nouvelle Commande")

    df_cmd = load("Commandes")

    with st.form("form_cmd", clear_on_submit=True):
        st.subheader("Client")
        c1, c2 = st.columns(2)
        with c1:
            client  = st.text_input("Nom *", placeholder="Aminata Diallo")
            tel     = st.text_input("Téléphone", placeholder="77 123 45 67")
            source  = st.selectbox("Comment il/elle a découvert Kafe Ndaanaan ?", SOURCES)
        with c2:
            zone    = st.selectbox("Zone *", ZONES)
            adresse = st.text_input("Adresse (France/Canada)", placeholder="12 rue des Lilas, 75010 Paris")
            comm    = st.text_area("Commentaire / Notes", height=70)

        st.subheader("Produit commandé")
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
        st.info(f"💰 **Montant total : {ca:,.2f} {devise}**")

        lot = st.text_input("N° de Lot (optionnel)", placeholder="Lot 4")

        submitted = st.form_submit_button("✅ Enregistrer", use_container_width=True, type="primary")

    if submitted:
        if not client.strip():
            st.error("Le nom du client est obligatoire.")
        else:
            new_id = next_id(df_cmd)
            append("Commandes", [
                date.today().strftime("%d/%m/%Y"),
                new_id, client.strip(), tel, adresse, zone,
                gamme, fmt, qty, prix, ca, devise,
                "À préparer", "Non payé",
                source, lot, comm,
            ])
            st.success(f"✅ Commande **{new_id}** enregistrée !")
            st.balloons()

# ─── Page : Commandes ──────────────────────────────────────────
def page_orders():
    st.title("📋 Commandes")

    df = load("Commandes")
    if df.empty:
        st.info("Aucune commande enregistrée.")
        return

    # ── Filtres ──
    with st.expander("🔍 Filtres", expanded=True):
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            fz = st.multiselect("Zone", ZONES, default=ZONES)
        with f2:
            fl = st.multiselect("Livraison", STATUTS_LIV, default=STATUTS_LIV)
        with f3:
            fp = st.multiselect("Paiement", STATUTS_PAY, default=STATUTS_PAY)
        with f4:
            fg = st.multiselect("Gamme", GAMMES, default=GAMMES)

    mask = (
        df["Zone"].isin(fz)
        & df["Statut_Livraison"].isin(fl)
        & df["Statut_Paiement"].isin(fp)
        & df["Gamme"].isin(fg)
    )
    df_f = df[mask].copy()
    df_f["CA"] = pd.to_numeric(df_f["CA"], errors="coerce").fillna(0)

    st.caption(f"{len(df_f)} commande(s) — CA filtré : "
               f"{df_f[df_f['Devise']=='FCFA']['CA'].sum():,.0f} FCFA | "
               f"{df_f[df_f['Devise']=='EUR']['CA'].sum():,.2f} €")

    COLS = ["Date","ID","Client","Zone","Gamme","Format","Quantité","CA","Devise",
            "Statut_Livraison","Statut_Paiement","Source"]
    COLS = [c for c in COLS if c in df_f.columns]

    st.dataframe(
        df_f[COLS],
        use_container_width=True,
        hide_index=True,
        column_config={
            "CA":               st.column_config.NumberColumn(format="%.0f"),
            "Statut_Livraison": st.column_config.SelectboxColumn(options=STATUTS_LIV),
            "Statut_Paiement":  st.column_config.SelectboxColumn(options=STATUTS_PAY),
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

        # Décrémenter stock si passage à "Livrée"
        if new_liv == "Livrée" and row["Statut_Livraison"] != "Livrée":
            _decrement_stock(row)

        bust()
        st.success(f"✅ Commande {sel_id} mise à jour !")
        st.rerun()

def _decrement_stock(row):
    """Réduit le stock dispo lors du passage d'une commande à 'Livrée'."""
    try:
        df_s = load("Stock")
        if df_s.empty:
            return
        zone = str(row.get("Zone", "Dakar"))
        loc  = zone if zone in LOCALISATIONS else "Dakar"
        mask = (
            (df_s["Gamme"]        == row["Gamme"])
            & (df_s["Format"]     == row["Format"])
            & (df_s["Localisation"] == loc)
        )
        if not mask.any():
            return
        idx     = df_s.index[mask][0]
        current = int(pd.to_numeric(df_s.loc[idx, "Stock_Dispo"], errors="coerce") or 0)
        qty     = int(pd.to_numeric(row.get("Quantité", 0), errors="coerce") or 0)
        new_val = max(0, current - qty)
        ws      = _ws("Stock")
        col_s   = list(df_s.columns).index("Stock_Dispo") + 1
        col_m   = list(df_s.columns).index("Derniere_MAJ") + 1
        ws.update_cell(idx + 2, col_s, new_val)
        ws.update_cell(idx + 2, col_m, date.today().strftime("%d/%m/%Y"))
        bust()
    except Exception as e:
        st.warning(f"⚠️ Stock non décrémenté automatiquement : {e}")

# ─── Page : Stock ──────────────────────────────────────────────
def page_stock():
    st.title("📦 Suivi des Stocks")

    df_s = load("Stock")

    if not df_s.empty:
        df_s["Stock_Dispo"] = pd.to_numeric(df_s["Stock_Dispo"], errors="coerce").fillna(0)

        f1, f2 = st.columns(2)
        with f1:
            loc_f = st.multiselect("Localisation", LOCALISATIONS, default=LOCALISATIONS)
        with f2:
            gam_f = st.multiselect("Gamme", GAMMES, default=GAMMES)

        mask = df_s["Localisation"].isin(loc_f) & df_s["Gamme"].isin(gam_f)
        df_f = df_s[mask].copy()

        def style_stock(val):
            if val <= 0:
                return "background-color:#fde8e8; color:#c0392b; font-weight:bold"
            if val <= 5:
                return "background-color:#fff3cd; color:#e67e22; font-weight:bold"
            return "background-color:#eafaf1; color:#27ae60"

        styled = df_f.style.applymap(style_stock, subset=["Stock_Dispo"])
        st.dataframe(styled, use_container_width=True, hide_index=True)

        # Graph
        fig = px.bar(
            df_f.groupby(["Gamme", "Format"])["Stock_Dispo"].sum().reset_index(),
            x="Gamme", y="Stock_Dispo", color="Format",
            title="Stock disponible par gamme", barmode="group",
        )
        fig.update_layout(height=300, margin=dict(t=40, b=0))
        st.plotly_chart(fig, use_container_width=True)

    # ── Mise à jour manuelle ──
    st.markdown("---")
    st.subheader("✏️ Modifier / Ajouter un stock")

    with st.form("form_stock", clear_on_submit=True):
        s1, s2, s3 = st.columns(3)
        with s1:
            gam_s = st.selectbox("Gamme", GAMMES, key="sg")
        with s2:
            fmt_s = st.selectbox("Format", FORMATS, key="sf")
        with s3:
            loc_s = st.selectbox("Localisation", LOCALISATIONS, key="sl")

        nv_stock = st.number_input("Stock disponible (unités)", min_value=0, value=0)

        if st.form_submit_button("💾 Mettre à jour", use_container_width=True, type="primary"):
            df_s_full = load("Stock")
            m = (
                (df_s_full["Gamme"]        == gam_s)
                & (df_s_full["Format"]     == fmt_s)
                & (df_s_full["Localisation"] == loc_s)
            )
            if not df_s_full.empty and m.any():
                idx   = df_s_full.index[m][0]
                ws    = _ws("Stock")
                col_s = list(df_s_full.columns).index("Stock_Dispo") + 1
                col_m = list(df_s_full.columns).index("Derniere_MAJ") + 1
                ws.update_cell(idx + 2, col_s, nv_stock)
                ws.update_cell(idx + 2, col_m, date.today().strftime("%d/%m/%Y"))
                bust()
                st.success("✅ Stock mis à jour !")
            else:
                append("Stock", [gam_s, fmt_s, loc_s, nv_stock, date.today().strftime("%d/%m/%Y")])
                st.success("✅ Nouvelle ligne stock créée !")
            st.rerun()

# ─── Page : Livreurs ───────────────────────────────────────────
def page_livreurs():
    st.title("🚚 Livreurs")

    df = load("Livreurs")

    if not df.empty:
        df["Montant"] = pd.to_numeric(df["Montant"], errors="coerce").fillna(0)
        df["Courses"] = pd.to_numeric(df["Courses"], errors="coerce").fillna(0)

        resume = (
            df.groupby("Livreur")
            .agg(Courses=("Courses", "sum"), Montant_Total=("Montant", "sum"))
            .reset_index()
        )
        st.subheader("Résumé")
        st.dataframe(resume, use_container_width=True, hide_index=True)

        st.subheader("Historique des courses")
        st.dataframe(df, use_container_width=True, hide_index=True)

    # ── Enregistrer une course ──
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
                st.success("✅ Course enregistrée !")
                st.rerun()
            else:
                st.error("Le nom du livreur est obligatoire.")

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
    """Ajoute au stock quand un lot est produit."""
    if qty <= 0:
        return
    try:
        df_s = load("Stock")
        if df_s.empty:
            return
        # Localisation = Touba (production)
        mask = (
            (df_s["Gamme"]         == gamme)
            & (df_s["Format"]      == fmt)
            & (df_s["Localisation"] == "Touba")
        )
        if not mask.any():
            return
        idx     = df_s.index[mask][0]
        current = int(pd.to_numeric(df_s.loc[idx, "Stock_Dispo"], errors="coerce") or 0)
        ws      = _ws("Stock")
        col_s   = list(df_s.columns).index("Stock_Dispo") + 1
        col_m   = list(df_s.columns).index("Derniere_MAJ") + 1
        ws.update_cell(idx + 2, col_s, current + qty)
        ws.update_cell(idx + 2, col_m, date.today().strftime("%d/%m/%Y"))
        bust()
    except Exception as e:
        st.warning(f"⚠️ Stock non incrémenté pour {gamme} {fmt} : {e}")


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
    elif page == "livreurs":
        page_livreurs()

if __name__ == "__main__":
    main()
