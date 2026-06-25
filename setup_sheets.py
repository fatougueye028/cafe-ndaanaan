"""
setup_sheets.py — Script d'initialisation du Google Sheet
Exécute ce script UNE SEULE FOIS pour créer la structure de base.

Usage :
    python setup_sheets.py
"""

import gspread
from google.oauth2.service_account import Credentials
import json, sys

# ── Charger les credentials ─────────────────────────────────────
try:
    with open(".streamlit/secrets.toml") as f:
        import re
        content = f.read()
    # Extraction simple du JSON de compte de service
    # Pour un vrai setup, utilise toml.load()
    print("ℹ️  Assure-toi d'avoir configuré .streamlit/secrets.toml")
except FileNotFoundError:
    print("❌ Fichier .streamlit/secrets.toml introuvable.")
    print("   Copie .streamlit/secrets.toml.example → .streamlit/secrets.toml et remplis les valeurs.")
    sys.exit(1)

# ── Connexion ───────────────────────────────────────────────────
import toml
secrets = toml.load(".streamlit/secrets.toml")

scopes = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
creds = Credentials.from_service_account_info(secrets["gcp_service_account"], scopes=scopes)
gc    = gspread.authorize(creds)

SHEET_NAME = secrets["SHEET_NAME"]

# ── Créer ou ouvrir le spreadsheet ─────────────────────────────
try:
    sh = gc.open(SHEET_NAME)
    print(f"✅ Google Sheet '{SHEET_NAME}' existant trouvé.")
except gspread.SpreadsheetNotFound:
    sh = gc.create(SHEET_NAME)
    sh.share(None, perm_type='anyone', role='writer')  # Partage optionnel
    print(f"✅ Google Sheet '{SHEET_NAME}' créé.")

# ── Définition des onglets et leurs en-têtes ───────────────────
SHEETS = {
    "Commandes": [
        "Date", "ID", "Client", "Téléphone", "Source", "Localisation", "Dépôt",
        "Gamme", "Format", "Quantité", "Prix_Unitaire", "CA", "Devise",
        "Type_Demande", "Statut_Livraison", "Statut_Paiement", "Lot", "Commentaire", "Date_Prevue", "Offre_Commerciale",
    ],
    "Stock": [
        "Date", "Lot", "Gamme", "Format",
        "Unites_Produites", "Unites_Commandees", "Unites_Vendues", "Stock_Restant",
        "Commentaire",
    ],
    "Livreurs": [
        "Livreur", "Date", "Zone", "Tarif", "Courses", "Montant",
    ],
    "Transporteurs": [
        "Nom", "Type", "Zone", "Tarif", "Telephone", "Delai", "Notes",
    ],
    "Depots": [
        "ID", "Nom", "Responsable", "Localisation", "Notes",
    ],
    "Stock_Depots": [
        "Depot", "Lot", "Gamme", "Format", "Stock_Restant", "Derniere_MAJ",
    ],
    "Mouvements_Stock": [
        "Date", "ID_Mouvement", "Depot_Origine", "Depot_Destination",
        "Gamme", "Format", "Quantite", "Statut", "Commentaire",
    ],
    "Sachets": [
        "Date", "Gamme", "Couleur", "Format",
        "Qte_Achetee", "Qte_Utilisee", "Stock_Restant", "Notes",
    ],
    "Affiches": [
        "Date", "Gamme", "Format",
        "Qte_Imprimee", "Qte_Utilisee", "Stock_Restant", "Notes",
    ],
    "Utilisateurs": [
        "ID", "Nom", "Email", "Téléphone", "Mot_de_passe", "Rôle", "Dépôt", "Actif", "Notes",
    ],
    "Production": [
        "Lot", "Gamme", "Date",
        "Cafe_Brut_kg", "Prix_Cafe_Brut", "Cout_Cafe_Brut",
        "Jar_kg", "Prix_Jar", "Cout_Jar",
        "Clous_FCFA", "Poivre_FCFA", "Gingembre_FCFA",
        "Frais_Torrefaction", "Frais_Transport", "Sachets_FCFA",
        "Main_Oeuvre", "Affiches_FCFA", "Emballage_FCFA", "Marketing_FCFA",
        "Cafe_Net_kg", "Cout_Total", "Cout_Revient_kg",
        "Qte_250g", "Qte_500g", "Qte_1kg", "Notes",
    ],
}

existing = [ws.title for ws in sh.worksheets()]

for sheet_name, headers in SHEETS.items():
    if sheet_name not in existing:
        ws = sh.add_worksheet(title=sheet_name, rows=1000, cols=len(headers))
        print(f"  ➕ Onglet '{sheet_name}' créé.")
    else:
        ws = sh.worksheet(sheet_name)
        print(f"  ✔  Onglet '{sheet_name}' existant.")

    # Écrire les en-têtes si la feuille est vide
    if not ws.row_values(1):
        ws.append_row(headers)
        print(f"     ✅ En-têtes écrits.")

print("\nℹ️  Stock : lance migrate_stock.py pour importer l'historique depuis l'Excel.")

# ── Supprimer la feuille par défaut ────────────────────────────
try:
    default_ws = sh.worksheet("Feuille 1")
    sh.del_worksheet(default_ws)
    print("  🗑  Feuille par défaut supprimée.")
except:
    pass

print("\n🎉 Google Sheet prêt ! URL :")
print(f"   https://docs.google.com/spreadsheets/d/{sh.id}/edit")
