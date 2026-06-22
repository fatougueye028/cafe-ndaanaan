"""
init_depots.py — Initialise les dépôts et le stock initial par dépôt.

Les 4 dépôts sont créés : Dakar (hub central), Touba, France, Partenaire France.
Le stock initial est basé sur le stock actuel de l'onglet Stock,
attribué au dépôt Dakar (stock de référence).

Usage :
    source venv/bin/activate
    python3 init_depots.py
"""

import gspread, toml
from google.oauth2.service_account import Credentials
from datetime import date
import os

SECRETS = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
secrets = toml.load(SECRETS)
scopes  = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds   = Credentials.from_service_account_info(secrets["gcp_service_account"], scopes=scopes)
gc      = gspread.authorize(creds)
sh      = gc.open(secrets["SHEET_NAME"])

today = date.today().strftime("%d/%m/%Y")

# ── 1. Créer les dépôts ─────────────────────────────────────────
print("=== Initialisation des dépôts ===")
ws_dep = sh.worksheet("Depots")
existing_dep = {r.get("Nom","") for r in ws_dep.get_all_records()}

DEPOTS = [
    {"ID": "DEP-001", "Nom": "Dakar",            "Responsable": "Fatou Gueye",  "Localisation": "Dakar, Sénégal",  "Notes": "Hub central — stock de référence"},
    {"ID": "DEP-002", "Nom": "Touba",            "Responsable": "",             "Localisation": "Touba, Sénégal",  "Notes": "Production et dépôt Touba"},
    {"ID": "DEP-003", "Nom": "France",           "Responsable": "",             "Localisation": "France",          "Notes": "Dépôt France"},
]

headers_dep = ws_dep.row_values(1)
new_deps = [d for d in DEPOTS if d["Nom"] not in existing_dep]
if new_deps:
    rows = [[d.get(h, "") for h in headers_dep] for d in new_deps]
    ws_dep.append_rows(rows, value_input_option="RAW")
    print(f"  ✅ {len(new_deps)} dépôt(s) créé(s)")
else:
    print("  ℹ️  Dépôts déjà présents")

# ── 2. Initialiser le stock par dépôt depuis le stock actuel ────
print("\n=== Initialisation du stock par dépôt (Dakar) ===")
ws_stock    = sh.worksheet("Stock")
ws_sd       = sh.worksheet("Stock_Depots")
headers_sd  = ws_sd.row_values(1)

# Lire le stock actuel
rows_stock = ws_stock.get_all_records(value_render_option="UNFORMATTED_VALUE")
existing_sd = {(r.get("Depot",""), r.get("Lot",""), r.get("Gamme",""), r.get("Format",""))
               for r in ws_sd.get_all_records()}

new_sd_rows = []
for r in rows_stock:
    lot    = str(r.get("Lot",   "")).strip()
    gamme  = str(r.get("Gamme", "")).strip()
    fmt    = str(r.get("Format","")).strip()
    restant = int(r.get("Stock_Restant", 0) or 0)

    if not gamme or not fmt:
        continue

    # Stock de Dakar = stock actuel de la feuille Stock
    if ("Dakar", lot, gamme, fmt) not in existing_sd:
        new_sd_rows.append({
            "Depot": "Dakar", "Lot": lot, "Gamme": gamme, "Format": fmt,
            "Stock_Restant": restant, "Derniere_MAJ": today
        })

    # Autres dépôts initialisés à 0 (seulement pour le dernier lot)
    for depot in ["Touba", "France"]:
        if (depot, lot, gamme, fmt) not in existing_sd:
            new_sd_rows.append({
                "Depot": depot, "Lot": lot, "Gamme": gamme, "Format": fmt,
                "Stock_Restant": 0, "Derniere_MAJ": today
            })

if new_sd_rows:
    rows = [[r.get(h, "") for h in headers_sd] for r in new_sd_rows]
    ws_sd.append_rows(rows, value_input_option="RAW")
    print(f"  ✅ {len(new_sd_rows)} ligne(s) stock initialisées")
    dakar_rows = [r for r in new_sd_rows if r["Depot"] == "Dakar" and r["Stock_Restant"] > 0]
    for r in dakar_rows:
        print(f"    Dakar | {r['Gamme']} {r['Format']} → {r['Stock_Restant']} unités")
else:
    print("  ℹ️  Stock dépôts déjà initialisé")

print(f"\n🎉 Initialisation terminée !")
print(f"📊 https://docs.google.com/spreadsheets/d/{sh.id}/edit")
