"""
migrate_sachets.py — Import des stocks Sachets et Affiches depuis Excel

2 tableaux dans "4. SUIVI STOCKS" (lignes 27-37) :
  - Gauche (cols 1-4)  : Stock Affiches → Date | Gamme  | Format | Unité
  - Droite (cols 6-9+) : Stock Sachets  → Date | Couleur | Format | Unité (+ col 11 = restants)

Usage :
    source venv/bin/activate
    python3 migrate_sachets.py
"""

import openpyxl, gspread, toml
from google.oauth2.service_account import Credentials
from datetime import datetime
import os

EXCEL_PATH   = os.path.join(os.path.dirname(__file__), "..", "Gestion_Cafe_Ndaanaan (1).xlsx")
SECRETS_PATH = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")

secrets = toml.load(SECRETS_PATH)
scopes  = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds   = Credentials.from_service_account_info(secrets["gcp_service_account"], scopes=scopes)
gc      = gspread.authorize(creds)
sh      = gc.open(secrets["SHEET_NAME"])

FORMAT_MAP = {"1 kg": "1kg", "500 g": "500g", "250 g": "250g"}
GAMME_MAP  = {"Epicé": "Épicé", "Epice": "Épicé", "Nooket": "Ñooket"}
COULEUR_GAMME = {
    "Blanc":    "Signature",
    "Noir":     "Prestige",
    "Doré":     "Original",
    "Doré vif": "Ñooket",
    "Argenté":  "Épicé",
}

def norm(val, mapping): return mapping.get(str(val).strip(), str(val).strip())
def safe_int(v):
    try: return int(float(v)) if v is not None else 0
    except: return 0
def fmt_date(v):
    if isinstance(v, datetime): return v.strftime("%d/%m/%Y")
    return str(v) if v else ""

wb    = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
ws_xl = wb["4. SUIVI STOCKS"]

# ── STOCK AFFICHES (tableau gauche : cols 1-4) ─────────────────
print("=== Import Stock Affiches (par gamme) ===")
ws_aff = sh.worksheet("Affiches")
rows_aff = []

for r in range(28, 38):
    date_v = ws_xl.cell(r, 1).value
    gamme  = ws_xl.cell(r, 2).value
    fmt    = ws_xl.cell(r, 3).value
    qte    = ws_xl.cell(r, 4).value

    if not gamme: continue

    gamme_s = norm(gamme, GAMME_MAP)
    fmt_s   = norm(fmt, FORMAT_MAP)
    stock   = safe_int(qte)

    rows_aff.append([
        fmt_date(date_v), gamme_s, fmt_s,
        stock,  # Qte_Imprimee
        0,      # Qte_Utilisee
        stock,  # Stock_Restant
        "",
    ])
    print(f"  ✔  {gamme_s} {fmt_s} → Stock: {stock}")

if rows_aff:
    ws_aff.append_rows(rows_aff, value_input_option="RAW")
    print(f"  ✅ {len(rows_aff)} affiches importées.")

# ── STOCK SACHETS (tableau droite : cols 6-9, restants col 11) ─
print("\n=== Import Stock Sachets (par couleur) ===")
ws_sac = sh.worksheet("Sachets")
rows_sac = []

for r in range(28, 38):
    date_v  = ws_xl.cell(r, 6).value
    couleur = ws_xl.cell(r, 7).value
    fmt     = ws_xl.cell(r, 8).value
    qte_ach = ws_xl.cell(r, 9).value
    restant = ws_xl.cell(r, 11).value

    if not couleur: continue

    couleur_s = str(couleur).strip()
    fmt_s     = norm(fmt, FORMAT_MAP)
    qte_a     = safe_int(qte_ach)
    rest_s    = safe_int(restant) if restant is not None else qte_a
    util_s    = max(0, qte_a - rest_s)

    gamme_s = COULEUR_GAMME.get(couleur_s, "")
    rows_sac.append([
        fmt_date(date_v), gamme_s, couleur_s, fmt_s,
        qte_a,   # Qte_Achetee
        util_s,  # Qte_Utilisee
        rest_s,  # Stock_Restant
        "",
    ])
    print(f"  ✔  {gamme_s} / {couleur_s} {fmt_s} → Achetés:{qte_a} Utilisés:{util_s} Restants:{rest_s}")

if rows_sac:
    ws_sac.append_rows(rows_sac, value_input_option="RAW")
    print(f"  ✅ {len(rows_sac)} sachets importés.")

print(f"\n🎉 Import terminé !")
print(f"📊 https://docs.google.com/spreadsheets/d/{sh.id}/edit")
