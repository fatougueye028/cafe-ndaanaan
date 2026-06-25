"""
init_users.py — Créer l'utilisateur admin initial et la feuille Utilisateurs.

Usage :
    cd kafe-ndaanaan-app
    source venv/bin/activate
    python3 init_users.py

Ce script :
  1. Crée l'onglet 'Utilisateurs' s'il n'existe pas
  2. Crée un compte Admin avec le mot de passe saisi
  3. Affiche un exemple pour créer d'autres utilisateurs

Format mot de passe stocké : sha256$<salt>$<hash>  (sécurisé, non réversible)

Pour ajouter d'autres utilisateurs :
  - Soit ajouter directement dans l'onglet Google Sheet 'Utilisateurs'
  - Soit relancer ce script avec un autre email
"""

import gspread, toml, os, hashlib, secrets
from google.oauth2.service_account import Credentials

# ── Connexion ───────────────────────────────────────────────────
SECRETS = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
secrets_data = toml.load(SECRETS)
scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds  = Credentials.from_service_account_info(secrets_data["gcp_service_account"], scopes=scopes)
gc     = gspread.authorize(creds)
sh     = gc.open(secrets_data["SHEET_NAME"])

HEADERS_USERS = ["ID", "Nom", "Email", "Téléphone", "Mot_de_passe", "Rôle", "Dépôt", "Actif", "Notes"]


def hash_pwd(password: str, salt: str = None) -> str:
    if salt is None:
        salt = secrets.token_hex(16)
    h = hashlib.sha256(f"{salt}{password}".encode()).hexdigest()
    return f"sha256${salt}${h}"


def get_or_create_sheet():
    try:
        ws = sh.worksheet("Utilisateurs")
        print("✅ Onglet Utilisateurs trouvé.")
    except Exception:
        ws = sh.add_worksheet("Utilisateurs", rows=200, cols=len(HEADERS_USERS))
        ws.append_row(HEADERS_USERS)
        print("✅ Onglet Utilisateurs créé avec les en-têtes.")
    return ws


def create_user(ws, uid, nom, email, tel, pwd, role, depot, notes=""):
    headers = ws.row_values(1)
    records = ws.get_all_records()
    existing = {str(r.get("Email","")).lower() for r in records}

    if email.lower() in existing:
        print(f"  ⚠️  Utilisateur {email} déjà présent — ignoré.")
        return

    user = {
        "ID":           uid,
        "Nom":          nom,
        "Email":        email,
        "Téléphone":    tel,
        "Mot_de_passe": hash_pwd(pwd),
        "Rôle":         role,
        "Dépôt":        depot,
        "Actif":        "1",
        "Notes":        notes,
    }
    row = [user.get(h, "") for h in headers]
    ws.append_row(row, value_input_option="RAW")
    print(f"  ✅ Utilisateur créé : {nom} ({email}) — Rôle: {role}{' · Dépôt: ' + depot if depot else ''}")


# ── Main ─────────────────────────────────────────────────────────
print("=== Initialisation des utilisateurs Kafe Ndaanaan ===\n")

ws_users = get_or_create_sheet()

print("\nCréation du compte Admin (Fatou Gueye)")
pwd_admin = input("Mot de passe admin (laisse vide pour 'ndaanaan2024') : ").strip()
if not pwd_admin:
    pwd_admin = "ndaanaan2024"
    print("  ⚠️  Mot de passe par défaut utilisé. Change-le après la première connexion !")

create_user(
    ws_users,
    uid="USR-001", nom="Fatou Gueye", email="admin@ndaanaan.com",
    tel="", pwd=pwd_admin, role="Admin", depot="",
    notes="Compte administrateur principal"
)

print("\n─────────────────────────────────────────")
print("Pour ajouter d'autres utilisateurs, modifie directement l'onglet")
print("'Utilisateurs' dans Google Sheets avec ces colonnes :")
print(f"  {', '.join(HEADERS_USERS)}")
print()
print("Rôles disponibles : Admin | Dakar | Touba | France | Partenaire")
print("Dépôts disponibles : Dakar | Touba | France  (vide pour Admin)")
print("Actif : 1 = actif, 0 = désactivé")
print()
print("Pour hasher un nouveau mot de passe, utilise ce code Python :")
print("  import hashlib, secrets")
print("  salt = secrets.token_hex(16)")
print("  h = hashlib.sha256(f'{salt}MON_MOT_DE_PASSE'.encode()).hexdigest()")
print("  print(f'sha256${salt}${h}')")
print()
print(f"🎉 Accède à l'app et connecte-toi avec :")
print(f"   Email : admin@ndaanaan.com")
print(f"   Mot de passe : {pwd_admin}")
print(f"📊 https://docs.google.com/spreadsheets/d/{sh.id}/edit#gid=0")
