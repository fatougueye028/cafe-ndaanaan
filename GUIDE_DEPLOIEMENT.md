# ☕ Kafe Ndaanaan — Guide de Déploiement

## Vue d'ensemble

L'application est construite avec **Streamlit** (Python) et utilise **Google Sheets** comme base de données.
Déploiement gratuit sur **Streamlit Cloud** — accessible depuis téléphone, tablette ou ordinateur.

---

## Structure Google Sheets

Le Google Sheet doit s'appeler exactement : **`Kafe Ndaanaan — Gestion`**

Il contient 3 onglets :

### Onglet `Commandes`
| Colonne | Description |
|---|---|
| Date | JJ/MM/AAAA |
| ID | CMD-001, CMD-002… |
| Client | Nom complet |
| Téléphone | Numéro |
| Adresse | Pour France/Canada |
| Zone | Touba / Dakar / France / Canada / Autre |
| Gamme | Signature / Original / Prestige / Épicé / Ñooket |
| Format | 250g / 500g / 1kg |
| Quantité | Nombre d'unités |
| Prix_Unitaire | Prix à l'unité |
| CA | Prix × Quantité |
| Devise | FCFA / EUR / CAD |
| Statut_Livraison | À préparer / Préparée / Livrée / Annulée |
| Statut_Paiement | Non payé / Partiel / Payé |
| Source | WhatsApp / Facebook / TikTok… |
| Lot | Numéro de lot de production |
| Commentaire | Notes libres |

### Onglet `Stock`
| Colonne | Description |
|---|---|
| Gamme | Version du café |
| Format | 250g / 500g / 1kg |
| Localisation | Touba / Dakar / France / Canada |
| Stock_Dispo | Unités disponibles |
| Derniere_MAJ | Date de dernière mise à jour |

### Onglet `Livreurs`
| Colonne | Description |
|---|---|
| Livreur | Nom du livreur |
| Date | Date de la course |
| Zone | Zone couverte |
| Tarif | Prix par course (FCFA) |
| Courses | Nombre de courses effectuées |
| Montant | Tarif × Courses |

---

## Étape 1 : Créer le compte de service Google

1. Va sur [console.cloud.google.com](https://console.cloud.google.com)
2. Crée un projet (ex: `kafe-ndaanaan`)
3. Active ces deux APIs :
   - **Google Sheets API**
   - **Google Drive API**
4. Va dans **IAM & Admin → Comptes de service**
5. Crée un compte de service (ex: `kafe-app@kafe-ndaanaan.iam.gserviceaccount.com`)
6. Génère une clé JSON → télécharge le fichier

---

## Étape 2 : Préparer le Google Sheet

1. Crée un nouveau Google Sheet sur [sheets.google.com](https://sheets.google.com)
2. Nomme-le : **`Kafe Ndaanaan — Gestion`**
3. **Partage** le sheet avec l'adresse email du compte de service (rôle : Éditeur)
4. Lance le script d'initialisation :

```bash
pip install gspread google-auth toml
# Place ton fichier JSON dans .streamlit/secrets.toml (voir secrets.toml.example)
python setup_sheets.py
```

---

## Étape 3 : Déployer sur Streamlit Cloud

### 3a — Mettre le code sur GitHub

1. Crée un dépôt GitHub **privé** (ex: `kafe-ndaanaan-app`)
2. Upload ces fichiers :
   ```
   app.py
   requirements.txt
   .streamlit/config.toml
   ```
   ⚠️ **NE PAS** uploader `secrets.toml` — le mettre dans `.gitignore`

### 3b — Créer l'application sur Streamlit Cloud

1. Va sur [share.streamlit.io](https://share.streamlit.io)
2. Connecte ton compte GitHub
3. Clique **New app** → sélectionne ton dépôt → fichier principal : `app.py`
4. Clique **Advanced settings → Secrets** et colle :

```toml
SHEET_NAME = "Kafe Ndaanaan — Gestion"

[gcp_service_account]
type = "service_account"
project_id = "kafe-ndaanaan"
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
client_email = "kafe-app@kafe-ndaanaan.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
```

5. Clique **Deploy** → l'app sera disponible à une URL du type :
   `https://kafe-ndaanaan-app.streamlit.app`

---

## Utilisation sur téléphone

- Ouvre l'URL dans Chrome ou Safari
- Dans Chrome → menu (⋮) → **Ajouter à l'écran d'accueil**
- L'app s'ouvre comme une application native

---

## Rôles de l'équipe

| Membre | Accès recommandé |
|---|---|
| Fatou (toi) | Tout — Dashboard, commandes, finances |
| Frère (Touba) | Stock Touba, production |
| Sœur (Dakar/France) | Commandes France + Dakar, paiements |
| Nièce | Stock, livraisons, livreurs |

> La gestion des rôles par login (mot de passe par utilisateur) peut être ajoutée en V2 via `streamlit-authenticator`.

---

## Évolutions prévues (V2)

- [ ] Authentification par rôle (Fatou / Frère / Sœur / Nièce)
- [ ] Export PDF des commandes
- [ ] Envoi automatique d'un résumé WhatsApp quotidien
- [ ] Gestion multi-lots de production
- [ ] Suivi des coûts de production et marges détaillées
- [ ] Notifications stock bas

---

## Support

Pour toute question technique, contact : fgueye@purse.eu
