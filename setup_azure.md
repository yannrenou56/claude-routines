# Configuration Azure AD — permissions requises

## Permissions Graph API nécessaires

Pour que `outlook_drafts.py` fonctionne, l'application Azure AD doit avoir ces permissions **application** (pas déléguées) :

| Permission | Type | Usage |
|---|---|---|
| `Mail.ReadWrite` | Application | Lire l'inbox + créer des brouillons |
| `Mail.Send` | Application | (déjà présent) Envoi des CR de réunion |
| `User.Read.All` | Application | (déjà présent) Résolution des participants |
| `Sites.Read.All` | Application | (déjà présent) Lecture SharePoint |

## Ajouter Mail.ReadWrite dans Azure

1. Allez sur **portal.azure.com** → Azure Active Directory → App registrations
2. Sélectionnez votre application (`claude-routines` ou son nom)
3. **API permissions** → Add a permission → Microsoft Graph → Application permissions
4. Cherchez `Mail.ReadWrite` → cochez → Add permissions
5. Cliquez **Grant admin consent** (bouton en haut de la liste)

## Variables .env à renseigner

```env
SENDER_EMAIL=yann.renou@olisma.fr
USER_DISPLAY_NAME=Yann RENOU
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
```

## Test

```bash
cd ~/claude-routines
cp .env.example .env
# Remplir les valeurs dans .env
python3 outlook_drafts.py --dry-run   # voir sans créer
python3 outlook_drafts.py             # créer les brouillons
```
