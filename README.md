# MCP Office – modèles d'entreprise pour LibreChat

Serveur MCP (HTTP streamable) qui transforme les réponses de l'IA en **Word, Excel, PowerPoint**, toujours à partir du **modèle le plus récent** du dossier SharePoint (styles, thème, masques, en-têtes/pieds de page conservés).

## Outils exposés
| Outil | Entrée | Sortie |
|---|---|---|
| `create_word` | `title`, `markdown` (titres, listes, tableaux, gras/italique, code) | .docx |
| `create_excel` | `title`, `sheets:[{name, rows}]` (1re ligne = en-têtes → tableau Excel) | .xlsx |
| `list_slide_types` | – | diapos types du modèle PowerPoint (n°, libellé, zones) |
| `create_powerpoint` | `title`, `subtitle`, `slides:[{type, title, bullets, bullets2, table, zones, notes}]` | .pptx |
| `list_templates` | – | modèles disponibles / utilisé |

`template_prefix` (optionnel) permet de choisir une famille de modèles (ex. `Note_`, `Rapport_`). Le fichier retourné est un lien de téléchargement à usage temporaire (`FILE_TTL`).

## Fonctionnement
- **Dernier modèle** : liste le dossier via Microsoft Graph, prend le fichier `.dotx/.docx`, `.xltx/.xlsx`, `.potx/.pptx` le plus récemment modifié ; cache invalidé par eTag.
- **Style par défaut** : Word → styles `Title`, `Heading n`, `List Bullet/Number`, `Table Grid` du modèle (repli si absents) ; le contenu d'exemple du modèle est supprimé, la mise en page conservée. Excel → 1re feuille du modèle dupliquée, données sous l'en-tête existant. PowerPoint → voir ci-dessous.

### Modèle PowerPoint à diapos types
Le modèle `.pptx` contient des diapos déjà préparées (couverture, sommaire, 2 colonnes, tableau, contact…). Pour chaque diapo demandée, le serveur **duplique la diapo type** (décor, images, logos, fond, mise en forme conservés), remplit ses zones, puis supprime les diapos types d'origine.
- Diapo 1 = couverture (`title`, `subtitle`).
- `type` : n° ou libellé (titre de la diapo type) ; omis → 1re diapo avec tableau ou corps de texte.
- `title` → titre ; `bullets` / `bullets2` → 1er / 2e espace de contenu (indentation de 2 espaces = niveau inférieur) ; `table` → tableau existant (lignes ajoutées/supprimées en gardant le style) ou nouveau tableau.
- `zones` → remplit toute forme par son **nom** (volet Sélection de PowerPoint), ex. `{"Contact":"…"}`. Conseil : nommer les zones du modèle de façon explicite.
- Les espaces de contenu non remplis sont supprimés (pas de texte d'exemple résiduel) ; les zones de texte hors espaces réservés sont conservées.
- **Accessibilité** : titre et langue du document renseignés, vrais styles de titres, en-tête de tableau répété (Word), tableaux Excel structurés, titre sur chaque diapo, notes orateur.

## Sécurité
- App Entra ID en **`Sites.Selected`** (lecture seule, un seul site) – pas de `Sites.Read.All`.
- `/mcp` protégé par `Authorization: Bearer $MCP_API_KEY` (comparaison à temps constant).
- Liens de fichiers non devinables (192 bits), expirés après `FILE_TTL`, `no-store`, `nosniff`, sans traversée de chemin.
- Taille d'entrée limitée (`MAX_INPUT`), conteneur non-root. À exposer uniquement derrière le reverse-proxy HTTPS interne.

Accorder l'accès au site (admin, une fois) :
```http
POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
{"roles":["read"],"grantedToIdentities":[{"application":{"id":"<client-id>","displayName":"mcp-office"}}]}
```
`SP_SITE_ID` : `GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/<NomDuSite>` → `id`.

## Lancement
```bash
cp .env.example .env   # compléter
docker build -t mcp-office . && docker run -d --env-file .env -p 8000:8000 mcp-office
```

## Configuration LibreChat (`librechat.yaml`)
```yaml
mcpServers:
  office:
    type: streamable-http
    url: http://mcp-office:8000/mcp
    headers:
      Authorization: "Bearer ${MCP_OFFICE_KEY}"
    timeout: 60000
    serverInstructions: |
      Quand l'utilisateur demande un Word, Excel ou PowerPoint, appelle l'outil correspondant
      (pour PowerPoint, appelle d'abord list_slide_types pour choisir les diapos types)
      et renvoie le lien de téléchargement tel quel.
```
Si l'instance LibreChat est hébergée par un tiers, demander à l'administrateur d'ajouter ce bloc (ou d'autoriser les serveurs MCP utilisateurs) et d'ouvrir le flux réseau vers le serveur.

## Test local sans SharePoint
```bash
pip install -r requirements.txt
TEMPLATE_DIR=./modeles MCP_API_KEY=test python server.py
```
