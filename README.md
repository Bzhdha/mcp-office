# MCP Office – modèles d'entreprise pour LibreChat

Serveur MCP (HTTP streamable) qui transforme les réponses de l'IA en **Word, Excel, PowerPoint**, toujours à partir du **modèle le plus récent** du dossier SharePoint (styles, thème, masques, en-têtes/pieds de page conservés).

## Outils exposés
| Outil | Rôle |
|---|---|
| `create_word` | .docx depuis du Markdown (titres, listes, tableaux, gras/italique, code) |
| `create_excel` | .xlsx depuis `sheets:[{name, rows}]` (1re ligne = en-têtes → tableau Excel) |
| `list_slide_types` | catalogue des dispositions du thème (`L n°`) ; avec `slides="L4,L13"` : zones de chaque disposition |
| `layout_catalog` | .pptx « catalogue » : une diapo par disposition, zones étiquetées `#id` |
| `create_powerpoint` | .pptx diapo par diapo, chacune sur une disposition du thème ; `dry_run` pour vérifier le plan |
| `list_templates` | modèles disponibles et modèle utilisé par défaut |

`template` (optionnel, tous les outils) : partie du nom du modèle (ex. `Theme`, `Proposition`).

## Choix du modèle
- Dans une même famille, le serveur prend la version la plus récente (année dans le nom, puis `vX.Y`, puis date de modification). Le cache est invalidé par eTag.
- Modèle par défaut de chaque type : `DEFAULT_TEMPLATE_DOCX|XLSX|PPTX` (partie du nom ; pour PowerPoint, `theme` par défaut).

## PowerPoint strictement conforme au thème
Le plan de la présentation se construit dans la conversation LibreChat ; le serveur garantit la conformité :
1. L'IA définit avec l'utilisateur l'objectif et le plan, puis propose pour chaque diapo la **disposition du thème** adaptée (chiffres clés, colonnes, équipe, citation, référence client…).
2. `list_slide_types("L4,L14,…")` donne les zones de chaque disposition, triées haut → bas et gauche → droite : `#id`, type, position (% de la diapo), capacité estimée (≤ lignes × caractères), texte d'invite = rôle attendu.
3. `create_powerpoint(dry_run=true)` vérifie le plan : texte trop long, zone inconnue, titre vide, image à insérer. Puis génération.

Garanties :
- chaque diapo est créée depuis une disposition du thème : polices, couleurs, positions, pieds de page (classification C2) et décor sont ceux du masque ;
- aucune forme libre : seules les zones prévues sont remplies (un tableau prend la place de la zone de contenu, avec le style de tableau du thème) ; seuls le gras et l'italique sont autorisés ;
- les zones non remplies sont supprimées : aucun texte d'invite ne reste ;
- les instructions du serveur MCP (envoyées au client) décrivent cette démarche à l'IA.

Options :
- `LAYOUT_GUIDE=/chemin/guide.json` : description d'usage par disposition, affichée dans le catalogue (ex. `{"Key Numbers":"3 chiffres clés avec légende"}`).
- Modèle à diapos préparées (ex. proposition commerciale) : `template="Proposition"`, diapos désignées par leur n° et dupliquées (décor, images, graphiques conservés). Les post-it de consigne (« EXEMPLE », « À REMPLIR », motif `PPTX_DROP`) sont supprimés. `PPTX_PREPARED_SLIDES=auto|on|off` (auto = tous les modèles sauf ceux dont le nom contient « theme »).

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
    serverInstructions: true   # utilise les instructions fournies par le serveur (démarche PowerPoint)
```
Si l'instance LibreChat est hébergée par un tiers, demander à l'administrateur d'ajouter ce bloc (ou d'autoriser les serveurs MCP utilisateurs) et d'ouvrir le flux réseau vers le serveur.

## Test local sans SharePoint
```bash
pip install -r requirements.txt
TEMPLATE_DIR=./modeles MCP_API_KEY=test python server.py
```
Ne pas versionner les modèles d'entreprise (documents classifiés C2) dans ce dépôt.
