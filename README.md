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
1. L'IA définit avec l'utilisateur l'objectif et le plan, puis propose pour chaque diapo une **famille de dispositions** adaptée au contenu (chiffres clés, colonnes, équipe, citation, intercalaire, référence client…) et la **déclinaison** voulue : couleur, nombre de colonnes, avec ou sans image ou description.
2. `list_slide_types("L4,L14,…")` donne les zones de chaque disposition, triées haut → bas et gauche → droite : `#id`, type, position (% de la diapo), capacité estimée (≤ lignes × caractères), texte d'invite = rôle attendu.
3. `create_powerpoint(dry_run=true)` vérifie le plan : texte trop long, zone inconnue, titre vide, image à insérer. Puis génération.

Garanties :
- chaque diapo est créée depuis une disposition du thème : polices, couleurs, positions, pieds de page (classification C2) et décor sont ceux du masque ;
- aucune forme libre : seules les zones prévues sont remplies (un tableau prend la place de la zone de contenu, avec le style de tableau du thème) ; seuls le gras et l'italique sont autorisés ;
- les zones non remplies sont supprimées : aucun texte d'invite ne reste ;
- les instructions du serveur MCP (envoyées au client) décrivent cette démarche à l'IA.

### Polices
Le serveur n'impose aucune police : le texte hérite des polices déclarées par le modèle (thème, masques, dispositions). Pour le thème Niji : titres N27 Medium, texte N27 Light, plus N27 et N27 Regular selon les dispositions. Le catalogue (`list_slide_types`) affiche les polices du modèle, et pour chaque zone sa police et sa taille effectives (ex. `[N27 Light 10.5pt]`).
- `**mots importants**` : écrits dans la police de mise en valeur du modèle, c'est-à-dire la variante plus grasse de la même famille déclarée dans le modèle (N27 Light → N27 Medium, conformément aux consignes de la charte), et non en gras synthétique. Si aucune variante n'existe, le gras est utilisé. `PPTX_EMPHASIS_FONT` permet d'imposer cette police.
- Les polices ne sont pas embarquées dans les modèles : elles doivent être installées sur les postes qui ouvrent les présentations.

### Familles et déclinaisons
Le catalogue regroupe les dispositions par famille : nom avant « - », sans préfixe `1_` ni numéros. Par exemple `Title + 2/3/4 Columns` → « Title + Columns » et `Detailed Section 2 - Magento/Red` → « Detailed Section ». Il indique si les déclinaisons ont les mêmes zones (seul le visuel change) ou non. Une diapo se désigne par `"type":"L31"`, ou par famille et déclinaison : `{"type":"Detailed Section","variant":"2 - Magento"}`. Une déclinaison inconnue renvoie la liste des déclinaisons possibles.

Options :
- `LAYOUT_GUIDE=/chemin/guide.json` : usage et/ou famille imposée par disposition, affichés dans le catalogue. Exemple :
  ```json
  {"Key Numbers": "3 chiffres clés avec légende",
   "One liners, Quotes, or small texts - Grayscale": {"famille": "One liners & Quotes", "usage": "citation ou message fort"},
   "Fiche Reference + Chiffres clés": {"famille": "Fiche Reference"}}
  ```
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

## Installation en local (poste de travail)

### 1. Prérequis
- Python 3.11 ou plus récent (`python --version`), ou Docker Desktop.
- Les modèles d'entreprise sur le poste. Le plus simple : synchroniser le dossier SharePoint des modèles avec OneDrive (bouton « Synchroniser » dans SharePoint) et pointer `TEMPLATE_DIR` dessus. Les modèles se mettent alors à jour tout seuls, sans configuration Entra ID.
- Pour voir les présentations avec la bonne police : polices N27 installées sur le poste.

### 2. Installation
```bash
git clone https://github.com/Bzhdha/mcp-office.git
cd mcp-office
python -m venv .venv
# macOS / Linux
source .venv/bin/activate
# Windows (PowerShell)
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 3. Lancement
macOS / Linux :
```bash
export TEMPLATE_DIR="$HOME/Library/CloudStorage/OneDrive-Niji/Modeles"   # dossier synchronisé (adapter le chemin)
export MCP_API_KEY="une-cle-longue-et-aleatoire"
export PUBLIC_BASE_URL="http://localhost:8000"
python server.py
```
Windows (PowerShell) :
```powershell
$env:TEMPLATE_DIR="$env:USERPROFILE\Niji\Modeles - Documents"   # dossier synchronisé (adapter le chemin)
$env:MCP_API_KEY="une-cle-longue-et-aleatoire"
$env:PUBLIC_BASE_URL="http://localhost:8000"
python server.py
```
Pour générer une clé : `python -c "import secrets;print(secrets.token_urlsafe(32))"`.

Avec Docker (même résultat, dossier de modèles monté en lecture seule) :
```bash
docker build -t mcp-office .
docker run --rm -p 8000:8000 -e MCP_API_KEY=... -e PUBLIC_BASE_URL=http://localhost:8000 \
  -e TEMPLATE_DIR=/templates -v "/chemin/vers/Modeles:/templates:ro" mcp-office
```

### 4. Vérification
- `http://localhost:8000/health` doit répondre `{"ok":true}`.
- Tester les outils sans client IA avec MCP Inspector : `npx @modelcontextprotocol/inspector`, transport **Streamable HTTP**, URL `http://localhost:8000/mcp`, en-tête `Authorization: Bearer <MCP_API_KEY>`. Appeler `list_templates`, puis `list_slide_types`.
- Les liens de téléchargement renvoyés pointent vers `PUBLIC_BASE_URL` et expirent après `FILE_TTL` secondes (1 h par défaut).

### 5. Brancher un client MCP
- **LibreChat lancé sur le même poste** (Docker) : bloc `mcpServers` ci-dessus avec `url: http://host.docker.internal:8000/mcp`.
- **LibreChat hébergé (instance de l'entreprise)** : l'instance ne peut pas joindre `localhost`. Le serveur doit être déployé sur une machine accessible par LibreChat (Docker derrière le reverse-proxy HTTPS interne, cf. « Lancement »), avec `PUBLIC_BASE_URL` égal à son adresse publique interne.
- **Tout client MCP compatible HTTP** (Claude Desktop, VS Code…) : URL `http://localhost:8000/mcp` et en-tête `Authorization`.

Ne pas versionner les modèles d'entreprise (documents classifiés C2) dans ce dépôt : le `.gitignore` les exclut.
