# MCP Office – modèles d'entreprise pour LibreChat

Serveur MCP (HTTP streamable) qui transforme les réponses de l'IA en **Word, Excel, PowerPoint**, à partir des **modèles d'entreprise** du dossier SharePoint (styles, thème, masques, en-têtes/pieds de page conservés) : le plus récent pour Word et Excel, le modèle **épinglé** pour PowerPoint (voir catalogue).

## Outils exposés
| Outil | Entrée | Sortie |
|---|---|---|
| `create_word` | `title`, `markdown` (titres, listes, tableaux, gras/italique, code) | .docx |
| `create_excel` | `title`, `sheets:[{name, rows}]` (1re ligne = en-têtes → tableau Excel) | .xlsx |
| `get_presentation_catalog` | – | diapos PowerPoint autorisées : usage, champs, limites, règles de rédaction (à appeler avant `create_powerpoint`) |
| `create_powerpoint` | `title`, `slides:[{model, fields, variante, notes}]`, `ambiance` | .pptx + liste des dépassements à corriger |
| `refresh_template` | – | épingle le dernier modèle PowerPoint et refait l'analyse (sur demande explicite) |
| `list_slide_types` | `layouts` (optionnel) | administration : inventaire brut des zones du modèle épinglé |
| `list_templates` | – | modèles disponibles / utilisé |

`template_prefix` (optionnel) permet de choisir une famille de modèles (ex. `Note_`, `Rapport_`). Le fichier retourné est un lien de téléchargement à usage temporaire (`FILE_TTL`).

## Fonctionnement
- **Dernier modèle** : liste le dossier via Microsoft Graph, prend le fichier `.dotx/.docx`, `.xltx/.xlsx`, `.potx/.pptx` le plus récemment modifié ; cache invalidé par eTag.
- **Style par défaut** : Word → styles `Title`, `Heading n`, `List Bullet/Number`, `Table Grid` du modèle (repli si absents) ; le contenu d'exemple du modèle est supprimé, la mise en page conservée. Excel → 1re feuille du modèle dupliquée, données sous l'en-tête existant. PowerPoint → voir ci-dessous.

### PowerPoint : catalogue de diapos
Le chat ne manipule pas le modèle directement : il choisit parmi des **diapos autorisées**, présentées par rubrique (Ouverture, Structure, Messages, Contenu, Zoom, Démarche, Chiffres, Niji, Annexes, Clôture), et remplit des **champs nommés** (`titre`, `col1_points`…). Le serveur s'occupe de la mise en forme.

**Variantes.** L'analyse regroupe automatiquement les dispositions de même structure (mêmes zones aux mêmes positions) en **familles** : 13 intercalaires, 5 sous-intercalaires, 4 couvertures, 3 conclusions, objectifs en 3 couleurs… Une diapo du catalogue peut pointer vers une famille (`"source":{"family":"Intercalaire-Niveau1"}`) ou une liste de diapos types (`{"type":[20,22,23]}`) : 
- la variante est choisie dans l'**ambiance couleur** du document (`ambiance` = `orange`, `magenta` ou `turquoise`, charte : une couleur primaire par document) et **alternée** d'une occurrence à l'autre (intercalaires tous différents) ;
- le chat peut imposer `"variante": n` (liste et libellés dans le catalogue) ;
- une variante ajoutée à une famille dans un nouveau modèle est prise en compte au `refresh_template`, sans modifier la configuration (libellé/ambiance à compléter dans `variants`).

Autres sources : `"asis": true` insère une diapo officielle telle quelle (pages Niji : à propos, clients, récompenses, RSE, certifications…), `"sequence": true` insère toutes les diapos de la source (CGV), `"variant_by": "champ"` choisit la variante d'après un champ.

`refresh_template` signale toute disposition du modèle **ni exposée ni écartée** : la vision du modèle reste complète. Les dispositions écartées le sont avec leur raison dans `ignored_layouts` (images ou logos à fournir, doublons, couleur hors charte).

Le catalogue a deux parties :

| | Où | Contenu | Mise à jour |
|---|---|---|---|
| **Configuration** | `catalog/pptx.json` (versionné, dans l'image) | diapos exposées : usage, source (diapo type n° ou disposition du thème), champs → zones, limites (`max` caractères, `max_lines`, `max_rows`/`max_cols`), textes fixes, règles de rédaction | à la main, à chaque nouveau gabarit |
| **État** | `CATALOG_DIR` (volume `catalog-state`) | copie du modèle **épinglé**, polices du thème, inventaire des zones | outil `refresh_template` uniquement |

Le modèle n'est donc **ni relu sur SharePoint ni ré-analysé** à chaque génération (≈ 1 s par présentation). Publier un nouveau modèle ne change rien tant que `refresh_template` n'a pas été appelé ; cet outil signale les champs de la configuration qui ne correspondent plus au modèle.

Rendu :
- chaque diapo est la **copie d'une diapo type** du modèle (décor, images, logos conservés) ou une diapo vierge d'une **disposition du thème** (`"source":{"layout":"Pyramid - Grayscale"}`) ; les diapos types d'origine sont retirées ;
- un champ non fourni retire sa zone (pas de texte d'exemple résiduel) ; `fixed` impose un texte ou supprime une forme (`null`, ex. étiquette « EXEMPLE ») ;
- **polices** : celles du thème (N27 Light pour le texte), `**mot**` → N27 Medium (règle de la charte) ; une police étrangère au thème héritée d'un texte d'exemple est retirée ;
- chaque ligne reprend le style du paragraphe correspondant de l'exemple. Si des lignes commencent par `- `, la structure est respectée : `- ` → paragraphe à puce, autre → paragraphe sans puce, `""` → ligne vide ;
- options de champ : `upper` (majuscules saisies dans le modèle, `1` = 1re ligne seulement), `breaks` (lignes → sauts de ligne d'un même paragraphe), `anchor` (`"t"` = texte en haut), `bullets: false` (pas de puce héritée), `prefix` (lignes fixes ajoutées en tête), `default` ;
- zones désignées par nom de forme (volet Sélection), `Nom#n` (n-ième forme de ce nom) ou `@idx` (espace réservé d'une disposition) : voir `list_slide_types`.
- les **dépassements** de limites sont renvoyés au chat avec le lien, pour correction.

Ajouter une diapo au catalogue : `list_slide_types` (ou `layouts=True`) → repérer les zones → ajouter un modèle dans `catalog/pptx.json` → `refresh_template` pour vérifier → contrôler le rendu.

### Accessibilité
Titre et langue du document renseignés, vrais styles de titres, en-tête de tableau répété (Word), tableaux Excel structurés, titre sur chaque diapo, notes orateur.

## Sécurité
- App Entra ID en **`Sites.Selected`** (lecture seule, un seul site) – pas de `Sites.Read.All`.
- `/mcp` protégé par `Authorization: Bearer $MCP_API_KEY` (comparaison à temps constant). **Si `MCP_API_KEY` est vide, `/mcp` n'est pas protégé** (avertissement au démarrage).
- Liens de fichiers non devinables (192 bits), expirés après `FILE_TTL`, `no-store`, `nosniff`, sans traversée de chemin.
- Taille d'entrée limitée (`MAX_INPUT`), conteneur non-root. À exposer uniquement derrière le reverse-proxy HTTPS interne.

Accorder l'accès au site (admin, une fois) :
```http
POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
{"roles":["read"],"grantedToIdentities":[{"application":{"id":"<client-id>","displayName":"mcp-office"}}]}
```
`SP_SITE_ID` : `GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/<NomDuSite>` → `id`.

## Variables d'environnement
| Variable | Défaut | Rôle |
|---|---|---|
| `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` | – | App Entra ID (Graph) |
| `SP_SITE_ID` | – | Site SharePoint des modèles |
| `SP_FOLDER` | `Modeles` | Dossier des modèles dans la bibliothèque par défaut |
| `TEMPLATE_DIR` | – | Dossier local de modèles (remplace SharePoint si défini) |
| `MCP_API_KEY` | – | Jeton Bearer exigé sur `/mcp` |
| `PUBLIC_BASE_URL` | `http://localhost:8000` | Base des liens de téléchargement |
| `FILE_TTL` | `3600` | Durée de validité des fichiers générés (s) |
| `TEMPLATE_CACHE` | `300` | Durée du cache des modèles (s) |
| `MAX_INPUT` | `500000` | Taille max. des entrées (caractères) |
| `DOC_LANG` | `fr-FR` | Langue des documents |
| `XLSX_TABLE_STYLE` | `TableStyleMedium2` | Style des tableaux Excel |
| `OUTPUT_DIR` | `/out` (Docker) | Dossier des fichiers générés |
| `CATALOG_DIR` | `/data/catalog` | État du catalogue PowerPoint (modèle épinglé, analyse) |
| `MODELS_DIR` | `./catalog` | Configuration des diapos (`pptx.json`, ou `pptx-<prefixe>.json` par famille de modèles) |
| `HOST`, `PORT` | `0.0.0.0`, `8000` | Écoute du serveur (garder 8000 en Docker) |

## Lancement (Docker)
```bash
cp .env.example .env   # compléter
docker compose up -d --build
docker compose logs -f            # suivi
curl http://127.0.0.1:8000/health # {"ok":true}
```
`docker-compose.yml` : redémarrage auto, volume pour les fichiers générés, système de fichiers en lecture seule, sans capacités, port lié à `127.0.0.1` (variable `BIND_ADDR` pour l'ouvrir, `HOST_PORT` pour le changer).
- **Même hôte que LibreChat** : décommenter `networks` (nom du réseau via `docker network ls`) ; LibreChat joint alors `http://mcp-office:8000/mcp`.
- **Sans SharePoint** : `TEMPLATE_DIR=/templates` dans `.env` et décommenter le montage `./modeles:/templates:ro`.
- `PUBLIC_BASE_URL` doit être l'URL vue par les utilisateurs (liens de téléchargement).

Sans compose : `docker build -t mcp-office . && docker run -d --env-file .env -p 8000:8000 mcp-office`

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
      et renvoie le lien de téléchargement tel quel.
      PowerPoint : appelle d'abord get_presentation_catalog, propose un plan diapo par diapo
      (modèle + contenu) structuré en parties (intercalaires) et rythmé (varier les mises en page),
      choisis une ambiance couleur, en respectant les limites, puis appelle create_powerpoint.
      Si des dépassements sont signalés, raccourcis les textes concernés et regénère.
      N'appelle refresh_template que si l'utilisateur annonce un nouveau modèle d'entreprise.
```
Si l'instance LibreChat est hébergée par un tiers, demander à l'administrateur d'ajouter ce bloc (ou d'autoriser les serveurs MCP utilisateurs) et d'ouvrir le flux réseau vers le serveur.

## Test local sans SharePoint
```bash
pip install -r requirements.txt
TEMPLATE_DIR=./modeles CATALOG_DIR=./.catalog MCP_API_KEY=test python server.py
```
Le rendu PowerPoint suppose les polices N27 installées sur le poste qui ouvre le fichier (elles ne sont pas embarquées).

## TODO

### Avant la mise en service
- [x] **Tester l'image Docker** (01/10/2026, Docker 29.3 sous WSL 2, modèles en local) : image de 207 Mo construite, conteneur `healthy`, utilisateur non-root `app`, système de fichiers en lecture seule sauf `/out` et `/data/catalog`, `/mcp` refusé sans clé (401), 7 outils, `refresh_template`, catalogue, présentation de 28 diapos générée en 0,7 s et téléchargée, catalogue conservé après redémarrage, 126 Mo de mémoire utilisés sur 512 Mo.
- [ ] Fuseau horaire : le conteneur est en UTC (la date « analysé le » a 2 h de décalage). Ajouter `tzdata` à l'image et `TZ=Europe/Paris`.
- [ ] **Tester avec SharePoint** : app Entra ID en `Sites.Selected`, `SP_SITE_ID` / `SP_FOLDER` réels, récupération du modèle épinglé (testé uniquement avec `TEMPLATE_DIR`).
- [ ] **Tester de bout en bout depuis LibreChat** : appel de `get_presentation_catalog` par le chat, qualité des plans proposés, prise en compte des dépassements signalés, lien de téléchargement via `PUBLIC_BASE_URL` derrière le reverse-proxy.
- [ ] **Restreindre `refresh_template`** aux administrateurs (clé distincte ou outil non exposé au chat) : aujourd'hui tout utilisateur du chat peut changer le modèle épinglé.
- [ ] Healthcheck du `Dockerfile` : utiliser `PORT` au lieu de 8000 en dur.

### PowerPoint
- [ ] **Images** : remplir les espaces réservés image (photo, logo, motif) à partir d'une bibliothèque SharePoint ou d'une URL autorisée, avec texte alternatif. Débloquerait les dispositions écartées : équipe (`Team`), fiches références, `Full Image`, `Images + Infos`, panneaux `Detailed Section`, intercalaires et couvertures « Editable ».
- [ ] **Diapos types complexes non exposées** : plannings Gantt (diapos 39-40), proposition financière et estimation budgétaire (46-47), méthodologies (19, 33), équipe (41-42), suivi continu (44).
- [ ] **Objective 1/2/3** : les zones diffèrent selon la variante. Il faudrait un mapping de zones par variante pour exposer cette frise (objectif n mis en avant).
- [ ] **Sommaire automatique** : le déduire des intercalaires et sous-intercalaires de la présentation, au lieu de le faire saisir au chat.
- [ ] **Contrôle des débordements** : les limites `max` sont des nombres de caractères estimés à l'œil. Il serait plus fiable de mesurer le texte avec les métriques des polices N27 et la taille des zones, et d'ajuster `max` d'après les rendus réels.
- [ ] Libellés et ambiances des variantes (`variants` dans `catalog/pptx.json`) : saisis à la main, à revoir à chaque nouveau modèle (`refresh_template` signale les nouvelles dispositions mais pas leur couleur).
- [ ] Polices N27 : vérifier la licence et l'éventuel embarquement dans les fichiers destinés aux clients.

### Word et Excel
- [ ] Appliquer la même approche que PowerPoint : modèle épinglé + catalogue (styles autorisés, blocs types) au lieu du « dernier modèle » relu à chaque génération.
- [ ] Vérifier le rendu avec les modèles Niji réels (`C2-Niji-Word_Modele de doc-2026`) : styles de titres, listes, tableaux, page de garde.

### Qualité et exploitation
- [ ] **Tests automatisés** : génération de chaque type de diapo du catalogue, validité du XML (ouverture sans réparation par PowerPoint), contrôle des polices, non-régression visuelle (export PNG comparé à une référence).
- [ ] Le code s'appuie sur des attributs privés de python-pptx (`_rels`, `_spTree`, `_txBody`) : épingler des versions précises dans `requirements.txt` et tester avant chaque montée de version.
- [ ] Plusieurs instances : le cache du catalogue est par processus et les fichiers générés sont sur un volume local. Pour plusieurs instances, prévoir un stockage partagé et un rechargement du catalogue après `refresh_template`.
- [ ] Journaliser les générations (modèle, nombre de diapos, dépassements) sans le contenu, pour suivre l'usage et la qualité.
