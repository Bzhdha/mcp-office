# MCP Office – modèles d'entreprise pour LibreChat

Serveur MCP (HTTP streamable) qui transforme les réponses de l'IA en **Word, Excel, PowerPoint**, à partir des **modèles d'entreprise** (styles, thème, masques, en-têtes/pieds de page conservés) : modèles PowerPoint et Word **épinglés** et embarqués dans l'image (voir catalogues), modèle Excel le plus récent du dossier SharePoint.

## Outils exposés
| Outil | Entrée | Sortie |
|---|---|---|
| `get_word_catalog` | – | blocs Word disponibles, cadres, ambiances, plans types, bibliothèque d'UO (à appeler avant `create_word_document`) |
| `create_word_document` | `titre`, `blocs:[{type,…}]`, `sous_titre`, `cadre`, `ambiance`, `historique`, `interlocuteurs`, `options` | .docx au format du modèle ou du cadre client + nombre de pages |
| `create_word` | `title`, `markdown` | .docx simple (Markdown converti en blocs, même moteur) |
| `search_uo` / `get_uo` | `recherche` / `id` | bibliothèque d'unités d'œuvre issues des réponses précédentes |
| `save_word_cadre` | `nom`, `parametres` | enregistre un cadre client réutilisable (police, taille, marges, pages…) |
| `create_excel` | `title`, `sheets:[{name, rows}]` (1re ligne = en-têtes → tableau Excel) | .xlsx |
| `get_presentation_catalog` | – | diapos PowerPoint autorisées : usage, champs, limites, règles de rédaction (à appeler avant `create_powerpoint`) |
| `create_powerpoint` | `title`, `slides:[{model, fields, variante, notes}]`, `ambiance` | .pptx + liste des dépassements à corriger |
| `refresh_template` | `type` (`pptx` ou `docx`) | épingle le dernier modèle PowerPoint ou Word de la source et refait l'analyse (sur demande explicite) |
| `add_slides_link` | – | lien de dépôt (24 h) pour ajouter des diapos de l'utilisateur à son catalogue |
| `get_slides_report` | `import_id` | écart au modèle de chaque diapo déposée : score, niveau, alertes, champs proposés |
| `add_user_slides` | `import_id`, `slides:[{diapo, nom, usage, champs}]`, `forcer` | ajoute des diapos au catalogue de l'utilisateur (« Mes diapos ») |
| `remove_user_slide` | `nom` | retire une diapo de « Mes diapos » |
| `list_slide_types` | `layouts` (optionnel) | administration : inventaire brut des zones du modèle épinglé |
| `list_templates` | – | modèles disponibles / utilisé |

`template_prefix` (optionnel) permet de choisir une famille de modèles (ex. `Note_`, `Rapport_`). Le fichier retourné est un lien de téléchargement à usage temporaire (`FILE_TTL`).

## Fonctionnement
- **Source des modèles** : dossier SharePoint via Microsoft Graph (ou `TEMPLATE_DIR`), fichier le plus récemment modifié par type ; cache invalidé par eTag. Excel l'utilise à chaque génération ; PowerPoint et Word seulement à l'épinglage (`refresh_template`, `python server.py bundle`).
- **Style par défaut** : Excel → 1re feuille du modèle dupliquée, données sous l'en-tête existant. Word et PowerPoint → voir ci-dessous.

### Word : blocs, cadres et unités d'œuvre
Le modèle `C2-Niji-Word*` est **épinglé** comme le modèle PowerPoint (embarqué dans `catalog/bundle`, `refresh_template(type="docx")` pour en changer). Moteur : `word.py` ; configuration : `catalog/docx.json`.

- **Début de document repris du modèle** : page de garde (titre, sous-titre), historique des versions, vos interlocuteurs (sans photo d'exemple), sommaire (mis à jour à l'ouverture dans Word). Le corps d'exemple est remplacé par les blocs ; la 4e de couverture est conservée ; présentation Niji et CGV en option.
- **Blocs** : `titre` (niveaux 1 à 4, numérotation du modèle), `paragraphe` (mise en avant, légende), `liste`, `liste_numerotee`, `tableau` (en-tête coloré répété), `encadre`, `citation`, `chiffres_cles`, `tableau_risques` (probabilité × gravité, criticité colorée), `fiche_uo`, `fiche_profil`, `saut_de_page`, `presentation_niji`. Listes à puces à la charte (puce colorée, sous-niveaux).
- **Ambiances** : couleur des encadrés, fiches et puces (violet, magenta, rose, orange, turquoise, bleu, sobre).
- **Cadres** : `niji` (N27 Light 10,5 pt, mots clés en N27 Medium, marges du modèle) ou cadre imposé par le client, passé par nom ou en paramètres : `police_texte`, `police_titres`, `taille`, `interligne`, `marges_cm`, `format`, `pages_max`, `couleurs` (`sobre` = gris), `couleur_titres`, `justifie`, `strict` (toute police du document suit le cadre, page de garde comprise), et les éléments `page_de_garde`, `historique`, `interlocuteurs`, `sommaire`, `presentation_niji`, `cgv`. Cadres d'exemple : `client_arial_11`, `client_compact_10`. `save_word_cadre` enregistre un cadre par utilisateur (ex. exigences d'un client récurrent).
- **Nombre de pages** : exact si LibreOffice (`soffice`) est présent dans l'image, sinon **estimation** (écart constaté de ± 2 pages sur 12 à 14 pages) ; un dépassement de `pages_max` est signalé avec le pourcentage à retirer.
- **Unités d'œuvre** : `fiche_uo` reproduit la présentation des réponses à appels d'offres (titre coloré, rubriques Objectif, Prérequis, Méthode, Livrables, Facteurs clés de succès, Principaux risques, Profils et hypothèses de charge, tableau de charge facultatif). La **bibliothèque d'UO** est extraite des réponses passées : `python server.py bundle-uo <dossier de .docx>` → `catalog/bundle/uo_library.json` (hors Git, contenu client confidentiel) ; le chat la consulte avec `search_uo` / `get_uo` et adapte la fiche.
- **Plans types** : mémoire technique, réponse organisation / RH, plan d'assurance sécurité (dans `catalog/docx.json`).

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

#### Modèle embarqué (prêt au déploiement)
L'image embarque dans `catalog/bundle/` les modèles PowerPoint et Word, leur analyse et la bibliothèque d'UO : au premier démarrage (volume `catalog-state` vide), les modèles sont recopiés dans `CATALOG_DIR` et le serveur produit présentations et documents Word **sans SharePoint ni `TEMPLATE_DIR`**. Le journal de démarrage indique les modèles utilisés et le nombre d'UO. `refresh_template` reste possible ensuite si une source est configurée.

Préparer le paquet avant `docker compose build` (depuis un dossier contenant les `.potx` et `.docx` à jour, ou depuis SharePoint avec les variables `SP_*`) :
```bash
TEMPLATE_DIR=/chemin/vers/modeles python server.py bundle
# → catalog/bundle/<modèle>.potx + pptx.json (analyse) et <modèle Word>.docx + docx.json
python server.py bundle-uo /chemin/vers/reponses-ao
# → catalog/bundle/uo_library.json (fiches UO extraites des mémoires techniques)
```
> ⚠ `catalog/bundle/` est **exclu de Git** : les modèles sont des documents internes (C2 - restricted), la bibliothèque d'UO reprend des extraits de réponses à appels d'offres, et ce dépôt est public. Il est copié dans l'image au build : ne pas publier l'image sur un registre public (utiliser un registre privé).

Pour livrer un nouveau modèle : refaire `python server.py bundle`, reconstruire l'image, puis **vider le volume** `catalog-state` (`docker compose down -v`) ou appeler `refresh_template`, sinon le modèle déjà épinglé dans le volume est conservé.

Rendu :
- chaque diapo est la **copie d'une diapo type** du modèle (décor, images, logos conservés) ou une diapo vierge d'une **disposition du thème** (`"source":{"layout":"Pyramid - Grayscale"}`) ; les diapos types d'origine sont retirées ;
- un champ non fourni retire sa zone (pas de texte d'exemple résiduel) ; `fixed` impose un texte ou supprime une forme (`null`, ex. étiquette « EXEMPLE ») ;
- **polices** : celles du thème (N27 Light pour le texte), `**mot**` → N27 Medium (règle de la charte) ; une police étrangère au thème héritée d'un texte d'exemple est retirée ;
- chaque ligne reprend le style du paragraphe correspondant de l'exemple. Si des lignes commencent par `- `, la structure est respectée : `- ` → paragraphe à puce, autre → paragraphe sans puce, `""` → ligne vide ;
- options de champ : `upper` (majuscules saisies dans le modèle, `1` = 1re ligne seulement), `breaks` (lignes → sauts de ligne d'un même paragraphe), `anchor` (`"t"` = texte en haut), `bullets: false` (pas de puce héritée), `prefix` (lignes fixes ajoutées en tête), `default` ;
- zones désignées par nom de forme (volet Sélection), `Nom#n` (n-ième forme de ce nom) ou `@idx` (espace réservé d'une disposition) : voir `list_slide_types`.
- les **dépassements** de limites sont renvoyés au chat avec le lien, pour correction.

**Accessibilité** (contrôlée à chaque génération, lecteurs d'écran et vérificateur d'accessibilité de PowerPoint) :
- **titre** : chaque diapo a un titre ; s'il n'en a pas (fiches, conclusions, pages Niji, CGV), un titre est ajouté hors de la zone visible (`titre_accessible` du modèle, sinon premier champ significatif, sinon l'usage) ;
- **ordre de lecture** (= ordre d'empilement des formes) : titre, puis champs dans l'ordre du catalogue ; chaque intitulé ou numéro fixe est lu juste avant le contenu qu'il introduit ; sans configuration (diapos utilisateur, pages Niji), découpage XY : colonnes puis bandes. Les formes décoratives gardent leur rang d'origine, et deux formes lues qui se chevauchent gardent l'ordre du modèle : le rendu visuel est inchangé ;
- **tableaux** : première ligne déclarée comme ligne d'en-têtes, résumé des en-têtes en texte de remplacement ; une cellule d'en-tête vide est signalée au chat ;
- **textes de remplacement** : images et formes sans texte de remplacement marquées **décoratives** ; texte de remplacement fourni par `"alt": {"Nom de forme": "texte"}` dans un modèle du catalogue ou à l'ajout d'une diapo utilisateur (`add_user_slides`, avec aussi `"ordre"` pour imposer l'ordre de lecture des champs).

Ajouter une diapo au catalogue : `list_slide_types` (ou `layouts=True`) → repérer les zones → ajouter un modèle dans `catalog/pptx.json` → `refresh_template` pour vérifier → contrôler le rendu.

### Diapos de l'utilisateur (« Mes diapos »)
Un utilisateur peut étendre son catalogue avec des diapos de ses propres présentations, pour générer ensuite des diapos du même type avec un autre contenu.

1. Le chat appelle `add_slides_link` et donne à l'utilisateur un **lien de dépôt** (`/import/<jeton>`, valable 24 h, 50 Mo max) ; l'utilisateur y dépose son `.pptx` depuis le navigateur.
2. Chaque diapo est **comparée au modèle épinglé** et notée de 0 à 100 ; la page de dépôt et `get_slides_report` affichent le résultat :

   | Écart contrôlé | Pénalité |
   |---|---|
   | Disposition absente du modèle d'entreprise | refus |
   | Contenu non reproductible : graphique, SmartArt, objet incorporé, vidéo | refus |
   | Polices du thème différentes | 25 |
   | Couleurs du thème différentes | 3 par couleur (max. 20) |
   | Police hors charte dans la diapo (ex. Arial) | 15 par police (max. 30) |
   | Couleur hors palette du modèle (les gris sont neutres) | 10 par couleur (max. 30) |
   | Format de diapo différent | 20 |
   | Disposition modifiée par rapport au modèle | 10 |
   | Formes hors de la diapo | 5 par forme (max. 15) |
   | Texte de moins de 8 pt | 5 |

   **Niveaux** : *conforme* (≥ 80 et aucune entorse à la charte), *alerte* (< 80, ou police, couleur ou thème hors charte), *refus* (< 60, ou cas bloquant). Seuils modifiables dans `catalog/pptx.json` (`"conformite": {"alerte": 80, "refus": 60}`).
3. `add_user_slides` ajoute les diapos choisies, avec un nom et un usage. Les zones de texte et tableaux deviennent des **champs** (ordre de lecture, exemple et limite de longueur tirés de la diapo d'origine) ; le chat les renomme d'après leurs exemples (`"champs": {"texte2": "col1_titre"}`). Les numéros décoratifs, formes, images et le reste du décor sont conservés tels quels. Une diapo en **alerte** est ajoutée avec son alerte, rappelée dans le catalogue pour que le chat la signale ; une diapo en **refus** n'est ajoutée qu'avec `forcer=true`, après accord explicite de l'utilisateur.
4. Ces diapos apparaissent dans `get_presentation_catalog` (rubrique « Mes diapos ») et s'utilisent dans `create_powerpoint` comme les autres. À la génération, la diapo est recopiée dans la présentation (images et liens compris) sur la disposition du modèle de même nom, et les textes saisis reprennent les polices du thème.

**Utilisateur** : identifié par l'en-tête `X-User-Id` transmis par LibreChat (voir configuration), stocké sous forme de hachage ; sans en-tête, le catalogue est commun à tous. Les présentations déposées et les catalogues personnels sont dans le volume `catalog-state` (`users/`, `imports/`).

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
| `MAX_UPLOAD` | `52428800` | Taille max. d'une présentation déposée (octets) |
| `DOC_LANG` | `fr-FR` | Langue des documents |
| `XLSX_TABLE_STYLE` | `TableStyleMedium2` | Style des tableaux Excel |
| `OUTPUT_DIR` | `/out` (Docker) | Dossier des fichiers générés |
| `CATALOG_DIR` | `/data/catalog` | État : modèles PowerPoint et Word épinglés, analyse, diapos et cadres des utilisateurs, dépôts |
| `MODELS_DIR` | `./catalog` | Configuration : `pptx.json` (diapos ; `pptx-<prefixe>.json` par famille de modèles), `docx.json` (blocs, cadres, plans types), `docx-styles.xml`, `bundle/` (modèles et bibliothèque d'UO embarqués) |
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
      X-User-Id: "{{LIBRECHAT_USER_ID}}"   # catalogue « Mes diapos » propre à chaque utilisateur
    timeout: 60000
    serverInstructions: |
      Quand l'utilisateur demande un Word, Excel ou PowerPoint, appelle l'outil correspondant
      et renvoie le lien de téléchargement tel quel.
      PowerPoint : appelle d'abord get_presentation_catalog, propose un plan diapo par diapo
      (modèle + contenu) structuré en parties (intercalaires) et rythmé (varier les mises en page),
      choisis une ambiance couleur, en respectant les limites, puis appelle create_powerpoint.
      Si des dépassements sont signalés, raccourcis les textes concernés et regénère.
      N'appelle refresh_template que si l'utilisateur annonce un nouveau modèle d'entreprise.
      Word : appelle d'abord get_word_catalog ; si le client impose un cadre (police, taille,
      marges, nombre de pages), passe-le dans « cadre » ; pour les unités d'œuvre, cherche une
      UO existante (search_uo, get_uo) et adapte-la au nouveau client avant create_word_document.
      Si l'utilisateur veut réutiliser des diapos de sa propre présentation : add_slides_link,
      puis get_slides_report ; présente-lui les alertes d'écart au modèle avant add_user_slides,
      et ne force jamais une diapo refusée sans son accord explicite.
```
Si l'instance LibreChat est hébergée par un tiers, demander à l'administrateur d'ajouter ce bloc (ou d'autoriser les serveurs MCP utilisateurs) et d'ouvrir le flux réseau vers le serveur.

## Test local sans SharePoint
```bash
pip install -r requirements.txt
TEMPLATE_DIR=./modeles CATALOG_DIR=./.catalog MCP_API_KEY=test python server.py
```
Le rendu PowerPoint et Word (cadre Niji) suppose les polices N27 installées sur le poste qui ouvre le fichier (elles ne sont pas embarquées).

## TODO

### Avant la mise en service
- [x] **Tester l'image Docker** (01/10/2026, Docker 29.3 sous WSL 2, modèles en local) : image de 207 Mo construite, conteneur `healthy`, utilisateur non-root `app`, système de fichiers en lecture seule sauf `/out` et `/data/catalog`, `/mcp` refusé sans clé (401), 7 outils, `refresh_template`, catalogue, présentation de 28 diapos générée en 0,7 s et téléchargée, catalogue conservé après redémarrage, 126 Mo de mémoire utilisés sur 512 Mo.
- [ ] Fuseau horaire : le conteneur est en UTC (la date « analysé le » a 2 h de décalage). Ajouter `tzdata` à l'image et `TZ=Europe/Paris`.
- [x] **Modèle PowerPoint embarqué dans l'image** (`catalog/bundle`, hors Git) : testé en déploiement à nu (volume vierge, sans SharePoint ni `TEMPLATE_DIR`), présentation de 28 diapos produite.
- [x] **Modèle Word et bibliothèque d'UO embarqués** : testés dans Docker sous WSL (16 outils exposés, catalogue Word servi, document généré et téléchargé).
- [ ] Embarquer aussi le modèle Excel (aujourd'hui relu sur SharePoint ou `TEMPLATE_DIR` à chaque génération).
- [ ] **Tester avec SharePoint** : app Entra ID en `Sites.Selected`, `SP_SITE_ID` / `SP_FOLDER` réels, récupération du modèle épinglé (testé uniquement avec `TEMPLATE_DIR`).
- [ ] **Tester de bout en bout depuis LibreChat** : appel de `get_presentation_catalog` et `get_word_catalog` par le chat, qualité des plans proposés, prise en compte des dépassements signalés (longueurs, pages), lien de téléchargement via `PUBLIC_BASE_URL` derrière le reverse-proxy.
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
- [x] **Accessibilité** : titre sur chaque diapo, ordre de lecture logique, en-têtes de tableaux déclarés, images et formes décoratives marquées. Mesuré sur la présentation de 28 diapos : 6 diapos sans titre, 11 images et 23 formes sans texte de remplacement, 2 tableaux sans en-tête déclaré → 0 ; rendu identique au pixel près (hors en-têtes nommés).
- [ ] Accessibilité, limites : quand un en-tête est posé sur un cadre plein (comitologie), le cadre est lu avant son en-tête pour ne pas le masquer ; les pages Niji figées ont leur texte dans la disposition (non lu par les lecteurs d'écran : seul le titre ajouté est annoncé) ; l'ordre des CGV suit la géométrie et peut placer un paragraphe avant son intertitre. Valider avec le vérificateur d'accessibilité de PowerPoint et un lecteur d'écran (NVDA).
- [ ] Accessibilité Word : en-têtes de tableaux répétés en place ; vérifier les textes de remplacement des images du modèle (page de garde, présentation Niji) et l'ordre de lecture des zones de texte.

### Diapos de l'utilisateur
- [x] Dépôt, analyse d'écart au modèle, ajout au catalogue personnel et génération : testé de bout en bout en HTTP (diapos conformes à 100/100, diapo retouchée en Arial avec couleur hors palette en alerte à 75/100, graphique et présentation hors modèle refusés, isolement entre deux utilisateurs, rendu vérifié).
- [x] Dans Docker : page de dépôt, analyse et écriture de `imports/` dans le volume `catalog-state` testées (déploiement de l'archive).
- [ ] Dans Docker : ajout au catalogue personnel (`users/`) et génération avec une diapo utilisateur.
- [ ] Tester l'en-tête `X-User-Id` avec LibreChat réel (`{{LIBRECHAT_USER_ID}}`) et le lien de dépôt derrière le reverse-proxy (taille maximale des requêtes).
- [ ] **Décisions à valider** :
  - Arrivée du fichier : lien de dépôt servi par le serveur (choix actuel, aucun droit SharePoint supplémentaire, reverse-proxy à configurer pour 50 Mo) ou lecture dans le OneDrive/SharePoint de l'utilisateur (droits Graph plus larges).
  - Identification : en-tête `X-User-Id` transmis par LibreChat, conservé sous forme d'empreinte ; sans en-tête, catalogue « Mes diapos » commun à tous.
  - Portée : catalogues personnels (choix actuel) ou catalogue d'équipe validé par un référent charte.
  - Seuils de conformité : alerte < 80, refus < 60, toute entorse à la charte (police, couleur, thème) en alerte ; pondérations du tableau ci-dessus.
- [ ] Catalogue d'équipe : promouvoir une diapo utilisateur validée vers un catalogue partagé (aujourd'hui personnel ou commun à tous sans en-tête), avec validation par un référent charte.
- [ ] Graphiques et SmartArt : aujourd'hui refusés car non recopiés ; prise en charge possible (copie des parties `chart` et de leur classeur).
- [ ] Textes dans des formes groupées : conservés tels quels, pas proposés comme champs.
- [ ] Nommage des champs : automatique (`texte1`…) puis renommé par le chat ; le rendre plus parlant d'après la position et le style.
- [ ] Durée de conservation des présentations déposées et des catalogues personnels (RGPD, confidentialité des contenus) ; outil de purge.
- [ ] Après un `refresh_template`, revérifier les diapos utilisateur (dispositions disparues du nouveau modèle).

### Word
- [x] Modèle Word épinglé et embarqué, catalogue de blocs, cadres client, bibliothèque d'UO (16 fiches extraites de 2 mémoires techniques) : testé en HTTP et rendu vérifié dans Word (charte Niji avec présentation et CGV, cadre Arial 11 sobre, cadre Times 12 avec dépassement de pages signalé, cadre client enregistré puis réutilisé).
- [ ] Nombre de pages exact : ajouter LibreOffice (`libreoffice-writer`) à l'image (environ 300 Mo) ou garder l'estimation ; recalibrer l'estimation sur davantage de documents.
- [ ] Images : logo client en page de garde, photos des interlocuteurs, illustrations dans le corps.
- [ ] Champs de la page de garde propres au client (nom du client, référence de la consultation, lot) et en-tête aux couleurs du client.
- [ ] Bibliothèque d'UO : l'alimenter avec d'autres réponses, la rendre consultable par équipe, purger les mentions propres à l'ancien client à l'extraction.
- [ ] Autres fiches récurrentes à extraire des réponses : fiches profils (CV), tableaux de références, matrices de compétences.
- [ ] Contrôle de conformité d'un document Word déposé par l'utilisateur (équivalent de « Mes diapos »).

### Excel
- [ ] Appliquer la même approche (modèle épinglé, catalogue) : aujourd'hui dernier modèle relu à chaque génération.

### Qualité et exploitation
- [ ] **Tests automatisés** : génération de chaque type de diapo du catalogue, validité du XML (ouverture sans réparation par PowerPoint), contrôle des polices, non-régression visuelle (export PNG comparé à une référence).
- [ ] Le code s'appuie sur des attributs privés de python-pptx (`_rels`, `_spTree`, `_txBody`) : épingler des versions précises dans `requirements.txt` et tester avant chaque montée de version.
- [ ] Plusieurs instances : le cache du catalogue est par processus et les fichiers générés sont sur un volume local. Pour plusieurs instances, prévoir un stockage partagé et un rechargement du catalogue après `refresh_template`.
- [ ] Journaliser les générations (modèle, nombre de diapos, dépassements) sans le contenu, pour suivre l'usage et la qualité.
