# Journal du projet — Reprise Magnitude → Tagetik

> Mémoire de travail. Mise à jour au fil des explications de l'utilisateur.
> Chaque section = une étape expliquée. Ne rien supposer avant explication.

## Objectif final

Un **moteur Excel** où :
- on **colle l'extraction Magnitude** dans le **premier onglet**,
- on obtient un **output Tagetik** dans le **dernier onglet**,
- les onglets intermédiaires portent les **tables de mapping**.
Construction pilotée par l'utilisateur, **step by step**. Sans VBA (Power Query privilégié).

## Fichiers fournis

- `travail/RDHEngineQ1.xlsx` (~9,9 Mo) — fourni le 2026-08-20
- `travail/RDHEngineQ1Copie.xlsx` (~227 Ko) — fourni le 2026-08-20
- Rôle exact de chacun : *à préciser par l'utilisateur*.

## Consignes de l'utilisateur

- Ne PAS tout analyser d'un coup (éviter la pollution). Avancer à son rythme.
- Il explique chaque partie ; j'attends son guidage avant d'agir.
- Tenir ce journal à jour et bien structurer.

## Étapes expliquées

_(vide pour l'instant — à compléter au fur et à mesure)_

## Décisions / conventions actées

- Forme : classeur Excel, sans VBA, Power Query.
- Emplacement : dossier `magnitude-tagetik-reprise/` dans le dépôt `demo-3`,
  branche `claude/eduservices-verification-aar4ni`.

## Questions en suspens

- Rôle respectif des deux fichiers RDHEngineQ1 vs Copie.
- Périmètre temporel de la reprise / scénario(s).

---

## Étape 1 — Onglet source & filtrage (onglet 2)

### Onglet source : « Liasse brute Q1 » (dans RDHEngineQ1.xlsx)
- Extraction Magnitude pour la période **2026.03** (mars 2026 / Q1).
- Dite « brute » mais ne l'est pas totalement : **les 3 dernières colonnes ont été
  ajoutées par l'utilisateur** (à préciser plus tard).
- Contient des données **mélangées** (avant et après allocation, etc.).

### Onglet 2 (à créer) : filtrage
Ne conserver **que les colonnes utiles** et appliquer les filtres suivants :

| Colonne source | Rôle / équivalent Tagetik | Règle de filtrage / transformation |
|----------------|---------------------------|------------------------------------|
| `D_DP`         | Période (année + mois)    | Sert à l'année et au mois          |
| `OA`           | Operational Activity ≈ **PMA** Tagetik | Garder **uniquement les lignes remplies** ; **exclure OA vide** |
| `RU`           | ≈ **Entité** Tagetik      | —                                  |
| `D_AC`         | Compte                    | Garder **uniquement les éléments fins** ; **exclure les nœuds** |
| `D_FL`         | Flux                      | Garder **uniquement `F99`** (donnée financière) et **`Q99`** (ETP) |
| `D_CU`         | Currency                  | —                                  |
| `P_AMOUNT`     | Montant                   | **Remplacer les `.` par `,`**      |

### Points de vigilance signalés (à traiter ensuite)
- **Entités (RU) et PMA (OA) : PAS du 1 pour 1** → mapping à définir (prochaine discussion).
- **Comptes (D_AC) : quelques sujets** à traiter aussi.

## Étape 1 bis — Onglet « rate » (taux de change)

- Il existe aussi un onglet **`rate`** (taux de change) dans le fichier source.
- Usage : servira à **contrôler la donnée en euro** dans un **deuxième temps**
  (conversion / réconciliation devise → EUR pour vérification).
- Pas utilisé tout de suite ; à garder de côté pour l'étape de contrôle.

---

## Étape 2 — Onglet « D_AC_Comptes » : mapping des comptes Magnitude → Tagetik

Rôle : table de correspondance **compte Magnitude → compte Tagetik**.
Parfois **1 pour 1**, souvent des **cas particuliers**. Détail ci-dessous.

### Cas 1 — GR6000 & GR6100 : un total + son détail complet
- `GR6000` et `GR6100` sont mappés vers :
  - `IND_00_070087` **et** `IND_26_060074` (le **détail**),
  - **et aussi** vers `IND_00_070008`.
- Dans Tagetik :
  - `IND_00_070008` = compte de **P&L** qui est la **somme** des deux.
  - `IND_00_070087` et `IND_26_060074` sont présentés comme des « **dont** » mais
    constituent en réalité le **détail complet** (leur somme = le total).
- ⇒ **CONFIRMÉ** : l'output Tagetik doit **alimenter les DEUX** pour ce cas précis :
  le **total** (`IND_00_070008`) **ET** le **détail** (`IND_00_070087` + `IND_26_060074`).
  Donc une même donnée source génère **plusieurs lignes** dans l'output.

### Cas 2 — GR2000 : nœud à NE PAS prendre (prendre le détail)
- `GR2000` = **Salaires variables (incl. charges sociales)** → **ne pas prendre**
  (c'est un agrégat), car il est **détaillé** en :
  - `GR2001`, `GR2002`, `GR2003`, `GR2004`, `GR2010`
- ⚠️ On **ne prend plus** `GR200211`.

### Cas 3 — GR2100 : salaires fixes, retraitement charges sociales
- `GR2100` = **Salaires fixes (incl. charges soc.)**. On connaît :
  - `GR2101` = **Indemnités de départ**
  - `GR2102` = **Engagements sociaux**
  - ⇒ `GR2101` et `GR2102` doivent être **déduits** de `GR2100`.
- Problème : dans **Tagetik**, on saisit les **salaires fixes HORS charges sociales**,
  et il existe un **compte séparé pour les charges sociales**.
- ⇒ **Prévoir un tableau de saisie PAR ENTITÉ** pour renseigner un
  **taux de charges sociales**, et **calculer automatiquement** la ventilation
  (salaires fixes hors charges soc. / charges sociales).

### Cas 4 — Mapping différent selon l'entité (France vs Pays)
- Le mapping d'un même compte **dépend de l'entité** :
  - **France** → comptes Tagetik trouvés « en face ».
  - **Pays (hors France)** → d'autres comptes, **juste à côté**, avec une
    **description explicite** indiquant que c'est pour les pays.
- Concerne notamment `GR3000` et `GR3200` (en plus des cas ci-dessus).

### Cas 5 — GFB110 & GFB130 : mapping inconnu (à laisser saisir)
- Mapping **pas encore connu** → les **prévoir dans la table de mapping**,
  les **contrôleurs de gestion saisiront** le mapping plus tard.

### Cas 6 — GR5000, GR5300, GR5400 : à préciser
- L'utilisateur **doit encore creuser** → mapping **TBD**.

### Implications pour la conception de la table de mapping
- La table doit gérer : correspondances **1→1**, **1→plusieurs** (total + détail),
  **exclusions** (nœuds/agrégats non repris), **variantes par entité** (France/Pays),
  et des **lignes à mapping vide** laissées à la saisie des contrôleurs.
- Retraitements calculés (charges sociales) pilotés par un **tableau de taux par entité**.

---

## Étape 3 — Indicateurs Tagetik proposés pour les comptes en attente (EN ROUGE = provisoire)

> Source : capture de la doc « Data Model » Tagetik, onglet **3.1 Indicators**.
> ⚠️ **À METTRE EN ROUGE** dans la table de mapping = **propositions provisoires**,
> à **valider par les contrôleurs de gestion** avant usage.

| Compte Magnitude | Nature       | Indicateur Tagetik (IND) | Libellé                                                    |
|------------------|--------------|--------------------------|------------------------------------------------------------|
| 🔴 `GFB130`      | Profit & Loss| `IND_26_060017`          | Refacturation frais de sièges (Headquarters…)              |
| 🔴 `GFB110`      | Profit & Loss| `IND_00_060004`          | Frais de gestion indirects - **Groupe**                    |
| 🔴 `GFB110`      | Profit & Loss| `IND_00_060323`          | Frais de gestion indirects - **Régaliens / Hors Métropole**|
| 🔴 `GR5000`      | Profit & Loss| `IND_00_060001`          | Cost of Risk - Depreciation                                |
| 🔴 `GR5300`      | Profit & Loss| `IND_00_070073`          | Autres résultats hors exploitation (IAS)                   |
| 🔴 `GR5400`      | Profit & Loss| `IND_00_070030`          | Résultat des sociétés mises en équivalence à FP Sociaux    |
| 🔴 `GR9201`      | Profit & Loss| `IND_00_060003`          | Of which External Management Costs                         |

### Notes
- `GFB110` apparaît sur **2 indicateurs** (`IND_00_060004` Groupe **et** `IND_00_060323`
  Régaliens/Hors Métropole) → mapping **conditionnel** (probable logique France/Groupe
  vs Hors Métropole, à rapprocher du cas « France vs Pays »). À préciser.
- `GR9201` = « Of which External Management Costs » → à relier au cas GR6000/GR6100
  (logique « of which / dont ») — **à confirmer**.
- Ceci **remplace/complète** les points « GFB110/GFB130 : mapping inconnu » et
  « GR5000/GR5300/GR5400 : TBD » du chapitre précédent — désormais **propositions en rouge**.

---

## Étape 4 — Exemple de sortie : P&L Tagetik (`Tagetik_PnL_exemple.xlsx`)

Fichier : `travail/Tagetik_PnL_exemple.xlsx` (fourni le 2026-08-20). 1 onglet, ~788 lignes.
C'est un **rendu type « rapport »** Tagetik (indicateurs en lignes, périodes en colonnes).

### Structure observée (lecture structurelle légère)
- **Lignes 6-7** = contexte période : `F6 = "Actuals 2026"`, `F7 = "March"`
  → la colonne **F** porte la valeur **Actuals 2026 / Mars** (cohérent avec 2026.03).
- À partir de la ligne 8, une ligne = un **indicateur Tagetik** :
  - **Col C** = flag `1`/`0` → probablement **1 = élément fin/saisissable**, **0 = agrégat/calculé** (à confirmer).
  - **Col D** = **code indicateur** `IND_xxx` (ex. `IND_00_070030`, `IND_00_060029`).
  - **Col E** = **libellé** de l'indicateur.
  - **Col F** = **montant** (Actuals 2026 / March).
  - **Col G…S** = colonnes additionnelles (autres mois/périodes ou scénarios) — à préciser.

### À retenir pour l'output du moteur
- La sortie finale devra produire, par indicateur `IND_`, un **montant** aligné sur ce format.
- On retrouve nos IND (ex. `IND_00_070030` = « Share of earnings of equity method »,
  déjà mappé 🔴 depuis `GR5400`).
- **À clarifier** : rôle exact de la colonne C (1/0), et si l'output moteur doit
  respecter **ce format rapport** (indicateurs en lignes + colonnes périodes) ou un
  **format d'import à plat** (une ligne = entité × compte × période × montant).

---

## Étape 4 bis — Hiérarchie complète de l'exemple Tagetik (au-delà du P&L)

Confirmé par l'utilisateur : **flag col C → `1` = élément fin (leaf), `0` = node (subtotal)**.
Le fichier ne contient **pas que le P&L**. Extraction complète → `docs/tagetik_indicateurs.csv`.

### Décomposition (736 indicateurs : 557 fins, 179 nodes)
- **Section P&L** : de l'index 0 jusqu'à **`IND_100000000` = « Profit & loss »** (index 342).
  Cascade finale : … `IND_111000000` Net Income → `IND_110000000` (w/o exceptional)
  → **`IND_100000000` Profit & loss**.
- **Section POST-P&L** (index 343→735) : au-delà du P&L :
  - **ETP / Headcount** (`IND_7xxxxxxxx`, ex. `IND_710000000` Headcount, `IND_26_030003`
    TOTAL ETP Opérationnel) → **cible des lignes ETP** (flux `Q99` côté Magnitude).
  - **KPI / Business Indicators** (`IND_720000000`, ESG, ratios de risque…).
  - **« Of which »** analytique (`IND_2xxxxxxxx`, ex. `IND_200000000` « Of which »).

### Lecture de la hiérarchie
- Les nodes numériques (`IND_1111322113` « Salaries incl. charges », `IND_111132211`
  « Staff costs »…) forment un **arbre par préfixe de chiffres** ; les codes alphanumériques
  `IND_00_0xxxxx` sont des **leaves** rattachées à l'arbre.

### Cibles clés repérées pour nos retraitements
- **`IND_00_060020` « Fixed remuneration »** → cible des **salaires fixes hors charges** (cas GR2100).
- **`IND_00_060074` « Social Charges »** → cible du **compte de charges sociales** séparé (cas GR2100).
  ⇒ Confirme la faisabilité du retraitement « taux de charges sociales par entité ».

### ⚠️ Écarts à reconcilier plus tard (ne pas trancher maintenant)
- Pour GR6000/GR6100 l'utilisateur avait cité `IND_00_070087` et **`IND_26_060074`**.
  Or dans cet exemple : `IND_00_070087` = « Of Which internal interest incomes » et
  `IND_00_060074` = « Social Charges » (préfixe `00` vs `26`). → **à revérifier ensemble**.
- Certains libellés diffèrent de la doc « 3.1 Indicators » (contextes distincts) → **à valider**
  au moment de figer la table de mapping.

---

## Étape 5 — Hiérarchies Entités & PMA (`Entites_PMA_hierarchie.xlsx`)

Fichier : `travail/Entites_PMA_hierarchie.xlsx` (fourni le 2026-08-20), 2 onglets, format rapport.
Ce sont les **listes de référence Tagetik** des cibles pour nos mappings **RU→Entité** et **OA→PMA**.
Extraits → `docs/hierarchie_pma.csv` et `docs/hierarchie_entites.csv`.

### Onglet 1 — PMA (50 codes `PMA_xxxx`)
- Équivalent Tagetik de l'**OA** (Operational Activity) Magnitude.
- Libellés préfixés **`S_`** (activité détaillée/sous-niveau) et **`A_`** (regroupement/agrégat),
  ex. `PMA_8000` « A_Services Immobiliers Promotion », `PMA_8015` « S_Office Property Developement ».

### Onglet 2 — Entités (356 codes, plusieurs préfixes)
- Équivalent Tagetik de la **RU** (entité) Magnitude. Répartition des préfixes :
  - **`EJ_`** (291) = **entités juridiques** → le gros du volume (niveau le plus fin).
  - **`EG_`** (59) = **entités de gestion/groupe** (ex. `EG_21700` « BNPP Real Estate France »,
    `EG_18` « Real Estate - Belgium »).
  - **`UG_`** (5) et **`EGMET`** (1) = autres regroupements.

### Structure hiérarchique (CE N'EST PAS UN LISTING PLAT)
- La hiérarchie est encodée par l'**indentation du libellé** (colonne D, via le style de cellule),
  **pas** par l'outline Excel (outline level = 0) ni par le flag col E (=0 partout ici).
- **Niveau = indentation** ; les **nodes/subtotals** sont les lignes **moins indentées**, placées
  **APRÈS** leurs enfants (layout « subtotal sous les enfants »). Feuilles = indentation max (3).

#### Onglet PMA — 45 feuilles (indent 3) + 6 nodes (indent 2)
Nodes (agrégats « A_ ») :
- `02_12430000` « Real Estate » (racine)
- `PMA_1617` A_Real Estate Property Management
- `PMA_8000` A_Services Immobiliers Promotion
- `PMA_8001` A_Services Immobiliers Advisory
- `PMA_8002` A_Services Immobiliers Residence Services
- `PMA_8003` A_Services Holding
→ Les 45 PMA « S_… » (feuilles) se regroupent sous ces agrégats.

#### Onglet Entités — 353 feuilles (indent 3) + 3 nodes
- `UG_39` (lvl1) « Real Estate Services scope of responsibility » (racine)
- `UG_390` (lvl2) « BNP Paribas Real Estate France - EUR »
- `EJ_21700` (lvl2) « BNPP Real Estate »
→ Les 353 entités feuilles (préfixes `EG_`, `EG_T_`, `EJ_`…) se rattachent à ces nodes.

#### Détail complet → CSV
`docs/hierarchie_pma.csv` et `docs/hierarchie_entites.csv` : colonnes
`code, libelle, niveau_indent, type (feuille|node/subtotal)`.

### Usage projet
- **Référentiel des cibles valides** pour construire les tables de correspondance
  **RU (Magnitude) → Entité** et **OA (Magnitude) → PMA**, qui sont **non 1-pour-1**
  (mapping à définir avec l'utilisateur — prochaine discussion).
- La hiérarchie servira aussi aux **contrôles** (rattacher une feuille au bon agrégat).

---

## Étape 6 — Mapping dimensionnel RU × OA × FA → Entité × PMA × Centre de coût

⭐ **Étape centrale du projet.** Le système cible Tagetik est nommé **« SOLARE »**
(colonne « Code SOLARE » = code Tagetik).

### Nomenclature des axes
| Magnitude | Tagetik (SOLARE)         | Remarque |
|-----------|--------------------------|----------|
| **RU**    | **Entité** (`EJ_`/`EG_`…)| 1 RU → **plusieurs** entités possibles (1→N) |
| **OA**    | **PMA** (`PMA_`)         | Business Line / activité opérationnelle |
| **FA** (Fonction) | **Centre de coût** | Axe Tagetik « Centre de coût » |

### Fichiers sources du mapping
- `RDHEngineQ1Copie.xlsx` :
  - onglet **« 1. Business Lines (2) »** → mapping **OA → PMA** (33 lignes) → `docs/mapping_OA_business_lines.csv`
  - onglet **« 2. Entities »** → mapping **RU → EJ** (33 lignes, avec devise) → `docs/mapping_RU_entities.csv`
- `RDHEngineQ1.xlsx` (le gros fichier) contient DÉJÀ un moteur avancé — onglets :
  `Feuil1, Liasse brute Q1 (~96 Mo données brutes), Liasse transfo, Rate,
  Filtres & Transfo, D_AC_Comptes, Synthèse, Mappable, Non mappable, Méthode & sources`.

### Le sujet OA : tags « NOT USED » non fiables
- Dans Business Lines, certains OA sont tagués `NOT USED` (ex. `OA011O`, `OA005O`,
  `OA061O`, `OA050`) **alors qu'ils apparaissent dans le jeu de données**.
  Ex. `G-SP × OA011O` figure dans Mappable (Vol 1.1). ⇒ **Ne pas se fier au tag NOT USED** ;
  se baser sur la présence réelle dans les données.

### Le sujet RU : 1 → N entités (désambiguïsation par OA/FA)
- Exemples 1→N : `G-UK` → 4 EJ, `G-SP` → 4 EJ, `G-DE` → 4 EJ, `G-ITMP` → 3 EJ,
  `G-BE` → 2 EJ, `G-PRTPROMO` → 2, `G-UKPROMO` → 3.
- ⇒ Le RU seul ne suffit pas : il faut la **combinaison RU × OA (× FA)** pour retrouver l'EJ.

### Travail déjà fait (session Claude précédente) : Mappable / Non mappable
Correspondances **RU × OA → EJ** reconstruites depuis le jeu de données réel :
- **`Mappable`** (90 combinaisons) → `docs/mapping_RUxOA_mappable.csv`
  Colonnes : `RU, Pays, OA, PMA, Activité, EJ cible, Niveau de confiance, Pourquoi, Vol`.
  La plupart = « Déterministe — RU mono-EJ » (le RU n'a qu'une EJ active).
- **`Non mappable`** (8 combinaisons) → `docs/mapping_RUxOA_non_mappable.csv`
  Colonnes : `RU, Pays, OA, Activité, EJ candidates, Raison, Explication, Pour débloquer, Vol`.
  Cas bloquants :
  - **`G-FRMP` (France)** promo/IM (OA015/OA041/OA045/OA060) : la source agrège sans
    dimension société ; la cible éclate sur 3 EJ (`EJ_21712`/`EJ_21665`/`EJ_67012`) →
    **clé de ventilation société absente**.
  - **`G-ASIEPF` (Plateforme Asie)** & **`G-MEPF` (Plateforme M-Orient)**, OA005 Holding :
    **aucune EJ Tagetik identifiée** dans le cube de mars.
  - **`GT-BCESTE`, `GT-TUT` (Transversal)**, OA005 Holding : **mailles techniques**, pas des EJ
    (centralisation / tutelle) → cible = maille technique (UG_OUT / tutelle), pas une société.

### Le sujet Centre de coût (FA) — et RU portés par l'axe Centre de coût
- **FA = Fonction** (Magnitude) → axe **Centre de coût** (Tagetik).
- Certains RU (**`G-ASIEPF`, `G-MEPF`, `G-SGP`**) sont en réalité **portés par l'axe Centre de coût** :
  au **croisement `EJ_41015` × `PMA_8033`** (Holding), **avec des Centres de coût différents**
  (à confirmer). ⇒ C'est la piste pour « débloquer » ces plateformes non mappables côté EJ.

### 🎯 Décision de conception : table de mapping à triple clé
- Construire **UNE table de mapping** à clé **(RU, OA, FA)** → **(ENTITE, PMA, CENTRE DE COÛT)**.
- **Pré-remplir** ce qui est déterministe (depuis Mappable / Business Lines / Entities),
  **laisser vide** ce qui ne l'est pas (Non mappable) → **saisie par les contrôleurs de gestion**.
- Conserver la colonne **Vol** (volume/matérialité) pour prioriser les cas à trancher.
- Prévoir la logique « un RU peut atterrir sur l'axe Centre de coût » (cas plateformes).

### Questions ouvertes
- Périmètre **IM** repris ou non (cf. G-FRMP OA060) — **à acter**.
- Cible Tagetik des **mailles techniques** (GT-BCESTE/GT-TUT) et des **plateformes** (Asie/M-Orient).
- Confirmer le triplet `EJ_41015 × PMA_8033 × Centre de coût` pour G-ASIEPF/G-MEPF/G-SGP.

### Principe : tout est modifiable (même le « mappable »)
- ⚠️ **Aucune valeur n'est verrouillée.** Même les combinaisons **déterministes/mappables**
  sont **pré-remplies comme propositions par défaut** que le contrôleur de gestion peut
  **modifier** dans la table de mapping.
- Implication conception :
  - La cible (`ENTITE` / `PMA` / `CENTRE DE COÛT`) de CHAQUE ligne est une **cellule éditable**,
    y compris pour les lignes déterministes.
  - On distingue visuellement l'**origine** de la valeur (proposée auto vs saisie/modifiée),
    ex. colonne « Source » (Auto déterministe / Auto proposé 🔴 / Saisie CDG / Modifié CDG),
    sans jamais empêcher la modification.
  - Le moteur (output Tagetik) doit **toujours lire la valeur de la table de mapping**
    (la version en vigueur, éventuellement modifiée), **jamais** une valeur figée en dur.

---

## Étape 7 — Schéma cible DÉFINITIF : table de fait Tagetik (`Tagetik_table_de_fait_EX4.xlsx`)

Fichier : `travail/Tagetik_table_de_fait_EX4.xlsx` — **extraction de la table de fait Tagetik**,
1 onglet `Sheet0`, **31 649 lignes**. C'est le **format d'import à plat** = ce que l'output du
moteur doit produire (question « rapport vs à plat » de l'étape 4 → **tranchée : à plat**).

### Colonnes (24) = dimensions × (code + description) + montant + origine
| Col | Dimension | Type | Valeur / source |
|-----|-----------|------|-----------------|
| A/B | **Scenario** | 🔒 constant | `2026AC` (Actuals 2026) |
| C/D | **Period** | 🔒 constant | `03` (March) — vient de `D_DP` |
| E/F | **Entity** | 🔗 mapping | 35 val. (`EJ_`/`EG_`/`EGMET_`) ← **RU** |
| G/H | **Indicator** | 🔗 mapping | 162 val. (`IND_`) ← **D_AC** (comptes) |
| I/J | **Counterparty** | 🔒 constant | `NA` |
| K/L | **PMA** | 🔗 mapping | 31 val. (dont `NA`) ← **OA** |
| M/N | **Product** | 🔒 constant | `NA` |
| O/P | **Vision** | 🔒 constant | `VIS_00_000001` (JV 100%) |
| Q/R | **Cost Center** | 🔗 mapping | 206 val. (`CC_`) ← **FA (Fonction)** |
| S/T | **Category** | ❓ à clarifier | 36 val. (`CTG…`) — voir ci-dessous |
| U/V | **Entity currency** | 🔗 donnée | 4 val. `EUR/GBP/PLN/SGD` ← **D_CU** |
| W | **Entity currency amount** | 💶 mesure | montant **en devise entité** (← `P_AMOUNT`) |
| X | **Origin** | 🔒 constant | `QDL` |

### Enseignements clés
- **6 dimensions constantes** pour cette reprise (Scenario, Period, Counterparty, Product,
  Vision, Origin) → à figer dans le moteur (paramètres du run).
- **Les 3 axes mappés se confirment** : `Entity ← RU`, `PMA ← OA`, `Cost Center ← FA`.
- **Le montant (W) est en DEVISE ENTITÉ** (pas toujours EUR : GBP/PLN/SGD présents)
  → ⚠️ d'où l'onglet **`rate`** pour convertir en EUR au moment du **contrôle** (étape 1 bis).
- **Nouvelle dimension `Category` (`CTG…`, col S)** : 36 valeurs, liée à la **cartographie/origine**
  du montant (ex. `CTG260084` « Amounts of Rotule IN imputed by Synthesis RE », `CTG260085`
  « Cartographie (dans PnL) », `CTG260034` « source DEF, isolé en LOC »).
  ❓ **À clarifier avec l'utilisateur** : comment déterminer la Category dans la reprise
  (constante ? déduite du compte/flux ? table dédiée ?).

### Impact conception moteur
- L'onglet **output** = ces 24 colonnes (ou au moins les codes A,C,E,G,I,K,M,O,Q,S,U,W,X requis
  à l'import + descriptions optionnelles).
- Chaque ligne source filtrée (étape 1) génère 1..N lignes output (cf. cas comptes 1→N)
  en résolvant : Entity/PMA/Cost Center via la **table de mapping (RU,OA,FA)**, Indicator via
  la **table comptes (D_AC)**, et en injectant les **constantes**.

### Category (`CTG…`) — RÉSOLU : hors périmètre
- La **Category** sert uniquement à **tracer** l'origine dans Tagetik (audit interne Tagetik).
- ⇒ **On ne s'en occupe pas** dans le moteur de reprise : dimension **ignorée** (non produite /
  laissée vide ou à une valeur par défaut selon ce qu'exige l'import — à voir au moment de l'import).

---

## Étape 8 — Mapping INTELLIGENT : règles « défaut + exceptions » (le plus spécifique gagne)

✅ **Oui, c'est possible et c'est le bon design.** La table de mapping n'est pas une simple
correspondance ligne à ligne : c'est un **jeu de règles** avec des **caractères génériques**
(`*` = « tout / toutes valeurs ») et une résolution par **spécificité**.

### Principe
- Clé source d'une règle = **(RU, OA, CC_source)** où **chaque composante** peut être :
  - une **valeur précise** (ex. `G-UK`, `OA010`, `CC_12001`), ou
  - **`*`** = « n'importe quelle valeur » (règle générale/défaut).
- Cible d'une règle = **(ENTITE, PMA, CC_cible)** (éditable, cf. étape « tout modifiable »).
- **Résolution** : pour une ligne de données `(RU, OA, CC)`, on prend la **règle qui matche
  la plus spécifique** (celle qui a le **moins de `*`**). En cas d'égalité → colonne
  **Priorité** (nombre, plus grand = gagne) pour trancher.

### Exemple (celui de l'utilisateur)
| # | RU | OA | CC source | → ENTITE | PMA | CC cible | Commentaire |
|---|----|----|-----------|----------|-----|----------|-------------|
| Règle défaut | `G-XX` | `OA010` | `*` (tous CC) | `EJ_aaa` | `PMA_bbb` | `CC_zzz` | cas général |
| Exception | `G-XX` | `OA010` | `CC_777` | `EJ_ccc` | `PMA_ddd` | `CC_888` | ce CC précis part ailleurs |
→ Une ligne `(G-XX, OA010, CC_777)` prend l'**exception** (plus spécifique).
   Toute autre `(G-XX, OA010, autre CC)` prend le **défaut**.

### Spécificité (du plus fort au plus faible)
1. RU + OA + CC tous précis
2. RU + OA précis, CC = `*`
3. RU précis, OA = `*`, CC = `*`
4. tout `*` (règle attrape-tout globale, optionnelle)
(La colonne **Priorité** permet de forcer un ordre si besoin, ex. exceptions transverses.)

### Mise en œuvre (Excel / Power Query, sans VBA)
- La table de règles porte des colonnes : `RU, OA, CC_source, ENTITE, PMA, CC_cible, Priorité, Source`.
- Le moteur, pour chaque ligne de données, calcule un **score de spécificité** par règle
  candidate (match exact = fort, `*` = faible) et retient la meilleure (puis Priorité).
- Réalisable en Power Query (jointures successives du plus spécifique au plus général, ou
  fonction de résolution) — **100 % faisable**.

### Bénéfices
- **Peu de lignes** pour couvrir beaucoup de cas (une règle défaut couvre « tous les CC »).
- **Exceptions ciblées** sans dupliquer tout le référentiel.
- Compatible avec le principe **« tout éditable »** et avec les cas **1→N** (comptes) déjà actés.

### Portée du mapping intelligent — LIMITÉE à RU × OA × CC
- ⚠️ **PAS de conso sur les flux & comptes.** Le mapping intelligent `*` s'applique **uniquement**
  aux axes **RU × OA × CC**.
- **Flux (`D_FL`)** = simple **filtre dur**, rien de plus : garder **`F99`** (financier) et
  **`Q99`** (ETP), point. Aucune règle, aucune combinaison, aucune conso sur le flux.
- **Comptes (`D_AC`)** = table de correspondance (étape 2), pas de mécanisme `*` dessus.
- ⇒ Ne PAS généraliser le wildcard aux flux/comptes (annule la piste évoquée précédemment).

---

## Environnement technique (acté)

- **Plateforme : Windows** → Power Query **complet**, aucune contrainte.
- **Version Excel : CONFIRMÉE** → **Excel pour Microsoft 365**, Version **2607** (Build 16.0.20228.20188),
  **64 bits**. Canal à jour, Power Query complet. Le **64 bits** est un vrai atout pour les gros
  volumes (extraction ~100 Mo, table de fait 31k+ lignes) : pas de limite mémoire du 32 bits.
- **Techno moteur : Power Query (langage M), SANS VBA / macro.** Actualisation par
  **Données → Actualiser tout**.

---

## Étape 9 — Ergonomie & public cible

### Séparation des rôles
- **Construire le moteur** = technique, fait **une seule fois** (nous). Les utilisateurs n'y touchent pas.
- **Utiliser le moteur** = **transparent** : coller l'extraction → Actualiser/Générer → lire l'output.
  Aucune manipulation de Power Query, aucune formule.

### Décisions
- **Habillage ergonomique (bouton unique vs Actualiser tout natif) : à décider PLUS TARD**
  (on construit le moteur d'abord, on choisit l'UX à la fin).
- **Saisie du mapping = par les Contrôleurs de Gestion (CDG)** — public **métier, non technique**.
  ⇒ Concevoir les onglets de mapping avec **garde-fous** :
  - **listes déroulantes** des cibles valides (Entités/PMA/CC/Indicateurs issus des référentiels),
  - **validation de données** (pas de saisie libre erronée),
  - **rouge = à valider** (propositions provisoires),
  - repères visuels d'origine (auto / à saisir / modifié).

### À prévoir pour l'usage transparent (quand on fera l'UX)
- Zone d'entrée claire « Collez ici l'extraction Magnitude ».
- Page « Mode d'emploi » (3 étapes).
- **Contrôles automatiques** avec voyants (totaux OK ✅ / lignes non mappées ⚠️ / équilibre EUR via rate).
- Signaler que la 1ʳᵉ actualisation prend quelques secondes (volume ~100 Mo).

---

## Étape 10 — Schéma réel de la source + axe FA→CC (fichier `CCEX`)

### 🔎 Schéma réel de « Liasse brute Q1 » (28 colonnes)
Les noms réels des colonnes source sont **préfixés `D_`** (l'étape 1 les nommait en raccourci) :
`A D_CA | B D_DP (période) | C D_OA (OA) | D **D_FA (Fonction)** | E D_VI | F D_TA | G D_PE |
H D_RU (RU) | I D_ORU | J D_AC (compte) | K D_FL (flux) | L D_AU | M D_T1 | N D_T2 |
O D_CU (devise) | P D_TO | Q D_GO | R D_LE | S D_NU | T D_DEST | U D_AREA | V D_MU | W D_PMU |
X P_AMOUNT (montant) | Y P_COMMENT | Z "PNB + MEE" | AA "OPEX" | AB "pre taxe incompe"`.
- **Les 3 dernières colonnes ajoutées par l'utilisateur** = **Z** « PNB + MEE », **AA** « OPEX »,
  **AB** « pre taxe incompe » (les « 3 dernières colonnes » mentionnées à l'étape 1).

### ⚠️ CORRECTION étape 1 : garder AUSSI `D_FA`
- La liste des colonnes à conserver devient : **`D_DP`, `D_OA`, `D_FA`, `D_RU`, `D_AC`, `D_FL`,
  `D_CU`, `P_AMOUNT`** (8 colonnes — `D_FA` avait été oublié, il est indispensable pour le Cost Center).
- Mapping des noms : `D_OA`=OA, `D_RU`=RU, `D_FA`=FA, `D_AC`=compte, `D_FL`=flux, `D_CU`=devise, `D_DP`=période.

### Fichier `CCEX` (`travail/Tagetik_CC_par_FA_CCEX.xlsx`) — CC groupés par FA
- Format rapport : **flag C** (`1`=CC feuille, `0`=FA node). Les **CC (`CC_…`)** sont les **enfants**
  de leur **FA (`FA00x`)**. **1 FA → N CC** (le CC est plus fin que le FA).
- Extrait → `docs/mapping_FA_to_CC.csv` : **62 liens**, **18 FA** (ex. FA015 Communication → 12 CC,
  FA025 Finance Services → 10 CC ; certains FA → 1 seul CC = déterministe).

### 🎯 Décision : CCEX = TABLE D'AIDE AU CHOIX (on ne choisit pas le CC nous-mêmes)
- ⛔ **On NE fait PAS le choix du CC automatiquement** — jugé **impossible** de façon fiable.
- ✅ On **transforme CCEX en table de mapping d'aide** : pour chaque **FA**, proposer la
  **liste des CC candidats** ; le **CDG choisit** le CC (liste déroulante, aide visuelle).
- Cas confirmé : **un même OA peut être détaillé sur deux CC** ⇒ le CC dépend aussi de l'**OA**,
  pas seulement du FA → d'où l'aide au choix (et le mécanisme de règles RU×OA×FA de l'étape 8).
- Quand un **FA n'a qu'un seul CC**, la proposition est **évidente** (pré-remplissable), mais reste
  **modifiable** (principe « tout éditable »).

### Reconciliation de la clé du mapping intelligent (étape 8)
- Source réelle de la clé = **(D_RU, D_OA, D_FA)** → cible **(ENTITE, PMA, CENTRE DE COÛT)**.
  (Le « CC_source » évoqué à l'étape 8 était imprécis : côté source c'est **FA** ; le **CC** est la **cible**,
  choisie par le CDG à l'aide de la table FA→CC.)

### Confirmations utilisateur (à figer)
- ✅ **`D_FA` est conservé** (correction de l'étape 1 validée).
- ✅ **Filtre OBLIGATOIRE : `D_OA` non vide** (exclure les OA vides) — reconfirmé.
- ✅ **Les 3 colonnes `Z` (PNB + MEE), `AA` (OPEX), `AB` (pre taxe incompe) = ajouts de CONTRÔLE
  de l'utilisateur.** La **vraie extraction Magnitude s'arrête juste avant** : colonnes **A → Y**
  (`D_CA` … `P_COMMENT`). ⇒ Le moteur **ne doit PAS dépendre** de Z/AA/AB (elles peuvent être
  absentes en production). Elles servent uniquement à recouper/valider pendant la mise au point.

---

## Étape 11 — Onglets d'AIDE AU CHOIX (référentiels) + architecture du classeur

Objectif utilisateur : **un maximum d'onglets d'aide au choix** (lecture seule), en **plus** de la
**table de mapping pilote**. Ils servent aux CDG pour saisir/valider sans se tromper, et alimentent
les **listes déroulantes** (validation de données).

### Référentiels à intégrer (déjà extraits en CSV → prêts à devenir des onglets)
| Onglet d'aide | Contenu | Source (CSV) |
|---------------|---------|--------------|
| **REF - Indicateurs Tagetik** | 736 IND, hiérarchie P&L + ETP/KPI/Of which, flag fin/node | `docs/tagetik_indicateurs.csv` |
| **REF - Hiérarchie Entités** | EJ/EG/UG, niveaux, feuilles/nodes | `docs/hierarchie_entites.csv` |
| **REF - Hiérarchie PMA** | PMA S_/A_, niveaux, feuilles/nodes | `docs/hierarchie_pma.csv` |
| **REF - CC par FA** | FA → CC candidats (18 FA, 62 CC), aide au choix du Cost Center | `docs/mapping_FA_to_CC.csv` |
| **REF - OA → PMA (actuel)** | liste des OA + mapping actuel Business Lines (dont NOT USED) | `docs/mapping_OA_business_lines.csv` |
| **REF - RU → EJ (actuel)** | liste des RU + mapping actuel + devise | `docs/mapping_RU_entities.csv` |
| **REF - RU×OA Mappable** | 90 combos résolus (EJ cible, confiance, Vol) | `docs/mapping_RUxOA_mappable.csv` |
| **REF - RU×OA Non mappable** | 8 combos bloqués (raison, pour débloquer, Vol) | `docs/mapping_RUxOA_non_mappable.csv` |
| **REF - Comptes D_AC** | mapping comptes Magnitude→IND + cas particuliers | (onglet `D_AC_Comptes` du fichier source, à reprendre) |

> Ces onglets **REF** sont **en lecture seule** (référence). Ils ne pilotent pas directement le moteur ;
> ils **aident la saisie** de la table de mapping pilote et **fournissent les listes** de valeurs valides.

### Table de mapping PILOTE (celle qui pilote réellement l'output)
- **MAP - RU×OA×FA → Entité×PMA×CC** : la table à **règles** (étape 8), **éditable**, avec
  wildcards `*`, priorité, colonne « Source » (auto/à valider 🔴/saisi/modifié). C'est **elle** que
  le moteur lit pour produire l'output. Pré-remplie depuis Mappable + Business Lines + Entities.
- Éventuellement une **MAP - Comptes** pilote (D_AC→IND) séparée, + un **paramétrage taux charges
  sociales par entité** (cas GR2100).

### Architecture d'onglets envisagée (ordre du classeur)
1. **Accueil / Mode d'emploi** (+ futur bouton, cf. UX à décider)
2. **1. ENTRÉE — Coller extraction Magnitude** (colonnes A→Y)
3. **2. Liasse filtrée** (Power Query : filtres F99/Q99, OA non vide, éléments fins…)
4. **MAP - …** (tables de mapping pilotes, éditables par CDG)
5. **Paramètres** (constantes Tagetik : Scenario 2026AC, Period 03, Vision, Origin… + taux charges soc.)
6. **REF - …** (tous les référentiels d'aide au choix, lecture seule)
7. **Contrôles** (voyants : totaux, non mappés, équilibre EUR via rate)
8. **N. SORTIE — Table de fait Tagetik** (24 colonnes, format à plat)

> Règle : **entrée en premier onglet, sortie en dernier onglet** (consigne initiale), mapping et
> référentiels au milieu.

---

## Étape 12 — Exemple de restitution Magnitude (onglets « Réalisé Vs Estimé » [Valeurs])

📍 **Emplacement réel** : ces 2 onglets sont dans **`RDHEngineQ1Copie.xlsx`** (pas le gros fichier).
C'est un **exemple de ce que donne une restitution via Magnitude** = le reporting **actuel**
(P&L de gestion **Réalisé vs Estimé**), pour cross-check — **pas** une brique du moteur.

### Les deux versions
- **« Réalisé Vs Estimé »** = avec les **formules GETDATA** (add-in Magnitude, tirage live du cube).
- **« Réalisé Vs Estimé (Valeurs) »** = les **valeurs/libellés résolus** (mêmes lignes, figées).

### Paramétrage GETDATA (en-tête, lignes 1-11)
`SC=BNPPI-GEST` (scope/société), `CC=EUR` (devise conso), `FA sum AR002`, **`FL=F99`**,
`CA=A` (Actual) / `CA=E` (Estimate), `DP=2026.03` / `PE=2026.03`, `AU sum GEST00`,
**`VA=RECOMPO`** (version), `OA sum (TYPE-OA=OA06)`. Comparaison **Actual Q1 2026 vs Estimate March**.

### Structure du corps (à partir de la ligne ~17)
| Col | Contenu |
|-----|---------|
| A | **définition OA** en syntaxe Magnitude : `{OA=OA040}+{OA=OA041}`, `OA sum AR028`… |
| B | **définition RU** : `{RU=G-FRMP}+{RU=P-FRMP}+{RU=G-FRAR}`, `RU sum AF026`… |
| C | **pays** (FR, ITA, ESP, DE, PRT, UK, LUX, BEL…) |
| D | **libellé de ligne** (Property Development, o/w France, Investment Management, o/w Italy…) |
| E…BR | **blocs de valeurs** : Actual Q1 2026 / Estimate March / Variances, répétés par agrégat |

### Enseignements / points d'attention
- Les lignes agrègent des **combinaisons RU × OA** (et des **groupes nommés** Magnitude
  `OA sum AR028`, `RU sum AF026`, `FA sum AR002`) → logique de reporting existante.
- Apparition de **RU préfixés `P-`** (`P-FRMP`, `P-CENTRAL`, `P-UK`…) à côté des `G-` :
  ❓ à clarifier (probable **périmètre/plan** vs réel `G-`) — **noté comme question**.
- Filtre **`FL=F99`** cohérent avec notre filtre flux. Conso en **EUR** (`CC=EUR`) → cohérent
  avec le contrôle EUR via l'onglet `rate`.
- Usage projet : **référence de recoupement** (structure P&L, périmètre RU/pays, définitions de
  lignes) et **modèle possible** pour un onglet de **contrôle/restitution** comparatif. Pas repris tel quel.

### Question ouverte ajoutée
- Signification des **RU `P-…`** vs `G-…` (périmètre/plan ?) et des **groupes nommés** `AR0xx`/`AF0xx`.
