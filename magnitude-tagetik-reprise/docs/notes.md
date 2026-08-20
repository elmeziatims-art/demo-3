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
