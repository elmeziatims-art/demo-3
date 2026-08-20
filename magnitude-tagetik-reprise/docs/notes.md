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
