# Architecture du classeur — Reprise Magnitude → Tagetik

> Vue d'ensemble de l'organisation du classeur Excel (Power Query, sans VBA),
> de l'extraction Magnitude jusqu'à l'output Tagetik. Détail métier dans `notes.md`.

## Flux de bout en bout

```
① Extraction Magnitude  →  ② Filtrage  →  ③④ Mapping (comptes + dimensions)  →  Constantes & lignes 1→N  →  ⑨ Table de fait Tagetik
      (coller A→Y)          (F99/Q99…)       (IND / Entité·PMA·CC)                 (Scenario, Period…)          (24 colonnes, à plat)
                                                                        ↘ ⑧ Contrôles (totaux, non mappés, EUR via rate)
```

## Onglets (ordre du classeur)

| # | Onglet | Rôle | Édition |
|---|--------|------|---------|
| 0 | **Accueil / Mode d'emploi** | 3 étapes + futur bouton | — |
| 1 | **① ENTRÉE – Extraction Magnitude** | zone « collez ici » (colonnes A→Y) | Utilisateur (colle) |
| 2 | **② Liasse filtrée** | résultat Power Query (filtres) | auto (masquable) |
| 3 | **MAP – Comptes** | `D_AC → IND` + cas + taux charges sociales/entité | CDG |
| 4 | **MAP – RU×OA×FA** | règles `→ Entité·PMA·CC` (wildcards `*`, priorité) | CDG |
| 5 | **PARAM – Constantes** | Scenario `2026AC`, Period `03`, Vision `VIS_00_000001`, Origin `QDL`, Counterparty `NA`, Product `NA` | Toi (rare) |
| 6 | **REF – …** | référentiels d'aide au choix (lecture seule) | — |
| 7 | **rate** | taux de change (contrôle EUR) | Toi (maj période) |
| 8 | **CONTRÔLES** | voyants totaux / non mappés / équilibre EUR | — |
| 9 | **⑨ SORTIE – Table de fait Tagetik** | 24 colonnes, format à plat | Utilisateur (copie) |

### Onglets REF (aide au choix, lecture seule)
Indicateurs Tagetik · Hiérarchie Entités · Hiérarchie PMA · CC par FA ·
OA→PMA (actuel) · RU→EJ (actuel) · RU×OA Mappable · RU×OA Non mappable · Comptes D_AC.
(Alimentent les listes déroulantes des onglets MAP.)

## Pipeline Power Query (requêtes, invisibles pour l'utilisateur)

1. **qSource** — lit l'onglet ENTRÉE ; garde 8 colonnes : `D_DP, D_OA, D_FA, D_RU, D_AC, D_FL, D_CU, P_AMOUNT`.
2. **qFiltre** — `D_FL ∈ {F99, Q99}` ; `D_OA` non vide ; `D_AC` éléments fins (exclut nœuds).
3. **qMapComptes** — `D_AC → IND` :
   - 1→N (total + détail : GR6000/GR6100 → `IND_00_070008` + détail) ;
   - exclusions (GR2000 = nœud ; GR200211 abandonné) ;
   - variante **France / Pays** ;
   - **ventilation charges sociales** (GR2100 : déduire GR2101/GR2102, split hors charges / charges via **taux par entité**).
4. **qMapDim** — résout `(D_RU, D_OA, D_FA) → (Entité, PMA, Cost Center)` par **règle la plus spécifique** (wildcards `*`, priorité). Cost Center aidé par table FA→CC (choix CDG).
5. **qConstantes** — injecte les 6 constantes + `Entity currency` (← `D_CU`) + `amount` (← `P_AMOUNT`).
6. **qSortie** — assemble les **24 colonnes** de la table de fait → charge dans l'onglet SORTIE.
7. **qContrôle** — agrégats de vérification (par entité / indicateur) ; conversion EUR via `rate` ; liste des lignes **non mappées**.

## Geste utilisateur (usage courant)
**Coller (①) → Données ▸ Actualiser tout → lire/copier (⑨).**
Les CDG n'interviennent que sur **MAP – Comptes** et **MAP – RU×OA×FA** (listes déroulantes, rouge = à valider).

## Principes transverses
- **Sans VBA** ; 100 % Power Query (langage M). Excel 365 64 bits (Windows).
- **Tout éditable** : les valeurs pré-remplies sont des propositions modifiables ; le moteur lit toujours la table en vigueur.
- **Format de sortie** : table de fait à plat (24 colonnes), montant en **devise entité**.
- **Category (CTG)** : hors périmètre (traçage Tagetik).

## Points ouverts (voir notes.md)
- Détail de calcul de la « dernière partie » (à préciser par l'utilisateur).
- RU `P-…` vs `G-…` (périmètre/plan ?), groupes nommés `AR0xx`/`AF0xx`.
- Périmètre IM ; cible des mailles techniques/plateformes ; triplet `EJ_41015 × PMA_8033 × CC`.
- Comptes TBD : mapping GFB110/GFB130 (rouge), GR5000/5300/5400 (rouge à valider).
