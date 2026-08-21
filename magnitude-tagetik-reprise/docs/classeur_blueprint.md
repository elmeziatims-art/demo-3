# Classeur « Reprise Magnitude → Tagetik » — Blueprint acté

> Structure définitive du classeur (Excel 365 64 bits, Power Query, sans VBA).
> Ordre : **entrée en premier**, **sortie en dernier**, mapping + référentiels au milieu.
> Principe : **tous les mappings éditables par les CDG** ; le moteur lit toujours la version en vigueur.

## Vue d'ensemble des onglets

| # | Onglet | Type | Rôle |
|---|--------|------|------|
| 0 | **🏠 Accueil** | Guide | Mode d'emploi (3 étapes) + voyants de contrôle + (bouton futur) |
| 1 | **① ENTRÉE – Magnitude** | Saisie | Coller l'extraction brute (colonnes A→Y) |
| 2 | **② Liasse filtrée** | Auto (PQ) | Résultat des filtres — masquable |
| 3 | **🗺️ MAP – Comptes P&L** | Éditable | `D_AC → IND` (France/Pays) |
| 4 | **🗺️ MAP – ETP** | Éditable | `D_AC (Q99) → IND` headcount |
| 5 | **🗺️ MAP – Dimensions** | Éditable | Règles `RU×OA×FA → Entité·PMA·CC` |
| 6 | **🗺️ MAP – Taux charges soc.** | Éditable | Un % par RU |
| 7 | **🗺️ MAP – Zone RU** | Éditable | RU → Pays → France/Pays |
| 8 | **⚙️ PARAM – Constantes** | Éditable | Scenario, Period, Vision, Origin… |
| 9 | **📚 REF – …** (plusieurs) | Lecture seule | Référentiels d'aide au choix |
| 10 | **💱 rate** | Éditable | Taux de change (contrôle EUR) |
| 11 | **✅ CONTRÔLES** | Auto | Totaux blocs, non mappés, équilibre EUR |
| 12 | **⑨ SORTIE – Table de fait** | Auto (PQ) | 24 colonnes, format à plat Tagetik |

## Détail des colonnes

### ① ENTRÉE – Magnitude
Zone « collez ici ». Colonnes source A→Y (`D_CA … P_COMMENT`). Les 3 colonnes de contrôle
(PNB+MEE/OPEX/Pre-tax) ne sont **pas** requises par le moteur.

### ② Liasse filtrée (Power Query)
8 colonnes retenues : `D_DP, D_OA, D_FA, D_RU, D_AC, D_FL, D_CU, P_AMOUNT`.
Filtres : `D_FL ∈ {F99, Q99}` · `D_OA` non vide · `D_AC` = comptes fins (exclut nœuds/agrégats).

### MAP – Exclusions / Ajustements (éditable)
`RU | OA | FA | Mode (Montant|%) | Valeur | Commentaire`. Mécanisme générique : retirer X% ou X€ d'un
`RU × OA × FA`, **case vide = toutes les valeurs** de la dimension. **Appliqué AVANT le mapping**,
au prorata sur les lignes source du périmètre, sur **F99 ET Q99**. (REIM = simple exemple.)

### ③ MAP – Comptes P&L
`Code | Libellé | Bloc_PnL | Type | blk_PNB_MEE | blk_OPEX | blk_PreTax | IND_France | IND_Pays | Statut | Commentaire`
Gère 1→N (GR6000/6100 total+détail), exclusions (GR2000, GR200211), retraitement charges (GR2100).

### ④ MAP – ETP (`Q99`)
`Code_Magnitude | IND_Tagetik | IND_libellé | Statut | Commentaire`
(GR050 🔴 à mapper, GR051→IND_26_100027, GR052→IND_26_100030.)

### ⑤ MAP – Dimensions (règles)
`RU | OA | FA | → ENTITE | PMA | CC | Priorité | Source | Commentaire`
Jokers `*` autorisés ; résolution « la plus spécifique gagne » (+ Priorité). Pré-rempli déterministe, éditable.

### ⑥ MAP – Taux charges sociales / RU
`RU (liste) | Taux | Commentaire` — un seul % par RU ; appliqué ligne par ligne.

### ⑦ MAP – Zone RU
`RU | Pays | Zone (France/Pays)` — dérivé de Mappable/Non mappable, éditable.

### ⑧ PARAM – Constantes
- **Scenario & Period DÉDUITS de `D_DP`** (colonnes calculées, pas figées) :
  `D_DP = 2026.03` → **Scenario = `2026AC`** (année + `AC`), **Period = `03`** (le mois).
  ⇒ Le moteur est **indépendant de la période** : il s'adapte à toute valeur de `D_DP`.
- Vraies constantes : `Vision=VIS_00_000001 | Origin=QDL | Counterparty=NA | Product=NA`.

### ⑨ REF – … (lecture seule, alimentent les listes déroulantes)
Indicateurs Tagetik · Hiérarchie Entités · Hiérarchie PMA · FA→CC · OA→PMA · RU→EJ ·
RU×OA Mappable · RU×OA Non mappable.

### ⑪ CONTRÔLES
Totaux par bloc P&L (PNB+MEE / OPEX / Pré-tax) ; liste des lignes **non mappées** ⚠️ ;
équilibre **EUR** (via `rate`) ; nb de lignes en entrée vs sortie.

### ÉTAT 1 – Restitution Magnitude (façon « Réalisé Vs Estimé Valeurs »)
Vue P&L en **EUR** construite sur la donnée Magnitude (blocs PNB+MEE / OPEX / Pré-tax) → réconciliation.

### ÉTAT 2 – Restitution SolaRE (PMA × Entités)
Vue **EUR**, PMA en tête + entités dessous, **hiérarchie complète** (nœuds agrégés PMA/entités, repli/dépli).

### ⑫ SORTIE – Table de fait Tagetik (24 colonnes)
`Scenario | Scenario-Desc | Period | Period-Desc | Entity | Entity-Desc | Indicator | Indicator-Desc |
Counterparty(NA) | Counterparty-Desc | PMA | PMA-Desc | Product(NA) | Product-Desc | Vision | Vision-Desc |
Cost Center | Cost Center-Desc | Category | Category-Desc | Entity currency | … | Entity currency amount | Origin`
Montant en **devise entité**. Category hors périmètre (traçage Tagetik).

## Geste utilisateur courant
**Coller (①) → Données ▸ Actualiser tout → lire/copier (⑫).** Les CDG interviennent uniquement sur les onglets **MAP** / **PARAM**.
