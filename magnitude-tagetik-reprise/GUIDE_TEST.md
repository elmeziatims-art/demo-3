# Guide de test — brancher le moteur Power Query (pas-à-pas)

Objectif : coller les 8 requêtes, actualiser, et voir la SORTIE se remplir.
Durée ~15 min. Aucun VBA. Les Tableaux Excel sont déjà créés.

## Étape 0 — Ouvrir
Ouvre `Reprise_Magnitude_Tagetik.xlsx`. Si un bandeau jaune « Activer la modification » apparaît, clique dessus.
(Un échantillon de 1 500 lignes est déjà dans « ① ENTREE » pour le test.)

## Étape 1 — Créer les requêtes (dans CET ORDRE)
Pour CHAQUE requête ci-dessous :
1. Ruban **Données** ▸ **Obtenir des données** ▸ **À partir d'autres sources** ▸ **Requête vide**.
2. Dans l'éditeur qui s'ouvre : **Accueil** ▸ **Éditeur avancé**.
3. **Efface tout**, colle le contenu du fichier `.pq` correspondant, clique **Terminé**.
4. À gauche (volet Requêtes), **clic droit sur la requête ▸ Renommer** → mets EXACTEMENT le nom indiqué.
5. **Accueil** ▸ **Fermer et charger dans…** ▸ choisis **« Ne créer que la connexion »** (sauf pour les 2 dernières, voir étape 2).

Ordre et noms (le nom doit être exact, les requêtes s'appellent entre elles) :
| # | Fichier | Nom de la requête |
|---|---------|-------------------|
| 1 | 01_Source.pq | `Source` |
| 2 | 02_Filtre.pq | `Filtre` |
| 3 | 03_Exclusions.pq | `Exclusions` |
| 4 | 04_MapComptes.pq | `MapComptes` |
| 5 | 05_MapDimensions.pq | `MapDimensions` |
| 6 | 06_Constantes.pq | `Constantes` |
| 7 | 07_Sortie.pq | `Sortie` |
| 8 | 08_AMapper.pq | `AMapper` |

> Astuce : si une requête affiche une erreur du type « Source introuvable », c'est que la
> requête qu'elle appelle n'a pas encore été créée / mal nommée. Respecte l'ordre et les noms.

## Étape 2 — Charger les 2 sorties dans des feuilles
- Requête **`Sortie`** : **Fermer et charger dans…** ▸ **Tableau** ▸ **Feuille de calcul existante**
  ▸ clique l'onglet **« ⑨ SORTIE - Table de fait »**, cellule **A5** ▸ OK.
- Requête **`AMapper`** : **Fermer et charger dans…** ▸ **Tableau** ▸ onglet
  **« À MAPPER (auto) »**, cellule **A5** ▸ OK.

## Étape 3 — Actualiser & lire
- **Données** ▸ **Actualiser tout**.
- Va dans **⑨ SORTIE** : la table de fait Tagetik se remplit.
- Va dans **À MAPPER (auto)** : liste des RU/OA/FA pas encore mappés (à compléter dans les MAP).

## Étape 4 — Vérifier
- Compare quelques montants avec l'échantillon.
- Regarde les colonnes Entity / Indicator / PMA / Cost Center : elles doivent être remplies.
- Les lignes non résolues → onglet « À MAPPER (auto) ».

## Remettre la vraie extraction
Quand le test est concluant : dans « ① ENTREE », remplace l'échantillon par ta vraie extraction
Magnitude (colle à partir de la ligne 5, sous les en-têtes), puis **Actualiser tout**.

## En cas de souci
- Le montant `P_AMOUNT` doit être numérique (pas de texte). Si l'extraction a des « . » décimaux,
  Power Query les gère ; si problème de format régional, dis-le-moi.
- Le retraitement charges sociales (GR2100) et quelques finitions sont balisés dans les .pq — on les
  active dans un 2e temps une fois le socle validé.
