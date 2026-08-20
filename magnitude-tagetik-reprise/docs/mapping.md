# Spécification du mapping — Magnitude → Tagetik

> À figer avec les fichiers d'exemple. Les colonnes ci-dessous sont un modèle
> générique EPM (consolidation) ; à remplacer par les noms de champs réels.

## Dimensions attendues

| Dimension Tagetik | Champ Magnitude (source) | Règle de transformation                    |
|-------------------|--------------------------|--------------------------------------------|
| Entité            | *(à préciser)*           | Correspondance code entité (table de corr.)|
| Compte            | *(à préciser)*           | Correspondance plan de comptes             |
| Période           | *(à préciser)*           | Format Tagetik (ex. `M01`…`M12` / `AAAA`)  |
| Scénario / Version| *(à préciser)*           | Ex. ACTUAL, historique figé                |
| Flux / Mouvement  | *(à préciser)*           | Si géré côté Tagetik                        |
| Devise            | *(à préciser)*           | Devise locale / de conso                    |
| Montant           | *(à préciser)*           | Signe, décimales, séparateur               |

## Tables de correspondance

- **Entités** : `code_magnitude → code_tagetik`
- **Comptes**  : `compte_magnitude → compte_tagetik`
- **Périodes** : mapping calendaire si les conventions diffèrent

Ces tables seront des onglets du classeur, chargés comme requêtes Power Query
et jointes à la source (merge) pour appliquer les correspondances.

## Contrôles de recette

1. **Équilibre par entité/période** : Σ montants Magnitude = Σ montants Tagetik.
2. **Lignes non mappées** : toute valeur source sans correspondance est listée
   dans un onglet « Rejets » (jamais perdue silencieusement).
3. **Comptes de contrôle** : totaux bilan / P&L avant et après reprise.
