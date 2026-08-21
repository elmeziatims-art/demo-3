# Moteur Power Query — mode d'emploi

Ces 7 requêtes forment le moteur. À coller dans Excel une par une :
**Données ▸ Obtenir des données ▸ À partir d'une requête vide ▸ Éditeur avancé**,
puis coller le contenu du `.pq` et nommer la requête comme le fichier (sans le numéro).

## Prérequis — nommer les plages en Tableaux Excel
Sélectionner chaque plage puis Insertion ▸ Tableau (Ctrl+T), et la renommer (onglet Création de tableau) :

| Onglet | Nom de tableau |
|--------|----------------|
| ① ENTREE | `tblEntree` |
| MAP - Comptes P&L | `tblComptes` |
| MAP - ETP (Q99) | `tblETP` |
| MAP - Dimensions | `tblDim` |
| MAP - Exclusions | `tblExcl` |
| MAP - Taux charges soc. | `tblTaux` |
| MAP - Zone RU | `tblZone` |
| rate | `tblRate` |

## Ordre d'enchaînement
Source → Filtre → Exclusions → MapComptes → MapDimensions → Constantes → Sortie
Seule **Sortie** est chargée dans une feuille (⑨ SORTIE) ; les autres restent en « connexion seule ».

## Note
Scripts fournis structurés et commentés. Certaines finitions (formats de nombre,
noms de colonnes exacts après jointure) sont à ajuster une fois dans Excel — le squelette
logique est complet.
