# Moteur Power Query — mode d'emploi

Les **Tableaux Excel sont déjà créés** dans le classeur (tblEntree, tblRUEJ, tblOAPMA,
tblFACC, tblDim, tblComptes, tblETP, tblExcl, tblTaux, tblZone, tblRate) — **rien à faire côté Ctrl+T**.

## Coller les requêtes
Pour chaque `.pq` : **Données ▸ Obtenir des données ▸ À partir d'une requête vide ▸ Éditeur avancé**,
coller le contenu, nommer la requête comme le fichier (sans le numéro).

## Ordre d'enchaînement
Source → Filtre → Exclusions → MapComptes → MapDimensions → Constantes → Sortie
Seule **Sortie** est chargée dans une feuille (⑨ SORTIE) ; les autres restent en « connexion seule ».

## Chaîne de mapping
- `tblRUEJ` (défaut), `tblOAPMA`, `tblFACC` (défaut) alimentent les défauts.
- **`tblDim` (MAP - Dimensions) a le dernier mot** : Entité toujours prise de tblDim ;
  PMA/CC pris de tblDim si renseignés, sinon des tables OA→PMA / FA→CC.

## Note
Scripts structurés et commentés. Quelques finitions (retraitement charges sociales GR2100 en lignes,
formats) se calent une fois dans Excel — balisées dans le code.
