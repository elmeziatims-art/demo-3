# Les 2 états de restitution (sans VBA)

Les requêtes `Etat_SolaRE` et `Etat_Magnitude` agrègent les montants **en EUR**.
On les affiche en **Tableau croisé dynamique (TCD)** pour avoir la hiérarchie repliable + les
segments Scénario / Période.

## Créer les requêtes
Comme les autres : Requête vide ▸ Éditeur avancé ▸ coller ▸ renommer.
- `09_Etat_SolaRE.pq`  → nom **`Etat_SolaRE`**
- `10_Etat_Magnitude.pq` → nom **`Etat_Magnitude`**
Charge chacune : **Fermer et charger dans… ▸ Tableau croisé dynamique** ▸ **Feuille existante**,
sur l'onglet correspondant (« ETAT - Restitution SolaRE » / « ETAT - Restitution Magnitude »), cellule A5.
*(Si tu préfères, charge-les en « connexion seule » puis Insertion ▸ TCD à partir de la requête.)*

## État SolaRE (PMA × Entités)
Dans le TCD :
- **Lignes** : `PMA` puis `Entity` (dans cet ordre → PMA en tête, entités dessous).
- **Valeurs** : `Montant_EUR` (Somme).
- **Segments** (Insertion ▸ Segment) : `Scenario` et `Period`.
- Le +/- à gauche des PMA replie/déplie les entités (la hiérarchie que tu voulais).

## État Magnitude (P&L)
Dans le TCD :
- **Lignes** : `Bloc_PnL` puis `D_AC` (compte).
- **Valeurs** : `Montant_EUR` (Somme) — tu peux ajouter `Montant_devise`.
- **Segments** : `Scenario` et `Period`.
→ Vue P&L par bloc (PNB+MEE / OPEX / Pré-tax), comparable à la restitution Magnitude.

## Astuce présentation
- Création du TCD ▸ **Disposition en mode Plan** (Outils de tableau croisé ▸ Création ▸ Disposition du rapport)
  pour un rendu proche d'une liasse.
- Active **les sous-totaux** (ils sont là par défaut) = les nœuds/agrégats.

> Note : les montants EUR ne s'affichent que si le montant source est bien lu — recolle d'abord
> le `01_Source.pq` blindé (sinon Montant_EUR = vide).
