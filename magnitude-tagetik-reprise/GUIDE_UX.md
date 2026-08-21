# Guide ergonomie (Étage A — sans VBA)

Après avoir recollé `07_Sortie.pq` (ajoute la colonne **Statut**) et `08_AMapper.pq`
(ajoute **Où mapper ?** et trie par volume), fais ces réglages **une seule fois**.

## 1) Rouge automatique dans la SORTIE (ce qui n'est pas mappé)
La requête Sortie a maintenant une colonne **Statut** = `OK` ou `À MAPPER`.
1. Clique dans le tableau **Sortie** (sur la feuille où il est chargé).
2. Sélectionne toute la colonne **Statut** (clic sur son en-tête dans le tableau).
3. **Accueil** ▸ **Mise en forme conditionnelle** ▸ **Règles de mise en surbrillance** ▸ **Texte qui contient…**
4. Tape `À MAPPER` ▸ choisis **Remplissage rouge clair / Texte rouge foncé** ▸ OK.

👉 Pour surligner **toute la ligne** (plus visible) :
- Sélectionne toute la plage de données du tableau Sortie (sans les en-têtes).
- **Mise en forme conditionnelle** ▸ **Nouvelle règle** ▸ **Utiliser une formule**.
- Formule : `=$K2="À MAPPER"`  *(remplace `K` par la lettre de la colonne Statut, et `2` par la 1re ligne de données)*.
- **Format** ▸ Remplissage rouge ▸ OK.

## 2) Rouge automatique dans AMapper
1. Clique dans le tableau **AMapper**.
2. Sélectionne la colonne **Code manquant**.
3. **Mise en forme conditionnelle** ▸ **Nouvelle règle** ▸ **Mettre en forme toutes les cellules** ▸
   applique un **remplissage rouge** (ou règle « Texte qui contient » vide = tout).
   (AMapper ne liste QUE des manquants → tu peux tout mettre en rouge.)
La colonne **« Où mapper ? »** indique déjà l'onglet à ouvrir pour chaque code.

## 3) Tableau de bord CONTRÔLES (formules toutes prêtes)
Sur l'onglet **CONTRÔLES**, en colonne B en face de chaque libellé, colle (le tableau chargé
de la requête Sortie s'appelle `Sortie`) :

| Contrôle | Formule à coller |
|----------|------------------|
| Lignes en sortie | `=NBVAL(Sortie[Scenario])` |
| Lignes NON mappées | `=NB.SI(Sortie[Statut];"À MAPPER")` |
| % de couverture | `=1-NB.SI(Sortie[Statut];"À MAPPER")/NBVAL(Sortie[Statut])` |
| Total montant (devise entité) | `=SOMME(Sortie[Entity currency amount])` |

Puis mets une **MFC** sur « % de couverture » : vert si `>=0,99`, orange sinon (icônes ou remplissage).

## 4) Protéger les référentiels (éviter les modifs par erreur)
Pour chaque onglet **REF - …** : clic droit sur l'onglet ▸ **Protéger la feuille** ▸ OK
(laisse le mot de passe vide). Les REF deviennent lecture seule ; les MAP restent éditables.

## 5) Navigation rapide
- L'onglet **0. Accueil** contient déjà un **sommaire cliquable** (chaque onglet est un lien).
- Depuis AMapper, la colonne **« Où mapper ? »** te dit où aller ; clique l'onglet correspondant
  (ou reviens à l'Accueil et clique le lien).

> Le « double-clic qui saute directement à la bonne ligne » et les boutons = Étage B (VBA / .xlsm),
> non inclus ici par choix. On pourra l'ajouter plus tard si tu veux.
