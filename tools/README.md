# Alignement de la mise en page COCKPIT → CAD_PIL

`CAD_PIL.xlsx` (onglets *Cadrage* et *Pilotage*) porte la charte de mise en page
maison. `COCKPIT.xlsx` avait été construit sur une charte différente. Ces deux
scripts réappliquent celle de CAD_PIL aux trois onglets visibles du cockpit.

```bash
python3 tools/harmonize_cockpit.py COCKPIT.xlsx COCKPIT_aligne.xlsx
```

* `xlsxpatch.py` — édition d'un `.xlsx` au niveau XML. Les graphiques, images,
  mises en forme conditionnelles (extensions x14 comprises) et noms définis
  traversent le traitement à l'octet près, ce qu'un aller-retour openpyxl ne
  permet pas : openpyxl ne sait pas relire ces graphiques et supprime les
  extensions de MFC.
* `harmonize_cockpit.py` — la charte elle-même : jetons de couleur, échelle
  typographique, bandeau de titre, bandeau d'indicateurs, titres de section,
  filets de tableau, rythme des lignes et réglages de feuille.

Aucune valeur ni formule n'est modifiée ; seuls les libellés de titres et
d'en-têtes sont réécrits, pour suivre la convention de CAD_PIL.
