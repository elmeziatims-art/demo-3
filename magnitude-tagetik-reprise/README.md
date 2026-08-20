# Moteur de reprise d'historique — Magnitude → Tagetik

Outil de reprise de l'historique de consolidation depuis **Magnitude** vers **Tagetik**,
sous forme d'un **classeur Excel piloté par Power Query** — sans VBA.

## Principe

L'export Magnitude est chargé dans Excel via Power Query, transformé selon un mapping
paramétrable, puis restitué dans un tableau conforme au **format d'import Tagetik**.
Le classeur se rafraîchit d'un clic (Données → Actualiser tout) — aucune macro.

```
  Export Magnitude (.xlsx/.csv)          Gabarit d'import Tagetik
            │                                      ▲
            ▼                                      │
   [ Power Query : requête « Source » ]            │
            │                                      │
            ▼                                      │
   [ Mapping paramétrable (table de correspondance) ]
            │                                      │
            ▼                                      │
   [ Requête « Sortie_Tagetik» ] ──────────────────┘
```

## Structure du dépôt

| Dossier      | Contenu                                                          |
|--------------|------------------------------------------------------------------|
| `entrees/`   | Exports Magnitude à reprendre (fichiers sources)                 |
| `sorties/`   | Fichiers d'import générés au format Tagetik                      |
| `docs/`      | Spécification du mapping des colonnes et des dimensions          |
| racine       | Le classeur `Reprise_Magnitude_Tagetik.xlsx` (moteur Power Query)|

## Étapes de mise en place

1. **Fichiers d'exemple** — déposer dans `entrees/` un export Magnitude réel et dans
   `docs/` le gabarit d'import Tagetik. Ils servent à figer le mapping exact.
2. **Mapping** — renseigner la table de correspondance (voir `docs/mapping.md`).
3. **Construction du classeur** — assemblage des requêtes Power Query.
4. **Recette** — comparaison des totaux Magnitude vs Tagetik (contrôle d'équilibre).

## Statut

🚧 En cours de cadrage — en attente des fichiers d'exemple (export Magnitude + gabarit Tagetik).
