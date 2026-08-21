// ============================================================
//  Filtre — règles de l'étape 1
//   • flux F99 (financier) et Q99 (ETP) uniquement
//   • OA (D_OA) non vide
//   • comptes fins uniquement (les nœuds/agrégats sont exclus
//     via la table tblComptes : Type = "Compte fin")
// ============================================================
let
    Source = Filtre_base,   // = requête "Source"

    // 1) Flux F99 / Q99
    fFlux = Table.SelectRows(Source, each ([D_FL] = "F99" or [D_FL] = "Q99")),

    // 2) OA non vide
    fOA = Table.SelectRows(fFlux, each [D_OA] <> null and Text.Trim([D_OA]) <> ""),

    // 3) Comptes fins : on récupère la liste des comptes fins depuis tblComptes
    Comptes = Excel.CurrentWorkbook(){[Name="tblComptes"]}[Content],
    Fins = Table.SelectRows(Comptes, each [Type] = "Compte fin"),
    ListeFins = List.Buffer(List.Distinct(Fins[Code])),
    fFin = Table.SelectRows(fOA, each List.Contains(ListeFins, [D_AC]))
in
    fFin
// NB : remplacez "Filtre_base" par la référence à la requête Source
//      (Éditeur : renommez, ou first line  Source = Source).
