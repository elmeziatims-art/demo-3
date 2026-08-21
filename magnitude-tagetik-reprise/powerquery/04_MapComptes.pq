// ============================================================
//  MapComptes — D_AC -> Indicateur (IND)
//   • Zone France/Pays (via tblZone) choisit IND_France ou IND_Pays
//   • flux F99 -> mapping P&L (tblComptes) ; Q99 -> mapping ETP (tblETP)
//   • cas 1->N (ex. GR6000/GR6100 : total 070008 + détail) : voir table
//     tblComptes qui contient déjà les lignes multiples
//   • charges sociales (GR2100) : traité dans une étape dédiée ci-dessous
// ============================================================
let
    Source   = Exclusions,                 // = requête "Exclusions"
    Comptes  = Excel.CurrentWorkbook(){[Name="tblComptes"]}[Content],
    ETP      = Excel.CurrentWorkbook(){[Name="tblETP"]}[Content],
    Zone     = Excel.CurrentWorkbook(){[Name="tblZone"]}[Content],
    Taux     = Excel.CurrentWorkbook(){[Name="tblTaux"]}[Content],

    // 1) rattache la zone France/Pays via le RU
    jZone = Table.NestedJoin(Source, {"D_RU"}, Zone, {"RU"}, "z", JoinKind.LeftOuter),
    wZone = Table.AddColumn(jZone, "Zone", each try [z]{0}[#"Zone (France/Pays)"] otherwise "France"),

    // 2) sépare financier (F99) et ETP (Q99)
    Fin = Table.SelectRows(wZone, each [D_FL] = "F99"),
    Etp = Table.SelectRows(wZone, each [D_FL] = "Q99"),

    // 3a) FINANCIER : jointure sur le compte, choix IND selon la zone
    jFin = Table.NestedJoin(Fin, {"D_AC"}, Comptes, {"Code"}, "c", JoinKind.LeftOuter),
    // (une ligne compte -> potentiellement plusieurs IND dans tblComptes : Expand)
    xFin = Table.ExpandTableColumn(jFin, "c", {"IND_France","IND_Pays"}, {"IND_France","IND_Pays"}),
    wFin = Table.AddColumn(xFin, "Indicator", each
        if [Zone] = "Pays" and [IND_Pays] <> null and [IND_Pays] <> "" then [IND_Pays]
        else [IND_France]),

    // 3b) ETP : jointure sur tblETP
    jEtp = Table.NestedJoin(Etp, {"D_AC"}, ETP, {"Code_Magnitude"}, "e", JoinKind.LeftOuter),
    wEtp = Table.AddColumn(jEtp, "Indicator", each try [e]{0}[IND_Tagetik] otherwise null),

    // 4) recombine
    colonnes = {"D_DP","D_OA","D_FA","D_RU","D_AC","D_FL","D_CU","P_AMOUNT","Zone","Indicator"},
    Fusion = Table.Combine({
        Table.SelectColumns(wFin, colonnes),
        Table.SelectColumns(wEtp, colonnes)})
in
    Fusion
// ------------------------------------------------------------
// NB charges sociales (GR2100) : à insérer ici en amont —
//   base = GR2100 - GR2101 - GR2102 (par RU/OA/FA/période),
//   salaire_hors = base / (1 + taux_RU)  -> IND_00_060020
//   charges      = base - salaire_hors   -> IND_00_060074
//   GR2101 / GR2102 -> leurs IND dédiés (sans charges).
// (Étape détaillée à brancher selon la structure réelle des lignes.)
