let
    Src = Exclusions,
    Comptes = Excel.CurrentWorkbook(){[Name="tblComptes"]}[Content],
    ETP = Excel.CurrentWorkbook(){[Name="tblETP"]}[Content],
    Zone = Excel.CurrentWorkbook(){[Name="tblZone"]}[Content],
    jZone = Table.NestedJoin(Src, {"D_RU"}, Zone, {"RU"}, "z", JoinKind.LeftOuter),
    wZone = Table.AddColumn(jZone, "Zone", each try [z]{0}[Zone] otherwise "France"),
    Fin = Table.SelectRows(wZone, each [D_FL] = "F99"),
    Etp = Table.SelectRows(wZone, each [D_FL] = "Q99"),
    jFin = Table.NestedJoin(Fin, {"D_AC"}, Comptes, {"Code"}, "c", JoinKind.LeftOuter),
    xFin = Table.ExpandTableColumn(jFin, "c", {"IND_France","IND_Pays"}, {"IND_France","IND_Pays"}),
    wFin = Table.AddColumn(xFin, "Indicator", each
        if [Zone] = "Pays" and [IND_Pays] <> null and [IND_Pays] <> "" then [IND_Pays] else [IND_France]),
    jEtp = Table.NestedJoin(Etp, {"D_AC"}, ETP, {"Code_Magnitude"}, "e", JoinKind.LeftOuter),
    wEtp = Table.AddColumn(jEtp, "Indicator", each try [e]{0}[IND_Tagetik] otherwise null),
    cols = {"D_DP","D_OA","D_FA","D_RU","D_AC","D_FL","D_CU","P_AMOUNT","Zone","Indicator"},
    Fusion = Table.Combine({Table.SelectColumns(wFin, cols), Table.SelectColumns(wEtp, cols)})
in
    Fusion
