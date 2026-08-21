// ============================================================
//  Constantes — déduit Scenario & Period de D_DP, ajoute les
//  dimensions fixes, et convertit en EUR (colonne de contrôle).
//  D_DP ex. "2026.03"  ->  Scenario "2026AC" , Period "03"
// ============================================================
let
    Source = MapDimensions,               // = requête "MapDimensions"
    Rate   = Excel.CurrentWorkbook(){[Name="tblRate"]}[Content],

    // Scenario & Period depuis D_DP
    wScen = Table.AddColumn(Source, "Scenario", each Text.BeforeDelimiter([D_DP], ".") & "AC", type text),
    wPer  = Table.AddColumn(wScen, "Period",  each Text.AfterDelimiter([D_DP], "."), type text),

    // constantes
    wVis  = Table.AddColumn(wPer,  "Vision",       each "VIS_00_000001", type text),
    wOri  = Table.AddColumn(wVis,  "Origin",       each "QDL", type text),
    wCpt  = Table.AddColumn(wOri,  "Counterparty", each "NA", type text),
    wPrd  = Table.AddColumn(wCpt,  "Product",      each "NA", type text),

    // devise & montant
    wCur  = Table.RenameColumns(wPrd, {{"D_CU","EntityCurrency"}}),

    // équivalent EUR (contrôle) via tblRate
    jR = Table.NestedJoin(wCur, {"EntityCurrency"}, Rate, {"Devise"}, "r", JoinKind.LeftOuter),
    wEUR = Table.AddColumn(jR, "Montant_EUR", each
        [P_AMOUNT] * (try Number.From([r]{0}[#"Taux vers EUR"]) otherwise 1), type number)
in
    wEUR
