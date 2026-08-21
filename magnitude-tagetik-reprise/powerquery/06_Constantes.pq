// ============================================================
//  Constantes — déduit Scenario & Period de D_DP, ajoute les
//  dimensions fixes, et convertit en EUR (contrôle) via tblRate
//  joint sur Scénario + Période + Devise.
//  D_DP "2026.03" -> Scenario "2026AC", Period "03"
// ============================================================
let
    Source = MapDimensions,                        // requête "MapDimensions"
    Rate   = Excel.CurrentWorkbook(){[Name="tblRate"]}[Content],

    wScen = Table.AddColumn(Source, "Scenario", each Text.BeforeDelimiter([D_DP], ".") & "AC", type text),
    wPer  = Table.AddColumn(wScen, "Period",  each Text.PadStart(Text.AfterDelimiter([D_DP], "."), 2, "0"), type text),
    wVis  = Table.AddColumn(wPer,  "Vision",       each "VIS_00_000001", type text),
    wOri  = Table.AddColumn(wVis,  "Origin",       each "QDL", type text),
    wCpt  = Table.AddColumn(wOri,  "Counterparty", each "NA", type text),
    wPrd  = Table.AddColumn(wCpt,  "Product",      each "NA", type text),
    wCur  = Table.RenameColumns(wPrd, {{"D_CU","EntityCurrency"}}),

    // équivalent EUR via tblRate (Scenario + Period + Devise)
    jR = Table.NestedJoin(wCur, {"Scenario","Period","EntityCurrency"},
                           Rate, {"Scenario","Period","Devise"}, "r", JoinKind.LeftOuter),
    wEUR = Table.AddColumn(jR, "Montant_EUR", each
        [P_AMOUNT] * (try Number.From([r]{0}[#"Taux vers EUR"]) otherwise 1), type number),
    Clean = Table.RemoveColumns(wEUR, {"r"})
in
    Clean
