let
    Src = Exclusions,
    cols = Table.ColumnNames(Src),
    Taux = Excel.CurrentWorkbook(){[Name="tblTaux"]}[Content],
    tauxOf = (ru) =>
        let v = try Table.SelectRows(Taux, each [RU] = ru){0}[Taux] otherwise null
        in if v = null then 0 else Number.From(v),
    Autres = Table.SelectRows(Src, each [D_AC] <> "GR2100"),
    Ded = Table.SelectRows(Src, each [D_AC] = "GR2101" or [D_AC] = "GR2102"),
    DedGrp = Table.Group(Ded, {"D_RU","D_OA","D_FA","D_DP"}, {{"ded", each List.Sum([P_AMOUNT]), type number}}),
    Sal = Table.SelectRows(Src, each [D_AC] = "GR2100"),
    jSal = Table.NestedJoin(Sal, {"D_RU","D_OA","D_FA","D_DP"}, DedGrp, {"D_RU","D_OA","D_FA","D_DP"}, "d", JoinKind.LeftOuter),
    wCalc = Table.AddColumn(jSal, "base", each [P_AMOUNT] - (try [d]{0}[ded] otherwise 0), type number),
    wCalc2 = Table.AddColumn(wCalc, "hors", each [base] / (1 + tauxOf([D_RU])), type number),
    wCalc3 = Table.AddColumn(wCalc2, "chg", each [base] - ([base] / (1 + tauxOf([D_RU]))), type number),
    recs = Table.ToRecords(wCalc3),
    expanded = List.Combine(List.Transform(recs, (r) =>
        {
            Record.SelectFields(r, cols) & [P_AMOUNT = r[hors]],
            Record.SelectFields(r, cols) & [P_AMOUNT = r[chg], D_AC = "GR21CS"]
        })),
    HorsChg = Table.FromRecords(expanded, Value.Type(Src)),
    Result = Table.Combine({Autres, HorsChg})
in
    Result
