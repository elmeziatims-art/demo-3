let
    Source = Excel.CurrentWorkbook(){[Name="tblEntree"]}[Content],
    Txt = Table.TransformColumnTypes(Source,{
        {"D_DP", type text},{"D_OA", type text},{"D_FA", type text},
        {"D_RU", type text},{"D_AC", type text},{"D_FL", type text},{"D_CU", type text}}),
    Amt = Table.AddColumn(Txt, "MNT", each
        let v = [P_AMOUNT] in
        if v = null then null
        else if Value.Is(v, type number) then v
        else try Number.FromText(Text.Replace(Text.From(v), ",", "."), "en-US")
             otherwise try Number.From(v) otherwise null, type number),
    Garde = Table.SelectColumns(Amt, {"D_DP","D_OA","D_FA","D_RU","D_AC","D_FL","D_CU","MNT"}),
    Ren = Table.RenameColumns(Garde, {{"MNT","P_AMOUNT"}})
in
    Ren
