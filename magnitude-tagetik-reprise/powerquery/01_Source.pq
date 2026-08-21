let
    Source = Excel.CurrentWorkbook(){[Name="tblEntree"]}[Content],
    Typ = Table.TransformColumnTypes(Source,{
        {"D_DP", type text}, {"D_OA", type text}, {"D_FA", type text},
        {"D_RU", type text}, {"D_AC", type text}, {"D_FL", type text},
        {"D_CU", type text}, {"P_AMOUNT", type number}}),
    Garde = Table.SelectColumns(Typ,
        {"D_DP","D_OA","D_FA","D_RU","D_AC","D_FL","D_CU","P_AMOUNT"})
in
    Garde
