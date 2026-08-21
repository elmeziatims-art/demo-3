let
    Data = Filtre,
    RUEJ = Excel.CurrentWorkbook(){[Name="tblRUEJ"]}[Content],
    OAP = Excel.CurrentWorkbook(){[Name="tblOAPMA"]}[Content],
    FACC = Excel.CurrentWorkbook(){[Name="tblFACC"]}[Content],
    ruMap = List.Buffer(List.Distinct(RUEJ[RU])),
    oaMap = List.Buffer(List.Distinct(OAP[#"Code MAGNITUDE (OA)"])),
    faMap = List.Buffer(List.Distinct(FACC[FA])),
    grp = (t as table, srcCol as text, mapped as list, typ as text, ou as text) =>
        let
            d = Table.Group(t, {srcCol}, {{"nb", each Table.RowCount(_), Int64.Type}}),
            manquants = Table.SelectRows(d, each not List.Contains(mapped, Record.Field(_, srcCol))),
            ren = Table.RenameColumns(manquants, {{srcCol, "Code manquant"}}),
            t2 = Table.AddColumn(ren, "Type", each typ),
            t3 = Table.AddColumn(t2, "Où mapper ?", each ou)
        in t3,
    ru = grp(Data, "D_RU", ruMap, "RU", "MAP - RU vers EJ  (+ MAP - Dimensions)"),
    oa = grp(Data, "D_OA", oaMap, "OA", "MAP - OA vers PMA"),
    fa = grp(Data, "D_FA", faMap, "FA", "MAP - FA vers CC"),
    Tout = Table.Combine({ru, oa, fa}),
    Sorted = Table.Sort(Tout, {{"nb", Order.Descending}}),
    Final = Table.SelectColumns(Sorted, {"Type","Code manquant","nb","Où mapper ?"})
in
    Final
