// ============================================================
//  AMapper — détection automatique (à charger dans « À MAPPER (auto) »)
//  Liste les RU / OA / FA présents dans les données filtrées mais
//  ABSENTS des tables de mapping. Se met à jour à chaque Actualisation.
//  (100% Power Query, aucun VBA.)
// ============================================================
let
    Data = Filtre,                               // requête "Filtre"
    RUEJ = Excel.CurrentWorkbook(){[Name="tblRUEJ"]}[Content],
    OAP  = Excel.CurrentWorkbook(){[Name="tblOAPMA"]}[Content],
    FACC = Excel.CurrentWorkbook(){[Name="tblFACC"]}[Content],

    ruMap = List.Buffer(List.Distinct(RUEJ[RU])),
    oaMap = List.Buffer(List.Distinct(OAP[#"Code MAGNITUDE (OA)"])),
    faMap = List.Buffer(List.Distinct(FACC[FA])),

    grp = (t as table, srcCol as text, mapped as list, typ as text) =>
        let d = Table.Group(t, {srcCol}, {{"nb", each Table.RowCount(_), Int64.Type}}),
            manquants = Table.SelectRows(d, each not List.Contains(mapped, Record.Field(_, srcCol))),
            ren = Table.RenameColumns(manquants, {{srcCol,"Code manquant"}}),
            avecType = Table.AddColumn(ren, "Type", each typ)
        in avecType,

    ru = grp(Data, "D_RU", ruMap, "RU"),
    oa = grp(Data, "D_OA", oaMap, "OA"),
    fa = grp(Data, "D_FA", faMap, "FA"),
    Tout = Table.Combine({ru, oa, fa}),
    Final = Table.SelectColumns(
        Table.AddColumn(Tout, "Action", each "À mapper — ajoutez ce code dans l'onglet MAP correspondant"),
        {"Type","Code manquant","nb","Action"})
in
    Final
