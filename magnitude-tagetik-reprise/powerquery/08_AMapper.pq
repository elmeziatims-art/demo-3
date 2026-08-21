let
    Src = MapDimensions,
    manque = (t as table, cond as function, keycols as list, typ as text, ou as text) =>
        let
            bad = Table.SelectRows(t, cond),
            grp = Table.Group(bad, keycols, {{"nb", each Table.RowCount(_), Int64.Type}}),
            code = Table.AddColumn(grp, "Code manquant", each Text.Combine(List.Transform(keycols, (k)=> Text.From(Record.Field(_, k))), " × ")),
            t2 = Table.AddColumn(code, "Type", each typ),
            t3 = Table.AddColumn(t2, "Où mapper ?", each ou)
        in Table.SelectColumns(t3, {"Type","Code manquant","nb","Où mapper ?"}),
    ent = manque(Src, each [Entity] = null or [Entity] = "", {"D_RU","D_OA"}, "Entité manquante", "MAP - Dimensions (+ RU vers EJ)"),
    cc  = manque(Src, each [CostCenter] = null or [CostCenter] = "", {"D_FA"}, "Cost Center manquant", "MAP - FA vers CC (+ Dimensions)"),
    ind = manque(Src, each [Indicator] = null or [Indicator] = "", {"D_AC"}, "Indicateur manquant", "MAP - Comptes P&L / MAP - ETP"),
    Tout = Table.Combine({ent, cc, ind}),
    Final = Table.Sort(Tout, {{"nb", Order.Descending}})
in
    Final
