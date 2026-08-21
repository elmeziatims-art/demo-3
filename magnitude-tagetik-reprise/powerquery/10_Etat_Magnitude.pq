let
    Src = Constantes,
    Comptes = Excel.CurrentWorkbook(){[Name="tblComptes"]}[Content],
    jB = Table.NestedJoin(Src, {"D_AC"}, Comptes, {"Code"}, "c", JoinKind.LeftOuter),
    wB = Table.AddColumn(jB, "Bloc_PnL", each try [c]{0}[Bloc_PnL] otherwise "(non classé)"),
    Grp = Table.Group(wB, {"Scenario","Period","Bloc_PnL","D_AC"},
        {{"Montant_EUR", each List.Sum([Montant_EUR]), type number},
         {"Montant_devise", each List.Sum([P_AMOUNT]), type number}}),
    Tri = Table.Sort(Grp, {{"Bloc_PnL", Order.Ascending}, {"D_AC", Order.Ascending}})
in
    Tri
