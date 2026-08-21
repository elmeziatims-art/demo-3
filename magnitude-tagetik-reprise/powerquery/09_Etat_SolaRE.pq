let
    Src = Constantes,
    Grp = Table.Group(Src, {"Scenario","Period","PMA","Entity"},
        {{"Montant_EUR", each List.Sum([Montant_EUR]), type number},
         {"Montant_devise", each List.Sum([P_AMOUNT]), type number}}),
    Tri = Table.Sort(Grp, {{"PMA", Order.Ascending}, {"Entity", Order.Ascending}})
in
    Tri
