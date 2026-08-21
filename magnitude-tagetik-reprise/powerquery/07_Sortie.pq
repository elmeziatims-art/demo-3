let
    Src = Constantes,
    Comptes = Excel.CurrentWorkbook(){[Name="tblComptes"]}[Content],
    jB = Table.NestedJoin(Src, {"D_AC"}, Comptes, {"Code"}, "c", JoinKind.LeftOuter),
    wBloc = Table.AddColumn(jB, "Bloc P&L", each try [c]{0}[Bloc_PnL] otherwise null, type text),
    wPNB  = Table.AddColumn(wBloc, "PNB+MEE", each try (if [c]{0}[blk_PNB_MEE] = "x" then "X" else "") otherwise "", type text),
    wOPEX = Table.AddColumn(wPNB,  "OPEX",    each try (if [c]{0}[blk_OPEX] = "x" then "X" else "") otherwise "", type text),
    wPT   = Table.AddColumn(wOPEX, "Pré-tax", each try (if [c]{0}[blk_PreTax] = "x" then "X" else "") otherwise "", type text),
    Sel = Table.SelectColumns(wPT, {
        "Scenario","Period","Entity","Indicator","Counterparty","PMA","Product",
        "CostCenter","EntityCurrency","P_AMOUNT","Bloc P&L","PNB+MEE","OPEX","Pré-tax"}),
    Stat = Table.AddColumn(Sel, "Statut", each
        if [Entity] = null or [Entity] = ""
        or [Indicator] = null or [Indicator] = ""
        or [CostCenter] = null or [CostCenter] = ""
        then "À MAPPER" else "OK", type text),
    Ren = Table.RenameColumns(Stat, {
        {"CostCenter","Cost Center"},
        {"EntityCurrency","Entity currency"},
        {"P_AMOUNT","Entity currency amount"}})
in
    Ren
