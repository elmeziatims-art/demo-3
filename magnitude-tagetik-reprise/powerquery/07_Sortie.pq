let
    Src = Constantes,
    Sel = Table.SelectColumns(Src, {
        "Scenario","Period","Entity","Indicator","Counterparty","PMA","Product",
        "CostCenter","EntityCurrency","P_AMOUNT"}),
    Ren = Table.RenameColumns(Sel, {
        {"CostCenter","Cost Center"},
        {"EntityCurrency","Entity currency"},
        {"P_AMOUNT","Entity currency amount"}})
in
    Ren
