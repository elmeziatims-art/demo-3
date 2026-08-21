let
    Src = Constantes,
    Sel = Table.SelectColumns(Src, {
        "Scenario","Period","Entity","Indicator","Counterparty","PMA","Product",
        "Vision","CostCenter","EntityCurrency","P_AMOUNT","Origin"}),
    Ren = Table.RenameColumns(Sel, {
        {"CostCenter","Cost Center"},
        {"EntityCurrency","Entity currency"},
        {"P_AMOUNT","Entity currency amount"}}),
    Final = Table.AddColumn(Ren, "Category", each "", type text)
in
    Final
