let
    Src = Constantes,
    Sel = Table.SelectColumns(Src, {
        "Scenario","Period","Entity","Indicator","Counterparty","PMA","Product",
        "CostCenter","EntityCurrency","P_AMOUNT"}),
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
