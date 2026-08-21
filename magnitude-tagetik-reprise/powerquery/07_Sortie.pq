// ============================================================
//  Sortie — table de fait Tagetik (format d'import à plat).
//  Charge cette requête dans l'onglet ⑨ SORTIE.
//  Montant = devise entité (P_AMOUNT). Category laissée vide (hors périmètre).
// ============================================================
let
    Source = Constantes,                  // = requête "Constantes"

    Sel = Table.SelectColumns(Source, {
        "Scenario","Period","Entity","Indicator","Counterparty","PMA","Product",
        "Vision","CostCenter","EntityCurrency","P_AMOUNT","Origin"}),

    Ren = Table.RenameColumns(Sel, {
        {"CostCenter","Cost Center"},
        {"EntityCurrency","Entity currency"},
        {"P_AMOUNT","Entity currency amount"}}),

    // Ajoute les colonnes de description vides + Category (hors périmètre)
    Final = Table.AddColumn(Ren, "Category", each "", type text)
in
    Final
