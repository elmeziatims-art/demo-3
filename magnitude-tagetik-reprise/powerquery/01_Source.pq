// ============================================================
//  Source — lit l'extraction Magnitude collée dans ① ENTREE
//  et ne garde que les 8 colonnes utiles.
// ============================================================
let
    Source = Excel.CurrentWorkbook(){[Name="tblEntree"]}[Content],

    // Sécurise les types texte des dimensions + numérique du montant
    Typ = Table.TransformColumnTypes(Source,{
        {"D_DP", type text}, {"D_OA", type text}, {"D_FA", type text},
        {"D_RU", type text}, {"D_AC", type text}, {"D_FL", type text},
        {"D_CU", type text}, {"P_AMOUNT", type number}
    }),

    // Ne garde que les colonnes nécessaires au moteur
    Garde = Table.SelectColumns(Typ,
        {"D_DP","D_OA","D_FA","D_RU","D_AC","D_FL","D_CU","P_AMOUNT"})
in
    Garde
