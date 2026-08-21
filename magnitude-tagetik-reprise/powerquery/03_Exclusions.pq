// ============================================================
//  Exclusions — retire un périmètre estimé AVANT le mapping.
//  Table tblExcl : RU | OA | FA | Mode(Montant|%) | Valeur | Commentaire
//  Case vide = toutes les valeurs. Application au prorata,
//  sur F99 ET Q99 (montant P_AMOUNT).
//  Résolution : on applique successivement chaque règle d'exclusion.
// ============================================================
let
    Source = Filtre,                       // = requête "Filtre"
    Regles = Excel.CurrentWorkbook(){[Name="tblExcl"]}[Content],
    ReglesOK = Table.SelectRows(Regles, each [Valeur] <> null and [Valeur] <> ""),

    // Applique une règle r à la table t
    Applique = (t as table, r as record) as table =>
        let
            match = (row) =>
                (r[RU]  = null or r[RU]  = "" or row[D_RU] = r[RU]) and
                (r[OA]  = null or r[OA]  = "" or row[D_OA] = r[OA]) and
                (r[FA]  = null or r[FA]  = "" or row[D_FA] = r[FA]),
            concernees = Table.SelectRows(t, each match(_)),
            total = List.Sum(concernees[P_AMOUNT]),
            // facteur de réduction : % -> (1 - v/100) ; Montant -> prorata (1 - montant/total)
            v = Number.From(r[Valeur]),
            facteur =
                if Text.Contains(r[#"Mode (Montant|%)"], "%") then 1 - v/100
                else if total = 0 then 1 else 1 - v/total,
            reduit = Table.TransformColumns(t, {"P_AMOUNT", each _}),  // placeholder
            // applique le facteur uniquement aux lignes concernées
            out = Table.FromRecords(
                List.Transform(Table.ToRecords(t), (row) =>
                    if match(row) then Record.TransformFields(row, {"P_AMOUNT", each _ * facteur}) else row))
        in
            out,

    Resultat = List.Accumulate(Table.ToRecords(ReglesOK), Source, (etat, r) => Applique(etat, r))
in
    Resultat
