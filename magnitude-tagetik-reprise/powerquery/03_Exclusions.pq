let
    Src = Filtre,
    Regles = Excel.CurrentWorkbook(){[Name="tblExcl"]}[Content],
    ReglesOK = Table.SelectRows(Regles, each [Valeur] <> null and [Valeur] <> ""),
    Applique = (t as table, r as record) as table =>
        let
            match = (row) =>
                (r[RU] = null or r[RU] = "" or row[D_RU] = r[RU]) and
                (r[OA] = null or r[OA] = "" or row[D_OA] = r[OA]) and
                (r[FA] = null or r[FA] = "" or row[D_FA] = r[FA]),
            total = List.Sum(Table.SelectRows(t, each match(_))[P_AMOUNT]),
            v = Number.From(r[Valeur]),
            facteur = if Text.Contains(r[Mode], "%") then 1 - v/100
                      else if total = 0 then 1 else 1 - v/total,
            out = Table.FromRecords(List.Transform(Table.ToRecords(t), (row) =>
                if match(row) then Record.TransformFields(row, {"P_AMOUNT", each _ * facteur}) else row))
        in out,
    Resultat = List.Accumulate(Table.ToRecords(ReglesOK), Src, (etat, r) => Applique(etat, r))
in
    Resultat
