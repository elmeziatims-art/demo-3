let
    Src = Source,
    fFlux = Table.SelectRows(Src, each ([D_FL] = "F99" or [D_FL] = "Q99")),
    fOA = Table.SelectRows(fFlux, each [D_OA] <> null and Text.Trim([D_OA]) <> ""),
    Comptes = Excel.CurrentWorkbook(){[Name="tblComptes"]}[Content],
    Fins = Table.SelectRows(Comptes, each [Type] = "Compte fin"),
    ListeFins = List.Buffer(List.Distinct(Fins[Code])),
    fFin = Table.SelectRows(fOA, each List.Contains(ListeFins, [D_AC]))
in
    fFin
