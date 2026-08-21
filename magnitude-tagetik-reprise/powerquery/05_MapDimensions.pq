let
    Src = MapComptes,
    Dim = Table.Buffer(Excel.CurrentWorkbook(){[Name="tblDim"]}[Content]),
    OAPMA = Excel.CurrentWorkbook(){[Name="tblOAPMA"]}[Content],
    FACC = Excel.CurrentWorkbook(){[Name="tblFACC"]}[Content],
    FAdef = Table.SelectRows(FACC, each Text.Upper(Text.From([Défaut])) = "X"),
    ccOfFA = (fa) => try Table.SelectRows(FAdef, each [FA] = fa){0}[CC] otherwise null,
    pmaOfOA = (oa) => try Table.SelectRows(OAPMA, each [#"Code MAGNITUDE (OA)"] = oa){0}[#"Code SOLARE (PMA)"] otherwise "NA",
    score = (dr, r) =>
        let
            okRU = (r[RU] = null or r[RU] = "" or dr[D_RU] = r[RU]),
            okOA = (r[OA] = null or r[OA] = "" or dr[D_OA] = r[OA]),
            okFA = (r[FA] = null or r[FA] = "" or dr[D_FA] = r[FA])
        in
            if okRU and okOA and okFA
            then (if r[RU] = "" or r[RU] = null then 0 else 4)
               + (if r[OA] = "" or r[OA] = null then 0 else 2)
               + (if r[FA] = "" or r[FA] = null then 0 else 1)
            else -1,
    best = (dr) =>
        let
            scored = List.Transform(Table.ToRecords(Dim), (r) => [rec = r, s = score(dr, r)]),
            ok = List.Select(scored, each [s] >= 0)
        in
            if List.Count(ok) = 0 then null else List.Last(List.Sort(ok, each [s]))[rec],
    withEnt = Table.AddColumn(Src, "Entity", (dr) => let m = best(dr) in if m = null then null else m[ENTITE]),
    withCC = Table.AddColumn(withEnt, "CostCenter", (dr) =>
        let m = best(dr),
            forced = if m <> null and m[Cost_Center] <> null and m[Cost_Center] <> "" then m[Cost_Center] else null
        in if forced <> null then forced else ccOfFA(dr[D_FA])),
    withPMA = Table.AddColumn(withCC, "PMA", (dr) =>
        let m = best(dr),
            forced = if m <> null and m[PMA] <> null and m[PMA] <> "" then m[PMA] else null
        in if forced <> null then forced else pmaOfOA(dr[D_OA]))
in
    withPMA
