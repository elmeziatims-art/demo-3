// ============================================================
//  MapDimensions — (RU × OA × FA) -> Entité, PMA, Cost Center
//  MAP - Dimensions (tblDim) a TOUJOURS le dernier mot :
//   • Entité : règle la plus spécifique de tblDim (défaut = EJ du RU) ;
//   • PMA   : si renseigné dans tblDim -> forcé ; sinon table OA→PMA (tblOAPMA) ;
//   • CC    : si renseigné dans tblDim -> forcé ; sinon CC "Défaut" du FA (tblFACC).
//  Case vide dans tblDim = joker ; Priorité départage.
// ============================================================
let
    Source = MapComptes,                              // requête "MapComptes"
    Dim    = Table.Buffer(Excel.CurrentWorkbook(){[Name="tblDim"]}[Content]),
    OAPMA  = Excel.CurrentWorkbook(){[Name="tblOAPMA"]}[Content],
    FACC   = Excel.CurrentWorkbook(){[Name="tblFACC"]}[Content],
    FAdef  = Table.SelectRows(FACC, each Text.Upper(Text.From([Défaut])) = "X"),

    // score de spécificité règle/ligne (-1 = ne matche pas)
    score = (row, r) =>
        let okRU=(r[RU]=null or r[RU]="" or row[D_RU]=r[RU]),
            okOA=(r[OA]=null or r[OA]="" or row[D_OA]=r[OA]),
            okFA=(r[FA]=null or r[FA]="" or row[D_FA]=r[FA])
        in if okRU and okOA and okFA then
              (if r[RU]="" or r[RU]=null then 0 else 4)
            + (if r[OA]="" or r[OA]=null then 0 else 2)
            + (if r[FA]="" or r[FA]=null then 0 else 1)
            + (try Number.From(r[Priorite]) otherwise 0)*100
           else -1,
    best = (row) =>
        let sc = List.Transform(Table.ToRecords(Dim), each [r=_, s=score(row,_)]),
            ok = List.Select(sc, each [s] >= 0)
        in if List.Count(ok)=0 then null else List.Last(List.Sort(ok, each [s]))[r],

    withEnt = Table.AddColumn(Source, "Entity", each let m=best(_) in if m=null then null else m[ENTITE]),
    withCC  = Table.AddColumn(withEnt, "CostCenter", each
                let m=best(_),
                    forced = if m<>null and m[Cost_Center]<>null and m[Cost_Center]<>"" then m[Cost_Center] else null,
                    parFA  = try Table.SelectRows(FAdef, (x)=> x[FA]=[D_FA]){0}[CC] otherwise null
                in if forced<>null then forced else parFA),
    // PMA : override tblDim sinon OA→PMA
    jP = Table.NestedJoin(withCC, {"D_OA"}, OAPMA, {"Code MAGNITUDE (OA)"}, "p", JoinKind.LeftOuter),
    withPMA = Table.AddColumn(jP, "PMA", each
                let m=best(_),
                    forced = if m<>null and m[PMA]<>null and m[PMA]<>"" then m[PMA] else null,
                    parOA  = try [p]{0}[#"Code SOLARE (PMA)"] otherwise "NA"
                in if forced<>null then forced else parOA),
    Clean = Table.RemoveColumns(withPMA, {"p"})
in
    Clean
