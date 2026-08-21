// ============================================================
//  MapDimensions — (RU × OA × FA) -> Entité, PMA, Cost Center
//  MAP - Dimensions (tblDim) a TOUJOURS le dernier mot.
//   • table dédoublonnée : UNE seule ligne par croisement (RU,OA,FA)
//   • règle la plus SPÉCIFIQUE gagne (le moins de cases vides) ; case vide = joker
//   • Entité : prise de tblDim
//   • PMA/CC : forcés si renseignés dans tblDim ; sinon table OA→PMA / CC "Défaut" du FA
//  (Pas de colonne Priorité : inutile car un seul enregistrement par croisement.)
// ============================================================
let
    Source = MapComptes,
    Dim    = Table.Buffer(Excel.CurrentWorkbook(){[Name="tblDim"]}[Content]),
    OAPMA  = Excel.CurrentWorkbook(){[Name="tblOAPMA"]}[Content],
    FACC   = Excel.CurrentWorkbook(){[Name="tblFACC"]}[Content],
    FAdef  = Table.SelectRows(FACC, each Text.Upper(Text.From([Défaut])) = "X"),

    score = (row, r) =>
        let okRU=(r[RU]=null or r[RU]="" or row[D_RU]=r[RU]),
            okOA=(r[OA]=null or r[OA]="" or row[D_OA]=r[OA]),
            okFA=(r[FA]=null or r[FA]="" or row[D_FA]=r[FA])
        in if okRU and okOA and okFA then
              (if r[RU]="" or r[RU]=null then 0 else 4)
            + (if r[OA]="" or r[OA]=null then 0 else 2)
            + (if r[FA]="" or r[FA]=null then 0 else 1)
           else -1,
    best = (row) =>
        let sc=List.Transform(Table.ToRecords(Dim), each [r=_, s=score(row,_)]),
            ok=List.Select(sc, each [s]>=0)
        in if List.Count(ok)=0 then null else List.Last(List.Sort(ok, each [s]))[r],

    withEnt = Table.AddColumn(Source, "Entity", each let m=best(_) in if m=null then null else m[ENTITE]),
    withCC  = Table.AddColumn(withEnt, "CostCenter", each
                let m=best(_),
                    forced=if m<>null and m[Cost_Center]<>null and m[Cost_Center]<>"" then m[Cost_Center] else null,
                    parFA =try Table.SelectRows(FAdef,(x)=>x[FA]=[D_FA]){0}[CC] otherwise null
                in if forced<>null then forced else parFA),
    jP = Table.NestedJoin(withCC, {"D_OA"}, OAPMA, {"Code MAGNITUDE (OA)"}, "p", JoinKind.LeftOuter),
    withPMA = Table.AddColumn(jP, "PMA", each
                let m=best(_),
                    forced=if m<>null and m[PMA]<>null and m[PMA]<>"" then m[PMA] else null,
                    parOA =try [p]{0}[#"Code SOLARE (PMA)"] otherwise "NA"
                in if forced<>null then forced else parOA),
    Clean = Table.RemoveColumns(withPMA, {"p"})
in
    Clean
