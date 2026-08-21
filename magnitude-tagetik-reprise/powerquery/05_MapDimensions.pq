// ============================================================
//  MapDimensions — (RU × OA × FA) -> Entité (+ PMA, CC)
//   • règle « la plus spécifique gagne » (case vide = joker)
//   • défaut = 1re EJ du RU (lignes génériques de tblDim)
//   • PMA vient de tblOAPMA (REF OA->PMA) ; CC de tblDim si renseigné
//   • Priorité départage les égalités
// ============================================================
let
    Source = MapComptes,                  // = requête "MapComptes"
    Dim    = Excel.CurrentWorkbook(){[Name="tblDim"]}[Content],
    DimOK  = Table.SelectRows(Dim, each [ENTITE] <> null and [ENTITE] <> ""),

    // score de spécificité d'une règle vis-à-vis d'une ligne
    matchScore = (row, r) =>
        let
            okRU = (r[RU] = null or r[RU] = "" or row[D_RU] = r[RU]),
            okOA = (r[OA] = null or r[OA] = "" or row[D_OA] = r[OA]),
            okFA = (r[FA] = null or r[FA] = "" or row[D_FA] = r[FA])
        in
            if okRU and okOA and okFA then
                (if r[RU]="" or r[RU]=null then 0 else 4)
              + (if r[OA]="" or r[OA]=null then 0 else 2)
              + (if r[FA]="" or r[FA]=null then 0 else 1)
              + (try Number.From(r[Priorite]) otherwise 0)*100
            else -1,

    regles = Table.Buffer(DimOK),

    choisit = (row) =>
        let
            scores = List.Transform(Table.ToRecords(regles), each [r=_, s=matchScore(row,_)]),
            valides = List.Select(scores, each [s] >= 0),
            meilleur = if List.Count(valides)=0 then null
                       else List.Last(List.Sort(valides, each [s]))[r]
        in meilleur,

    avecEnt = Table.AddColumn(Source, "Entity", each
        let m = choisit(_) in if m = null then null else m[ENTITE]),
    avecCC  = Table.AddColumn(avecEnt, "CostCenter", each
        let m = choisit(_) in if m = null then null else m[Cost_Center]),

    // PMA depuis le référentiel OA -> PMA
    OAPMA = Excel.CurrentWorkbook(){[Name="tblOAPMA"]}[Content],
    jPMA = Table.NestedJoin(avecCC, {"D_OA"}, OAPMA, {"Code MAGNITUDE (OA)"}, "p", JoinKind.LeftOuter),
    avecPMA = Table.AddColumn(jPMA, "PMA", each try [p]{0}[#"Code SOLARE (PMA)"] otherwise "NA")
in
    avecPMA
