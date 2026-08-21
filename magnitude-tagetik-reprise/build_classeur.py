#!/usr/bin/env python3
"""Construit le classeur Reprise Magnitude -> Tagetik (structure + donnees de mapping).
Le moteur Power Query (M) est fourni a part (dossier powerquery/) a coller dans Excel."""
import csv, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

DOCS = "docs"
def rd(name, header=True):
    with open(os.path.join(DOCS, name), encoding="utf-8") as f:
        r = list(csv.reader(f))
    return (r[0], r[1:]) if header else (None, r)

# ---- palette / styles
NAVY="17262A"; ACC="0E7C66"; ACCSOFT="E2EFE9"; AMBER="B26A13"; AMBERSOFT="F5E9D6"
VIOLET="6A4B8A"; VIOLETSOFT="EDE7F4"; BLUE="2A6F8E"; GREY="5C6F72"; LINE="C6D0CB"
WHITE="FFFFFF"; PAPER="F4F6F4"; REDTXT="C0392B"; REDFILL="FBE9E7"; YELLOW="FFF6D8"
F="Arial"
def font(sz=10,b=False,color=NAVY,it=False): return Font(name=F,size=sz,bold=b,color=color,italic=it)
def fill(c): return PatternFill("solid",fgColor=c)
thin=Side(style="thin",color=LINE)
border=Border(left=thin,right=thin,top=thin,bottom=thin)
center=Alignment(horizontal="center",vertical="center",wrap_text=True)
left=Alignment(horizontal="left",vertical="center",wrap_text=True)
top=Alignment(horizontal="left",vertical="top",wrap_text=True)

wb=Workbook(); wb.remove(wb.active)

def sheet(name,tab_color=None):
    ws=wb.create_sheet(name)
    if tab_color: ws.sheet_properties.tabColor=tab_color
    ws.sheet_view.showGridLines=False
    return ws

def title_block(ws,title,subtitle,color=ACC,badge=None):
    ws["A1"]=title; ws["A1"].font=font(15,True,color)
    ws["A2"]=subtitle; ws["A2"].font=font(9,False,GREY,it=True)
    if badge:
        ws["A3"]=badge; ws["A3"].font=font(8,True,color)
    ws.row_dimensions[1].height=22

def header_row(ws,headers,row,color=NAVY,txt=WHITE):
    for j,h in enumerate(headers,1):
        c=ws.cell(row=row,column=j,value=h)
        c.font=font(9,True,txt); c.fill=fill(color); c.alignment=center; c.border=border
    ws.freeze_panes=ws.cell(row=row+1,column=1)

def widths(ws,ws_widths):
    for col,w in ws_widths.items(): ws.column_dimensions[col].width=w

# =========================================================== 0. ACCUEIL
ws=sheet("0. Accueil",ACC)
ws["B2"]="Reprise Magnitude  →  Tagetik (SolaRE)"; ws["B2"].font=font(18,True,ACC)
ws["B3"]="Moteur Excel / Power Query — sans VBA"; ws["B3"].font=font(10,False,GREY,it=True)
steps=[("1","Coller","Collez l'extraction Magnitude (colonnes A→Y) dans l'onglet « ① ENTREE »."),
       ("2","Actualiser","Données ▸ Actualiser tout — le moteur recalcule tout."),
       ("3","Lire","Récupérez la « SORTIE — Table de fait » et les 2 états de restitution.")]
r=6
ws[f"B{r}"]="Mode d'emploi (3 étapes)"; ws[f"B{r}"].font=font(12,True,NAVY); r+=1
for n,t,d in steps:
    ws[f"B{r}"]=n; ws[f"B{r}"].font=font(14,True,WHITE); ws[f"B{r}"].fill=fill(ACC); ws[f"B{r}"].alignment=center
    ws[f"C{r}"]=t; ws[f"C{r}"].font=font(11,True,NAVY)
    ws[f"D{r}"]=d; ws[f"D{r}"].font=font(10,False,NAVY); ws[f"D{r}"].alignment=left
    ws.row_dimensions[r].height=28; r+=1
r+=1
ws[f"B{r}"]="Légende des onglets"; ws[f"B{r}"].font=font(12,True,NAVY); r+=1
leg=[("Saisie utilisateur",AMBER),("Éditable par les CDG",ACC),("Référentiel — lecture seule",VIOLET),("Automatique — Power Query",BLUE)]
for lab,c in leg:
    ws[f"B{r}"]="   "; ws[f"B{r}"].fill=fill(c)
    ws[f"C{r}"]=lab; ws[f"C{r}"].font=font(10,False,NAVY); r+=1
r+=1
ws[f"B{r}"]="Règle d'or : toutes les tables de mapping restent modifiables. Le rouge = « à valider / à saisir »."
ws[f"B{r}"].font=font(10,True,REDTXT)
widths(ws,{"A":2,"B":22,"C":26,"D":80})

# =========================================================== 1. ENTREE
ws=sheet("1. ① ENTREE - Magnitude",AMBER)
title_block(ws,"① ENTREE — Extraction Magnitude","Collez ici l'extraction brute. La vraie extraction s'arrête à P_COMMENT (colonne Y).","0"+AMBER if False else AMBER)
cols_src=["D_CA","D_DP","D_OA","D_FA","D_VI","D_TA","D_PE","D_RU","D_ORU","D_AC","D_FL","D_AU","D_T1","D_T2","D_CU","D_TO","D_GO","D_LE","D_NU","D_DEST","D_AREA","D_MU","D_PMU","P_AMOUNT","P_COMMENT"]
header_row(ws,cols_src,5)
ws["A4"]="▼ COLLEZ L'EXTRACTION À PARTIR DE LA LIGNE 6 (l'en-tête ligne 5 correspond aux colonnes source)"
ws["A4"].font=font(9,True,AMBER)
for j in range(1,len(cols_src)+1): ws.column_dimensions[get_column_letter(j)].width=12

# =========================================================== MAP - Comptes P&L
ws=sheet("MAP - Comptes P&L",ACC)
title_block(ws,"MAP — Comptes P&L (D_AC → Indicateur)","Éditable. Variante France (IND_France) / Pays (IND_Pays). Rouge = à valider/compléter.")
h,rows=rd("mapping_comptes_D_AC.csv")
header_row(ws,h,5)
rr=6
for row in rows:
    for j,val in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=val); c.font=font(9); c.alignment=left; c.border=border
    # rouge si compte fin sans IND_France
    code=row[0]; typ=row[3] if len(row)>3 else ""; indf=row[7] if len(row)>7 else ""
    if typ=="Compte fin" and not indf:
        for j in range(1,len(h)+1):
            ws.cell(row=rr,column=j).font=font(9,True,REDTXT); ws.cell(row=rr,column=j).fill=fill(REDFILL)
    rr+=1
widths(ws,{"A":10,"B":34,"C":16,"D":11,"E":10,"F":8,"G":9,"H":15,"I":30,"J":15,"K":30,"L":40})

# =========================================================== MAP - ETP
ws=sheet("MAP - ETP (Q99)",ACC)
title_block(ws,"MAP — ETP (flux Q99 → Indicateur headcount)","Éditable. GR050 reste à mapper (rouge).")
h,rows=rd("mapping_ETP_Q99.csv")
header_row(ws,h,5); rr=6
for row in rows:
    for j,val in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=val); c.font=font(9); c.alignment=left; c.border=border
    if not row[1]:
        for j in range(1,len(h)+1):
            ws.cell(row=rr,column=j).font=font(9,True,REDTXT); ws.cell(row=rr,column=j).fill=fill(REDFILL)
    rr+=1
widths(ws,{"A":16,"B":18,"C":30,"D":16,"E":40})

# =========================================================== MAP - Dimensions
ws=sheet("MAP - Dimensions",ACC)
title_block(ws,"MAP — Dimensions (RU × OA × FA → Entité · PMA · CC)","Règle générique (défaut = 1re EJ du RU) + exceptions. Case vide = toutes les valeurs. « Le plus spécifique gagne » + Priorité.")
dh=["RU","OA","FA","ENTITE","PMA","Cost_Center","Priorite","Source","Commentaire"]
header_row(ws,dh,5); rr=6
_,drows=rd("mapping_dim_defaut_RU_entite.csv")
# example row (input format) — first
ws.cell(row=rr,column=1,value="G-UK").font=font(9,it=True,color=GREY)
ws.cell(row=rr,column=2,value="OA060").font=font(9,it=True,color=GREY)
ws.cell(row=rr,column=4,value="EJ_45220").font=font(9,it=True,color=GREY)
ws.cell(row=rr,column=8,value="EXEMPLE exception").font=font(9,it=True,color=GREY)
ws.cell(row=rr,column=9,value="ligne d'exemple — supprimez-la").font=font(9,it=True,color=GREY)
for j in range(1,len(dh)+1): ws.cell(row=rr,column=j).border=border
rr+=1
for row in drows:
    ru,oa,fa,ent,lib,src=row
    vals=[ru,oa,fa,ent,"","", "", src, lib]
    for j,val in enumerate(vals,1):
        c=ws.cell(row=rr,column=j,value=val); c.font=font(9); c.alignment=left; c.border=border
        if j in (4,) and val: c.fill=fill(YELLOW)
    rr+=1
widths(ws,{"A":12,"B":10,"C":10,"D":14,"E":12,"F":14,"G":9,"H":16,"I":34})

# =========================================================== MAP - Exclusions
ws=sheet("MAP - Exclusions",ACC)
title_block(ws,"MAP — Exclusions / Ajustements","Retirer X% ou X€ d'un RU × OA × FA (case vide = tout). Appliqué AVANT le mapping, au prorata, sur F99 ET ETP.")
eh=["RU","OA","FA","Mode (Montant|%)","Valeur","Commentaire"]
header_row(ws,eh,5)
ex=[["G-REIMLUX","","","%","30","EXEMPLE — retire 30% de tout ce RU"],
    ["","OA060","","Montant","200000","EXEMPLE — 200k€ sur cet OA, tous RU/FA"]]
rr=6
for row in ex:
    for j,val in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=val); c.font=font(9,it=True,color=GREY); c.border=border; c.alignment=left
    rr+=1
dv=DataValidation(type="list",formula1='"Montant,%"',allow_blank=True); ws.add_data_validation(dv); dv.add(f"D6:D200")
widths(ws,{"A":14,"B":10,"C":10,"D":16,"E":12,"F":44})

# =========================================================== MAP - Taux charges
ws=sheet("MAP - Taux charges soc.",ACC)
title_block(ws,"MAP — Taux de charges sociales / RU","Un seul % par RU. Appliqué ligne par ligne : base = GR2100 − GR2101 − GR2102 ; salaire_hors = base/(1+taux) ; charges = base − salaire_hors.")
header_row(ws,["RU","Taux charges sociales","Commentaire"],5)
_,zrows=rd("ref_RU_pays_zone.csv"); rus=[r[0] for r in zrows]
rr=6
for ru in rus:
    ws.cell(row=rr,column=1,value=ru).font=font(9); ws.cell(row=rr,column=1).border=border
    tc=ws.cell(row=rr,column=2); tc.fill=fill(YELLOW); tc.border=border; tc.number_format="0.0%"
    ws.cell(row=rr,column=3).border=border
    rr+=1
widths(ws,{"A":14,"B":18,"C":40})

# =========================================================== MAP - Zone RU
ws=sheet("MAP - Zone RU",ACC)
title_block(ws,"MAP — Zone RU (France / Pays)","Détermine la variante d'indicateur (cas comptes France vs Pays). Dérivé de Mappable/Non mappable, éditable.")
h,rows=rd("ref_RU_pays_zone.csv"); header_row(ws,["RU","Pays","Zone (France/Pays)"],5); rr=6
dvz=DataValidation(type="list",formula1='"France,Pays"',allow_blank=True)
ws.add_data_validation(dvz)
for row in rows:
    for j,val in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=val); c.font=font(9); c.border=border; c.alignment=left
        if j==3: c.fill=fill(YELLOW)
    rr+=1
dvz.add(f"C6:C{rr}")
widths(ws,{"A":14,"B":20,"C":18})

# =========================================================== PARAM
ws=sheet("PARAM - Constantes",ACC)
title_block(ws,"PARAM — Constantes du run","Scenario & Period sont DÉDUITS de D_DP (2026.03 → 2026AC / 03). Ci-dessous les vraies constantes.")
header_row(ws,["Dimension","Valeur","Note"],5)
params=[("Scenario","= déduit de D_DP","année + « AC » (Actuals)"),
        ("Period","= déduit de D_DP","mois (2 chiffres)"),
        ("Vision","VIS_00_000001","JV 100%"),
        ("Origin","QDL",""),
        ("Counterparty","NA",""),
        ("Product","NA",""),
        ("Category","(hors périmètre)","traçage Tagetik uniquement")]
rr=6
for d,v,n in params:
    ws.cell(row=rr,column=1,value=d).font=font(9,True); ws.cell(row=rr,column=1).border=border
    cv=ws.cell(row=rr,column=2,value=v); cv.border=border
    if v.startswith("="): cv.font=font(9,it=True,color=GREY)
    else: cv.font=font(9,color="0000FF"); cv.fill=fill(YELLOW)
    ws.cell(row=rr,column=3,value=n).font=font(9,color=GREY); ws.cell(row=rr,column=3).border=border
    rr+=1
widths(ws,{"A":16,"B":22,"C":34})

# =========================================================== REF sheets (lecture seule)
def ref_sheet(name,csvfile,headers=None,widths_map=None,has_header=True):
    ws=sheet(name,VIOLET)
    title_block(ws,name.replace("REF - ","REF — "),"Référentiel — lecture seule (aide au choix / listes déroulantes).",VIOLET)
    h,rows=rd(csvfile,header=has_header)
    hh=h if has_header else headers
    header_row(ws,hh,5,color=VIOLET); rr=6
    for row in rows:
        for j,val in enumerate(row,1):
            c=ws.cell(row=rr,column=j,value=val); c.font=font(9); c.border=border; c.alignment=left
        rr+=1
    if widths_map: widths(ws,widths_map)
    return ws

ref_sheet("REF - Indicateurs","tagetik_indicateurs.csv",widths_map={"A":8,"B":26,"C":7,"D":14,"E":18,"F":52})
ref_sheet("REF - Hier Entites","hierarchie_entites.csv",widths_map={"A":14,"B":40,"C":14,"D":16})
ref_sheet("REF - Hier PMA","hierarchie_pma.csv",widths_map={"A":14,"B":44,"C":14,"D":16})
ref_sheet("REF - FA vers CC","mapping_FA_to_CC.csv",widths_map={"A":10,"B":34,"C":12,"D":40})
ref_sheet("REF - OA vers PMA","mapping_OA_business_lines.csv",headers=["Code MAGNITUDE (OA)","Code SOLARE (PMA)","Description"],has_header=False,widths_map={"A":20,"B":18,"C":46})
ref_sheet("REF - RU vers EJ","mapping_RU_entities.csv",headers=["RU","Lib Magnitude","Devise","EJ (SOLARE)","Lib SOLARE"],has_header=False,widths_map={"A":14,"B":26,"C":8,"D":12,"E":40})
ref_sheet("REF - RUxOA Mappable","mapping_RUxOA_mappable.csv",widths_map={"A":12,"B":14,"C":8,"D":10,"E":36,"F":22,"G":22,"H":40,"I":10})
ref_sheet("REF - RUxOA NonMappable","mapping_RUxOA_non_mappable.csv",widths_map={"A":12,"B":12,"C":8,"D":30,"E":26,"F":26,"G":50,"H":50,"I":8})

# =========================================================== rate
ws=sheet("rate",ACC)
title_block(ws,"rate — Taux de change (→ EUR)","Éditable. Sert au contrôle en euro et aux 2 états de restitution.")
header_row(ws,["Devise","Taux vers EUR","Période (D_DP)","Note"],5)
for i,(cu) in enumerate(["EUR","GBP","PLN","SGD"]):
    ws.cell(row=6+i,column=1,value=cu).font=font(9); ws.cell(row=6+i,column=1).border=border
    tc=ws.cell(row=6+i,column=2); tc.fill=fill(YELLOW); tc.border=border; tc.number_format="0.000000"
    if cu=="EUR": tc.value=1
    ws.cell(row=6+i,column=3).border=border; ws.cell(row=6+i,column=4).border=border
widths(ws,{"A":10,"B":16,"C":14,"D":40})

# =========================================================== CONTROLES
ws=sheet("CONTROLES",BLUE)
title_block(ws,"CONTRÔLES","Voyants générés par Power Query (à brancher). Totaux par bloc, lignes non mappées, équilibre EUR, lignes sur entité par défaut.",BLUE)
checks=[("Total lignes en entrée","(PQ)"),("Total lignes en sortie","(PQ)"),
        ("Lignes NON mappées (compte/dimension)","(PQ) — à 0 idéalement"),
        ("Lignes sur ENTITÉ PAR DÉFAUT (à affiner)","(PQ)"),
        ("Total exclu (montant / %)","(PQ)"),
        ("Équilibre EUR (Magnitude vs SolaRE)","(PQ)"),
        ("Total bloc PNB+MEE","(PQ)"),("Total bloc OPEX","(PQ)"),("Total bloc Pré-tax","(PQ)")]
header_row(ws,["Contrôle","Valeur / statut"],5); rr=6
for lab,v in checks:
    ws.cell(row=rr,column=1,value=lab).font=font(10,True); ws.cell(row=rr,column=1).border=border
    ws.cell(row=rr,column=2,value=v).font=font(10,color=GREY); ws.cell(row=rr,column=2).border=border; rr+=1
widths(ws,{"A":44,"B":30})

# =========================================================== ETATS
for nm,desc in [("ETAT - Restitution Magnitude","Vue P&L façon « Réalisé vs Estimé (Valeurs) », en EUR. Sélecteur Scénario/Période. À générer via TCD/PQ."),
                ("ETAT - Restitution SolaRE","PMA × Entités, hiérarchie complète (nœuds agrégés), en EUR. Sélecteur Scénario/Période. À générer via TCD/PQ.")]:
    ws=sheet(nm,BLUE); title_block(ws,nm.replace("ETAT - ","ÉTAT — "),desc,BLUE)
    ws["A6"]="▾ Scénario :   [ segment ]        ▾ Période :   [ segment ]"; ws["A6"].font=font(10,True,BLUE)
    ws["A8"]="(Zone de restitution — tableau croisé dynamique à brancher sur la sortie du moteur.)"
    ws["A8"].font=font(9,it=True,color=GREY)
    widths(ws,{"A":90})

# =========================================================== SORTIE
ws=sheet("⑨ SORTIE - Table de fait",BLUE)
title_block(ws,"⑨ SORTIE — Table de fait Tagetik","Format d'import à plat (24 colonnes). Généré par Power Query. Montant en devise entité.",BLUE)
out=["Scenario","Scenario - Description","Period","Period - Description","Entity","Entity - Description",
     "Indicator","Indicator - Description","Counterparty","Counterparty - Description","PMA","PMA - Description",
     "Product","Product - Description","Vision","Vision - Description","Cost Center","Cost Center - Description",
     "Category","Category - Description","Entity currency","Entity currency - Description","Entity currency amount","Origin"]
header_row(ws,out,5)
for j in range(1,len(out)+1): ws.column_dimensions[get_column_letter(j)].width=16

# =========================================================== POWER QUERY (M) doc tab
ws=sheet("Requêtes Power Query (M)",GREY)
title_block(ws,"Requêtes Power Query (M) — à coller dans Excel","Données ▸ Obtenir des données ▸ À partir d'une requête vide ▸ Éditeur avancé. Une requête par bloc ci-dessous.",NAVY)
ws["A5"]="Voir le dossier « powerquery/ » du dépôt pour chaque script .pq (Source, Filtre, Exclusions, MapComptes, MapDimensions, Constantes, Sortie)."
ws["A5"].font=font(10,True,NAVY)
ws["A6"]="Ordre d'enchaînement : Source → Filtre → Exclusions → MapComptes → MapDimensions → Constantes → Sortie."
ws["A6"].font=font(10,False,NAVY)
widths(ws,{"A":120})

os.makedirs("powerquery",exist_ok=True)
out_path="Reprise_Magnitude_Tagetik.xlsx"
wb.save(out_path)
print("Classeur écrit:",out_path,"—",len(wb.sheetnames),"onglets")
print("Onglets:",wb.sheetnames)
