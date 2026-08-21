#!/usr/bin/env python3
"""Classeur Reprise Magnitude -> Tagetik.
Onglets sources de mapping (RU->EJ, OA->PMA, FA->CC) avec colonne Defaut,
MAP - Dimensions pre-rempli (DERNIER MOT), comptes/ETP/exclusions/taux/zone,
referentiels, rate reel, echantillon liasse, sortie. Tableaux Excel crees
automatiquement (pas de Ctrl+T a faire). Moteur = scripts powerquery/."""
import csv, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

DOCS="docs"
def rd(name, header=True):
    with open(os.path.join(DOCS,name),encoding="utf-8") as f:
        r=list(csv.reader(f))
    return (r[0],r[1:]) if header else (None,r)

NAVY="17262A"; ACC="0E7C66"; AMBER="B26A13"; VIOLET="6A4B8A"; BLUE="2A6F8E"
GREY="5C6F72"; LINE="C6D0CB"; WHITE="FFFFFF"; REDTXT="C0392B"; REDFILL="FBE9E7"; YELLOW="FFF6D8"
F="Arial"
def font(sz=10,b=False,color=NAVY,it=False): return Font(name=F,size=sz,bold=b,color=color,italic=it)
def fill(c): return PatternFill("solid",fgColor=c)
thin=Side(style="thin",color=LINE); border=Border(left=thin,right=thin,top=thin,bottom=thin)
center=Alignment(horizontal="center",vertical="center",wrap_text=True)
left=Alignment(horizontal="left",vertical="center",wrap_text=True)

wb=Workbook(); wb.remove(wb.active)
def sheet(name,color=None):
    ws=wb.create_sheet(name)
    if color: ws.sheet_properties.tabColor=color
    ws.sheet_view.showGridLines=False
    return ws
def title_block(ws,title,sub,color=ACC):
    ws["A1"]=title; ws["A1"].font=font(15,True,color)
    ws["A2"]=sub; ws["A2"].font=font(9,False,GREY,it=True)
    ws.row_dimensions[1].height=22
def hrow(ws,headers,row,color=NAVY,txt=WHITE):
    for j,h in enumerate(headers,1):
        c=ws.cell(row=row,column=j,value=h); c.font=font(9,True,txt); c.fill=fill(color)
        c.alignment=center; c.border=border
    ws.freeze_panes=ws.cell(row=row+1,column=1)
def widths(ws,m):
    for k,v in m.items(): ws.column_dimensions[k].width=v
def add_table(ws,name,first_row,ncols,nrows):
    """Cree un Tableau Excel nomme sur l'entete (first_row) + nrows lignes."""
    last=get_column_letter(ncols)
    ref=f"A{first_row}:{last}{first_row+max(nrows,1)}"
    t=Table(displayName=name,ref=ref)
    t.tableStyleInfo=TableStyleInfo(name="TableStyleLight9",showRowStripes=True)
    ws.add_table(t)

# ---------- 0. Accueil
ws=sheet("0. Accueil",ACC)
ws["B2"]="Reprise Magnitude  →  Tagetik (SolaRE)"; ws["B2"].font=font(18,True,ACC)
ws["B3"]="Moteur Excel / Power Query — sans VBA. Tableaux déjà créés (aucun Ctrl+T à faire)."; ws["B3"].font=font(10,False,GREY,it=True)
r=6
ws[f"B{r}"]="Mode d'emploi"; ws[f"B{r}"].font=font(12,True,NAVY); r+=1
for n,t,d in [("1","Coller","Collez votre extraction Magnitude (A→Y) dans « ① ENTREE » (un échantillon est déjà là)."),
              ("2","Actualiser","Données ▸ Actualiser tout."),
              ("3","Lire","Récupérez « ⑨ SORTIE » + les 2 états de restitution.")]:
    ws[f"B{r}"]=n; ws[f"B{r}"].font=font(14,True,WHITE); ws[f"B{r}"].fill=fill(ACC); ws[f"B{r}"].alignment=center
    ws[f"C{r}"]=t; ws[f"C{r}"].font=font(11,True,NAVY)
    ws[f"D{r}"]=d; ws[f"D{r}"].font=font(10); ws[f"D{r}"].alignment=left; ws.row_dimensions[r].height=26; r+=1
r+=1
ws[f"B{r}"]="Chaîne de mapping"; ws[f"B{r}"].font=font(12,True,NAVY); r+=1
for d in ["RU → EJ  (+ défaut) ,   OA → PMA ,   FA → CC  (+ défaut)   alimentent…",
          "…MAP - Dimensions  →  qui a TOUJOURS le dernier mot sur l'output.",
          "Rouge = à valider / à saisir.   Toutes les tables sont modifiables."]:
    ws[f"B{r}"]=d; ws[f"B{r}"].font=font(10,color=NAVY); r+=1
widths(ws,{"A":2,"B":22,"C":24,"D":86})

# ---------- 1. ENTREE (+ echantillon)
ws=sheet("1. ① ENTREE - Magnitude",AMBER)
title_block(ws,"① ENTREE — Extraction Magnitude","Échantillon réel (1 500 lignes) pré-collé. Remplacez/complétez par votre extraction (A→Y).",AMBER)
_,srows=rd("_sample_entree.csv",header=False)
cols_src=srows[0]
hrow(ws,cols_src,4)
rr=5
for row in srows[1:]:
    for j,v in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(8); c.border=border
    rr+=1
for j in range(1,len(cols_src)+1): ws.column_dimensions[get_column_letter(j)].width=11
add_table(ws,"tblEntree",4,len(cols_src),len(srows)-1)

# ---------- MAP - RU vers EJ (+ Defaut)
ws=sheet("MAP - RU vers EJ",ACC)
title_block(ws,"MAP — RU → EJ (entité)","Mapping 1→N. Colonne « Défaut » = l'EJ retenue par défaut pour le RU (pré-tagué sur la 1ʳᵉ). Éditable.")
_,ruej=rd("mapping_RU_entities.csv",header=False)
hrow(ws,["RU","Lib Magnitude","Devise","EJ (SOLARE)","Lib SOLARE","Défaut"],4)
seen=set(); rr=5
dvx=DataValidation(type="list",formula1='"X"',allow_blank=True); ws.add_data_validation(dvx)
for row in ruej:
    ru=row[0]
    vals=row[:5]+["X" if ru not in seen else ""]
    for j,v in enumerate(vals,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(9); c.border=border; c.alignment=left
    ws.cell(row=rr,column=6).fill=fill(YELLOW); ws.cell(row=rr,column=6).alignment=center
    seen.add(ru); rr+=1
dvx.add(f"F5:F{rr}")
widths(ws,{"A":13,"B":26,"C":8,"D":13,"E":42,"F":9})
add_table(ws,"tblRUEJ",4,6,len(ruej))

# ---------- MAP - OA vers PMA (NOT USED rouge)
ws=sheet("MAP - OA vers PMA",ACC)
title_block(ws,"MAP — OA → PMA","Business Lines. NOT USED en rouge (mais reste utilisable si présent dans les données). Éditable.")
_,oapma=rd("mapping_OA_business_lines.csv",header=False)
hrow(ws,["Code MAGNITUDE (OA)","Code SOLARE (PMA)","Description"],4); rr=5
for row in oapma:
    for j,v in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(9); c.border=border; c.alignment=left
    if "NOT USED" in (row[1] or ""):
        for j in range(1,4):
            ws.cell(row=rr,column=j).font=font(9,True,REDTXT); ws.cell(row=rr,column=j).fill=fill(REDFILL)
    rr+=1
widths(ws,{"A":22,"B":18,"C":48})
add_table(ws,"tblOAPMA",4,3,len(oapma))

# ---------- MAP - FA vers CC (+ Defaut)
ws=sheet("MAP - FA vers CC",ACC)
title_block(ws,"MAP — FA → Cost Center","Aide au choix 1→N. Colonne « Défaut » = le CC retenu par défaut pour le FA (pré-tagué sur le 1ᵉʳ). Éditable.")
h,facc=rd("mapping_FA_to_CC.csv")
hrow(ws,["FA","FA_libellé","CC","CC_libellé","Défaut"],4)
seen=set(); rr=5
dvx2=DataValidation(type="list",formula1='"X"',allow_blank=True); ws.add_data_validation(dvx2)
for row in facc:
    fa=row[0]
    vals=row[:4]+["X" if fa not in seen else ""]
    for j,v in enumerate(vals,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(9); c.border=border; c.alignment=left
    ws.cell(row=rr,column=5).fill=fill(YELLOW); ws.cell(row=rr,column=5).alignment=center
    seen.add(fa); rr+=1
dvx2.add(f"E5:E{rr}")
widths(ws,{"A":10,"B":34,"C":12,"D":40,"E":9})
add_table(ws,"tblFACC",4,5,len(facc))

# ---------- MAP - Dimensions (PRE-REMPLI, DERNIER MOT)
ws=sheet("MAP - Dimensions",ACC)
title_block(ws,"MAP — Dimensions (RU × OA × FA → Entité · PMA · CC)  —  DERNIER MOT","Pré-rempli avec ce qu'on sait (Entité = EJ défaut du RU). Case vide = tout. « Le plus spécifique gagne » + Priorité. PMA/CC vides ⇒ pris des tables OA→PMA / FA→CC ; renseignés ici ⇒ forcés.")
dh=["RU","OA","FA","ENTITE","PMA","Cost_Center","Priorite","Source","Commentaire"]
hrow(ws,dh,4)
_,drows=rd("mapping_dim_defaut_RU_entite.csv"); rr=5
for row in drows:
    ru,oa,fa,ent,lib,src=row
    vals=[ru,oa,fa,ent,"","","",src,lib]
    for j,v in enumerate(vals,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(9); c.border=border; c.alignment=left
    ws.cell(row=rr,column=4).fill=fill(YELLOW)
    rr+=1
widths(ws,{"A":12,"B":10,"C":10,"D":14,"E":12,"F":14,"G":9,"H":16,"I":34})
add_table(ws,"tblDim",4,len(dh),len(drows))

# ---------- MAP - Comptes P&L
ws=sheet("MAP - Comptes P&L",ACC)
title_block(ws,"MAP — Comptes P&L (D_AC → Indicateur)","Variante France (IND_France) / Pays (IND_Pays). Rouge = à valider/compléter.")
h,rows=rd("mapping_comptes_D_AC.csv"); hrow(ws,h,4); rr=5
for row in rows:
    for j,v in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(9); c.border=border; c.alignment=left
    if len(row)>7 and row[3]=="Compte fin" and not row[7]:
        for j in range(1,len(h)+1):
            ws.cell(row=rr,column=j).font=font(9,True,REDTXT); ws.cell(row=rr,column=j).fill=fill(REDFILL)
    rr+=1
widths(ws,{"A":10,"B":34,"C":16,"D":11,"E":10,"F":8,"G":9,"H":15,"I":30,"J":15,"K":30,"L":40})
add_table(ws,"tblComptes",4,len(h),len(rows))

# ---------- MAP - ETP
ws=sheet("MAP - ETP (Q99)",ACC)
title_block(ws,"MAP — ETP (flux Q99 → Indicateur headcount)","GR050 reste à mapper (rouge).")
h,rows=rd("mapping_ETP_Q99.csv"); hrow(ws,h,4); rr=5
for row in rows:
    for j,v in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(9); c.border=border; c.alignment=left
    if not row[1]:
        for j in range(1,len(h)+1):
            ws.cell(row=rr,column=j).font=font(9,True,REDTXT); ws.cell(row=rr,column=j).fill=fill(REDFILL)
    rr+=1
widths(ws,{"A":16,"B":18,"C":30,"D":16,"E":40})
add_table(ws,"tblETP",4,len(h),len(rows))

# ---------- MAP - Exclusions
ws=sheet("MAP - Exclusions",ACC)
title_block(ws,"MAP — Exclusions / Ajustements","Retirer X% ou X€ d'un RU × OA × FA (case vide = tout). AVANT le mapping, au prorata, F99 ET ETP.")
eh=["RU","OA","FA","Mode","Valeur","Commentaire"]; hrow(ws,eh,4)
ex=[["G-REIMLUX","","","%","30","EXEMPLE — retire 30% de tout ce RU"],
    ["","OA060","","Montant","200000","EXEMPLE — 200k€ sur cet OA"]]
rr=5
for row in ex:
    for j,v in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(9,it=True,color=GREY); c.border=border; c.alignment=left
    rr+=1
for _ in range(20):
    for j in range(1,7): ws.cell(row=rr,column=j).border=border
    rr+=1
dv=DataValidation(type="list",formula1='"Montant,%"',allow_blank=True); ws.add_data_validation(dv); dv.add("D5:D200")
widths(ws,{"A":14,"B":10,"C":10,"D":12,"E":12,"F":44})
add_table(ws,"tblExcl",4,6,22)

# ---------- MAP - Taux charges
ws=sheet("MAP - Taux charges soc.",ACC)
title_block(ws,"MAP — Taux de charges sociales / RU","Un % par RU. base = GR2100−GR2101−GR2102 ; salaire_hors = base/(1+taux) ; charges = base−salaire_hors.")
hrow(ws,["RU","Taux","Commentaire"],4)
_,zrows=rd("ref_RU_pays_zone.csv"); rus=[r[0] for r in zrows]; rr=5
for ru in rus:
    ws.cell(row=rr,column=1,value=ru).font=font(9); ws.cell(row=rr,column=1).border=border
    tc=ws.cell(row=rr,column=2); tc.fill=fill(YELLOW); tc.border=border; tc.number_format="0.0%"
    ws.cell(row=rr,column=3).border=border; rr+=1
widths(ws,{"A":14,"B":14,"C":40})
add_table(ws,"tblTaux",4,3,len(rus))

# ---------- MAP - Zone RU
ws=sheet("MAP - Zone RU",ACC)
title_block(ws,"MAP — Zone RU (France / Pays)","Choisit la variante d'indicateur France/Pays. Éditable.")
h,rows=rd("ref_RU_pays_zone.csv"); hrow(ws,["RU","Pays","Zone"],4); rr=5
dvz=DataValidation(type="list",formula1='"France,Pays"',allow_blank=True); ws.add_data_validation(dvz)
for row in rows:
    for j,v in enumerate(row,1):
        c=ws.cell(row=rr,column=j,value=v); c.font=font(9); c.border=border; c.alignment=left
        if j==3: c.fill=fill(YELLOW)
    rr+=1
dvz.add(f"C5:C{rr}")
widths(ws,{"A":14,"B":20,"C":14})
add_table(ws,"tblZone",4,3,len(rows))

# ---------- REF (lecture seule)
def ref_sheet(name,csvfile,headers=None,wm=None,has_header=True):
    ws=sheet(name,VIOLET)
    title_block(ws,name.replace("REF - ","REF — "),"Référentiel — lecture seule.",VIOLET)
    h,rows=rd(csvfile,header=has_header); hh=h if has_header else headers
    hrow(ws,hh,4,color=VIOLET); rr=5
    for row in rows:
        for j,v in enumerate(row,1):
            c=ws.cell(row=rr,column=j,value=v); c.font=font(9); c.border=border; c.alignment=left
        rr+=1
    if wm: widths(ws,wm)
ref_sheet("REF - Indicateurs","tagetik_indicateurs.csv",wm={"A":8,"B":26,"C":7,"D":14,"E":18,"F":52})
ref_sheet("REF - Hier Entites","hierarchie_entites.csv",wm={"A":14,"B":40,"C":14,"D":16})
ref_sheet("REF - Hier PMA","hierarchie_pma.csv",wm={"A":14,"B":44,"C":14,"D":16})
ref_sheet("REF - RUxOA Mappable","mapping_RUxOA_mappable.csv",wm={"A":12,"B":14,"C":8,"D":10,"E":36,"F":22,"G":22,"H":40,"I":10})
ref_sheet("REF - RUxOA NonMappable","mapping_RUxOA_non_mappable.csv",wm={"A":12,"B":12,"C":8,"D":30,"E":26,"F":26,"G":50,"H":50,"I":8})

# ---------- rate (reel)
ws=sheet("rate",ACC)
title_block(ws,"rate — Taux de change (période 03) → EUR","Taux finaux réels (mars 2026). Montant_EUR = montant × « Taux vers EUR ».")
h,rrows=rd("_rate_03.csv"); hrow(ws,["Devise","Taux final 03","Taux vers EUR","Note"],4); rr=5
for row in rrows:
    ws.cell(row=rr,column=1,value=row[0]).font=font(9); ws.cell(row=rr,column=1).border=border
    c2=ws.cell(row=rr,column=2,value=float(row[1])); c2.number_format="0.000000"; c2.border=border
    c3=ws.cell(row=rr,column=3,value=float(row[2])); c3.number_format="0.000000"; c3.border=border; c3.fill=fill(YELLOW)
    ws.cell(row=rr,column=4).border=border; rr+=1
widths(ws,{"A":10,"B":14,"C":14,"D":36})
add_table(ws,"tblRate",4,4,len(rrows))

# ---------- CONTROLES
ws=sheet("CONTROLES",BLUE)
title_block(ws,"CONTRÔLES","Voyants Power Query (à brancher).",BLUE)
hrow(ws,["Contrôle","Valeur / statut"],4); rr=5
for lab,v in [("Lignes en entrée","(PQ)"),("Lignes en sortie","(PQ)"),
    ("Lignes NON mappées","(PQ) — idéalement 0"),("Lignes sur ENTITÉ PAR DÉFAUT","(PQ)"),
    ("Total exclu","(PQ)"),("Équilibre EUR (Magnitude vs SolaRE)","(PQ)"),
    ("Total PNB+MEE","(PQ)"),("Total OPEX","(PQ)"),("Total Pré-tax","(PQ)")]:
    ws.cell(row=rr,column=1,value=lab).font=font(10,True); ws.cell(row=rr,column=1).border=border
    ws.cell(row=rr,column=2,value=v).font=font(10,color=GREY); ws.cell(row=rr,column=2).border=border; rr+=1
widths(ws,{"A":40,"B":30})

# ---------- ETATS
for nm,desc in [("ETAT - Restitution Magnitude","Vue P&L façon « Réalisé vs Estimé », en EUR. Segments Scénario/Période. TCD à brancher sur la sortie."),
                ("ETAT - Restitution SolaRE","PMA × Entités, hiérarchie complète (nœuds), en EUR. Segments Scénario/Période. TCD à brancher.")]:
    ws=sheet(nm,BLUE); title_block(ws,nm.replace("ETAT - ","ÉTAT — "),desc,BLUE)
    ws["A6"]="▾ Scénario : [segment]      ▾ Période : [segment]"; ws["A6"].font=font(10,True,BLUE)
    ws["A8"]="(Zone de restitution — tableau croisé dynamique branché sur la sortie du moteur.)"; ws["A8"].font=font(9,it=True,color=GREY)
    widths(ws,{"A":90})

# ---------- SORTIE
ws=sheet("⑨ SORTIE - Table de fait",BLUE)
title_block(ws,"⑨ SORTIE — Table de fait Tagetik","Format d'import à plat (24 colonnes). Généré par Power Query.",BLUE)
out=["Scenario","Scenario - Description","Period","Period - Description","Entity","Entity - Description",
 "Indicator","Indicator - Description","Counterparty","Counterparty - Description","PMA","PMA - Description",
 "Product","Product - Description","Vision","Vision - Description","Cost Center","Cost Center - Description",
 "Category","Category - Description","Entity currency","Entity currency - Description","Entity currency amount","Origin"]
hrow(ws,out,4)
for j in range(1,len(out)+1): ws.column_dimensions[get_column_letter(j)].width=16

# ---------- PQ doc
ws=sheet("Requêtes Power Query (M)",GREY)
title_block(ws,"Requêtes Power Query (M) — à coller","Données ▸ Obtenir des données ▸ Requête vide ▸ Éditeur avancé. Voir dossier powerquery/.",NAVY)
ws["A4"]="Ordre : Source → Filtre → Exclusions → MapComptes → MapDimensions → Constantes → Sortie."; ws["A4"].font=font(10,True,NAVY)
ws["A5"]="Les Tableaux sont déjà créés (tblEntree, tblRUEJ, tblOAPMA, tblFACC, tblDim, tblComptes, tblETP, tblExcl, tblTaux, tblZone, tblRate)."; ws["A5"].font=font(10,color=NAVY)
widths(ws,{"A":120})

out_path="Reprise_Magnitude_Tagetik.xlsx"
wb.save(out_path)
print("OK",out_path,"—",len(wb.sheetnames),"onglets")
print(wb.sheetnames)
