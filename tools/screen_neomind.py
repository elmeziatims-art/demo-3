#!/usr/bin/env python3
"""Enrichit l'onglet « contacts à valider » : OK ou KO, et pourquoi.

Neomind Advisory vend des projets EPM / SPM (Tagetik, Anaplan, Board, SAP,
Amelkis) à des directions financières. L'ICP inscrit dans le fichier : un
groupe français de 80 M€ à 1 Md€ avec un propriétaire de la consolidation, du
reporting ou de la planification.

Un KO tombe pour une de ces raisons, dans cet ordre :
    pays hors France · société non identifiable · concurrent ou société
    absorbée · CA hors de la fourchette · rattachement SIREN erroné ·
    entité trop petite · persona hors cible · ligne déjà écartée par le
    propriétaire de la liste.

Chaque raison porte sa preuve : [F] contradiction interne au fichier,
[W] source publique vérifiée en septembre 2026, [C] appréciation à confirmer.

    python3 tools/screen_neomind.py neomind.xlsx neomind_enrichi.xlsx
"""

import sys

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from xlsxpatch import Sheet, SharedStrings, Styles, Workbook  # noqa: E402

SHEET = 'contacts à valider'
FIRST_ROW, LAST_ROW = 2, 167

# --------------------------------------------------------------- pays
# Le fichier n'a pas de champ pays et LinkedIn est inaccessible depuis ce
# poste : le pays retenu est celui de la SOCIETE. Il n'est jamais deduit du nom
# de la personne - ce serait faux autant qu'injuste, un DAF base en France
# pouvant porter n'importe quel nom.
PAYS = {
    6: 'Luxembourg', 16: 'Brésil', 17: 'Maroc', 18: 'Maroc', 21: 'Maroc',
    22: 'Maroc', 30: 'Maroc', 41: 'Congo', 53: 'Portugal / Maroc',
    106: 'Maroc', 108: 'Suisse / Hongrie', 109: 'Turquie', 113: 'Sénégal',
    116: 'Maroc', 123: 'Suisse / Italie', 127: 'Maroc', 129: 'Maroc',
    146: 'Maroc', 157: 'Maroc',
    27: 'Indéterminé', 37: 'Indéterminé', 56: 'Indéterminé',
    133: 'Indéterminé', 141: 'Indéterminé',
    3: 'France', 29: 'France', 42: 'France', 80: 'France', 101: 'France',
}

# --------------------------------------------------------------- verdicts
KO, OK = {}, {}


def ko(raison, *rows):
    for r in rows:
        KO[r] = raison


def note(raison, *rows):
    for r in rows:
        OK[r] = raison


# -- pays hors France ------------------------------------------------------
ko("Société luxembourgeoise (siège 48 rue de Bragance, L-1255) [W]", 6)
ko("Société marocaine : AXESS PHARMA, zone industrielle de Berrechid [W]", 21)
ko("Société marocaine : Câble d'Or, Meknès (distribution de matériel électrique) [W]", 30)
ko("Holding marocaine ex-FinanceCom (famille Benjelloun). Le SIREN français "
   "retenu par l'enrichissement n'est qu'une SCI parisienne [W][F]", 116)
ko("Société marocaine : Casablanca, bd Abdelhadi Boutaleb [W]", 127)
ko("Société marocaine : Rabat, Hay Riad (conditionnement agricole) [W]", 129)
ko("Société marocaine : Tanger, Z.I. Gzenaya (plasturgie) [W]", 146)
ko("Société marocaine : Casablanca, route de Zenata (agroalimentaire) [W]", 157)
ko("Périmètre Congo, écrit dans l'intitulé du poste [F]", 41)
ko("Périmètre Sénégal, écrit dans l'intitulé du poste [F]", 113)
ko("Atacadão est l'enseigne de gros de Carrefour au Brésil ; le SIREN retenu "
   "(RH AULNAY, 94 M€) n'a aucun rapport avec elle [F][C]", 16)
ko("Site aéronautique de Casablanca [C]", 17)
ko("Aucune entité française identifiable, faisceau d'indices Maroc [F][C]", 18, 106)
ko("Groupe agricole marocain basé à Agadir [C]", 22)
ko("Poste « Country Finance Director » hors France, aucune entité française claire [F][C]", 53)
ko("MET Group est un négociant en énergie suisse ; le rattachement "
   "GROUPE MET (DAVIS) ne correspond pas au groupe [C]", 108)
ko("Entité de 3 à 5 personnes, contact hors France [F][C]", 109)
ko("Société suisse ou italienne, aucune entité française identifiable [F][C]", 123)

# -- société non identifiable ---------------------------------------------
ko("Société non identifiable : aucun rattachement au registre français, "
   "ni CA ni taille vérifiables. Rien à qualifier en l'état [F]", 37, 56, 133, 141)

# -- concurrents et sociétés absorbées -------------------------------------
ko("Concurrent direct : intégrateur IBM Planning Analytics / TM1, positionné "
   "sur le pilotage de la performance financière (21 personnes) [W]", 43)
ko("Racheté par Accenture, acquisition finalisée le 01/04/2022 : plus de "
   "direction financière autonome, et Accenture est concurrent sur l'EPM [W]", 3)
ko("Courtage vocal cédé à Marex, finalisé le 01/02/2023 : le SIREN du fichier "
   "pointe déjà sur MAREX SA, la décision se prend au Royaume-Uni [W][F]", 121)
ko("Cabinet de DAF et RH à temps partagé : vend du conseil financier, "
   "n'en achète pas [C]", 74)
ko("Cabinet de recrutement finance, et le contact est Responsable Recrutement : "
   "société et persona hors cible [F][C]", 36)
ko("Cabinet de conseil de 1 à 2 personnes, contact Managing Partner [F][C]", 120)
ko("Plateforme de mise en relation avec des consultants indépendants et de "
   "management de transition, 10-19 personnes (EOD SAS, Paris 9e) [W]", 29)
ko("Cabinet de conseil : votre base contient chez eux un « Directeur practice "
   "consolidation », intitulé de prestataire et non de client [F][C]", 33)

# -- CA hors de la fourchette ICP -----------------------------------------
ko("Hors ICP par le haut : CA 1 412 M€, soit 41 % au-dessus du plafond de 1 Md€, "
   "alors que la ligne est marquée « Dans ICP (80M-1Md) » [F]", 7)
ko("Hors ICP par le haut : CA 1 407 M€ contre un plafond de 1 Md€ [F]", 65)
ko("Hors ICP par le haut : CA 1 208 M€ contre un plafond de 1 Md€ [F]", 105)
ko("Hors ICP par le haut : CA 1 244 M€ contre un plafond de 1 Md€ [F]", 107)
ko("Hors ICP par le haut : Asmodee Group pèse environ 1,3 Md€ ; les 110 M€ "
   "affichés sont ceux d'une entité intermédiaire. À noter : votre base "
   "contient chez eux un « Directeur Transformation Finance Groupe » mieux "
   "placé que ce contact, si vous décidez de garder le compte [F][C]", 13)

# -- rattachement SIREN erroné --------------------------------------------
ko("Rattachement SIREN aberrant : l'entité trouvée est une personne physique "
   "(ARDELIN CHARLES, livreur). Société réelle non identifiée [F]", 8)
ko("Rattachement SIREN aberrant : l'entité trouvée est une personne physique "
   "(ALI SAADOUN). Etanco appartient par ailleurs au groupe suisse SFS [F][C]", 57)
ko("Le SIREN pointe sur un franchisé (MCDONALD S EST PARISIEN, 98 M€) et non "
   "sur McDonald's France : mono-activité, pas de besoin de consolidation [F][C]", 104)
ko("SEM publique locale de 50 à 99 personnes, SIREN rattaché à IMMOBILIERE "
   "PARIS SUD EST : hors ICP en taille, et achat sous marché public [F][C]", 134)

# -- entité trop petite ----------------------------------------------------
TAILLE = {2: '10-19', 23: '20-49', 34: '10-19', 84: '20-49', 90: '10-19',
          96: '10-19', 128: '1-2', 131: '6-9', 138: '10-19', 139: '20-49',
          142: '20-49', 159: '20-49', 165: '10-19'}
for r, eff in TAILLE.items():
    ko(f"Société opérationnelle de {eff} salariés : très loin du plancher de "
       f"80 M€, et aucune structure de groupe à consolider [F]", r)
ko("Association de 1 à 2 salariés : ni la taille ni le modèle ne correspondent [F]", 95)

# -- filiale d'un groupe étranger -----------------------------------------
ko("Filiale française d'un groupe étranger : la décision sur l'outil de "
   "consolidation se prend au siège, hors de votre périmètre [C]", 11, 26, 110)
ko("Entité française de 20 à 49 personnes d'un groupe suisse : ni la taille ni "
   "l'autonomie de décision ne correspondent [F][C]", 71)

# -- persona hors cible ----------------------------------------------------
ko("Persona hors cible : rôle produit crypto-actifs, pas une direction "
   "financière ; société par ailleurs non identifiée [F]", 20)
ko("Persona hors cible : poste d'ingénieur, et société non identifiable [F]", 27)

# -- déjà écartées par vous ------------------------------------------------
ko("Ligne que vous aviez déjà cochée. Elle passe pourtant tous les critères "
   "ICP : le motif vous appartient (client existant, deal en cours, contact "
   "obsolète ?) et mérite d'être écrit, sinon le crible la fera revenir", 9, 10, 118)
ko("Ligne que vous aviez déjà cochée, et l'intitulé « Finance Transformation » "
   "dans une société de 20 à 49 personnes suggère effectivement un cabinet", 87)
ko("Ligne que vous aviez déjà cochée. Effectif 100-199, CA non confirmé", 93, 94)

# -- OK méritant une précision --------------------------------------------
note("OK — CA 1 025 M€, soit 2,5 % au-dessus du plafond : dans la marge "
     "d'incertitude, à trancher plutôt qu'à écarter [F]", 89)
note("OK — CA estimé 50-100 M€, donc à la limite basse de l'ICP : confirmer "
     "avant d'investir du temps [F]", 5)
note("OK — bailleur social des Hauts-de-France, 270 salariés, 20 000 logements. "
     "Entité française multi-entités : le crible l'avait écartée à tort [W]", 80)
note("OK — Le Port (La Réunion), environ 300 M€ et 900 collaborateurs, donc "
     "France. Le crible l'avait laissée en doute [W]", 101)
note("OK — CA 89 M€, juste au-dessus du plancher [F]", 156)
note("OK — CA 80 M€, exactement au plancher de l'ICP [F]", 24)
note("OK — CA 83 M€, juste au-dessus du plancher [F]", 114)
note("OK — opérateur de data centers français (Marcoussis), CA estimé "
     "200-400 M€. Pas de rattachement INPI mais société bien française [F][C]", 42)
note("FAUX REJET du crible : les 3 à 5 salariés sont ceux de la holding. "
     "Bertrand Franchise pèse 950 M€ de CA (Moody's, 12 mois à septembre 2025) "
     "et le contact est Responsable consolidation. Cible prioritaire [W]", 25)
note("FAUX REJET du crible : holding de tête (6-9 salariés) du groupe Lov / FL "
     "Entertainment. Un Directeur consolidation suppose un vrai périmètre "
     "groupe : requalifier le CA consolidé [F][C]", 99)
note("FAUX REJET du crible : le SIREN est une holding de 1 à 2 salariés, mais "
     "le poste est DG délégué aux finances. Requalifier sur le CA groupe [F]", 40)
note("FAUX REJET du crible : le SIREN est une holding de 1 à 2 salariés alors "
     "que le poste est DAF Groupe. Requalifier sur le CA consolidé [F]", 166)
note("FAUX REJET du crible : CA affiché 0 € alors que le fichier estime "
     "lui-même 800 M€-1 Md€. Contrôleur de gestion d'un groupe multi-entités [F]", 122)
note("FAUX REJET du crible : CA affiché 0 €, et le fichier note que le SIREN "
     "retenu n'est pas le bon. Transport et logistique, estimé 100-200 M€ [F]", 124)
note("FAUX REJET probable : les 50 à 99 salariés sont ceux de l'entité "
     "matchée ; Guyader Gastronomie pèse nettement plus [F][C]", 79)
note("OK — mais 5 contacts sur ce seul compte dans la liste : garder les rôles "
     "groupe, mettre les périmètres France au second rang [F]", 59, 60, 61, 62, 63)


def main(src, dst):
    wb = Workbook(src)
    st = Styles(wb.text('xl/styles.xml'))
    sst = SharedStrings(wb.text('xl/sharedStrings.xml'))
    part = wb.sheet_part(SHEET)
    sheet = Sheet(wb.text(part))

    head = st.xf(font=st.font('Calibri', 11, b=True), fill=st.fill('FFE7E6E6'),
                 border=st.border(bottom=('thin', 'FF262626')), halign='left')
    style = {
        'KO': st.xf(font=st.font('Calibri', 11, b=True, rgb='FFC41822'),
                    fill=st.fill('FFFDECEC'), halign='left', wrap=True),
        'OK': st.xf(font=st.font('Calibri', 11, b=True, rgb='FF1F6F3C'),
                    fill=st.fill('FFEAF6EE'), halign='left', wrap=True),
    }

    for col, text in zip((15, 16, 17), ('OK / KO', 'Raison', 'Pays (société)')):
        sheet.set_style(1, col, head)
        sheet.set_text(1, col, sst.add(text))

    tally = {'OK': 0, 'KO': 0}
    for row in range(FIRST_ROW, LAST_ROW + 1):
        verdict = 'KO' if row in KO else 'OK'
        raison = KO.get(row) or OK.get(row) or (
            "OK — entité française, persona finance, "
            + ("CA dans la fourchette 80 M€-1 Md€" if sheet.has_cell(row, 9)
               else "CA non déclaré au registre : à confirmer avant contact"))
        pays = PAYS.get(row, 'France' if sheet.has_cell(row, 11) else 'Indéterminé')
        tally[verdict] += 1
        for col, text in zip((15, 16, 17), (verdict, raison, pays)):
            sheet.set_style(row, col, style[verdict])
            sheet.set_text(row, col, sst.add(text))
        if verdict == 'KO':                      # coche « cocher pour retirer »
            cell = sheet.cell(row, 1)
            cell.attrs['t'] = 'b'
            cell.inner = '<v>1</v>'

    sheet.column_width(15, 15, 10)
    sheet.column_width(16, 16, 104)
    sheet.column_width(17, 17, 18)
    sheet.head = sheet.head.replace('$A$1:$N$167', '$A$1:$Q$167')

    wb.put(part, sheet.serialize())
    wb.put('xl/styles.xml', st.serialize())
    wb.put('xl/sharedStrings.xml', sst.serialize())
    wb.save(dst)
    print(f'wrote {dst}')
    print(f'  OK {tally["OK"]:>4}\n  KO {tally["KO"]:>4}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
