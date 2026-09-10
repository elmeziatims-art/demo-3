#!/usr/bin/env python3
"""Screen the 'contacts a valider' list against Neomind Advisory's ICP.

Neomind Advisory sells EPM / SPM / performance-management projects (Tagetik,
Anaplan, Board, SAP, Amelkis) to finance departments. The ICP encoded in the
sheet is a French group doing 80 M EUR - 1 Md EUR with a consolidation,
reporting or planning owner.

Each verdict below is annotated with the evidence it rests on:
  [F] a contradiction inside the file itself
  [W] checked against a public source, September 2026
  [C] my own reading of the company, to be confirmed

    python3 tools/screen_neomind.py neomind.xlsx neomind_screened.xlsx
"""

import collections
import sys

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from xlsxpatch import Sheet, SharedStrings, Styles, Workbook, col_letter  # noqa: E402

SHEET = 'contacts à valider'

# ---------------------------------------------------------------- verdicts
# row -> (verdict, motif). Rows absent from the table pass the screen.
#   FAUX   : out of ICP, tick "cocher pour retirer"
#   DOUTE  : one check needed before any outreach
#   GARDER : the file rejected it wrongly - keep, and fix the enrichment
V = {}

def mark(verdict, motif, *rows):
    for r in rows:
        V[r] = (verdict, motif)

# --- CA above the 1 Bn ceiling, yet flagged "Dans ICP (80M-1Md)" -----------
mark('FAUX', "Hors ICP par le haut : CA 1 412 M€, au-dessus du plafond 1 Md [F]", 7)
mark('FAUX', "Hors ICP par le haut : CA 1 407 M€, au-dessus du plafond 1 Md [F]", 65)
mark('FAUX', "Hors ICP par le haut : CA 1 208 M€, au-dessus du plafond 1 Md [F]", 105)
mark('FAUX', "Hors ICP par le haut : CA 1 244 M€, au-dessus du plafond 1 Md [F]", 107)
mark('DOUTE', "CA 1 025 M€ : dépasse le plafond de 2,5 % seulement, arbitrage à faire [F]", 89)

# --- competitors and same-market suppliers --------------------------------
mark('FAUX', "Concurrent direct : intégrateur IBM Planning Analytics / TM1, "
             "positionné sur le pilotage de la performance financière (21 pers.) [W]", 43)
mark('FAUX', "Absorbé par Accenture (acquisition finalisée le 01/04/2022) : plus de DAF "
             "autonome, et Accenture est concurrent sur l'EPM [W]", 3)
mark('FAUX', "Concurrent / partenaire : DAF et RH à temps partagé, ne sous-traite pas l'EPM [C]", 74)
mark('FAUX', "Cabinet de recrutement finance, et le contact est Responsable Recrutement : "
             "persona et société hors cible [F][C]", 36)
mark('FAUX', "Cabinet de conseil (1-2 salariés), Managing Partner : ne consomme pas d'EPM [F][C]", 120)
mark('FAUX', "Activité de courtage vocal cédée à Marex, finalisée le 01/02/2023 : "
             "le SIREN pointe déjà sur MAREX SA, décision au Royaume-Uni [W]", 121)

# --- the SIREN match landed on the wrong entity ---------------------------
mark('FAUX', "Rattachement SIREN aberrant : l'entité trouvée est une personne physique "
             "(ARDELIN CHARLES, livreur). Société réelle non identifiée [F]", 8)
mark('FAUX', "Rattachement SIREN aberrant : l'entité trouvée est une personne physique "
             "(ALI SAADOUN). Etanco appartient au groupe suisse SFS [F][C]", 57)
mark('FAUX', "Rattachement SIREN erroné : RH AULNAY n'est pas Atacadão (enseigne "
             "Carrefour Brésil). Le CA de 94 M€ ne concerne pas cette société [F][C]", 16)
mark('FAUX', "Le SIREN pointe sur un franchisé (MCDONALD S EST PARISIEN, 98 M€), pas sur "
             "McDonald's France : mono-activité, pas de besoin de consolidation [F][C]", 104)
mark('FAUX', "Le SIREN pointe sur une SCI (1-2 salariés), pas sur le groupe opérationnel : "
             "CA et taille non exploitables en l'état [F]", 116)
mark('FAUX', "SEM publique locale (50-99 pers.), SIREN rattaché à IMMOBILIERE PARIS SUD EST : "
             "hors ICP en taille et achat sous marché public [F][C]", 134)

# --- outside the addressable French perimeter -----------------------------
mark('FAUX', "Périmètre Congo, indiqué dans l'intitulé du poste [F]", 41)
mark('FAUX', "Périmètre Sénégal, indiqué dans l'intitulé du poste [F]", 113)
mark('FAUX', "Groupe agricole marocain (siège Agadir) [C]", 22)
mark('FAUX', "Société marocaine, aucun rattachement INPI trouvé [F][C]", 21)
mark('FAUX', "Site industriel marocain (aéronautique, Casablanca) [C]", 17)
mark('FAUX', "Trader énergie suisse/hongrois ; le rattachement GROUPE MET (DAVIS) "
             "ne correspond pas au groupe [C]", 108)
mark('FAUX', "Entité de 3-5 personnes, contact basé en Turquie [F][C]", 109)

# --- operating company well under the 80 M floor --------------------------
TOO_SMALL = {
    2: '10-19', 23: '20-49', 29: '10-19', 34: '10-19', 84: '20-49',
    90: '10-19', 96: '10-19', 128: '1-2', 129: '10-19', 131: '6-9', 138: '10-19',
    139: '20-49', 142: '20-49', 159: '20-49', 165: '10-19',
}
for row, eff in TOO_SMALL.items():
    mark('FAUX', f"Société opérationnelle de {eff} salariés : très loin du plancher "
                 f"de 80 M€, aucune structure de groupe à consolider [F]", row)
mark('FAUX', "Association (1-2 salariés) : ni la taille ni le modèle ne correspondent [F]", 95)

# --- persona is not a finance decision-maker ------------------------------
mark('FAUX', "Persona hors cible : rôle produit crypto-actifs, pas une direction "
             "financière ; société par ailleurs non identifiée [F]", 20)
mark('FAUX', "Persona hors cible : poste d'ingénieur, société non identifiée à l'INPI [F]", 27)

# --- the file rejected these wrongly --------------------------------------
mark('GARDER', "FAUX REJET : les 3-5 salariés sont ceux de la holding. Bertrand Franchise "
               "pèse 950 M€ de CA (Moody's, 12 mois à sept. 2025) et le contact est "
               "Responsable consolidation : cible prioritaire [W]", 25)
mark('GARDER', "FAUX REJET : holding de tête (6-9 salariés) du groupe Lov / FL "
               "Entertainment. Un Directeur consolidation implique un vrai périmètre "
               "groupe : requalifier le CA avant de trancher [F][C]", 99)
mark('GARDER', "FAUX REJET : le SIREN est une holding (1-2 salariés). Le poste est "
               "DG délégué aux finances : requalifier sur le CA groupe [F]", 40)
mark('GARDER', "FAUX REJET : le SIREN est une holding (1-2 salariés) alors que le poste "
               "est DAF Groupe : requalifier sur le CA consolidé [F]", 166)
mark('GARDER', "FAUX REJET : CA affiché 0 € alors que le fichier estime lui-même "
               "800 M€-1 Md€. Contrôleur de gestion d'un groupe multi-entités [F]", 122)
mark('GARDER', "FAUX REJET : CA affiché 0 €, le fichier note que le SIREN retenu n'est "
               "pas le bon. Transport/logistique estimé 100-200 M€ [F]", 124)
mark('GARDER', "FAUX REJET probable : les 50-99 salariés sont ceux de l'entité matchée ; "
               "Guyader Gastronomie pèse nettement plus [F][C]", 79)

# --- one check needed -----------------------------------------------------
mark('DOUTE', "Filiale d'un groupe étranger : la décision EPM se prend au siège. "
              "À qualifier sur l'autonomie du périmètre France [C]", 11, 26, 110, 71)
mark('DOUTE', "Périmètre géographique à confirmer : faisceau d'indices hors France "
              "(raison sociale, absence de rattachement INPI, localisation du contact) [F][C]",
     18, 30, 53, 106, 123, 127, 141, 146, 157)
mark('DOUTE', "Société non identifiée à l'INPI : ni CA ni taille vérifiables, "
              "à qualifier avant tout contact [F]", 37, 56)
mark('', "Bailleur social des Hauts-de-France, 270 salariés, 20 000 logements : "
         "entité française, périmètre multi-entités [W]", 80)
mark('FAUX', "Plateforme de mise en relation avec des consultants indépendants et de "
             "management de transition (10-19 pers.) : concurrent sur le conseil "
             "financier, pas un acheteur d'EPM [W]", 29)
mark('DOUTE', "Holding ou entité de tête peu peuplée : requalifier sur le CA consolidé "
              "avant de conclure [F][C]", 86, 152)
mark('DOUTE', "Intitulé « Finance Transformation » dans une société de 20-49 personnes : "
              "vérifier qu'il ne s'agit pas d'un cabinet de conseil concurrent [F][C]", 87)
mark('DOUTE', "CA estimé 50-100 M€, sous ou juste au plancher de l'ICP [F]", 5)
mark('DOUTE', "CA groupe autour de 1,3 Md€ (Asmodee Group, coté) : le 110 M€ affiché est "
              "une entité intermédiaire. Requalifier avant de conclure [F][C]", 13)
mark('DOUTE', "Effectif 100-199 en zone incertaine : confirmer le CA [F]", 93, 94, 164)
mark('DOUTE', "Effectif 50-99, CA probablement sous le plancher : confirmer [F]", 167)
mark('DOUTE', "Rattachement INPI faible ou entité intermédiaire : confirmer que le CA "
              "affiché est bien celui de la société visée [F]", 32, 46, 47, 48, 82, 161, 164)
mark('DOUTE', "CA non déclaré et non estimé : à qualifier [F]",
     12, 33, 39, 66, 72, 76, 78, 83, 85, 91, 103, 125, 126, 133, 137, 140, 153, 154)


# ---------------------------------------------------------------- pays
# Le fichier ne porte aucun champ pays, et LinkedIn est inaccessible depuis ce
# poste : le pays retenu est celui de la SOCIETE, jamais deduit du nom de la
# personne - un DAF francais peut porter n'importe quel nom.
#   row -> (pays, comment on le sait)
PAYS = {
    # verifie sur source publique, septembre 2026
    6:   ('Luxembourg', "vérifié — siège 48 rue de Bragance, L-1255 Luxembourg"),
    21:  ('Maroc', "vérifié — AXESS PHARMA, zone industrielle de Berrechid"),
    30:  ('Maroc', "vérifié — Câble d'Or, Meknès (distribution de matériel électrique)"),
    116: ('Maroc', "vérifié — holding marocaine ex-FinanceCom (famille Benjelloun) ; "
                   "le SIREN français retenu n'est qu'une SCI parisienne"),
    127: ('Maroc', "vérifié — Casablanca, bd Abdelhadi Boutaleb"),
    129: ('Maroc', "vérifié — Rabat, Hay Riad (conditionnement agricole)"),
    146: ('Maroc', "vérifié — Tanger, Z.I. Gzenaya (plasturgie)"),
    157: ('Maroc', "vérifié — Casablanca, route de Zenata (agroalimentaire)"),
    80:  ('France', "vérifié — Habilians, bailleur social à Villeneuve-d'Ascq "
                    "(fusion Habitat du Nord / Logis Métropole), 270 salariés"),
    101: ('France', "vérifié — Marbour, Le Port (La Réunion), ~300 M€, 900 collaborateurs"),
    29:  ('France', "vérifié — EOD SAS, Paris 9e"),
    # ecrit noir sur blanc dans l'intitule du poste
    41:  ('Congo', "indiqué dans l'intitulé du poste"),
    113: ('Sénégal', "indiqué dans l'intitulé du poste"),
    # forte presomption, a confirmer
    16:  ('Brésil', "à confirmer — Atacadão est l'enseigne de gros de Carrefour au Brésil"),
    17:  ('Maroc', "à confirmer — site aéronautique de Casablanca"),
    18:  ('Maroc', "à confirmer — aucune entité française identifiable"),
    22:  ('Maroc', "à confirmer — groupe agricole basé à Agadir"),
    53:  ('Portugal / Maroc', "à confirmer — poste « Country Finance Director », "
                              "pas d'entité française claire"),
    106: ('Maroc', "à confirmer — aucune entité française identifiable"),
    108: ('Suisse / Hongrie', "à confirmer — MET Group est un négociant en énergie suisse"),
    109: ('Turquie', "à confirmer — entité de 3-5 personnes"),
    123: ('Suisse / Italie', "à confirmer"),
    # non resolu
    27:  ('Indéterminé', "société non identifiée"),
    37:  ('Indéterminé', "société non identifiée"),
    56:  ('Indéterminé', "société non identifiée"),
    133: ('Indéterminé', "société non identifiée"),
    141: ('Indéterminé', "société non identifiée"),
}


def pays_de(row, siren, ca):
    """Pays retenu et son niveau de preuve."""
    if row in PAYS:
        return PAYS[row]
    if siren and ca:
        return ('France', "entité française au registre, chiffre d'affaires déclaré")
    if siren:
        return ('France', "entité française au registre, CA non déclaré")
    return ('Indéterminé', "aucun rattachement au registre français")


HEADERS = ['Verdict Neomind', 'Motif', 'Action', 'Pays (société)', 'Comment on le sait']
ACTIONS = {'FAUX': 'Retirer', 'DOUTE': 'Vérifier avant contact',
           'GARDER': 'Conserver - corriger l enrichissement', '': 'Contacter'}


def main(src, dst):
    wb = Workbook(src)
    st = Styles(wb.text('xl/styles.xml'))
    sst = SharedStrings(wb.text('xl/sharedStrings.xml'))
    part = wb.sheet_part(SHEET)
    sheet = Sheet(wb.text(part))

    head = st.xf(font=st.font('Calibri', 11, b=True), fill=st.fill('FFE7E6E6'),
                 border=st.border(bottom=('thin', 'FF262626')), halign='left')
    body = {
        'FAUX': st.xf(font=st.font('Calibri', 11, b=True, rgb='FFC41822'),
                      fill=st.fill('FFFDECEC'), halign='left', wrap=True),
        'DOUTE': st.xf(font=st.font('Calibri', 11, rgb='FF8A6D0B'),
                       fill=st.fill('FFFFFBEA'), halign='left', wrap=True),
        'GARDER': st.xf(font=st.font('Calibri', 11, b=True, rgb='FF1F6F3C'),
                        fill=st.fill('FFEAF6EE'), halign='left', wrap=True),
        '': st.xf(font=st.font('Calibri', 11, rgb='FF262626'), halign='left', wrap=True),
    }

    for col, text in zip((15, 16, 17, 18, 19), HEADERS):
        sheet.set_style(1, col, head)
        sheet.set_text(1, col, sst.add(text))

    counts = {'FAUX': 0, 'DOUTE': 0, 'GARDER': 0, '': 0}
    geo = collections.Counter()
    already = 0
    for row in range(2, 168):
        verdict, motif = V.get(row, ('', ''))
        counts[verdict] += 1
        style = body[verdict]
        action = ACTIONS[verdict]
        # ticks the owner made before this pass are never undone
        if verdict != 'FAUX' and sheet.has_cell(row, 1) \
                and sheet.rows[row].cells[1].inner == '<v>1</v>':
            already += 1
            action = 'Déjà coché par vous - le crible ICP ne le retire pas'
        pays, preuve = pays_de(row, sheet.has_cell(row, 11), sheet.has_cell(row, 9))
        if pays not in ('France', 'Indéterminé') and verdict != 'FAUX':
            verdict, style = 'FAUX', body['FAUX']
            counts['FAUX'] += 1
            counts[V.get(row, ('', ''))[0]] -= 1
            action = ACTIONS['FAUX']
        if pays not in ('France', 'Indéterminé'):
            motif = f"Hors périmètre France — {pays}. " + (motif or '')
        geo[pays] += 1
        for col, text in zip((15, 16, 17, 18, 19),
                             (verdict, motif, action, pays, preuve)):
            sheet.set_style(row, col, style)
            if text:
                sheet.set_text(row, col, sst.add(text))
        if verdict == 'FAUX':                       # tick "cocher pour retirer"
            cell = sheet.cell(row, 1)
            cell.attrs['t'] = 'b'
            cell.inner = '<v>1</v>'

    sheet.column_width(15, 15, 16)
    sheet.column_width(16, 16, 90)
    sheet.column_width(17, 17, 30)
    sheet.column_width(18, 18, 18)
    sheet.column_width(19, 19, 64)
    sheet.head = sheet.head.replace('$A$1:$N$167', '$A$1:$S$167')

    wb.put(part, sheet.serialize())
    wb.put('xl/styles.xml', st.serialize())
    wb.put('xl/sharedStrings.xml', sst.serialize())
    wb.save(dst)
    print(f'wrote {dst}')
    for k in ('FAUX', 'DOUTE', 'GARDER', ''):
        print(f'  {k or "OK":<7} {counts[k]:>4}')
    print(f'  déjà cochées hors crible : {already}')
    print('  pays :')
    for k, v in geo.most_common():
        print(f'     {k:<20} {v:>4}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
