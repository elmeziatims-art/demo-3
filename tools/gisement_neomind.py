#!/usr/bin/env python3
"""Second pass on the Neomind list: what the promotion filter left behind.

The 166 rows of "contacts a valider" were drawn from an export of 1 818
LinkedIn connections held on the hidden "contacts base" tab. This pass works
the other way round: it takes every connection carrying the strongest EPM
buying signal - consolidation, EPM, finance transformation, or a named tool
(Tagetik, Anaplan, OneStream, Hyperion, SAP FC) - and reports the ones that
never reached the validation list, sorted into what they actually are.

    python3 tools/gisement_neomind.py neomind.xlsx gisement.xlsx
"""

import collections
import csv
import io
import re
import sys
import zipfile

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- the buying signal Neomind actually sells against ---------------------
EPM = re.compile(r'consolidation|\bepm\b|finance transformation|transformation financ|'
                 r'normes ifrs|hyperion|tagetik|anaplan|onestream|planning analytics|'
                 r'\btm1\b|sap fc|magnitude|sigma conso|lucanet|\bbfc\b', re.I)
NOISE = re.compile(r'\b(stage|stagiaire|intern|alternan|apprenti|junior|assistant|'
                   r'[ée]tudiant|student|open to work)\b', re.I)

# --- who is not a prospect ------------------------------------------------
RIVAL = re.compile(
    r'\b(deloitte|\bey\b|ernst|pwc|kpmg|mazars|forvis|grant thornton|\brsm\b|\bbdo\b|bm&a|'
    r'primexis|arthaud|satriun|newreport|sqorus|micropole|viseo|visea|talan|wavestone|'
    r'accenture|capgemini|sopra|\bcgi\b|keyrus|business ?& ?decision|hardis|timspirit|'
    r'denjean|ginini|ifb group|artefact|eight advisory|alvarez|mckinsey|\bbcg\b|bain|'
    r'oliver wyman|kearney|2cfinance|absoluce|baker tilly|\bexco\b|fiducial|in extenso|'
    r'tagetik|anaplan|onestream|jedox|insightsoftware|board international|amelkis|'
    r'cegid|lucanet|sigma conso|morgan philips|robert half|hays|michael page|fed finance|'
    r'conseil|consulting|consultants|advisory|audit|expertise.comptable|cabinet|'
    r'recrutement|int[ée]rim)\b', re.I)
MEGA = re.compile(
    r'\b(lvmh|louis vuitton|kering|herm[èe]s|chanel|dior|carrefour|auchan|leclerc|casino|'
    r'veolia|suez|engie|\bedf\b|totalenergies|renault|stellantis|valeo|forvia|faurecia|'
    r'michelin|airbus|safran|thales|dassault|naval group|sanofi|l.or[ée]al|danone|nestl[ée]|'
    r'unilever|pepsico|coca.cola|\bbel\b|lactalis|vinci|bouygues|eiffage|saint.gobain|'
    r'schneider|legrand|air liquide|arcelor|solvay|orange|\bsfr\b|bouygues telecom|'
    r'publicis|havas|\bwpp\b|vivendi|canal\+|\btf1\b|\bm6\b|bnp|soci[ée]t[ée] g[ée]n[ée]rale|'
    r'cr[ée]dit agricole|\bbpce\b|cr[ée]dit mutuel|ark[ée]a|caisse des d[ée]p[ôo]ts|\baxa\b|'
    r'generali|allianz|\bcnp\b|groupama|maif|macif|cov[ée]a|amundi|natixis|hsbc|abn amro|'
    r'accor|sodexo|elior|atos|worldline|teleperformance|ubisoft|\bsncf\b|\bratp\b|la poste|'
    r'eurofins|vallourec|edenred|veepee|manitou|exclusive networks|ceva logistics|'
    r'africa global logistics|cma cgm|bollor[ée]|\bsaur\b|emeis|invivo|\bsuez\b|'
    r'dsm.firmenich|gsk|pfizer|novartis|bayer|siemens|bosch|foundever|unilabs|'
    r'showroomprive|ovhcloud)\b', re.I)
# Marqueurs geographiques explicites dans l'intitule du poste ou le nom de la
# societe. Seule preuve de pays disponible : le fichier n'a pas de champ pays et
# LinkedIn est inaccessible. Jamais deduit du nom de la personne.
GEO = [
    ('Maroc', r"\b(maroc|morocco|casablanca|rabat|tanger|marrakech|agadir|kenitra|meknes|"
              r"berrechid)\b|marjane|risma s\.a|aradei|saham bank|attijari|\bocp\b|cimaf|"
              r"tesca morocco|groupe managem|cosumar"),
    ('Afrique subsaharienne', r"\b(s[ée]n[ée]gal|dakar|c[ôo]te d.ivoire|abidjan|cameroun|gabon|"
                              r"congo|kinshasa|togo|b[ée]nin|burkina|guin[ée]e|nigeria|ghana|"
                              r"kenya|afrique|africa)\b"),
    ('Maghreb hors Maroc', r"\b(tunisie|tunisia|tunis|alg[ée]rie|algeria|alger)\b"),
    ('Moyen-Orient', r"\b(dubai|dubaï|uae|emirates|qatar|riyad|saudi|middle east|beyrouth|liban)\b"),
    ('Canada', r"\b(canada|montr[ée]al|qu[ée]bec|toronto|ottawa|vancouver|calgary)\b"),
    ('Europe hors France', r"\b(belgique|bruxelles|luxembourg|suisse|switzerland|gen[èe]ve|zurich|"
                           r"espagne|spain|madrid|italie|italy|milan|portugal|lisbon|allemagne|"
                           r"germany|munich|pays-bas|netherlands|amsterdam|london|londres|uk|"
                           r"united kingdom|irlande|ireland|pologne|poland|hongrie|hungary|"
                           r"turquie|turkey|istanbul|su[èe]de|sweden|slovenia)\b"),
    ('Amériques', r"\b(usa|united states|new york|miami|brazil|br[ée]sil|mexico|mexique)\b"),
    ('Asie / Pacifique', r"\b(china|chine|shanghai|hong kong|singapore|singapour|japan|japon|"
                         r"india|inde|australia)\b"),
]
GEO = [(name, re.compile(pat, re.I)) for name, pat in GEO]

AFRICA = re.compile(
    r'\b(maroc|morocco|marjane|risma|aradei|saham|banque centrale populaire|\bbcp\b|'
    r'attijari|\bocp\b|ciments de l.afrique|cimaf|tesca|groupe cdg|asma invest|'
    r'label.vie|cosumar|managem|akwa|ynna|addoha|alliances|s[ée]n[ée]gal|c[ôo]te d.ivoire|'
    r'abidjan|dakar|tunisie|tunisia|alg[ée]rie|algeria|cameroun|congo|gabon)\b', re.I)

CLASSES = [
    ('CONCURRENT', RIVAL,
     "Conseil, audit ou intégrateur EPM : pas un prospect. Utile en veille "
     "concurrentielle, en recrutement ou en partenariat."),
    ('HORS ICP', MEGA,
     "Groupe très au-dessus du plafond de 1 Md€ : hors ICP tel que défini dans le fichier."),
    ('AFRIQUE', AFRICA,
     "Périmètre Maroc / Afrique : décision d'ouvrir ou non ce marché, pas un tri de données."),
]


def perimetre(company, position):
    """Pays declare dans le poste ou le nom de societe ; France par defaut."""
    text = f'{company} {position}'
    for name, pattern in GEO:
        if pattern.search(text):
            return name
    return 'France présumée'


def classify(company, position):
    zone = perimetre(company, position)
    if zone != 'France présumée':
        return 'HORS FRANCE', zone
    for label, pattern, _ in CLASSES:
        if pattern.search(company):
            return label, zone
    return 'A QUALIFIER', zone


def read_sheet(path, name):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h or '') for h in rows[0]]
    return [dict(zip(head, r)) for r in rows[1:]]


def norm(url):
    u = str(url or '').strip().lower().rstrip('/')
    return re.sub(r'^https?://(www\.)?', '', re.sub(r'^[a-z]{2}\.linkedin', 'linkedin', u))


def main(src, dst):
    base = read_sheet(src, 'contacts base')
    val = read_sheet(src, 'contacts à valider')
    promoted = {norm(r['LinkedIn']) for r in val}
    retained = {str(r['Entreprise'] or '').strip().lower() for r in val}

    pool = []
    for b in base:
        pos, comp = str(b['Position'] or '').strip(), str(b['Company'] or '').strip()
        if not pos or not comp or NOISE.search(pos):
            continue
        if not EPM.search(pos) or norm(b['URL']) in promoted:
            continue
        pool.append((b, comp, pos))

    depth = collections.Counter(c for _, c, _ in pool)
    rows = []
    for b, comp, pos in pool:
        cls, zone = classify(comp, pos)
        rows.append({
            'Classement': cls,
            'Périmètre': zone,
            'Société': comp,
            'Contacts sur ce compte': depth[comp],
            'Compte déjà dans la liste': 'oui' if comp.lower() in retained else '',
            'Prénom': b['First Name'], 'Nom': b['Last Name'],
            'Poste (LinkedIn)': pos,
            'LinkedIn': b['URL'],
            'Connexion': str(b['Connected On'] or '')[:11],
        })
    order = {'A QUALIFIER': 0, 'HORS ICP': 1, 'CONCURRENT': 2,
             'HORS FRANCE': 3, 'AFRIQUE': 4}
    rows.sort(key=lambda r: (order[r['Classement']], -r['Contacts sur ce compte'],
                             r['Société'].lower()))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'gisement EPM'
    cols = list(rows[0])
    head_fill = PatternFill('solid', fgColor='E7E6E6')
    head_font = Font(name='Arial', size=9, bold=True, color='262626')
    rule = Border(bottom=Side('thin', color='262626'))
    tint = {'A QUALIFIER': 'EAF6EE', 'AFRIQUE': 'FFFBEA', 'HORS FRANCE': 'FFFBEA',
            'HORS ICP': 'F5F6F7', 'CONCURRENT': 'FDECEC'}

    for c, name in enumerate(cols, 1):
        cell = ws.cell(1, c, name)
        cell.fill, cell.font, cell.border = head_fill, head_font, rule
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    for r, row in enumerate(rows, 2):
        fill = PatternFill('solid', fgColor=tint[row['Classement']])
        for c, name in enumerate(cols, 1):
            cell = ws.cell(r, c, row[name])
            cell.font = Font(name='Arial', size=8.5, color='262626',
                             bold=(name == 'Société' and row['Contacts sur ce compte'] > 1))
            cell.fill = fill
            cell.alignment = Alignment(horizontal='left', vertical='center')
    for c, w in enumerate([15, 22, 34, 10, 12, 16, 20, 58, 46, 12], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(cols))}{len(rows) + 1}'
    ws.sheet_view.showGridLines = False

    leg = wb.create_sheet('lecture')
    leg.column_dimensions['A'].width = 16
    leg.column_dimensions['B'].width = 108
    counts = collections.Counter(r['Classement'] for r in rows)
    leg.append(['Classement', 'Ce que ça veut dire'])
    for label, _, meaning in CLASSES + [
        ('HORS FRANCE', None,
         "Pays déclaré dans l'intitulé du poste ou le nom de la société. "
         "Le pays n'est jamais déduit du nom de la personne : sur ce fichier ce "
         "serait à la fois faux et injuste, un DAF basé en France pouvant porter "
         "n'importe quel nom."),
        ('A QUALIFIER', None,
            "Le gisement : profil au signal EPM le plus fort, société ni concurrente, "
            "ni mega-cap, ni hors périmètre. À passer au même crible ICP que les 166.")]:
        leg.append([f'{label} ({counts.get(label, 0)})', meaning])
    for row in leg.iter_rows(min_row=1, max_row=leg.max_row):
        for cell in row:
            cell.font = Font(name='Arial', size=9, bold=cell.row == 1, color='262626')
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    leg.sheet_view.showGridLines = False

    wb.save(dst)
    print(f'wrote {dst}  ({len(rows)} profils, {len(depth)} sociétés)')
    for k, v in counts.most_common():
        print(f'   {k:<12} {v:>4}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
