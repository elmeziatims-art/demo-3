#!/usr/bin/env python3
"""Align COCKPIT.xlsx onto the CAD_PIL.xlsx page-layout system.

CAD_PIL (sheets Cadrage / Pilotage) defines the house layout; COCKPIT was built
on a different one. This rewrites COCKPIT's three visible sheets to CAD_PIL's:
same palette, type scale, title band, KPI band, section headings, table rules,
row rhythm and sheet chrome. Values, formulas, charts, images and conditional
formatting are untouched.

    python3 tools/harmonize_cockpit.py COCKPIT.xlsx COCKPIT_aligned.xlsx
"""

import re
import sys

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from xlsxpatch import Sheet, SharedStrings, Styles, Workbook, split_ref  # noqa: E402

# ---------------------------------------------------------------- tokens
# Read off CAD_PIL.xlsx; every colour below appears in that workbook.
BRAND = 'FF007AC3'      # rule under the title, active-scenario chip, key figures
CANVAS = 'FFF5F6F7'     # page background
BAND = 'FFE7E6E6'       # KPI strip, table headings, total rows
SURFACE = 'FFFFFFFF'    # table body
ZEBRA = 'FFE8F1F9'      # grouped-row tint
INPUT_BG = 'FFFFFFCC'   # editable cell
RULE = 'FFD5D7DA'       # section rule
HAIRLINE = 'FFEDEEF0'   # between body rows
STRONG = 'FF262626'     # under a table heading; also body text
MUTED = 'FF6B7075'      # captions
POSITIVE = 'FF5F8A17'
NEGATIVE = 'FFC41822'
WARN_BG = 'FFF0EDE4'
TITLE_THEME, TITLE_TINT = 4, '-0.5'   # theme4 at -50% = the CAD_PIL navy

# CAD_PIL number formats (note the en dash for "nothing to show").
F_EUR = '#,##0" €";\\-#,##0" €";"–"'
F_PCT = '0.0%;\\-0.0%;"–"'
F_INT = '#,##0;\\-#,##0;"–"'
F_ARROW_PCT = '"▲ "0.0%;"▼ "0.0%;"–"'
F_ARROW_PT = '"▲ "0.00" pt";"▼ "0.00" pt";"–"'
F_FREE_SEATS = '#,##0" places libres";\\-#,##0" places libres";"–"'

# CAD_PIL row rhythm.
H_SPACER_TOP, H_TITLE, H_RULE = 14.25, 46.5, 3.75
H_CAPTION, H_META, H_GAP = 15.75, 19.5, 8.25
H_KPI_TOP, H_KPI_LABEL, H_KPI_VALUE, H_KPI_CAPTION = 3.75, 15.75, 29.25, 15.75
H_SECTION, H_SUBTITLE, H_HEADER, H_BODY = 30.0, 14.25, 30.0, 18.0
H_TIGHT = 6.0


def build_palette(st):
    """Register the CAD_PIL tokens and return the cell styles the layout needs."""
    f = dict(
        title=st.font('Montserrat ExtraBold', 20, b=True, theme=TITLE_THEME, tint=TITLE_TINT),
        caption=st.font('Arial', 9, i=True, rgb=MUTED),
        small=st.font('Arial', 8, i=True, rgb=MUTED),
        meta_label=st.font('Arial', 12, b=True, rgb=STRONG),
        meta_value=st.font('Arial', 10, b=True, rgb=STRONG),
        chip=st.font('Arial', 10, b=True, rgb=BRAND),
        kpi_label=st.font('Arial', 7.5, b=True, rgb=STRONG),
        kpi_value=st.font('Arial', 20, rgb=STRONG),
        section=st.font('Arial', 10, b=True, rgb=STRONG),
        blank=st.font('Aptos Narrow', 11, theme=1),
        on_brand=st.font('Aptos Narrow', 11, theme=1),
    )
    fill = dict(none=st.fill(None), canvas=st.fill(CANVAS), band=st.fill(BAND),
                surface=st.fill(SURFACE), brand=st.fill(BRAND), chip=st.fill(INPUT_BG))
    bd = dict(none=st.border(),
              under_rule=st.border(bottom=('thin', RULE)),
              chip=st.border(left=('thin', BRAND), right=('thin', BRAND),
                             top=('thin', BRAND), bottom=('thin', BRAND)))

    s = {}
    s['band_white'] = st.xf(font=f['blank'], fill=fill['none'], border=bd['none'])
    s['title'] = st.xf(font=f['title'], fill=fill['none'], border=bd['none'],
                       halign='centerContinuous')
    s['rule'] = st.xf(font=f['on_brand'], fill=fill['brand'], border=bd['none'])
    s['caption'] = st.xf(font=f['caption'], fill=fill['canvas'], border=bd['none'],
                         halign='centerContinuous')
    s['canvas'] = st.xf(font=f['blank'], fill=fill['canvas'], border=bd['none'])
    s['meta_label'] = st.xf(font=f['meta_label'], fill=fill['canvas'], border=bd['none'],
                            halign='left', indent=1)
    s['meta_small'] = st.xf(font=f['small'], fill=fill['canvas'], border=bd['none'],
                            halign='left', indent=1)
    s['meta_value'] = st.xf(font=f['meta_value'], fill=fill['canvas'], border=bd['none'],
                            halign='left', indent=1)
    s['chip'] = st.xf(font=f['chip'], fill=fill['chip'], border=bd['chip'], halign='center')
    s['kpi_top'] = st.xf(font=f['blank'], fill=fill['band'], border=bd['none'])
    s['kpi_label'] = st.xf(font=f['kpi_label'], fill=fill['band'], border=bd['under_rule'],
                           halign='centerContinuous')
    s['section'] = st.xf(font=f['section'], fill=fill['canvas'], border=bd['none'],
                         halign='left', indent=1)
    s['subtitle'] = st.xf(font=f['small'], fill=fill['canvas'], border=bd['none'],
                          halign='left', indent=1)
    for key, code in (('eur', F_EUR), ('pct', F_PCT), ('int', F_INT)):
        s['kpi_value_' + key] = st.xf(font=f['kpi_value'], fill=fill['surface'],
                                      border=bd['none'], numfmt=code,
                                      halign='centerContinuous')
    for key, code in (('pct', F_ARROW_PCT), ('pt', F_ARROW_PT), ('seats', F_FREE_SEATS)):
        s['kpi_caption_' + key] = st.xf(font=f['caption'], fill=fill['surface'],
                                        border=bd['under_rule'], numfmt=code,
                                        halign='centerContinuous')
    s['kpi_value_blank'] = st.xf(font=f['kpi_value'], fill=fill['surface'],
                                 border=bd['none'], halign='centerContinuous')
    s['kpi_caption_blank'] = st.xf(font=f['caption'], fill=fill['surface'],
                                   border=bd['under_rule'], halign='centerContinuous')
    # the strip runs the full table width; its outer columns carry no text
    s['kpi_label_edge'] = st.xf(font=f['blank'], fill=fill['band'], border=bd['under_rule'])
    s['kpi_value_edge'] = st.xf(font=f['blank'], fill=fill['surface'], border=bd['none'])
    s['kpi_caption_edge'] = st.xf(font=f['blank'], fill=fill['surface'],
                                  border=bd['under_rule'])
    return s


def recolour(st):
    """Swap COCKPIT's own palette and type for CAD_PIL's, entry by entry.

    Every style the sheets already reference keeps its index, so nothing has to
    be re-pointed: the same cells simply render in the house colours.
    """
    st.remap('fonts', {
        'FF202733': STRONG, 'FF69778B': MUTED, 'FF526071': MUTED,
        'FF7C8798': MUTED, 'FFB8C6DA': MUTED,
        'FF1E9E89': POSITIVE, 'FF00B050': POSITIVE, 'FFD64545': NEGATIVE,
        'FF2A78D6': BRAND, 'FFF3F6FA': CANVAS,
        '"Fira Sans Medium"': '"Arial"', '"Segoe UI"': '"Arial"',
    })
    # White-on-slate no longer works once the slate becomes a light band.
    st.replace('fonts', 32, f'<font><b/><sz val="7.5"/><color rgb="{STRONG}"/>'
                            f'<name val="Arial"/></font>')
    # The two banner fonts become the CAD_PIL title face.
    for idx in (21, 35):
        st.replace('fonts', idx,
                   f'<font><b/><sz val="20"/><color theme="{TITLE_THEME}" '
                   f'tint="{TITLE_TINT}"/><name val="Montserrat ExtraBold"/></font>')
    st.replace('fonts', 36, f'<font><i/><sz val="9"/><color rgb="{MUTED}"/>'
                            f'<name val="Arial"/></font>')
    st.replace('fonts', 0, '<font><sz val="11"/><color theme="1"/>'
                           '<name val="Aptos Narrow"/><scheme val="minor"/></font>')

    st.remap('fills', {
        'FFF3F6FA': CANVAS, 'FFEAF0F6': CANVAS,
        'FF526071': BAND, 'FFD9E2EF': BAND,
        'FF172033': SURFACE,          # the dark drill banner turns into the white band
        'FFEAF2FC': ZEBRA, 'FFFFF1E8': WARN_BG,
    })

    st.remap('borders', {
        'FFE9EDF3': HAIRLINE, 'FFB8CFEC': RULE,
        "style='medium'": "style='thin'", 'style="medium"': 'style="thin"',
        'FF526071': STRONG,
    })
    # Table headings: CAD_PIL rules them above and below, never boxes them.
    st.replace('borders', 21,
               f'<border><left/><right/><top style="thin"><color rgb="{RULE}"/></top>'
               f'<bottom style="thin"><color rgb="{STRONG}"/></bottom><diagonal/></border>')
    # The old banner underline is replaced by the brand rule row.
    st.replace('borders', 25, '<border><left/><right/><top/><bottom/><diagonal/></border>')

    st.remap_dxfs({
        'FF202733': STRONG, 'FFE9EDF3': HAIRLINE,
        'FF526071': STRONG, "style='medium'": "style='thin'",
        'FFEAF2FC': ZEBRA, 'FFD9E2EF': BAND, 'FFFFF1E8': WARN_BG,
        'FF1E9E89': POSITIVE, 'FF00B050': POSITIVE, 'FFD64545': NEGATIVE,
    })


# CAD_PIL's three-stop scale, and a brand-blue bar in place of COCKPIT's.
CF_COLOURS = {
    'FFF7CFCF': 'FFF6C9CC',     # low
    'FFFDF0CE': 'FFF7F7F5',     # mid
    'FFCDE9DF': 'FFDCEBC0',     # high
    'FFB8CFEC': 'FFA6D0EA',     # data bar
}


def recolour_scales(xml):
    for old, new in CF_COLOURS.items():
        xml = re.sub(r'(<(?:x14:)?color rgb=")%s(")' % old, r'\g<1>%s\g<2>' % new, xml)
    return xml


def banner(sheet, s, sst, title_from, title_to, caption_from, caption_to,
           first_col, last_col, title_span_first, title_span_last):
    """The CAD_PIL masthead: white band, centred title, brand rule, caption."""
    sheet.style_range(1, first_col, 2, last_col, s['band_white'])
    if title_from != title_to:
        sheet.move(title_from, title_to)
    for c in range(title_span_first, title_span_last + 1):
        sheet.set_style(2, c, s['title'])
    sheet.style_range(3, first_col, 3, last_col, s['rule'])
    if caption_from:
        if caption_from != caption_to:
            sheet.move(caption_from, caption_to)
        row, _ = split_ref(caption_to)
        sheet.style_range(row, first_col, row, last_col, s['canvas'])
        for c in range(title_span_first, title_span_last + 1):
            sheet.set_style(row, c, s['caption'])
    for r, h in ((1, H_SPACER_TOP), (2, H_TITLE), (3, H_RULE)):
        sheet.row_height(r, h)


def section(sheet, s, sst, row, col, text, first_col, last_col, subtitle=None):
    sheet.style_range(row, first_col, row, last_col, s['canvas'])
    sheet.set_style(row, col, s['section'])
    sheet.set_text(row, col, sst.add(text))
    sheet.row_height(row, H_SECTION)
    if subtitle is not None:
        sheet.style_range(row + 1, first_col, row + 1, last_col, s['canvas'])
        sheet.set_style(row + 1, col, s['subtitle'])
        sheet.set_text(row + 1, col, sst.add(subtitle))
        sheet.row_height(row + 1, H_SUBTITLE)


# ------------------------------------------------------------------ sheets
def do_cockpit(sheet, s, sst):
    FIRST, LAST = 1, 17                      # A .. Q is the visible width
    TAB_L, TAB_R = 2, 17                     # the detail table spans B .. Q
    banner(sheet, s, sst, 'G2', 'C2', 'G3', 'C4',
           FIRST, LAST, title_span_first=TAB_L, title_span_last=TAB_R)
    sheet.row_height(4, H_CAPTION)

    # -- meta strip: active scenario, version, entity
    sheet.style_range(5, FIRST, 5, LAST, s['canvas'])
    sheet.set_style(5, 3, s['meta_label']); sheet.set_text(5, 3, sst.add('Scénario'))
    sheet.set_style(5, 6, s['chip'])
    sheet.set_style(5, 10, s['meta_small']); sheet.set_text(5, 10, sst.add('Version'))
    sheet.set_style(5, 11, s['meta_value'])
    sheet.set_style(5, 13, s['meta_small']); sheet.set_text(5, 13, sst.add('Entité'))
    sheet.set_style(5, 14, s['meta_value'])
    sheet.row_height(5, H_META)
    sheet.style_range(6, FIRST, 6, LAST, s['canvas'])
    sheet.row_height(6, H_GAP)

    # -- KPI band: label / figure / variation, one block per indicator
    blocks = [
        (3, 6, 'Chiffre d\'affaires', 'eur', 'pct'),
        (7, 8, 'EBITDA', 'eur', 'pct'),
        (9, 10, 'Marge EBITDA', 'pct', 'pt'),
        (11, 12, 'Inscrits (nouveaux)', 'int', 'pct'),
        (13, 14, 'Coût d\'acquisition', 'eur', 'pct'),
        (15, 16, 'Remplissage moyen', 'pct', 'seats'),
    ]
    sheet.style_range(7, FIRST, 7, LAST, s['canvas'])
    sheet.style_range(7, TAB_L, 7, TAB_R, s['kpi_top'])
    for r, key in ((8, 'kpi_label_edge'), (9, 'kpi_value_edge'), (10, 'kpi_caption_edge')):
        sheet.style_range(r, FIRST, r, LAST, s['canvas'])
        sheet.style_range(r, TAB_L, r, TAB_R, s[key])
    for c0, c1, label, vfmt, cfmt in blocks:
        for c in range(c0, c1 + 1):
            sheet.set_style(8, c, s['kpi_label'])
            sheet.set_style(9, c, s['kpi_value_' + vfmt])
            sheet.set_style(10, c, s['kpi_caption_' + cfmt])
        sheet.set_text(8, c0, sst.add(label))
    # a seventh, empty block was left over past the table's right edge
    sheet.style_range(7, 18, 10, 18, s['canvas'])
    for r, h in ((7, H_KPI_TOP), (8, H_KPI_LABEL), (9, H_KPI_VALUE), (10, H_KPI_CAPTION)):
        sheet.row_height(r, h)
    sheet.style_range(11, FIRST, 11, LAST, s['canvas'])
    sheet.row_height(11, H_TIGHT)

    # -- chart captions
    sheet.style_range(12, FIRST, 12, LAST, s['canvas'])
    sheet.clear_value(12, 3)
    sheet.clear_value(12, 11)
    for col, text in ((2, "1 ·  Pont d'EBITDA 2025 → 2026"),      # each sits on the
                      (9, '2 ·  Marge EBITDA par marque'),        # left edge of the
                      (14, "3 ·  Dépenses d'acquisition et inscrits (base 100)")):
        sheet.set_style(12, col, s['section'])                    # chart it names
        sheet.set_text(12, col, sst.add(text))
    sheet.row_height(12, H_SECTION)
    sheet.style_range(13, FIRST, 13, LAST, s['canvas'])
    sheet.row_height(13, H_TIGHT)

    # -- detail table
    sheet.style_range(34, FIRST, 34, LAST, s['canvas'])
    sheet.row_height(34, H_TIGHT)
    section(sheet, s, sst, 35, TAB_L, '4 ·  Le détail par marque et par campus',
            FIRST, LAST,
            subtitle='niveau 2 · groupe   ·   niveau 3 · marque   ·   niveau 4 · campus')
    for col, text in ((2, 'Niv.'), (3, 'Périmètre'), (4, 'Programme'), (5, 'Modalité')):
        sheet.set_text(37, col, sst.add(text))
    sheet.row_height(37, H_HEADER)
    for r in range(38, 58):
        sheet.row_height(r, H_BODY)

    sheet.set_view(zoom=95, freeze='C11', tab_selected=True)
    sheet.column_width(3, 3, 34)    # full campus names, no ellipsis
    sheet.hide_columns(29, 29)      # spacer left over between working blocks
    sheet.hide_columns(49, 64)      # chart feed: base-100 indices


def do_drill_one(sheet, s, sst):
    FIRST, LAST = 1, 17
    banner(sheet, s, sst, 'C3', 'B2', 'C4', 'B4',
           FIRST, LAST, title_span_first=2, title_span_last=16)
    sheet.row_height(4, H_CAPTION)
    sheet.style_range(5, FIRST, 5, LAST, s['canvas'])
    sheet.row_height(5, H_TIGHT)
    sheet.style_range(9, FIRST, 9, LAST, s['canvas'])
    sheet.row_height(9, H_TIGHT)

    section(sheet, s, sst, 10, 2, '1 ·  Le pont, chiffre par chiffre', FIRST, LAST)
    sheet.row_height(11, H_HEADER)
    for r in range(12, 22):
        sheet.row_height(r, H_BODY)
    section(sheet, s, sst, 41, 2, '2 ·  Le détail par campus', FIRST, LAST)
    sheet.row_height(42, H_HEADER)
    for r in range(43, 74):
        sheet.row_height(r, H_BODY)

    sheet.set_view(zoom=95, freeze='B5')
    sheet.hide_columns(17, 82)      # gutter, effect workings and chart feed


def do_drill_two(sheet, s, sst):
    FIRST, LAST = 1, 8
    banner(sheet, s, sst, 'D2', 'B2', None, None,
           FIRST, LAST, title_span_first=2, title_span_last=8)
    sheet.style_range(4, FIRST, 4, LAST, s['canvas'])
    sheet.row_height(4, H_GAP)

    section(sheet, s, sst, 32, 2, "1 ·  Le compte d'exploitation, poste par poste",
            FIRST, LAST)
    sheet.row_height(33, H_HEADER)
    for r in range(34, 48):
        sheet.row_height(r, H_BODY)

    sheet.set_view(zoom=95, freeze='B4')
    sheet.hide_columns(9, 25)       # lookup table and chart feed


# ----------------------------------------------------------------- drawings
LOGO_X, LOGO_Y, LOGO_W, LOGO_H = 419100, 90488, 1514475, 562518   # CAD_PIL's masthead


def place_logos(xml):
    """Pin the logo where CAD_PIL puts it, at CAD_PIL's size."""
    def fix(m):
        block = m.group(0)
        if '<xdr:pic>' not in block:
            return block
        block = re.sub(r'<xdr:from>.*?</xdr:from>',
                       f'<xdr:from><xdr:col>0</xdr:col><xdr:colOff>{LOGO_X}</xdr:colOff>'
                       f'<xdr:row>0</xdr:row><xdr:rowOff>{LOGO_Y}</xdr:rowOff></xdr:from>',
                       block, flags=re.S)
        block = re.sub(r'<xdr:to>.*?</xdr:to>',
                       f'<xdr:ext cx="{LOGO_W}" cy="{LOGO_H}"/>', block, flags=re.S)
        block = re.sub(r'<a:off[^/]*/>', f'<a:off x="{LOGO_X}" y="{LOGO_Y}"/>', block)
        block = re.sub(r'<a:ext cx="\d+" cy="\d+"/>',
                       f'<a:ext cx="{LOGO_W}" cy="{LOGO_H}"/>', block)
        block = re.sub(r'^<xdr:twoCellAnchor[^>]*>', '<xdr:oneCellAnchor>', block)
        return block.replace('</xdr:twoCellAnchor>', '</xdr:oneCellAnchor>')
    return re.sub(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', fix, xml, flags=re.S)


def align_charts(xml, from_row, from_off, to_row, to_off):
    """Give every chart on a sheet the same top and bottom edge."""
    def fix(m):
        block = m.group(0)
        if '<xdr:pic>' in block:
            return block
        block = re.sub(r'(<xdr:from>.*?<xdr:row>)\d+(</xdr:row><xdr:rowOff>)\d+',
                       lambda x: f'{x.group(1)}{from_row}{x.group(2)}{from_off}',
                       block, flags=re.S)
        block = re.sub(r'(<xdr:to>.*?<xdr:row>)\d+(</xdr:row><xdr:rowOff>)\d+',
                       lambda x: f'{x.group(1)}{to_row}{x.group(2)}{to_off}',
                       block, flags=re.S)
        return block
    return re.sub(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', fix, xml, flags=re.S)


def main(src, dst):
    wb = Workbook(src)
    st = Styles(wb.text('xl/styles.xml'))
    sst = SharedStrings(wb.text('xl/sharedStrings.xml'))
    recolour(st)
    s = build_palette(st)

    for name, fn in (('Cockpit', do_cockpit),
                     ('Drill EBITDA 1', do_drill_one),
                     ('Drill EBITDA 2', do_drill_two)):
        part = wb.sheet_part(name)
        sheet = Sheet(wb.text(part))
        fn(sheet, s, sst)
        wb.put(part, recolour_scales(sheet.serialize()))

    wb.put('xl/styles.xml', st.serialize())
    wb.put('xl/sharedStrings.xml', sst.serialize())

    wb.put('xl/drawings/drawing1.xml',
           align_charts(place_logos(wb.text('xl/drawings/drawing1.xml')), 13, 0, 31, 76200))
    for part in ('xl/drawings/drawing2.xml', 'xl/drawings/drawing3.xml'):
        wb.put(part, place_logos(wb.text(part)))

    # CAD_PIL widens the tab strip so every sheet name stays readable.
    book = wb.text('xl/workbook.xml')
    if 'tabRatio' not in book:
        book = book.replace('<workbookView ', '<workbookView tabRatio="1000" ', 1)
    wb.put('xl/workbook.xml', book)

    wb.save(dst)
    print(f'wrote {dst}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
