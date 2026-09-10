"""Minimal OOXML spreadsheet patcher.

Edits an .xlsx in place at the XML level: everything the file already contains
(charts, images, conditional-formatting extensions, defined names) is carried
through byte-for-byte, which a full openpyxl round-trip would not do.

Only what a layout pass needs is modelled: the style tables, row heights,
column widths, sheet views and per-cell style indices.
"""

import re
import zipfile
from collections import OrderedDict

A_ORD = ord('A')


def col_letter(idx):
    """1 -> A, 27 -> AA."""
    s = ''
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(A_ORD + r) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - A_ORD + 1)
    return n


def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref)
    return int(m.group(2)), col_index(m.group(1))


def attrs_of(tag_xml):
    return OrderedDict(re.findall(r'([\w:]+)="([^"]*)"', tag_xml))


def render_attrs(attrs):
    return ''.join(f' {k}="{v}"' for k, v in attrs.items())


def children(xml, tag):
    """Top-level <tag> elements of a fragment, self-closing or paired."""
    pat = re.compile(r'<%s\b[^>]*/>|<%s\b[^>]*>.*?</%s>' % (tag, tag, tag), re.S)
    return pat.findall(xml)


class Styles:
    """The workbook style tables, addressable by index."""

    _BUILTIN = {0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
                9: '0%', 10: '0.00%', 49: '@'}

    def __init__(self, xml):
        self.xml = xml
        self.numfmts = OrderedDict(
            (code, int(i)) for i, code in
            re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/>', self._block('numFmts')))
        self.fonts = children(self._block('fonts'), 'font')
        self.fills = children(self._block('fills'), 'fill')
        self.borders = children(self._block('borders'), 'border')
        self.xfs = children(self._block('cellXfs'), 'xf')
        self.dxfs_xml = self._block('dxfs')

    def _block(self, tag):
        m = re.search(r'<%s\b[^>]*/>|<%s\b[^>]*>(.*?)</%s>' % (tag, tag, tag), self.xml, re.S)
        if m is None:
            return ''
        return m.group(1) or ''

    # -- component tables ------------------------------------------------
    def numfmt(self, code):
        """`code` is a raw format string; stored codes are XML-escaped."""
        for i, c in self._BUILTIN.items():
            if c == code:
                return i
        esc = (code.replace('&', '&amp;').replace('<', '&lt;')
               .replace('>', '&gt;').replace('"', '&quot;'))
        if esc in self.numfmts:
            return self.numfmts[esc]
        nid = max(list(self.numfmts.values()) + [999]) + 1
        self.numfmts[esc] = nid
        return nid

    def _intern(self, table, xml):
        if xml in table:
            return table.index(xml)
        table.append(xml)
        return len(table) - 1

    def font(self, name='Arial', sz=11, b=False, i=False, rgb=None, theme=None, tint=None):
        if rgb is not None:
            color = f'<color rgb="{rgb}"/>'
        elif theme is not None:
            color = f'<color theme="{theme}"' + (f' tint="{tint}"' if tint is not None else '') + '/>'
        else:
            color = '<color theme="1"/>'
        return self._intern(self.fonts,
                            f'<font><b val="{int(b)}"/><i val="{int(i)}"/><sz val="{sz}"/>'
                            f'{color}<name val="{name}"/></font>')

    def fill(self, rgb=None):
        if rgb is None:
            return self._intern(self.fills, '<fill><patternFill patternType="none"/></fill>')
        return self._intern(self.fills,
                            f'<fill><patternFill patternType="solid"><fgColor rgb="{rgb}"/>'
                            f'<bgColor rgb="{rgb}"/></patternFill></fill>')

    def border(self, left=None, right=None, top=None, bottom=None):
        def side(tag, spec):
            if spec is None:
                return f'<{tag}/>'
            style, rgb = spec
            return f'<{tag} style="{style}"><color rgb="{rgb}"/></{tag}>'
        return self._intern(self.borders,
                            '<border>' + side('left', left) + side('right', right)
                            + side('top', top) + side('bottom', bottom) + '<diagonal/></border>')

    def xf(self, font=0, fill=0, border=0, numfmt='General', halign=None,
           valign='center', wrap=False, indent=0, quote_prefix=False):
        nid = self.numfmt(numfmt)
        align = ''
        if halign or valign or wrap or indent:
            bits = ''
            if halign:
                bits += f' horizontal="{halign}"'
            if valign:
                bits += f' vertical="{valign}"'
            if wrap:
                bits += ' wrapText="1"'
            if indent:
                bits += f' indent="{indent}"'
            align = f'<alignment{bits}/>'
        head = (f'<xf numFmtId="{nid}" fontId="{font}" fillId="{fill}" borderId="{border}"'
                f' xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"'
                f' applyAlignment="1"' + (' quotePrefix="1"' if quote_prefix else '') + '>')
        return self._intern(self.xfs, head + align + '</xf>')

    # -- in-place recolouring -------------------------------------------
    def remap(self, table_name, mapping):
        """Apply literal substitutions to one style table."""
        table = {'fonts': self.fonts, 'fills': self.fills, 'borders': self.borders}[table_name]
        for i, xml in enumerate(table):
            for old, new in mapping.items():
                xml = xml.replace(old, new)
            table[i] = xml

    def remap_dxfs(self, mapping):
        for old, new in mapping.items():
            self.dxfs_xml = self.dxfs_xml.replace(old, new)

    def replace(self, table_name, index, xml):
        {'fonts': self.fonts, 'fills': self.fills, 'borders': self.borders}[table_name][index] = xml

    # -- serialisation ---------------------------------------------------
    def serialize(self):
        out = self.xml

        def sub_block(text, tag, body, count):
            pat = re.compile(r'<%s\b[^>]*/>|<%s\b[^>]*>.*?</%s>' % (tag, tag, tag), re.S)
            return pat.sub(lambda m: f'<{tag} count="{count}">{body}</{tag}>', text, count=1)

        nf = ''.join(f'<numFmt numFmtId="{i}" formatCode="{c}"/>'
                     for c, i in self.numfmts.items())
        if '<numFmts' in out:
            out = sub_block(out, 'numFmts', nf, len(self.numfmts))
        else:
            out = out.replace('<fonts', f'<numFmts count="{len(self.numfmts)}">{nf}</numFmts><fonts', 1)
        out = sub_block(out, 'fonts', ''.join(self.fonts), len(self.fonts))
        out = sub_block(out, 'fills', ''.join(self.fills), len(self.fills))
        out = sub_block(out, 'borders', ''.join(self.borders), len(self.borders))
        out = sub_block(out, 'cellXfs', ''.join(self.xfs), len(self.xfs))
        out = sub_block(out, 'dxfs', self.dxfs_xml,
                        len(children(self.dxfs_xml, 'dxf')))
        return out


class SharedStrings:
    def __init__(self, xml):
        self.xml = xml
        self.items = children(re.search(r'<sst\b[^>]*>(.*)</sst>', xml, re.S).group(1), 'si')
        self.index = {s: i for i, s in enumerate(self.items)}

    def add(self, text):
        esc = (text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))
        si = f'<si><t xml:space="preserve">{esc}</t></si>'
        if si in self.index:
            return self.index[si]
        self.items.append(si)
        self.index[si] = len(self.items) - 1
        return len(self.items) - 1

    def serialize(self):
        head = re.match(r'.*?<sst\b[^>]*>', self.xml, re.S).group(0)
        head = re.sub(r'count="\d+"', f'count="{len(self.items)}"', head)
        head = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{len(self.items)}"', head)
        return head + ''.join(self.items) + '</sst>'


class Cell:
    __slots__ = ('attrs', 'inner')

    def __init__(self, attrs, inner):
        self.attrs, self.inner = attrs, inner

    def render(self):
        if self.inner:
            return f'<c{render_attrs(self.attrs)}>{self.inner}</c>'
        return f'<c{render_attrs(self.attrs)}/>'


class Row:
    __slots__ = ('attrs', 'cells')

    def __init__(self, attrs, cells):
        self.attrs, self.cells = attrs, cells

    def render(self):
        if not self.cells:
            return f'<row{render_attrs(self.attrs)}/>'
        body = ''.join(self.cells[c].render() for c in sorted(self.cells))
        return f'<row{render_attrs(self.attrs)}>{body}</row>'


class Sheet:
    def __init__(self, xml):
        m = re.search(r'(.*?)<sheetData\b[^>]*>(.*?)</sheetData>(.*)$', xml, re.S)
        if m is None:
            m = re.search(r'(.*?)<sheetData\b[^>]*/>(.*)$', xml, re.S)
            self.head, self.tail = m.group(1), m.group(2)
            body = ''
        else:
            self.head, body, self.tail = m.groups()
        self.rows = OrderedDict()
        for rx in children(body, 'row'):
            attrs = attrs_of(re.match(r'<row\b[^>]*?/?>', rx).group(0))
            cells = OrderedDict()
            inner = re.match(r'<row\b[^>]*>(.*)</row>$', rx, re.S)
            if inner:
                for cx in children(inner.group(1), 'c'):
                    ca = attrs_of(re.match(r'<c\b[^>]*?/?>', cx).group(0))
                    ci = re.match(r'<c\b[^>]*>(.*)</c>$', cx, re.S)
                    _, col = split_ref(ca['r'])
                    cells[col] = Cell(ca, ci.group(1) if ci else '')
            self.rows[int(attrs['r'])] = Row(attrs, cells)

    # -- rows / cells ----------------------------------------------------
    def row(self, r):
        if r not in self.rows:
            self.rows[r] = Row(OrderedDict([('r', str(r))]), OrderedDict())
        return self.rows[r]

    def cell(self, r, c):
        row = self.row(r)
        if c not in row.cells:
            row.cells[c] = Cell(OrderedDict([('r', f'{col_letter(c)}{r}')]), '')
        return row.cells[c]

    def has_cell(self, r, c):
        return r in self.rows and c in self.rows[r].cells

    def set_style(self, r, c, xf):
        self.cell(r, c).attrs['s'] = str(xf)

    def style_range(self, r1, c1, r2, c2, xf):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self.set_style(r, c, xf)

    def set_text(self, r, c, sst_index):
        cell = self.cell(r, c)
        cell.attrs['t'] = 's'
        cell.inner = f'<v>{sst_index}</v>'

    def move(self, src, dst):
        """Relocate a cell's value/formula; both cells keep their own style."""
        sr, sc = split_ref(src)
        dr, dc = split_ref(dst)
        if not self.has_cell(sr, sc):
            return
        s = self.rows[sr].cells[sc]
        d = self.cell(dr, dc)
        d.inner = s.inner
        for k in ('t', 'cm', 'vm', 'ph'):
            if k in s.attrs:
                d.attrs[k] = s.attrs[k]
            else:
                d.attrs.pop(k, None)
        self.clear_value(sr, sc)

    def clear_value(self, r, c):
        if self.has_cell(r, c):
            cell = self.rows[r].cells[c]
            cell.inner = ''
            for k in ('t', 'cm', 'vm', 'ph'):
                cell.attrs.pop(k, None)

    def row_height(self, r, height):
        a = self.row(r).attrs
        a['ht'] = str(height)
        a['customHeight'] = '1'

    # -- sheet chrome ----------------------------------------------------
    def set_view(self, gridlines=False, headers=False, zoom=None, freeze=None,
                 tab_selected=None):
        attrs = ['topLeftCell="A1"', 'workbookViewId="0"']
        if not gridlines:
            attrs.insert(0, 'showGridLines="0"')
        if not headers:
            attrs.insert(0, 'showRowColHeaders="0"')
        if zoom:
            attrs.append(f'zoomScale="{zoom}"')
            attrs.append(f'zoomScaleNormal="{zoom}"')
        if tab_selected:
            attrs.append('tabSelected="1"')
        inner = ''
        if freeze:
            fr, fc = split_ref(freeze)
            inner = (f'<pane xSplit="{fc - 1}" ySplit="{fr - 1}" topLeftCell="{freeze}"'
                     f' activePane="bottomRight" state="frozen"/>'
                     f'<selection pane="bottomRight" activeCell="{freeze}" sqref="{freeze}"/>')
        view = f'<sheetViews><sheetView {" ".join(attrs)}>{inner}</sheetView></sheetViews>'
        self.head = re.sub(r'<sheetViews\b[^>]*/>|<sheetViews\b.*?</sheetViews>',
                           view, self.head, count=1, flags=re.S)

    def column_width(self, first, last, width):
        """Widen or narrow a column span, splitting existing <col> runs as needed."""
        self._edit_columns(first, last, lambda a: a.update(
            {'width': str(width), 'customWidth': '1'}))

    def hide_columns(self, first, last, default_width=9.140625, style=None):
        """Mark a column span hidden, splitting existing <col> runs as needed."""
        self._edit_columns(first, last, lambda a: a.update({'hidden': '1'}),
                           default_width=default_width, style=style)

    def _edit_columns(self, first, last, apply, default_width=9.140625, style=None):
        m = re.search(r'<cols\b[^>]*>(.*?)</cols>', self.head, re.S)
        cols = children(m.group(1), 'col') if m else []
        spans = []
        for cx in cols:
            a = attrs_of(cx)
            spans.append([int(a.pop('min')), int(a.pop('max')), a])
        out = []
        for lo, hi, a in spans:
            for s_lo, s_hi, inside in _split(lo, hi, first, last):
                b = OrderedDict(a)
                if inside:
                    apply(b)
                out.append([s_lo, s_hi, b])
        covered = set()
        for lo, hi, _ in out:
            covered.update(range(lo, hi + 1))
        for c in range(first, last + 1):
            if c not in covered:
                b = OrderedDict()
                if style is not None:
                    b['style'] = str(style)
                b['width'] = str(default_width)
                b['customWidth'] = '1'
                apply(b)
                out.append([c, c, b])
        out.sort()
        body = ''.join(f'<col min="{lo}" max="{hi}"{render_attrs(a)}/>' for lo, hi, a in out)
        new = f'<cols>{body}</cols>'
        if m:
            self.head = self.head[:m.start()] + new + self.head[m.end():]
        else:
            self.head = re.sub(r'(<sheetFormatPr\b[^>]*/>)', r'\1' + new, self.head, count=1)

    def serialize(self):
        body = ''.join(self.rows[r].render() for r in sorted(self.rows))
        return f'{self.head}<sheetData>{body}</sheetData>{self.tail}'


def _split(lo, hi, first, last):
    """Cut [lo,hi] against [first,last]; yields (lo, hi, is_inside)."""
    parts = []
    if hi < first or lo > last:
        return [(lo, hi, False)]
    if lo < first:
        parts.append((lo, first - 1, False))
    parts.append((max(lo, first), min(hi, last), True))
    if hi > last:
        parts.append((last + 1, hi, False))
    return parts


class Workbook:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.names = [i.filename for i in z.infolist()]
            self.parts = {i.filename: z.read(i.filename) for i in z.infolist()}

    def text(self, name):
        return self.parts[name].decode('utf-8')

    def put(self, name, text):
        self.parts[name] = text.encode('utf-8')

    def sheet_part(self, sheet_name):
        wb = self.text('xl/workbook.xml')
        rels = self.text('xl/_rels/workbook.xml.rels')
        relmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        for sx in children(wb, 'sheet'):
            a = attrs_of(sx)
            if a.get('name') == sheet_name:
                return 'xl/' + relmap[a['r:id']]
        raise KeyError(sheet_name)

    def save(self, path):
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            for name in self.names:
                z.writestr(name, self.parts[name])
