"""Escritura de partes en las casillas del Excel original (una pestaña por proyecto).

Edita directamente el XML del .xlsx para conservar formatos, fórmulas y pestañas
existentes. Solo toca las columnas de datos (fecha, mes, tarea y horas).
"""
from __future__ import annotations
from datetime import date
from html import escape
import os
import posixpath
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from xlsx_io import NS, RN, BASE_DATE

MONTHS = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
          'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
ROW_RE = re.compile(r'<row\b[^>]*?(?:/>|>.*?</row>)', re.S)
CELL_RE = re.compile(r'<c\b[^>]*?(?:/>|>.*?</c>)', re.S)
SHEET_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
CALC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'


def minutes_to_excel_hours(minutes):
    """Convención del libro: 8.15 = 8h15, 8.5 = 8h30, 8.45 = 8h45."""
    if minutes is None:
        return None
    h, m = divmod(int(minutes), 60)
    if m == 30:
        return h + 0.5
    return round(h + m / 100, 2)


def col_letter(idx):
    s = ''
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(ref):
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n - 1


def _attr(tag, name):
    m = re.search(rf'\s{name}="([^"]*)"', tag)
    return m.group(1) if m else None


def _num(v):
    return repr(float(v)) if not float(v).is_integer() else str(int(v))


class _Sheet:
    def __init__(self, xml, shared):
        self.xml = xml
        self.shared = shared
        start = xml.index('<sheetData')
        head_end = xml.index('>', start) + 1
        if xml[head_end - 2] == '/':  # <sheetData/>
            self.prefix, self.suffix = xml[:start] + '<sheetData>', '</sheetData>' + xml[head_end:]
            body = ''
        else:
            end = xml.index('</sheetData>')
            self.prefix, self.suffix = xml[:head_end], xml[end:]
            body = xml[head_end:end]
        self.rows = {}
        for m in ROW_RE.finditer(body):
            r = int(_attr(m.group(0)[:m.group(0).index('>') + 1], 'r'))
            self.rows[r] = m.group(0)

    # --- lectura -------------------------------------------------------
    def cells(self, rn):
        row = self.rows.get(rn)
        if not row:
            return {}
        out = {}
        for c in CELL_RE.findall(row):
            tag = c[:c.index('>') + 1]
            out[col_index(_attr(tag, 'r'))] = c
        return out

    def value(self, cell):
        if cell is None:
            return None
        tag = cell[:cell.index('>') + 1]
        typ = _attr(tag, 't')
        if typ == 'inlineStr':
            return ''.join(re.findall(r'<t[^>]*>(.*?)</t>', cell, re.S)) or None
        v = re.search(r'<v>(.*?)</v>', cell, re.S)
        if not v:
            return None
        if typ == 's':
            return self.shared[int(v.group(1))]
        if typ in ('str', 'e', 'b'):
            return v.group(1)
        try:
            return float(v.group(1))
        except ValueError:
            return v.group(1)

    def layout(self):
        """Devuelve (columnas, tipo_mes, estilos) según las filas de datos existentes."""
        three = four = 0
        month_kind = 'name'
        last = None
        for rn in sorted(self.rows):
            c = self.cells(rn)
            a, b, cc = (self.value(c.get(i)) for i in range(3))
            if not isinstance(a, float):
                continue
            if isinstance(b, str) and isinstance(cc, float):
                three += 1
            elif b is not None or isinstance(cc, str):
                four += 1
                if isinstance(b, float):
                    month_kind = 'number'
                elif isinstance(b, str):
                    month_kind = 'name'
            last = rn
        cols = ('fecha', 'desc', 'horas') if three > four else ('fecha', 'mes', 'desc', 'horas')
        styles = {}
        if last:
            for i, cell in self.cells(last).items():
                if i < len(cols):
                    s = _attr(cell[:cell.index('>') + 1], 's')
                    if s:
                        styles[i] = s
        return cols, month_kind, styles

    def last_data_row(self, ncols):
        for rn in sorted(self.rows, reverse=True):
            c = self.cells(rn)
            if any(self.value(c.get(i)) not in (None, '') for i in range(ncols)):
                return rn
        return 0

    # --- escritura ----------------------------------------------------
    def set_cells(self, rn, values, styles):
        """values: {col: valor|None}. None deja la casilla vacía conservando su formato."""
        cells = self.cells(rn)
        for col, val in values.items():
            ref = f'{col_letter(col)}{rn}'
            old = cells.get(col)
            s = _attr(old[:old.index('>') + 1], 's') if old else None
            if styles.get(col) and (old is None or self.value(old) in (None, '')):
                s = styles[col]
            sattr = f' s="{s}"' if s else ''
            if val is None or val == '':
                cells[col] = f'<c r="{ref}"{sattr}/>'
            elif isinstance(val, str):
                cells[col] = (f'<c r="{ref}"{sattr} t="inlineStr"><is><t xml:space="preserve">'
                              f'{escape(val, quote=False)}</t></is></c>')
            else:
                cells[col] = f'<c r="{ref}"{sattr}><v>{_num(val)}</v></c>'
        old_row = self.rows.get(rn)
        if old_row:
            open_tag = old_row[:old_row.index('>') + 1]
            if open_tag.endswith('/>'):
                open_tag = open_tag[:-2] + '>'
            open_tag = re.sub(r'\sspans="[^"]*"', '', open_tag)
        else:
            open_tag = f'<row r="{rn}">'
        self.rows[rn] = open_tag + ''.join(cells[k] for k in sorted(cells)) + '</row>'

    def to_xml(self):
        body = ''.join(self.rows[k] for k in sorted(self.rows))
        xml = self.prefix + body + self.suffix
        max_row = max(self.rows) if self.rows else 1
        max_col = 0
        for rn in self.rows:
            if self.cells(rn):
                max_col = max(max_col, max(self.cells(rn)))
        dim = f'A1:{col_letter(max_col)}{max_row}'
        if '<dimension ' in xml:
            xml = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', xml, count=1)
        return xml


def _row_values(cols, month_kind, rec):
    out = {}
    d = date.fromisoformat(rec['fecha']) if rec.get('fecha') else None
    for i, key in enumerate(cols):
        if key == 'fecha':
            out[i] = (d - BASE_DATE).days if d else None
        elif key == 'mes':
            out[i] = (d.month if month_kind == 'number' else MONTHS[d.month - 1]) if d else None
        elif key == 'desc':
            out[i] = rec['descripcion']
        else:
            out[i] = minutes_to_excel_hours(rec.get('minutos'))
    return out


def _new_sheet_xml():
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<worksheet xmlns="{NS}" xmlns:r="{RN}"><dimension ref="A1"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            '<cols><col min="1" max="1" width="11" customWidth="1"/>'
            '<col min="2" max="2" width="12" customWidth="1"/>'
            '<col min="3" max="3" width="110" customWidth="1"/>'
            '<col min="4" max="4" width="8" customWidth="1"/></cols>'
            '<sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
            'header="0.3" footer="0.3"/></worksheet>')


def write_records(path, upserts, clears):
    """Escribe los partes en el Excel.

    upserts: lista de dicts con id, proyecto, fecha, descripcion, minutos y
             (opcional) hoja/fila si ya ocupa una fila del Excel.
    clears:  lista de (hoja, fila) cuyas casillas de datos se vacían.
    Devuelve {id: (hoja, fila)} con la posición final de cada parte.
    """
    with zipfile.ZipFile(path) as z:
        files = {i.filename: z.read(i.filename) for i in z.infolist()}
        infos = {i.filename: i for i in z.infolist()}
    wb_xml = files['xl/workbook.xml'].decode('utf-8')
    rels_xml = files['xl/_rels/workbook.xml.rels'].decode('utf-8')
    rels = {e.attrib['Id']: e.attrib['Target'] for e in ET.fromstring(rels_xml)}
    sheet_paths = {}
    for m in re.finditer(r'<sheet\b[^>]*/>', wb_xml):
        tag = m.group(0)
        name = _unescape(_attr(tag, 'name'))
        target = rels[_attr(tag, 'r:id')]
        sheet_paths[name] = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
    shared = []
    if 'xl/sharedStrings.xml' in files:
        root = ET.fromstring(files['xl/sharedStrings.xml'])
        for si in root.findall(f'{{{NS}}}si'):
            shared.append(''.join((t.text or '') for t in si.iter(f'{{{NS}}}t')))

    sheets = {}
    def sheet(name):
        if name not in sheets:
            if name not in sheet_paths:
                nonlocal wb_xml, rels_xml
                wb_xml, rels_xml = _add_sheet(files, wb_xml, rels_xml, sheet_paths, name)
            sheets[name] = _Sheet(files[sheet_paths[name]].decode('utf-8'), shared)
        return sheets[name]

    template_cache = []
    def template():
        if not template_cache:
            template_cache.append(_template_layout(dict(sheet_paths), files, shared))
        return template_cache[0]

    for hoja, fila in clears:
        if hoja in sheet_paths and fila:
            sh = sheet(hoja)
            cols, _, styles = sh.layout()
            sh.set_cells(fila, {i: None for i in range(len(cols))}, styles)

    positions = {}
    next_row = {}
    for rec in upserts:
        hoja, fila = rec.get('hoja'), rec.get('fila')
        if hoja != rec['proyecto'] or not fila or hoja not in sheet_paths:
            if hoja and fila and hoja in sheet_paths:
                sh = sheet(hoja)
                cols, _, styles = sh.layout()
                sh.set_cells(fila, {i: None for i in range(len(cols))}, styles)
            hoja, fila = rec['proyecto'], None
        sh = sheet(hoja)
        cols, month_kind, styles = sh.layout()
        if not sh.rows and template():
            cols, month_kind, styles = template()
        if fila is None:
            fila = next_row.get(hoja) or sh.last_data_row(len(cols)) + 1
            next_row[hoja] = fila + 1
        sh.set_cells(fila, _row_values(cols, month_kind, rec), styles)
        positions[rec['id']] = (hoja, fila)

    for name, sh in sheets.items():
        files[sheet_paths[name]] = sh.to_xml().encode('utf-8')
    wb_xml = _force_recalc(wb_xml)
    files['xl/workbook.xml'] = wb_xml.encode('utf-8')
    rels_xml, files = _drop_calc_chain(rels_xml, files)
    files['xl/_rels/workbook.xml.rels'] = rels_xml.encode('utf-8')

    folder = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(suffix='.xlsx', dir=folder)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as out:
            for name, data in files.items():
                info = infos.get(name)
                if info is not None:
                    zi = zipfile.ZipInfo(name, date_time=info.date_time)
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    zi.external_attr = info.external_attr
                    out.writestr(zi, data)
                else:
                    out.writestr(name, data)
        os.replace(tmp, path)
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    return positions


def _unescape(s):
    return (s.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"')
             .replace('&apos;', "'").replace('&amp;', '&'))


def _template_layout(sheet_paths, files, shared):
    """Formato de la pestaña con la fecha más reciente, para pestañas nuevas."""
    best, best_serial = None, -1
    for name, p in sheet_paths.items():
        if p not in files:
            continue
        sh = _Sheet(files[p].decode('utf-8'), shared)
        for rn in sh.rows:
            a = sh.value(sh.cells(rn).get(0))
            if isinstance(a, float) and a > best_serial:
                best, best_serial = sh, a
    return best.layout() if best else None


def _add_sheet(files, wb_xml, rels_xml, sheet_paths, name):
    n = 1
    while f'xl/worksheets/sheet{n}.xml' in files:
        n += 1
    path = f'xl/worksheets/sheet{n}.xml'
    files[path] = _new_sheet_xml().encode('utf-8')
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
    rid = f'rId{max(ids, default=0) + 1}'
    rels_xml = rels_xml.replace('</Relationships>',
        f'<Relationship Id="{rid}" Type="{SHEET_REL}" Target="worksheets/sheet{n}.xml"/></Relationships>')
    sheet_id = max((int(x) for x in re.findall(r'<sheet\b[^>]*\ssheetId="(\d+)"', wb_xml)), default=0) + 1
    wb_xml = wb_xml.replace('</sheets>',
        f'<sheet name="{escape(name[:31])}" sheetId="{sheet_id}" r:id="{rid}"/></sheets>')
    ct = files['[Content_Types].xml'].decode('utf-8')
    ct = ct.replace('</Types>', f'<Override PartName="/{path}" ContentType="application/'
                    'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    files['[Content_Types].xml'] = ct.encode('utf-8')
    sheet_paths[name] = path
    return wb_xml, rels_xml


def _force_recalc(wb_xml):
    """Pide a Excel recalcular las SUMAS al abrir el libro."""
    if '<calcPr' in wb_xml:
        m = re.search(r'<calcPr\b[^>]*/>', wb_xml)
        tag = m.group(0)
        if 'fullCalcOnLoad' not in tag:
            wb_xml = wb_xml.replace(tag, tag[:-2] + ' fullCalcOnLoad="1"/>', 1)
        return wb_xml
    return re.sub(r'(</sheets>(?:<functionGroups[^>]*/>)?(?:<externalReferences>.*?</externalReferences>)?'
                  r'(?:<definedNames>.*?</definedNames>)?)', r'\1<calcPr fullCalcOnLoad="1"/>', wb_xml,
                  count=1, flags=re.S)


def _drop_calc_chain(rels_xml, files):
    """La cadena de cálculo es una caché; Excel la regenera. Así evitamos avisos de reparación."""
    if 'xl/calcChain.xml' in files:
        del files['xl/calcChain.xml']
        rels_xml = re.sub(rf'<Relationship\b[^>]*Type="{re.escape(CALC_REL)}"[^>]*/>', '', rels_xml)
        ct = files['[Content_Types].xml'].decode('utf-8')
        ct = re.sub(r'<Override\b[^>]*PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
        files['[Content_Types].xml'] = ct.encode('utf-8')
    return rels_xml, files
