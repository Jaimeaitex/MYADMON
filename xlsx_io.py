"""Lector y exportador XLSX con la biblioteca estándar (sin dependencias externas).
El importador se orienta al libro original con una pestaña por proyecto.
"""
from __future__ import annotations
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from html import escape
import io
import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PN = 'http://schemas.openxmlformats.org/package/2006/relationships'
BASE_DATE = date(1899, 12, 30)


def serial_to_date(value):
    try:
        return BASE_DATE + timedelta(days=int(float(value)))
    except (TypeError, ValueError, OverflowError):
        return None


def date_to_serial(value):
    return (date.fromisoformat(value) - BASE_DATE).days


def normalize_hour_value(value):
    """Admite escritura histórica mixta 8.5 = 8h30 y 8.45 = 8h45.
    Devuelve (minutos, aviso). Nunca aplica 100 * hora como minutos.
    """
    if value is None or str(value).strip() == '':
        return None, 'Horas sin informar'
    try:
        d = Decimal(str(value).replace(',', '.'))
    except InvalidOperation:
        return None, f'Horas no numéricas: {value}'
    if not d.is_finite() or d < 0:
        return None, f'Horas fuera de rango: {value}'
    h = int(d)
    fraction = int(((d - Decimal(h)) * 100).quantize(Decimal('1')))
    if fraction == 50:  # .5 se ha usado en el libro como media hora
        minute = 30
    elif fraction in (0, 15, 30, 45):
        minute = fraction
    elif 0 <= fraction < 60:
        minute = fraction
        return h * 60 + minute, f'Fracción poco habitual; revisar: {value}'
    else:
        return None, f'Fracción no interpretable automáticamente: {value}'
    return h * 60 + minute, ''


def _colnum(ref):
    letters = re.match(r'[A-Z]+', ref)
    if not letters:
        return -1
    n = 0
    for c in letters.group(0):
        n = n * 26 + ord(c) - 64
    return n - 1


def read_xlsx_sheets(path):
    """Devuelve {nombre_hoja: [(número_fila_excel, {índice_col: valor}), ...]}"""
    with zipfile.ZipFile(path, 'r') as z:
        if 'xl/workbook.xml' not in z.namelist():
            raise ValueError('No parece un archivo .xlsx válido')
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rel = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relations = {e.attrib['Id']: e.attrib['Target'] for e in rel}
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall(f'{{{NS}}}si'):
                shared.append(''.join((t.text or '') for t in si.iter(f'{{{NS}}}t')))
        result = {}
        for sh in wb.iter(f'{{{NS}}}sheet'):
            name = sh.attrib['name']
            target = relations[sh.attrib[f'{{{RN}}}id']]
            target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            if target not in z.namelist():
                continue
            doc = ET.fromstring(z.read(target))
            rows = []
            for row in doc.iter(f'{{{NS}}}row'):
                rn = int(row.attrib.get('r', len(rows) + 1))
                cells = {}
                for c in row.findall(f'{{{NS}}}c'):
                    col = _colnum(c.attrib.get('r', ''))
                    if col < 0:
                        continue
                    typ = c.attrib.get('t', '')
                    v = c.find(f'{{{NS}}}v')
                    val = v.text if v is not None else None
                    if typ == 'inlineStr':
                        inline = c.find(f'{{{NS}}}is')
                        val = ''.join((t.text or '') for t in inline.iter(f'{{{NS}}}t')) if inline is not None else ''
                    elif typ == 's' and val is not None:
                        val = shared[int(val)]
                    elif typ in ('str', 'e'):
                        pass
                    elif val is not None:
                        try:
                            val = float(val)
                        except ValueError:
                            pass
                    if val is not None:
                        cells[col] = val
                if cells:
                    rows.append((rn, cells))
            result[name] = rows
        return result


def extract_original_records(path):
    """Conserva trazabilidad por pestaña/fila y no elimina actividades sin horas."""
    out = []
    warnings = []
    sheets = read_xlsx_sheets(path)
    for name, rows in sheets.items():
        if name.casefold() in ('registros', 'resumen'):
            continue
        year_match = re.search(r'20\d\d', name)
        default_year = int(year_match.group()) if year_match else 2024
        for rnum, cells in rows:
            a = cells.get(0)
            b = cells.get(1)
            c = cells.get(2)
            d = cells.get(3)
            valid_date = None
            if isinstance(a, (int, float)) and a >= 10000:
                valid_date = serial_to_date(a)
            elif isinstance(a, (int, float)) and 1 <= a <= 31 and isinstance(b, (int, float)) and 1 <= b <= 12:
                try:
                    valid_date = date(default_year, int(b), int(a))
                except ValueError:
                    warnings.append(f'{name} fila {rnum}: fecha inválida')
                    continue
            is_3_col = isinstance(b, str) and (isinstance(c, (int, float)) or (a is None and not isinstance(c, str)))
            desc = b if is_3_col else c
            hours = c if is_3_col else d
            if not isinstance(desc, str) or not desc.strip():
                continue
            if a is not None and valid_date is None:
                continue  # cabeceras o totales
            minutes, warning = normalize_hour_value(hours)
            if warning:
                warnings.append(f'{name} fila {rnum}: {warning}')
            out.append({
                'fecha': valid_date.isoformat() if valid_date else None,
                'proyecto': name, 'tecnico': 'Mi usuario',
                'actividad': 'Otras actividades', 'descripcion': desc.strip(),
                'minutos': minutes,
                'horas_original': float(hours) if isinstance(hours, (int, float)) else None,
                'origen_archivo': os.path.basename(path), 'origen_hoja': name,
                'origen_fila': rnum, 'observaciones': warning,
            })
    return out, warnings


def _column_label(n):
    x = ''
    while n:
        n, r = divmod(n - 1, 26)
        x = chr(65 + r) + x
    return x


def _cell_xml(row, col, value, style=0, raw_num=False):
    ref = f'{_column_label(col)}{row}'
    st = f' s="{style}"' if style else ''
    if value is None:
        return ''
    if raw_num and isinstance(value, (float, int)):
        return f'<c r="{ref}"{st}><v>{value}</v></c>'
    s = escape(str(value), quote=False)
    return f'<c r="{ref}" t="inlineStr"{st}><is><t xml:space="preserve">{s}</t></is></c>'


def _sheet_xml(rows, widths=None, numeric_cols=(), dates_cols=(), auto_filter=True):
    ncols = max((len(r) for r in rows), default=1)
    cols = ''
    for i, w in enumerate(widths or [], 1):
        cols += f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
    body = []
    for i, values in enumerate(rows, 1):
        cells = []
        for j, v in enumerate(values, 1):
            style = 1 if i == 1 else 2 if j in dates_cols else 3 if j in numeric_cols else 0
            cells.append(_cell_xml(i, j, v, style, j in numeric_cols or j in dates_cols))
        body.append(f'<row r="{i}">{"".join(cells)}</row>')
    end_ref = f'{_column_label(ncols)}{len(rows)}'
    autofilter = f'<autoFilter ref="A1:{end_ref}"/>' if auto_filter and len(rows) > 1 else ''
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{NS}">'
            f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="17"/><cols>{cols}</cols>'
            f'<sheetData>{"".join(body)}</sheetData>{autofilter}</worksheet>')


def export_xlsx(path, records):
    """Exporta actividades en hoja única + resumen, con horas exactas en minutos."""
    headings = ['Fecha', 'Proyecto', 'Técnico', 'Actividad', 'Descripción', 'Horas (HH:MM)',
                'Horas decimales', 'Valor original Excel', 'Observaciones', 'Hoja de origen']
    out = [headings]
    by_project = defaultdict(lambda: [0, 0, 0])
    for r in records:
        minutes = r['minutos']
        if minutes is not None:
            by_project[r['proyecto']][0] += minutes
            by_project[r['proyecto']][1] += 1
        else:
            by_project[r['proyecto']][2] += 1
        out.append([
            date_to_serial(r['fecha']) if r['fecha'] else None,
            r['proyecto'], r['tecnico'], r['actividad'], r['descripcion'],
            f'{minutes // 60:02d}:{minutes % 60:02d}' if minutes is not None else '',
            round(minutes / 60, 4) if minutes is not None else None,
            r['horas_original'], r['observaciones'], r['origen_hoja'],
        ])
    summary = [['Proyecto', 'Horas (HH:MM)', 'Horas decimales', 'Partes con horas', 'Partes pendientes']]
    for name in sorted(by_project, key=str.casefold):
        mins, count, pending = by_project[name]
        summary.append([name, f'{mins//60:02d}:{mins%60:02d}', round(mins / 60, 4), count, pending])
    total_min = sum(i[0] for i in by_project.values())
    summary.append(['TOTAL', f'{total_min//60:02d}:{total_min%60:02d}', round(total_min/60, 4),
                    sum(i[1] for i in by_project.values()), sum(i[2] for i in by_project.values())])
    xml_reg = _sheet_xml(out, [15, 24, 19, 28, 80, 18, 20, 21, 50, 24],
                         numeric_cols=(7, 8), dates_cols=(1,))
    xml_sum = _sheet_xml(summary, [27, 21, 21, 24, 23], numeric_cols=(3, 4, 5))
    styles = f'''<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="{NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF304554"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'''
    manifest = '''<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'''
    wbxml = f'''<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{NS}" xmlns:r="{RN}"><sheets><sheet name="Registros" sheetId="1" r:id="rId1"/><sheet name="Resumen" sheetId="2" r:id="rId2"/></sheets></workbook>'''
    rels = f'''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PN}"><Relationship Id="rId1" Type="{RN}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{RN}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{RN}/styles" Target="styles.xml"/></Relationships>'''
    rootrels = f'''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PN}"><Relationship Id="rId1" Type="{RN}/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for fn, contents in {'[Content_Types].xml': manifest, '_rels/.rels': rootrels,
                             'xl/workbook.xml': wbxml, 'xl/_rels/workbook.xml.rels': rels,
                             'xl/styles.xml': styles, 'xl/worksheets/sheet1.xml': xml_reg,
                             'xl/worksheets/sheet2.xml': xml_sum}.items():
            z.writestr(fn, contents.encode('utf-8'))
