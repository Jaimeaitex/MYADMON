"""Autocomprobación de datos, SQLite y XLSX; no modifica los datos reales."""
from pathlib import Path
from tempfile import TemporaryDirectory
import json
import sqlite3
import zipfile
import xml.etree.ElementTree as ET
from partes_aitex import Store, SEED_PATH
from xlsx_io import export_xlsx, extract_original_records, read_xlsx_sheets, normalize_hour_value


def check(seed_path=SEED_PATH):
    with open(seed_path, encoding='utf8') as f:
        seed = json.load(f)
    assert len(seed) == 780
    assert len({r['proyecto'] for r in seed}) == 18
    assert sum(r['fecha'] is not None for r in seed) == 778
    assert sum(r['minutos'] is not None for r in seed) == 775
    assert normalize_hour_value(8.5)[0] == 510
    assert normalize_hour_value(8.15)[0] == 495
    assert normalize_hour_value(9.27)[0] == 567

    with TemporaryDirectory() as td:
        db = Path(td) / 'partes.db'
        backup = Path(td) / 'seguridad.db'
        exp = Path(td) / 'exportacion.xlsx'
        store = Store(db, seed_path)
        all_records = store.records()
        assert len(all_records) == 780
        assert len(store.project_list()) == 18
        nov_2026 = store.records('2026-11-01', '2026-11-30')
        assert nov_2026
        pk = all_records[0]['id']
        store.save({'fecha':'2026-11-12','proyecto':seed[0]['proyecto'],'tecnico':'Mi usuario',
                    'actividad':'Otras actividades','descripcion':'Prueba del formulario',
                    'minutos':135, 'observaciones':''})
        assert len(store.records()) == 781
        added_id = max(r['id'] for r in store.records())
        store.save({'fecha':'2026-11-12','proyecto':seed[0]['proyecto'],'tecnico':'Mi usuario',
                    'actividad':'Otras actividades','descripcion':'Prueba editada',
                    'minutos':145, 'observaciones':''}, id=added_id)
        assert store.record(added_id)['minutos'] == 145
        store.delete(added_id)
        assert len(store.records()) == 780
        store.backup(backup)
        assert sqlite3.connect(backup).execute('select count(*) from partes').fetchone()[0] == 780
        export_xlsx(exp, store.records())
        with zipfile.ZipFile(exp) as z:
            assert z.testzip() is None
            for f in ('xl/workbook.xml','xl/styles.xml','xl/worksheets/sheet1.xml',
                      'xl/worksheets/sheet2.xml'):
                ET.fromstring(z.read(f))
        exported = read_xlsx_sheets(exp)
        assert len(exported['Registros']) == 781  # header + records
        assert len(exported['Resumen']) == 20  # header + 18 projects + total
        store.close()
        reopen = Store(db, seed_path)
        assert len(reopen.records()) == 780
        reopen.close()
        check_excel_write(Path(td))
    from copilot_chat import parse_suggestion
    assert parse_suggestion('Idea:\nPARTE: Pruebas de visión\nHORAS: 7:30') == ('Pruebas de visión', 450)
    assert parse_suggestion('**PARTE:** «Reunión»\n**HORAS:** 8,5') == ('Reunión', 510)
    assert parse_suggestion('Sin formato') == (None, None)
    print('OK: 18 proyectos, 780 registros, fechas, horas, altas, ediciones, borrados, copia SQLite y Excel XLSX.')


def check_excel_write(td):
    book = td / 'libro.xlsx'
    export_xlsx(book, [])
    store = Store(td / 'escritura.db')
    store.set_setting('excel_path', str(book))
    base = {'tecnico': 'yo', 'actividad': 'Otras actividades', 'observaciones': ''}
    store.save({**base, 'fecha': '2026-11-12', 'proyecto': 'Proyecto X', 'descripcion': 'Uno', 'minutos': 510})
    store.save({**base, 'fecha': '2026-11-13', 'proyecto': 'Proyecto X', 'descripcion': 'Dos', 'minutos': 135})
    assert store.sync_excel(td / 'copias') == 2 and store.excel_pending_count() == 0
    rows = read_xlsx_sheets(book)['Proyecto X']
    assert [r[0] for r in rows] == [1, 2]
    assert rows[0][1][2] == 'Uno' and rows[0][1][3] == 8.5 and rows[1][1][3] == 2.15
    data, _ = extract_original_records(book)
    assert store.import_records(data) == (0, 2)
    first, second = store.records(project='Proyecto X')
    store.save({**base, 'fecha': first['fecha'], 'proyecto': 'Proyecto X', 'descripcion': 'Uno bis', 'minutos': 60},
               id=first['id'])
    store.delete(second['id'])
    store.sync_excel()
    rows = dict(read_xlsx_sheets(book)['Proyecto X'])
    assert rows[1][2] == 'Uno bis' and rows[1][3] == 1.0 and 2 not in rows
    with zipfile.ZipFile(book) as z:
        assert z.testzip() is None
    store.close()


if __name__ == '__main__':
    if SEED_PATH.is_file():
        check()
    else:
        # El repositorio público no contiene partes personales.
        with TemporaryDirectory() as td:
            seed_path = Path(td) / 'datos_prueba.json'
            records = [
                {'fecha': f'2026-11-{i % 28 + 1:02d}' if i < 778 else None,
                 'proyecto': f'Proyecto {i % 18 + 1:02d}',
                 'descripcion': f'Actividad de prueba {i + 1}',
                 'minutos': 60 if i < 775 else None,
                 'origen_hoja': f'Proyecto {i % 18 + 1:02d}',
                 'origen_fila': i + 1}
                for i in range(780)
            ]
            seed_path.write_text(json.dumps(records), encoding='utf-8')
            check(seed_path)
