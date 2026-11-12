#!/usr/bin/env python3
"""MYADMON · Partes horarios — aplicación de escritorio para Windows/Linux/macOS.
Python >= 3.10, sin paquetes de terceros. Guarda los partes en SQLite local y los
escribe en las pestañas del Excel importado.
"""
from __future__ import annotations
import calendar
import csv
import json
import os
from pathlib import Path
import shutil
import sqlite3
import subprocess
import sys
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from xlsx_io import extract_original_records, export_xlsx
from xlsx_writer import write_records, minutes_to_excel_hours
import copilot_chat
import queue
import threading

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = (Path(os.environ['LOCALAPPDATA']) if os.name == 'nt' and os.getenv('LOCALAPPDATA')
            else Path.home() / '.local' / 'share') / 'MyAdmonPartes'
DB_PATH = DATA_DIR / 'partes.sqlite3'
SEED_PATH = APP_DIR / 'data' / 'horas_iniciales.json'

BG = '#F5F7FA'
NAVY = '#304351'
NAVY_DARK = '#263848'
BLUE = '#244D64'
INK = '#243544'
MUTED = '#697C8D'
LINE = '#DFE6ED'
WHITE = '#FFFFFF'
ORANGE = '#E96655'
TEAL = '#1D8174'
WEEKEND = '#F6D6D0'
FUTURE = '#F9F8EB'
SELECTED = '#E8F2F8'
COLORS = ['#254D66', '#237C85', '#6674A9', '#A16A47', '#8664A8', '#1E7664',
          '#9B6374', '#497B4F', '#B17739', '#536B8A', '#5F7180', '#926458']
MONTHS = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
          'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
DAYS = ['LUN.', 'MAR.', 'MIÉ.', 'JUE.', 'VIE.', 'SÁB.', 'DOM.']


def hhmm(minutes):
    return 'Sin horas' if minutes is None else f'{minutes//60:02d}:{minutes%60:02d}'


def nice_date(day):
    if not day:
        return 'Sin fecha'
    return date.fromisoformat(day).strftime('%d/%m/%Y')


def parse_user_date(value):
    return datetime.strptime(value.strip(), '%d/%m/%Y').date().isoformat()


def cut(text, width):
    return text if len(text) <= width else text[:max(0, width - 1)] + '…'


class Store:
    def __init__(self, path=DB_PATH, seed=None):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute('PRAGMA foreign_keys=ON')
        self.conn.execute('PRAGMA busy_timeout=4000')
        self.conn.executescript('''
        CREATE TABLE IF NOT EXISTS ajustes(clave TEXT PRIMARY KEY, valor TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS proyectos(
          nombre TEXT PRIMARY KEY, color TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS partes(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          fecha TEXT, proyecto TEXT NOT NULL REFERENCES proyectos(nombre),
          tecnico TEXT NOT NULL DEFAULT 'Mi usuario',
          actividad TEXT NOT NULL DEFAULT 'Otras actividades',
          descripcion TEXT NOT NULL,
          minutos INTEGER CHECK(minutos IS NULL OR minutos >= 0),
          horas_original REAL, origen_archivo TEXT, origen_hoja TEXT,
          origen_fila INTEGER, observaciones TEXT NOT NULL DEFAULT '',
          creado TEXT NOT NULL DEFAULT (datetime('now')),
          modificado TEXT NOT NULL DEFAULT (datetime('now')),
          UNIQUE(origen_archivo, origen_hoja, origen_fila));
        CREATE INDEX IF NOT EXISTS idx_partes_fecha ON partes(fecha);
        CREATE INDEX IF NOT EXISTS idx_partes_proy_fecha ON partes(proyecto, fecha);
        CREATE TABLE IF NOT EXISTS excel_borrados(hoja TEXT NOT NULL, fila INTEGER NOT NULL);
        ''')
        cols = {r['name'] for r in self.conn.execute('PRAGMA table_info(partes)')}
        if 'excel_pendiente' not in cols:
            self.conn.execute('ALTER TABLE partes ADD COLUMN excel_pendiente INTEGER NOT NULL DEFAULT 0')
            self.conn.commit()
        first = self.conn.execute("SELECT valor FROM ajustes WHERE clave='inicializado'").fetchone()
        if not first:
            if seed and Path(seed).is_file():
                with open(seed, encoding='utf-8') as f:
                    values = json.load(f)
                self.import_records(values, origin_default='Horas MyAdmon.xlsx')
            self.conn.execute("INSERT INTO ajustes(clave,valor) VALUES('inicializado','1')")
            self.conn.commit()

    def setting(self, key, default=''):
        row = self.conn.execute('SELECT valor FROM ajustes WHERE clave=?', (key,)).fetchone()
        return row['valor'] if row else default

    def set_setting(self, key, value):
        self.conn.execute('INSERT INTO ajustes(clave,valor) VALUES(?,?) '
                          'ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor', (key, value))
        self.conn.commit()

    def project_list(self):
        return [dict(r) for r in self.conn.execute('SELECT * FROM proyectos ORDER BY nombre COLLATE NOCASE')]

    def project_color(self, name):
        row = self.conn.execute('SELECT color FROM proyectos WHERE nombre=?', (name,)).fetchone()
        return row['color'] if row else BLUE

    def ensure_project(self, name):
        if not name.strip():
            raise ValueError('Introduce un nombre de proyecto')
        n = self.conn.execute('SELECT COUNT(*) FROM proyectos').fetchone()[0]
        self.conn.execute('INSERT OR IGNORE INTO proyectos(nombre,color) VALUES(?,?)',
                          (name.strip(), COLORS[n % len(COLORS)]))
        self.conn.commit()

    def import_records(self, records, origin_default=None):
        added = skipped = 0
        for r in records:
            self.ensure_project(r['proyecto'])
            origin_file = r.get('origen_archivo') or origin_default
            cursor = self.conn.execute('''
              INSERT OR IGNORE INTO partes(fecha,proyecto,tecnico,actividad,descripcion,
               minutos,horas_original,origen_archivo,origen_hoja,origen_fila,observaciones)
               VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
              (r.get('fecha'), r['proyecto'], r.get('tecnico') or 'Mi usuario',
               r.get('actividad') or 'Otras actividades', r['descripcion'],
               r.get('minutos'), r.get('horas_original'), origin_file,
               r.get('origen_hoja'), r.get('origen_fila'), r.get('observaciones') or ''))
            if cursor.rowcount:
                added += 1
            else:
                skipped += 1
        self.conn.commit()
        return added, skipped

    def records(self, first=None, last=None, project=None, day=None, undated=False):
        q = 'SELECT * FROM partes WHERE 1=1'
        args = []
        if undated:
            q += ' AND fecha IS NULL'
        if day is not None:
            q += ' AND fecha=?'; args.append(day)
        if first is not None:
            q += ' AND fecha>=?'; args.append(first)
        if last is not None:
            q += ' AND fecha<=?'; args.append(last)
        if project and project != 'Todos los proyectos':
            q += ' AND proyecto=?'; args.append(project)
        q += ' ORDER BY fecha,proyecto, id'
        return [dict(r) for r in self.conn.execute(q, args)]

    def record(self, id):
        row = self.conn.execute('SELECT * FROM partes WHERE id=?', (id,)).fetchone()
        return dict(row) if row else None

    def save(self, r, id=None):
        self.ensure_project(r['proyecto'])
        data = (r.get('fecha'), r['proyecto'], r['tecnico'], r['actividad'],
                r['descripcion'], r.get('minutos'), r.get('observaciones', ''))
        if id:
            self.conn.execute('''UPDATE partes SET fecha=?,proyecto=?,tecnico=?,actividad=?,
                 descripcion=?,minutos=?,observaciones=?,modificado=datetime('now'),
                 excel_pendiente=1 WHERE id=?''', (*data, id))
        else:
            self.conn.execute('''INSERT INTO partes(fecha,proyecto,tecnico,actividad,
               descripcion,minutos,observaciones,excel_pendiente) VALUES(?,?,?,?,?,?,?,1)''', data)
        self.conn.commit()

    def delete(self, id):
        r = self.record(id)
        if r and r['origen_hoja'] and r['origen_fila'] and r['origen_archivo'] == self.excel_name():
            self.conn.execute('INSERT INTO excel_borrados(hoja,fila) VALUES(?,?)',
                              (r['origen_hoja'], r['origen_fila']))
        self.conn.execute('DELETE FROM partes WHERE id=?', (id,))
        self.conn.commit()

    def excel_path(self):
        return self.setting('excel_path', '')

    def excel_name(self):
        p = self.excel_path()
        return os.path.basename(p) if p else None

    def excel_pending_count(self):
        n = self.conn.execute('SELECT COUNT(*) FROM partes WHERE excel_pendiente=1').fetchone()[0]
        return n + self.conn.execute('SELECT COUNT(*) FROM excel_borrados').fetchone()[0]

    def sync_excel(self, backup_dir=None):
        """Escribe en el Excel vinculado los partes nuevos/editados/borrados. Devuelve nº de cambios."""
        path = self.excel_path()
        if not path or not Path(path).is_file():
            raise FileNotFoundError('No se encuentra el Excel vinculado. Vuelve a importarlo.')
        name = os.path.basename(path)
        pending = [dict(r) for r in self.conn.execute(
            'SELECT * FROM partes WHERE excel_pendiente=1 ORDER BY fecha IS NULL, fecha, id')]
        clears = [(r['hoja'], r['fila']) for r in self.conn.execute('SELECT hoja, fila FROM excel_borrados')]
        if not pending and not clears:
            return 0
        if backup_dir:
            backup_dir = Path(backup_dir)
            backup_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, backup_dir / f'{Path(path).stem}_{datetime.now():%Y%m%d_%H%M%S}.xlsx')
            for old in sorted(backup_dir.glob(f'{Path(path).stem}_*.xlsx'))[:-20]:
                old.unlink()
        upserts = [{**r, 'hoja': r['origen_hoja'] if r['origen_archivo'] == name else None,
                    'fila': r['origen_fila'] if r['origen_archivo'] == name else None} for r in pending]
        positions = write_records(path, upserts, clears)
        for r in pending:
            hoja, fila = positions[r['id']]
            self.conn.execute('''UPDATE partes SET origen_fila=NULL WHERE origen_archivo=? AND
                 origen_hoja=? AND origen_fila=? AND id<>?''', (name, hoja, fila, r['id']))
            self.conn.execute('''UPDATE partes SET origen_archivo=?, origen_hoja=?, origen_fila=?,
                 horas_original=?, excel_pendiente=0 WHERE id=?''',
                              (name, hoja, fila, minutes_to_excel_hours(r['minutos']), r['id']))
        self.conn.execute('DELETE FROM excel_borrados')
        self.conn.commit()
        return len(pending) + len(clears)

    def backup(self, dest):
        backup_conn = sqlite3.connect(dest)
        try:
            self.conn.backup(backup_conn)
        finally:
            backup_conn.close()

    def close(self):
        self.conn.close()


class PartDialog(tk.Toplevel):
    def __init__(self, app, original=None, target_date=None, template=None):
        super().__init__(app)
        self.app = app
        self.original = original
        self.title('Editar parte' if original else 'Nuevo parte horario')
        self.configure(bg=WHITE)
        self.geometry('640x645')
        self.minsize(550, 590)
        self.transient(app)
        self.grab_set()
        self.columnconfigure(0, weight=1)
        pad = tk.Frame(self, bg=WHITE, padx=24, pady=17)
        pad.pack(fill='both', expand=True)
        title = 'Editar parte' if original else 'Registrar nuevo parte'
        tk.Label(pad, text=title, bg=WHITE, fg=INK, font=('Segoe UI', 17, 'bold')).pack(anchor='w')
        tk.Label(pad, text='Los cambios se guardan en tu base de datos local y en el Excel vinculado.',
                 bg=WHITE, fg=MUTED, font=('Segoe UI', 9)).pack(anchor='w', pady=(4, 14))
        record = original or template or {}
        def field(label):
            tk.Label(pad, text=label, bg=WHITE, fg=INK, font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(8, 4))
        field('Fecha (dd/mm/aaaa)')
        init_date = record.get('fecha') or target_date or date.today().isoformat()
        self.date_v = tk.StringVar(value=nice_date(init_date))
        ttk.Entry(pad, textvariable=self.date_v, font=('Segoe UI', 11)).pack(fill='x')
        field('Proyecto')
        self.proj_v = tk.StringVar(value=record.get('proyecto') or app.current_project_or_default())
        self.proj_combo = ttk.Combobox(pad, textvariable=self.proj_v, state='readonly',
                                       values=[x['nombre'] for x in app.store.project_list()], font=('Segoe UI', 10))
        self.proj_combo.pack(fill='x')
        field('Actividad / código de tarea')
        self.act_v = tk.StringVar(value=record.get('actividad') or 'Otras actividades')
        ttk.Combobox(pad, textvariable=self.act_v, font=('Segoe UI', 10),
                     values=['Otras actividades', 'T3.2 Delivery of TBI Services',
                             'Reunión / Coordinación', 'Investigación / Estado del arte',
                             'Desarrollo / Ensayos', 'Informe / Entregable']).pack(fill='x')
        field('Descripción del trabajo realizado')
        self.description = tk.Text(pad, height=5, wrap='word', relief='solid', bd=1,
                                   fg=INK, font=('Segoe UI', 10), padx=9, pady=7)
        self.description.pack(fill='x')
        self.description.insert('1.0', record.get('descripcion', ''))
        row = tk.Frame(pad, bg=WHITE)
        row.pack(fill='x', pady=(10, 0))
        hours_block = tk.Frame(row, bg=WHITE)
        hours_block.pack(side='left')
        tk.Label(hours_block, text='Horas', bg=WHITE, fg=INK, font=('Segoe UI', 10, 'bold')).pack(anchor='w')
        dur = record.get('minutos')
        self.hours_v = tk.StringVar(value=str(dur // 60) if dur is not None else '0')
        ttk.Spinbox(hours_block, from_=0, to=999, width=8, textvariable=self.hours_v).pack(anchor='w', pady=(4, 0))
        minutes_block = tk.Frame(row, bg=WHITE)
        minutes_block.pack(side='left', padx=(15, 0))
        tk.Label(minutes_block, text='Minutos', bg=WHITE, fg=INK, font=('Segoe UI', 10, 'bold')).pack(anchor='w')
        self.mins_v = tk.StringVar(value=f'{dur%60:02d}' if dur is not None else '00')
        ttk.Combobox(minutes_block, textvariable=self.mins_v, width=7,
                     values=['00', '15', '30', '45']).pack(anchor='w', pady=(4, 0))
        self.draft_v = tk.BooleanVar(value=bool(original) and original.get('minutos') is None)
        ttk.Checkbutton(row, text='Pendiente de horas', variable=self.draft_v).pack(side='left', padx=19, pady=16)
        field('Técnico')
        self.tech_v = tk.StringVar(value=(record.get('tecnico') or app.technician.get()))
        ttk.Entry(pad, textvariable=self.tech_v).pack(fill='x')
        field('Observaciones / revisión')
        self.obs_v = tk.StringVar(value=record.get('observaciones') or '')
        ttk.Entry(pad, textvariable=self.obs_v).pack(fill='x')
        if original and original.get('horas_original') is not None:
            tk.Label(pad, text=f'Valor original de Excel (conservado para auditoría): {original["horas_original"]}',
                     bg=WHITE, fg=MUTED, font=('Segoe UI', 9)).pack(anchor='w', pady=(9, 0))
        footer = tk.Frame(self, bg=BG, padx=24, pady=12)
        footer.pack(fill='x', side='bottom')
        tk.Button(footer, text='Guardar parte', font=('Segoe UI', 10, 'bold'), bg=TEAL, fg=WHITE,
                  activebackground='#156c5f', relief='flat', padx=18, pady=8,
                  command=self.save).pack(side='right')
        tk.Button(footer, text='Cancelar', font=('Segoe UI', 10), bg=WHITE, fg=INK,
                  relief='flat', padx=16, pady=8, command=self.destroy).pack(side='right', padx=8)
        self.bind('<Escape>', lambda _: self.destroy())
        self.description.focus_set()

    def save(self):
        try:
            parsed_date = parse_user_date(self.date_v.get())
            if not self.proj_v.get().strip():
                raise ValueError('Selecciona un proyecto')
            desc = self.description.get('1.0', 'end').strip()
            if not desc:
                raise ValueError('Escribe la descripción de la actividad')
            if not self.tech_v.get().strip():
                raise ValueError('Escribe el técnico')
            h, m = int(self.hours_v.get()), int(self.mins_v.get())
            if not (0 <= h <= 999 and 0 <= m < 60):
                raise ValueError('Introduce una duración válida (0–59 minutos)')
            duration = None if self.draft_v.get() else 60*h+m
            if duration == 0:
                raise ValueError('La duración debe ser superior a cero o marcarse como pendiente')
        except (ValueError, OverflowError) as exc:
            messagebox.showerror('Revisa el parte', str(exc), parent=self)
            return
        self.app.store.save({
            'fecha': parsed_date, 'proyecto': self.proj_v.get().strip(),
            'tecnico': self.tech_v.get().strip(), 'actividad': self.act_v.get().strip() or 'Otras actividades',
            'descripcion': desc, 'minutos': duration, 'observaciones': self.obs_v.get().strip(),
        }, id=self.original['id'] if self.original else None)
        self.app.store.set_setting('tecnico', self.tech_v.get().strip())
        self.app.technician.set(self.tech_v.get().strip())
        self.app.selected = date.fromisoformat(parsed_date)
        self.app.year = self.app.selected.year
        self.app.month = self.app.selected.month
        self.destroy()
        self.app.refresh_all()
        self.app.sync_excel()


class PartesApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.store = Store()
        self.title('MyAdmon · Partes horarios')
        self.configure(bg=BG)
        self.geometry(f'{min(1900, self.winfo_screenwidth() - 40)}x900')
        self.minsize(1080, 660)
        today = date.today()
        self.year, self.month = today.year, today.month
        self.selected = today
        self.project = tk.StringVar(value='Todos los proyectos')
        self.technician = tk.StringVar(value=self.store.setting('tecnico', 'Mi usuario'))
        self.technician_filter = tk.BooleanVar(value=False)
        self.current_day_records = []
        self.month_records = []
        self._drawing = False
        self.make_style()
        self.build_ui()
        self.protocol('WM_DELETE_WINDOW', self.quit_app)
        self.refresh_all()

    def make_style(self):
        s = ttk.Style(self)
        if 'clam' in s.theme_names():
            s.theme_use('clam')
        s.configure('TCombobox', font=('Segoe UI', 10), padding=5)
        s.configure('TEntry', padding=6)
        s.configure('Treeview', font=('Segoe UI', 9), rowheight=31, background=WHITE,
                    fieldbackground=WHITE, foreground=INK, borderwidth=0)
        s.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'), background=BG,
                    foreground=INK, relief='flat')
        s.map('Treeview', background=[('selected', '#DBEEF6')], foreground=[('selected', INK)])

    def btn(self, parent, txt, cmd, bg=BLUE, fg=WHITE, **kwargs):
        return tk.Button(parent, text=txt, command=cmd, bg=bg, fg=fg,
                         activebackground=bg, activeforeground=fg,
                         relief='flat', bd=0, cursor='hand2', font=('Segoe UI', 10, 'bold'),
                         padx=11, pady=9, **kwargs)

    def build_ui(self):
        header = tk.Frame(self, bg=NAVY, height=48)
        header.pack(side='top', fill='x')
        header.pack_propagate(False)
        tk.Label(header, text='◉', bg=ORANGE, fg=WHITE,
                 font=('Segoe UI', 18, 'bold'), width=3).pack(side='left', fill='y')
        tk.Label(header, text='MYADMON  /  GESTIÓN DE PROYECTOS  /  CALENDARIO',
                 bg=NAVY, fg=WHITE, font=('Segoe UI', 10, 'bold'), padx=19).pack(side='left')
        tk.Label(header, text='PARTES HORARIOS · LOCAL', bg=NAVY, fg='#C9DCE6',
                 font=('Segoe UI', 9), padx=20).pack(side='right')
        tk.Button(header, text='💬  Asistente Copilot', command=self.toggle_chat, bg=TEAL, fg=WHITE,
                  activebackground=TEAL, activeforeground=WHITE, relief='flat', bd=0, cursor='hand2',
                  font=('Segoe UI', 9, 'bold'), padx=12).pack(side='right', fill='y', pady=9)
        body = tk.Frame(self, bg=BG)
        body.pack(fill='both', expand=True)
        self.side = tk.Frame(body, bg='#F2F5F7', width=265, padx=17, pady=18,
                             highlightbackground=LINE, highlightthickness=1)
        self.side.pack(side='left', fill='y')
        self.side.pack_propagate(False)
        tk.Label(self.side, text='CALENDARIO', fg=INK, bg='#F2F5F7',
                 font=('Segoe UI', 17)).pack(anchor='w', pady=(4, 28))
        self.side_title('Proyecto')
        self.proj_combo = ttk.Combobox(self.side, state='readonly', textvariable=self.project)
        self.proj_combo.pack(fill='x', pady=(4, 11))
        self.proj_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_all())
        self.btn(self.side, '+  Añadir proyecto', self.new_project, bg='#E5ECF0', fg=INK).pack(fill='x', pady=(0, 15))
        self.side_title('Técnico')
        ttk.Entry(self.side, textvariable=self.technician).pack(fill='x', pady=(4, 7))
        ttk.Checkbutton(self.side, text='Filtrar por este técnico', variable=self.technician_filter,
                        command=self.refresh_all).pack(anchor='w', pady=(0, 20))
        self.btn(self.side, '+  Nuevo parte', self.new_part, bg=TEAL).pack(fill='x', pady=(0, 9))
        self.btn(self.side, '☷  Pendientes / sin fecha', self.pending, bg='#E2EBF1', fg=INK).pack(fill='x', pady=(0, 25))
        tk.Frame(self.side, bg=LINE, height=1).pack(fill='x', pady=(0, 15))
        self.side_title('RESUMEN DEL MES')
        self.month_stats = tk.Label(self.side, text='', bg='#F2F5F7', fg=BLUE,
                                    font=('Segoe UI', 21, 'bold'), anchor='w')
        self.month_stats.pack(fill='x', pady=(7, 0))
        self.month_substats = tk.Label(self.side, text='', bg='#F2F5F7', fg=MUTED,
                                       justify='left', font=('Segoe UI', 10))
        self.month_substats.pack(anchor='w', pady=(0, 24))
        self.side_title('DATOS Y COPIAS')
        self.btn(self.side, 'Exportar a Excel (.xlsx)', self.export_excel, bg='#E2EBF1', fg=INK).pack(fill='x', pady=(7, 7))
        self.btn(self.side, 'Importar Excel original', self.import_excel, bg='#E2EBF1', fg=INK).pack(fill='x', pady=(0, 7))
        self.btn(self.side, 'Escribir en Excel ahora', lambda: self.sync_excel(manual=True),
                 bg='#E2EBF1', fg=INK).pack(fill='x', pady=(0, 7))
        self.btn(self.side, 'Copia de seguridad (.db)', self.backup_db, bg='#E2EBF1', fg=INK).pack(fill='x', pady=(0, 7))
        self.btn(self.side, 'Abrir carpeta de datos', self.open_data_dir, bg='#E2EBF1', fg=INK).pack(fill='x', pady=(0, 14))
        self.excel_status = tk.Label(self.side, text='', bg='#F2F5F7', fg=MUTED, justify='left',
                                     wraplength=230, font=('Segoe UI', 9))
        self.excel_status.pack(anchor='w')
        self.update_excel_status()
        self.side.update_idletasks()

        self.detail = tk.Frame(body, bg=WHITE, width=345, padx=14, pady=14,
                               highlightbackground=LINE, highlightthickness=1)
        self.detail.pack(side='right', fill='y')
        self.detail.pack_propagate(False)
        self.right_title = tk.Label(self.detail, text='', bg=WHITE, fg=INK,
                                    font=('Segoe UI', 14, 'bold'), anchor='w')
        self.right_title.pack(fill='x', pady=(7, 5))
        self.right_subtitle = tk.Label(self.detail, text='', bg=WHITE, fg=MUTED,
                                       font=('Segoe UI', 10), anchor='w')
        self.right_subtitle.pack(fill='x', pady=(0, 11))
        self.btn(self.detail, '+ Añadir parte a este día', self.new_part, bg=TEAL).pack(fill='x', pady=(0, 17))
        self.day_tree = ttk.Treeview(self.detail, columns=('project', 'time'), show='headings', selectmode='browse')
        self.day_tree.heading('project', text='Proyecto / Parte')
        self.day_tree.heading('time', text='Horas')
        self.day_tree.column('project', width=220, stretch=True)
        self.day_tree.column('time', width=70, anchor='center', stretch=False)
        self.day_tree.pack(fill='both', expand=True)
        self.day_tree.bind('<<TreeviewSelect>>', lambda e: self.update_detail())
        self.day_tree.bind('<Double-1>', lambda e: self.edit_selected())
        self.part_desc = tk.Label(self.detail, text='Selecciona un parte para ver su descripción.',
                                  bg=WHITE, fg=INK, anchor='nw', justify='left',
                                  font=('Segoe UI', 9), wraplength=310, height=6)
        self.part_desc.pack(fill='x', pady=(15, 8))
        btnrow = tk.Frame(self.detail, bg=WHITE)
        btnrow.pack(fill='x')
        self.btn(btnrow, 'Editar', self.edit_selected, bg=BLUE).pack(side='left', expand=True, fill='x', padx=(0, 4))
        self.btn(btnrow, 'Duplicar', self.duplicate_selected, bg='#E5ECF0', fg=INK).pack(side='left', expand=True, fill='x', padx=4)
        self.btn(btnrow, 'Eliminar', self.delete_selected, bg='#FDEAE7', fg='#963C2E').pack(side='left', expand=True, fill='x', padx=(4, 0))
        self.build_chat(body)

        center = tk.Frame(body, bg=WHITE)
        center.pack(side='left', fill='both', expand=True)
        nav = tk.Frame(center, bg=WHITE, height=92, padx=15, pady=12)
        nav.pack(fill='x')
        nav.columnconfigure(3, weight=1)
        self.btn(nav, '◀', lambda: self.change_month(-1), bg='#E8EEF2', fg=INK).grid(row=0, column=0, padx=3)
        self.btn(nav, '▶', lambda: self.change_month(1), bg='#E8EEF2', fg=INK).grid(row=0, column=1, padx=3)
        self.btn(nav, 'Hoy', self.to_today, bg='#E8EEF2', fg=INK).grid(row=0, column=2, padx=8)
        self.month_title = tk.Label(nav, text='', bg=WHITE, fg=BLUE, font=('Segoe UI', 13, 'bold'))
        self.month_title.grid(row=0, column=3, sticky='nsew')
        self.month_choice = tk.StringVar()
        self.month_box = ttk.Combobox(nav, width=12, state='readonly', textvariable=self.month_choice, values=MONTHS)
        self.month_box.grid(row=0, column=4, padx=3)
        self.month_box.bind('<<ComboboxSelected>>', lambda e: self.jump_month())
        self.year_choice = tk.StringVar()
        ttk.Entry(nav, textvariable=self.year_choice, width=7).grid(row=0, column=5, padx=4)
        self.btn(nav, 'Ir', self.jump_month, bg='#E8EEF2', fg=INK).grid(row=0, column=6, padx=3)
        self.legend = tk.Label(center, text='Haz clic en un día para ver sus partes · Doble clic en un parte para editarlo',
                               bg=WHITE, fg=MUTED, font=('Segoe UI', 9), anchor='w', padx=18)
        self.legend.pack(fill='x', pady=(0, 8))
        board = tk.Frame(center, bg=WHITE)
        board.pack(fill='both', expand=True)
        self.canvas = tk.Canvas(board, bg=WHITE, highlightthickness=0, bd=0)
        self.canvas.pack(side='left', fill='both', expand=True)
        self.canvas.bind('<Configure>', lambda e: self.draw_calendar())
        self.refresh_projects()

    def side_title(self, title):
        tk.Label(self.side, text=title, bg='#F2F5F7', fg=INK,
                 font=('Segoe UI', 10, 'bold')).pack(anchor='w')

    # --- Asistente Copilot -------------------------------------------------
    def build_chat(self, body):
        self.chat_history = []
        self.chat_queue = queue.Queue()
        self.chat_proc = None
        self.chat_busy = False
        self.chat_last = None
        self.chat = tk.Frame(body, bg=WHITE, width=330, padx=12, pady=12,
                             highlightbackground=LINE, highlightthickness=1)
        self.chat.pack(side='right', fill='y', before=self.detail)
        self.chat.pack_propagate(False)
        top = tk.Frame(self.chat, bg=WHITE)
        top.pack(fill='x')
        tk.Label(top, text='💬 Asistente Copilot', bg=WHITE, fg=INK,
                 font=('Segoe UI', 13, 'bold')).pack(side='left')
        tk.Button(top, text='Nueva', command=self.chat_reset, bg='#E5ECF0', fg=INK, relief='flat',
                  bd=0, cursor='hand2', font=('Segoe UI', 8), padx=8, pady=3).pack(side='right')
        tk.Label(self.chat, text='Proyecto para las ideas', bg=WHITE, fg=MUTED,
                 font=('Segoe UI', 9)).pack(anchor='w', pady=(8, 2))
        self.chat_project = tk.StringVar()
        self.chat_proj_combo = ttk.Combobox(self.chat, state='readonly', textvariable=self.chat_project)
        self.chat_proj_combo.pack(fill='x')
        log_frame = tk.Frame(self.chat, bg=WHITE)
        log_frame.pack(fill='both', expand=True, pady=(10, 8))
        scroll = ttk.Scrollbar(log_frame, orient='vertical')
        self.chat_log = tk.Text(log_frame, wrap='word', relief='flat', bg='#F7F9FB', fg=INK,
                                font=('Segoe UI', 9), padx=8, pady=8, state='disabled',
                                yscrollcommand=scroll.set, cursor='arrow')
        scroll.configure(command=self.chat_log.yview)
        scroll.pack(side='right', fill='y')
        self.chat_log.pack(side='left', fill='both', expand=True)
        self.chat_log.tag_configure('user', foreground=BLUE, font=('Segoe UI', 9, 'bold'),
                                    spacing1=8, lmargin1=40, lmargin2=40)
        self.chat_log.tag_configure('bot', foreground=INK, spacing1=8)
        self.chat_log.tag_configure('info', foreground=MUTED, font=('Segoe UI', 8, 'italic'), spacing1=6)
        self.chat_log.tag_configure('error', foreground='#963C2E', spacing1=6)
        self.btn(self.chat, '💡  Dame una idea para el parte de este día', lambda: self.chat_send(
            'Dame una idea de qué poner en el parte de este día para el proyecto seleccionado.'),
            bg=TEAL).pack(fill='x', pady=(0, 6))
        self.chat_use_btn = self.btn(self.chat, '✚  Crear parte con la última sugerencia',
                                     self.chat_use_suggestion, bg='#E2EBF1', fg=INK)
        self.chat_use_btn.pack(fill='x', pady=(0, 8))
        entry_row = tk.Frame(self.chat, bg=WHITE)
        entry_row.pack(fill='x')
        self.chat_entry = tk.Text(entry_row, height=3, wrap='word', relief='solid', bd=1,
                                  font=('Segoe UI', 9), padx=6, pady=4)
        self.chat_entry.pack(side='left', fill='x', expand=True)
        self.chat_entry.bind('<Return>', self._chat_enter)
        self.chat_send_btn = self.btn(entry_row, 'Enviar', self.chat_send, bg=BLUE)
        self.chat_send_btn.pack(side='right', fill='y', padx=(6, 0))
        tk.Label(self.chat, text='Intro envía · Mayús+Intro salto de línea.\nSe envían a GitHub Copilot '
                 'los partes recientes del proyecto elegido.', bg=WHITE, fg=MUTED, justify='left',
                 font=('Segoe UI', 8)).pack(anchor='w', pady=(6, 0))
        self.chat_write('info', 'Pídeme ideas para el parte: tengo en cuenta tus últimos partes del '
                        'proyecto seleccionado y lo que ya has apuntado en el día.')

    def toggle_chat(self):
        if self.chat.winfo_ismapped():
            self.chat.pack_forget()
        else:
            self.chat.pack(side='right', fill='y', before=self.detail)

    def refresh_chat_projects(self):
        names = [r['nombre'] for r in self.store.project_list()]
        self.chat_proj_combo.configure(values=names)
        if self.current_filter():
            self.chat_project.set(self.current_filter())
        elif self.chat_project.get() not in names:
            row = self.store.conn.execute('SELECT proyecto FROM partes WHERE fecha IS NOT NULL '
                                          'ORDER BY fecha DESC, id DESC LIMIT 1').fetchone()
            self.chat_project.set(row['proyecto'] if row else (names[0] if names else ''))

    def chat_context(self):
        project = self.chat_project.get()
        day = self.selected
        lines = [f'Hoy es {nice_date(date.today().isoformat())}. '
                 f'Día del parte: {DAYS[day.weekday()].rstrip(".").title()} {nice_date(day.isoformat())}.',
                 f'Proyecto seleccionado: {project or "(ninguno)"}']
        if project:
            recent = self.store.conn.execute(
                'SELECT fecha, minutos, descripcion FROM partes WHERE proyecto=? AND fecha IS NOT NULL '
                'ORDER BY fecha DESC, id DESC LIMIT 30', (project,)).fetchall()
            lines.append(f'\nÚltimos partes de «{project}» (del más reciente al más antiguo):')
            lines += [f'- {nice_date(r["fecha"])} · {hhmm(r["minutos"])} · {r["descripcion"]}' for r in recent] \
                or ['- (sin partes previos)']
        today = self.store.records(day=day.isoformat())
        lines.append(f'\nPartes ya registrados el {nice_date(day.isoformat())}:')
        lines += [f'- {r["proyecto"]} · {hhmm(r["minutos"])} · {r["descripcion"]}' for r in today] or ['- (ninguno)']
        if today:
            lines.append(f'Total ya imputado ese día: {hhmm(sum(r["minutos"] or 0 for r in today))}')
        return '\n'.join(lines)

    def chat_write(self, tag, text):
        self.chat_log.configure(state='normal')
        prefix = {'user': 'Tú: ', 'bot': 'Copilot: '}.get(tag, '')
        self.chat_log.insert('end', prefix + text.strip() + '\n', tag)
        self.chat_log.configure(state='disabled')
        self.chat_log.see('end')

    def _chat_enter(self, event):
        if event.state & 0x0001:  # Mayús
            return None
        self.chat_send()
        return 'break'

    def chat_send(self, question=None):
        if self.chat_busy:
            return
        if question is None:
            question = self.chat_entry.get('1.0', 'end').strip()
            if not question:
                return
            self.chat_entry.delete('1.0', 'end')
        prompt = copilot_chat.build_prompt(self.chat_context(), self.chat_history, question)
        self.chat_history.append(('user', question))
        self.chat_write('user', question)
        self.chat_write('info', 'Copilot está pensando…')
        self.chat_busy = True
        self.chat_send_btn.configure(state='disabled')
        def work():
            try:
                self.chat_queue.put(('ok', copilot_chat.ask(prompt, DATA_DIR / 'copilot',
                                                            on_start=lambda p: setattr(self, 'chat_proc', p))))
            except Exception as e:
                self.chat_queue.put(('error', str(e)))
        threading.Thread(target=work, daemon=True).start()
        self.after(200, self._chat_poll)

    def _chat_poll(self):
        try:
            kind, text = self.chat_queue.get_nowait()
        except queue.Empty:
            self.after(200, self._chat_poll)
            return
        self.chat_busy = False
        self.chat_proc = None
        self.chat_send_btn.configure(state='normal')
        self.chat_log.configure(state='normal')
        last = self.chat_log.search('Copilot está pensando…', 'end', backwards=True)
        if last:
            self.chat_log.delete(last, f'{last} lineend +1c')
        self.chat_log.configure(state='disabled')
        if kind == 'ok':
            self.chat_history.append(('assistant', text))
            self.chat_write('bot', text)
            desc, minutes = copilot_chat.parse_suggestion(text)
            if desc:
                self.chat_last = (desc, minutes, self.chat_project.get())
        else:
            self.chat_history.pop()
            self.chat_write('error', text)

    def chat_use_suggestion(self):
        if not self.chat_last:
            messagebox.showinfo('Asistente', 'Pide primero una idea al asistente.', parent=self)
            return
        desc, minutes, project = self.chat_last
        PartDialog(self, target_date=self.selected.isoformat(),
                   template={'descripcion': desc, 'minutos': minutes, 'proyecto': project})

    def chat_reset(self):
        self.chat_history.clear()
        self.chat_last = None
        self.chat_log.configure(state='normal')
        self.chat_log.delete('1.0', 'end')
        self.chat_log.configure(state='disabled')
        self.chat_write('info', 'Nueva conversación.')

    def current_project_or_default(self):
        proj = self.project.get()
        ps = self.store.project_list()
        return proj if proj != 'Todos los proyectos' else (ps[0]['nombre'] if ps else '')

    def current_filter(self):
        return None if self.project.get() == 'Todos los proyectos' else self.project.get()

    def technician_filtered(self, records):
        if not self.technician_filter.get():
            return records
        t = self.technician.get().strip().casefold()
        return [r for r in records if r['tecnico'].casefold() == t]

    def refresh_projects(self):
        names = ['Todos los proyectos'] + [r['nombre'] for r in self.store.project_list()]
        current = self.project.get()
        self.proj_combo.configure(values=names)
        if current not in names:
            self.project.set('Todos los proyectos')

    def refresh_all(self):
        self.refresh_projects()
        self.refresh_chat_projects()
        first = date(self.year, self.month, 1)
        last = date(self.year, self.month, calendar.monthrange(self.year, self.month)[1])
        self.month_records = self.technician_filtered(self.store.records(first.isoformat(), last.isoformat(),
                                                                            self.current_filter()))
        self.month_title.configure(text=f'{MONTHS[self.month - 1].upper()} {self.year}')
        self.month_choice.set(MONTHS[self.month-1])
        self.year_choice.set(str(self.year))
        minutes = sum(r['minutos'] or 0 for r in self.month_records)
        drafts = sum(r['minutos'] is None for r in self.month_records)
        self.month_stats.configure(text=hhmm(minutes) + ' h')
        self.month_substats.configure(text=f'{len(self.month_records)} actividades · {drafts} sin horas')
        self.draw_calendar()
        self.refresh_day()

    def draw_calendar(self):
        if self._drawing:
            return
        self._drawing = True
        try:
            can = self.canvas
            can.delete('all')
            w, h = max(550, can.winfo_width()), max(400, can.winfo_height())
            cell_w = w / 7.0
            weeks = calendar.Calendar(firstweekday=0).monthdatescalendar(self.year, self.month)
            header_h = 33
            cell_h = max(88, (h-header_h)/len(weeks))
            can.configure(scrollregion=(0, 0, w, header_h + cell_h * len(weeks)))
            for i, day_label in enumerate(DAYS):
                x = i * cell_w
                can.create_rectangle(x, 0, x+cell_w, header_h, fill='#F0F4F7', outline=LINE)
                can.create_text(x+cell_w/2, header_h/2, text=day_label, fill=MUTED,
                                font=('Segoe UI', 9, 'bold'))
            by_day = {}
            for r in self.month_records:
                by_day.setdefault(r['fecha'], []).append(r)
            for wi, week in enumerate(weeks):
                for di, day in enumerate(week):
                    x0, y0 = di*cell_w, header_h+wi*cell_h
                    x1, y1 = x0+cell_w, y0+cell_h
                    in_month = day.month == self.month
                    bg = WEEKEND if day.weekday() >= 5 and in_month else WHITE
                    if in_month and day > date.today() and day.weekday() < 5:
                        bg = FUTURE
                    if not in_month: bg = '#F9FAFB'
                    if day == self.selected: bg = SELECTED
                    iso = day.isoformat()
                    tag = f'dia_{iso}'
                    can.create_rectangle(x0, y0, x1, y1, fill=bg, outline=LINE,
                                         width=1, tags=(tag,))
                    can.create_text(x0+9, y0+14, text=str(day.day), anchor='w',
                                    font=('Segoe UI', 10, 'bold' if day == date.today() else 'normal'),
                                    fill=ORANGE if day == date.today() else (MUTED if not in_month else INK),
                                    tags=(tag,))
                    if day == date.today():
                        can.create_oval(x1-16, y0+8, x1-8, y0+16, fill=TEAL, outline='', tags=(tag,))
                    if in_month:
                        entries = by_day.get(iso, [])
                        max_events = max(1, min(3, int((cell_h-46)/24)))
                        for ix, r in enumerate(entries[:max_events]):
                            yy = y0 + 29 + ix * 23
                            c = self.store.project_color(r['proyecto'])
                            eventtag = f'part_{r["id"]}'
                            can.create_rectangle(x0+5, yy, x1-5, yy+20, fill=c,
                                                 outline='', tags=(eventtag,))
                            txt = f'{r["proyecto"]}: {hhmm(r["minutos"])}'
                            can.create_text(x0+9, yy+10, text=cut(txt, max(10,int(cell_w/7.4))),
                                            anchor='w', fill=WHITE, font=('Segoe UI', 8), tags=(eventtag,))
                            can.tag_bind(eventtag, '<Button-1>', lambda ev, id=r['id']: self.focus_part(id))
                        if len(entries) > max_events:
                            can.create_text(x0+9, y0+32+max_events*23,
                                            text=f'+ {len(entries)-max_events} más', anchor='w',
                                            fill=BLUE, font=('Segoe UI', 8, 'bold'), tags=(tag,))
                        if entries:
                            min_sum = sum(r['minutos'] or 0 for r in entries)
                            can.create_text(x1-9, y1-10, text=hhmm(min_sum) + ' h', anchor='e',
                                            fill=BLUE, font=('Segoe UI', 8, 'bold'), tags=(tag,))
                    can.tag_bind(tag, '<Button-1>', lambda ev, value=day: self.pick_day(value))
        finally:
            self._drawing = False

    def pick_day(self, d):
        self.selected = d
        if self.month != d.month or self.year != d.year:
            self.year, self.month = d.year, d.month
            self.refresh_all()
        else:
            self.draw_calendar()
            self.refresh_day()

    def change_month(self, diff):
        pos = self.year*12 + self.month-1 + diff
        self.year = pos // 12
        self.month = pos % 12 + 1
        self.selected = date(self.year, self.month, 1)
        self.refresh_all()

    def to_today(self):
        self.selected = date.today()
        self.year, self.month = self.selected.year, self.selected.month
        self.refresh_all()

    def jump_month(self):
        try:
            y = int(self.year_choice.get())
            m = MONTHS.index(self.month_choice.get())+1
            if not 1900 <= y <= 2100: raise ValueError
        except (ValueError, IndexError):
            messagebox.showwarning('Fecha', 'Selecciona un mes y un año entre 1900 y 2100', parent=self)
            return
        self.year, self.month = y, m
        self.selected = date(y, m, 1)
        self.refresh_all()

    def refresh_day(self):
        self.right_title.configure(text=self.selected.strftime('%d/%m/%Y'))
        records = self.technician_filtered(self.store.records(day=self.selected.isoformat(),
                                                              project=self.current_filter()))
        self.current_day_records = records
        total = sum(r['minutos'] or 0 for r in records)
        self.right_subtitle.configure(text=f'{len(records)} actividades  ·  {hhmm(total)} h')
        for id in self.day_tree.get_children():
            self.day_tree.delete(id)
        for r in records:
            descr = cut(r['descripcion'].replace('\n',' '), 36)
            name = cut(r['proyecto'], 23)
            self.day_tree.insert('', 'end', iid=str(r['id']), values=(name + ' · ' + descr, hhmm(r['minutos'])))
        self.part_desc.configure(text='Selecciona un parte para ver su descripción.')

    def update_detail(self):
        sel = self.day_tree.selection()
        if not sel: return
        r = self.store.record(int(sel[0]))
        if r:
            orig = f'\nOriginal Excel: {r["horas_original"]}' if r['horas_original'] is not None else ''
            self.part_desc.configure(text=f'{r["descripcion"]}\n\nActividad: {r["actividad"]}{orig}')

    def selected_id(self):
        sel = self.day_tree.selection()
        return int(sel[0]) if sel else None

    def focus_part(self, id):
        rec = self.store.record(id)
        if rec and rec['fecha']:
            self.pick_day(date.fromisoformat(rec['fecha']))
            if self.day_tree.exists(str(id)):
                self.day_tree.selection_set(str(id))
                self.day_tree.see(str(id))
                self.update_detail()

    def new_part(self):
        PartDialog(self, target_date=self.selected.isoformat())

    def edit_selected(self):
        id = self.selected_id()
        if not id:
            messagebox.showinfo('Editar', 'Selecciona primero una actividad del listado de la derecha.', parent=self)
            return
        PartDialog(self, original=self.store.record(id))

    def duplicate_selected(self):
        id = self.selected_id()
        if not id:
            messagebox.showinfo('Duplicar', 'Selecciona un parte para duplicarlo.', parent=self)
            return
        r = self.store.record(id)
        r['fecha'] = self.selected.isoformat()
        PartDialog(self, original=None, target_date=r['fecha'], template=r)

    def delete_selected(self):
        id = self.selected_id()
        if id is None: return
        if messagebox.askyesno('Eliminar parte', '¿Eliminar definitivamente este parte?\n'
                               'Puedes hacer una copia de seguridad antes.', parent=self):
            self.store.delete(id)
            self.refresh_all()
            self.sync_excel()

    def new_project(self):
        dialog = tk.Toplevel(self)
        dialog.title('Nuevo proyecto')
        dialog.geometry('440x178')
        dialog.resizable(False, False)
        dialog.configure(bg=WHITE)
        dialog.transient(self)
        dialog.grab_set()
        tk.Label(dialog, text='Nombre del nuevo proyecto', bg=WHITE, fg=INK,
                 font=('Segoe UI', 13, 'bold')).pack(anchor='w', padx=20, pady=(20, 8))
        name = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=name)
        entry.pack(fill='x', padx=20)
        entry.focus()
        def save():
            try: self.store.ensure_project(name.get())
            except ValueError as e:
                messagebox.showwarning('Proyecto', str(e), parent=dialog)
                return
            dialog.destroy()
            self.project.set(name.get().strip())
            self.refresh_all()
        self.btn(dialog, 'Añadir proyecto', save, bg=TEAL).pack(anchor='e', padx=20, pady=14)
        dialog.bind('<Return>', lambda e: save())

    def pending(self):
        records = self.technician_filtered(self.store.records(project=self.current_filter()))
        records = [r for r in records if r['minutos'] is None or r['fecha'] is None]
        top = tk.Toplevel(self)
        top.title('Actividades pendientes de revisar')
        top.geometry('1000x520')
        top.configure(bg=WHITE)
        tk.Label(top, text=f'{len(records)} actividades pendientes', bg=WHITE, fg=INK,
                 font=('Segoe UI', 16, 'bold')).pack(anchor='w', padx=18, pady=15)
        columns = ('fecha', 'proyecto', 'horas', 'descripcion')
        tree = ttk.Treeview(top, columns=columns, show='headings')
        for k, t, w in [('fecha', 'Fecha', 130), ('proyecto', 'Proyecto', 180),
                         ('horas', 'Horas', 90), ('descripcion', 'Descripción', 540)]:
            tree.heading(k, text=t)
            tree.column(k, width=w, stretch=(k == 'descripcion'))
        tree.pack(fill='both', expand=True, padx=18, pady=5)
        for r in records:
            tree.insert('', 'end', iid=str(r['id']), values=(nice_date(r['fecha']), r['proyecto'],
                                                               hhmm(r['minutos']), r['descripcion']))
        def edit():
            sel = tree.selection()
            if sel:
                r = self.store.record(int(sel[0]))
                top.destroy()
                PartDialog(self, original=r)
        self.btn(top, 'Editar actividad seleccionada', edit, bg=TEAL).pack(anchor='e', padx=18, pady=15)
        tree.bind('<Double-1>', lambda ev: edit())

    def export_excel(self):
        fn = filedialog.asksaveasfilename(title='Exportar todos los partes a Excel',
               initialfile='Mis_partes_horarios.xlsx', defaultextension='.xlsx',
               filetypes=[('Libro Excel', '*.xlsx')], parent=self)
        if not fn: return
        try:
            export_xlsx(fn, self.store.records())
            messagebox.showinfo('Exportación finalizada',
                                'Excel generado con las hojas Registros y Resumen.\n'
                                'Los valores originales y las horas normalizadas se conservan en columnas distintas.', parent=self)
        except Exception as e:
            messagebox.showerror('No se ha podido exportar', str(e), parent=self)

    def import_excel(self):
        fn = filedialog.askopenfilename(title='Selecciona el Excel original de partes',
               filetypes=[('Libro Excel', '*.xlsx')], parent=self)
        if not fn: return
        try:
            data, warnings = extract_original_records(fn)
            if not data:
                raise ValueError('No se han reconocido registros en formato de una pestaña por proyecto.')
            preview = f'Se han detectado {len(data)} registros y {len(warnings)} avisos de revisión.\n\n'
            preview += ('Los registros ya importados desde las mismas pestañas/filas se omitirán; '
                        'no se sobrescribirán cambios hechos en la aplicación.\n\n¿Continuar?')
            if not messagebox.askyesno('Importar Excel', preview, parent=self): return
            added, skipped = self.store.import_records(data)
            self.store.set_setting('excel_path', str(Path(fn).resolve()))
            self.refresh_all()
            self.update_excel_status()
            messagebox.showinfo('Importación', f'Añadidos: {added}\nExistentes (omitidos): {skipped}\n'
                                f'Avisos sobre horas/fechas: {len(warnings)}\n\n'
                                'A partir de ahora los partes que crees, edites o borres se escribirán '
                                'en este Excel.', parent=self)
            self.sync_excel()
        except Exception as e:
            messagebox.showerror('Error al importar', str(e), parent=self)

    def update_excel_status(self):
        path = self.store.excel_path()
        if not path:
            text = 'Sin Excel vinculado: usa «Importar Excel original».'
        else:
            pend = self.store.excel_pending_count()
            text = f'Excel vinculado: {os.path.basename(path)}'
            text += f'\n⚠ {pend} cambio(s) pendientes de escribir.' if pend else '\n✓ Al día.'
        self.excel_status.configure(text=text, fg=ORANGE if path and self.store.excel_pending_count() else MUTED)

    def sync_excel(self, manual=False):
        if not self.store.excel_path():
            if manual:
                messagebox.showinfo('Excel', 'Primero importa el Excel con «Importar Excel original».', parent=self)
            return
        try:
            n = self.store.sync_excel(DATA_DIR / 'copias_excel')
            if manual:
                messagebox.showinfo('Excel actualizado', f'Cambios escritos en el Excel: {n}', parent=self)
        except PermissionError:
            messagebox.showwarning('Excel abierto',
                'No se puede escribir en el Excel porque está abierto o bloqueado.\n'
                'Ciérralo y pulsa «Escribir en Excel ahora». Los cambios quedan guardados en la app.',
                parent=self)
        except Exception as e:
            messagebox.showerror('Error al escribir en Excel', str(e), parent=self)
        self.update_excel_status()

    def backup_db(self):
        name = f'partes_backup_{datetime.now():%Y%m%d_%H%M}.sqlite3'
        fn = filedialog.asksaveasfilename(title='Guardar copia de seguridad',
               initialfile=name, defaultextension='.sqlite3',
               filetypes=[('Base de datos SQLite', '*.sqlite3'), ('Todos los archivos', '*.*')], parent=self)
        if not fn: return
        try:
            self.store.backup(fn)
            messagebox.showinfo('Copia guardada', 'Copia íntegra de tus datos guardada correctamente.', parent=self)
        except Exception as e:
            messagebox.showerror('Error de copia', str(e), parent=self)

    def open_data_dir(self):
        try:
            if os.name == 'nt':
                os.startfile(DATA_DIR)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(DATA_DIR)])
            else:
                subprocess.Popen(['xdg-open', str(DATA_DIR)])
        except Exception as e:
            messagebox.showinfo('Carpeta de datos', str(DATA_DIR), parent=self)

    def quit_app(self):
        if self.chat_proc and self.chat_proc.poll() is None:
            self.chat_proc.kill()
        self.store.close()
        self.destroy()


if __name__ == '__main__':
    PartesApp().mainloop()
