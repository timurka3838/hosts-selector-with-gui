#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
table_tool.py  —  Обработка таблиц (единое приложение)
=======================================================
Этап 1 — «Обработка файлов»
  • Выберите папку с .csv / .xls / .xlsx файлами
  • Удаляются столбцы n3, n4, n7, n8 (файлы перезаписываются)
  • Собираются уникальные хосты (включая пустые)
  • hosts_result.xlsx сохраняется в ту же папку

Этап 2 — «Управление хостами»
  • Фильтр по файлу (выпадающий список) + текстовый поиск
  • Сценарий А   → заменить все хосты на метку файла
  • Сценарий Б   → добавить колонку "dang" у выбранных
  • Удалить      → удалить строки с выбранными хостами

Зависимости:  pip install pandas openpyxl xlrd
"""

import os
import re
import logging
import threading
from pathlib import Path
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, List, Optional, Set, Tuple
from tkinter import filedialog, scrolledtext

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════════════════
#  Константы
# ══════════════════════════════════════════════════════════════════════════
CHUNK_SIZE       = 100_000
MAX_WORKERS      = os.cpu_count() or 4
OUTPUT_FILENAME  = 'hosts_result.xlsx'
COLUMNS_TO_DROP  = {'n3', 'n4', 'n7', 'n8'}
EMPTY_DISPLAY    = '(пусто)'

# ══════════════════════════════════════════════════════════════════════════
#  Логирование
# ══════════════════════════════════════════════════════════════════════════
log = logging.getLogger('table_tool')
log.setLevel(logging.DEBUG)

_fh = logging.FileHandler('table_tool.log', encoding='utf-8')
_fh.setFormatter(logging.Formatter('%(asctime)s  %(levelname)-8s  %(message)s'))
log.addHandler(_fh)

_sh = logging.StreamHandler()
_sh.setFormatter(logging.Formatter('%(asctime)s  %(levelname)-8s  %(message)s'))
log.addHandler(_sh)


# ══════════════════════════════════════════════════════════════════════════
#  Утилиты
# ══════════════════════════════════════════════════════════════════════════

def extract_label(filename: str) -> str:
    """'n7nsdg (word).xlsx'  →  'word'"""
    m = re.search(r'\(([^)]+)\)', Path(filename).stem)
    return m.group(1).strip() if m else Path(filename).stem


def detect_csv_encoding(path: Path) -> str:
    for enc in ('utf-8-sig', 'utf-8', 'cp1251', 'latin-1'):
        try:
            pd.read_csv(path, nrows=2, encoding=enc)
            return enc
        except Exception:
            continue
    return 'utf-8'


def _remove(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def get_source_files(folder_path: str) -> List[Path]:
    """Все .csv/.xlsx/.xlsm, кроме выходного файла hosts_result.xlsx."""
    folder = Path(folder_path)
    return sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in {'.csv', '.xlsx', '.xlsm'}
        and p.name != OUTPUT_FILENAME
    )


# ══════════════════════════════════════════════════════════════════════════
#  Этап 1: обработка исходных файлов
# ══════════════════════════════════════════════════════════════════════════

def process_csv(filepath: str, label: str) -> List[Tuple[str, str]]:
    path = Path(filepath)
    tmp  = path.with_suffix('.tmp.csv')
    hosts: set = set()
    enc   = detect_csv_encoding(path)
    try:
        first = True
        for chunk in pd.read_csv(path, chunksize=CHUNK_SIZE,
                                  encoding=enc, low_memory=False):
            if 'host' in chunk.columns:
                hosts.update(chunk['host'].fillna('').astype(str).unique())
            drop = [c for c in COLUMNS_TO_DROP if c in chunk.columns]
            if drop:
                chunk.drop(columns=drop, inplace=True)
            chunk.to_csv(tmp, index=False,
                         mode='w' if first else 'a',
                         header=first, encoding='utf-8-sig')
            first = False
        os.replace(tmp, path)
        log.info(f"[CSV]   готов: {path.name}  ({len(hosts)} хостов)")
    except Exception:
        _remove(tmp)
        raise
    return [(label, h) for h in hosts]


def process_xlsx(filepath: str, label: str) -> List[Tuple[str, str]]:
    path = Path(filepath)
    tmp  = path.with_suffix('.tmp.xlsx')
    hosts: set = set()
    try:
        wb_r = load_workbook(path, read_only=True, data_only=True)
        ws_r = wb_r.active
        it   = ws_r.iter_rows(values_only=True)
        hdr  = list(next(it))
        drop_idx = {i for i, h in enumerate(hdr)
                    if h is not None and str(h).strip() in COLUMNS_TO_DROP}
        host_idx = next((i for i, h in enumerate(hdr)
                         if h is not None and str(h).strip() == 'host'), None)
        new_hdr  = [h for i, h in enumerate(hdr) if i not in drop_idx]

        wb_w = openpyxl.Workbook(write_only=True)
        ws_w = wb_w.create_sheet()
        ws_w.append(new_hdr)
        buf  = []
        for row in it:
            if host_idx is not None:
                hosts.add('' if row[host_idx] is None else str(row[host_idx]))
            buf.append([v for i, v in enumerate(row) if i not in drop_idx])
            if len(buf) >= CHUNK_SIZE:
                for r in buf: ws_w.append(r)
                buf.clear()
        for r in buf: ws_w.append(r)
        wb_r.close()
        wb_w.save(tmp)
        os.replace(tmp, path)
        log.info(f"[XLSX]  готов: {path.name}  ({len(hosts)} хостов)")
    except Exception:
        _remove(tmp)
        raise
    return [(label, h) for h in hosts]


def process_xls(filepath: str, label: str) -> List[Tuple[str, str]]:
    path  = Path(filepath)
    hosts: set = set()
    try:
        df = pd.read_excel(path, engine='xlrd')
        if 'host' in df.columns:
            hosts.update(df['host'].fillna('').astype(str).unique())
        drop = [c for c in COLUMNS_TO_DROP if c in df.columns]
        if drop:
            df.drop(columns=drop, inplace=True)
        new_path = path.with_suffix('.xlsx')
        df.to_excel(new_path, index=False)
        path.unlink()
        log.info(f"[XLS]   конвертирован → {new_path.name}  ({len(hosts)} хостов)")
    except Exception:
        raise
    return [(label, h) for h in hosts]


def process_file(filepath: str) -> List[Tuple[str, str]]:
    p   = Path(filepath)
    lbl = extract_label(p.name)
    ext = p.suffix.lower()
    log.info(f"Начало: {p.name}  (метка «{lbl}»)")
    if ext == '.csv':               return process_csv(filepath, lbl)
    elif ext in ('.xlsx', '.xlsm'): return process_xlsx(filepath, lbl)
    elif ext == '.xls':             return process_xls(filepath, lbl)
    else:
        log.warning(f"Пропущен (неизвестный формат): {p.name}")
        return []


def run_phase1(folder_path: str,
               progress_cb: Optional[Callable] = None) -> List[Tuple[str, str]]:
    folder = Path(folder_path)
    files  = sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in {'.csv', '.xls', '.xlsx', '.xlsm'}
        and p.name != OUTPUT_FILENAME
    )
    if not files:
        raise FileNotFoundError(f'Файлы не найдены в: {folder_path}')

    all_rows: List[Tuple[str, str]] = []
    done = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futs = {pool.submit(process_file, str(f)): f for f in files}
        for fut in as_completed(futs):
            fname = futs[fut].name
            done += 1
            try:
                rows = fut.result()
                all_rows.extend(rows)
                log.info(f"[{done}/{len(files)}] готов: {fname}  (+{len(rows)})")
            except Exception as exc:
                log.error(f"[{done}/{len(files)}] ошибка: {fname}: {exc}",
                           exc_info=True)
            if progress_cb:
                progress_cb(done, len(files))
    return all_rows


def save_hosts_xlsx(all_rows: List[Tuple[str, str]], output_path: Path) -> None:
    df = (
        pd.DataFrame(all_rows, columns=['Название файла', 'Имя хоста'])
        .drop_duplicates()
        .sort_values('Имя хоста',
                     key=lambda s: s.str.lower(),
                     ignore_index=True)
    )
    df.insert(0, 'Запись',
              df['Название файла'] + ' - ' +
              df['Имя хоста'].apply(lambda h: EMPTY_DISPLAY if h == '' else h))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Хосты'
    hfill = PatternFill('solid', start_color='1F4E79')
    hfont = Font(bold=True, color='FFFFFF', name='Arial', size=11)
    cfont = Font(name='Arial', size=10)
    left  = Alignment(horizontal='left',   vertical='center')
    ctr   = Alignment(horizontal='center', vertical='center')

    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(1, ci, col)
        c.fill = hfill; c.font = hfont; c.alignment = ctr

    for ri, row in enumerate(df.itertuples(index=False), 2):
        for ci, val in enumerate(row, 1):
            c = ws.cell(ri, ci, val)
            c.font = cfont; c.alignment = left

    for ci, col in enumerate(df.columns, 1):
        w = max(len(str(col)), df[col].astype(str).str.len().max())
        ws.column_dimensions[get_column_letter(ci)].width = min(w + 4, 80)

    ws.freeze_panes = 'A2'
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    log.info(f"Сохранено: {output_path}  ({len(df)} строк)")


# ══════════════════════════════════════════════════════════════════════════
#  Сценарий А: заменить все хосты на метку файла
# ══════════════════════════════════════════════════════════════════════════

def _replace_hosts_csv(fp: Path, val: str) -> None:
    enc = detect_csv_encoding(fp)
    tmp = fp.with_suffix('.tmp.csv')
    try:
        first = True
        for chunk in pd.read_csv(fp, chunksize=CHUNK_SIZE,
                                  encoding=enc, low_memory=False):
            if 'host' in chunk.columns:
                chunk['host'] = val
            chunk.to_csv(tmp, index=False,
                         mode='w' if first else 'a',
                         header=first, encoding='utf-8-sig')
            first = False
        os.replace(tmp, fp)
    except Exception:
        _remove(tmp); raise


def _replace_hosts_xlsx(fp: Path, val: str) -> None:
    tmp = fp.with_suffix('.tmp.xlsx')
    try:
        wb_r = load_workbook(fp, read_only=True, data_only=True)
        ws_r = wb_r.active
        it   = ws_r.iter_rows(values_only=True)
        hdr  = list(next(it))
        hidx = next((i for i, h in enumerate(hdr)
                     if h is not None and str(h).strip() == 'host'), None)
        wb_w = openpyxl.Workbook(write_only=True)
        ws_w = wb_w.create_sheet()
        ws_w.append(hdr)
        buf  = []
        for row in it:
            r = list(row)
            if hidx is not None: r[hidx] = val
            buf.append(r)
            if len(buf) >= CHUNK_SIZE:
                for rr in buf: ws_w.append(rr)
                buf.clear()
        for rr in buf: ws_w.append(rr)
        wb_r.close(); wb_w.save(tmp); os.replace(tmp, fp)
    except Exception:
        _remove(tmp); raise


def scenario_a_apply(folder_path: str,
                     progress_cb: Optional[Callable] = None) -> int:
    files = get_source_files(folder_path)
    done  = 0
    def _do(fp: Path):
        lbl = extract_label(fp.name)
        (_replace_hosts_csv if fp.suffix.lower() == '.csv'
         else _replace_hosts_xlsx)(fp, lbl)
        log.info(f"[А] хосты → «{lbl}»: {fp.name}")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futs = {pool.submit(_do, f): f for f in files}
        for fut in as_completed(futs):
            fut.result(); done += 1
            if progress_cb: progress_cb(done, len(files))
    return done


# ══════════════════════════════════════════════════════════════════════════
#  Сценарий Б: добавить колонку 'dang'
# ══════════════════════════════════════════════════════════════════════════

def _mark_dang_csv(fp: Path, hosts: Set[str]) -> None:
    enc = detect_csv_encoding(fp)
    tmp = fp.with_suffix('.tmp.csv')
    emp = '' in hosts
    try:
        first = True
        for chunk in pd.read_csv(fp, chunksize=CHUNK_SIZE,
                                  encoding=enc, low_memory=False):
            if 'host' in chunk.columns:
                chunk['dang'] = chunk['host'].apply(
                    lambda h: 'dang' if (
                        (pd.isna(h) and emp) or
                        (not pd.isna(h) and str(h) in hosts)
                    ) else '')
            chunk.to_csv(tmp, index=False,
                         mode='w' if first else 'a',
                         header=first, encoding='utf-8-sig')
            first = False
        os.replace(tmp, fp)
    except Exception:
        _remove(tmp); raise


def _mark_dang_xlsx(fp: Path, hosts: Set[str]) -> None:
    tmp = fp.with_suffix('.tmp.xlsx')
    emp = '' in hosts
    try:
        wb_r = load_workbook(fp, read_only=True, data_only=True)
        ws_r = wb_r.active
        it   = ws_r.iter_rows(values_only=True)
        hdr  = list(next(it))
        hidx = next((i for i, h in enumerate(hdr)
                     if h is not None and str(h).strip() == 'host'), None)
        sh   = [str(h).strip() if h else '' for h in hdr]
        if 'dang' in sh:
            didx = sh.index('dang'); new_hdr = list(hdr)
        else:
            didx = len(hdr); new_hdr = list(hdr) + ['dang']
        wb_w = openpyxl.Workbook(write_only=True)
        ws_w = wb_w.create_sheet()
        ws_w.append(new_hdr)
        buf  = []
        for row in it:
            r  = list(row)
            hv = '' if (hidx is None or r[hidx] is None) else str(r[hidx])
            mk = 'dang' if ((not hv and emp) or hv in hosts) else ''
            if didx < len(r): r[didx] = mk
            else: r.append(mk)
            buf.append(r)
            if len(buf) >= CHUNK_SIZE:
                for rr in buf: ws_w.append(rr)
                buf.clear()
        for rr in buf: ws_w.append(rr)
        wb_r.close(); wb_w.save(tmp); os.replace(tmp, fp)
    except Exception:
        _remove(tmp); raise


def scenario_b_apply(folder_path: str,
                     checked: Set[Tuple[str, str]],
                     progress_cb: Optional[Callable] = None) -> int:
    lth: dict = defaultdict(set)
    for lbl, h in checked: lth[lbl].add(h)
    files = get_source_files(folder_path)
    done  = 0
    def _do(fp: Path):
        lbl = extract_label(fp.name)
        if lbl not in lth: return
        (_mark_dang_csv if fp.suffix.lower() == '.csv'
         else _mark_dang_xlsx)(fp, lth[lbl])
        log.info(f"[Б] dang → {len(lth[lbl])} хост(ов): {fp.name}")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futs = {pool.submit(_do, f): f for f in files}
        for fut in as_completed(futs):
            fut.result(); done += 1
            if progress_cb: progress_cb(done, len(files))
    return done


# ══════════════════════════════════════════════════════════════════════════
#  Удаление строк с выбранными хостами
# ══════════════════════════════════════════════════════════════════════════

def _delete_hosts_csv(fp: Path, hosts: Set[str]) -> None:
    enc = detect_csv_encoding(fp)
    tmp = fp.with_suffix('.tmp.csv')
    emp = '' in hosts
    try:
        first = True
        for chunk in pd.read_csv(fp, chunksize=CHUNK_SIZE,
                                  encoding=enc, low_memory=False):
            if 'host' in chunk.columns:
                mask = ~chunk['host'].apply(
                    lambda h: (pd.isna(h) and emp) or
                              (not pd.isna(h) and str(h) in hosts))
                chunk = chunk[mask]
            if not chunk.empty:
                chunk.to_csv(tmp, index=False,
                             mode='w' if first else 'a',
                             header=first, encoding='utf-8-sig')
                first = False
        if not first:
            os.replace(tmp, fp)
        else:
            # все строки удалены — оставляем файл с шапкой
            cols = pd.read_csv(fp, nrows=0, encoding=enc).columns
            pd.DataFrame(columns=cols).to_csv(fp, index=False,
                                               encoding='utf-8-sig')
            _remove(tmp)
    except Exception:
        _remove(tmp); raise


def _delete_hosts_xlsx(fp: Path, hosts: Set[str]) -> None:
    tmp = fp.with_suffix('.tmp.xlsx')
    emp = '' in hosts
    try:
        wb_r = load_workbook(fp, read_only=True, data_only=True)
        ws_r = wb_r.active
        it   = ws_r.iter_rows(values_only=True)
        hdr  = list(next(it))
        hidx = next((i for i, h in enumerate(hdr)
                     if h is not None and str(h).strip() == 'host'), None)
        wb_w = openpyxl.Workbook(write_only=True)
        ws_w = wb_w.create_sheet()
        ws_w.append(hdr)
        buf  = []
        for row in it:
            hv = '' if (hidx is None or row[hidx] is None) else str(row[hidx])
            if (not hv and emp) or hv in hosts:
                continue
            buf.append(list(row))
            if len(buf) >= CHUNK_SIZE:
                for rr in buf: ws_w.append(rr)
                buf.clear()
        for rr in buf: ws_w.append(rr)
        wb_r.close(); wb_w.save(tmp); os.replace(tmp, fp)
    except Exception:
        _remove(tmp); raise


def delete_hosts_apply(folder_path: str,
                       checked: Set[Tuple[str, str]],
                       progress_cb: Optional[Callable] = None) -> int:
    lth: dict = defaultdict(set)
    for lbl, h in checked: lth[lbl].add(h)
    files = get_source_files(folder_path)
    done  = 0
    def _do(fp: Path):
        lbl = extract_label(fp.name)
        if lbl not in lth: return
        (_delete_hosts_csv if fp.suffix.lower() == '.csv'
         else _delete_hosts_xlsx)(fp, lth[lbl])
        log.info(f"[Del] удалены строки с {len(lth[lbl])} хост(ами): {fp.name}")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futs = {pool.submit(_do, f): f for f in files}
        for fut in as_completed(futs):
            fut.result(); done += 1
            if progress_cb: progress_cb(done, len(files))
    return done


# ══════════════════════════════════════════════════════════════════════════
#  GUI: обработчик лога
# ══════════════════════════════════════════════════════════════════════════

class _GUILogHandler(logging.Handler):
    _COLORS = {
        logging.DEBUG:    '#888888',
        logging.INFO:     '#111111',
        logging.WARNING:  '#b07000',
        logging.ERROR:    '#cc0000',
        logging.CRITICAL: '#990000',
    }

    def __init__(self, widget):
        super().__init__()
        self.widget = widget
        self.setFormatter(logging.Formatter(
            '%(asctime)s  %(levelname)-8s  %(message)s', datefmt='%H:%M:%S'))

    def emit(self, record: logging.LogRecord) -> None:
        msg   = self.format(record) + '\n'
        color = self._COLORS.get(record.levelno, '#111111')
        lvl   = record.levelname
        def _write():
            self.widget.configure(state='normal')
            self.widget.insert('end', msg, lvl)
            self.widget.tag_config(lvl, foreground=color)
            self.widget.see('end')
            self.widget.configure(state='disabled')
        self.widget.after(0, _write)


# ══════════════════════════════════════════════════════════════════════════
#  GUI: главное приложение
# ══════════════════════════════════════════════════════════════════════════

class App:
    _CHK_ON  = '☑'
    _CHK_OFF = '☐'

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title('Обработка таблиц')
        self.root.geometry('1100x740')
        self.root.minsize(820, 560)

        self._folder: str = ''
        self._all: List[Tuple[str, str]] = []       # (label, host_internal)
        self._checked: Set[Tuple[str, str]] = set()
        self._busy    = False
        self._sort_asc = {'label': True, 'host': True}
        self._filter_job: Optional[str] = None

        self._build_ui()

    # ── Построение интерфейса ─────────────────────────────────────────────

    def _build_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        style.configure('Action.TButton', font=('Arial', 10, 'bold'), padding=7)
        style.configure('Del.TButton',
                        font=('Arial', 10, 'bold'), padding=7,
                        foreground='#cc0000')

        self._nb = ttk.Notebook(self.root)
        self._nb.pack(fill='both', expand=True, padx=6, pady=6)

        self._tab1 = ttk.Frame(self._nb)
        self._tab2 = ttk.Frame(self._nb)
        self._nb.add(self._tab1, text='  📂  Обработка файлов  ')
        self._nb.add(self._tab2, text='  🔍  Управление хостами  ')
        self._nb.tab(1, state='disabled')

        self._build_tab1()
        self._build_tab2()

    # ── Вкладка 1 ─────────────────────────────────────────────────────────

    def _build_tab1(self) -> None:
        t = self._tab1

        # Папка
        frm = ttk.LabelFrame(t, text=' Папка с исходными файлами ', padding=10)
        frm.pack(fill='x', padx=12, pady=(12, 6))
        self._folder_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self._folder_var,
                  font=('Arial', 10)).pack(side='left', fill='x',
                                            expand=True, padx=(0, 6))
        ttk.Button(frm, text='📁  Выбрать папку…',
                   command=self._on_folder_browse).pack(side='left')

        # Кнопки
        bfr = ttk.Frame(t, padding=(12, 0))
        bfr.pack(fill='x')
        self._btn_run = ttk.Button(
            bfr, text='▶  Запустить обработку файлов',
            style='Action.TButton', command=self._run_phase1)
        self._btn_run.pack(side='left', padx=(0, 8))
        self._btn_load = ttk.Button(
            bfr, text='📋  Загрузить существующий hosts_result.xlsx',
            command=self._load_existing)
        self._btn_load.pack(side='left')

        # Прогресс
        pfr = ttk.Frame(t, padding=(12, 6, 12, 0))
        pfr.pack(fill='x')
        self._p1_var = tk.DoubleVar(value=0)
        ttk.Progressbar(pfr, variable=self._p1_var,
                        mode='determinate', maximum=100).pack(fill='x')
        self._p1_lbl = tk.StringVar(value='')
        ttk.Label(pfr, textvariable=self._p1_lbl,
                  font=('Arial', 9), anchor='w').pack(fill='x', pady=(2, 0))

        # Лог
        lfr = ttk.LabelFrame(t, text=' Лог обработки ', padding=6)
        lfr.pack(fill='both', expand=True, padx=12, pady=(8, 12))
        self._log_txt = scrolledtext.ScrolledText(
            lfr, height=14, state='disabled',
            font=('Consolas', 9), wrap='word', bg='#fafafa')
        self._log_txt.pack(fill='both', expand=True)

        log.addHandler(_GUILogHandler(self._log_txt))

    # ── Вкладка 2 ─────────────────────────────────────────────────────────

    def _build_tab2(self) -> None:
        t = self._tab2

        # Фильтры
        ffr = ttk.LabelFrame(t, text=' Фильтры ', padding=(10, 8))
        ffr.pack(fill='x', padx=12, pady=(10, 4))

        # Строка 1 — фильтр по файлу
        r1 = ttk.Frame(ffr)
        r1.pack(fill='x', pady=(0, 5))
        ttk.Label(r1, text='Файл:', width=8,
                  font=('Arial', 10)).pack(side='left')
        self._file_var = tk.StringVar(value='Все файлы')
        self._file_cb  = ttk.Combobox(
            r1, textvariable=self._file_var,
            state='readonly', width=38, font=('Arial', 10))
        self._file_cb['values'] = ['Все файлы']
        self._file_cb.pack(side='left', padx=(4, 0))
        self._file_cb.bind('<<ComboboxSelected>>',
                           lambda _: self._schedule_filter())

        # Строка 2 — текстовый поиск + кнопки
        r2 = ttk.Frame(ffr)
        r2.pack(fill='x')
        ttk.Label(r2, text='Поиск:', width=8,
                  font=('Arial', 10)).pack(side='left')
        self._qry_var = tk.StringVar()
        self._qry_var.trace_add('write', lambda *_: self._schedule_filter())
        ttk.Entry(r2, textvariable=self._qry_var,
                  width=38, font=('Arial', 10)).pack(side='left', padx=(4, 2))
        ttk.Button(r2, text='✕', width=3,
                   command=lambda: self._qry_var.set('')).pack(side='left',
                                                                padx=(0, 14))
        ttk.Separator(r2, orient='vertical').pack(side='left', fill='y', padx=4)
        ttk.Button(r2, text='☑  Выбрать все',
                   command=self._check_all).pack(side='left', padx=4)
        ttk.Button(r2, text='☐  Снять всё',
                   command=self._uncheck_all).pack(side='left', padx=4)

        # Treeview
        tfr = ttk.Frame(t, padding=(12, 0, 12, 0))
        tfr.pack(fill='both', expand=True)

        cols = ('chk', 'label', 'host')
        self._tree = ttk.Treeview(tfr, columns=cols, show='headings',
                                   selectmode='browse')
        self._tree.heading('chk',   text='')
        self._tree.heading('label', text='Название файла',
                           command=lambda: self._sort_by('label'))
        self._tree.heading('host',  text='Имя хоста',
                           command=lambda: self._sort_by('host'))
        self._tree.column('chk',   width=36,  minwidth=36,  stretch=False, anchor='center')
        self._tree.column('label', width=220, minwidth=120, stretch=False, anchor='w')
        self._tree.column('host',  width=700, minwidth=200, stretch=True,  anchor='w')

        vsb = ttk.Scrollbar(tfr, orient='vertical',   command=self._tree.yview)
        hsb = ttk.Scrollbar(tfr, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right',  fill='y')
        hsb.pack(side='bottom', fill='x')
        self._tree.pack(fill='both', expand=True)

        self._tree.bind('<Button-1>', self._on_click)
        self._tree.bind('<space>',    self._on_space)
        self._tree.tag_configure('checked',      background='#c8e6c9', foreground='#1b5e20')
        self._tree.tag_configure('unchecked',    background='',         foreground='')
        self._tree.tag_configure('empty_chk',   background='#c8e6c9', foreground='#5d7a5d')
        self._tree.tag_configure('empty_unchk', background='',         foreground='#999999')

        # Строка состояния
        ttk.Separator(t, orient='horizontal').pack(fill='x', padx=12, pady=(4, 0))
        self._stat_var = tk.StringVar(value='—')
        self._stat_lbl = ttk.Label(t, textvariable=self._stat_var,
                                    font=('Arial', 9), anchor='w',
                                    padding=(12, 3))
        self._stat_lbl.pack(fill='x')

        # Прогрессбар (скрыт до начала операции)
        self._p2_var = tk.DoubleVar(value=0)
        self._p2     = ttk.Progressbar(t, variable=self._p2_var,
                                        mode='determinate', maximum=100)

        # Кнопки действий
        ttk.Separator(t, orient='horizontal').pack(fill='x', padx=12)
        bot = ttk.Frame(t, padding=(12, 8, 12, 10))
        bot.pack(fill='x')

        self._btn_a = ttk.Button(
            bot,
            text='Нет интересных хостов\n→ Заменить на имя файла  (Сцен. А)',
            style='Action.TButton', command=self._run_scenario_a, width=37)
        self._btn_a.pack(side='left', padx=(0, 6))

        self._btn_b = ttk.Button(
            bot,
            text='✔  Отметить выбранные как "dang"\n    (Сценарий Б)',
            style='Action.TButton', command=self._run_scenario_b, width=30)
        self._btn_b.pack(side='left', padx=(0, 6))

        self._btn_del = ttk.Button(
            bot,
            text='🗑  Удалить строки\n    с выбранными хостами',
            style='Del.TButton', command=self._run_delete, width=22)
        self._btn_del.pack(side='right')

    # ── Логика вкладки 1 ──────────────────────────────────────────────────

    def _on_folder_browse(self) -> None:
        path = filedialog.askdirectory(title='Выберите папку с файлами')
        if not path:
            return
        self._folder_var.set(path)
        self._folder = path
        out = Path(path) / OUTPUT_FILENAME
        if out.exists():
            log.info(f'Найден существующий файл: {out.name}')
            self._load_hosts_from(out)

    def _run_phase1(self) -> None:
        folder = self._folder_var.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showerror('Ошибка', 'Укажите корректный путь к папке.')
            return
        if self._busy:
            return
        self._folder = folder
        self._set_busy1(True)
        self._p1_var.set(0)

        def run():
            try:
                all_rows = run_phase1(folder, self._p1_progress)
                out      = Path(folder) / OUTPUT_FILENAME
                save_hosts_xlsx(all_rows, out)
                self.root.after(0, lambda: self._on_phase1_done(all_rows, out))
            except Exception as exc:
                log.error(f'Ошибка этапа 1: {exc}', exc_info=True)
                msg = str(exc)
                self.root.after(0, lambda: [
                    self._set_busy1(False),
                    messagebox.showerror('Ошибка', msg)])

        threading.Thread(target=run, daemon=True).start()

    def _p1_progress(self, done: int, total: int) -> None:
        pct = done / total * 100 if total else 100
        d, t = done, total
        def _u():
            self._p1_var.set(pct)
            self._p1_lbl.set(f'Обработано файлов: {d} / {t}')
        self.root.after(0, _u)

    def _on_phase1_done(self, all_rows: list, out: Path) -> None:
        self._set_busy1(False)
        self._p1_var.set(100)
        log.info(f'Этап 1 завершён. Всего хостов: {len(all_rows)}')
        self._load_hosts_from(out)
        self._nb.tab(1, state='normal')
        self._nb.select(1)

    def _load_existing(self) -> None:
        folder = self._folder_var.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showerror('Ошибка', 'Укажите корректный путь к папке.')
            return
        self._folder = folder
        out = Path(folder) / OUTPUT_FILENAME
        if not out.exists():
            messagebox.showerror(
                'Файл не найден',
                f'Файл {OUTPUT_FILENAME} не найден в:\n{folder}')
            return
        self._load_hosts_from(out)
        self._nb.tab(1, state='normal')
        self._nb.select(1)

    def _set_busy1(self, busy: bool) -> None:
        self._busy = busy
        st = ('disabled',) if busy else ('!disabled',)
        self._btn_run.state(st)
        self._btn_load.state(st)

    # ── Загрузка хостов ───────────────────────────────────────────────────

    def _load_hosts_from(self, path: Path) -> None:
        try:
            df = pd.read_excel(path)
            if 'Название файла' in df.columns and 'Имя хоста' in df.columns:
                lc, hc = 'Название файла', 'Имя хоста'
            else:
                cols = list(df.columns)
                lc, hc = cols[-2], cols[-1]
                log.warning(f'Авто-определение колонок: «{lc}», «{hc}»')
            self._all = list(zip(
                df[lc].fillna('').astype(str),
                df[hc].fillna('').astype(str)))
            self._checked.clear()
            self._update_file_combo(preserve=False)
            self._apply_filter()
            log.info(f'Загружено хостов: {len(self._all)}')
        except Exception as exc:
            log.error(f'Ошибка загрузки: {exc}', exc_info=True)
            messagebox.showerror('Ошибка загрузки', str(exc))

    def _update_file_combo(self, preserve: bool = True) -> None:
        current = self._file_var.get()
        labels  = sorted(set(lbl for lbl, _ in self._all))
        values  = ['Все файлы'] + labels
        self._file_cb['values'] = values
        if preserve and current in values:
            self._file_var.set(current)
        else:
            self._file_var.set('Все файлы')

    # ── Фильтрация ────────────────────────────────────────────────────────

    def _schedule_filter(self) -> None:
        if self._filter_job:
            self.root.after_cancel(self._filter_job)
        self._filter_job = self.root.after(200, self._apply_filter)

    def _apply_filter(self) -> None:
        qry  = self._qry_var.get().strip().lower()
        fsel = self._file_var.get()
        use_file = fsel not in ('', 'Все файлы')

        self._tree.delete(*self._tree.get_children())

        for lbl, host in self._all:
            if use_file and lbl != fsel:
                continue
            disp = EMPTY_DISPLAY if not host else host
            if qry and qry not in lbl.lower() and qry not in disp.lower():
                continue
            chk     = (lbl, host) in self._checked
            is_emp  = not host
            tag = ('empty_chk'   if is_emp else 'checked'   ) if chk else \
                  ('empty_unchk' if is_emp else 'unchecked')
            self._tree.insert('', 'end',
                               values=(self._CHK_ON if chk else self._CHK_OFF,
                                       lbl, disp),
                               tags=(tag,))
        self._update_status()

    # ── Взаимодействие с деревом ──────────────────────────────────────────

    def _toggle(self, iid: str) -> None:
        v    = self._tree.item(iid, 'values')
        disp = v[2]
        host = '' if disp == EMPTY_DISPLAY else disp
        key  = (v[1], host)
        if key in self._checked:
            self._checked.discard(key)
            chk, is_emp = False, not host
        else:
            self._checked.add(key)
            chk, is_emp = True, not host
        tag = ('empty_chk'   if is_emp else 'checked'   ) if chk else \
              ('empty_unchk' if is_emp else 'unchecked')
        self._tree.item(iid,
                         values=(self._CHK_ON if chk else self._CHK_OFF, v[1], disp),
                         tags=(tag,))
        self._update_status()

    def _on_click(self, event: tk.Event) -> None:
        iid = self._tree.identify_row(event.y)
        if iid:
            self._toggle(iid)

    def _on_space(self, event: tk.Event) -> None:
        sel = self._tree.selection()
        if sel:
            self._toggle(sel[0])

    def _check_all(self) -> None:
        for iid in self._tree.get_children():
            v    = self._tree.item(iid, 'values')
            disp = v[2]
            host = '' if disp == EMPTY_DISPLAY else disp
            self._checked.add((v[1], host))
            is_emp = not host
            self._tree.item(iid,
                             values=(self._CHK_ON, v[1], disp),
                             tags=('empty_chk' if is_emp else 'checked',))
        self._update_status()

    def _uncheck_all(self) -> None:
        for iid in self._tree.get_children():
            v    = self._tree.item(iid, 'values')
            disp = v[2]
            host = '' if disp == EMPTY_DISPLAY else disp
            self._checked.discard((v[1], host))
            is_emp = not host
            self._tree.item(iid,
                             values=(self._CHK_OFF, v[1], disp),
                             tags=('empty_unchk' if is_emp else 'unchecked',))
        self._update_status()

    def _sort_by(self, col: str) -> None:
        idx = 0 if col == 'label' else 1
        asc = self._sort_asc[col]
        self._all.sort(key=lambda x: x[idx].lower(), reverse=not asc)
        self._sort_asc[col] = not asc
        self._apply_filter()

    def _update_status(self) -> None:
        total   = len(self._all)
        visible = len(self._tree.get_children())
        sel     = len(self._checked)
        flt     = f'  |  Показано: {visible}' if visible != total else ''
        self._stat_var.set(f'Всего хостов: {total}{flt}   |   Выбрано: {sel}')

    # ── Состояние «занят» ─────────────────────────────────────────────────

    def _set_busy2(self, busy: bool) -> None:
        self._busy = busy
        st = ('disabled',) if busy else ('!disabled',)
        for btn in (self._btn_a, self._btn_b, self._btn_del):
            btn.state(st)
        if busy:
            self._p2_var.set(0)
            self._p2.pack(fill='x', padx=12, pady=(0, 2),
                           after=self._stat_lbl)
        else:
            self._p2.pack_forget()

    def _p2_progress(self, done: int, total: int) -> None:
        pct = done / total * 100 if total else 100
        d, t = done, total
        def _u():
            self._p2_var.set(pct)
            self._stat_var.set(f'Обработка файлов: {d} / {t}')
        self.root.after(0, _u)

    # ── Сценарии ──────────────────────────────────────────────────────────

    def _ensure_folder(self) -> bool:
        if not self._folder or not Path(self._folder).is_dir():
            messagebox.showerror('Ошибка',
                                  'Папка с исходными файлами не задана.')
            return False
        return True

    def _run_scenario_a(self) -> None:
        if self._busy or not self._ensure_folder(): return
        if not messagebox.askyesno(
            'Сценарий А — Подтверждение',
            'Заменить ВСЕ имена хостов во всех исходных файлах\n'
            'на соответствующую метку файла (слово из скобок)?\n\n'
            '⚠ Действие необратимо.'
        ): return
        self._set_busy2(True)
        def run():
            try:
                n = scenario_a_apply(self._folder, self._p2_progress)
                self.root.after(0, lambda: [
                    self._set_busy2(False),
                    messagebox.showinfo('Готово — Сцен. А',
                                         f'Хосты заменены в {n} файле(ах).')])
            except Exception as exc:
                msg = str(exc)
                self.root.after(0, lambda: [self._set_busy2(False),
                                             messagebox.showerror('Ошибка', msg)])
        threading.Thread(target=run, daemon=True).start()

    def _run_scenario_b(self) -> None:
        if self._busy or not self._ensure_folder(): return
        if not self._checked:
            messagebox.showwarning('Ничего не выбрано',
                                    'Отметьте хотя бы один хост.'); return
        n = len(self._checked)
        if not messagebox.askyesno(
            'Сценарий Б — Подтверждение',
            f'Добавить метку "dang" для {n} хоста(-ов) '
            'в исходных файлах?\n\n⚠ Необратимо.'
        ): return
        self._set_busy2(True)
        snap = frozenset(self._checked)
        def run():
            try:
                scenario_b_apply(self._folder, snap, self._p2_progress)
                self.root.after(0, lambda: [
                    self._set_busy2(False),
                    messagebox.showinfo('Готово — Сцен. Б',
                                         f'"dang" проставлен для {len(snap)} хоста(-ов).')])
            except Exception as exc:
                msg = str(exc)
                self.root.after(0, lambda: [self._set_busy2(False),
                                             messagebox.showerror('Ошибка', msg)])
        threading.Thread(target=run, daemon=True).start()

    def _run_delete(self) -> None:
        if self._busy or not self._ensure_folder(): return
        if not self._checked:
            messagebox.showwarning('Ничего не выбрано',
                                    'Отметьте хотя бы один хост.'); return
        n = len(self._checked)
        if not messagebox.askyesno(
            'Удаление — Подтверждение',
            f'Удалить ВСЕ строки с {n} выбранным(и) хостом(-ами)\n'
            'из исходных файлов?\n\n⚠ Необратимо.'
        ): return
        self._set_busy2(True)
        snap = frozenset(self._checked)
        def run():
            try:
                delete_hosts_apply(self._folder, snap, self._p2_progress)
                self.root.after(0, lambda: self._after_delete(snap))
            except Exception as exc:
                msg = str(exc)
                self.root.after(0, lambda: [self._set_busy2(False),
                                             messagebox.showerror('Ошибка', msg)])
        threading.Thread(target=run, daemon=True).start()

    def _after_delete(self, deleted: frozenset) -> None:
        self._all     = [(l, h) for l, h in self._all if (l, h) not in deleted]
        self._checked -= deleted
        self._update_file_combo(preserve=True)
        self._apply_filter()
        self._set_busy2(False)
        messagebox.showinfo('Готово — Удаление',
                             f'Удалены строки с хостами: {len(deleted)}.')


# ══════════════════════════════════════════════════════════════════════════
#  Точка входа
# ══════════════════════════════════════════════════════════════════════════

def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()