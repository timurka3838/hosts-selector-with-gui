#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
host_selector_gui.py  —  Этап 2
================================
Запускается ПОСЛЕ table_processor.py.
Загружает hosts_result.xlsx и показывает GUI для выбора хостов.

Сценарий А: пользователь не нашёл интересных хостов
    → в каждом исходном файле колонка 'host' заменяется
      на слово из скобок имени файла (метку).

Сценарий Б: пользователь выбрал хосты
    → в исходных файлах рядом с выбранными хостами
      добавляется колонка 'dang' со значением "dang".

Зависимости (те же, что у table_processor.py):
    pip install pandas openpyxl
"""

import os
import re
import logging
import threading
from pathlib import Path
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Tuple, Set

import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd
import openpyxl
from openpyxl import load_workbook

# ══════════════════════════════════════════════════════════════════════════
#  НАСТРОЙКИ — измените под себя
# ══════════════════════════════════════════════════════════════════════════
HOSTS_FILE  = r"D:\Programming\test\hosts_result.xlsx"       # итоговый файл этапа 1
FOLDER_PATH = r"D:\Programming\test"         # папка с исходниками
CHUNK_SIZE  = 100_000
MAX_WORKERS = os.cpu_count() or 4
# ══════════════════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-8s  %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('selector.log', encoding='utf-8'),
    ],
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────
#  Утилиты (дублируем из table_processor, чтобы файл был автономным)
# ─────────────────────────────────────────────────────────────────────────

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
    folder = Path(folder_path)
    return sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in {'.csv', '.xlsx', '.xlsm'}
    )


# ─────────────────────────────────────────────────────────────────────────
#  Сценарий А: заменить все хосты на метку файла
# ─────────────────────────────────────────────────────────────────────────

def _replace_hosts_csv(filepath: Path, new_value: str) -> None:
    enc = detect_csv_encoding(filepath)
    tmp = filepath.with_suffix('.tmp.csv')
    try:
        first = True
        for chunk in pd.read_csv(filepath, chunksize=CHUNK_SIZE,
                                  encoding=enc, low_memory=False):
            if 'host' in chunk.columns:
                chunk['host'] = new_value
            chunk.to_csv(tmp, index=False,
                         mode='w' if first else 'a',
                         header=first, encoding='utf-8-sig')
            first = False
        os.replace(tmp, filepath)
    except Exception:
        _remove(tmp)
        raise


def _replace_hosts_xlsx(filepath: Path, new_value: str) -> None:
    tmp = filepath.with_suffix('.tmp.xlsx')
    try:
        wb_r = load_workbook(filepath, read_only=True, data_only=True)
        ws_r = wb_r.active
        rows_iter = ws_r.iter_rows(values_only=True)
        headers  = list(next(rows_iter))
        host_idx = next(
            (i for i, h in enumerate(headers)
             if h is not None and str(h).strip() == 'host'), None
        )
        wb_w = openpyxl.Workbook(write_only=True)
        ws_w = wb_w.create_sheet()
        ws_w.append(headers)
        buf = []
        for row in rows_iter:
            row_list = list(row)
            if host_idx is not None:
                row_list[host_idx] = new_value
            buf.append(row_list)
            if len(buf) >= CHUNK_SIZE:
                for r in buf:
                    ws_w.append(r)
                buf.clear()
        for r in buf:
            ws_w.append(r)
        wb_r.close()
        wb_w.save(tmp)
        os.replace(tmp, filepath)
    except Exception:
        _remove(tmp)
        raise


def scenario_a_apply(folder_path: str, progress_cb=None) -> int:
    """Заменяет все значения 'host' в исходных файлах на метку файла."""
    files = get_source_files(folder_path)
    done  = 0

    def process(fp: Path):
        label = extract_label(fp.name)
        if fp.suffix.lower() == '.csv':
            _replace_hosts_csv(fp, label)
        else:
            _replace_hosts_xlsx(fp, label)
        log.info(f"[А] хосты → «{label}»: {fp.name}")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futs = {pool.submit(process, f): f for f in files}
        for fut in as_completed(futs):
            fut.result()   # пробрасываем исключения
            done += 1
            if progress_cb:
                progress_cb(done, len(files))

    return done


# ─────────────────────────────────────────────────────────────────────────
#  Сценарий Б: добавить колонку 'dang' для выбранных хостов
# ─────────────────────────────────────────────────────────────────────────

def _mark_dang_csv(filepath: Path, hosts: Set[str]) -> None:
    enc = detect_csv_encoding(filepath)
    tmp = filepath.with_suffix('.tmp.csv')
    try:
        first = True
        for chunk in pd.read_csv(filepath, chunksize=CHUNK_SIZE,
                                  encoding=enc, low_memory=False):
            if 'host' in chunk.columns:
                chunk['dang'] = chunk['host'].apply(
                    lambda h: 'dang' if str(h) in hosts else '')
            chunk.to_csv(tmp, index=False,
                         mode='w' if first else 'a',
                         header=first, encoding='utf-8-sig')
            first = False
        os.replace(tmp, filepath)
    except Exception:
        _remove(tmp)
        raise


def _mark_dang_xlsx(filepath: Path, hosts: Set[str]) -> None:
    tmp = filepath.with_suffix('.tmp.xlsx')
    try:
        wb_r = load_workbook(filepath, read_only=True, data_only=True)
        ws_r = wb_r.active
        rows_iter = ws_r.iter_rows(values_only=True)
        headers  = list(next(rows_iter))
        host_idx = next(
            (i for i, h in enumerate(headers)
             if h is not None and str(h).strip() == 'host'), None
        )

        str_headers = [str(h).strip() if h else '' for h in headers]
        if 'dang' in str_headers:
            dang_idx    = str_headers.index('dang')
            new_headers = list(headers)
        else:
            dang_idx    = len(headers)
            new_headers = list(headers) + ['dang']

        wb_w = openpyxl.Workbook(write_only=True)
        ws_w = wb_w.create_sheet()
        ws_w.append(new_headers)

        buf = []
        for row in rows_iter:
            row_list = list(row)
            host_val = (str(row_list[host_idx])
                        if host_idx is not None and row_list[host_idx] is not None
                        else '')
            mark = 'dang' if host_val in hosts else ''
            if dang_idx < len(row_list):
                row_list[dang_idx] = mark
            else:
                row_list.append(mark)
            buf.append(row_list)
            if len(buf) >= CHUNK_SIZE:
                for r in buf:
                    ws_w.append(r)
                buf.clear()
        for r in buf:
            ws_w.append(r)

        wb_r.close()
        wb_w.save(tmp)
        os.replace(tmp, filepath)
    except Exception:
        _remove(tmp)
        raise


def scenario_b_apply(folder_path: str,
                     checked: Set[Tuple[str, str]],
                     progress_cb=None) -> int:
    """Добавляет 'dang' в исходные файлы рядом с выбранными хостами."""
    label_to_hosts: dict = defaultdict(set)
    for label, host in checked:
        label_to_hosts[label].add(host)

    files = get_source_files(folder_path)
    done  = 0

    def process(fp: Path):
        label = extract_label(fp.name)
        if label not in label_to_hosts:
            return
        hosts = label_to_hosts[label]
        if fp.suffix.lower() == '.csv':
            _mark_dang_csv(fp, hosts)
        else:
            _mark_dang_xlsx(fp, hosts)
        log.info(f"[Б] помечено {len(hosts)} хост(ов) в: {fp.name}")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futs = {pool.submit(process, f): f for f in files}
        for fut in as_completed(futs):
            fut.result()
            done += 1
            if progress_cb:
                progress_cb(done, len(files))

    return done


# ─────────────────────────────────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────────────────────────────────

class HostSelectorApp:
    """Главное окно приложения выбора хостов."""

    CHK_ON  = '☑'
    CHK_OFF = '☐'

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title('Выбор хостов — Этап 2')
        self.root.geometry('1020x680')
        self.root.minsize(720, 500)

        self._all: List[Tuple[str, str]] = []      # все (метка, хост)
        self._checked: Set[Tuple[str, str]] = set()  # отмеченные
        self._busy     = False
        self._sort_asc = {'label': True, 'host': True}
        self._filter_job = None   # id отложенного after-вызова (debounce)

        self._build_ui()
        self.root.after(100, self._load)

    # ── Построение интерфейса ─────────────────────────────────────────────

    def _build_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        style.configure('Action.TButton', font=('Arial', 10, 'bold'), padding=8)
        style.configure('THeading', font=('Arial', 10, 'bold'))

        # ── Верхняя панель: фильтр + кнопки ──────────────────────────────
        top = ttk.Frame(self.root, padding=(10, 8, 10, 6))
        top.pack(fill='x')

        ttk.Label(top, text='🔍  Фильтр:', font=('Arial', 10)).pack(side='left')

        self._filter_var = tk.StringVar()
        self._filter_var.trace_add('write', lambda *_: self._schedule_filter())
        ttk.Entry(top, textvariable=self._filter_var,
                  width=40, font=('Arial', 10)).pack(side='left', padx=(4, 2))
        ttk.Button(top, text='✕', width=3,
                   command=lambda: self._filter_var.set('')).pack(side='left', padx=(0, 14))

        ttk.Separator(top, orient='vertical').pack(side='left', fill='y', padx=6)
        ttk.Button(top, text='☑  Выбрать все',
                   command=self._check_all).pack(side='left', padx=4)
        ttk.Button(top, text='☐  Снять всё',
                   command=self._uncheck_all).pack(side='left', padx=4)

        # ── Treeview ──────────────────────────────────────────────────────
        mid = ttk.Frame(self.root, padding=(10, 0, 10, 0))
        mid.pack(fill='both', expand=True)

        cols = ('chk', 'label', 'host')
        self._tree = ttk.Treeview(mid, columns=cols, show='headings',
                                   selectmode='browse')

        self._tree.heading('chk',   text='')
        self._tree.heading('label', text='Название файла  ▲▼',
                           command=lambda: self._sort_by('label'))
        self._tree.heading('host',  text='Имя хоста  ▲▼',
                           command=lambda: self._sort_by('host'))

        self._tree.column('chk',   width=36,  minwidth=36,  stretch=False, anchor='center')
        self._tree.column('label', width=220, minwidth=120, stretch=False, anchor='w')
        self._tree.column('host',  width=660, minwidth=200, stretch=True,  anchor='w')

        vsb = ttk.Scrollbar(mid, orient='vertical',   command=self._tree.yview)
        hsb = ttk.Scrollbar(mid, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right',  fill='y')
        hsb.pack(side='bottom', fill='x')
        self._tree.pack(fill='both', expand=True)

        self._tree.bind('<Button-1>', self._on_click)
        self._tree.bind('<space>',    self._on_space)
        self._tree.tag_configure('checked',   background='#c8e6c9', foreground='#1b5e20')
        self._tree.tag_configure('unchecked', background='',         foreground='')

        # ── Строка состояния ──────────────────────────────────────────────
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=(4, 0))
        status_row = ttk.Frame(self.root, padding=(10, 3, 10, 3))
        status_row.pack(fill='x')
        self._status_var = tk.StringVar(value='Загрузка...')
        ttk.Label(status_row, textvariable=self._status_var,
                  font=('Arial', 9), anchor='w').pack(fill='x')

        # ── Прогрессбар (скрыт по умолчанию) ─────────────────────────────
        self._progress_var = tk.DoubleVar(value=0)
        self._progress = ttk.Progressbar(
            self.root, variable=self._progress_var,
            mode='determinate', maximum=100)

        # ── Нижняя панель: кнопки действий ───────────────────────────────
        ttk.Separator(self.root, orient='horizontal').pack(fill='x')
        bot = ttk.Frame(self.root, padding=(10, 10, 10, 12))
        bot.pack(fill='x')

        self._btn_a = ttk.Button(
            bot,
            text='Нет интересных хостов\n→ Заменить хосты именем файла (Сценарий А)',
            style='Action.TButton',
            command=self._run_scenario_a,
            width=46,
        )
        self._btn_a.pack(side='left', padx=(0, 8))

        self._btn_b = ttk.Button(
            bot,
            text='✔  Отметить выбранные хосты как "dang"  (Сценарий Б)',
            style='Action.TButton',
            command=self._run_scenario_b,
            width=46,
        )
        self._btn_b.pack(side='right', padx=(8, 0))

    # ── Загрузка данных ───────────────────────────────────────────────────

    def _load(self) -> None:
        try:
            df = pd.read_excel(HOSTS_FILE)

            if 'Название файла' in df.columns and 'Имя хоста' in df.columns:
                lbl_col, host_col = 'Название файла', 'Имя хоста'
            else:
                # авто-определение: последние две колонки
                cols = list(df.columns)
                lbl_col, host_col = cols[-2], cols[-1]
                log.warning(f"Авто-определение колонок: «{lbl_col}», «{host_col}»")

            self._all = list(zip(df[lbl_col].astype(str), df[host_col].astype(str)))
            self._apply_filter()
            self._update_status()

        except FileNotFoundError:
            messagebox.showerror('Файл не найден',
                                  f'Не удалось найти файл хостов:\n{HOSTS_FILE}')
        except Exception as exc:
            log.error('Ошибка загрузки хостов', exc_info=True)
            messagebox.showerror('Ошибка загрузки', str(exc))

    # ── Управление Treeview ───────────────────────────────────────────────

    def _schedule_filter(self) -> None:
        """Запускает фильтрацию с задержкой 200 мс (debounce)."""
        if self._filter_job:
            self.root.after_cancel(self._filter_job)
        self._filter_job = self.root.after(200, self._apply_filter)

    def _apply_filter(self) -> None:
        query = self._filter_var.get().strip().lower()
        self._tree.delete(*self._tree.get_children())

        for label, host in self._all:
            if query and query not in label.lower() and query not in host.lower():
                continue
            checked = (label, host) in self._checked
            self._tree.insert(
                '', 'end',
                values=(self.CHK_ON if checked else self.CHK_OFF, label, host),
                tags=('checked' if checked else 'unchecked',),
            )
        self._update_status()

    def _toggle(self, iid: str) -> None:
        v     = self._tree.item(iid, 'values')
        key   = (v[1], v[2])
        on    = key in self._checked
        self._checked.discard(key) if on else self._checked.add(key)
        new_v = (self.CHK_OFF if on else self.CHK_ON, v[1], v[2])
        tag   = 'unchecked' if on else 'checked'
        self._tree.item(iid, values=new_v, tags=(tag,))
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
            v = self._tree.item(iid, 'values')
            self._checked.add((v[1], v[2]))
            self._tree.item(iid, values=(self.CHK_ON, v[1], v[2]),
                             tags=('checked',))
        self._update_status()

    def _uncheck_all(self) -> None:
        for iid in self._tree.get_children():
            v = self._tree.item(iid, 'values')
            self._checked.discard((v[1], v[2]))
            self._tree.item(iid, values=(self.CHK_OFF, v[1], v[2]),
                             tags=('unchecked',))
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
        checked = len(self._checked)
        flt     = f'  |  Показано: {visible}' if visible != total else ''
        self._status_var.set(
            f'Всего хостов: {total}{flt}   |   Выбрано: {checked}')

    # ── Состояние «занят» ─────────────────────────────────────────────────

    def _set_busy(self, busy: bool, progress: float = 0) -> None:
        self._busy = busy
        st = ('disabled',) if busy else ('!disabled',)
        self._btn_a.state(st)
        self._btn_b.state(st)
        if busy:
            self._progress_var.set(progress)
            self._progress.pack(fill='x', padx=10, before=self._tree.master)
        else:
            self._progress.pack_forget()

    def _update_progress(self, done: int, total: int) -> None:
        pct = (done / total * 100) if total else 100
        self.root.after(0, lambda: self._progress_var.set(pct))
        self.root.after(0, lambda: self._status_var.set(
            f'Обработка файлов: {done} / {total}'))

    # ── Сценарии ──────────────────────────────────────────────────────────

    def _run_scenario_a(self) -> None:
        if self._busy:
            return
        if not messagebox.askyesno(
            'Подтверждение — Сценарий А',
            'Заменить ВСЕ имена хостов во всех исходных файлах\n'
            'на соответствующую метку файла (слово из скобок)?\n\n'
            '⚠  Действие необратимо.'
        ):
            return

        self._set_busy(True)

        def run() -> None:
            try:
                n = scenario_a_apply(FOLDER_PATH, self._update_progress)
                self.root.after(0, lambda: [
                    self._set_busy(False),
                    messagebox.showinfo(
                        'Готово — Сценарий А',
                        f'Все хосты заменены на метки файлов.\n'
                        f'Обработано файлов: {n}'),
                ])
            except Exception as exc:
                log.error('Сценарий А: ошибка', exc_info=True)
                msg = str(exc)
                self.root.after(0, lambda: [
                    self._set_busy(False),
                    messagebox.showerror('Ошибка', msg),
                ])

        threading.Thread(target=run, daemon=True).start()

    def _run_scenario_b(self) -> None:
        if self._busy:
            return
        if not self._checked:
            messagebox.showwarning('Ничего не выбрано',
                                    'Отметьте хотя бы один хост.')
            return
        n = len(self._checked)
        if not messagebox.askyesno(
            'Подтверждение — Сценарий Б',
            f'Добавить метку "dang" для {n} хоста(-ов)\n'
            'в исходных файлах?\n\n'
            '⚠  Действие необратимо.'
        ):
            return

        self._set_busy(True)
        snapshot = frozenset(self._checked)

        def run() -> None:
            try:
                scenario_b_apply(FOLDER_PATH, snapshot, self._update_progress)
                self.root.after(0, lambda: [
                    self._set_busy(False),
                    messagebox.showinfo(
                        'Готово — Сценарий Б',
                        f'Метка "dang" проставлена для {len(snapshot)} хоста(-ов).'),
                ])
            except Exception as exc:
                log.error('Сценарий Б: ошибка', exc_info=True)
                msg = str(exc)
                self.root.after(0, lambda: [
                    self._set_busy(False),
                    messagebox.showerror('Ошибка', msg),
                ])

        threading.Thread(target=run, daemon=True).start()


# ─────────────────────────────────────────────────────────────────────────
#  Точка входа
# ─────────────────────────────────────────────────────────────────────────

def launch_gui() -> None:
    """Запускает GUI (можно вызывать из table_processor.py)."""
    root = tk.Tk()
    HostSelectorApp(root)
    root.mainloop()


if __name__ == '__main__':
    launch_gui()