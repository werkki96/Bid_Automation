# 파일 역할: Tkinter 화면에서 조회조건을 입력하고 백그라운드 수집 결과를 표시합니다.
# 마지막 업데이트: 2026-10-27
# 작업자: SMI 장지수

"""Desktop app; package with PyInstaller on Windows."""
import copy
import os
import subprocess
import sys
import queue
import threading
import webbrowser
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont
from urllib.parse import urlparse
from config import KST, BASE_DIR
from config import load_settings, save_ui_keywords
from exporter import EXCEL_COLUMNS
from ui_runner import parse_keywords, run_ui
from ui_helpers import format_price

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('입찰공고 통합 검색')
        self.geometry('1380x920')
        self.minsize(1100, 800)
        self.configure(background='#f5f7fb')
        families = set(tkfont.families(self))
        self.ui_font = next((name for name in ('맑은 고딕', 'Malgun Gothic', 'Noto Sans CJK KR', 'Apple SD Gothic Neo')
                             if name in families), 'TkDefaultFont')
        self.queue = queue.Queue()
        self.running = False
        self.rows = {'nara': [], 'd2b': []}
        self.highlighted_count = {'nara': 0, 'd2b': 0}
        self.excluded_count = {'nara': 0, 'd2b': 0}
        self.tables = {}
        self.selected_cells = {}
        self.keyword_editors = {}
        settings = load_settings()
        self.keyword_values = {
            'highlight': list(settings.get('highlight_keywords', [])),
            'exclude': list(settings.get('exclude_keywords', [])),
        }
        self.count_labels = {}
        self.source_status = {'nara': '대기 중', 'd2b': '대기 중'}
        self.saved_files = []
        self._configure_style()

        today = datetime.now(KST).date()
        header = tk.Frame(self, background='#142d4e', padx=24, pady=12)
        header.pack(fill='x')
        tk.Label(header, text='SMI  /  BID NOTICE WORKSPACE', font=(self.ui_font, 9, 'bold'),
                 foreground='#a9cde7', background='#142d4e').pack(anchor='w')
        tk.Label(header, text='검토할 공고를 한눈에.', font=(self.ui_font, 20, 'bold'),
                 foreground='white', background='#142d4e').pack(anchor='w', pady=(4, 3))
        tk.Label(header, text='나라장터 · 국방전자조달  |  검색하고, 검토하고, Excel로 저장하세요.',
                 foreground='#d1e1ef', background='#142d4e', font=(self.ui_font, 9)).pack(anchor='w')

        fields = ttk.LabelFrame(self, text='  조회 조건  ', padding=12, style='Card.TLabelframe')
        fields.pack(fill='x', padx=22, pady=(12, 10))
        ttk.Label(fields, text='조회 기간 · KST', style='Card.TLabel').grid(row=0, column=0, columnspan=4, sticky='w', pady=(0, 6))
        ttk.Label(fields, text='포함 검색어', style='Card.TLabel').grid(row=0, column=4, columnspan=2, sticky='w', padx=(18, 0), pady=(0, 6))
        self.start = ttk.Entry(fields, width=13)
        self.start.insert(0, str(today-timedelta(days=1)))
        self.start.grid(row=1, column=0, sticky='ew')
        ttk.Label(fields, text=' — ', style='Card.TLabel').grid(row=1, column=1)
        self.end = ttk.Entry(fields, width=13)
        self.end.insert(0, str(today))
        self.end.grid(row=1, column=2, sticky='ew')
        self.keywords = ttk.Entry(fields)
        self.keywords.insert(0, ', '.join(settings.get('include_keywords', [])))
        self.keywords.grid(row=1, column=4, columnspan=2, sticky='ew', padx=(18, 18))
        self.keywords.bind('<Return>', self._start_from_keyword_entry)
        self.button = ttk.Button(fields, text='공고 검색 시작', command=self.start_job, style='Accent.TButton')
        self.button.grid(row=0, column=6, rowspan=2, sticky='nsew', ipadx=14)
        quick = ttk.Frame(fields, style='Card.TFrame')
        quick.grid(row=2, column=0, columnspan=4, sticky='w', pady=(8, 0))
        for text, days in [('오늘', 1), ('어제~오늘', 2), ('최근 7일', 7)]:
            ttk.Button(quick, text=text, command=lambda d=days: self._set_period(d), style='Small.TButton').pack(side='left', padx=(0, 4))
        ttk.Label(fields, text='쉼표로 구분 · 비워두면 전체 공고', style='Muted.Card.TLabel').grid(row=2, column=4, columnspan=2, sticky='w', padx=(18, 0))
        options = ttk.Frame(fields, style='Card.TFrame')
        options.grid(row=3, column=0, columnspan=7, sticky='ew', pady=(10, 0))
        self.nara = tk.BooleanVar(value=True)
        self.d2b = tk.BooleanVar(value=True)
        ttk.Checkbutton(options, text='나라장터', variable=self.nara, style='Card.TCheckbutton').pack(side='left', padx=(0, 12))
        ttk.Checkbutton(options, text='국방전자조달 · D2B', variable=self.d2b, style='Card.TCheckbutton').pack(side='left', padx=(0, 24))
        ttk.Button(options, text='강조 편집', command=lambda: self._open_keyword_editor('highlight')).pack(side='left', padx=(0, 6))
        ttk.Button(options, text='강조 제외 편집', command=lambda: self._open_keyword_editor('exclude')).pack(side='left')
        self.save_keywords_button = ttk.Button(options, text='기본 키워드 저장', command=self.save_keywords)
        self.save_keywords_button.pack(side='right')
        fields.columnconfigure(4, weight=1)

        summary = ttk.Frame(self)
        summary.pack(fill='x', padx=22, pady=(0, 12))
        self.summary_values = {}
        for index, (key, title) in enumerate([('total', '수집 공고'), ('highlight', '강조 공고'), ('files', '저장된 Excel')]):
            card = ttk.Frame(summary, style='Card.TFrame', padding=(18, 6))
            card.grid(row=0, column=index, sticky='ew', padx=(0, 10 if index < 2 else 0))
            summary.columnconfigure(index, weight=1)
            ttk.Label(card, text=title, style='Muted.Card.TLabel').pack(anchor='w')
            value = tk.StringVar(value='—')
            self.summary_values[key] = value
            ttk.Label(card, textvariable=value, style='Metric.TLabel').pack(anchor='w')

        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=22, pady=(0, 8))
        ttk.Label(toolbar, text='결과 내 검색').pack(side='left', padx=(0, 8))
        self.result_query = tk.StringVar()
        ttk.Entry(toolbar, textvariable=self.result_query, width=28).pack(side='left')
        self.only_highlighted = tk.BooleanVar(value=False)
        ttk.Checkbutton(toolbar, text='강조 공고만 보기', variable=self.only_highlighted,
                        command=self._render_results).pack(side='left', padx=16)
        ttk.Button(toolbar, text='결과 폴더 열기', command=self._open_output).pack(side='right')
        self.result_query.trace_add('write', lambda *_: self._render_results())

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=22, pady=(0, 8))
        self._create_source_tab('nara', '나라장터', '나라장터 수집 결과')
        self._create_source_tab('d2b', '국방전자조달 · D2B', '국방전자조달 수집 결과')
        detail = ttk.Frame(self, style='Card.TFrame', padding=(16, 10))
        detail.pack(fill='x', padx=22, pady=(0, 8))
        self.detail_title = tk.StringVar(value='공고를 선택하면 제목과 강조 근거를 확인할 수 있습니다.')
        self.detail_reason = tk.StringVar(value='노란색: 검토 후보  ·  셀 클릭 후 Ctrl+C: 복사  ·  상세보기 주소 클릭: 공고 열기')
        ttk.Label(detail, textvariable=self.detail_title, style='Card.TLabel', wraplength=1020).pack(anchor='w')
        ttk.Label(detail, textvariable=self.detail_reason, style='Muted.Card.TLabel', wraplength=1020).pack(anchor='w', pady=(4, 0))
        self.status = tk.StringVar(value='조회 조건을 설정하고 검색을 시작하세요. 화면 필터는 Excel 전체 결과에 영향을 주지 않습니다.')
        status_bar = ttk.Frame(self, padding=(22, 6))
        status_bar.pack(fill='x')
        ttk.Label(status_bar, textvariable=self.status, foreground='#52637a').pack(side='left')
        self.progress = ttk.Progressbar(status_bar, mode='indeterminate', length=140)
        self.progress.pack(side='right')
        # 작은 화면에서도 상태·상세 영역을 먼저 확보하고 표가 남은 공간을 사용합니다.
        status_bar.pack_configure(side='bottom', before=self.notebook)
        detail.pack_configure(side='bottom', after=status_bar)
        self.after(150,self.poll)

    def _start_from_keyword_entry(self, _event):
        self.start_job()
        return 'break'

    def _configure_style(self):
        style = ttk.Style(self)
        # clam은 Windows/Linux에서 색상과 패딩을 일관되게 적용할 수 있습니다.
        style.theme_use('clam')
        style.configure('.', font=(self.ui_font, 10), background='#f5f7fb', foreground='#203650')
        style.configure('TFrame', background='#f5f7fb')
        style.configure('TLabel', background='#f5f7fb')
        style.configure('Card.TFrame', background='white')
        style.configure('Card.TLabel', background='white')
        style.configure('Muted.Card.TLabel', background='white', foreground='#61738a', font=(self.ui_font, 9))
        style.configure('Metric.TLabel', background='white', foreground='#20568b', font=(self.ui_font, 18, 'bold'))
        style.configure('Card.TLabelframe', background='white', bordercolor='#dce4ef', relief='solid')
        style.configure('Card.TLabelframe.Label', background='#f5f7fb', foreground='#203f62', font=(self.ui_font, 10, 'bold'))
        style.configure('TButton', padding=(12, 7), background='white', bordercolor='#dce4ef')
        style.map('TButton', background=[('active', '#eaf0f8')])
        style.configure('Small.TButton', padding=(8, 4), font=(self.ui_font, 9))
        style.configure('Accent.TButton', background='#2260a4', foreground='white', font=(self.ui_font, 11, 'bold'))
        style.map('Accent.TButton', background=[('disabled', '#9aadc3'), ('active', '#174b85')], foreground=[('disabled', 'white')])
        style.configure('TEntry', padding=7, fieldbackground='white', bordercolor='#dce4ef')
        style.configure('TCheckbutton', background='#f5f7fb')
        style.configure('Card.TCheckbutton', background='white')
        style.configure('TNotebook', background='#f5f7fb', borderwidth=0)
        style.configure('TNotebook.Tab', padding=(20, 8), background='#e8eef6')
        style.map('TNotebook.Tab', background=[('selected', 'white')], foreground=[('selected', '#2260a4')])
        style.configure('Treeview', rowheight=34, background='white', fieldbackground='white', borderwidth=0)
        style.map('Treeview', background=[('selected', '#dbeafe')], foreground=[('selected', '#142d4e')])
        style.configure('Treeview.Heading', padding=(8, 10), background='#edf2f8', foreground='#304b6b', font=(self.ui_font, 9, 'bold'))
        style.configure('Keyword.Treeview', rowheight=34)
        style.configure('TProgressbar', background='#2260a4', troughcolor='#e2e8f0', borderwidth=0)

    def _set_period(self, days):
        today = datetime.now(KST).date()
        for entry, value in [(self.start, today-timedelta(days=days-1)), (self.end, today)]:
            entry.delete(0, 'end')
            entry.insert(0, str(value))

    def _open_output(self):
        path = (BASE_DIR / load_settings().get('output_directory', 'output')).resolve()
        try:
            if not path.is_relative_to(BASE_DIR.resolve()):
                raise ValueError('결과 폴더는 프로젝트 내부여야 합니다.')
            path.mkdir(parents=True, exist_ok=True)
            if sys.platform == 'win32':
                os.startfile(path)
            else:
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', str(path)])
        except (OSError, ValueError) as exc:
            messagebox.showerror('결과 폴더', f'폴더를 열 수 없습니다.\n{exc}', parent=self)

    def _show_detail(self, source):
        table = self.tables[source]
        if not table.selection():
            return
        values = table.item(table.selection()[0], 'values')
        self.detail_title.set(values[EXCEL_COLUMNS.index('공고명')])
        reason = values[EXCEL_COLUMNS.index('강조 근거')]
        self.detail_reason.set('강조 근거: ' + reason if reason else '일반 공고 · 자동 강조 조건에 해당하지 않습니다.')

    def _render_results(self):
        query = self.result_query.get().strip().casefold()
        self.selected_cells.clear()
        for source, table in self.tables.items():
            selected = table.selection()
            selected_values = table.item(selected[0], 'values') if selected else None
            table.delete(*table.get_children())
            rows = [row for row in self.rows[source]
                    if (not self.only_highlighted.get() or row.get('강조 근거'))
                    and (not query or query in ' '.join(str(row.get(c) or '') for c in
                         ('공고명', '공고기관(발주기관)', '수요기관')).casefold())]
            rows.sort(key=lambda row: not bool(row.get('강조 근거')))
            for index, row in enumerate(rows):
                values = [self._display_value(c, row.get(c, '')) for c in EXCEL_COLUMNS]
                tag = 'highlight' if row.get('강조 근거') else ('even' if index % 2 == 0 else 'odd')
                item = table.insert('', 'end', values=values, tags=(tag,))
                if selected_values is not None and tuple(str(v) for v in values) == tuple(selected_values):
                    table.selection_set(item)
            self.count_labels[source].set(f'{self.source_status[source]}  ·  {len(rows):,} / {len(self.rows[source]):,}건')
        self.summary_values['total'].set(f'{sum(len(rows) for rows in self.rows.values()):,}건')
        self.summary_values['highlight'].set(f'{sum(bool(row.get("강조 근거")) for rows in self.rows.values() for row in rows):,}건')
        self.summary_values['files'].set(f'{len(self.saved_files)}개')

    def _create_source_tab(self, source, title, description):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text=title)
        heading = ttk.Frame(tab)
        heading.pack(fill='x', pady=(0, 8))
        ttk.Label(heading, text=description,
                  font=(self.ui_font, 11, 'bold')).pack(side='left')
        count = tk.StringVar(value='0건')
        self.count_labels[source] = count
        ttk.Label(heading, textvariable=count, foreground='#1677c8',
                  font=(self.ui_font, 10, 'bold')).pack(side='right')

        frame = ttk.Frame(tab)
        frame.pack(fill='both', expand=True)
        table = ttk.Treeview(frame, columns=EXCEL_COLUMNS, show='headings', height=6)
        self.tables[source] = table
        column_widths = {
            '공고위치': 70,
            '공고번호': 160,
            '공고구분': 75,
            '공고일자': 145,
            '공고명': 350,
            '공고기관(발주기관)': 170,
            '수요기관': 150,
            '추정가격(원)': 100,
            '입찰참가등록 마감일시': 100,
            '상세보기 주소': 260,
        }
        for col in EXCEL_COLUMNS:
            table.heading(col, text=col)
            table.column(col, width=column_widths.get(col, 150), stretch=False)
        table.bind('<Button-1>', self._handle_detail_click)
        table.bind('<Control-c>', self._copy_selected_cell)
        table.bind('<<TreeviewSelect>>', lambda _event, key=source: self._show_detail(key))
        table.tag_configure('even', background='#f7f9fc')
        table.tag_configure('odd', background='white')
        table.tag_configure('highlight', background='#fff4bf')
        hscroll = ttk.Scrollbar(frame, orient='horizontal', command=table.xview)
        vscroll = ttk.Scrollbar(frame, orient='vertical', command=table.yview)
        table.configure(xscrollcommand=hscroll.set, yscrollcommand=vscroll.set)
        table.grid(row=0, column=0, sticky='nsew')
        vscroll.grid(row=0, column=1, sticky='ns')
        hscroll.grid(row=1, column=0, sticky='ew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def save_keywords(self):
        try:
            save_ui_keywords(
                parse_keywords(self.keywords.get()),
                self.keyword_values['exclude'],
                self.keyword_values['highlight'],
            )
        except (OSError, ValueError, TypeError) as exc:
            messagebox.showerror('저장 오류', f'키워드를 저장하지 못했습니다.\n{exc}', parent=self)
            return
        self.status.set('키워드 저장 완료')

    def _open_keyword_editor(self, kind):
        existing = self.keyword_editors.get(kind)
        if existing is not None and existing.winfo_exists():
            existing.lift()
            return
        title = '하이라이트 키워드 편집' if kind == 'highlight' else '강조 제외 조건 편집'
        window = tk.Toplevel(self)
        window.title(title)
        window.geometry('760x560')
        window.minsize(620, 420)
        window.resizable(True, True)
        window.transient(self)
        self.keyword_editors[kind] = window

        ttk.Label(
            window,
            text='일치하는 공고는 결과에 남겨두고 강조만 해제합니다. 예: 학생|학교|교육 + 연수'
            if kind == 'exclude' else '한 행에 하나의 하이라이트 키워드를 입력하세요.',
            wraplength=440,
        ).pack(fill='x', padx=12, pady=(12, 6))
        frame = ttk.Frame(window)
        frame.pack(fill='both', expand=True, padx=12)
        table = ttk.Treeview(
            frame,
            columns=('number', 'keyword'),
            show='headings',
            selectmode='browse',
            style='Keyword.Treeview',
        )
        table.heading('number', text='번호')
        table.heading('keyword', text='조건' if kind == 'exclude' else '키워드')
        table.column('number', width=70, minwidth=70, stretch=False, anchor='center')
        table.column('keyword', width=620, minwidth=400, stretch=True)
        table.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=table.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        table.configure(yscrollcommand=scrollbar.set)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        def display_value(value):
            return ' + '.join(value) if isinstance(value, list) else str(value)

        def refresh_numbers():
            for number, row_id in enumerate(table.get_children(), 1):
                current = table.item(row_id, 'values')
                table.item(row_id, values=(number, current[1]))
                table.item(row_id, tags=('even' if number % 2 == 0 else 'odd',))

        table.tag_configure('even', background='#f3f6fa')
        table.tag_configure('odd', background='white')

        for value in self.keyword_values[kind]:
            table.insert('', 'end', values=('', display_value(value)))
        refresh_numbers()

        entry = ttk.Entry(window)
        entry.pack(fill='x', padx=12, pady=(8, 4))

        def add_row():
            value = entry.get().strip()
            if not value:
                return
            table.insert('', 'end', values=('', value))
            refresh_numbers()
            entry.delete(0, 'end')

        def remove_row():
            selected = table.selection()
            if selected:
                table.delete(selected[0])
                refresh_numbers()

        def apply_rows():
            values = []
            for row_id in table.get_children():
                value = str(table.item(row_id, 'values')[1]).strip()
                if value:
                    parts = [part.strip() for part in value.split('+') if part.strip()]
                    values.append(parts if kind == 'exclude' and len(parts) > 1 else value)
            self.keyword_values[kind] = values
            close_editor()

        buttons = ttk.Frame(window)
        buttons.pack(fill='x', padx=12, pady=(4, 12))
        ttk.Button(buttons, text='+ 추가', command=add_row).pack(side='left')
        ttk.Button(buttons, text='- 제거', command=remove_row).pack(side='left', padx=6)
        ttk.Button(buttons, text='변경 적용', command=apply_rows, style='Accent.TButton').pack(side='right')
        entry.bind('<Return>', lambda _event: add_row())

        def close_editor():
            self.keyword_editors.pop(kind, None)
            window.destroy()

        window.protocol('WM_DELETE_WINDOW', close_editor)

    def _handle_detail_click(self, event):
        table = event.widget
        if table.identify_region(event.x, event.y) != 'cell':
            return
        row_id = table.identify_row(event.y)
        column_id = table.identify_column(event.x)
        if not row_id or not column_id.startswith('#'):
            return
        self.selected_cells[table] = (row_id, int(column_id[1:]) - 1)
        detail_column = f'#{EXCEL_COLUMNS.index("상세보기 주소") + 1}'
        if column_id != detail_column:
            return
        values = table.item(row_id, 'values')
        url = str(values[EXCEL_COLUMNS.index("상세보기 주소")]).strip()
        parsed = urlparse(url)
        if parsed.scheme not in ('http', 'https') or not parsed.netloc:
            messagebox.showerror('주소 오류', '유효한 상세보기 주소가 없습니다.')
            return
        if not messagebox.askyesno(
                '상세보기 이동',
                f'다음 주소로 이동하시겠습니까?\n\n{url}',
                parent=self):
            return
        if not webbrowser.open_new_tab(url):
            messagebox.showerror('브라우저 오류', '상세보기 주소를 열 수 없습니다.')

    def _copy_selected_cell(self, event):
        table = event.widget
        selected = self.selected_cells.get(table)
        if selected is None:
            return 'break'
        row_id, column_index = selected
        values = table.item(row_id, 'values')
        if column_index >= len(values):
            return 'break'
        self.clipboard_clear()
        self.clipboard_append(str(values[column_index]))
        self.update()
        return 'break'

    @staticmethod
    def _display_value(column, value):
        if column != '추정가격(원)' or value in (None, ''):
            return value or ''
        return format_price(value)

    def start_job(self):
        if self.running:
            return
        try:
            first = datetime.strptime(self.start.get().strip(),'%Y-%m-%d').date()
            last = datetime.strptime(self.end.get().strip(),'%Y-%m-%d').date()
            if first > last:
                raise ValueError('시작일은 종료일보다 늦을 수 없습니다.')
        except ValueError as exc:
            messagebox.showerror('날짜 오류',str(exc))
            return
        sources = tuple(name for name, flag in (('nara',self.nara.get()),('d2b',self.d2b.get())) if flag)
        if not sources:
            messagebox.showerror('선택 오류','검색 대상을 선택하세요.')
            return
        for source, table in self.tables.items():
            table.delete(*table.get_children())
            self.rows[source] = []
            self.highlighted_count[source] = 0
            self.excluded_count[source] = 0
            self.count_labels[source].set('0건')
            self.source_status[source] = '대기 중'
        self.saved_files = []
        self.selected_cells.clear()
        self.detail_title.set('수집 중입니다. 공고를 선택하면 내용을 확인할 수 있습니다.')
        self.detail_reason.set('노란색은 검토 후보입니다.')
        self._render_results()
        self.running = True
        self.progress.start(12)
        self.button.config(state='disabled', text='수집 중…')
        self.status.set('수집 시작...')
        # Capture tkinter state before crossing thread boundary.
        search_words = self.keywords.get()
        excluded_words = copy.deepcopy(self.keyword_values['exclude'])
        highlighted_words = list(self.keyword_values['highlight'])
        def safe_worker():
            try:
                run_ui(str(first), str(last), search_words, self.queue.put, sources,
                       excluded_words, highlighted_words)
            except Exception as exc:
                self.queue.put({'type':'error','message':str(exc)})
                self.queue.put({'type':'done','failures':['setup']})
        threading.Thread(target=safe_worker,daemon=True).start()

    def poll(self):
        changed = False
        # 한 번에 처리할 이벤트 수를 제한하여 수집량이 많아도 화면 이벤트 처리가 멈추지 않게 합니다.
        for _ in range(400):
            try:
                event = self.queue.get_nowait()
            except queue.Empty:
                break
            if event['type']=='page':
                source = event['source']
                table = self.tables[source]
                self.rows[source].extend(event['rows'])
                self.highlighted_count[source] = sum(bool(row.get('강조 근거')) for row in self.rows[source])
                self.source_status[source] = f"{event['page']}/{event['pages']}페이지"
                changed = True
                self.status.set(
                    f"{'나라장터' if source == 'nara' else 'D2B'} "
                    f"{event['page']}/{event['pages']}페이지 · "
                    f"{len(self.rows[source])}건")
            elif event['type']=='saved':
                source = event['source']
                self.source_status[source] = '저장 완료'
                self.saved_files.append(event['path'])
                changed = True
                self.status.set(f"{'나라장터' if source == 'nara' else 'D2B'} 저장 완료")
            elif event['type']=='error':
                source = event.get('source')
                if source in self.source_status:
                    self.source_status[source] = '확인 필요'
                    changed = True
                self.status.set('오류: '+event['message'])
                messagebox.showerror('수집 오류',event['message'])
            elif event['type']=='done':
                self.running=False
                self.progress.stop()
                self.button.config(state='normal', text='다시 검색')
                total = sum(len(rows) for rows in self.rows.values())
                self.status.set(
                    f"수집 종료 · 화면 표시 {total}건" +
                    (' (일부 실패)' if event['failures'] else ''))
        if changed:
            self._render_results()
        self.after(150,self.poll)

if __name__=='__main__':
    App().mainloop()
