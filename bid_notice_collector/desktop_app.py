# 파일 역할: Tkinter 화면에서 조회조건을 입력하고 백그라운드 수집 결과를 표시합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""Desktop app; package with PyInstaller on Windows."""
import queue
import threading
import webbrowser
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox
from urllib.parse import urlparse
from config import KST
from config import load_settings, save_ui_keywords
from exporter import EXCEL_COLUMNS
from ui_runner import parse_keywords, run_ui

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('입찰공고 통합 검색')
        self.geometry('1280x760')
        self.minsize(1000, 620)
        self.configure(background='#f4f6f8')
        self.queue = queue.Queue()
        self.running = False
        self.rows = {'nara': [], 'd2b': []}
        self.highlighted_count = {'nara': 0, 'd2b': 0}
        self.excluded_count = {'nara': 0, 'd2b': 0}
        self.tables = {}
        self.count_labels = {}
        self.source_status = {'nara': '대기 중', 'd2b': '대기 중'}
        self._configure_style()

        today = datetime.now(KST).date()
        header = tk.Frame(self, background='#17324d', padx=24, pady=18)
        header.pack(fill='x')
        tk.Label(header, text='입찰공고 통합 검색', font=('맑은 고딕', 18, 'bold'),
                 foreground='white', background='#17324d').pack(anchor='w')
        tk.Label(header, text='나라장터와 D2B 공고를 한 화면에서 조회합니다',
                 font=('맑은 고딕', 9), foreground='#c9d8e6',
                 background='#17324d').pack(anchor='w', pady=(3, 0))

        fields = ttk.LabelFrame(self, text='검색 조건', padding=14)
        fields.pack(fill='x', padx=18, pady=(16, 10))
        ttk.Label(fields, text='조회 기간').grid(row=0, column=0, padx=(0, 6), sticky='w')
        self.start = ttk.Entry(fields, width=14)
        self.start.insert(0, str(today-timedelta(days=1)))
        self.start.grid(row=0, column=1, padx=(0, 5))
        ttk.Label(fields, text='~').grid(row=0, column=2, padx=3)
        self.end = ttk.Entry(fields, width=14)
        self.end.insert(0, str(today))
        self.end.grid(row=0, column=3, padx=(0, 20))
        ttk.Label(fields, text='검색 키워드').grid(row=0, column=4, padx=(0, 6))
        settings = load_settings()
        self.keywords = ttk.Entry(fields, width=28)
        self.keywords.insert(0, ', '.join(settings.get('include_keywords', [])))
        self.keywords.grid(row=0, column=5, padx=(0, 18), sticky='ew')
        keyword_row = ttk.Frame(fields)
        keyword_row.grid(row=1, column=0, columnspan=6, pady=(10, 0), sticky='ew')
        ttk.Label(keyword_row, text='하이라이트').pack(side='left', padx=(0, 6))
        self.highlight_keywords = ttk.Entry(keyword_row, width=28)
        self.highlight_keywords.insert(
            0, ', '.join(settings.get('highlight_keywords', [])))
        self.highlight_keywords.pack(side='left', padx=(0, 18))
        ttk.Label(keyword_row, text='제외').pack(side='left', padx=(0, 6))
        self.exclude_keywords = ttk.Entry(keyword_row, width=28)
        self.exclude_keywords.insert(0, ', '.join(settings.get('exclude_keywords', [])))
        self.exclude_keywords.pack(side='left', padx=(0, 6))
        self.save_keywords_button = ttk.Button(
            fields, text='키워드 저장', command=self.save_keywords)
        self.save_keywords_button.grid(row=1, column=6, padx=(8, 0), pady=(10, 0))
        self.nara = tk.BooleanVar(value=True)
        self.d2b = tk.BooleanVar(value=True)
        ttk.Checkbutton(fields, text='나라장터', variable=self.nara).grid(
            row=2, column=1, pady=(12, 0), sticky='w')
        ttk.Checkbutton(fields, text='D2B', variable=self.d2b).grid(
            row=2, column=3, pady=(12, 0), sticky='w')
        ttk.Label(fields, text='쉼표로 구분 · 비워두면 전체 조회',
                  foreground='#6b7280').grid(row=2, column=5, pady=(12, 0), sticky='w')
        self.button = tk.Button(
            fields,
            text='검색 시작',
            command=self.start_job,
            background='#1677c8',
            foreground='white',
            activebackground='#0f5fa5',
            activeforeground='white',
            disabledforeground='#9ca3af',
            relief='flat',
            borderwidth=0,
            font=('맑은 고딕', 10, 'bold'),
            cursor='hand2',
        )
        self.button.grid(row=0, column=6, rowspan=3, padx=(8, 0),
                         ipadx=14, ipady=5)
        fields.columnconfigure(5, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=18, pady=(0, 10))
        self._create_source_tab('nara', '나라장터', '나라장터 수집 결과')
        self._create_source_tab('d2b', 'D2B', 'D2B 수집 결과')

        self.status = tk.StringVar(value='대기 중')
        status_bar = ttk.Frame(self, padding=(18, 8))
        status_bar.pack(fill='x')
        ttk.Label(status_bar, textvariable=self.status, foreground='#52606d').pack(side='left')
        self.after(150,self.poll)

    def _configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use('vista')
        except tk.TclError:
            pass
        style.configure('TFrame', background='#f4f6f8')
        style.configure('TLabelframe', background='#f4f6f8')
        style.configure('TLabelframe.Label', background='#f4f6f8',
                        foreground='#17324d', font=('맑은 고딕', 10, 'bold'))
        style.configure('TLabel', background='#f4f6f8', foreground='#263238')
        style.configure('TNotebook', background='#f4f6f8', borderwidth=0)
        style.configure('TNotebook.Tab', padding=(18, 8), font=('맑은 고딕', 10))
        style.configure('Treeview', rowheight=28, font=('맑은 고딕', 9),
                        background='white', fieldbackground='white')
        style.configure('Treeview.Heading', font=('맑은 고딕', 9, 'bold'),
                        background='#e8eef5', foreground='#17324d')

    def _create_source_tab(self, source, title, description):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text=title)
        heading = ttk.Frame(tab)
        heading.pack(fill='x', pady=(0, 8))
        ttk.Label(heading, text=description,
                  font=('맑은 고딕', 11, 'bold')).pack(side='left')
        count = tk.StringVar(value='0건')
        self.count_labels[source] = count
        ttk.Label(heading, textvariable=count, foreground='#1677c8',
                  font=('맑은 고딕', 10, 'bold')).pack(side='right')

        frame = ttk.Frame(tab)
        frame.pack(fill='both', expand=True)
        table = ttk.Treeview(frame, columns=EXCEL_COLUMNS, show='headings')
        self.tables[source] = table
        column_widths = {
            '공고위치': 70,
            '공고번호': 55,
            '공고구분': 75,
            '공고일자': 100,
            '공고명': 220,
            '공고기관(발주기관)': 70,
            '수요기관': 70,
            '추정가격(원)': 125,
            '입찰참가등록 마감일시': 170,
            '상세보기 주소': 260,
        }
        for col in EXCEL_COLUMNS:
            table.heading(col, text=col)
            table.column(col, width=column_widths.get(col, 150), stretch=False)
        table.bind('<Button-1>', self._handle_detail_click)
        table.tag_configure('highlight', background='#fff4bf')
        table.tag_configure('excluded', background='#ffc7ce')
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
                parse_keywords(self.exclude_keywords.get()),
                parse_keywords(self.highlight_keywords.get()),
            )
        except (OSError, ValueError, TypeError) as exc:
            messagebox.showerror('저장 오류', f'키워드를 저장하지 못했습니다.\n{exc}', parent=self)
            return
        self.status.set('키워드 저장 완료')

    def _handle_detail_click(self, event):
        table = event.widget
        detail_column = f'#{EXCEL_COLUMNS.index("상세보기 주소") + 1}'
        if table.identify_region(event.x, event.y) != 'cell':
            return
        if table.identify_column(event.x) != detail_column:
            return
        row_id = table.identify_row(event.y)
        if not row_id:
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

    @staticmethod
    def _display_value(column, value):
        if column != '추정가격(원)' or value in (None, ''):
            return value or ''
        try:
            number = int(str(value).replace(',', '').strip())
        except (TypeError, ValueError):
            return value
        return f'{number:,}'

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
        self.running = True
        self.button.config(state='disabled')
        self.status.set('수집 시작...')
        # Capture tkinter state before crossing thread boundary.
        search_words = self.keywords.get()
        excluded_words = self.exclude_keywords.get()
        def safe_worker():
            try:
                run_ui(str(first), str(last), search_words, self.queue.put, sources,
                       excluded_words)
            except Exception as exc:
                self.queue.put({'type':'error','message':str(exc)})
                self.queue.put({'type':'done','failures':['setup']})
        threading.Thread(target=safe_worker,daemon=True).start()

    def poll(self):
        # 한 번에 처리할 이벤트 수를 제한하여 수집량이 많아도 화면 이벤트 처리가 멈추지 않게 합니다.
        for _ in range(400):
            try:
                event = self.queue.get_nowait()
            except queue.Empty:
                break
            if event['type']=='page':
                source = event['source']
                table = self.tables[source]
                for row in event['rows']:
                    self.rows[source].append(row)
                    highlighted = bool(row.get('강조 근거'))
                    excluded = bool(row.get('_ui_excluded'))
                    if excluded:
                        position = 'end'
                    elif highlighted:
                        position = self.highlighted_count[source]
                    else:
                        position = max(
                            0,
                            len(table.get_children()) - self.excluded_count[source],
                        )
                    values = [self._display_value(c, row.get(c, '')) for c in EXCEL_COLUMNS]
                    table.insert('', position, values=values,
                                tags=('excluded' if excluded else
                                      'highlight' if highlighted else ''))
                    if highlighted:
                        self.highlighted_count[source] += 1
                    if excluded:
                        self.excluded_count[source] += 1
                    self.count_labels[source].set(f"{len(self.rows[source])}건")
                    self.source_status[source] = (
                        f"{event['page']}/{event['pages']}페이지")
                self.status.set(
                    f"{'나라장터' if source == 'nara' else 'D2B'} "
                    f"{event['page']}/{event['pages']}페이지 · "
                    f"{len(self.rows[source])}건")
            elif event['type']=='saved':
                source = event['source']
                self.source_status[source] = '저장 완료'
                self.status.set(f"{'나라장터' if source == 'nara' else 'D2B'} 저장 완료")
            elif event['type']=='error':
                source = event.get('source')
                if source in self.source_status:
                    self.source_status[source] = '오류'
                self.status.set('오류: '+event['message'])
                messagebox.showerror('수집 오류',event['message'])
            elif event['type']=='done':
                self.running=False
                self.button.config(state='normal')
                total = sum(len(rows) for rows in self.rows.values())
                self.status.set(
                    f"수집 종료 · 화면 표시 {total}건" +
                    (' (일부 실패)' if event['failures'] else ''))
        self.after(150,self.poll)

if __name__=='__main__':
    App().mainloop()
