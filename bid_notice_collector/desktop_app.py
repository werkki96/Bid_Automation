# 파일 역할: Tkinter 화면에서 조회조건을 입력하고 백그라운드 수집 결과를 표시합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""Desktop app; package with PyInstaller on Windows."""
import queue
import threading
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox
from config import KST
from exporter import EXCEL_COLUMNS
from ui_runner import run_ui

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('입찰공고 통합 검색')
        self.geometry('1250x700')
        self.queue = queue.Queue()
        self.running = False
        self.rows = []
        self.highlighted_count = 0
        today = datetime.now(KST).date()
        fields = ttk.Frame(self, padding=12)
        fields.pack(fill='x')
        ttk.Label(fields, text='시작일').grid(row=0, column=0, padx=4)
        self.start = ttk.Entry(fields, width=14)
        self.start.insert(0, str(today-timedelta(days=1)))
        self.start.grid(row=0,column=1,padx=4)
        ttk.Label(fields, text='종료일').grid(row=0,column=2,padx=4)
        self.end = ttk.Entry(fields, width=14)
        self.end.insert(0, str(today))
        self.end.grid(row=0,column=3,padx=4)
        ttk.Label(fields, text='포함 키워드').grid(row=0,column=4,padx=4)
        self.keywords = ttk.Entry(fields, width=38)
        self.keywords.grid(row=0,column=5,padx=4)
        self.nara = tk.BooleanVar(value=True)
        self.d2b = tk.BooleanVar(value=True)
        ttk.Checkbutton(fields,text='나라장터',variable=self.nara).grid(row=1,column=1)
        ttk.Checkbutton(fields,text='D2B',variable=self.d2b).grid(row=1,column=2)
        self.button = ttk.Button(fields,text='검색 시작',command=self.start_job)
        self.button.grid(row=1,column=5,pady=8)
        frame = ttk.Frame(self)
        frame.pack(fill='both',expand=True,padx=12)
        self.table = ttk.Treeview(frame, columns=EXCEL_COLUMNS, show='headings')
        for col in EXCEL_COLUMNS:
            self.table.heading(col,text=col)
            self.table.column(col,width=170,stretch=False)
        # Treeview는 셀별 배경색을 지원하지 않아 화면에서는 강조 공고의 행 전체를 표시합니다.
        self.table.tag_configure('highlight', background='#FFFF00')
        hscroll = ttk.Scrollbar(frame,orient='horizontal',command=self.table.xview)
        vscroll = ttk.Scrollbar(frame,orient='vertical',command=self.table.yview)
        self.table.configure(xscrollcommand=hscroll.set,yscrollcommand=vscroll.set)
        self.table.grid(row=0,column=0,sticky='nsew')
        vscroll.grid(row=0,column=1,sticky='ns')
        hscroll.grid(row=1,column=0,sticky='ew')
        frame.rowconfigure(0,weight=1)
        frame.columnconfigure(0,weight=1)
        self.status = tk.StringVar(value='대기 중')
        ttk.Label(self,textvariable=self.status,padding=10).pack(fill='x')
        self.after(150,self.poll)

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
        self.table.delete(*self.table.get_children())
        self.rows = []
        self.highlighted_count = 0
        self.running = True
        self.button.config(state='disabled')
        self.status.set('수집 시작...')
        # Capture tkinter state before crossing thread boundary.
        search_words = self.keywords.get()
        def safe_worker():
            try:
                run_ui(str(first),str(last),search_words,self.queue.put,sources)
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
                for row in event['rows']:
                    self.rows.append(row)
                    highlighted = bool(row.get('강조 근거'))
                    position = self.highlighted_count if highlighted else 'end'
                    self.table.insert('', position, values=[row.get(c,'') for c in EXCEL_COLUMNS],
                                      tags=('highlight',) if highlighted else ())
                    if highlighted:
                        self.highlighted_count += 1
                self.status.set(f"{event['source']} {event['page']}/{event['pages']}페이지 · {len(self.rows)}건")
            elif event['type']=='saved':
                self.status.set(f"저장 완료: {event['path']}")
            elif event['type']=='error':
                self.status.set('오류: '+event['message'])
                messagebox.showerror('수집 오류',event['message'])
            elif event['type']=='done':
                self.running=False
                self.button.config(state='normal')
                self.status.set(f"수집 종료 · 화면 표시 {len(self.rows)}건" + (' (일부 실패)' if event['failures'] else ''))
        self.after(150,self.poll)

if __name__=='__main__':
    App().mainloop()
