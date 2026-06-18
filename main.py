import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import re
from datetime import date
import openpyxl
import pandas as pd

# ── 상품 매핑 로드 ──────────────────────────────────────────
def load_mapping(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = None
    for name in wb.sheetnames:
        if '상품' in name and '매칭' in name:
            sheet = wb[name]
            break
    if sheet is None:
        sheet = wb[wb.sheetnames[1]]

    mapping = {}
    for row in sheet.iter_rows(values_only=True):
        ss, c24 = row[0], row[1]
        if ss and c24:
            try:
                ss_key = str(int(float(ss)))
                c24_val = int(float(c24))
                if c24_val > 0:
                    mapping[ss_key] = c24_val
            except (ValueError, TypeError):
                pass
    wb.close()
    return mapping

# ── 리뷰 파일 로드 ──────────────────────────────────────────
def load_review(path):
    # pandas로 읽기 (openpyxl이 일부 네이버 파일을 잘못 읽는 문제 방지)
    df = pd.read_excel(path, header=None, dtype=str)
    df = df.fillna('')

    # 헤더 행 찾기 ('상품번호' 포함된 행)
    header_idx = 0
    for i, row in df.iterrows():
        if any('상품번호' in str(v) for v in row.values):
            header_idx = i
            break

    headers = list(df.iloc[header_idx])
    data = [list(row) for _, row in df.iloc[header_idx + 1:].iterrows()
            if any(v.strip() for v in row)]
    return headers, data

# ── 변환 로직 ────────────────────────────────────────────────
OUT_HEADERS = [
    '리뷰_id','상품_id','리뷰_작성_일자','리뷰_작성_시간','리뷰_작성자명',
    '리뷰_제목','리뷰_내용','리뷰_별점','관리자_댓글',
    '구매옵션_옵션명1','구매옵션_옵션값1','구매옵션_옵션명2','구매옵션_옵션값2',
    '구매옵션_옵션명3','구매옵션_옵션값3','구매옵션_옵션명4','구매옵션_옵션값4',
    '구매옵션_옵션명5','구매옵션_옵션값5',
    '고객정보_정보명1','고객정보_답변값1','고객정보_정보명2','고객정보_답변값2',
    '고객정보_정보명3','고객정보_답변값3','고객정보_정보명4','고객정보_답변값4',
    '고객정보_정보명5','고객정보_답변값5',
    'URL_이미지1','URL_이미지2','URL_이미지3','URL_이미지4','URL_이미지5',
    'URL_이미지6','URL_이미지7','URL_이미지8','URL_이미지9','URL_이미지10',
    'URL_동영상1','URL_동영상2','URL_동영상3','URL_동영상4','URL_동영상5',
    'URL_동영상6','URL_동영상7','URL_동영상8','URL_동영상9','URL_동영상10',
]

def convert(mapping, headers, data, date_str):
    def cidx(keyword):
        for i, h in enumerate(headers):
            if keyword in h:
                return i
        return -1

    i_ss   = cidx('상품번호')
    i_rate = cidx('평점')
    i_img  = cidx('포토')
    i_cont = cidx('리뷰상세내용')
    i_auth = cidx('등록자')
    i_date = cidx('등록일')

    unmapped = set()
    out_rows = []

    for seq, row in enumerate(data, 1):
        def g(i):
            v = row[i] if 0 <= i < len(row) else None
            return str(v).strip() if v is not None else ''

        ss_num = re.sub(r'\.0$', '', g(i_ss))
        cafe24_id = mapping.get(ss_num, '')
        if ss_num and not cafe24_id:
            unmapped.add(ss_num)

        # 날짜 파싱
        raw_date = g(i_date)
        m = re.search(r'(\d{4})\.(\d{2})\.(\d{2})\.\s*(\d{2}:\d{2}:\d{2})', raw_date)
        if m:
            rev_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            rev_time = m.group(4)
        else:
            parts = raw_date.split()
            rev_date = parts[0].replace('.', '-').rstrip('-') if parts else ''
            rev_time = parts[1] if len(parts) > 1 else ''

        # 이미지 URL
        img_raw = g(i_img)
        img_urls = [u.strip() for u in re.split(r'[,\n]+', img_raw) if u.strip().startswith('http')]

        out = [''] * len(OUT_HEADERS)
        out[0]  = f"{date_str}{seq:06d}"
        out[1]  = str(cafe24_id) if cafe24_id else ''
        out[2]  = rev_date
        out[3]  = rev_time
        out[4]  = g(i_auth)
        out[5]  = ''
        out[6]  = g(i_cont)
        out[7]  = re.sub(r'\.0$', '', g(i_rate))
        for j, url in enumerate(img_urls[:10]):
            out[29 + j] = url

        out_rows.append(out)

    return out_rows, unmapped

def save_xlsx(out_rows, save_path):
    wb = openpyxl.Workbook()

    # 안내 시트
    ws_guide = wb.active
    ws_guide.title = '안내 사항'
    ws_guide['A1'] = '브이리뷰 어드민 > 리뷰 연동 > 리뷰 파일 이관 메뉴에서 업로드해 주세요.'

    # review 시트
    ws = wb.create_sheet('review')
    ws.append(OUT_HEADERS)
    for row in out_rows:
        ws.append(row)

    # 컬럼 너비
    col_widths = {'A': 22, 'B': 12, 'C': 14, 'D': 12, 'E': 16, 'G': 40}
    for col in ws.iter_cols(min_col=30, max_col=49):
        ws.column_dimensions[col[0].column_letter].width = 60
    for col_letter, w in col_widths.items():
        ws.column_dimensions[col_letter].width = w

    wb.save(save_path)


# ══════════════════════════════════════════════════════════════
#  GUI
# ══════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('브이리뷰 변환기')
        self.resizable(False, False)
        self.configure(bg='#f5f6fa')

        self.mapping_path = tk.StringVar()
        self.review_path  = tk.StringVar()
        self.date_var     = tk.StringVar(value=date.today().strftime('%Y%m%d'))

        self.mapping_data = None
        self.review_headers = None
        self.review_rows = None

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f'+{(sw-w)//2}+{(sh-h)//2}')

    def _build_ui(self):
        PAD = dict(padx=18, pady=8)
        BG = '#f5f6fa'
        CARD = '#ffffff'
        PURPLE = '#6c5ce7'

        # ── 헤더 ──
        header = tk.Frame(self, bg=PURPLE)
        header.pack(fill='x')
        tk.Label(header, text='📋  스마트스토어 → 브이리뷰 변환기',
                 font=('Malgun Gothic', 13, 'bold'), bg=PURPLE, fg='white',
                 pady=14, padx=20).pack(anchor='w')
        tk.Label(header, text='스마트스토어 리뷰 엑셀을 업로드하면 브이리뷰 이관 파일을 만들어드려요',
                 font=('Malgun Gothic', 9), bg=PURPLE, fg='#d9d4ff',
                 pady=0, padx=20).pack(anchor='w')
        tk.Frame(header, height=10, bg=PURPLE).pack()

        outer = tk.Frame(self, bg=BG)
        outer.pack(fill='both', expand=True, padx=20, pady=16)

        # ── STEP 1 ──
        self._card(outer, '① 연동작업 스프레드시트 (매핑 파일)',
                   '브이리뷰_스마트스토어_연동작업.xlsx')
        row1 = tk.Frame(self._last_card, bg=CARD)
        row1.pack(fill='x', pady=(0, 6))
        self.mapping_entry = tk.Entry(row1, textvariable=self.mapping_path,
                                      width=40, relief='flat', bg='#f0eeff',
                                      font=('Malgun Gothic', 9))
        self.mapping_entry.pack(side='left', ipady=5, padx=(0, 6))
        tk.Button(row1, text='파일 선택', command=self._pick_mapping,
                  bg=PURPLE, fg='white', relief='flat', cursor='hand2',
                  font=('Malgun Gothic', 9), padx=10).pack(side='left')
        self.mapping_status = tk.Label(self._last_card, text='', bg=CARD,
                                       font=('Malgun Gothic', 9), fg='#636e72')
        self.mapping_status.pack(anchor='w')

        # ── STEP 2 ──
        self._card(outer, '② 스마트스토어 리뷰 파일',
                   '스마트스토어에서 다운받은 리뷰 엑셀')
        row2 = tk.Frame(self._last_card, bg=CARD)
        row2.pack(fill='x', pady=(0, 6))
        self.review_entry = tk.Entry(row2, textvariable=self.review_path,
                                     width=40, relief='flat', bg='#f0eeff',
                                     font=('Malgun Gothic', 9))
        self.review_entry.pack(side='left', ipady=5, padx=(0, 6))
        tk.Button(row2, text='파일 선택', command=self._pick_review,
                  bg=PURPLE, fg='white', relief='flat', cursor='hand2',
                  font=('Malgun Gothic', 9), padx=10).pack(side='left')
        self.review_status = tk.Label(self._last_card, text='', bg=CARD,
                                      font=('Malgun Gothic', 9), fg='#636e72')
        self.review_status.pack(anchor='w')

        # ── STEP 3 ──
        self._card(outer, '③ 리뷰 ID 날짜', '리뷰_id 생성에 사용 (예: 20260114000001)')
        row3 = tk.Frame(self._last_card, bg=CARD)
        row3.pack(fill='x', pady=(0, 2))
        tk.Label(row3, text='날짜 (YYYYMMDD):', bg=CARD,
                 font=('Malgun Gothic', 9), fg='#636e72').pack(side='left', padx=(0, 8))
        tk.Entry(row3, textvariable=self.date_var, width=14,
                 relief='flat', bg='#f0eeff',
                 font=('Malgun Gothic', 10)).pack(side='left', ipady=5)

        # ── 변환 버튼 ──
        self.convert_btn = tk.Button(
            outer, text='⚡  브이리뷰 이관 파일 생성',
            command=self._start_convert,
            bg=PURPLE, fg='white', relief='flat', cursor='hand2',
            font=('Malgun Gothic', 11, 'bold'),
            pady=12, activebackground='#5a4bd1', activeforeground='white'
        )
        self.convert_btn.pack(fill='x', pady=(10, 4))

        # ── 진행바 ──
        self.progress = ttk.Progressbar(outer, mode='indeterminate', length=400)
        self.progress.pack(fill='x', pady=(0, 4))

        # ── 상태 메시지 ──
        self.status_label = tk.Label(outer, text='', bg=BG,
                                     font=('Malgun Gothic', 9), fg='#636e72',
                                     wraplength=440, justify='left')
        self.status_label.pack(anchor='w', pady=(0, 8))

    def _card(self, parent, title, subtitle=''):
        frame = tk.Frame(parent, bg='#ffffff', bd=0,
                         highlightthickness=1, highlightbackground='#e9e4ff')
        frame.pack(fill='x', pady=(0, 12))
        inner = tk.Frame(frame, bg='#ffffff')
        inner.pack(fill='x', padx=16, pady=12)
        tk.Label(inner, text=title, bg='#ffffff',
                 font=('Malgun Gothic', 10, 'bold'), fg='#6c5ce7').pack(anchor='w')
        if subtitle:
            tk.Label(inner, text=subtitle, bg='#ffffff',
                     font=('Malgun Gothic', 8), fg='#b2bec3').pack(anchor='w', pady=(1, 8))
        self._last_card = inner

    # ── 파일 선택 ──────────────────────────────────────────────
    def _pick_mapping(self):
        path = filedialog.askopenfilename(
            title='연동작업 파일 선택',
            filetypes=[('Excel', '*.xlsx *.xls'), ('All', '*.*')]
        )
        if not path:
            return
        self.mapping_path.set(path)
        self.mapping_status.config(text='⏳ 매핑 데이터 읽는 중...', fg='#a29bfe')
        self.update()
        try:
            self.mapping_data = load_mapping(path)
            self.mapping_status.config(
                text=f'✅ {len(self.mapping_data)}개 상품 매핑 로드 완료', fg='#00b894')
        except Exception as e:
            self.mapping_data = None
            self.mapping_status.config(text=f'❌ 오류: {e}', fg='#d63031')

    def _pick_review(self):
        path = filedialog.askopenfilename(
            title='스마트스토어 리뷰 파일 선택',
            filetypes=[('Excel', '*.xlsx *.xls'), ('All', '*.*')]
        )
        if not path:
            return
        self.review_path.set(path)
        self.review_status.config(text='⏳ 리뷰 파일 읽는 중...', fg='#a29bfe')
        self.update()
        try:
            self.review_headers, self.review_rows = load_review(path)
            n = len(self.review_rows)
            self.review_status.config(text=f'✅ {n}건 로드 완료', fg='#00b894')
        except Exception as e:
            self.review_headers = self.review_rows = None
            self.review_status.config(text=f'❌ 오류: {e}', fg='#d63031')

    # ── 변환 실행 ──────────────────────────────────────────────
    def _start_convert(self):
        if not self.mapping_data:
            messagebox.showwarning('알림', '① 연동작업 파일을 먼저 선택해주세요.')
            return
        if not self.review_rows:
            messagebox.showwarning('알림', '② 리뷰 파일을 먼저 선택해주세요.')
            return
        date_str = self.date_var.get().strip().replace('-', '')
        if len(date_str) != 8 or not date_str.isdigit():
            messagebox.showwarning('알림', '날짜를 YYYYMMDD 형식으로 입력해주세요.')
            return

        save_path = filedialog.asksaveasfilename(
            title='저장 위치 선택',
            defaultextension='.xlsx',
            initialfile=f'브이리뷰_이관파일_{date_str}.xlsx',
            filetypes=[('Excel', '*.xlsx')]
        )
        if not save_path:
            return

        self.convert_btn.config(state='disabled')
        self.progress.start(10)
        self.status_label.config(text='변환 중...', fg='#a29bfe')

        def run():
            try:
                out_rows, unmapped = convert(
                    self.mapping_data, self.review_headers,
                    self.review_rows, date_str
                )
                save_xlsx(out_rows, save_path)
                self.after(0, lambda: self._done(len(out_rows), unmapped, save_path))
            except Exception as e:
                self.after(0, lambda: self._error(str(e)))

        threading.Thread(target=run, daemon=True).start()

    def _done(self, count, unmapped, save_path):
        self.progress.stop()
        self.convert_btn.config(state='normal')
        msg = f'✅ 변환 완료! {count}건 → {os.path.basename(save_path)}'
        if unmapped:
            msg += f'\n⚠️ 매핑 불가 상품번호 {len(unmapped)}개: {", ".join(list(unmapped)[:5])}'
            if len(unmapped) > 5:
                msg += ' ...'
        self.status_label.config(text=msg, fg='#00b894')
        messagebox.showinfo('완료', f'{count}건 변환 완료!\n저장 위치: {save_path}')

    def _error(self, msg):
        self.progress.stop()
        self.convert_btn.config(state='normal')
        self.status_label.config(text=f'❌ 오류: {msg}', fg='#d63031')
        messagebox.showerror('오류', msg)


if __name__ == '__main__':
    app = App()
    app.mainloop()
